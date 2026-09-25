"""Snapshots imutáveis e cache efêmero por versão técnica.

O cache pertence ao coordenador de uma execução.  Ele deduplica solicitações
concorrentes, mantém referências explícitas por comparação e nunca persiste o
conteúdo no banco oficial.
"""

from __future__ import annotations

from collections.abc import Iterator, Mapping
from concurrent.futures import Future, Executor
from dataclasses import dataclass, field
from enum import StrEnum
import logging
import os
from pathlib import Path
import pickle
import threading
import time
from types import MappingProxyType
from typing import Callable

from app.excel.reader import CellValue, Snapshot, read_workbook


logger = logging.getLogger("auditoria_excel.parallel")
SnapshotKey = tuple[str, str]


class SnapshotState(StrEnum):
    NOT_REQUESTED = "NOT_REQUESTED"
    PARSING = "PARSING"
    READY = "READY"
    FAILED = "FAILED"
    RELEASED = "RELEASED"


class FrozenSheet(Mapping[str, CellValue]):
    """Mapa serializável que não expõe nenhuma operação de mutação."""

    __slots__ = ("_items", "_values")

    def __init__(self, cells: Mapping[str, CellValue]) -> None:
        self._items = tuple(cells.items())
        self._values = MappingProxyType(dict(self._items))

    def __getitem__(self, key: str) -> CellValue:
        return self._values[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._values)

    def __len__(self) -> int:
        return len(self._values)

    def __reduce__(self):
        return (type(self)._from_items, (self._items,))

    @classmethod
    def _from_items(cls, items: tuple[tuple[str, CellValue], ...]) -> "FrozenSheet":
        instance = cls.__new__(cls)
        instance._items = items
        instance._values = MappingProxyType(dict(items))
        return instance


class VersionSnapshot(Mapping[str, FrozenSheet]):
    """Resultado completo, imutável e serializável de um único XLSX."""

    __slots__ = ("workbook_identity", "technical_version_id", "_items", "_sheets")

    def __init__(self, workbook_identity: str, technical_version_id: str,
                 snapshot: Mapping[str, Mapping[str, CellValue]]) -> None:
        self.workbook_identity = workbook_identity
        self.technical_version_id = technical_version_id
        self._items = tuple((name, FrozenSheet(cells)) for name, cells in snapshot.items())
        self._sheets = MappingProxyType(dict(self._items))

    def __getitem__(self, key: str) -> FrozenSheet:
        return self._sheets[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._sheets)

    def __len__(self) -> int:
        return len(self._sheets)

    def __reduce__(self):
        return (type(self)._from_items,
                (self.workbook_identity, self.technical_version_id, self._items))

    @classmethod
    def _from_items(cls, identity: str, technical_id: str,
                    items: tuple[tuple[str, FrozenSheet], ...]) -> "VersionSnapshot":
        instance = cls.__new__(cls)
        instance.workbook_identity = identity
        instance.technical_version_id = technical_id
        instance._items = items
        instance._sheets = MappingProxyType(dict(items))
        return instance


@dataclass(frozen=True, slots=True)
class VersionSnapshotTask:
    workbook_identity: str
    technical_version_id: str
    local_path: str


@dataclass(frozen=True, slots=True)
class SnapshotParseResult:
    snapshot: VersionSnapshot
    duration: float
    estimated_bytes: int
    serialized_bytes: int
    serialize_duration: float
    worker_pid: int


def parse_version_snapshot(task: VersionSnapshotTask) -> SnapshotParseResult:
    """Parseia um arquivo local em isolamento; nenhum reader é compartilhado."""
    logger.info("SNAPSHOT_PARSE_STARTED technical_version_id=%s worker_pid=%d",
                task.technical_version_id, os.getpid())
    started = time.perf_counter()
    raw: Snapshot = read_workbook(Path(task.local_path))
    snapshot = VersionSnapshot(task.workbook_identity, task.technical_version_id, raw)
    duration = time.perf_counter() - started
    serialize_started = time.perf_counter()
    serialized_bytes = len(pickle.dumps(snapshot, protocol=5))
    serialize_duration = time.perf_counter() - serialize_started
    return SnapshotParseResult(
        snapshot, duration, serialized_bytes, serialized_bytes,
        serialize_duration, os.getpid(),
    )


@dataclass(slots=True)
class _Entry:
    dependencies: set[int] = field(default_factory=set)
    state: SnapshotState = SnapshotState.NOT_REQUESTED
    future: Future[SnapshotParseResult] | None = None
    result: SnapshotParseResult | None = None
    error: BaseException | None = None


@dataclass(frozen=True, slots=True)
class SnapshotCacheSummary:
    unique_versions: int
    parse_count: int
    reuse_count: int
    duplicate_parse_prevented: int
    peak_cache_count: int
    estimated_peak_bytes: int

    @property
    def parse_amplification(self) -> float:
        return self.parse_count / self.unique_versions if self.unique_versions else 0.0


class SnapshotCache:
    """Registry thread-safe com limite de residentes e referências explícitas."""

    def __init__(self, capacity: int) -> None:
        if capacity < 2:
            raise ValueError("snapshot cache capacity deve ser ao menos 2")
        self.capacity = capacity
        self._entries: dict[SnapshotKey, _Entry] = {}
        self._lock = threading.RLock()
        self.parse_count = 0
        self.reuse_count = 0
        self.duplicate_parse_prevented = 0
        self.peak_cache_count = 0
        self.estimated_bytes = 0
        self.estimated_peak_bytes = 0
        self.read_xlsx_total = 0.0
        self.serialize_duration_total = 0.0
        self.serialized_bytes_total = 0

    def register(self, key: SnapshotKey, dependencies: set[int]) -> None:
        with self._lock:
            entry = self._entries.setdefault(key, _Entry())
            entry.dependencies.update(dependencies)
        logger.info("SNAPSHOT_REQUESTED technical_version_id=%s sequence_dependencies=%s",
                    key[1], sorted(dependencies))

    def _resident_count(self) -> int:
        return sum(entry.state in {SnapshotState.PARSING, SnapshotState.READY}
                   for entry in self._entries.values())

    def can_request(self, key: SnapshotKey) -> bool:
        with self._lock:
            entry = self._entries[key]
            return entry.state in {SnapshotState.PARSING, SnapshotState.READY} or (
                entry.state is SnapshotState.NOT_REQUESTED
                and self._resident_count() < self.capacity
            )

    def request(self, key: SnapshotKey, local_path: str, consumer_sequence: int,
                executor: Executor,
                parser: Callable[[VersionSnapshotTask], SnapshotParseResult] = parse_version_snapshot,
                ) -> Future[SnapshotParseResult]:
        with self._lock:
            entry = self._entries[key]
            if consumer_sequence not in entry.dependencies:
                raise RuntimeError(f"dependência não registrada: {key} sequence={consumer_sequence}")
            if entry.state in {SnapshotState.PARSING, SnapshotState.READY}:
                self.reuse_count += 1
                self.duplicate_parse_prevented += 1
                logger.info("SNAPSHOT_REUSED technical_version_id=%s consumer_sequence=%d",
                            key[1], consumer_sequence)
                assert entry.future is not None
                return entry.future
            if entry.state is SnapshotState.FAILED:
                raise RuntimeError(f"snapshot FAILED: technical_version_id={key[1]}") from entry.error
            if entry.state is SnapshotState.RELEASED:
                raise RuntimeError(f"snapshot RELEASED: technical_version_id={key[1]}")
            if self._resident_count() >= self.capacity:
                raise BufferError("snapshot cache cheio; aplicar backpressure")
            entry.state = SnapshotState.PARSING
            self.parse_count += 1
            future = executor.submit(parser, VersionSnapshotTask(key[0], key[1], local_path))
            entry.future = future
            future.add_done_callback(lambda completed, cache_key=key: self._publish(cache_key, completed))
            self.peak_cache_count = max(self.peak_cache_count, self._resident_count())
            return future

    def _publish(self, key: SnapshotKey, future: Future[SnapshotParseResult]) -> None:
        with self._lock:
            entry = self._entries[key]
            try:
                result = future.result()
            except BaseException as error:
                entry.state = SnapshotState.FAILED
                entry.error = error
                entry.result = None
                logger.error("SNAPSHOT_PARSE_FAILED technical_version_id=%s", key[1], exc_info=error)
                return
            entry.result = result
            entry.state = SnapshotState.READY
            self.estimated_bytes += result.estimated_bytes
            self.read_xlsx_total += result.duration
            self.serialize_duration_total += result.serialize_duration
            self.serialized_bytes_total += result.serialized_bytes
            self.estimated_peak_bytes = max(self.estimated_peak_bytes, self.estimated_bytes)
            logger.info(
                "SNAPSHOT_PARSE_READY technical_version_id=%s duration=%.6f estimated_bytes=%d "
                "worker_pid=%d snapshot_serialize_duration=%.6f snapshot_serialized_bytes=%d",
                key[1], result.duration, result.estimated_bytes, result.worker_pid,
                result.serialize_duration, result.serialized_bytes,
            )

    def result(self, key: SnapshotKey) -> SnapshotParseResult:
        with self._lock:
            entry = self._entries[key]
            future = entry.future
        if future is None:
            raise RuntimeError(f"snapshot não solicitado: {key}")
        result = future.result()
        with self._lock:
            if self._entries[key].state is not SnapshotState.READY:
                raise RuntimeError(f"snapshot não está READY: {key}")
        return result

    def release(self, key: SnapshotKey, consumer_sequence: int) -> bool:
        with self._lock:
            entry = self._entries[key]
            if consumer_sequence not in entry.dependencies:
                return False
            entry.dependencies.remove(consumer_sequence)
            if entry.dependencies:
                return False
            if entry.state is SnapshotState.READY and entry.result is not None:
                self.estimated_bytes -= entry.result.estimated_bytes
            entry.result = None
            entry.future = None
            entry.state = SnapshotState.RELEASED
            logger.info("SNAPSHOT_RELEASED technical_version_id=%s remaining_refs=0", key[1])
            return True

    def state(self, key: SnapshotKey) -> SnapshotState:
        with self._lock:
            return self._entries[key].state

    def reset_failed(self, key: SnapshotKey) -> None:
        """Invalida atomicamente uma falha antes de um retry explícito."""
        with self._lock:
            entry = self._entries[key]
            if entry.state is not SnapshotState.FAILED:
                raise RuntimeError(f"somente snapshot FAILED pode ser resetado: {key}")
            entry.future = None
            entry.result = None
            entry.error = None
            entry.state = SnapshotState.NOT_REQUESTED
        logger.warning("SNAPSHOT_RETRY technical_version_id=%s previous_state=FAILED", key[1])

    @property
    def resident_count(self) -> int:
        with self._lock:
            return self._resident_count()

    def summary(self) -> SnapshotCacheSummary:
        with self._lock:
            return SnapshotCacheSummary(
                len(self._entries), self.parse_count, self.reuse_count,
                self.duplicate_parse_prevented, self.peak_cache_count,
                self.estimated_peak_bytes,
            )
