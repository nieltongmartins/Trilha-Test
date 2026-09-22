"""Leitura XLSX antecipada, limitada e identificada por versao."""

from __future__ import annotations

from concurrent.futures import Future, ProcessPoolExecutor
import ctypes
from dataclasses import dataclass
import os
from pathlib import Path
import pickle
import time
from uuid import uuid4

from app.excel.reader import Snapshot, read_workbook


READ_AHEAD_BUFFER_SIZE = 1


def _rss_bytes() -> int:
    """Retorna o RSS corrente sem adicionar uma dependencia operacional."""
    if os.name == "nt":
        class ProcessMemoryCounters(ctypes.Structure):
            _fields_ = [
                ("cb", ctypes.c_ulong),
                ("PageFaultCount", ctypes.c_ulong),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]

        counters = ProcessMemoryCounters()
        counters.cb = ctypes.sizeof(counters)
        handle = ctypes.windll.kernel32.GetCurrentProcess()  # type: ignore[attr-defined]
        if ctypes.windll.psapi.GetProcessMemoryInfo(  # type: ignore[attr-defined]
            handle, ctypes.byref(counters), counters.cb
        ):
            return int(counters.WorkingSetSize)
        return 0
    try:
        fields = Path("/proc/self/statm").read_text(encoding="ascii").split()
        return int(fields[1]) * os.sysconf("SC_PAGE_SIZE")
    except (OSError, ValueError, IndexError):
        return 0


@dataclass(frozen=True, slots=True)
class PreparedSnapshot:
    version_id: str
    version_label: str
    origin: str
    token: str
    snapshot: Snapshot
    read_seconds: float
    worker_rss_bytes: int
    snapshot_bytes: int
    cell_count: int


def prepare_snapshot(
    version_id: str, version_label: str, origin: str, token: str
) -> PreparedSnapshot:
    """Funcao top-level deliberadamente serializavel pelo multiprocessing."""
    started = time.perf_counter()
    snapshot = read_workbook(origin)
    return PreparedSnapshot(
        version_id=version_id,
        version_label=version_label,
        origin=origin,
        token=token,
        snapshot=snapshot,
        read_seconds=time.perf_counter() - started,
        worker_rss_bytes=_rss_bytes(),
        snapshot_bytes=len(pickle.dumps(snapshot, protocol=pickle.HIGHEST_PROTOCOL)),
        cell_count=sum(len(cells) for cells in snapshot.values()),
    )


class ReadAheadExecutor:
    """Mantem exatamente uma leitura pendente em um unico processo."""

    def __init__(self) -> None:
        self._executor = ProcessPoolExecutor(max_workers=READ_AHEAD_BUFFER_SIZE)
        self._future: Future[PreparedSnapshot] | None = None
        self._identity: tuple[str, str, str, str] | None = None

    def submit(self, version_id: str, version_label: str, path: Path) -> str:
        if self._future is not None:
            raise RuntimeError("buffer de leitura antecipada ja esta ocupado")
        token = uuid4().hex
        origin = str(path.resolve())
        self._identity = (version_id, version_label, origin, token)
        self._future = self._executor.submit(
            prepare_snapshot, version_id, version_label, origin, token
        )
        return token

    def result(
        self, version_id: str, version_label: str, path: Path, token: str
    ) -> tuple[PreparedSnapshot, bool]:
        if self._future is None or self._identity is None:
            raise RuntimeError("nao existe snapshot antecipado")
        expected = (version_id, version_label, str(path.resolve()), token)
        if expected != self._identity:
            raise RuntimeError("identidade do snapshot antecipado nao corresponde a versao")
        ready = self._future.done()
        try:
            prepared = self._future.result()
            if (
                prepared.version_id,
                prepared.version_label,
                prepared.origin,
                prepared.token,
            ) != expected:
                raise RuntimeError("worker retornou snapshot associado a versao incorreta")
            return prepared, ready
        finally:
            self._future = None
            self._identity = None

    def close(self) -> None:
        if self._future is not None:
            self._future.cancel()
        self._executor.shutdown(wait=True, cancel_futures=True)
        self._future = None
        self._identity = None


main_process_rss_bytes = _rss_bytes
