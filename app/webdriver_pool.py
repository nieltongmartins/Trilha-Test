"""Pool de atores Selenium: uma fila compartilhada e um owner por WebDriver."""

from __future__ import annotations

from concurrent.futures import Future
from dataclasses import dataclass, field
from enum import StrEnum
import logging
from queue import PriorityQueue
import threading
import time
from typing import Callable, Generic, TypeVar


logger = logging.getLogger("auditoria_excel.webdriver_pool")
T = TypeVar("T")


class DriverState(StrEnum):
    STARTING = "STARTING"
    READY = "READY"
    IDLE = "IDLE"
    DOWNLOADING = "DOWNLOADING"
    RECOVERING = "RECOVERING"
    FAILED = "FAILED"
    STOPPED = "STOPPED"


class DownloadState(StrEnum):
    NOT_REQUESTED = "NOT_REQUESTED"
    QUEUED = "QUEUED"
    DOWNLOADING = "DOWNLOADING"
    READY = "READY"
    FAILED = "FAILED"


@dataclass(frozen=True, slots=True)
class DownloadTask(Generic[T]):
    workbook_identity: str
    technical_version_id: str
    version_label: str
    sequence_priority: int
    url: str
    payload: T
    attempt: int = 0


@dataclass(slots=True)
class DriverWorker:
    driver_id: int
    webdriver: object
    owner_thread: threading.Thread | None = None
    state: DriverState = DriverState.STARTING
    current_version: str | None = None
    last_activity: float = field(default_factory=time.monotonic)
    health: str = "starting"
    downloads: int = 0
    busy_time: float = 0.0
    started_at: float = field(default_factory=time.monotonic)


class WebDriverPool(Generic[T]):
    """Executa cada comando no ator que possui exclusivamente seu driver.

    A deduplicação ocorre antes da fila. Requisições simultâneas do mesmo ID
    recebem o mesmo ``Future`` e, portanto, jamais originam dois downloads.
    """

    def __init__(self, drivers: list[object], download: Callable[[object, DownloadTask[T]], object],
                 *, validate: Callable[[object], None] | None = None,
                 recover: Callable[[object, BaseException], None] | None = None,
                 close_driver: Callable[[object], None] | None = None,
                 max_attempts: int = 2) -> None:
        if not 1 <= len(drivers) <= 4:
            raise ValueError("driver_count deve estar entre 1 e 4")
        self._download = download
        self._validate = validate
        self._recover = recover
        self._close_driver = close_driver
        self._max_attempts = max_attempts
        self._queue: PriorityQueue[tuple[int, int, DownloadTask[T], Future]] = PriorityQueue()
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._serial = 0
        self._futures: dict[str, Future] = {}
        self.download_state: dict[str, DownloadState] = {}
        self.workers = [DriverWorker(index, driver) for index, driver in enumerate(drivers, 1)]
        for worker in self.workers:
            thread = threading.Thread(target=self._run, args=(worker,),
                                      name=f"webdriver-owner-{worker.driver_id}", daemon=True)
            worker.owner_thread = thread
            thread.start()

    def submit(self, task: DownloadTask[T]) -> Future:
        with self._lock:
            existing = self._futures.get(task.technical_version_id)
            if existing is not None:
                return existing
            future: Future = Future()
            self._futures[task.technical_version_id] = future
            self.download_state[task.technical_version_id] = DownloadState.QUEUED
            self._serial += 1
            self._queue.put((task.sequence_priority, self._serial, task, future))
            return future

    def _run(self, worker: DriverWorker) -> None:
        try:
            if self._validate:
                self._validate(worker.webdriver)
            worker.state = DriverState.READY
            worker.health = "ready"
            logger.info("WEBDRIVER_READY driver_id=%d", worker.driver_id)
            while not self._stop.is_set():
                try:
                    _, _, task, future = self._queue.get(timeout=.1)
                except Exception:
                    worker.state = DriverState.IDLE
                    continue
                if future.cancelled() or future.done():
                    self._queue.task_done()
                    continue
                started = time.monotonic()
                worker.state = DriverState.DOWNLOADING
                worker.current_version = task.technical_version_id
                with self._lock:
                    self.download_state[task.technical_version_id] = DownloadState.DOWNLOADING
                logger.info("WEBDRIVER_DOWNLOAD_STARTED driver_id=%d technical_version_id=%s sequence=%d",
                            worker.driver_id, task.technical_version_id, task.sequence_priority)
                try:
                    result = self._download(worker.webdriver, task)
                except BaseException as error:
                    if task.attempt + 1 < self._max_attempts and not self._stop.is_set():
                        worker.state = DriverState.RECOVERING
                        logger.warning("WEBDRIVER_RECOVERY driver_id=%d reason=%s", worker.driver_id, error)
                        try:
                            if self._recover:
                                self._recover(worker.webdriver, error)
                            retry = DownloadTask(task.workbook_identity, task.technical_version_id,
                                                 task.version_label, task.sequence_priority, task.url,
                                                 task.payload, task.attempt + 1)
                            with self._lock:
                                self.download_state[task.technical_version_id] = DownloadState.QUEUED
                                self._serial += 1
                                self._queue.put((retry.sequence_priority, self._serial, retry, future))
                        except BaseException as recovery_error:
                            with self._lock:
                                self.download_state[task.technical_version_id] = DownloadState.FAILED
                            future.set_exception(recovery_error)
                    else:
                        with self._lock:
                            self.download_state[task.technical_version_id] = DownloadState.FAILED
                        future.set_exception(error)
                        logger.error("WEBDRIVER_FAILED driver_id=%d reason=%s", worker.driver_id, error)
                else:
                    elapsed = time.monotonic() - started
                    worker.downloads += 1
                    worker.busy_time += elapsed
                    with self._lock:
                        self.download_state[task.technical_version_id] = DownloadState.READY
                    future.set_result(result)
                    logger.info("WEBDRIVER_DOWNLOAD_FINISHED driver_id=%d technical_version_id=%s duration=%.3f",
                                worker.driver_id, task.technical_version_id, elapsed)
                finally:
                    worker.last_activity = time.monotonic()
                    worker.current_version = None
                    self._queue.task_done()
        except BaseException as error:
            worker.state = DriverState.FAILED
            worker.health = str(error)
            logger.exception("WEBDRIVER_FAILED driver_id=%d reason=%s", worker.driver_id, error)
        finally:
            worker.state = DriverState.STOPPED

    def close(self) -> None:
        self._stop.set()
        for worker in self.workers:
            if worker.owner_thread:
                worker.owner_thread.join(timeout=5)
            try:
                if self._close_driver is not None:
                    self._close_driver(worker.webdriver)
                else:
                    quit_driver = getattr(worker.webdriver, "quit", None)
                    if callable(quit_driver):
                        quit_driver()
            except Exception:
                logger.warning("Falha ao encerrar driver_id=%d", worker.driver_id, exc_info=True)
        elapsed = max(.001, max((time.monotonic() - w.started_at for w in self.workers), default=.001))
        logger.info("WEBDRIVER_POOL_SUMMARY driver_count_configured=%d drivers_created=%d "
                    "drivers_ready=%d drivers_failed=%d total_downloads=%d downloads_per_driver=%s "
                    "utilization_per_driver=%s aggregate_downloads_per_minute=%.3f",
                    len(self.workers), len(self.workers),
                    sum(w.health == "ready" for w in self.workers),
                    sum(w.health != "ready" for w in self.workers), sum(w.downloads for w in self.workers),
                    {w.driver_id: w.downloads for w in self.workers},
                    {w.driver_id: round(w.busy_time / elapsed, 4) for w in self.workers},
                    sum(w.downloads for w in self.workers) * 60 / elapsed)

    def forget(self, technical_version_id: str) -> None:
        """Remove resultado READY depois que o consumidor libera o arquivo."""
        with self._lock:
            self._futures.pop(technical_version_id, None)
            self.download_state.pop(technical_version_id, None)
