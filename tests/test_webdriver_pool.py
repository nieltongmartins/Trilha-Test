from __future__ import annotations

import threading
import time

from app.webdriver_pool import DownloadState, DownloadTask, WebDriverPool


class Driver:
    def __init__(self) -> None:
        self.active = 0
        self.maximum = 0
        self.threads: set[int] = set()

    def quit(self) -> None:
        pass


def task(identifier: str, priority: int = 1) -> DownloadTask[str]:
    return DownloadTask("book", identifier, identifier, priority, "/file", identifier)


def test_pool_deduplicates_and_never_calls_one_driver_concurrently() -> None:
    drivers = [Driver(), Driver(), Driver()]
    calls: list[str] = []

    def download(driver: Driver, item: DownloadTask[str]) -> str:
        driver.active += 1
        driver.maximum = max(driver.maximum, driver.active)
        driver.threads.add(threading.get_ident())
        calls.append(item.technical_version_id)
        time.sleep(.02)
        driver.active -= 1
        return item.payload

    pool = WebDriverPool(drivers, download)
    try:
        duplicate_a = pool.submit(task("10"))
        duplicate_b = pool.submit(task("10"))
        results = [pool.submit(task(str(value))).result(timeout=2) for value in range(11, 15)]
        assert duplicate_a is duplicate_b
        assert duplicate_a.result(timeout=2) == "10"
        assert results == [str(value) for value in range(11, 15)]
        assert calls.count("10") == 1
        assert pool.download_state["10"] is DownloadState.READY
        assert all(driver.maximum == 1 and len(driver.threads) == 1 for driver in drivers)
    finally:
        pool.close()


def test_failed_driver_task_is_requeued_without_duplicate_future() -> None:
    attempts = 0

    def flaky(_driver: Driver, item: DownloadTask[str]) -> str:
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            raise RuntimeError("janela perdida")
        return item.payload

    pool = WebDriverPool([Driver(), Driver(), Driver()], flaky, recover=lambda *_: None)
    try:
        future = pool.submit(task("20"))
        assert future.result(timeout=2) == "20"
        assert attempts == 2
        assert pool.download_state["20"] is DownloadState.READY
    finally:
        pool.close()


def test_shared_queue_dispatches_nearest_priority_after_variable_latency() -> None:
    release_slow = threading.Event()
    started = threading.Event()
    order: list[str] = []

    def variable(_driver: Driver, item: DownloadTask[str]) -> str:
        order.append(item.technical_version_id)
        if item.technical_version_id == "slow":
            started.set()
            release_slow.wait(timeout=2)
        return item.payload

    pool = WebDriverPool([Driver()], variable)
    try:
        slow = pool.submit(task("slow", 1))
        assert started.wait(timeout=2)
        far = pool.submit(task("far", 40))
        nearest = pool.submit(task("nearest", 10))
        duplicate = pool.submit(task("nearest", 0))
        release_slow.set()
        assert slow.result(timeout=2) == "slow"
        assert nearest.result(timeout=2) == "nearest"
        assert far.result(timeout=2) == "far"
        assert duplicate is nearest
        assert order == ["slow", "nearest", "far"]
    finally:
        pool.close()
