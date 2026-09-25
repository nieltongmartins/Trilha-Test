from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
import pickle
import threading
import time

import pytest

from app.version_snapshots import (
    SnapshotCache,
    SnapshotParseResult,
    SnapshotState,
    VersionSnapshot,
)


IDENTITY = "site/drive/workbook"


def parsed(task, *, delay=0.0):
    time.sleep(delay)
    snapshot = VersionSnapshot(task.workbook_identity, task.technical_version_id, {
        "Sheet": {"A1": task.technical_version_id, "B1": 0, "C1": False},
    })
    size = len(pickle.dumps(snapshot, protocol=5))
    return SnapshotParseResult(snapshot, delay, size, size, 0.0, 123)


def test_consecutive_pairs_parse_each_technical_version_once():
    cache = SnapshotCache(capacity=4)
    pairs = [("A", "B"), ("B", "C"), ("C", "D")]
    for sequence, pair in enumerate(pairs, 1):
        for technical_id in pair:
            cache.register((IDENTITY, technical_id), {sequence})

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = []
        for sequence, pair in enumerate(pairs, 1):
            futures.extend(cache.request(
                (IDENTITY, technical_id), f"{technical_id}.xlsx", sequence,
                executor, parsed,
            ) for technical_id in pair)
        assert all(future.result().snapshot for future in futures)

    summary = cache.summary()
    assert summary.unique_versions == 4
    assert summary.parse_count == 4
    assert summary.duplicate_parse_prevented == 2
    assert summary.parse_amplification == 1.0


def test_simultaneous_requests_share_the_same_parse_future():
    cache = SnapshotCache(capacity=2)
    key = (IDENTITY, "B")
    cache.register(key, {1, 2})
    barrier = threading.Barrier(2)

    with ThreadPoolExecutor(max_workers=3) as executor:
        def request(sequence):
            barrier.wait()
            return cache.request(key, "B.xlsx", sequence, executor,
                                 lambda task: parsed(task, delay=.03))

        callers = [executor.submit(request, sequence) for sequence in (1, 2)]
        first, second = (caller.result() for caller in callers)
        assert first is second
        assert first.result().snapshot["Sheet"]["B1"] == 0

    assert cache.parse_count == 1
    assert cache.reuse_count == 1


def test_snapshot_is_immutable_and_release_waits_for_every_consumer():
    cache = SnapshotCache(capacity=2)
    key = (IDENTITY, "B")
    cache.register(key, {1, 2})
    with ThreadPoolExecutor(max_workers=1) as executor:
        result = cache.request(key, "B.xlsx", 1, executor, parsed).result()
    with pytest.raises(TypeError):
        result.snapshot["Sheet"]["A1"] = "mutated"  # type: ignore[index]
    assert cache.release(key, 1) is False
    assert cache.state(key) is SnapshotState.READY
    assert cache.release(key, 2) is True
    assert cache.state(key) is SnapshotState.RELEASED
    assert cache.resident_count == 0


def test_failed_parse_never_publishes_partial_snapshot_and_can_be_reset():
    cache = SnapshotCache(capacity=2)
    key = (IDENTITY, "B")
    cache.register(key, {1})

    def crash(_task):
        raise SystemExit("worker morreu")

    with ThreadPoolExecutor(max_workers=1) as executor:
        future = cache.request(key, "B.xlsx", 1, executor, crash)
        with pytest.raises(SystemExit, match="worker morreu"):
            future.result()
        assert cache.state(key) is SnapshotState.FAILED
        cache.reset_failed(key)
        assert cache.request(key, "B.xlsx", 1, executor, parsed).result().snapshot

    assert cache.state(key) is SnapshotState.READY
    assert cache.parse_count == 2  # o segundo parse é um retry real e explícito


def test_sliding_window_never_exceeds_capacity():
    cache = SnapshotCache(capacity=3)
    with ThreadPoolExecutor(max_workers=1) as executor:
        for sequence in range(20):
            key = (IDENTITY, str(sequence))
            cache.register(key, {sequence})
            cache.request(key, f"{sequence}.xlsx", sequence, executor, parsed).result()
            cache.release(key, sequence)
    assert cache.resident_count == 0
    assert cache.summary().peak_cache_count <= cache.capacity


def test_result_waits_for_atomic_ready_publication(monkeypatch):
    """Future.done() must never be mistaken for registry publication."""
    cache = SnapshotCache(capacity=2)
    key = (IDENTITY, "B")
    cache.register(key, {1})
    publish_entered = threading.Event()
    permit_publish = threading.Event()
    original_publish = cache._publish

    def delayed_publish(cache_key, future):
        publish_entered.set()
        permit_publish.wait(1)
        original_publish(cache_key, future)

    monkeypatch.setattr(cache, "_publish", delayed_publish)
    with ThreadPoolExecutor(max_workers=2) as executor:
        future = cache.request(key, "B.xlsx", 1, executor, parsed)
        assert future.result().snapshot
        assert publish_entered.wait(1)
        consumer = executor.submit(cache.result, key)
        time.sleep(.02)
        assert not consumer.done()
        permit_publish.set()
        assert consumer.result(timeout=1).snapshot
    assert cache.state(key) is SnapshotState.READY


def test_ready_callbacks_are_event_driven_and_out_of_order():
    cache = SnapshotCache(capacity=4)
    awakened: list[str] = []
    cache.add_ready_callback(lambda key: awakened.append(key[1]))
    order = ("C", "A", "D", "B")
    with ThreadPoolExecutor(max_workers=1) as executor:
        for sequence, technical_id in enumerate(order, 1):
            key = (IDENTITY, technical_id)
            cache.register(key, {sequence})
            cache.request(key, f"{technical_id}.xlsx", sequence, executor, parsed).result()
        # Publication callback can trail Future.result(), so consume through the
        # atomic registry API before asserting the event order.
        for technical_id in order:
            cache.result((IDENTITY, technical_id))
    assert awakened == list(order)
    assert all(cache.is_ready((IDENTITY, technical_id)) for technical_id in order)


def test_twenty_thousand_backlog_materializes_only_active_window():
    cache = SnapshotCache(capacity=16)
    backlog = [(str(index), str(index + 1)) for index in range(20_000)]
    for sequence, pair in enumerate(backlog[:15], 1):
        for technical_id in pair:
            cache.register((IDENTITY, technical_id), {sequence})
    assert cache.summary().execution_unique_versions == 16
    assert cache.resident_count == 0
