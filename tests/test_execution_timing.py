from __future__ import annotations

import pytest

from app.execution_timing import SharedExecutionTimingModel, TimedStage


def test_bootstrap_first_sample_and_shared_slots():
    model = SharedExecutionTimingModel()
    assert not model.has_samples(TimedStage.READ_XLSX)
    assert model.average(TimedStage.READ_XLSX) > 0
    model.observe(TimedStage.READ_XLSX, 8, slot_id=1)
    model.observe(TimedStage.READ_XLSX, 10, slot_id=5)
    assert model.average(TimedStage.READ_XLSX) == 9
    assert {sample[2] for sample in model.raw_observations} == {1, 5}


def test_rolling_window_and_outlier_is_retained_but_robust():
    model = SharedExecutionTimingModel(window_size=3)
    for value in (1, 2, 3, 3600):
        model.observe(TimedStage.COMPARE, value)
    assert model.samples(TimedStage.COMPARE) == (2, 3, 3600)
    assert model.raw_observations[-1][1] == 3600
    assert model.average(TimedStage.COMPARE) < 10


def test_progress_is_continuous_asymptotic_and_only_finishes_on_fact():
    model = SharedExecutionTimingModel(baseline={
        TimedStage.DOWNLOAD_FETCH: .001,
        TimedStage.DOWNLOAD_TRANSFER: 5,
        TimedStage.SHA: .001,
        TimedStage.READ_XLSX: 10,
        TimedStage.COMPARE: 2,
        TimedStage.STAGING: .001,
    })
    early = model.estimate_task(TimedStage.DOWNLOAD_TRANSFER, 1)
    late = model.estimate_task(TimedStage.DOWNLOAD_TRANSFER, 50)
    assert 0 < early.progress < late.progress < 100
    parsing = model.estimate_task(
        TimedStage.READ_XLSX, 1,
        completed=(TimedStage.DOWNLOAD_FETCH, TimedStage.DOWNLOAD_TRANSFER, TimedStage.SHA),
    )
    assert parsing.progress > late.progress
    assert model.estimate_task(TimedStage.COMPARE, 200).progress < 100
    assert model.estimate_task(TimedStage.COMPARE, 200, finished=True).progress == 100


def test_individual_eta_excludes_completed_stages():
    model = SharedExecutionTimingModel()
    at_download = model.estimate_task(TimedStage.DOWNLOAD_TRANSFER, 0)
    at_compare = model.estimate_task(
        TimedStage.COMPARE, 0,
        completed=(TimedStage.DOWNLOAD_FETCH, TimedStage.DOWNLOAD_TRANSFER,
                   TimedStage.SHA, TimedStage.READ_XLSX),
    )
    assert at_compare.remaining < at_download.remaining


def test_global_eta_bootstraps_with_parallelism_then_uses_commit_throughput():
    model = SharedExecutionTimingModel(baseline={TimedStage.TOTAL_TASK: 20})
    assert model.global_eta(10, 5) == pytest.approx(40)
    for stamp in (0, 10, 20):
        model.record_commit(stamp)
    assert model.throughput_per_minute() == pytest.approx(6)
    assert model.global_eta(12, 1) == pytest.approx(120)


def test_pause_and_stop_freeze_visual_clock():
    model = SharedExecutionTimingModel()
    model.pause(100)
    assert model.active_now(150) == pytest.approx(100)
    model.resume(160)
    assert model.active_now(170) == pytest.approx(110)
    model.stop(180)
    assert model.active_now(999) == pytest.approx(120)


@pytest.mark.parametrize("fast,slow", [(0.01, 100.0), (0.1, 1000.0)])
def test_very_fast_and_slow_slots_share_a_bounded_estimate(fast, slow):
    model = SharedExecutionTimingModel()
    for value in (1, 1.1, 1.2, fast, slow):
        model.observe(TimedStage.COMPARE, value)
    assert fast <= model.average(TimedStage.COMPARE) < slow
