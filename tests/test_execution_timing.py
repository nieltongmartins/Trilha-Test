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
    assert model.estimate_task(TimedStage.STAGING, 200).progress <= 99


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
    samples_before = model.samples(TimedStage.READ_XLSX)
    model.continue_execution(200)
    assert model.active_now(210) == pytest.approx(130)
    assert model.samples(TimedStage.READ_XLSX) == samples_before


def test_throughput_excludes_pause_and_samples_survive_pause():
    model = SharedExecutionTimingModel()
    model.observe(TimedStage.TOTAL_TASK, 10, slot_id=1)
    model.record_commit(100)
    model.pause(105)
    model.resume(165)
    model.record_commit(170)
    assert model.throughput_per_minute() == pytest.approx(6)
    assert model.task_average() == 10


@pytest.mark.parametrize("fast,slow", [(0.01, 100.0), (0.1, 1000.0)])
def test_very_fast_and_slow_slots_share_a_bounded_estimate(fast, slow):
    model = SharedExecutionTimingModel()
    for value in (1, 1.1, 1.2, fast, slow):
        model.observe(TimedStage.COMPARE, value)
    assert fast <= model.average(TimedStage.COMPARE) < slow


def test_controlled_stage_samples_advance_inside_download_and_read():
    model = SharedExecutionTimingModel(baseline={
        TimedStage.DOWNLOAD_FETCH: .001,
        TimedStage.DOWNLOAD_TRANSFER: 2,
        TimedStage.SHA: .001,
        TimedStage.READ_XLSX: 4,
        TimedStage.COMPARE: 1,
        TimedStage.STAGING: .001,
    })
    download = [model.estimate_task(TimedStage.DOWNLOAD_TRANSFER, t).progress
                for t in (.5, 1, 1.5, 2)]
    assert download == sorted(download)
    assert len(set(download)) == 4
    completed = (TimedStage.DOWNLOAD_FETCH, TimedStage.DOWNLOAD_TRANSFER,
                 TimedStage.SHA)
    reading = [model.estimate_task(TimedStage.READ_XLSX, t, completed).progress
               for t in (1, 2, 3)]
    assert reading == sorted(reading)
    assert download[-1] < reading[0] < reading[-1] < 100


def test_slow_stage_keeps_moving_fast_stage_finishes_only_on_fact_and_mean_cannot_regress():
    model = SharedExecutionTimingModel(baseline={TimedStage.READ_XLSX: 2})
    slow = [model.estimate_task(TimedStage.READ_XLSX, t).progress for t in (2, 4, 8)]
    assert slow[0] < slow[1] < slow[2] < 100
    fast_running = model.estimate_task(TimedStage.READ_XLSX, 1)
    assert fast_running.progress < 100
    assert model.estimate_task(TimedStage.READ_XLSX, 1, finished=True).progress == 100
    displayed = fast_running.progress
    model.observe(TimedStage.READ_XLSX, 10)
    recalculated = model.estimate_task(TimedStage.READ_XLSX, 1).progress
    assert max(displayed, recalculated) >= displayed
