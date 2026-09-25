import pytest

from app.execution_timing import SharedExecutionTimingModel, TimedStage
from app.slot_visual_progress import SlotVisualProgress, TaskProgressPlan


def _plan(**seconds):
    baseline = {TimedStage[name]: value for name, value in seconds.items()}
    return TaskProgressPlan.snapshot(SharedExecutionTimingModel(baseline=baseline))


def test_dynamic_ranges_and_controlled_progress_each_second():
    plan = _plan(DOWNLOAD_TRANSFER=5, SHA=.01, READ_XLSX=20, COMPARE=4, STAGING=.01)
    expected_download_end = 99 * 5 / 29.02
    assert plan.ranges[TimedStage.DOWNLOAD_TRANSFER].end == pytest.approx(expected_download_end)

    visual = SlotVisualProgress("task", plan, TimedStage.DOWNLOAD_TRANSFER, 0)
    download = [visual.tick(second) for second in range(1, 6)]
    visual.enter_phase(TimedStage.READ_XLSX, 5)
    read = [visual.tick(second) for second in range(6, 26)]
    visual.enter_phase(TimedStage.COMPARE, 25)
    compare = [visual.tick(second) for second in range(26, 29)]

    assert len(set(download)) == 5
    assert len(set(read)) == 20
    assert len(set(compare)) == 3
    assert download + read + compare == sorted(download + read + compare)


def test_slow_phase_has_asymptotic_tail_and_only_event_completes_range():
    plan = _plan(READ_XLSX=10)
    visual = SlotVisualProgress("task", plan, TimedStage.READ_XLSX, 0)
    values = [visual.tick(second) for second in (5, 10, 20, 29)]
    phase_range = plan.ranges[TimedStage.READ_XLSX]
    local = [(value - phase_range.start) / (phase_range.end - phase_range.start)
             for value in values]
    assert local[0] == pytest.approx(.425)
    assert local[1] == pytest.approx(.85)
    assert local[1] < local[2] < local[3] < .99
    visual.enter_phase(TimedStage.COMPARE, 30)
    assert visual.last_displayed_progress == pytest.approx(phase_range.end)


def test_fast_phase_finishes_on_real_transition_and_plan_is_frozen():
    model = SharedExecutionTimingModel(baseline={TimedStage.READ_XLSX: 30})
    plan_a = TaskProgressPlan.snapshot(model)
    visual = SlotVisualProgress("task", plan_a, TimedStage.READ_XLSX, 0)
    running = visual.tick(5)
    assert running < plan_a.ranges[TimedStage.READ_XLSX].end

    model.observe(TimedStage.READ_XLSX, 60)
    assert visual.plan.ranges[TimedStage.READ_XLSX].estimate == 30
    plan_b = TaskProgressPlan.snapshot(model)
    assert plan_b.ranges[TimedStage.READ_XLSX].estimate == 60
    visual.enter_phase(TimedStage.COMPARE, 5)
    assert visual.last_displayed_progress == pytest.approx(
        plan_a.ranges[TimedStage.READ_XLSX].end
    )


def test_wait_promotion_stays_below_complete_and_progress_never_regresses():
    plan = _plan(STAGING=.1)
    visual = SlotVisualProgress("task", plan, TimedStage.STAGING, 0)
    before = visual.tick(100)
    visual.enter_phase(TimedStage.WAIT_PROMOTION, 100)
    waiting = visual.tick(1000)
    assert before <= waiting < 100
    visual.complete = True
    assert visual.tick(1001) == 100
    assert visual.tick(999) == 100
