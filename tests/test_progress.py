"""Testes do progresso preditivo, sem dependência do Tk."""

import threading

import pytest

from app.progress import SmoothVersionProgress


def test_advances_between_events_and_real_event_forces_milestone() -> None:
    progress = SmoothVersionProgress()
    progress.observe("2.1", 0, 0)
    progress.observe("2.1", 5, 0)
    assert 5 < progress.estimate(2) < 55
    assert progress.observe("2.1", 55, 2.1) == 55


def test_never_regresses_or_crosses_unconfirmed_limit() -> None:
    progress = SmoothVersionProgress()
    progress.observe("2.1", 0, 0)
    progress.observe("2.1", 5, 0)
    later = progress.estimate(8)
    assert progress.estimate(3) == later
    assert progress.estimate(10_000) == pytest.approx(54.5)


def test_slow_stage_keeps_moving_after_expected_time() -> None:
    progress = SmoothVersionProgress()
    progress.observe("2.1", 0, 0)
    progress.observe("2.1", 5, 0)
    at_expected = progress.estimate(10)
    assert at_expected < progress.estimate(20) < 55


def test_retry_does_not_reset_and_next_version_does() -> None:
    progress = SmoothVersionProgress()
    progress.observe("2.1", 0, 0)
    progress.observe("2.1", 5, 0)
    value = progress.estimate(4)
    # Mensagens de retry não são marcos e, portanto, não alteram o estado.
    assert progress.version == "2.1"
    assert progress.value == value
    assert progress.observe("2.2", 0, 5) == 0


def test_timeout_cannot_complete_and_only_checkpoint_reaches_100() -> None:
    progress = SmoothVersionProgress()
    progress.observe("2.1", 0, 0)
    progress.observe("2.1", 97, 1)
    assert progress.estimate(10_000) == pytest.approx(99.5)
    assert progress.observe("2.1", 100, 10_001) == 100


def test_median_resists_extreme_timeout_and_eta_remains_coherent() -> None:
    progress = SmoothVersionProgress()
    progress.histories[5.0].extend([8, 9, 10, 11, 10_000])
    progress.total_history.extend([18, 20, 22, 21, 50_000])
    assert progress.phase_average(5.0) == 10
    assert progress.total_average() == 21


def test_visual_updates_can_be_confined_to_main_thread() -> None:
    progress = SmoothVersionProgress()
    main_thread = threading.get_ident()
    progress.observe("2.1", 0, 0)
    progress.observe("2.1", 5, 0)
    assert threading.get_ident() == main_thread
    assert progress.estimate(1) > 5
