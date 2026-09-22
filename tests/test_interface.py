from pathlib import Path
import queue
import threading

from app.interface import AuditApplication
from app.sources import SharePointReadError, SpreadsheetInfo


class VariableFake:
    def __init__(self) -> None:
        self.value = ""
        self.set_threads: list[int] = []

    def set(self, value: str) -> None:
        self.value = value
        self.set_threads.append(threading.get_ident())

    def get(self) -> str:
        return self.value


class SelectorFake:
    def __init__(self) -> None:
        self.index = 0
        self.values = []

    def current(self, index: int | None = None) -> int:
        if index is not None:
            self.index = index
        return self.index

    def __setitem__(self, key: str, value: object) -> None:
        assert key == "values"
        self.values = value


class ButtonFake:
    def __init__(self) -> None:
        self.state = "normal"
        self.text = ""

    def configure(self, **kwargs: object) -> None:
        if "state" in kwargs:
            self.state = kwargs["state"]
        if "text" in kwargs:
            self.text = kwargs["text"]


class ProgressbarFake:
    def __init__(self) -> None:
        self.mode = "determinate"
        self.running = False

    def configure(self, **kwargs: object) -> None:
        if "mode" in kwargs:
            self.mode = str(kwargs["mode"])

    def start(self, _interval: int) -> None:
        self.running = True

    def stop(self) -> None:
        self.running = False


def test_version_discovery_is_indeterminate_and_reports_only_received_items() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application._version_scan_updates = queue.SimpleQueue()
    application._version_scan_started_at = None
    application._version_scan_active = False
    application.progress_bar = ProgressbarFake()
    application.progress_value = VariableFake()
    application.progress_text = VariableFake()

    application._begin_version_scan("5.129")
    assert application.progress_bar.mode == "indeterminate"
    assert application.progress_bar.running is True
    assert "checkpoint 5.129" in application.progress_text.value
    assert "Versões encontradas: 0" in application.progress_text.value

    application._version_scan_updates.put(1000)
    application._version_scan_updates.put(2000)
    application._poll_version_scan_updates()

    assert "Versões encontradas: 2,000" in application.progress_text.value
    assert application.progress_text.set_threads[-1] == threading.get_ident()


def test_pause_continue_and_stop_buttons_update_control_events() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application._audit_active = True
    application._audit_paused = False
    application._pause_event = threading.Event()
    application._stop_event = threading.Event()
    application.pause_button = ButtonFake()
    application.stop_button = ButtonFake()
    application.status = VariableFake()

    application.toggle_pause()
    assert application._pause_event.is_set()
    assert "Finalizando a versão atual" in application.status.value

    application._control_updates = queue.SimpleQueue()
    application._control_updates.put(("paused", "4.365"))
    application._poll_control_updates()
    assert application._audit_paused is True
    assert application.pause_button.text == "Continuar"
    assert application.status.value == "Auditoria pausada no checkpoint 4.365."

    application.toggle_pause()
    assert not application._pause_event.is_set()
    assert application.pause_button.text == "Pausar"
    assert application.status.value == "Auditoria retomada."

    application.stop_audit()
    assert application._stop_event.is_set()
    assert application.pause_button.state == "disabled"
    assert application.stop_button.state == "disabled"


class SchedulerFake:
    def __init__(self) -> None:
        self.callbacks: list[tuple[object, tuple[object, ...]]] = []
        self.destroyed = False
        self.after_threads: list[int] = []

    def winfo_exists(self) -> int:
        return 0 if self.destroyed else 1

    def winfo_viewable(self) -> int:
        return 0 if self.destroyed else 1

    def state(self) -> str:
        return "normal"

    def after(self, _delay: int, callback, *args: object) -> None:
        self.after_threads.append(threading.get_ident())
        self.callbacks.append((callback, args))

    def run_next(self) -> None:
        callback, args = self.callbacks.pop(0)
        callback(*args)

    def destroy(self) -> None:
        self.destroyed = True


def test_close_is_blocked_while_background_operation_is_active(monkeypatch) -> None:
    application = AuditApplication.__new__(AuditApplication)
    scheduler = SchedulerFake()
    application._busy = True
    application.winfo_toplevel = lambda: scheduler
    warnings: list[tuple[str, str]] = []
    monkeypatch.setattr(
        "app.interface.messagebox.showwarning",
        lambda title, message: warnings.append((title, message)),
    )

    application.request_close()

    assert scheduler.destroyed is False
    assert warnings and "Aguarde" in warnings[0][1]


def test_close_destroys_window_after_background_operation_finishes() -> None:
    application = AuditApplication.__new__(AuditApplication)
    scheduler = SchedulerFake()
    application._busy = False
    application.winfo_toplevel = lambda: scheduler

    application.request_close()

    assert scheduler.destroyed is True


def _application_for_connection(
    connect_source, save_configuration=lambda *_args: None
) -> tuple[AuditApplication, SchedulerFake]:
    application = AuditApplication.__new__(AuditApplication)
    scheduler = SchedulerFake()
    application.after = scheduler.after
    application.winfo_toplevel = lambda: scheduler
    application.connect_source = connect_source
    application.save_configuration = save_configuration
    application.site_url = VariableFake()
    application.site_url.value = "https://tenant.sharepoint.com/site"
    application.scope_paths = VariableFake()
    application.scope_paths.value = "/site/Documentos"
    application.folder_names = VariableFake()
    application.folder_paths = []
    application.folder_selector = SelectorFake()
    application.status = VariableFake()
    application.source = None
    application._busy = False
    application._work_results = queue.SimpleQueue()
    application._progress_updates = queue.SimpleQueue()
    application._version_scan_updates = queue.SimpleQueue()
    application._version_scan_started_at = None
    application._version_scan_active = False
    application._version_cache = {}
    application._version_cache_ttl = 300.0
    application.connect_button = ButtonFake()
    application.refresh_button = ButtonFake()
    application.audit_button = ButtonFake()
    application.report_button = ButtonFake()
    return application, scheduler


def test_selected_folder_is_copied_to_scope_and_active_source() -> None:
    configured: list[tuple[str, ...]] = []

    class Source:
        def set_scope_paths(self, scopes: tuple[str, ...]) -> None:
            configured.append(scopes)

    application, _ = _application_for_connection(lambda *_args: None)
    application.source = Source()
    application.folder_paths = ["/site/Documentos/Qualidade"]

    application.copy_folder_to_scope()

    assert application.scope_paths.value == "/site/Documentos/Qualidade"
    assert configured == [("/site/Documentos/Qualidade",)]
    assert application.status.value == "Escopo atualizado. Clique em Atualizar lista."


def test_folder_selection_does_not_change_scope_before_copy() -> None:
    application, _ = _application_for_connection(lambda *_args: None)
    application.folder_names.value = "Qualidade"

    assert application._configured_values() == (
        "https://tenant.sharepoint.com/site",
        ("/site/Documentos",),
    )


def test_loaded_folders_populate_readonly_selection() -> None:
    application, _ = _application_for_connection(lambda *_args: None)

    application._folders_loaded(
        [
            ("Financeiro", "/site/Documentos/Financeiro"),
            ("Qualidade", "/site/Documentos/Qualidade"),
        ]
    )

    assert application.folder_selector.values == ["Financeiro", "Qualidade"]
    assert application.folder_paths == [
        "/site/Documentos/Financeiro",
        "/site/Documentos/Qualidade",
    ]


def test_connect_keeps_interface_alive_and_finishes_on_tk_thread() -> None:
    worker_started = threading.Event()
    release_worker = threading.Event()
    operation_threads: list[int] = []
    source = object()

    def connect_source(*_args):
        operation_threads.append(threading.get_ident())
        worker_started.set()
        assert release_worker.wait(timeout=1)
        return source

    application, scheduler = _application_for_connection(connect_source)
    main_thread = threading.get_ident()

    application.connect()
    assert worker_started.wait(timeout=1)
    assert operation_threads != [main_thread]
    # O callback agendado roda enquanto a conexão está deliberadamente parada:
    # isso representa o mainloop continuando a despachar eventos/redesenhos.
    scheduler.run_next()
    assert scheduler.callbacks
    assert application._busy is True
    assert scheduler.destroyed is False

    release_worker.set()
    outcome = application._work_results.get(timeout=1)
    application._work_results.put(outcome)
    scheduler.run_next()

    assert scheduler.destroyed is False
    assert scheduler.winfo_exists() == 1
    assert scheduler.winfo_viewable() == 1
    assert scheduler.state() == "normal"
    assert application.source is source
    assert application.status.value.startswith("Conectado ao SharePoint.")
    assert scheduler.after_threads == [main_thread, main_thread]
    assert application.status.set_threads
    assert set(application.status.set_threads) == {main_thread}


def test_connection_failure_keeps_interface_alive_and_reports_error() -> None:
    attempted = threading.Event()

    def connect_source(*_args):
        attempted.set()
        raise RuntimeError("Edge indisponível")

    application, scheduler = _application_for_connection(connect_source)

    application.connect()
    assert attempted.wait(timeout=1)
    outcome = application._work_results.get(timeout=1)
    application._work_results.put(outcome)
    scheduler.run_next()

    assert scheduler.destroyed is False
    assert scheduler.winfo_exists() == 1
    assert application.source is None
    assert application.status.value == "Falha na operação: Edge indisponível"
    assert set(application.status.set_threads) == {threading.get_ident()}


class FailingSource:
    def list_versions(self, _spreadsheet: SpreadsheetInfo):
        raise SharePointReadError("resposta REST incompatível")


class DatabaseFake:
    class Connection:
        def execute(self, _query: str, _parameters: object):
            class Result:
                @staticmethod
                def fetchone():
                    return None

            return Result()

    connection = Connection()


def test_versions_read_error_is_presented_without_closing_interface() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application.spreadsheets = [
        SpreadsheetInfo(
            "site", "sharepoint-rest", "uuid", "arquivo.xlsx", "/arquivo.xlsx"
        )
    ]
    application.selector = SelectorFake()
    application.source = FailingSource()
    application.database = DatabaseFake()
    application.details = VariableFake()
    application.status = VariableFake()
    application.audit_button = ButtonFake()
    application._version_cache = {}
    application._version_cache_ttl = 300.0
    application._version_scan_updates = queue.SimpleQueue()
    application._version_scan_active = False
    application._audit_started_at = None

    try:
        application._spreadsheet_status(application.spreadsheets[0])
    except SharePointReadError as error:
        application._work_failed(error)

    assert application.status.value == "Falha na operação: resposta REST incompatível"


def test_audit_progress_displays_percentage_estimate_and_elapsed(monkeypatch) -> None:
    application = AuditApplication.__new__(AuditApplication)
    application._audit_started_at = 90.0
    application._progress_completed = 0
    application._progress_total = 0
    application.progress_value = VariableFake()
    application.progress_text = VariableFake()
    monkeypatch.setattr("app.interface.time.monotonic", lambda: 100.0)

    application._update_progress(1, 4)

    assert application.progress_value.value == 25
    assert application.progress_text.value == (
        "1 / 4 (25%) | Estimativa: 00:30 | Tempo total: 00:10"
    )


def test_confirmed_checkpoints_update_details_only_when_polled_on_main_thread() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application.details = VariableFake()
    application._latest_available = "40.467"
    application._checkpoint_updates = queue.SimpleQueue()
    main_thread = threading.get_ident()

    def worker() -> None:
        application._checkpoint_updates.put(("2.18", 1, 19346))
        application._checkpoint_updates.put(("2.19", 2, 19345))

    thread = threading.Thread(target=worker)
    thread.start()
    thread.join()

    # A worker apenas enfileira: nenhuma variável Tk é tocada antes do polling.
    assert application.details.set_threads == []
    application._poll_checkpoint_updates()

    assert application.details.value == (
        "Última auditada: 2.19 | Última disponível: 40.467 | Pendentes: 19345"
    )
    assert application.details.set_threads == [main_thread, main_thread]


def test_checkpoint_update_keeps_latest_available_from_explicit_refresh() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application.details = VariableFake()
    application.audit_button = ButtonFake()
    application.status = VariableFake()
    application._version_scan_active = False

    application._show_status_finished(("2.17", "40.467", 19347, 19349))
    application._set_audit_details("2.18", 19346)

    assert application.details.value == (
        "Última auditada: 2.18 | Última disponível: 40.467 | Pendentes: 19346"
    )


def test_duration_uses_hours_only_when_needed() -> None:
    assert AuditApplication._format_duration(65) == "01:05"
    assert AuditApplication._format_duration(3661) == "01:01:01"


def test_complete_deletion_cancellation_does_not_touch_local_database(monkeypatch) -> None:
    application = AuditApplication.__new__(AuditApplication)
    calls = []
    application.storage = type("Storage", (), {"delete_all": lambda self: calls.append("delete")})()
    application._start_work = lambda *_args: calls.append("work")
    monkeypatch.setattr("app.interface.messagebox.askyesno", lambda *_args, **_kwargs: False)

    application.delete_all_audits()

    assert calls == []


def test_local_management_does_not_use_sharepoint_source() -> None:
    application = AuditApplication.__new__(AuditApplication)
    calls = []
    application.source = type("ForbiddenSource", (), {"list_spreadsheets": lambda self: calls.append("sharepoint")})()
    application.storage = type("Storage", (), {"list_audits": lambda self: []})()
    application.stored_tree = type("Tree", (), {"get_children": lambda self: (), "delete": lambda *_: None, "insert": lambda *_a, **_k: None})()

    application.refresh_stored()

    assert calls == []


def test_constructor_waits_for_manual_authentication_before_sharepoint_calls(
    monkeypatch,
) -> None:
    calls: list[str] = []

    monkeypatch.setattr(
        "app.interface.ttk.Frame.__init__", lambda self, master, padding: None
    )
    monkeypatch.setattr("app.interface.tk.StringVar", lambda **kwargs: VariableFake())
    monkeypatch.setattr(AuditApplication, "_build", lambda self: None)
    monkeypatch.setattr(
        AuditApplication, "refresh", lambda self: calls.append("refresh")
    )

    application = AuditApplication(
        object(),
        object(),
        None,
        Path("relatorios"),
        connect_source=lambda *_: calls.append("connect"),
    )

    assert calls == []
    assert application.source is None
    assert hasattr(application, "folder_names")


def test_stored_report_button_follows_selection() -> None:
    application = AuditApplication.__new__(AuditApplication)
    application._busy = False
    application.open_stored_report_button = ButtonFake()
    application.stored_tree = type(
        "Tree", (), {"selection": lambda self: ()}
    )()
    application._stored_selection_changed()
    assert application.open_stored_report_button.state == "disabled"

    application.stored_tree = type(
        "Tree", (), {"selection": lambda self: ("1",)}
    )()
    application._stored_selection_changed()
    assert application.open_stored_report_button.state == "normal"


def test_hidden_gesture_requires_five_clicks_and_does_not_duplicate(monkeypatch) -> None:
    application = AuditApplication.__new__(AuditApplication)
    application._hidden_clicks = []
    calls = []
    application._show_comparator = lambda: calls.append("show")
    ticks = iter((1.0, 1.1, 1.2, 1.3, 1.4, 1.5))
    monkeypatch.setattr("app.interface.time.monotonic", lambda: next(ticks))

    for _ in range(4):
        application._hidden_comparator_gesture()
    assert calls == []
    application._hidden_comparator_gesture()
    assert calls == ["show"]
    application._hidden_comparator_gesture()
    assert calls == ["show"]


def test_open_stored_report_opens_exact_existing_artifact(monkeypatch, tmp_path: Path) -> None:
    application = AuditApplication.__new__(AuditApplication)
    report = tmp_path / "exact.xlsx"
    report.write_bytes(b"xlsx")
    application._selected_stored = lambda: (7, "Audit")
    application.report_artifacts = type(
        "Artifacts", (), {"locate": lambda self, identifier: report if identifier == 7 else None}
    )()
    opened = []
    application._open_file = lambda path: opened.append(path)
    infos = []
    monkeypatch.setattr("app.interface.messagebox.showinfo", lambda *args: infos.append(args))

    application.open_stored_report()

    assert opened == [report]
    assert infos == []


def test_open_stored_report_does_not_generate_when_missing(monkeypatch) -> None:
    application = AuditApplication.__new__(AuditApplication)
    application._selected_stored = lambda: (7, "Audit")
    application.report_artifacts = type(
        "Artifacts", (), {"locate": lambda self, identifier: None}
    )()
    messages = []
    monkeypatch.setattr("app.interface.messagebox.showinfo", lambda *args: messages.append(args))

    application.open_stored_report()

    assert "Não existe relatório gerado" in messages[0][1]


def _application_for_version_progress() -> AuditApplication:
    application = AuditApplication.__new__(AuditApplication)
    application._version_progress_updates = queue.SimpleQueue()
    application.version_progress_value = VariableFake()
    application.version_stage_text = VariableFake()
    application.version_timing_text = VariableFake()
    application.current_version_frame = ButtonFake()
    application._current_version_started_at = None
    application._current_version = "—"
    from app.progress import SmoothVersionProgress
    application._smooth_version_progress = SmoothVersionProgress()
    application._version_animation_interval = 0.15
    application._version_last_animation = 0.0
    application._version_last_clock_update = 0.0
    application._progress_completed = 0
    application._progress_total = 30
    return application


def test_version_bar_starts_at_zero_advances_and_resets(monkeypatch) -> None:
    from app.audit_service import VersionProgress

    application = _application_for_version_progress()
    ticks = iter((10.0, 11.0, 12.0, 13.0, 14.0))
    monkeypatch.setattr("app.interface.time.monotonic", lambda: next(ticks))
    application._version_progress_updates.put(
        VersionProgress("2.18", 0, "Obtendo versão 2.18...", 10.0)
    )
    application._version_progress_updates.put(
        VersionProgress("2.18", 55, "Validando arquivo...", 11.0)
    )
    application._poll_version_progress_updates()
    assert 55 < application.version_progress_value.value < 70
    assert application._current_version == "2.18"

    application._version_progress_updates.put(
        VersionProgress("2.19", 0, "Obtendo versão 2.19...", 13.0)
    )
    application._poll_version_progress_updates()
    assert application.version_progress_value.value == 0
    assert application._current_version == "2.19"


def test_recent_average_keeps_only_twenty_and_eta_uses_it(monkeypatch) -> None:
    application = _application_for_version_progress()
    application._smooth_version_progress.total_history.extend(range(1, 22))
    application._progress_completed = 25
    application._progress_total = 30
    monkeypatch.setattr("app.interface.time.monotonic", lambda: 100.0)

    application._update_version_timing()

    assert list(application._smooth_version_progress.total_history) == list(range(2, 22))
    assert "Média recente: 00:12" in application.version_timing_text.value
    assert "Estimativa restante: 00:58" in application.version_timing_text.value


def test_version_progress_from_worker_only_touches_tk_during_main_poll(monkeypatch) -> None:
    from app.audit_service import VersionProgress

    application = _application_for_version_progress()
    monkeypatch.setattr("app.interface.time.monotonic", lambda: 10.0)
    main_thread = threading.get_ident()
    thread = threading.Thread(
        target=lambda: application._version_progress_updates.put(
            VersionProgress("2.18", 5, "Baixando dados...")
        )
    )
    thread.start()
    thread.join()
    assert application.version_stage_text.set_threads == []

    application._poll_version_progress_updates()

    assert application.version_stage_text.value == "Baixando dados..."
    assert set(application.version_stage_text.set_threads) == {main_thread}


def test_timeout_and_retry_message_keep_current_version(monkeypatch) -> None:
    application = _application_for_version_progress()
    application._current_version = "2.18"
    application._current_version_started_at = 1.0
    application._report_updates = queue.SimpleQueue()
    application.status = VariableFake()
    application._report_updates.put("SharePoint demorando para responder...")
    application._report_updates.put("Retry 1/1...")

    application._poll_report_updates()

    assert application._current_version == "2.18"
    assert application.version_stage_text.value == "Retry 1/1..."
