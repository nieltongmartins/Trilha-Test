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

    def configure(self, **kwargs: object) -> None:
        if "state" in kwargs:
            self.state = kwargs["state"]


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
        "Progresso da auditoria: 1/4 (25%) | Estimativa: 00:30 | "
        "Tempo total: 00:10"
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
