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
    def configure(self, **_kwargs: object) -> None:
        pass


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
