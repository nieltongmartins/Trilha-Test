import logging
from pathlib import Path

import main


def test_application_starts_and_creates_database(tmp_path: Path, monkeypatch) -> None:
    database_path = tmp_path / "database" / "audit.db"
    log_path = tmp_path / "logs" / "audit.log"
    monkeypatch.setenv("AUDIT_DATABASE_PATH", str(database_path))
    monkeypatch.setenv("AUDIT_LOG_PATH", str(log_path))

    assert main.main(launch_ui=False) == 0
    assert database_path.is_file()
    assert log_path.is_file()


def test_ui_is_created_without_opening_sharepoint(tmp_path: Path, monkeypatch) -> None:
    events: list[str] = []

    class RootFake:
        def __init__(self) -> None:
            self.destroyed = False
            self.protocols = {}

        def title(self, _title: str) -> None:
            events.append("tk")

        def minsize(self, *_size: int) -> None:
            pass

        def mainloop(self) -> None:
            events.append("mainloop")

        def protocol(self, name, callback) -> None:
            self.protocols[name] = callback

        def bind(self, *_args, **_kwargs) -> None:
            pass

        def destroy(self) -> None:
            self.destroyed = True

        def winfo_exists(self) -> int:
            return int(not self.destroyed)

        def winfo_viewable(self) -> int:
            return int(not self.destroyed)

        def state(self) -> str:
            return "normal" if not self.destroyed else "indisponível"

    class ApplicationFake:
        def __init__(self, _root, _database, source, _reports, **_kwargs) -> None:
            assert source is None
            events.append("ui")

        def close_source(self) -> None:
            pass

    monkeypatch.setenv("AUDIT_DATABASE_PATH", str(tmp_path / "audit.db"))
    monkeypatch.setenv("AUDIT_LOG_PATH", str(tmp_path / "audit.log"))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "config"))
    monkeypatch.delenv("SHAREPOINT_SITE_URL", raising=False)
    monkeypatch.delenv("SHAREPOINT_SCOPE_PATHS", raising=False)
    monkeypatch.setattr(main.tk, "Tk", RootFake)
    monkeypatch.setattr(main, "AuditApplication", ApplicationFake)
    monkeypatch.setattr(
        main.BrowserSharePointSource,
        "open_edge",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("Edge não deve abrir durante a inicialização")
        ),
    )

    assert main.main(launch_ui=True) == 0
    assert events == ["tk", "ui", "mainloop"]


def test_mainloop_base_exception_is_logged_and_propagated(
    tmp_path: Path, monkeypatch
) -> None:
    class InterruptingRoot:
        def title(self, _title: str) -> None:
            pass

        def minsize(self, *_size: int) -> None:
            pass

        def mainloop(self) -> None:
            raise KeyboardInterrupt("interrupção de diagnóstico")

        def protocol(self, *_args) -> None:
            pass

        def bind(self, *_args, **_kwargs) -> None:
            pass

        def winfo_exists(self) -> int:
            return 1

        def winfo_viewable(self) -> int:
            return 1

        def state(self) -> str:
            return "normal"

    class ApplicationFake:
        def __init__(self, *_args, **_kwargs) -> None:
            pass

        def close_source(self) -> None:
            pass

    monkeypatch.setenv("AUDIT_DATABASE_PATH", str(tmp_path / "audit.db"))
    log_path = tmp_path / "audit.log"
    monkeypatch.setenv("AUDIT_LOG_PATH", str(log_path))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "config"))
    monkeypatch.setattr(main.tk, "Tk", InterruptingRoot)
    monkeypatch.setattr(main, "AuditApplication", ApplicationFake)
    try:
        main.main(launch_ui=True)
    except KeyboardInterrupt as error:
        assert repr(error) == "KeyboardInterrupt('interrupção de diagnóstico')"
    else:
        raise AssertionError("KeyboardInterrupt deveria ser propagado")

    log_text = log_path.read_text(encoding="utf-8")
    assert "root.mainloop terminou por BaseException" in log_text
    assert "tipo_erro=KeyboardInterrupt" in log_text


class LifecycleRootFake:
    def __init__(self) -> None:
        self.protocols = {}
        self.bindings = {}
        self.destroyed = False

    def protocol(self, name, callback) -> None:
        self.protocols[name] = callback

    def bind(self, sequence, callback, add=None) -> None:
        self.bindings[sequence] = (callback, add)

    def destroy(self) -> None:
        self.destroyed = True

    def winfo_exists(self) -> int:
        return int(not self.destroyed)

    def winfo_viewable(self) -> int:
        return int(not self.destroyed)

    def state(self) -> str:
        return "normal" if not self.destroyed else "indisponível"


def test_window_lifecycle_only_closes_on_wm_delete(caplog) -> None:
    caplog.set_level(logging.INFO)
    root = LifecycleRootFake()
    lifecycle = main.WindowLifecycle(root, logging.getLogger("test.lifecycle"))

    assert root.destroyed is False
    assert lifecycle.close_requested is False
    root.protocols["WM_DELETE_WINDOW"]()

    assert root.destroyed is True
    assert lifecycle.close_requested is True
    assert "WM_DELETE_WINDOW solicitado pelo usuário" in caplog.text
    assert "destroy solicitado" in caplog.text


def test_window_lifecycle_records_callback_exception(caplog) -> None:
    root = LifecycleRootFake()
    lifecycle = main.WindowLifecycle(root, logging.getLogger("test.lifecycle"))

    try:
        raise ValueError("callback inválido")
    except ValueError as error:
        lifecycle.report_callback_exception(type(error), error, error.__traceback__)

    assert root.destroyed is False
    assert "Exceção em callback Tkinter" in caplog.text
    assert "ValueError: callback inválido" in caplog.text
