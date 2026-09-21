from pathlib import Path
import subprocess
import sys

import main


def test_startup_imports_do_not_load_optional_heavy_modules() -> None:
    probe = subprocess.run(
        [
            sys.executable,
            "-c",
            "import main,sys; "
            "assert 'selenium' not in sys.modules; "
            "assert 'openpyxl' not in sys.modules; "
            "assert 'app.sources.graph' not in sys.modules; "
            "assert 'tools.benchmark_xlsx_reader' not in sys.modules",
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert probe.returncode == 0, probe.stderr


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
        def title(self, _title: str) -> None:
            events.append("tk")

        def minsize(self, *_size: int) -> None:
            pass

        def protocol(self, name: str, callback) -> None:
            assert name == "WM_DELETE_WINDOW"
            assert callable(callback)
            events.append("close-protocol")

        def mainloop(self) -> None:
            events.append("mainloop")

    class ApplicationFake:
        def __init__(self, _root, _database, source, _reports, **_kwargs) -> None:
            assert source is None
            events.append("ui")

        def close_source(self) -> None:
            pass

        def request_close(self) -> None:
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
    assert events == ["tk", "ui", "close-protocol", "mainloop"]
