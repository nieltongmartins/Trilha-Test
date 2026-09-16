from pathlib import Path

from app.interface import AuditApplication
from app.sources import SharePointReadError, SpreadsheetInfo


class VariableFake:
    def __init__(self) -> None:
        self.value = ""

    def set(self, value: str) -> None:
        self.value = value


class SelectorFake:
    def current(self) -> int:
        return 0


class ButtonFake:
    def configure(self, **_kwargs: object) -> None:
        raise AssertionError("o botão não deve ser alterado depois da falha")


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

    application.show_status()

    assert "indisponível" in application.details.value
    assert application.status.value == (
        "Falha ao consultar versões: resposta REST incompatível"
    )


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

    AuditApplication(object(), object(), object(), Path("relatorios"))

    assert calls == []
