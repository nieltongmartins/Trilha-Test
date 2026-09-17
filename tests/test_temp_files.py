import json
from pathlib import Path

from app import temp_files
from app.temp_files import TemporaryWorkspace


def test_current_process_check_never_sends_signal(monkeypatch) -> None:
    monkeypatch.setattr(
        temp_files.os,
        "kill",
        lambda *_args: (_ for _ in ()).throw(AssertionError("os.kill foi chamado")),
    )

    assert temp_files._process_is_running(temp_files.os.getpid()) is True


def test_windows_process_check_does_not_use_os_kill(monkeypatch) -> None:
    checked: list[int] = []
    monkeypatch.setattr(temp_files.os, "name", "nt")
    monkeypatch.setattr(
        temp_files.os,
        "kill",
        lambda *_args: (_ for _ in ()).throw(AssertionError("os.kill foi chamado")),
    )
    monkeypatch.setattr(
        temp_files,
        "_windows_process_is_running",
        lambda pid: checked.append(pid) or True,
    )

    assert temp_files._process_is_running(987654) is True
    assert checked == [987654]


def test_workspace_isolated_names_do_not_collide_and_cleanup_is_scoped(
    tmp_path: Path,
) -> None:
    unrelated = tmp_path / "nao-remover.xlsx"
    unrelated.write_bytes(b"externo")
    first = TemporaryWorkspace(tmp_path)
    second = TemporaryWorkspace(tmp_path)
    first_path = first.filename("site", "drive", "planilha-a", "versao-1")
    other_version = first.filename("site", "drive", "planilha-a", "versao-2")
    other_sheet = first.filename("site", "drive", "planilha-b", "versao-1")
    same_identity_other_run = second.filename(
        "site", "drive", "planilha-a", "versao-1"
    )

    assert len({first_path, other_version, other_sheet}) == 3
    assert first_path.name == same_identity_other_run.name
    assert first_path.parent != same_identity_other_run.parent
    first_path.write_bytes(b"temporario")
    first.release(first_path)
    assert not first_path.exists()

    first_directory = first.path
    second_directory = second.path
    first.close()
    assert not first_directory.exists()
    assert second_directory.exists()
    assert unrelated.read_bytes() == b"externo"
    second.close()
    assert unrelated.read_bytes() == b"externo"


def test_restart_removes_only_recognized_orphan_workspace(tmp_path: Path) -> None:
    orphan = tmp_path / "auditoria-xlsx-orphan"
    orphan.mkdir()
    (orphan / ".auditoria-temp.json").write_text(
        json.dumps({"owner": "auditoria_excel", "pid": 999_999_999}),
        encoding="utf-8",
    )
    (orphan / "historico.xlsx").write_bytes(b"temporario")
    foreign = tmp_path / "auditoria-xlsx-foreign"
    foreign.mkdir()
    (foreign / ".auditoria-temp.json").write_text("{}", encoding="utf-8")

    workspace = TemporaryWorkspace(tmp_path)

    assert not orphan.exists()
    assert foreign.exists()
    workspace.close()
