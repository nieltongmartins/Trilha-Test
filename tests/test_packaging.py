from pathlib import Path


ROOT = Path(__file__).parents[1]


def test_windows_package_is_one_folder_and_has_no_node_dependency() -> None:
    spec = (ROOT / "auditor_planilhas.spec").read_text(encoding="utf-8")
    build = (ROOT / "build_windows.bat").read_text(encoding="utf-8")

    assert "COLLECT(" in spec
    assert "console=False" in spec
    assert "requirements-build.txt" in build
    assert "PyInstaller" in build
    assert "node" not in (spec + build).lower()


def test_launcher_requires_local_non_versioned_configuration() -> None:
    launcher = (ROOT / "executar_auditor.bat").read_text(encoding="utf-8")
    example = (ROOT / "configuracao.exemplo.bat").read_text(encoding="utf-8")
    gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")

    assert 'cd /d "%~dp0"' in launcher
    assert 'if not exist "configuracao.bat"' in launcher
    assert 'call "configuracao.bat"' in launcher
    assert "SHAREPOINT_SITE_URL" in example
    assert "SHAREPOINT_SCOPE_PATHS" in example
    assert "configuracao.bat" in gitignore
    forbidden = ("PASSWORD", "TOKEN", "COOKIE", "CLIENT_SECRET")
    assert not any(f'set "{name}=' in example.upper() for name in forbidden)
