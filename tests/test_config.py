from pathlib import Path

import json

import pytest

from app.config import DEFAULT_SHAREPOINT_SITE_URL, Settings


def test_settings_create_required_directories(tmp_path: Path) -> None:
    settings = Settings(
        database_path=tmp_path / "database" / "audit.db",
        log_path=tmp_path / "logs" / "audit.log",
        temp_directory=tmp_path / "temp",
        reports_directory=tmp_path / "reports",
    )

    settings.create_directories()

    assert settings.database_path.parent.is_dir()
    assert settings.log_path.parent.is_dir()
    assert settings.temp_directory.is_dir()
    assert settings.reports_directory.is_dir()


def test_settings_start_without_sharepoint_environment(
    tmp_path: Path, monkeypatch
) -> None:
    monkeypatch.delenv("SHAREPOINT_SITE_URL", raising=False)
    monkeypatch.delenv("SHAREPOINT_SCOPE_PATHS", raising=False)

    settings = Settings.from_environment(config_path=tmp_path / "config.json")

    assert settings.sharepoint_site_url == DEFAULT_SHAREPOINT_SITE_URL
    assert settings.sharepoint_scope_paths == ()


def test_local_configuration_is_saved_and_loaded_without_credentials(
    tmp_path: Path, monkeypatch
) -> None:
    monkeypatch.delenv("SHAREPOINT_SITE_URL", raising=False)
    monkeypatch.delenv("SHAREPOINT_SCOPE_PATHS", raising=False)
    path = tmp_path / "profile" / "config.json"
    settings = Settings.from_environment(config_path=path)

    settings.save_browser_sharepoint(
        "https://tenant.sharepoint.com/site", ("/site/Documentos",)
    )
    restarted = Settings.from_environment(config_path=path)

    assert restarted.sharepoint_site_url == "https://tenant.sharepoint.com/site"
    assert restarted.sharepoint_scope_paths == ("/site/Documentos",)
    assert set(json.loads(path.read_text(encoding="utf-8"))) == {
        "sharepoint_site_url",
        "sharepoint_scope_paths",
    }


def test_environment_has_precedence_over_local_configuration(
    tmp_path: Path, monkeypatch
) -> None:
    path = tmp_path / "config.json"
    Settings.from_environment(config_path=path).save_browser_sharepoint(
        "https://local.sharepoint.com/site", ("/local/documentos",)
    )
    monkeypatch.setenv("SHAREPOINT_SITE_URL", "https://env.sharepoint.com/site")
    monkeypatch.setenv("SHAREPOINT_SCOPE_PATHS", "/env/a;/env/b")

    settings = Settings.from_environment(config_path=path)

    assert settings.sharepoint_site_url == "https://env.sharepoint.com/site"
    assert settings.sharepoint_scope_paths == ("/env/a", "/env/b")


@pytest.mark.parametrize(
    ("site_url", "scopes", "message"),
    [
        ("http://inseguro", ("/documentos",), "HTTPS"),
        ("https://tenant.sharepoint.com", (), "ao menos um"),
        ("https://tenant.sharepoint.com", ("sem-barra",), "começar com '/'"),
    ],
)
def test_invalid_configuration_has_clear_message(
    tmp_path: Path, site_url: str, scopes: tuple[str, ...], message: str
) -> None:
    settings = Settings.from_environment(config_path=tmp_path / "config.json")

    with pytest.raises(ValueError, match=message):
        settings.save_browser_sharepoint(site_url, scopes)
