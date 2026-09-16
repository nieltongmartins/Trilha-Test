"""Aquisição SharePoint REST dentro da sessão autenticada do Edge."""

from __future__ import annotations

import base64
from collections.abc import Mapping, Sequence
from pathlib import Path
import re
import tempfile
from typing import Any, Protocol
from urllib.parse import quote, urlsplit
import zipfile

from app.sources.base import SpreadsheetInfo, VersionInfo


class SharePointReadError(RuntimeError):
    """Falha de aquisição normalizada sem dados da sessão autenticada."""


class BrowserSession(Protocol):
    """Superfície mínima do WebDriver; não expõe cookies nem tokens."""

    def get(self, url: str) -> None: ...

    def execute_async_script(self, script: str, *args: object) -> object: ...

    def quit(self) -> None: ...


_FETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const binary = arguments[1];
fetch(url, {method: 'GET', credentials: 'same-origin', headers: {'Accept': 'application/json;odata=nometadata'}})
  .then(async response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    if (!binary) return {json: await response.json()};
    const bytes = new Uint8Array(await response.arrayBuffer());
    let text = '';
    for (let offset = 0; offset < bytes.length; offset += 32768) {
      text += String.fromCharCode(...bytes.subarray(offset, offset + 32768));
    }
    return {base64: btoa(text)};
  }).then(done).catch(error => done({error: String(error)}));
"""


def _odata_results(payload: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    value: object = payload.get("value")
    if value is None and isinstance(payload.get("d"), Mapping):
        value = payload["d"].get("results", payload["d"])
    if not isinstance(value, list):
        raise SharePointReadError("SharePoint REST retornou coleção inválida")
    return [item for item in value if isinstance(item, Mapping)]


def _escape_odata_path(path: str) -> str:
    """Escapa primeiro para literal OData e depois para uma URL segura."""
    return quote(path.replace("'", "''"), safe="/'()$=,:?&")


class BrowserSharePointSource:
    """Descobre e baixa versões via GET executado exclusivamente no Edge."""

    def __init__(
        self,
        site_url: str,
        scope_paths: Sequence[str],
        browser: BrowserSession,
        *,
        temp_directory: str | Path = "data/temp",
        owns_browser: bool = False,
    ) -> None:
        parsed = urlsplit(site_url)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("A URL do site SharePoint deve usar HTTPS")
        if not scope_paths:
            raise ValueError("Ao menos um escopo SharePoint deve ser configurado")
        self.site_url = site_url.rstrip("/")
        self._origin = f"{parsed.scheme}://{parsed.netloc}"
        self.scope_paths = tuple(dict.fromkeys(path.rstrip("/") for path in scope_paths))
        self._browser = browser
        self._owns_browser = owns_browser
        Path(temp_directory).mkdir(parents=True, exist_ok=True)
        self._temporary_directory = tempfile.TemporaryDirectory(
            prefix="sharepoint-browser-", dir=temp_directory
        )

    @classmethod
    def open_edge(
        cls, site_url: str, scope_paths: Sequence[str], *, temp_directory: str | Path = "data/temp"
    ) -> "BrowserSharePointSource":
        """Abre Edge visível; o usuário autentica normalmente no navegador."""
        from selenium import webdriver

        browser = webdriver.Edge()
        browser.get(site_url)
        return cls(site_url, scope_paths, browser, temp_directory=temp_directory, owns_browser=True)

    def close(self) -> None:
        self._temporary_directory.cleanup()
        if self._owns_browser:
            self._browser.quit()

    def __enter__(self) -> "BrowserSharePointSource":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def _endpoint(self, relative: str) -> str:
        url = f"{self.site_url}/_api/{relative.lstrip('/')}"
        parsed = urlsplit(url)
        if f"{parsed.scheme}://{parsed.netloc}" != self._origin or "/_api/" not in parsed.path:
            raise SharePointReadError("Endpoint fora do site ou da API REST permitida")
        return url

    def _fetch(self, url: str, *, binary: bool = False) -> object:
        parsed = urlsplit(url)
        if f"{parsed.scheme}://{parsed.netloc}" != self._origin or "/_api/" not in parsed.path:
            raise SharePointReadError("Somente endpoints GET da API REST configurada são permitidos")
        result = self._browser.execute_async_script(_FETCH_SCRIPT, url, binary)
        if not isinstance(result, Mapping) or result.get("error"):
            detail = result.get("error") if isinstance(result, Mapping) else "resposta inválida"
            raise SharePointReadError(f"Falha de leitura SharePoint REST: {detail}")
        key = "base64" if binary else "json"
        if key not in result:
            raise SharePointReadError("Resposta incompleta do navegador")
        return result[key]

    def _json(self, relative: str) -> Mapping[str, Any]:
        result = self._fetch(self._endpoint(relative))
        if not isinstance(result, Mapping):
            raise SharePointReadError("SharePoint REST retornou JSON inválido")
        return result

    def list_spreadsheets(self) -> tuple[SpreadsheetInfo, ...]:
        found: dict[str, SpreadsheetInfo] = {}
        pending = list(self.scope_paths)
        visited: set[str] = set()
        while pending:
            folder = pending.pop(0)
            if folder in visited:
                continue
            visited.add(folder)
            encoded = _escape_odata_path(folder)
            files = _odata_results(self._json(
                f"web/GetFolderByServerRelativeUrl('{encoded}')/Files"
                "?$select=Name,ServerRelativeUrl,UniqueId,Length"
            ))
            for item in files:
                name, path, unique_id = item.get("Name"), item.get("ServerRelativeUrl"), item.get("UniqueId")
                if not all(isinstance(value, str) for value in (name, path, unique_id)):
                    raise SharePointReadError("Arquivo retornado sem identidade REST comprovável")
                if name.lower().endswith(".xlsx"):
                    found[unique_id] = SpreadsheetInfo(
                        site_id=self.site_url,
                        drive_id=self.scope_paths[0],
                        drive_item_id=unique_id,
                        name=name,
                        path=path,
                        folder=path.rsplit("/", 1)[0],
                    )
            folders = _odata_results(self._json(
                f"web/GetFolderByServerRelativeUrl('{encoded}')/Folders?$select=Name,ServerRelativeUrl"
            ))
            pending.extend(
                path for item in folders
                if isinstance((path := item.get("ServerRelativeUrl")), str)
                and item.get("Name") != "Forms"
            )
        return tuple(sorted(found.values(), key=lambda item: item.path or ""))

    def list_versions(self, spreadsheet: SpreadsheetInfo) -> tuple[VersionInfo, ...]:
        self._validate_spreadsheet(spreadsheet)
        encoded = _escape_odata_path(spreadsheet.path or "")
        payload = self._json(
            f"web/GetFileByServerRelativeUrl('{encoded}')/Versions"
            "?$expand=CreatedBy&$select=ID,VersionLabel,Created,CreatedBy/Title,"
            "CreatedBy/Email,CreatedBy/LoginName,CheckInComment,Size,Url,IsCurrentVersion"
        )
        versions: list[VersionInfo] = []
        for item in _odata_results(payload):
            version_id, label = item.get("ID"), item.get("VersionLabel")
            if not isinstance(version_id, int) or not isinstance(label, str):
                raise SharePointReadError("Versão REST sem ID técnico ou VersionLabel")
            author = item.get("CreatedBy")
            versions.append(VersionInfo(
                id=str(version_id), number=label,
                modified_at=item.get("Created") if isinstance(item.get("Created"), str) else None,
                author=author.get("Title") if isinstance(author, Mapping) and isinstance(author.get("Title"), str) else None,
                author_email=author.get("Email") if isinstance(author, Mapping) and isinstance(author.get("Email"), str) else None,
                author_login=author.get("LoginName") if isinstance(author, Mapping) and isinstance(author.get("LoginName"), str) else None,
                comment=item.get("CheckInComment") if isinstance(item.get("CheckInComment"), str) else None,
                size=int(item["Size"]) if str(item.get("Size", "")).isdigit() else None,
                source_url=item.get("Url") if isinstance(item.get("Url"), str) else None,
                is_current=bool(item.get("IsCurrentVersion", False)),
            ))
        return tuple(sorted(versions, key=lambda version: int(version.id)))

    def get_version(self, spreadsheet: SpreadsheetInfo, version: VersionInfo) -> Path:
        self._validate_spreadsheet(spreadsheet)
        if not re.fullmatch(r"\d+", version.id):
            raise SharePointReadError("ID de versão REST inválido")
        encoded = _escape_odata_path(spreadsheet.path or "")
        url = self._endpoint(
            f"web/GetFileByServerRelativeUrl('{encoded}')/Versions({version.id})/$value"
        )
        encoded_body = self._fetch(url, binary=True)
        if not isinstance(encoded_body, str):
            raise SharePointReadError("Conteúdo histórico inválido")
        body = base64.b64decode(encoded_body, validate=True)
        safe_label = re.sub(r"[^A-Za-z0-9._-]", "_", version.number)
        path = Path(self._temporary_directory.name) / f"{spreadsheet.drive_item_id}_{version.id}_{safe_label}.xlsx"
        path.write_bytes(body)
        self._validate_xlsx(path)
        return path

    @staticmethod
    def _validate_xlsx(path: Path) -> None:
        if not path.is_file() or path.stat().st_size == 0:
            raise SharePointReadError("Versão histórica vazia ou ausente")
        try:
            with zipfile.ZipFile(path) as archive:
                if "xl/workbook.xml" not in archive.namelist():
                    raise SharePointReadError("Arquivo não contém xl/workbook.xml")
        except zipfile.BadZipFile as error:
            raise SharePointReadError("Versão histórica não é um XLSX/ZIP válido") from error

    def _validate_spreadsheet(self, spreadsheet: SpreadsheetInfo) -> None:
        if spreadsheet.site_id != self.site_url or not spreadsheet.drive_item_id or not spreadsheet.path:
            raise ValueError("A planilha não pertence ao site REST configurado")


# Nome público da fonte oficial da V1. O provider Graph permanece preservado no
# histórico Git, mas não é dependência da aquisição pelo navegador.
SharePointSource = BrowserSharePointSource
