"""Aquisição SharePoint REST estritamente dentro da sessão autenticada do Edge."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
import logging
import re
import shutil
import time
from typing import Any, Protocol
from urllib.parse import quote, unquote, urlsplit
import zipfile

from app.sources.base import SpreadsheetInfo, VersionInfo
from app.temp_files import TemporaryWorkspace


logger = logging.getLogger("auditoria_excel.sharepoint")


class SharePointReadError(RuntimeError):
    """Falha de aquisição normalizada sem dados da sessão autenticada."""


class BrowserSession(Protocol):
    """Superfície mínima do WebDriver; deliberadamente sem cookies ou tokens."""

    def get(self, url: str) -> None: ...
    def execute_async_script(self, script: str, *args: object) -> object: ...
    def quit(self) -> None: ...
    @property
    def current_url(self) -> str: ...


_FETCH_JSON_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
fetch(url, {method: 'GET', credentials: 'same-origin', headers: {'Accept': 'application/json;odata=nometadata'}})
  .then(async response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return {json: await response.json()};
  }).then(done).catch(error => done({error: String(error)}));
"""

_DOWNLOAD_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const filename = arguments[1];
fetch(url, {method: 'GET', credentials: 'same-origin'})
  .then(async response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.blob();
  }).then(blob => {
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
    done({ok: true});
  }).catch(error => done({error: String(error)}));
"""


def _odata_value(payload: Mapping[str, Any]) -> object:
    value: object = payload.get("value")
    if value is None and isinstance(payload.get("d"), Mapping):
        value = payload["d"].get("results", payload["d"])
    return value


def _odata_results(payload: Mapping[str, Any]) -> list[Mapping[str, Any]]:
    value = _odata_value(payload)
    if not isinstance(value, list):
        raise SharePointReadError("SharePoint REST retornou coleção inválida")
    return [item for item in value if isinstance(item, Mapping)]


def _odata_object(payload: Mapping[str, Any]) -> Mapping[str, Any]:
    # With ``odata=nometadata`` SharePoint returns a single entity as the
    # top-level JSON object.  Collections still use ``value`` and older OData
    # modes can use ``d``.  Keep all three documented wire shapes separate so
    # a direct entity is not mistaken for a malformed collection response.
    if "value" in payload:
        value = payload["value"]
    elif isinstance(payload.get("d"), Mapping):
        value = payload["d"]
    else:
        value = payload
    if not isinstance(value, Mapping):
        raise SharePointReadError("SharePoint REST retornou objeto inválido")
    return value


def _escape_odata_path(path: str) -> str:
    return quote(path.replace("'", "''"), safe="/'()$=,:?&")


def _normalize_scope_path(scope_path: str, site_path: str) -> str:
    """Converte um caminho relativo ao site em server-relative URL."""
    scope = "/" + unquote(scope_path).strip("/")
    site = "/" + unquote(site_path).strip("/") if site_path.strip("/") else ""
    if site and scope.casefold() != site.casefold() and not scope.casefold().startswith(
        site.casefold() + "/"
    ):
        return f"{site}/{scope.lstrip('/')}"
    return scope


class BrowserSharePointSource:
    """Descobre, enumera e baixa XLSX por GET read-only no próprio Edge."""

    REST_CONTEXT = "sharepoint-rest"

    def __init__(
        self,
        site_url: str,
        scope_paths: Sequence[str],
        browser: BrowserSession,
        *,
        temp_directory: str | Path = "data/temp",
        owns_browser: bool = False,
        download_timeout: float = 60,
        poll_interval: float = 0.1,
        download_directory: str | Path | None = None,
    ) -> None:
        parsed = urlsplit(site_url)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("A URL do site SharePoint deve usar HTTPS")
        scopes = tuple(
            dict.fromkeys(
                _normalize_scope_path(path, parsed.path)
                for path in scope_paths
                if path.strip("/")
            )
        )
        if not scopes:
            raise ValueError("Ao menos um escopo SharePoint deve ser configurado")
        self.site_url = site_url.rstrip("/")
        self._origin = f"{parsed.scheme}://{parsed.netloc}"
        self.scope_paths = scopes
        self._browser = browser
        self._owns_browser = owns_browser
        self._download_timeout = download_timeout
        self._poll_interval = poll_interval
        self._workspace = TemporaryWorkspace(temp_directory)
        self.download_directory = Path(
            download_directory or self._workspace.path
        )
        self.download_directory.mkdir(parents=True, exist_ok=True)

    @classmethod
    def open_edge(
        cls,
        site_url: str,
        scope_paths: Sequence[str],
        *,
        temp_directory: str | Path = "data/temp",
    ) -> "BrowserSharePointSource":
        """Abre Edge visível; login/MFA continuam inteiramente sob controle do usuário."""
        from selenium import webdriver

        root = Path(temp_directory)
        root.mkdir(parents=True, exist_ok=True)
        download_workspace = TemporaryWorkspace(root)
        download_directory = download_workspace.path
        options = webdriver.EdgeOptions()
        options.add_experimental_option(
            "prefs",
            {
                "download.default_directory": str(download_directory),
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
            },
        )
        logger.info(
            "Abrindo Edge para autenticação manual SharePoint site=%s escopos=%d modo=read-only",
            site_url,
            len(scope_paths),
        )
        browser = webdriver.Edge(options=options)
        try:
            source = cls(
                site_url,
                scope_paths,
                browser,
                temp_directory=root,
                owns_browser=True,
                download_directory=download_directory,
            )
            source._download_workspace = download_workspace
        except Exception:
            browser.quit()
            download_workspace.close()
            raise
        browser.get(site_url)
        logger.info(
            "Edge aberto; aguardando autenticação manual do usuário site=%s",
            site_url,
        )
        return source

    def wait_until_authenticated(self, timeout: float = 600) -> None:
        """Espera deterministicamente a sessão alcançar e ler o site configurado."""
        from selenium.common.exceptions import TimeoutException
        from selenium.webdriver.support.ui import WebDriverWait

        logger.info("Aguardando autenticação manual SharePoint site=%s", self.site_url)

        def authenticated(_browser: BrowserSession) -> bool:
            current = urlsplit(_browser.current_url)
            site = urlsplit(self.site_url)
            if current.scheme != site.scheme or current.netloc != site.netloc:
                return False
            try:
                payload = self._json("web?$select=Id")
                entity = _odata_object(payload)
                return isinstance(entity.get("Id"), str)
            except SharePointReadError:
                return False

        try:
            WebDriverWait(self._browser, timeout, poll_frequency=0.5).until(
                authenticated
            )
        except TimeoutException as error:
            raise SharePointReadError(
                "Tempo esgotado aguardando autenticação manual no SharePoint"
            ) from error
        logger.info("Autenticação SharePoint detectada e site validado")

    def close(self) -> None:
        self._workspace.close()
        download_workspace = getattr(self, "_download_workspace", None)
        if download_workspace is not None:
            download_workspace.close()
        if self._owns_browser:
            self._browser.quit()

    def __enter__(self) -> "BrowserSharePointSource":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def _endpoint(self, relative: str) -> str:
        url = f"{self.site_url}/_api/{relative.lstrip('/')}"
        parsed = urlsplit(url)
        if (
            f"{parsed.scheme}://{parsed.netloc}" != self._origin
            or "/_api/" not in parsed.path
        ):
            raise SharePointReadError("Endpoint fora do site ou da API REST permitida")
        return url

    def _json(self, relative: str) -> Mapping[str, Any]:
        url = self._endpoint(relative)
        parsed = urlsplit(url)
        if f"{parsed.scheme}://{parsed.netloc}" != self._origin:
            raise SharePointReadError(
                "Somente endpoints GET same-origin são permitidos"
            )
        result = self._browser.execute_async_script(_FETCH_JSON_SCRIPT, url)
        if not isinstance(result, Mapping) or result.get("error"):
            detail = (
                result.get("error")
                if isinstance(result, Mapping)
                else "resposta inválida"
            )
            raise SharePointReadError(f"Falha de leitura SharePoint REST: {detail}")
        payload = result.get("json")
        if not isinstance(payload, Mapping):
            raise SharePointReadError("SharePoint REST retornou JSON inválido")
        return payload

    def list_spreadsheets(self) -> tuple[SpreadsheetInfo, ...]:
        logger.info(
            "Descoberta de planilhas iniciada site=%s escopos=%d modo=read-only",
            self.site_url,
            len(self.scope_paths),
        )
        found: dict[str, SpreadsheetInfo] = {}
        pending = list(self.scope_paths)
        visited: set[str] = set()
        while pending:
            folder = pending.pop(0)
            if folder in visited:
                continue
            visited.add(folder)
            encoded = _escape_odata_path(folder)
            files = _odata_results(
                self._json(
                    f"web/GetFolderByServerRelativeUrl('{encoded}')/Files"
                    "?$select=Name,ServerRelativeUrl,UniqueId,Length"
                )
            )
            for item in files:
                name, path, unique_id = (
                    item.get("Name"),
                    item.get("ServerRelativeUrl"),
                    item.get("UniqueId"),
                )
                if (
                    not isinstance(name, str)
                    or not isinstance(path, str)
                    or not isinstance(unique_id, str)
                ):
                    raise SharePointReadError(
                        "Arquivo retornado sem identidade REST comprovável"
                    )
                if name.lower().endswith(".xlsx"):
                    found[unique_id.lower()] = SpreadsheetInfo(
                        site_id=self.site_url,
                        drive_id=self.REST_CONTEXT,
                        drive_item_id=unique_id,
                        name=name,
                        path=path,
                        folder=path.rsplit("/", 1)[0],
                    )
            folders = _odata_results(
                self._json(
                    f"web/GetFolderByServerRelativeUrl('{encoded}')/Folders?$select=Name,ServerRelativeUrl"
                )
            )
            pending.extend(
                path
                for item in folders
                if isinstance((path := item.get("ServerRelativeUrl")), str)
                and item.get("Name") != "Forms"
            )
        result = tuple(sorted(found.values(), key=lambda item: item.path or ""))
        logger.info(
            "Descoberta de planilhas concluída site=%s pastas=%d planilhas=%d",
            self.site_url,
            len(visited),
            len(result),
        )
        return result

    def _file_metadata(self, spreadsheet: SpreadsheetInfo) -> Mapping[str, Any]:
        encoded = _escape_odata_path(spreadsheet.path or "")
        return _odata_object(
            self._json(
                f"web/GetFileByServerRelativeUrl('{encoded}')"
                "?$select=Name,ServerRelativeUrl,UniqueId,UIVersion,UIVersionLabel,"
                "TimeLastModified,Length,ModifiedBy/Title,ModifiedBy/Email,ModifiedBy/LoginName"
                "&$expand=ModifiedBy"
            )
        )

    def list_versions(self, spreadsheet: SpreadsheetInfo) -> tuple[VersionInfo, ...]:
        self._validate_spreadsheet(spreadsheet)
        encoded = _escape_odata_path(spreadsheet.path or "")
        payload = self._json(
            f"web/GetFileByServerRelativeUrl('{encoded}')/Versions"
            "?$expand=CreatedBy&$select=ID,VersionLabel,Created,CreatedBy/Title,"
            "CreatedBy/Email,CreatedBy/LoginName,CheckInComment,Size,Length,Url,IsCurrentVersion"
        )
        historical: list[VersionInfo] = []
        for item in _odata_results(payload):
            version_id, label = item.get("ID"), item.get("VersionLabel")
            if not isinstance(version_id, int) or not isinstance(label, str):
                raise SharePointReadError("Versão REST sem ID técnico ou VersionLabel")
            historical.append(self._historical_version(item, version_id, label))
        historical.sort(key=lambda version: int(version.id))

        metadata = self._file_metadata(spreadsheet)
        unique_id = metadata.get("UniqueId")
        if (
            not isinstance(unique_id, str)
            or unique_id.lower() != spreadsheet.drive_item_id.lower()
        ):
            raise SharePointReadError(
                "UniqueId atual diverge da identidade da planilha"
            )
        ui_version, current_label = (
            metadata.get("UIVersion"),
            metadata.get("UIVersionLabel"),
        )
        if not isinstance(ui_version, int) or not isinstance(current_label, str):
            raise SharePointReadError("Arquivo atual sem UIVersion/UIVersionLabel")
        modified_by = metadata.get("ModifiedBy")
        current = VersionInfo(
            id=str(ui_version),
            number=current_label,
            modified_at=metadata.get("TimeLastModified")
            if isinstance(metadata.get("TimeLastModified"), str)
            else None,
            author=modified_by.get("Title")
            if isinstance(modified_by, Mapping)
            and isinstance(modified_by.get("Title"), str)
            else None,
            author_email=modified_by.get("Email")
            if isinstance(modified_by, Mapping)
            and isinstance(modified_by.get("Email"), str)
            else None,
            author_login=modified_by.get("LoginName")
            if isinstance(modified_by, Mapping)
            and isinstance(modified_by.get("LoginName"), str)
            else None,
            size=self._size(metadata),
            source_url=spreadsheet.path,
            is_current=True,
        )
        # O endpoint histórico pode eventualmente repetir a atual. A identidade/label
        # retornados, nunca uma derivação entre ambos, controlam a deduplicação.
        combined = [
            v for v in historical if v.id != current.id and v.number != current.number
        ]
        combined.append(current)
        logger.info(
            "Versões enumeradas planilha=%s identidade=%s total=%d modo=read-only",
            spreadsheet.name,
            spreadsheet.drive_item_id,
            len(combined),
        )
        return tuple(combined)

    @staticmethod
    def _historical_version(
        item: Mapping[str, Any], version_id: int, label: str
    ) -> VersionInfo:
        author = item.get("CreatedBy")
        return VersionInfo(
            id=str(version_id),
            number=label,
            modified_at=item.get("Created")
            if isinstance(item.get("Created"), str)
            else None,
            author=author.get("Title")
            if isinstance(author, Mapping) and isinstance(author.get("Title"), str)
            else None,
            author_email=author.get("Email")
            if isinstance(author, Mapping) and isinstance(author.get("Email"), str)
            else None,
            author_login=author.get("LoginName")
            if isinstance(author, Mapping) and isinstance(author.get("LoginName"), str)
            else None,
            comment=item.get("CheckInComment")
            if isinstance(item.get("CheckInComment"), str)
            else None,
            size=BrowserSharePointSource._size(item),
            source_url=item.get("Url") if isinstance(item.get("Url"), str) else None,
            is_current=bool(item.get("IsCurrentVersion", False)),
        )

    @staticmethod
    def _size(item: Mapping[str, Any]) -> int | None:
        value = item.get("Size", item.get("Length"))
        return (
            int(value)
            if isinstance(value, (int, str)) and str(value).isdigit()
            else None
        )

    def get_version(self, spreadsheet: SpreadsheetInfo, version: VersionInfo) -> Path:
        self._validate_spreadsheet(spreadsheet)
        if not re.fullmatch(r"\d+", version.id):
            raise SharePointReadError("ID de versão REST inválido")
        encoded = _escape_odata_path(spreadsheet.path or "")
        relative = f"web/GetFileByServerRelativeUrl('{encoded}')"
        if not version.is_current:
            relative += f"/Versions({version.id})"
        url = self._endpoint(relative + "/$value")
        destination = self._workspace.filename(
            spreadsheet.site_id,
            spreadsheet.drive_id,
            spreadsheet.drive_item_id,
            version.id,
        )
        before = set(self.download_directory.iterdir())
        result = self._browser.execute_async_script(
            _DOWNLOAD_SCRIPT, url, destination.name
        )
        if not isinstance(result, Mapping) or result.get("error") or not result.get("ok"):
            detail = (
                result.get("error")
                if isinstance(result, Mapping)
                else "resposta inválida"
            )
            raise SharePointReadError(f"Falha no download SharePoint REST: {detail}")
        downloaded = self._wait_for_download(before, destination.name)
        if downloaded.resolve() != destination.resolve():
            shutil.move(str(downloaded), destination)
        try:
            self._validate_xlsx(destination)
        except Exception:
            destination.unlink(missing_ok=True)
            raise
        logger.debug(
            "Versão adquirida planilha=%s identidade=%s versao=%s atual=%s",
            spreadsheet.name,
            spreadsheet.drive_item_id,
            version.number,
            version.is_current,
        )
        return destination

    def release_version(self, path: Path) -> None:
        self._workspace.release(path)

    def _wait_for_download(self, before: set[Path], expected_name: str) -> Path:
        deadline = time.monotonic() + self._download_timeout
        while time.monotonic() < deadline:
            files = {
                path for path in self.download_directory.iterdir() if path.is_file()
            }
            partial = [
                path
                for path in files
                if path.name == f"{expected_name}.crdownload"
            ]
            completed = [path for path in files - before if path.name == expected_name]
            if completed and not partial:
                return completed[0]
            time.sleep(self._poll_interval)
        raise SharePointReadError(
            "Download não foi concluído; arquivo .crdownload presente ou ausente"
        )

    @staticmethod
    def _validate_xlsx(path: Path) -> None:
        if not path.is_file() or path.stat().st_size == 0:
            raise SharePointReadError("Versão vazia ou ausente")
        try:
            with zipfile.ZipFile(path) as archive:
                if (
                    "[Content_Types].xml" not in archive.namelist()
                    or "xl/workbook.xml" not in archive.namelist()
                ):
                    raise SharePointReadError(
                        "Arquivo não é Open XML com xl/workbook.xml"
                    )
        except zipfile.BadZipFile as error:
            raise SharePointReadError("Versão não é um XLSX/ZIP válido") from error

    def _validate_spreadsheet(self, spreadsheet: SpreadsheetInfo) -> None:
        if (
            spreadsheet.site_id != self.site_url
            or spreadsheet.drive_id != self.REST_CONTEXT
            or not spreadsheet.drive_item_id
            or not spreadsheet.path
        ):
            raise ValueError("A planilha não pertence ao site REST configurado")


SharePointSource = BrowserSharePointSource
