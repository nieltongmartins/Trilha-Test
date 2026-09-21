"""Aquisição SharePoint REST estritamente dentro da sessão autenticada do Edge."""

from __future__ import annotations

import base64
from collections.abc import Callable, Mapping, Sequence
from pathlib import Path
import logging
import re
import time
from typing import Any, Protocol
from urllib.parse import quote, unquote, urljoin, urlsplit
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
    def set_script_timeout(self, time_to_wait: float) -> None: ...
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

_BEGIN_DOWNLOAD_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
fetch(url, {method: 'GET', credentials: 'same-origin'})
  .then(async response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.blob();
  }).then(blob => {
    window.__auditDownloadBlob = blob;
    done({ok: true, size: blob.size});
  }).catch(error => done({error: String(error)}));
"""

_READ_DOWNLOAD_CHUNK_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const offset = arguments[0];
const length = arguments[1];
const blob = window.__auditDownloadBlob;
if (!(blob instanceof Blob)) {
  done({error: 'download não inicializado'});
} else {
  blob.slice(offset, offset + length).arrayBuffer()
    .then(buffer => {
      const bytes = new Uint8Array(buffer);
      let binary = '';
      for (let index = 0; index < bytes.length; index += 1) {
        binary += String.fromCharCode(bytes[index]);
      }
      done({ok: true, data: btoa(binary), length: bytes.length});
    }).catch(error => done({error: String(error)}));
}
"""

_CLEAR_DOWNLOAD_SCRIPT = r"""
const done = arguments[arguments.length - 1];
window.__auditDownloadBlob = null;
done({ok: true});
"""

_START_PREFETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const current = window.__auditPrefetch;
if (current && current.url === url && (current.state === 'loading' || current.state === 'ready')) {
  done({ok: true, state: current.state});
} else {
  const slot = {url, state: 'loading', blob: null, size: 0, error: null};
  window.__auditPrefetch = slot;
  fetch(url, {method: 'GET', credentials: 'same-origin'})
    .then(async response => {
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return response.blob();
    })
    .then(blob => {
      if (window.__auditPrefetch === slot) {
        slot.blob = blob;
        slot.size = blob.size;
        slot.state = 'ready';
      }
    })
    .catch(error => {
      if (window.__auditPrefetch === slot) {
        slot.error = String(error);
        slot.state = 'error';
      }
    });
  done({ok: true, state: 'started'});
}
"""

_WAIT_PREFETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
function poll() {
  const slot = window.__auditPrefetch;
  if (!slot || slot.url !== url) {
    done({error: 'prefetch indisponível'});
    return;
  }
  if (slot.state === 'ready' && slot.blob instanceof Blob) {
    done({ok: true, size: slot.size});
    return;
  }
  if (slot.state === 'error') {
    done({error: slot.error || 'falha no prefetch'});
    return;
  }
  setTimeout(poll, 50);
}
poll();
"""

_READ_PREFETCH_CHUNK_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const offset = arguments[1];
const length = arguments[2];
const slot = window.__auditPrefetch;
if (!slot || slot.url !== url || !(slot.blob instanceof Blob)) {
  done({error: 'prefetch não inicializado'});
} else {
  slot.blob.slice(offset, offset + length).arrayBuffer()
    .then(buffer => {
      const bytes = new Uint8Array(buffer);
      let binary = '';
      for (let index = 0; index < bytes.length; index += 1) {
        binary += String.fromCharCode(bytes[index]);
      }
      done({ok: true, data: btoa(binary), length: bytes.length});
    }).catch(error => done({error: String(error)}));
}
"""

_CLEAR_PREFETCH_SCRIPT = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];
const slot = window.__auditPrefetch;
if (!slot || !url || slot.url === url) {
  window.__auditPrefetch = null;
}
done({ok: true});
"""
_DOWNLOAD_CHUNK_SIZE = 512 * 1024
_VERSION_PAGE_SIZE = 1000
_VERSION_PAGE_RETRIES = 3


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
    ) -> None:
        parsed = urlsplit(site_url)
        if parsed.scheme != "https" or not parsed.netloc:
            raise ValueError("A URL do site SharePoint deve usar HTTPS")
        self.site_url = site_url.rstrip("/")
        self._origin = f"{parsed.scheme}://{parsed.netloc}"
        self._site_path = parsed.path
        scopes = tuple(
            dict.fromkeys(
                _normalize_scope_path(path, self._site_path)
                for path in scope_paths
                if path.strip("/")
            )
        )
        if not scopes:
            raise ValueError("Ao menos um escopo SharePoint deve ser configurado")
        self.scope_paths = scopes
        self._browser = browser
        self._owns_browser = owns_browser
        self._workspace = TemporaryWorkspace(temp_directory)
        self._prefetch_url: str | None = None
        self._prefetch_started_at: float | None = None

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
        options = webdriver.EdgeOptions()
        logger.info(
            "Abrindo Edge para autenticação manual SharePoint site=%s escopos=%d modo=read-only",
            site_url,
            len(scope_paths),
        )
        browser = webdriver.Edge(options=options)
        # Há dois timeouts distintos no Selenium:
        # 1) o timeout do JavaScript assíncrono executado no Edge;
        # 2) o timeout HTTP usado pelo Python para aguardar o WebDriver local.
        # A enumeração de dezenas de milhares de versões pode ultrapassar os
        # 120 s padrão do segundo limite, mesmo com script_timeout=600.
        browser.set_script_timeout(600)
        command_executor = getattr(browser, "command_executor", None)
        if command_executor is not None:
            set_timeout = getattr(command_executor, "set_timeout", None)
            if callable(set_timeout):
                set_timeout(600)
            else:
                client_config = getattr(command_executor, "_client_config", None)
                if client_config is not None:
                    client_config.timeout = 600
        logger.info("Timeouts Selenium configurados script=600s webdriver_http=600s")
        try:
            source = cls(
                site_url,
                scope_paths,
                browser,
                temp_directory=root,
                owns_browser=True,
            )
        except Exception:
            browser.quit()
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
        """Fecha a sessão sem transformar arquivo temporário bloqueado em falha fatal."""
        try:
            self.cancel_prefetch()
        except Exception:
            logger.debug("Falha ao cancelar prefetch durante encerramento", exc_info=True)

        if self._owns_browser:
            try:
                self._browser.quit()
            except Exception:
                logger.debug("Falha ao encerrar Edge durante encerramento", exc_info=True)

        # No Windows, antivírus/indexador ou uma thread que ainda esteja terminando
        # pode manter o XLSX aberto por alguns instantes. Fazemos tentativas curtas
        # e, se continuar bloqueado, deixamos a pasta temporária para a próxima
        # limpeza em vez de derrubar a aplicação com WinError 32.
        last_error: Exception | None = None
        for attempt in range(5):
            try:
                self._workspace.close()
                return
            except PermissionError as error:
                last_error = error
                time.sleep(0.2 * (attempt + 1))
            except FileNotFoundError:
                return
        if last_error is not None:
            logger.warning(
                "Não foi possível remover imediatamente o diretório temporário; "
                "ele permanecerá para limpeza posterior erro=%s",
                last_error,
            )

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
        return self._json_url(self._endpoint(relative))

    def _json_url(self, url: str) -> Mapping[str, Any]:
        """Lê um endpoint REST absoluto, sempre limitado ao mesmo site SharePoint."""
        parsed = urlsplit(url)
        if (
            f"{parsed.scheme}://{parsed.netloc}" != self._origin
            or "/_api/" not in parsed.path
        ):
            raise SharePointReadError(
                "Somente endpoints GET same-origin da API REST são permitidos"
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

    def _json_url_with_retry(self, url: str, *, page: int) -> Mapping[str, Any]:
        """Repete somente a página que falhou, sem reiniciar toda a enumeração."""
        last_error: Exception | None = None
        for attempt in range(1, _VERSION_PAGE_RETRIES + 1):
            try:
                return self._json_url(url)
            except Exception as error:
                last_error = error
                if attempt >= _VERSION_PAGE_RETRIES:
                    break
                delay = float(2 ** (attempt - 1))
                logger.warning(
                    "Falha temporária ao enumerar versões pagina=%d tentativa=%d/%d "
                    "aguardando=%.0fs erro=%s",
                    page,
                    attempt,
                    _VERSION_PAGE_RETRIES,
                    delay,
                    error,
                )
                time.sleep(delay)
        assert last_error is not None
        raise last_error

    @staticmethod
    def _next_page_url(payload: Mapping[str, Any]) -> str | None:
        next_link = (
            payload.get("@odata.nextLink")
            or payload.get("odata.nextLink")
            or payload.get("__next")
        )
        if next_link is None and isinstance(payload.get("d"), Mapping):
            next_link = payload["d"].get("__next")
        return next_link if isinstance(next_link, str) and next_link else None

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

    def list_folders(self) -> tuple[tuple[str, str], ...]:
        """Lista recursivamente as subpastas dos escopos para seleção na interface."""
        found: dict[str, str] = {}
        pending = list(self.scope_paths)
        visited: set[str] = set()
        while pending:
            folder = pending.pop(0)
            if folder in visited:
                continue
            visited.add(folder)
            encoded = _escape_odata_path(folder)
            payload = self._json(
                f"web/GetFolderByServerRelativeUrl('{encoded}')/Folders"
                "?$select=Name,ServerRelativeUrl"
            )
            for item in _odata_results(payload):
                name, path = item.get("Name"), item.get("ServerRelativeUrl")
                if (
                    isinstance(name, str)
                    and isinstance(path, str)
                    and name != "Forms"
                ):
                    found[path] = name
                    pending.append(path)
        return tuple(
            (name, path)
            for path, name in sorted(found.items(), key=lambda item: item[1].casefold())
        )

    def set_scope_paths(self, scope_paths: Sequence[str]) -> None:
        """Atualiza os escopos de leitura sem recriar a sessão autenticada."""
        scopes = tuple(
            dict.fromkeys(
                _normalize_scope_path(path, self._site_path)
                for path in scope_paths
                if path.strip("/")
            )
        )
        if not scopes:
            raise ValueError("Ao menos um escopo SharePoint deve ser configurado")
        self.scope_paths = scopes

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

    def list_versions(
        self,
        spreadsheet: SpreadsheetInfo,
        progress_callback: Callable[[int], None] | None = None,
    ) -> tuple[VersionInfo, ...]:
        """Enumera todo o histórico em páginas, sem aceitar truncamento silencioso.

        O endpoint ``File/Versions`` do SharePoint Online nem sempre devolve um
        ``@odata.nextLink`` quando ``$top`` é usado. Por isso, seguimos o
        ``nextLink`` quando ele existir e, se uma página vier cheia sem link de
        continuação, avançamos ordinalmente com ``$skip``. Se o servidor ignorar
        o ``$skip`` e repetir a mesma página, a operação é abortada em vez de
        entregar uma lista incompleta à auditoria.
        """
        self._validate_spreadsheet(spreadsheet)
        encoded = _escape_odata_path(spreadsheet.path or "")
        base_relative = (
            f"web/GetFileByServerRelativeUrl('{encoded}')/Versions"
            "?$expand=CreatedBy&$select=ID,VersionLabel,Created,CreatedBy/Title,"
            "CreatedBy/Email,CreatedBy/LoginName,CheckInComment,Size,Length,Url,IsCurrentVersion"
        )

        next_url: str | None = self._endpoint(
            base_relative + f"&$top={_VERSION_PAGE_SIZE}&$skip=0"
        )
        historical: list[VersionInfo] = []
        seen_ids: set[str] = set()
        page = 0
        skip = 0
        started = time.perf_counter()

        while next_url:
            page += 1
            payload = self._json_url_with_retry(next_url, page=page)
            items = _odata_results(payload)
            new_in_page = 0

            for item in items:
                version_id, label = item.get("ID"), item.get("VersionLabel")
                if not isinstance(version_id, int) or not isinstance(label, str):
                    raise SharePointReadError(
                        "Versão REST sem ID técnico ou VersionLabel"
                    )
                version_key = str(version_id)
                if version_key in seen_ids:
                    continue
                seen_ids.add(version_key)
                historical.append(self._historical_version(item, version_id, label))
                new_in_page += 1

            if items and new_in_page == 0:
                raise SharePointReadError(
                    "SharePoint repetiu a mesma página de versões; "
                    "a enumeração foi interrompida para evitar histórico incompleto."
                )

            if progress_callback is not None:
                try:
                    progress_callback(len(historical))
                except Exception:
                    logger.warning(
                        "Falha ao publicar progresso da enumeração de versões",
                        exc_info=True,
                    )

            logger.info(
                "Enumeração de versões pagina=%d lote=%d novas=%d acumulado=%d skip=%d planilha=%s",
                page,
                len(items),
                new_in_page,
                len(historical),
                skip,
                spreadsheet.name,
            )

            next_link = self._next_page_url(payload)
            if next_link is not None:
                next_url = urljoin(self.site_url + "/", next_link)
                # Mantemos o skip apenas para log/fallback caso o próximo payload
                # deixe de fornecer continuação.
                skip += len(items)
            elif len(items) >= _VERSION_PAGE_SIZE:
                # Alguns tenants/coleções de versões aplicam $top, mas não expõem
                # nextLink. Nesse caso, avançamos com $skip.
                skip += len(items)
                next_url = self._endpoint(
                    base_relative
                    + f"&$top={_VERSION_PAGE_SIZE}&$skip={skip}"
                )
            else:
                next_url = None

        historical.sort(key=lambda version: int(version.id))
        logger.info(
            "PERF enumeracao_versoes planilha=%s paginas=%d historicas=%d total=%.3fs",
            spreadsheet.name,
            page,
            len(historical),
            time.perf_counter() - started,
        )

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

    def _version_download_url(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> str:
        self._validate_spreadsheet(spreadsheet)
        if not re.fullmatch(r"\d+", version.id):
            raise SharePointReadError("ID de versão REST inválido")
        encoded = _escape_odata_path(spreadsheet.path or "")
        relative = f"web/GetFileByServerRelativeUrl('{encoded}')"
        if not version.is_current:
            relative += f"/Versions({version.id})"
        return self._endpoint(relative + "/$value")

    def prefetch_version(
        self, spreadsheet: SpreadsheetInfo, version: VersionInfo
    ) -> bool:
        """Inicia o download da próxima versão no Edge sem bloquear o Python.

        Apenas um prefetch fica ativo. O ``fetch`` continua no próprio Edge
        enquanto o Python calcula hash, lê o XLSX e compara a versão atual.
        """
        url = self._version_download_url(spreadsheet, version)
        started = time.perf_counter()
        result = self._browser.execute_async_script(_START_PREFETCH_SCRIPT, url)
        if not isinstance(result, Mapping) or result.get("error") or not result.get("ok"):
            detail = result.get("error") if isinstance(result, Mapping) else "resposta inválida"
            logger.warning(
                "Prefetch não iniciado planilha=%s versao=%s erro=%s",
                spreadsheet.name,
                version.number,
                detail,
            )
            return False
        self._prefetch_url = url
        self._prefetch_started_at = started
        logger.info(
            "PERF prefetch_iniciado planilha=%s versao=%s estado=%s",
            spreadsheet.name,
            version.number,
            result.get("state", "started"),
        )
        return True

    def cancel_prefetch(self) -> None:
        """Descarta o blob pré-baixado, sem alterar qualquer arquivo no SharePoint."""
        url = self._prefetch_url
        try:
            self._browser.execute_async_script(_CLEAR_PREFETCH_SCRIPT, url or "")
        except Exception:
            logger.debug("Falha ao limpar prefetch", exc_info=True)
        finally:
            self._prefetch_url = None
            self._prefetch_started_at = None

    def _download_to_destination(
        self,
        spreadsheet: SpreadsheetInfo,
        version: VersionInfo,
        url: str,
        destination: Path,
        *,
        use_prefetch: bool,
    ) -> Path:
        total_started = time.perf_counter()
        begin_started = time.perf_counter()
        if use_prefetch:
            result = self._browser.execute_async_script(_WAIT_PREFETCH_SCRIPT, url)
        else:
            result = self._browser.execute_async_script(_BEGIN_DOWNLOAD_SCRIPT, url)
        begin_seconds = time.perf_counter() - begin_started

        if (
            not isinstance(result, Mapping)
            or result.get("error")
            or not result.get("ok")
        ):
            detail = result.get("error") if isinstance(result, Mapping) else "resposta inválida"
            raise SharePointReadError(f"Falha no download SharePoint REST: {detail}")

        temporary = destination.with_suffix(destination.suffix + ".part")
        clear_seconds = 0.0
        size = result.get("size")
        if not isinstance(size, int) or size <= 0:
            raise SharePointReadError("Download SharePoint vazio ou com tamanho inválido")

        try:
            chunk_calls = 0
            chunk_transfer_seconds = 0.0
            decode_write_seconds = 0.0
            with temporary.open("wb") as output:
                offset = 0
                while offset < size:
                    chunk_started = time.perf_counter()
                    if use_prefetch:
                        chunk = self._browser.execute_async_script(
                            _READ_PREFETCH_CHUNK_SCRIPT,
                            url,
                            offset,
                            min(_DOWNLOAD_CHUNK_SIZE, size - offset),
                        )
                    else:
                        chunk = self._browser.execute_async_script(
                            _READ_DOWNLOAD_CHUNK_SCRIPT,
                            offset,
                            min(_DOWNLOAD_CHUNK_SIZE, size - offset),
                        )
                    chunk_transfer_seconds += time.perf_counter() - chunk_started
                    chunk_calls += 1
                    if (
                        not isinstance(chunk, Mapping)
                        or chunk.get("error")
                        or not isinstance(chunk.get("data"), str)
                        or not isinstance(chunk.get("length"), int)
                        or chunk["length"] <= 0
                    ):
                        detail = chunk.get("error") if isinstance(chunk, Mapping) else "resposta inválida"
                        raise SharePointReadError(
                            f"Falha ao ler download SharePoint: {detail}"
                        )
                    decode_started = time.perf_counter()
                    decoded = base64.b64decode(chunk["data"], validate=True)
                    if len(decoded) != chunk["length"]:
                        raise SharePointReadError(
                            "Tamanho do bloco baixado diverge do informado"
                        )
                    output.write(decoded)
                    offset += chunk["length"]
                    decode_write_seconds += time.perf_counter() - decode_started
            temporary.replace(destination)
            validate_started = time.perf_counter()
            self._validate_xlsx(destination)
            validate_seconds = time.perf_counter() - validate_started
        except Exception:
            temporary.unlink(missing_ok=True)
            destination.unlink(missing_ok=True)
            raise
        finally:
            clear_started = time.perf_counter()
            if use_prefetch:
                self._browser.execute_async_script(_CLEAR_PREFETCH_SCRIPT, url)
                self._prefetch_url = None
                self._prefetch_started_at = None
            else:
                self._browser.execute_async_script(_CLEAR_DOWNLOAD_SCRIPT)
            clear_seconds = time.perf_counter() - clear_started

        total_seconds = time.perf_counter() - total_started
        prefetch_age = 0.0
        if use_prefetch and self._prefetch_started_at is not None:
            prefetch_age = time.perf_counter() - self._prefetch_started_at
        logger.info(
            "PERF download_sharepoint planilha=%s versao=%s modo=%s bytes=%d blocos=%d "
            "bloco_bytes=%d espera_fetch=%.3fs transferencia_blocos=%.3fs "
            "decode_gravacao=%.3fs validacao_xlsx=%.3fs limpeza=%.3fs total=%.3fs",
            spreadsheet.name,
            version.number,
            "prefetch" if use_prefetch else "normal",
            size,
            chunk_calls,
            _DOWNLOAD_CHUNK_SIZE,
            begin_seconds,
            chunk_transfer_seconds,
            decode_write_seconds,
            validate_seconds,
            clear_seconds,
            total_seconds,
        )
        logger.debug(
            "Versão adquirida planilha=%s identidade=%s versao=%s atual=%s modo=%s",
            spreadsheet.name,
            spreadsheet.drive_item_id,
            version.number,
            version.is_current,
            "prefetch" if use_prefetch else "normal",
        )
        return destination

    def get_version(self, spreadsheet: SpreadsheetInfo, version: VersionInfo) -> Path:
        url = self._version_download_url(spreadsheet, version)
        destination = self._workspace.filename(
            spreadsheet.site_id,
            spreadsheet.drive_id,
            spreadsheet.drive_item_id,
            version.id,
        )
        use_prefetch = self._prefetch_url == url
        if use_prefetch:
            try:
                return self._download_to_destination(
                    spreadsheet, version, url, destination, use_prefetch=True
                )
            except SharePointReadError as error:
                logger.warning(
                    "Prefetch falhou; repetindo download normal planilha=%s versao=%s erro=%s",
                    spreadsheet.name,
                    version.number,
                    error,
                )
                self.cancel_prefetch()
        return self._download_to_destination(
            spreadsheet, version, url, destination, use_prefetch=False
        )

    def release_version(self, path: Path) -> None:
        self._workspace.release(path)

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
