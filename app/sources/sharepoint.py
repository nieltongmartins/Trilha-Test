"""Fonte Microsoft Graph estritamente limitada a operações HTTP GET."""

from __future__ import annotations

from collections.abc import Callable, Mapping
import json
from pathlib import Path
import tempfile
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from app.config import Settings
from app.sources.base import SpreadsheetInfo, VersionInfo


GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
JsonObject = Mapping[str, Any]
TokenProvider = Callable[[], str]
GetRequest = Callable[[str, str], tuple[bytes, Mapping[str, str]]]


class GraphReadError(RuntimeError):
    """Falha de leitura normalizada sem incluir token ou segredo."""


def _app_token_provider(tenant_id: str, client_id: str, client_secret: str) -> TokenProvider:
    def acquire() -> str:
        url = (
            f"https://login.microsoftonline.com/{quote(tenant_id, safe='')}"
            "/oauth2/v2.0/token"
        )
        body = urlencode(
            {
                "client_id": client_id,
                "client_secret": client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            }
        ).encode()
        request = Request(
            url,
            data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            with urlopen(request, timeout=60) as response:
                result = json.loads(response.read())
        except (HTTPError, URLError, json.JSONDecodeError) as error:
            raise GraphReadError("Falha na autenticação Microsoft Graph") from error
        token = result.get("access_token")
        if not token:
            code = result.get("error", "authentication_failed")
            raise GraphReadError(f"Falha na autenticação Microsoft Graph: {code}")
        return str(token)

    return acquire


def _http_get(url: str, token: str) -> tuple[bytes, Mapping[str, str]]:
    request = Request(url, headers={"Authorization": f"Bearer {token}"}, method="GET")
    try:
        with urlopen(request, timeout=60) as response:
            return response.read(), dict(response.headers.items())
    except HTTPError as error:
        raise GraphReadError(f"Microsoft Graph retornou HTTP {error.code}") from error
    except URLError as error:
        raise GraphReadError("Não foi possível conectar ao Microsoft Graph") from error


class SharePointSource:
    """Implementa o contrato de versões via endpoints de leitura do Graph."""

    def __init__(
        self,
        site_id: str,
        drive_id: str,
        token_provider: TokenProvider,
        *,
        temp_directory: str | Path = "data/temp",
        get_request: GetRequest = _http_get,
        graph_base_url: str = GRAPH_BASE_URL,
    ) -> None:
        self.site_id = site_id
        self.drive_id = drive_id
        self._token_provider = token_provider
        self._get_request = get_request
        self._graph_base_url = graph_base_url.rstrip("/")
        Path(temp_directory).mkdir(parents=True, exist_ok=True)
        self._temporary_directory = tempfile.TemporaryDirectory(
            prefix="sharepoint-", dir=temp_directory
        )

    @classmethod
    def from_settings(cls, settings: Settings) -> "SharePointSource":
        tenant, client, secret, site, drive = settings.require_sharepoint()
        return cls(
            site,
            drive,
            _app_token_provider(tenant, client, secret),
            temp_directory=settings.temp_directory,
        )

    def close(self) -> None:
        self._temporary_directory.cleanup()

    def __enter__(self) -> "SharePointSource":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def _get(self, url: str) -> tuple[bytes, Mapping[str, str]]:
        return self._get_request(url, self._token_provider())

    def _get_json(self, url: str) -> JsonObject:
        body, _ = self._get(url)
        try:
            result = json.loads(body)
        except (UnicodeDecodeError, json.JSONDecodeError) as error:
            raise GraphReadError("Microsoft Graph retornou JSON inválido") from error
        if not isinstance(result, dict):
            raise GraphReadError("Microsoft Graph retornou estrutura inesperada")
        return result

    def _paged_values(self, url: str) -> list[JsonObject]:
        values: list[JsonObject] = []
        next_url: str | None = url
        while next_url:
            payload = self._get_json(next_url)
            page = payload.get("value")
            if not isinstance(page, list):
                raise GraphReadError("Resposta paginada do Microsoft Graph é inválida")
            values.extend(item for item in page if isinstance(item, dict))
            candidate = payload.get("@odata.nextLink")
            next_url = candidate if isinstance(candidate, str) else None
        return values

    def list_spreadsheets(self) -> tuple[SpreadsheetInfo, ...]:
        fields = "id,name,parentReference,file"
        url = (
            f"{self._graph_base_url}/drives/{quote(self.drive_id, safe='')}/root/children"
            f"?$select={fields}"
        )
        spreadsheets = []
        for item in self._paged_values(url):
            name = item.get("name")
            item_id = item.get("id")
            if not isinstance(name, str) or not name.lower().endswith(".xlsx"):
                continue
            if not isinstance(item_id, str) or not isinstance(item.get("file"), dict):
                continue
            parent = item.get("parentReference")
            path = parent.get("path") if isinstance(parent, dict) else None
            spreadsheets.append(
                SpreadsheetInfo(
                    site_id=self.site_id,
                    drive_id=self.drive_id,
                    drive_item_id=item_id,
                    name=name,
                    path=path if isinstance(path, str) else None,
                )
            )
        return tuple(spreadsheets)

    def list_versions(self, spreadsheet: SpreadsheetInfo) -> tuple[VersionInfo, ...]:
        self._validate_identity(spreadsheet)
        item_id = quote(spreadsheet.drive_item_id, safe="")
        url = f"{self._graph_base_url}/drives/{quote(self.drive_id, safe='')}/items/{item_id}/versions"
        versions = []
        for item in self._paged_values(url):
            version_id = item.get("id")
            if not isinstance(version_id, str):
                raise GraphReadError("Versão retornada sem identificador")
            modified_by = item.get("lastModifiedBy")
            user = modified_by.get("user") if isinstance(modified_by, dict) else None
            author = user.get("displayName") if isinstance(user, dict) else None
            size = item.get("size")
            versions.append(
                VersionInfo(
                    id=version_id,
                    number=version_id,
                    modified_at=item.get("lastModifiedDateTime") if isinstance(item.get("lastModifiedDateTime"), str) else None,
                    author=author if isinstance(author, str) else None,
                    # DriveItemVersion não expõe comentário de versão no
                    # contrato usado aqui; não atribuímos outro metadado a ele.
                    comment=None,
                    size=size if isinstance(size, int) else None,
                )
            )
        # A validação controlada da F4 deve confirmar a ordem recebida. A API
        # normalmente entrega a coleção em ordem decrescente, enquanto o
        # AuditService exige ordem cronológica para formar N -> N+1.
        versions.reverse()
        return tuple(versions)

    def get_version(self, spreadsheet: SpreadsheetInfo, version: VersionInfo) -> Path:
        self._validate_identity(spreadsheet)
        drive = quote(self.drive_id, safe="")
        item = quote(spreadsheet.drive_item_id, safe="")
        version_id = quote(version.id, safe="")
        url = f"{self._graph_base_url}/drives/{drive}/items/{item}/versions/{version_id}/content"
        body, _ = self._get(url)
        path = Path(self._temporary_directory.name) / f"{len(list(Path(self._temporary_directory.name).iterdir())):06d}.xlsx"
        path.write_bytes(body)
        return path

    def _validate_identity(self, spreadsheet: SpreadsheetInfo) -> None:
        if spreadsheet.site_id != self.site_id or spreadsheet.drive_id != self.drive_id:
            raise ValueError("A planilha não pertence ao site e drive configurados")
