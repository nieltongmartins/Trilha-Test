"""Ciclo de vida seguro dos XLSX temporários adquiridos pela aplicação."""

from __future__ import annotations

import hashlib
import json
import logging
import os
from pathlib import Path
import shutil
import tempfile
from uuid import uuid4


logger = logging.getLogger("auditoria_excel.temp")
_PREFIX = "auditoria-xlsx-"
_MARKER = ".auditoria-temp.json"


def _windows_process_is_running(pid: int) -> bool:
    """Consulta um PID no Windows sem enviar sinais ao processo."""
    import ctypes
    from ctypes import wintypes

    process_query_limited_information = 0x1000
    still_active = 259
    error_invalid_parameter = 87
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.GetExitCodeProcess.argtypes = (wintypes.HANDLE, wintypes.LPDWORD)
    kernel32.GetExitCodeProcess.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL

    handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
    if not handle:
        # PID inexistente gera ERROR_INVALID_PARAMETER. Falta de permissão não
        # demonstra que o processo terminou, então é tratada como ativo.
        return ctypes.get_last_error() != error_invalid_parameter
    try:
        exit_code = wintypes.DWORD()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
            return True
        return exit_code.value == still_active
    finally:
        kernel32.CloseHandle(handle)


def _process_is_running(pid: int) -> bool:
    if pid <= 0:
        return False
    if pid == os.getpid():
        return True
    if os.name == "nt":
        return _windows_process_is_running(pid)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except (PermissionError, OSError):
        return True
    return True


class TemporaryWorkspace:
    """Diretório isolado, identificável e removível sem limpeza ampla."""

    def __init__(self, root: str | Path) -> None:
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)
        self.remove_orphans(self.root)
        self.path = Path(tempfile.mkdtemp(prefix=_PREFIX, dir=self.root)).resolve()
        marker = {"owner": "auditoria_excel", "pid": os.getpid(), "run": uuid4().hex}
        (self.path / _MARKER).write_text(json.dumps(marker), encoding="utf-8")
        self._closed = False
        logger.debug("Diretório temporário criado diretorio=%s", self.path)

    @classmethod
    def remove_orphans(cls, root: str | Path) -> int:
        """Remove somente workspaces reconhecidos cujo processo não existe mais."""
        base = Path(root)
        if not base.is_dir():
            return 0
        removed = 0
        for candidate in base.iterdir():
            if not candidate.is_dir() or not candidate.name.startswith(_PREFIX):
                continue
            marker_path = candidate / _MARKER
            try:
                marker = json.loads(marker_path.read_text(encoding="utf-8"))
                owned = marker.get("owner") == "auditoria_excel"
                pid = int(marker["pid"])
            except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
                continue
            if owned and not _process_is_running(pid):
                shutil.rmtree(candidate)
                removed += 1
        if removed:
            logger.info("Arquivos temporários órfãos removidos diretorios=%d", removed)
        return removed

    def filename(self, *identity: str) -> Path:
        """Produz nome determinístico e seguro, sem incorporar dados em caminhos."""
        digest = hashlib.sha256("\0".join(identity).encode("utf-8")).hexdigest()
        return self.path / f"{digest}.xlsx"

    def owns(self, path: str | Path) -> bool:
        try:
            return Path(path).resolve().parent == self.path
        except OSError:
            return False

    def release(self, path: str | Path) -> None:
        if not self.owns(path):
            raise ValueError("Arquivo temporário não pertence a esta execução")
        Path(path).unlink(missing_ok=True)

    def close(self) -> None:
        if not self._closed:
            shutil.rmtree(self.path, ignore_errors=False)
            self._closed = True
            logger.debug("Diretório temporário removido diretorio=%s", self.path)
