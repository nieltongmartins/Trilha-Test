"""Infraestrutura compartilhada pelos probes isolados; não é código operacional."""

from __future__ import annotations

import argparse
import ctypes
import logging
import os
from pathlib import Path
import signal
import sys
import threading
import time
from types import FrameType
from typing import Callable
from urllib.parse import urlsplit

LOGGER = logging.getLogger("sigint_probe")


def configure(test_name: str) -> Path:
    """Configure log em arquivo e stderr, sem importar módulos da aplicação."""
    log_path = Path(__file__).with_name(f"{test_name.lower()}-{os.getpid()}.log")
    formatter = logging.Formatter(
        "%(asctime)s.%(msecs)03d %(levelname)s pid=%(process)d "
        "thread=%(threadName)s %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
    )
    LOGGER.setLevel(logging.INFO)
    LOGGER.handlers.clear()
    for handler in (logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()):
        handler.setFormatter(formatter)
        LOGGER.addHandler(handler)
    LOGGER.propagate = False
    LOGGER.info(
        "início teste=%s python=%s plataforma=%s cwd=%s log=%s",
        test_name,
        sys.version.replace("\n", " "),
        sys.platform,
        os.getcwd(),
        log_path,
    )
    log_console_members()
    return log_path


def install_sigint_observer(stage: Callable[[], str]) -> Callable[[], None]:
    """Registre SIGINT e preserve exatamente a semântica do handler anterior."""
    previous = signal.getsignal(signal.SIGINT)

    def observe(signum: int, frame: FrameType | None) -> None:
        LOGGER.error(
            "SIGINT recebido sinal=%s estágio=%s frame=%s:%s handler_anterior=%r",
            signum,
            stage(),
            frame.f_code.co_filename if frame else None,
            frame.f_lineno if frame else None,
            previous,
        )
        if callable(previous):
            previous(signum, frame)
        elif previous == signal.SIG_IGN:
            return
        else:
            signal.default_int_handler(signum, frame)

    signal.signal(signal.SIGINT, observe)

    def restore() -> None:
        signal.signal(signal.SIGINT, previous)

    return restore


def log_console_members() -> None:
    """No Windows, liste PIDs anexados ao console atual (não revela o emissor)."""
    if sys.platform != "win32":
        LOGGER.info("console_pids=indisponível (não-Windows)")
        return
    pids = (ctypes.c_ulong * 64)()
    count = ctypes.windll.kernel32.GetConsoleProcessList(pids, len(pids))
    LOGGER.info("console_pids=%s", list(pids[: min(count, len(pids))]))


def add_common_edge_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--driver-path", type=Path, help="msedgedriver.exe explícito; omita para Selenium Manager"
    )
    parser.add_argument(
        "--hold-seconds", type=float, default=30.0, help="tempo de observação após abrir a página"
    )


def create_edge(driver_path: Path | None):
    """Crie sessão própria e registre o PID oficial exposto por Service.process."""
    from selenium import webdriver
    from selenium.webdriver.edge.service import Service

    service = Service(executable_path=str(driver_path)) if driver_path else Service()
    LOGGER.info("estágio=webdriver.Edge início driver_path=%s", driver_path or "Selenium Manager")
    started = time.monotonic()
    try:
        driver = webdriver.Edge(service=service)
    except KeyboardInterrupt:
        LOGGER.exception(
            "KeyboardInterrupt durante webdriver.Edge; encerrando Service próprio pid=%s",
            getattr(service.process, "pid", None),
        )
        service.stop()
        raise
    LOGGER.info(
        "estágio=webdriver.Edge concluído duração=%.3fs edgedriver_pid=%s",
        time.monotonic() - started,
        getattr(service.process, "pid", None),
    )
    log_console_members()
    return driver, service


def navigate(driver, url: str, public_label: str) -> None:
    LOGGER.info("estágio=driver.get início destino=%s", public_label)
    started = time.monotonic()
    driver.get(url)
    LOGGER.info("estágio=driver.get concluído duração=%.3fs destino=%s", time.monotonic() - started, public_label)


def safe_quit(driver) -> None:
    if driver is None:
        return
    LOGGER.info("estágio=driver.quit início (somente sessão criada pelo probe)")
    try:
        driver.quit()
    except Exception:
        LOGGER.exception("driver.quit falhou")
    else:
        LOGGER.info("estágio=driver.quit concluído")


def sanitized_url_label(url: str) -> str:
    parsed = urlsplit(url)
    return f"{parsed.scheme}://{parsed.netloc}/<caminho-omitido>"
