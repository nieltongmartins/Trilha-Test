"""Diagnóstico manual isolado de Tk + thread + Edge no Windows.

Este script não importa nem configura qualquer módulo da aplicação.
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
import signal
import threading
import time
import tkinter as tk


LOGGER = logging.getLogger("tk_edge_probe")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--mode",
        choices=("control", "edge", "edge-get"),
        default="edge",
        help="controle sem Selenium, só cria Edge, ou cria Edge e abre about:blank",
    )
    parser.add_argument(
        "--non-daemon",
        action="store_true",
        help="executa a worker como não-daemon para comparação controlada",
    )
    parser.add_argument(
        "--driver-path",
        type=Path,
        help="usa msedgedriver.exe explícito e não executa o Selenium Manager",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(threadName)s %(message)s",
    )
    stop = threading.Event()
    driver_holder: list[object] = []
    close_requested = False

    previous_sigint_handler = signal.getsignal(signal.SIGINT)

    def sigint_handler(signum, frame) -> None:
        LOGGER.error(
            "SIGINT recebido sinal=%s thread=%s frame=%s:%s handler_anterior=%r",
            signum,
            threading.current_thread().name,
            frame.f_code.co_filename if frame is not None else None,
            frame.f_lineno if frame is not None else None,
            previous_sigint_handler,
        )
        if callable(previous_sigint_handler):
            previous_sigint_handler(signum, frame)
        elif previous_sigint_handler == signal.SIG_IGN:
            return
        else:
            signal.default_int_handler(signum, frame)

    def thread_exception(hook_args: threading.ExceptHookArgs) -> None:
        LOGGER.error(
            "worker exceção tipo=%s erro=%r ident=%s daemon=%s",
            hook_args.exc_type.__name__,
            hook_args.exc_value,
            hook_args.thread.ident if hook_args.thread else None,
            hook_args.thread.daemon if hook_args.thread else None,
            exc_info=(
                hook_args.exc_type,
                hook_args.exc_value,
                hook_args.exc_traceback,
            ),
        )

    threading.excepthook = thread_exception
    signal.signal(signal.SIGINT, sigint_handler)
    root = tk.Tk()
    root.title(f"Tk Edge probe: {args.mode}")
    tk.Label(root, text="Feche esta janela para encerrar o diagnóstico.").pack(
        padx=24, pady=24
    )

    def worker() -> None:
        current = threading.current_thread()
        LOGGER.info(
            "worker iniciou mode=%s ident=%s daemon=%s",
            args.mode,
            current.ident,
            current.daemon,
        )
        try:
            if args.mode == "control":
                LOGGER.info("controle aguardando Event (sem Selenium)")
                stop.wait()
                return

            from selenium import webdriver
            from selenium.webdriver.edge.service import Service

            started = time.monotonic()
            LOGGER.info("webdriver.Edge iniciado")
            service = (
                Service(executable_path=str(args.driver_path))
                if args.driver_path is not None
                else None
            )
            LOGGER.info(
                "resolução do driver=%s",
                "caminho explícito (sem Selenium Manager)"
                if service is not None
                else "Selenium Manager",
            )
            driver = webdriver.Edge(service=service) if service else webdriver.Edge()
            driver_holder.append(driver)
            LOGGER.info(
                "webdriver.Edge concluído duração=%.3fs", time.monotonic() - started
            )
            if args.mode == "edge-get":
                started = time.monotonic()
                LOGGER.info("browser.get iniciado url=about:blank")
                driver.get("about:blank")
                LOGGER.info(
                    "browser.get concluído duração=%.3fs", time.monotonic() - started
                )
            stop.wait()
        finally:
            if driver_holder:
                LOGGER.info("driver.quit iniciado")
                driver_holder.pop().quit()  # type: ignore[attr-defined]
                LOGGER.info("driver.quit concluído")
            LOGGER.info("worker terminou")

    worker_thread = threading.Thread(
        target=worker,
        name="tk-edge-probe-worker",
        daemon=not args.non_daemon,
    )
    LOGGER.info(
        "worker criada mode=%s ident=%s daemon=%s",
        args.mode,
        worker_thread.ident,
        worker_thread.daemon,
    )

    def close() -> None:
        nonlocal close_requested
        close_requested = True
        LOGGER.info("WM_DELETE_WINDOW recebido")
        stop.set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", close)
    worker_thread.start()
    LOGGER.info("Entrando em root.mainloop")
    try:
        root.mainloop()
        if close_requested:
            LOGGER.info("root.mainloop retornou após WM_DELETE_WINDOW")
        else:
            LOGGER.warning(
                "root.mainloop RETORNOU NORMALMENTE sem solicitação de fechamento"
            )
    except BaseException as error:
        LOGGER.exception(
            "root.mainloop terminou por BaseException tipo=%s erro=%r",
            type(error).__name__,
            error,
        )
        raise
    finally:
        stop.set()
        signal.signal(signal.SIGINT, previous_sigint_handler)
        LOGGER.info(
            "finally mainloop worker_alive=%s root_exists=%s",
            worker_thread.is_alive(),
            _root_exists(root),
        )
    if not worker_thread.daemon:
        worker_thread.join(timeout=30)
        if worker_thread.is_alive():
            LOGGER.error("worker não terminou em 30s")
            return 2
    return 0


def _root_exists(root: tk.Tk) -> object:
    try:
        return root.winfo_exists()
    except tk.TclError as error:
        return f"indisponível:{type(error).__name__}"


if __name__ == "__main__":
    raise SystemExit(main())
