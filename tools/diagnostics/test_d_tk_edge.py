"""TESTE D: Tkinter + Edge; separa criação do driver e about:blank."""
import argparse
import threading
import tkinter as tk
from probe_common import (LOGGER, add_common_edge_arguments, configure, create_edge,
                          install_sigint_observer, navigate, safe_quit)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    add_common_edge_arguments(parser)
    args = parser.parse_args()
    configure("teste-d")
    stage = "tk-create"
    restore = install_sigint_observer(lambda: stage)
    stop = threading.Event()
    holder = {"driver": None}
    root = tk.Tk()
    root.title("Teste D - Tk + Edge")

    def worker() -> None:
        nonlocal stage
        try:
            stage = "webdriver.Edge"
            holder["driver"], _service = create_edge(args.driver_path)
            stage = "webdriver-created"
            if stop.wait(3):
                return
            stage = "driver.get-about-blank"
            navigate(holder["driver"], "about:blank", "about:blank")
            stage = "hold"
            stop.wait(args.hold_seconds)
        finally:
            stage = "driver.quit"
            safe_quit(holder["driver"])
            holder["driver"] = None
            stop.set()
            root.after(0, root.destroy)

    thread = threading.Thread(target=worker, name="selenium-worker", daemon=False)
    root.protocol("WM_DELETE_WINDOW", lambda: (stop.set(), root.destroy()))
    thread.start()
    try:
        LOGGER.info("estágio=tk-mainloop selenium_worker=%s", thread.ident)
        root.mainloop()
        return 0
    except KeyboardInterrupt:
        LOGGER.exception("KeyboardInterrupt preservado estágio=%s", stage)
        raise
    finally:
        stop.set()
        thread.join(timeout=60)
        LOGGER.info("selenium_worker_alive=%s", thread.is_alive())
        restore()


if __name__ == "__main__":
    raise SystemExit(main())
