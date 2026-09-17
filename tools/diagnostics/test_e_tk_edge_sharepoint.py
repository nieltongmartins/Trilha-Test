"""TESTE E: Tkinter + Selenium + URL real, sem importar a aplicação."""
import argparse
import threading
import tkinter as tk
from probe_common import (LOGGER, add_common_edge_arguments, configure, create_edge,
                          install_sigint_observer, navigate, safe_quit, sanitized_url_label)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("url", help="URL HTTPS real do SharePoint")
    add_common_edge_arguments(parser)
    args = parser.parse_args()
    if not args.url.lower().startswith("https://"):
        parser.error("a URL deve usar HTTPS")
    configure("teste-e")
    stage = "tk-create"
    restore = install_sigint_observer(lambda: stage)
    stop = threading.Event()
    holder = {"driver": None}
    root = tk.Tk()
    root.title("Teste E - Tk + Edge + SharePoint")

    def worker() -> None:
        nonlocal stage
        try:
            stage = "webdriver.Edge"
            holder["driver"], _service = create_edge(args.driver_path)
            stage = "driver.get-sharepoint"
            navigate(holder["driver"], args.url, sanitized_url_label(args.url))
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
