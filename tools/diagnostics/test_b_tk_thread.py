"""TESTE B: Tkinter + thread normal, sem Selenium."""
import threading
import tkinter as tk
from probe_common import LOGGER, configure, install_sigint_observer


def main() -> int:
    configure("teste-b")
    stage = "tk-create"
    restore = install_sigint_observer(lambda: stage)
    stop = threading.Event()
    worker = threading.Thread(target=stop.wait, name="probe-worker", daemon=False)
    root = tk.Tk()
    root.title("Teste B - Tk + thread")
    root.protocol("WM_DELETE_WINDOW", lambda: (stop.set(), root.destroy()))
    worker.start()
    try:
        stage = "tk-mainloop-with-thread"
        LOGGER.info("estágio=%s worker_ident=%s", stage, worker.ident)
        root.mainloop()
        return 0
    except KeyboardInterrupt:
        LOGGER.exception("KeyboardInterrupt preservado")
        raise
    finally:
        stop.set()
        worker.join(timeout=10)
        LOGGER.info("worker_alive=%s", worker.is_alive())
        restore()


if __name__ == "__main__":
    raise SystemExit(main())
