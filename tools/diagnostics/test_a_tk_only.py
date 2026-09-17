"""TESTE A: somente Python + Tkinter."""
import tkinter as tk
from probe_common import LOGGER, configure, install_sigint_observer


def main() -> int:
    configure("teste-a")
    stage = "tk-create"
    restore = install_sigint_observer(lambda: stage)
    try:
        root = tk.Tk()
        root.title("Teste A - Tk somente")
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        stage = "tk-mainloop"
        LOGGER.info("estágio=%s", stage)
        root.mainloop()
        LOGGER.info("mainloop retornou")
        return 0
    except KeyboardInterrupt:
        LOGGER.exception("KeyboardInterrupt preservado")
        raise
    finally:
        restore()


if __name__ == "__main__":
    raise SystemExit(main())
