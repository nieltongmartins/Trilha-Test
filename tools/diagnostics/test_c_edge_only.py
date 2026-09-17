"""TESTE C: Selenium Edge sem Tkinter; abre about:blank."""
import argparse
import time
from probe_common import (LOGGER, add_common_edge_arguments, configure, create_edge,
                          install_sigint_observer, navigate, safe_quit)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    add_common_edge_arguments(parser)
    args = parser.parse_args()
    configure("teste-c")
    stage = "before-webdriver"
    restore = install_sigint_observer(lambda: stage)
    driver = None
    try:
        stage = "webdriver.Edge"
        driver, _service = create_edge(args.driver_path)
        stage = "driver.get-about-blank"
        navigate(driver, "about:blank", "about:blank")
        stage = "hold"
        LOGGER.info("estágio=hold segundos=%s", args.hold_seconds)
        time.sleep(args.hold_seconds)
        return 0
    except KeyboardInterrupt:
        LOGGER.exception("KeyboardInterrupt preservado estágio=%s", stage)
        raise
    finally:
        stage = "driver.quit"
        safe_quit(driver)
        restore()


if __name__ == "__main__":
    raise SystemExit(main())
