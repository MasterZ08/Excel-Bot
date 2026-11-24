"""Random Excel Data Generator - entry point.  Run:  python main.py"""
import sys
import threading


def _fatal(msg: str) -> None:
    print(msg, file=sys.stderr)
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("Random Excel Data Generator", msg)
        root.destroy()
    except Exception:
        pass
    sys.exit(1)


def main() -> None:
    from config_manager import ConfigManager
    from utils import setup_logging

    cfg = ConfigManager()
    cfg.ensure_dirs()
    log = setup_logging(cfg.logs_dir)

    def _hook(args):  # uncaught exception in any thread -> log file, never a silent death
        log.error("Uncaught thread exception", exc_info=(args.exc_type, args.exc_value, args.exc_traceback))

    threading.excepthook = _hook
    sys.excepthook = lambda t, v, tb: log.error("Uncaught exception", exc_info=(t, v, tb))

    try:
        import customtkinter  # noqa: F401
    except ImportError:
        _fatal("CustomTkinter is not installed.\nRun:  pip install -r requirements.txt")
    from gui import App

    app = App(cfg)
    app.mainloop()


if __name__ == "__main__":
    main()
