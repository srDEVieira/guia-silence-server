from __future__ import annotations


def _is_benign_webview_shutdown(exc: Exception) -> bool:
    tb = exc.__traceback__
    has_webview_start = False
    has_winforms_join = False

    while tb is not None:
      filename = tb.tb_frame.f_code.co_filename.replace("\\", "/").lower()
      function_name = tb.tb_frame.f_code.co_name

      if filename.endswith("/src/webview_app.py") and function_name == "run_webview_app":
          has_webview_start = True
      if filename.endswith("/webview/platforms/winforms.py"):
          has_winforms_join = True

      tb = tb.tb_next

    return has_webview_start and has_winforms_join


def main() -> None:
    try:
        from src.webview_app import run_webview_app

        run_webview_app()
    except KeyboardInterrupt:
        return
    except Exception as exc:
        if _is_benign_webview_shutdown(exc):
            return
        raise
    except ModuleNotFoundError:
        from src.gui import DocumentGeneratorApp

        app = DocumentGeneratorApp()
        app.mainloop()


if __name__ == "__main__":
    main()
