from __future__ import annotations

import tkinter as tk

from .app import ReleaseNoteApp


def main() -> None:
    root = tk.Tk()
    style = root.tk.call("ttk::style", "theme", "use")
    _ = style
    ReleaseNoteApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
