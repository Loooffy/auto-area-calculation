from __future__ import annotations

from pathlib import Path

from tkinter import Tk
from tkinter import ttk

from gui import TemplateBuilderApp


def _labelframe(widget: object, title: str) -> ttk.LabelFrame | None:
    for child in widget.winfo_children():  # type: ignore[attr-defined]
        if isinstance(child, ttk.LabelFrame) and child.cget("text") == title:
            return child
        found = _labelframe(child, title)
        if found is not None:
            return found
    return None


def test_porch_option_is_a_checkbox_only(tmp_path: Path) -> None:
    plain = tmp_path / "plain.csv"
    plain.write_text(",主建物,陽台\n1F-A,10,2\n", encoding="utf-8-sig")
    with_porch = tmp_path / "porch.csv"
    with_porch.write_text(",主建物,陽台,門廊\n1F-A,10,2,1.5\n", encoding="utf-8-sig")

    root = Tk()
    try:
        root.withdraw()
        app = TemplateBuilderApp(root)
        section = _labelframe(root, "門廊")
        assert section is not None
        assert not any(isinstance(child, ttk.Entry) for child in section.winfo_children())
        assert app.include_porch.get() is False

        app._sync_porch_option(with_porch)
        assert app.include_porch.get() is True
        app._sync_porch_option(plain)
        assert app.include_porch.get() is False
    finally:
        root.destroy()
