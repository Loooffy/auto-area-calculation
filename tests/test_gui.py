from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

import openpyxl
import pytest
from tkinter import Tk
from tkinter import ttk

from gui import (
    DEFAULT_PARKING_FLOORS,
    PARKING_DEFAULTS,
    TemplateBuilderApp,
    next_parking_floor_label,
    parse_parking_floor_labels,
)
from template_builder import PARKING_RATIO_AREA, PARKING_RATIO_FIXED, BuildResult


def _labelframe(widget: object, title: str) -> ttk.LabelFrame | None:
    for child in widget.winfo_children():  # type: ignore[attr-defined]
        if isinstance(child, ttk.LabelFrame) and child.cget("text") == title:
            return child
        found = _labelframe(child, title)
        if found is not None:
            return found
    return None


PORCH_CSV = """\
,主建物,陽台,門廊
1F-A,10,2,1.5
1F-B,20,3,0
1F-A1,8,1,2
2F-A1,30,4,5
10F-A,40,1,6
10~14F-A1,50,2,7
"""


@contextmanager
def _app():
    root = Tk()
    try:
        root.withdraw()
        yield TemplateBuilderApp(root)
    finally:
        root.destroy()


def _write_csv(directory: Path, text: str, name: str = "units.csv") -> Path:
    path = directory / name
    path.write_text(text, encoding="utf-8-sig")
    return path


def test_gui_has_no_porch_section() -> None:
    with _app() as app:
        assert _labelframe(app.root, "門廊") is None


def test_select_csv_sets_status(tmp_path: Path) -> None:
    path = _write_csv(tmp_path, PORCH_CSV)

    with _app() as app:
        app._apply_selected_csv(path)
        assert app.status.get() == "已選擇 CSV：units.csv"


def test_build_uses_csv_porch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("gui.messagebox.showinfo", lambda *_args, **_kwargs: None)
    monkeypatch.setattr("gui.messagebox.showwarning", lambda *_args, **_kwargs: None)

    def _fail_on_error(*args: object, **_kwargs: object) -> None:
        raise AssertionError(args)

    monkeypatch.setattr("gui.messagebox.showerror", _fail_on_error)
    path = _write_csv(tmp_path, PORCH_CSV)
    output = tmp_path / "out.xlsx"

    with _app() as app:
        app._apply_selected_csv(path)
        app._build_worker(path, None, output, [], None)
        app.root.update()

        assert app.status.get() == (
            "完成：out.xlsx（6 戶、0 筆公設、含門廊）"
        )

    workbook = openpyxl.load_workbook(output, data_only=False)
    sheet = workbook.active
    assert sheet["E2"].value == "門廊"
    assert sheet["E5"].value == 1.5
    assert sheet["E6"].value == 0
    assert sheet["E7"].value == 2
    assert sheet["E8"].value == 5
    assert sheet["E9"].value == 6
    assert sheet["E10"].value == 7
    assert sheet["F5"].value == "=C5+D5+E5"
    workbook.close()


def test_build_skips_porch_when_all_zero(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("gui.messagebox.showinfo", lambda *_args, **_kwargs: None)
    monkeypatch.setattr("gui.messagebox.showerror", lambda *_args, **_kwargs: None)
    path = _write_csv(
        tmp_path, ",主建物,陽台,門廊\n1F-A,10,2,0\n1F-B,8,1,\n", "zeros.csv"
    )
    output = tmp_path / "out.xlsx"

    with _app() as app:
        app._build_worker(path, None, output, [], None)
        app.root.update()
        assert "含門廊" not in app.status.get()

    workbook = openpyxl.load_workbook(output, data_only=False)
    sheet = workbook.active
    assert sheet["E2"].value != "門廊"
    workbook.close()


def test_success_status_without_porch_warning(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("gui.messagebox.showinfo", lambda *_args, **_kwargs: None)

    with _app() as app:
        app._on_build_success(BuildResult(path=tmp_path / "試算.xlsx"), 2, 1, False)
        assert app.status.get() == "完成：試算.xlsx（2 戶、1 筆公設）"


def test_reset_clears_paths(tmp_path: Path) -> None:
    path = _write_csv(tmp_path, PORCH_CSV)

    with _app() as app:
        app.csv_path.set(str(path))
        app._apply_selected_csv(path)
        app._reset()
        assert app.csv_path.get() == ""
        assert app.status.get() == "請選擇 CSV 檔案並輸入公設比。"


def test_unreadable_csv_shows_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    messages: list[str] = []
    monkeypatch.setattr(
        "gui.messagebox.showerror", lambda _title, message: messages.append(message)
    )
    good = _write_csv(tmp_path, PORCH_CSV, "good.csv")
    bad = _write_csv(tmp_path, "不是各戶資料\n", "bad.csv")

    with _app() as app:
        app._apply_selected_csv(good)
        app._apply_selected_csv(bad)
        assert app.status.get() == "CSV 載入失敗：bad.csv"
        assert messages
        assert "欄標題不符" in messages[-1]
        assert messages[-1].count("欄標題不符") == 1
        assert "主建物" in messages[-1]


def test_invalid_label_csv_shows_all_errors(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    messages: list[str] = []
    monkeypatch.setattr(
        "gui.messagebox.showerror", lambda _title, message: messages.append(message)
    )
    bad = _write_csv(
        tmp_path,
        ",主建物,陽台\n3F~6F-A,10,0\n2F-AB,8,0\n公設,樓別,面積\n,3~6F*3,100\n",
        "bad-labels.csv",
    )

    with _app() as app:
        app._apply_selected_csv(bad)
        assert app.status.get() == "CSV 載入失敗：bad-labels.csv"
        assert messages
        assert "3F~6F-A" in messages[-1]
        assert "2F-AB" in messages[-1]
        assert "層數應為 4" in messages[-1]


def test_public_ratio_rejects_over_100(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    messages: list[tuple[str, str]] = []
    monkeypatch.setattr(
        "gui.messagebox.showerror",
        lambda title, message: messages.append((title, message)),
    )
    path = _write_csv(tmp_path, PORCH_CSV)

    with _app() as app:
        app.csv_path.set(str(path))
        app.public_ratio.set("100")
        assert app._validate() is None
        assert messages[-1] == ("公設比無效", "公設比必須小於 100。")

        messages.clear()
        app.public_ratio.set("100.1")
        assert app._validate() is None
        assert messages[-1] == ("公設比無效", "公設比必須小於 100。")

        messages.clear()
        app.public_ratio.set("-1")
        assert app._validate() is None
        assert messages[-1] == ("公設比無效", "公設比不可為負數。")

        messages.clear()
        app.public_ratio.set("32.5")
        validated = app._validate()
        assert validated is not None
        assert validated[1] == 32.5
        assert not messages


def test_parking_floor_labels() -> None:
    assert parse_parking_floor_labels([" B1 ", "地下一"]) == ["B1", "地下一"]
    with pytest.raises(ValueError, match="空白"):
        parse_parking_floor_labels(["B1", "  "])
    with pytest.raises(ValueError, match="重複"):
        parse_parking_floor_labels(["B1", "B1"])
    assert next_parking_floor_label(["B1", "B2", "屋頂"]) == "B3"


def test_parking_form_collects_counts_by_floor() -> None:
    root = Tk()
    try:
        root.withdraw()
        app = TemplateBuilderApp(root)
        section = _labelframe(root, "車位設定")
        assert section is not None
        labels: list[str] = []

        def _collect_labels(widget: object) -> None:
            for child in widget.winfo_children():  # type: ignore[attr-defined]
                if isinstance(child, ttk.Label):
                    labels.append(str(child.cget("text")))
                _collect_labels(child)

        _collect_labels(section)
        assert "平面車位" in labels
        assert "機械車位" in labels
        assert "長 (cm)" in labels
        assert "寬 (cm)" in labels
        assert app.parking_ratio_mode.get() == PARKING_RATIO_AREA
        assert len(app.parking_sizes) == len(PARKING_DEFAULTS)
        assert [row.label.get() for row in app.parking_floors] == list(DEFAULT_PARKING_FLOORS)
        assert app.parking_sizes[0]["length"].get() == "550"
        assert app.parking_sizes[0]["width"].get() == "230"

        app.parking_floors[0].flat_counts[0].set("10")
        app.parking_floors[1].flat_counts[0].set("2")
        app.parking_floors[0].mechanical_counts[0].set("4")
        app.mechanical_ratio.set("0.6")
        app.parking_ratio_mode.set(PARKING_RATIO_FIXED)
        collected = app._collect_parking_spaces()
        assert collected is not None
        columns, floors = collected
        assert floors == list(DEFAULT_PARKING_FLOORS)
        assert len(columns) == 2
        assert [column.count for column in columns] == [12, 4]
        assert columns[0].floor_counts == (10, 2, 0, 0, 0, 0)
        assert columns[1].floor_counts == (4, 0, 0, 0, 0, 0)
        assert [column.share_ratio for column in columns] == [1, 0.6]
        assert columns[0].label.endswith("(平面)")
        assert columns[1].label.endswith("(機械)")
        assert all("無障礙" not in column.label for column in columns)

        app._add_parking_floor()
        assert app.parking_floors[-1].label.get() == "B7"
        removed = app.parking_floors[2]
        app._remove_parking_floor(removed)
        assert [row.label.get() for row in app.parking_floors] == [
            "B1",
            "B2",
            "B4",
            "B5",
            "B6",
            "B7",
        ]
    finally:
        root.destroy()


def _set_parking_floors(app: TemplateBuilderApp, labels: list[str]) -> None:
    while len(app.parking_floors) > len(labels):
        app._remove_parking_floor(app.parking_floors[-1])
    while len(app.parking_floors) < len(labels):
        app._add_parking_floor()
    for row, label in zip(app.parking_floors, labels, strict=True):
        row.label.set(label)


BASEMENT_CSV = """\
,主建物,陽台
1F-A,10,0
公設,樓別,面積
,1F,100
,屋突一層,50
,B1,200
,B2,200
"""


def test_basement_mismatch_asks_without_rewriting(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    prompts: list[tuple[str, str]] = []

    def _ask(title: str, message: str, **_kwargs: object) -> bool:
        prompts.append((title, message))
        return True

    monkeypatch.setattr("gui.messagebox.askyesno", _ask)
    path = _write_csv(tmp_path, BASEMENT_CSV, "basement.csv")

    with _app() as app:
        _set_parking_floors(app, ["B2", "B1"])
        assert all(
            not count.get()
            for row in app.parking_floors
            for count in row.flat_counts + row.mechanical_counts
        )
        before = [row.label.get() for row in app.parking_floors]
        assert app._confirm_basement_floors(path, ["B2", "B1"]) is True
        assert [row.label.get() for row in app.parking_floors] == before

    assert prompts
    title, message = prompts[-1]
    assert title == "地下室樓層不一致"
    assert "CSV：B1、B2" in message
    assert "畫面：B2、B1" in message
    assert "不會自動改寫" in message


def test_basement_match_skips_dialog(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def _fail(*_args: object, **_kwargs: object) -> bool:
        raise AssertionError("吻合時不應詢問")

    monkeypatch.setattr("gui.messagebox.askyesno", _fail)
    path = _write_csv(tmp_path, BASEMENT_CSV, "basement.csv")

    with _app() as app:
        assert app._confirm_basement_floors(path, ["B1", "B2"]) is True


def test_basement_comparison_is_case_sensitive(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    prompts: list[str] = []
    monkeypatch.setattr(
        "gui.messagebox.askyesno",
        lambda _title, message, **_kwargs: prompts.append(message) or False,
    )
    path = _write_csv(
        tmp_path,
        ",主建物,陽台\n1F-A,10,0\n公設,樓別,面積\n,1F,100\n,b1,20\n",
        "lower.csv",
    )

    with _app() as app:
        assert app._confirm_basement_floors(path, ["B1"]) is False

    assert "CSV：b1" in prompts[-1]
    assert "畫面：B1" in prompts[-1]


def test_empty_csv_basements_still_ask(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    prompts: list[str] = []
    monkeypatch.setattr(
        "gui.messagebox.askyesno",
        lambda _title, message, **_kwargs: prompts.append(message) or False,
    )
    path = _write_csv(tmp_path, ",主建物,陽台\n1F-A,10,0\n", "units-only.csv")

    with _app() as app:
        assert app._confirm_basement_floors(path, ["B1"]) is False

    assert "CSV：（無）" in prompts[-1]
    assert "畫面：B1" in prompts[-1]


def test_both_sides_without_basements_match(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    def _fail(*_args: object, **_kwargs: object) -> bool:
        raise AssertionError("兩邊都沒有地下室時不應詢問")

    monkeypatch.setattr("gui.messagebox.askyesno", _fail)
    path = _write_csv(
        tmp_path,
        ",主建物,陽台\n1F-A,10,0\n公設,樓別,面積\n,1F,100\n,屋突一層,50\n",
        "no-basement.csv",
    )

    with _app() as app:
        assert app._confirm_basement_floors(path, []) is True


def test_build_stops_when_basement_mismatch_declined(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("gui.messagebox.askyesno", lambda *_args, **_kwargs: False)
    started: list[object] = []

    class _Thread:
        def __init__(self, *args: object, **kwargs: object) -> None:
            started.append((args, kwargs))

        def start(self) -> None:
            started.append("start")

    monkeypatch.setattr("gui.threading.Thread", _Thread)
    path = _write_csv(tmp_path, BASEMENT_CSV, "basement.csv")

    with _app() as app:
        app.csv_path.set(str(path))
        _set_parking_floors(app, ["B1"])
        app._on_build_clicked()
        assert str(app.build_button.cget("state")) == "normal"

    assert started == []
