from __future__ import annotations

import re
import threading
import traceback
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from tkinter import (
    BOTH,
    RIGHT,
    StringVar,
    Tk,
    W,
    filedialog,
    messagebox,
)
from tkinter import ttk

from csv_import import basement_floor_labels, parse_import_csv
from parking_balance import excel_round
from template_builder import (
    PARKING_RATIO_AREA,
    PARKING_RATIO_FIXED,
    BuildResult,
    ParkingSizeInput,
    ParkingSpace,
    create_template,
    expand_parking_sizes,
)

DEFAULT_OUTPUT_NAME = "試算面積-混公.xlsx"

# 四種車位的預設尺寸（寬 cm, 長 cm）。試算表表頭寫成「寬*長」。
PARKING_SIZE_COUNT = 4
PARKING_DEFAULTS: tuple[tuple[int, int], ...] = (
    (230, 550),
    (250, 550),
    (250, 600),
    (350, 600),
)
DEFAULT_PARKING_FLOORS = tuple(f"B{index}" for index in range(1, 7))


def next_parking_floor_label(existing: Sequence[str]) -> str:
    """下一個預設樓層名稱。已有 B3 時用 B4，不管中間有沒有空號。"""
    numbers: list[int] = []
    for label in existing:
        match = re.fullmatch(r"B(\d+)", label.strip())
        if match:
            numbers.append(int(match.group(1)))
    return f"B{max(numbers, default=0) + 1}"


def parse_parking_floor_labels(labels: Sequence[str]) -> list[str]:
    """樓層名稱去掉空白。空白或重複時失敗。"""
    cleaned: list[str] = []
    seen: set[str] = set()
    for raw in labels:
        label = raw.strip()
        if not label:
            raise ValueError("樓層名稱不可空白。")
        if label in seen:
            raise ValueError(f"樓層名稱重複：{label}")
        seen.add(label)
        cleaned.append(label)
    return cleaned


@dataclass
class ParkingFloorRow:
    label: StringVar
    flat_counts: list[StringVar]
    mechanical_counts: list[StringVar]
    label_entry: ttk.Entry
    flat_entries: list[ttk.Entry]
    mechanical_entries: list[ttk.Entry]
    mechanical_label: ttk.Label
    remove: ttk.Button


class TemplateBuilderApp:
    """選擇 CSV、輸入公設比並產生試算表的 tkinter 介面。"""

    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title("自動建立試算面積")
        root.minsize(760, 0)

        self.csv_path = StringVar()
        self.public_ratio = StringVar()
        self.output_path = StringVar()
        self.parking_ratio_mode = StringVar(value=PARKING_RATIO_AREA)
        self.mechanical_ratio = StringVar()

        # 四種尺寸只填一次（長、寬）。各層數量在平面、機械兩個表。
        self.parking_sizes: list[dict[str, StringVar]] = []
        for width, length in PARKING_DEFAULTS:
            self.parking_sizes.append(
                {
                    "width": StringVar(value=str(width)),
                    "length": StringVar(value=str(length)),
                }
            )
        self.parking_floors: list[ParkingFloorRow] = []

        self._build_widgets()

    def _build_widgets(self) -> None:
        padding = {"padx": 12, "pady": 6}

        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill=BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        # CSV 檔案
        ttk.Label(frame, text="CSV 檔案").grid(row=0, column=0, sticky=W, **padding)
        ttk.Entry(frame, textvariable=self.csv_path).grid(
            row=0, column=1, sticky="ew", **padding
        )
        ttk.Button(frame, text="瀏覽…", command=self._choose_csv).grid(
            row=0, column=2, **padding
        )

        # 輸出路徑
        ttk.Label(frame, text="輸出檔案").grid(row=1, column=0, sticky=W, **padding)
        ttk.Entry(frame, textvariable=self.output_path).grid(
            row=1, column=1, sticky="ew", **padding
        )
        ttk.Button(frame, text="另存為…", command=self._choose_output).grid(
            row=1, column=2, **padding
        )

        # 公設比
        ttk.Label(frame, text="公設比 (%)").grid(row=2, column=0, sticky=W, **padding)
        ratio_entry = ttk.Entry(frame, textvariable=self.public_ratio)
        ratio_entry.grid(row=2, column=1, sticky="ew", **padding)
        ttk.Label(frame, text="例如 32 代表 32%").grid(
            row=2, column=2, sticky=W, **padding
        )

        # 車位設定
        self._build_parking_section(frame, row=3, padding=padding)

        # 動作按鈕
        button_bar = ttk.Frame(frame)
        button_bar.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(12, 6))
        self.build_button = ttk.Button(
            button_bar, text="建立試算表", command=self._on_build_clicked
        )
        self.build_button.pack(side=RIGHT, padx=4)
        ttk.Button(button_bar, text="清除", command=self._reset).pack(
            side=RIGHT, padx=4
        )

        # 狀態列
        self.status = StringVar(value="請選擇 CSV 檔案並輸入公設比。")
        ttk.Label(
            frame, textvariable=self.status, foreground="#444", wraplength=640
        ).grid(row=5, column=0, columnspan=3, sticky=W, **padding)

    def _apply_selected_csv(self, path: Path) -> None:
        try:
            parse_import_csv(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("CSV 格式錯誤", str(exc))
            self.status.set(f"CSV 載入失敗：{path.name}")
            return
        self.status.set(f"已選擇 CSV：{path.name}")

    def _build_parking_section(
        self, parent: ttk.Frame, row: int, padding: dict[str, int]
    ) -> None:
        section = ttk.LabelFrame(parent, text="車位設定", padding=10)
        section.grid(row=row, column=0, columnspan=3, sticky="ew", **padding)
        section.columnconfigure(0, weight=1)

        grid = ttk.Frame(section)
        grid.grid(row=0, column=0, sticky="ew")
        self.parking_grid = grid
        grid.columnconfigure(0, weight=0)
        for col in range(1, PARKING_SIZE_COUNT + 1):
            grid.columnconfigure(col, weight=1)

        for index in range(PARKING_SIZE_COUNT):
            ttk.Label(grid, text=f"尺寸 {index + 1}").grid(
                row=0, column=index + 1, padx=6, pady=(0, 4), sticky=W
            )
        ttk.Label(grid, text="長 (cm)").grid(row=1, column=0, padx=6, pady=3, sticky=W)
        ttk.Label(grid, text="寬 (cm)").grid(row=2, column=0, padx=6, pady=3, sticky=W)
        for index, size in enumerate(self.parking_sizes):
            ttk.Entry(grid, textvariable=size["length"], width=8).grid(
                row=1, column=index + 1, padx=6, pady=3, sticky="ew"
            )
            ttk.Entry(grid, textvariable=size["width"], width=8).grid(
                row=2, column=index + 1, padx=6, pady=3, sticky="ew"
            )

        ttk.Label(grid, text="平面車位").grid(
            row=3, column=0, padx=6, pady=(12, 4), sticky=W
        )
        ttk.Button(grid, text="新增樓層", command=self._add_parking_floor).grid(
            row=3, column=PARKING_SIZE_COUNT + 1, padx=6, pady=(12, 4), sticky="e"
        )
        self._flat_floor_headings = self._build_floor_headings(grid)
        self._mechanical_title = ttk.Label(grid, text="機械車位")
        self._mechanical_floor_headings = self._build_floor_headings(grid)

        for floor in DEFAULT_PARKING_FLOORS:
            self._add_parking_floor(floor)

        mode_bar = ttk.Frame(section)
        mode_bar.grid(row=1, column=0, padx=6, pady=(10, 3), sticky=W)
        ttk.Label(mode_bar, text="比例算法").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(
            mode_bar,
            text="依實際面積",
            value=PARKING_RATIO_AREA,
            variable=self.parking_ratio_mode,
        ).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(
            mode_bar,
            text="固定比例",
            value=PARKING_RATIO_FIXED,
            variable=self.parking_ratio_mode,
        ).pack(side="left")

        ratio_fields = ttk.Frame(section)
        ratio_fields.grid(row=2, column=0, sticky="ew")
        ttk.Label(ratio_fields, text="機械位比例").grid(
            row=0, column=0, padx=6, pady=3, sticky=W
        )
        ttk.Entry(ratio_fields, textvariable=self.mechanical_ratio, width=8).grid(
            row=0, column=1, padx=6, pady=3, sticky=W
        )
        ttk.Label(
            ratio_fields,
            text="平面位為 1.00。兩邊都有數量時才需要填，例如 0.60。",
        ).grid(row=0, column=2, padx=6, pady=3, sticky=W)

    def _add_parking_floor(self, label: str | None = None) -> None:
        if label is None:
            label = next_parking_floor_label(
                [row.label.get() for row in self.parking_floors]
            )
        label_var = StringVar(value=label)
        flat_counts = [StringVar() for _ in range(PARKING_SIZE_COUNT)]
        mechanical_counts = [StringVar() for _ in range(PARKING_SIZE_COUNT)]
        label_entry = ttk.Entry(self.parking_grid, textvariable=label_var, width=10)
        flat_entries = [
            ttk.Entry(self.parking_grid, textvariable=count, width=8)
            for count in flat_counts
        ]
        mechanical_entries = [
            ttk.Entry(self.parking_grid, textvariable=count, width=8)
            for count in mechanical_counts
        ]
        mechanical_label = ttk.Label(self.parking_grid, textvariable=label_var, width=10)
        remove = ttk.Button(self.parking_grid, text="刪除")
        row = ParkingFloorRow(
            label=label_var,
            flat_counts=flat_counts,
            mechanical_counts=mechanical_counts,
            label_entry=label_entry,
            flat_entries=flat_entries,
            mechanical_entries=mechanical_entries,
            mechanical_label=mechanical_label,
            remove=remove,
        )
        remove.configure(command=lambda row=row: self._remove_parking_floor(row))
        self.parking_floors.append(row)
        self._layout_parking_floors()

    def _remove_parking_floor(self, row: ParkingFloorRow) -> None:
        self.parking_floors.remove(row)
        row.label_entry.destroy()
        row.remove.destroy()
        row.mechanical_label.destroy()
        for entry in row.flat_entries:
            entry.destroy()
        for entry in row.mechanical_entries:
            entry.destroy()
        self._layout_parking_floors()

    def _build_floor_headings(self, frame: ttk.Frame) -> tuple[ttk.Label, ...]:
        headings = [ttk.Label(frame, text="樓層")]
        headings.extend(ttk.Label(frame, text="數量") for _ in range(PARKING_SIZE_COUNT))
        return tuple(headings)

    def _layout_parking_floors(self) -> None:
        for heading in self._flat_floor_headings + self._mechanical_floor_headings:
            heading.grid_remove()
        self._mechanical_title.grid_remove()
        for row in self.parking_floors:
            row.label_entry.grid_remove()
            row.remove.grid_remove()
            row.mechanical_label.grid_remove()
            for entry in row.flat_entries + row.mechanical_entries:
                entry.grid_remove()

        floor_count = len(self.parking_floors)
        # 長寬占 0–2 列，平面標題在第 3 列。樓層表從第 4 列接下去。
        self._mechanical_title.grid(
            row=5 + floor_count,
            column=0,
            columnspan=PARKING_SIZE_COUNT + 1,
            padx=6,
            pady=(12, 4),
            sticky=W,
        )
        if floor_count == 0:
            return
        for column, heading in enumerate(self._flat_floor_headings):
            heading.grid(row=4, column=column, padx=6, pady=(0, 4), sticky=W)
        for index, row in enumerate(self.parking_floors):
            grid_row = 5 + index
            row.label_entry.grid(row=grid_row, column=0, padx=6, pady=3, sticky="ew")
            for column, entry in enumerate(row.flat_entries, start=1):
                entry.grid(row=grid_row, column=column, padx=6, pady=3, sticky="ew")
            row.remove.grid(
                row=grid_row, column=PARKING_SIZE_COUNT + 1, padx=6, pady=3, sticky=W
            )
        mechanical_heading_row = 6 + floor_count
        for column, heading in enumerate(self._mechanical_floor_headings):
            heading.grid(
                row=mechanical_heading_row, column=column, padx=6, pady=(0, 4), sticky=W
            )
        for index, row in enumerate(self.parking_floors):
            grid_row = mechanical_heading_row + 1 + index
            row.mechanical_label.grid(row=grid_row, column=0, padx=6, pady=3, sticky=W)
            for column, entry in enumerate(row.mechanical_entries, start=1):
                entry.grid(row=grid_row, column=column, padx=6, pady=3, sticky="ew")

    def _choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="選擇試算表 CSV",
            filetypes=[("CSV 檔案", "*.csv"), ("所有檔案", "*.*")],
        )
        if not path:
            return
        self.csv_path.set(path)
        if not self.output_path.get():
            default_output = Path(path).with_name(DEFAULT_OUTPUT_NAME)
            self.output_path.set(str(default_output))
        self._apply_selected_csv(Path(path))

    def _choose_output(self) -> None:
        initial = self.output_path.get() or DEFAULT_OUTPUT_NAME
        initial_path = Path(initial)
        path = filedialog.asksaveasfilename(
            title="輸出試算表另存為",
            defaultextension=".xlsx",
            initialfile=initial_path.name,
            initialdir=str(initial_path.parent) if initial_path.parent != Path() else None,
            filetypes=[("Excel 活頁簿", "*.xlsx"), ("所有檔案", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _reset(self) -> None:
        self.csv_path.set("")
        self.public_ratio.set("")
        self.output_path.set("")
        for index, size in enumerate(self.parking_sizes):
            width, length = PARKING_DEFAULTS[index]
            size["width"].set(str(width))
            size["length"].set(str(length))
        for row in list(self.parking_floors):
            self._remove_parking_floor(row)
        for floor in DEFAULT_PARKING_FLOORS:
            self._add_parking_floor(floor)
        self.parking_ratio_mode.set(PARKING_RATIO_AREA)
        self.mechanical_ratio.set("")
        self.status.set("請選擇 CSV 檔案並輸入公設比。")

    def _parse_count(self, text: str, label: str, kind: str) -> int | None:
        if not text:
            return 0
        try:
            count = int(text)
        except ValueError:
            messagebox.showerror(
                "車位數量格式錯誤", f"{label} 的{kind}數量必須是整數。"
            )
            return None
        if count < 0:
            messagebox.showerror(
                "車位數量無效", f"{label} 的{kind}數量不可為負數。"
            )
            return None
        return count

    def _collect_parking_spaces(self) -> tuple[list[ParkingSpace], list[str]] | None:
        """讀取車位輸入並展開成試算表欄；驗證失敗回傳 None。"""
        try:
            floors = parse_parking_floor_labels(
                [row.label.get() for row in self.parking_floors]
            )
        except ValueError as exc:
            messagebox.showerror("樓層名稱無效", str(exc))
            return None

        sizes: list[ParkingSizeInput] = []
        for index, size in enumerate(self.parking_sizes):
            length_text = size["length"].get().strip()
            width_text = size["width"].get().strip()
            label = f"尺寸 {index + 1}"
            if not length_text or not width_text:
                messagebox.showerror("車位尺寸不完整", f"{label} 需同時填寫長與寬。")
                return None
            try:
                length = float(length_text)
                width = float(width_text)
            except ValueError:
                messagebox.showerror("車位尺寸格式錯誤", f"{label} 的長與寬必須是數字。")
                return None
            if length <= 0 or width <= 0:
                messagebox.showerror("車位尺寸無效", f"{label} 的長與寬必須大於 0。")
                return None

            flat_counts: list[int] = []
            mechanical_counts: list[int] = []
            for floor_row, floor_label in zip(self.parking_floors, floors, strict=True):
                flat_count = self._parse_count(
                    floor_row.flat_counts[index].get().strip(), floor_label, "平面"
                )
                if flat_count is None:
                    return None
                mechanical_count = self._parse_count(
                    floor_row.mechanical_counts[index].get().strip(),
                    floor_label,
                    "機械",
                )
                if mechanical_count is None:
                    return None
                flat_counts.append(flat_count)
                mechanical_counts.append(mechanical_count)
            sizes.append(
                ParkingSizeInput(
                    width=width,
                    length=length,
                    flat_counts=tuple(flat_counts),
                    mechanical_counts=tuple(mechanical_counts),
                )
            )

        needs_ratio = any(
            count > 0 for size in sizes for count in (size.flat_counts or ())
        ) and any(
            count > 0 for size in sizes for count in (size.mechanical_counts or ())
        )
        mechanical_ratio: float | None = None
        if needs_ratio:
            ratio_text = self.mechanical_ratio.get().strip()
            if not ratio_text:
                messagebox.showerror(
                    "尚未填寫機械位比例",
                    "平面與機械都有數量時，需填寫機械位比例（例如 0.60）。",
                )
                return None
            try:
                mechanical_ratio = excel_round(float(ratio_text), 2)
            except ValueError:
                messagebox.showerror(
                    "機械位比例格式錯誤",
                    "機械位比例必須是數字，並以小數兩位計算（例如 0.60）。",
                )
                return None
            if mechanical_ratio <= 0:
                messagebox.showerror("機械位比例無效", "機械位比例必須大於 0。")
                return None

        return (
            expand_parking_sizes(
                sizes,
                mode=self.parking_ratio_mode.get(),
                mechanical_ratio=mechanical_ratio,
            ),
            floors,
        )

    def _validate(
        self,
    ) -> tuple[Path, float | None, Path, list[ParkingSpace], list[str]] | None:
        csv_text = self.csv_path.get().strip()
        if not csv_text:
            messagebox.showwarning("尚未選擇 CSV", "請先選擇一個 CSV 檔案。")
            return None
        csv_file = Path(csv_text)
        if not csv_file.exists():
            messagebox.showerror("找不到檔案", f"找不到 CSV：\n{csv_file}")
            return None

        ratio_text = self.public_ratio.get().strip()
        ratio: float | None = None
        if ratio_text:
            try:
                ratio = float(ratio_text)
            except ValueError:
                messagebox.showerror(
                    "公設比格式錯誤", "公設比必須是數字（例如 32 或 32.5）。"
                )
                return None
            if ratio < 0:
                messagebox.showerror("公設比無效", "公設比不可為負數。")
                return None
            if ratio >= 100:
                messagebox.showerror("公設比無效", "公設比必須小於 100。")
                return None

        parking = self._collect_parking_spaces()
        if parking is None:
            return None
        parking_spaces, parking_floors = parking

        output_text = self.output_path.get().strip()
        if output_text:
            output_file = Path(output_text)
        else:
            output_file = csv_file.with_name(DEFAULT_OUTPUT_NAME)
            self.output_path.set(str(output_file))

        return csv_file, ratio, output_file, parking_spaces, parking_floors

    def _confirm_basement_floors(self, csv_file: Path, parking_floors: list[str]) -> bool:
        """地下室樓層與介面不一致時請使用者確認。

        依 CSV 出現順序、區分大小寫比對，數量全為 0 的樓層也算。
        吻合，或使用者選擇仍要建立時回傳 True。不改寫任一邊的名稱。
        """
        try:
            import_data = parse_import_csv(csv_file)
        except (OSError, ValueError) as exc:
            messagebox.showerror("CSV 格式錯誤", str(exc))
            self.status.set(f"CSV 載入失敗：{csv_file.name}")
            return False

        csv_floors = basement_floor_labels(import_data.public_areas)
        if csv_floors == parking_floors:
            return True

        def _format(floors: list[str]) -> str:
            return "、".join(floors) if floors else "（無）"

        return bool(
            messagebox.askyesno(
                "地下室樓層不一致",
                "CSV 公設的地下室樓層與畫面上的車位樓層不一致，不會自動改寫。\n\n"
                f"CSV：{_format(csv_floors)}\n"
                f"介面：{_format(parking_floors)}\n\n"
                "仍要建立試算表嗎？",
                default=messagebox.NO,
            )
        )

    def _on_build_clicked(self) -> None:
        validated = self._validate()
        if validated is None:
            return
        csv_file, ratio, output_file, parking_spaces, parking_floors = validated
        if not self._confirm_basement_floors(csv_file, parking_floors):
            return

        self.build_button.configure(state="disabled")
        self.status.set("建立中…")
        thread = threading.Thread(
            target=self._build_worker,
            args=(csv_file, ratio, output_file, parking_spaces, parking_floors),
            daemon=True,
        )
        thread.start()

    def _build_worker(
        self,
        csv_file: Path,
        ratio: float | None,
        output_file: Path,
        parking_spaces: list[ParkingSpace],
        parking_floors: list[str] | None,
    ) -> None:
        try:
            import_data = parse_import_csv(csv_file.resolve())
            result = create_template(
                output_file.resolve(),
                unit_data=import_data.units,
                public_areas=import_data.public_areas,
                public_ratio=ratio,
                parking_spaces=parking_spaces,
                parking_floors=parking_floors,
                include_porch=import_data.include_porch,
                include_terrace=import_data.include_terrace,
            )
        except PermissionError:
            self.root.after(0, self._on_permission_error, output_file)
        except Exception as exc:  # noqa: BLE001 - 回報給使用者
            traceback.print_exc()
            self.root.after(0, self._on_build_error, exc)
        else:
            self.root.after(
                0,
                self._on_build_success,
                result,
                len(import_data.units),
                len(import_data.public_areas),
                import_data.include_porch,
                import_data.include_terrace,
            )

    def _on_build_success(
        self,
        result: BuildResult,
        unit_count: int,
        public_count: int,
        include_porch: bool = False,
        include_terrace: bool = False,
    ) -> None:
        self.build_button.configure(state="normal")
        porch = "、含門廊" if include_porch else ""
        terrace = "、含露台" if include_terrace else ""
        summary = f"（{unit_count} 戶、{public_count} 筆公設{porch}{terrace}）"
        incomplete = result.incomplete_adjustments
        if incomplete:
            self.status.set(f"完成（部分持分未調整）：{result.path.name}{summary}")
            details = "\n".join(
                f"・{item.name}：目標 {item.target}、實際 {item.actual}"
                f"（差 {item.actual - item.target}）"
                for item in incomplete
            )
            messagebox.showwarning(
                "持分調整未完成",
                "試算表已產生，但下列持分無法自動調整到目標，需人工處理：\n\n"
                f"{details}",
            )
        else:
            self.status.set(f"完成：{result.path.name}{summary}")
            messagebox.showinfo(
                "建立完成",
                "已產生試算表，所有持分均已調整到目標。",
            )

    def _on_permission_error(self, output_file: Path) -> None:
        self.build_button.configure(state="normal")
        self.status.set("無法寫入檔案，請先關閉已開啟的 Excel 檔案再試一次。")
        messagebox.showerror(
            "無法寫入檔案",
            f"無法寫入：\n{output_file}\n\n"
            "該檔案可能正在 Excel 中開啟，請先關閉後再建立一次。",
        )

    def _on_build_error(self, exc: Exception) -> None:
        self.build_button.configure(state="normal")
        self.status.set("建立失敗，請檢查 CSV 內容或設定。")
        messagebox.showerror("建立失敗", str(exc))


def main() -> None:
    root = Tk()
    TemplateBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
