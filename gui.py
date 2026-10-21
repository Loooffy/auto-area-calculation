from __future__ import annotations

import threading
import traceback
from pathlib import Path
from tkinter import (
    BOTH,
    DISABLED,
    NORMAL,
    RIGHT,
    StringVar,
    Tk,
    W,
    filedialog,
    messagebox,
)
from tkinter import ttk

from csv_import import parse_import_csv
from template_builder import BuildResult, ParkingSpace, create_template

DEFAULT_OUTPUT_NAME = "試算面積-混公.xlsx"

# 四種車位的預設尺寸（寬 cm, 長 cm, 備註）
PARKING_DEFAULTS: tuple[tuple[int, int, str | None], ...] = (
    (230, 550, None),
    (250, 550, None),
    (250, 600, None),
    (350, 600, "無障礙"),
)


class TemplateBuilderApp:
    """選擇 CSV、輸入公設比並產生試算表的 tkinter 介面。"""

    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title("自動建立試算面積")
        root.minsize(560, 0)

        self.csv_path = StringVar()
        self.public_ratio = StringVar()
        self.output_path = StringVar()

        # 每種車位的 寬 / 長 / 數量 輸入欄
        self.parking_rows: list[dict[str, object]] = []
        for width, length, note in PARKING_DEFAULTS:
            self.parking_rows.append(
                {
                    "width": StringVar(value=str(width)),
                    "length": StringVar(value=str(length)),
                    "count": StringVar(),
                    "note": note,
                }
            )

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
            frame, textvariable=self.status, foreground="#444", wraplength=520
        ).grid(row=5, column=0, columnspan=3, sticky=W, **padding)

    def _build_parking_section(
        self, parent: ttk.Frame, row: int, padding: dict[str, int]
    ) -> None:
        section = ttk.LabelFrame(parent, text="車位設定（尺寸與數量）", padding=10)
        section.grid(row=row, column=0, columnspan=3, sticky="ew", **padding)
        for col in range(4):
            section.columnconfigure(col, weight=1)

        headers = ("車位類別", "寬 (cm)", "長 (cm)", "數量")
        for col, text in enumerate(headers):
            ttk.Label(section, text=text).grid(
                row=0, column=col, padx=6, pady=(0, 4), sticky=W
            )

        for idx, row_vars in enumerate(self.parking_rows):
            note = row_vars["note"]
            name = f"車位 {idx + 1}"
            if note:
                name += f"（{note}）"
            ttk.Label(section, text=name).grid(
                row=idx + 1, column=0, padx=6, pady=3, sticky=W
            )
            ttk.Entry(section, textvariable=row_vars["width"], width=10).grid(
                row=idx + 1, column=1, padx=6, pady=3, sticky="ew"
            )
            ttk.Entry(section, textvariable=row_vars["length"], width=10).grid(
                row=idx + 1, column=2, padx=6, pady=3, sticky="ew"
            )
            ttk.Entry(section, textvariable=row_vars["count"], width=10).grid(
                row=idx + 1, column=3, padx=6, pady=3, sticky="ew"
            )

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
        self.status.set(f"已選擇 CSV：{Path(path).name}")

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
        for idx, row_vars in enumerate(self.parking_rows):
            width, length, _ = PARKING_DEFAULTS[idx]
            row_vars["width"].set(str(width))
            row_vars["length"].set(str(length))
            row_vars["count"].set("")
        self.status.set("請選擇 CSV 檔案並輸入公設比。")

    def _collect_parking_spaces(self) -> list[ParkingSpace | None] | None:
        """讀取車位輸入，回傳依欄位順序排列的清單；驗證失敗回傳 None。"""
        spaces: list[ParkingSpace | None] = []
        for idx, row_vars in enumerate(self.parking_rows):
            width_text = row_vars["width"].get().strip()
            length_text = row_vars["length"].get().strip()
            count_text = row_vars["count"].get().strip()
            label = f"車位 {idx + 1}"

            if not width_text and not length_text and not count_text:
                spaces.append(None)
                continue

            if not width_text or not length_text:
                messagebox.showerror(
                    "車位尺寸不完整", f"{label} 需同時填寫寬與長。"
                )
                return None
            try:
                width = float(width_text)
                length = float(length_text)
            except ValueError:
                messagebox.showerror(
                    "車位尺寸格式錯誤", f"{label} 的寬與長必須是數字。"
                )
                return None
            if width <= 0 or length <= 0:
                messagebox.showerror(
                    "車位尺寸無效", f"{label} 的寬與長必須大於 0。"
                )
                return None

            count = 0
            if count_text:
                try:
                    count = int(count_text)
                except ValueError:
                    messagebox.showerror(
                        "車位數量格式錯誤", f"{label} 的數量必須是整數。"
                    )
                    return None
                if count < 0:
                    messagebox.showerror(
                        "車位數量無效", f"{label} 的數量不可為負數。"
                    )
                    return None

            spaces.append(
                ParkingSpace(
                    width=width,
                    length=length,
                    count=count,
                    note=row_vars["note"],
                )
            )
        return spaces

    def _validate(
        self,
    ) -> tuple[Path, float | None, Path, list[ParkingSpace | None]] | None:
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

        parking_spaces = self._collect_parking_spaces()
        if parking_spaces is None:
            return

        output_text = self.output_path.get().strip()
        if output_text:
            output_file = Path(output_text)
        else:
            output_file = csv_file.with_name(DEFAULT_OUTPUT_NAME)
            self.output_path.set(str(output_file))

        return csv_file, ratio, output_file, parking_spaces

    def _on_build_clicked(self) -> None:
        validated = self._validate()
        if validated is None:
            return
        csv_file, ratio, output_file, parking_spaces = validated

        self.build_button.configure(state=DISABLED)
        self.status.set("建立中…")
        thread = threading.Thread(
            target=self._build_worker,
            args=(csv_file, ratio, output_file, parking_spaces),
            daemon=True,
        )
        thread.start()

    def _build_worker(
        self,
        csv_file: Path,
        ratio: float | None,
        output_file: Path,
        parking_spaces: list[ParkingSpace | None],
    ) -> None:
        try:
            import_data = parse_import_csv(csv_file.resolve())
            result = create_template(
                output_file.resolve(),
                unit_data=import_data.units,
                public_areas=import_data.public_areas,
                public_ratio=ratio,
                parking_spaces=parking_spaces,
            )
        except PermissionError:
            self.root.after(0, self._on_permission_error, output_file)
        except Exception as exc:  # noqa: BLE001 - 回報給使用者
            traceback.print_exc()
            self.root.after(0, self._on_build_error, exc)
        else:
            unit_count = len(import_data.units)
            public_count = len(import_data.public_areas)
            self.root.after(
                0, self._on_build_success, result, unit_count, public_count
            )

    def _on_build_success(
        self, result: BuildResult, unit_count: int, public_count: int
    ) -> None:
        self.build_button.configure(state=NORMAL)
        incomplete = result.incomplete_adjustments
        if incomplete:
            self.status.set(
                f"完成（部分持分未調整）：{result.path.name}"
                f"（{unit_count} 戶、{public_count} 筆公設）"
            )
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
            self.status.set(
                f"完成：{result.path.name}（{unit_count} 戶、{public_count} 筆公設）"
            )
            messagebox.showinfo(
                "建立完成",
                "已產生試算表，所有持分均已調整到目標。",
            )

    def _on_permission_error(self, output_file: Path) -> None:
        self.build_button.configure(state=NORMAL)
        self.status.set("無法寫入檔案，請先關閉已開啟的 Excel 檔案再試一次。")
        messagebox.showerror(
            "無法寫入檔案",
            f"無法寫入：\n{output_file}\n\n"
            "該檔案可能正在 Excel 中開啟，請先關閉後再建立一次。",
        )

    def _on_build_error(self, exc: Exception) -> None:
        self.build_button.configure(state=NORMAL)
        self.status.set("建立失敗，請檢查 CSV 內容或設定。")
        messagebox.showerror("建立失敗", str(exc))


def main() -> None:
    root = Tk()
    TemplateBuilderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
