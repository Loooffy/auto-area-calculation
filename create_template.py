from __future__ import annotations

import argparse
from pathlib import Path

from csv_import import parse_import_csv
from template_builder import create_template

DEFAULT_OUTPUT = "試算面積-混公.xlsx"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="自動建立面積持分試算表（含標題、合併儲存格與算式）。"
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(DEFAULT_OUTPUT),
        help=f"輸出路徑（預設：{DEFAULT_OUTPUT}）",
    )
    parser.add_argument(
        "--units",
        type=int,
        default=None,
        help="各戶資料列數（未指定 --csv 時預設 10）",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=None,
        help="試算表 CSV（各戶面積 + 公設面積）；預設不匯入",
    )
    parser.add_argument(
        "--public-ratio",
        type=float,
        default=None,
        help="公設比（百分比數值，例如 32 代表 32%%）；寫入公設比欄位 C73",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    unit_data = None
    public_areas = None
    if args.csv is not None:
        csv_path = args.csv.resolve()
        if not csv_path.exists():
            raise SystemExit(f"找不到 CSV：{csv_path}")
        import_data = parse_import_csv(csv_path)
        unit_data = import_data.units
        public_areas = import_data.public_areas

    if unit_data is None:
        units = args.units if args.units is not None else 10
        if units < 1:
            raise SystemExit("--units 必須 >= 1")
    elif args.units is not None and args.units != len(unit_data):
        print(
            f"注意：--units {args.units} 已忽略，改以 CSV 的 {len(unit_data)} 戶為準"
        )

    output_path = args.output.resolve()
    try:
        result = create_template(
            output_path,
            units=args.units if args.units is not None else 10,
            unit_data=unit_data,
            public_areas=public_areas,
            public_ratio=args.public_ratio,
            include_porch=import_data.include_porch if args.csv is not None else False,
        )
    except PermissionError:
        raise SystemExit(
            f"無法寫入檔案：{output_path}\n"
            "該檔案可能正在 Excel 中開啟，請先關閉後再執行一次。"
        )

    output = result.path
    if unit_data is not None:
        public_count = len(public_areas) if public_areas else 0
        porch = "，含門廊" if import_data.include_porch else ""
        print(
            f"OK: {output} ({len(unit_data)} units, {public_count} public areas{porch})"
        )
    else:
        print(f"OK: {output}")

    for item in result.incomplete_adjustments:
        print(
            f"注意：{item.name} 未能調整到目標（目標 {item.target}、"
            f"實際 {item.actual}、差 {item.actual - item.target}），請人工處理。"
        )


if __name__ == "__main__":
    main()
