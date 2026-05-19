"""
複数のExcelファイルを1つに集約するスクリプト

使い方:
    python aggregate_excel.py

設定:
    INPUT_DIR  : 作業Excelが入っているフォルダ
    OUTPUT_FILE: 集約後に出力するファイル名
    SHEET_NAME : 各Excelの何番目(または名前)のシートを読むか
"""

import glob
import os
from pathlib import Path

import pandas as pd

# ── 設定 ──────────────────────────────────────────────
INPUT_DIR   = "./input"        # 作業Excelが入っているフォルダ
OUTPUT_FILE = "./output/集約データ.xlsx"  # 出力先
SHEET_NAME  = 0               # 0=先頭シート。シート名で指定する場合は "Sheet1" など
SKIP_ROWS   = 0               # 読み飛ばす行数（ヘッダー前に余分な行がある場合）
# ──────────────────────────────────────────────────────


def read_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_NAME, skiprows=SKIP_ROWS)
    # ファイル名を「出所」列として追加（どのファイル由来か分かるように）
    df.insert(0, "出所ファイル", Path(path).name)
    return df


def main():
    files = glob.glob(os.path.join(INPUT_DIR, "*.xlsx"))
    if not files:
        print(f"[警告] {INPUT_DIR} にExcelファイルが見つかりません。")
        return

    print(f"{len(files)} 件のファイルを読み込みます...")

    frames = []
    for f in sorted(files):
        try:
            df = read_excel(f)
            frames.append(df)
            print(f"  ✓ {Path(f).name}  ({len(df)} 行)")
        except Exception as e:
            print(f"  ✗ {Path(f).name}  スキップ: {e}")

    if not frames:
        print("読み込めたファイルがありません。処理を終了します。")
        return

    result = pd.concat(frames, ignore_index=True)
    print(f"\n合計 {len(result)} 行を集約しました。")

    # 出力フォルダがなければ作成
    os.makedirs(os.path.dirname(OUTPUT_FILE) or ".", exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="集約データ")

        # 列幅を自動調整
        ws = writer.sheets["集約データ"]
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    print(f"出力完了: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
