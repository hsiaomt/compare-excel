import os
import pandas as pd

file_dir = os.path.dirname(os.path.abspath(__file__))
result_file_name = 'compare_result.xlsx'
result_file_path = os.path.join(file_dir, result_file_name)
if os.path.exists(result_file_path):
    os.remove(result_file_path)
    print(f"{result_file_name} 已刪除")
else:
    print(f"找不到 {result_file_name}")
extensions = ('.xls', '.xlsx')
files = [f for f in os.listdir(file_dir) if f.endswith(extensions)]
files.sort()
file1_path = os.path.join(file_dir, files[0])
file2_path = os.path.join(file_dir, files[1])

xls1 = pd.read_excel(file1_path, sheet_name=None, dtype=str)
xls2 = pd.read_excel(file2_path, sheet_name=None, dtype=str)

with pd.ExcelWriter(result_file_path, engine="openpyxl") as writer:

    # 取得所有工作表名稱
    all_sheets = set(xls1.keys()) | set(xls2.keys())

    for sheet in all_sheets:

        df1 = xls1.get(sheet, pd.DataFrame()).fillna("")
        df2 = xls2.get(sheet, pd.DataFrame()).fillna("")

        # 統一欄位
        all_cols = sorted(set(df1.columns) | set(df2.columns))

        df1 = df1.reindex(columns=all_cols, fill_value="")
        df2 = df2.reindex(columns=all_cols, fill_value="")

        # 增加來源標記
        df1["_Source"] = "File1"
        df2["_Source"] = "File2"

        # 找出不同列
        merged = pd.concat([df1, df2])

        diff = merged.drop_duplicates(
            subset=all_cols,
            keep=False
        )

        if not diff.empty:
            diff.to_excel(
                writer,
                sheet_name=sheet[:31],
                index=False
            )

print(f"完成，差異已輸出至：{result_file_path}")
