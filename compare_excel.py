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

# 1. 讀取兩個 Excel 檔案
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
'''
# 2. 將整行轉換為字串來比對差異（合併整行所有欄位）
# 若欄位名稱不同，請替換成具體的欄位名（例如 df1['ID']）
# 若只要比對整row，可使用 pd.merge(..., how='outer', indicator=True)
merged = pd.merge(df1, df2, how='outer', indicator=True)

# 3. 找出完全一樣的 row
same_rows = merged[merged['_merge'] == 'both'].drop(columns=['_merge'])

# 4. 找出不同的 row
diff_rows = merged[merged['_merge'] != 'both']

# 5. 將比對結果匯出到新的 Excel
with pd.ExcelWriter('compare_result.xlsx') as writer:
    same_rows.to_excel(writer, sheet_name='完全相同', index=False)
    diff_rows.to_excel(writer, sheet_name='差異資料', index=False)

print("比對完成！結果已儲存至 compare_result.xlsx")
'''