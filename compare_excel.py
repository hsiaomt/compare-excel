import os
import pandas as pd

file_dir = os.path.dirname(os.path.abspath(__file__))
result_file_path = os.path.join(file_dir, 'compare_result.xlsx')
if os.path.exists(result_file_path):
    os.remove(result_file_path)
    print("compare_result.xlsx 已刪除")
else:
    print("找不到 compare_result.xlsx")
extensions = ('.xls', '.xlsx')
files = [f for f in os.listdir(file_dir) if f.endswith(extensions)]
files.sort()
file1_path = os.path.join(file_dir, files[0])
file2_path = os.path.join(file_dir, files[1])

# 1. 讀取兩個 Excel 檔案
df1 = pd.read_excel(file1_path, '工作表1')
df2 = pd.read_excel(file2_path, '工作表1')

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