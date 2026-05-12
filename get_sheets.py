import openpyxl
import sys

# Windows側のファイルパスを引数から受け取る、または直接書き換える
file_path = '/mnt/c/Users/K157227/Desktop/004.作業手順書/利用状況報告.xlsx'

try:
    # 読み取り専用モードで開く（ファイルが開いている際のエラー回避）
    wb = openpyxl.load_workbook(file_path, read_only=True)
    for name in wb.sheetnames:
        print(name)
except FileNotFoundError:
    print(f"Error: ファイルが見つかりません。パスを確認してください: {file_path}")
except Exception as e:
    print(f"Error: {e}")
