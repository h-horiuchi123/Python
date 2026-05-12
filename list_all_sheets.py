import openpyxl
import glob
import os

# 走査対象のWindows側ディレクトリパスを指定
target_dir = '/mnt/c/Users/K157227/Desktop/004.作業手順書/フロー'

# 指定ディレクトリ内のxlsxファイルを再帰的に取得
# サブディレクトリも含める場合は recursive=True を使用
files = glob.glob(os.path.join(target_dir, "*.xlsx"))

if not files:
    print("指定されたディレクトリにxlsxファイルが見つかりません。")
else:
    for file_path in files:
        print(f"--- File: {os.path.basename(file_path)} ---")
        try:
            # 読み取り専用でシート名を取得
            wb = openpyxl.load_workbook(file_path, read_only=True, keep_links=False)
            for name in wb.sheetnames:
                print(f"  - {name}")
        except Exception as e:
            print(f"  [Error] ファイルを読み込めませんでした: {e}")
        print("") # 空行で区切りを明確化
