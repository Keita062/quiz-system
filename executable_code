pip install openpyxl

import random
import openpyxl

# 新しいシート「ControlSheet」を作成する関数
def setup_button():
    # Excelファイルを読み込む
    wb = openpyxl.load_workbook("/content/極める年表.xlsx")
    
    # 'ControlSheet'が存在するか確認
    if "ControlSheet" not in wb.sheetnames:
        ws = wb.create_sheet("ControlSheet")
    else:
        ws = wb["ControlSheet"]
    
    # ボタンに関してはopenpyxlでは追加できませんが、関数として処理を実行する
    print("ControlSheetが作成されました。'add_random_asterisks'関数を実行してください。")
    
    # 変更を保存
    wb.save("your_workbook.xlsx")

# ランダムなアスタリスクを上書きする関数
def add_random_asterisks():
    # Excelファイルを読み込む
    wb = openpyxl.load_workbook("/content/極める年表.xlsx")
    
    # 対象シート名と列番号、上書きするアスタリスクの数を設定
    target_sheets = ["縄文から鎌倉.問題", "室町から江戸.問題", "明治から平成.問題"]
    target_columns = [1, 2, 3, 4, 5, 6]  # 対応する列 A, B, C, D, E, F (0-indexed は A=1, B=2, C=3 ...)
    blank_counts = [3, 7, 6, 4, 5, 5]  # 各列に挿入するアスタリスクの数
    
    # 対象シートごとに処理を行う
    for sheet_name in target_sheets:
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # 各列に対してランダムにアスタリスクを上書き
            for col_index, blank_count in zip(target_columns, blank_counts):
                # 現在の列の最終行を取得
                last_row = ws.max_row
                
                # 空でないセルをランダムに選んでアスタリスクを上書き
                filled_cells = []
                for row in range(2, last_row + 1):  # ヘッダー行を避けるため、2行目からスタート
                    cell = ws.cell(row=row, column=col_index)
                    if cell.value is not None:  # 空でないセルのみ対象
                        filled_cells.append(row)
                
                # ランダムに空白セルにアスタリスクを上書き
                for _ in range(blank_count):
                    if filled_cells:
                        random_row = random.choice(filled_cells)  # ランダムにセルを選ぶ
                        ws.cell(row=random_row, column=col_index).value = "*"  # 上書き
                        filled_cells.remove(random_row)  # 挿入したセルをリストから削除
    
    # 変更を保存
    wb.save("your_workbookD.xlsx")
    print("指定されたシートにランダムなアスタリスクを上書きしました。")

# 実行例
setup_button()  # 'ControlSheet'を作成
add_random_asterisks()  # ランダムにアスタリスクを上書き
