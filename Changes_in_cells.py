#python ver.3.8.0
#pip ver 20.3.3
#注意：作業フォルダにtest.xlsxを作ること

#セルの変更

import openpyxl #openpyxlをインポートする
wb = openpyxl.load_workbook("test.xlsx") #作業フォルダからエクセルファイルを指定
wa = wb.active #アクティブのシートを指定
wa.cell(row = 1 ,column = 2).value = "こんばんは" #セルのrow:行 column:列を指定する
wb.save("test.xlsx") #セーブする
wb.close() #ファイル閉じる
