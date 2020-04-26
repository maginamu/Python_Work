import openpyxl
import os
import sys

fileno = 1
hanni = "B54:B59"

testpass = input("集計対象のファイルがあるパスを入力してください。:")
print(testpass)

testhanni = input("集計対象のセルを入力（B54:B59）:")
print(testhanni)

files = os.listdir(testpass+'\\')
print(files)

hantei = input("集計開始する場合は、1を入力:")
if hantei != '1':
    print('実行を中止します')
    sys.exit(1)


wb_result = openpyxl.Workbook()
sheet_result = wb_result.active
sheet_result.title = '集計結果'

for fcnt in files:
    wb = openpyxl.load_workbook(testpass+'\\'+fcnt)
    sheet_result.cell(fileno,1,fcnt)
    ws = wb["Sheet1"]
    rng1 = ws[testhanni]
    for cnt in range(6):
        sheet_result.cell(fileno, cnt+2, rng1[cnt][0].value)
    fileno+=1

wb_result.save('C:\\Users\\magin\\Desktop\\Tasks\\200410_アンケート収集ツール作成\\集計結果.xlsx')