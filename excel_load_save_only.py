import openpyxl as op
from openpyxl.chart import LineChart, Reference, Series



if __name__ == "__main__":

    #グラフの作成
    #https://atmarkit.itmedia.co.jp/ait/articles/2203/08/news028.html
    #https://www.shibutan-bloomers.com/python_libraly_openpyxl-9/3126/

    wb = op.load_workbook("/home/ubuntu/conda_src/excel/data/test_ver0.xlsx", data_only=False)# ワークシートの読み込み
    ws = wb['Sheet1'] # ワークシートの有効化






    wb.save("/home/ubuntu/conda_src/excel/data/load_save_only.xlsx")


