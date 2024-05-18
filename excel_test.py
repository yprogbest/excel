import openpyxl as op
from openpyxl.chart import LineChart, Reference, Series



if __name__ == "__main__":

    #グラフの作成
    #https://atmarkit.itmedia.co.jp/ait/articles/2203/08/news028.html
    #https://www.shibutan-bloomers.com/python_libraly_openpyxl-9/3126/

    wb = op.load_workbook("/home/ubuntu/conda_src/excel/data/test_ver0.xlsx")# ワークシートの読み込み
    ws = wb['Sheet1'] # ワークシートの有効化

    # rmin = ws.min_row
    # rmax = ws.max_row
    # cmin = ws.min_column
    # cmax = ws.max_column


    #データ範囲の設定
    min_col = 1
    max_col = 6
    min_row = 1
    max_row = 4



    chart = LineChart()

    cats = Reference(ws, min_col=min_col+1, min_row=min_row, max_col=max_col, max_row=min_row)

    # 各行のデータを系列として追加
    for i in range(2, max_row+1):
        data = Reference(ws, min_col=min_col+1, min_row=i, max_col=max_col, max_row=i)
        series = Series(data, title=ws.cell(row=i, column=1).value)
        chart.append(series)
        
    
    chart.set_categories(cats)
    chart.anchor = 'I1'  # グラフの表示位置
    chart.width = 16  # グラフのサイズ
    chart.height = 8

    # 各シリーズに自動マーカーを設定
    for s in chart.series:
        s.marker.symbol = "auto"
        s.smooth = False

    ws.add_chart(chart)
    wb.save("/home/ubuntu/conda_src/excel/data/sample_chart.xlsx")


