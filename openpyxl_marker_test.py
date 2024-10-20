import openpyxl
import zipfile
import xml.etree.ElementTree as ET


if __name__ == '__main__':
    input_file = "/home/ubuntu/conda_src/excel/!openpyxlでグラフの作成/data/openpyxl_two_type_graph_test.xlsx"
    outputfile = "/home/ubuntu/conda_src/excel/!openpyxlでグラフの作成/data/openpyxl_two_type_graph_test_output.xlsx"
    wb = openpyxl.load_workbook(input_file)

    for each_ws in wb.worksheets:
        charts = each_ws._charts
        # chartの数を数える
        chart_cnt = len(charts)
    #     for chart in charts:
    #         for series in chart.series:
    #             series.marker.symbol = "auto"
    #             series.smooth = False

    print(chart_cnt)

    # エクセルファイルを保存する
    wb.save(outputfile)

    # グラフのカスタマイズを行うためにZIPとして開く
    with zipfile.ZipFile(input_file, 'r') as z:
        print(z.infolist())
        # 'chart2.xml'の読み込み
        chart_xml = z.read('xl/charts/chart2.xml')
    
    # # XMLを解析
    # root = ET.fromstring(chart_xml)


    # for ser in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}ser'):
    #     marker = ET.SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}marker')
    #     symbol = ET.SubElement(marker, '{http://schemas.openxmlformats.org/drawingml/2006/chart}symbol')
    #     symbol.set('val', 'auto')  # マーカーの形状を設定
    #     size = ET.SubElement(marker, '{http://schemas.openxmlformats.org/drawingml/2006/chart}size')
    #     size.set('val', '5')  # マーカーのサイズを設定

    # # XMLを文字列に変換する
    # new_chart_xml = ET.tostring(root, method='xml')

    # # XMLを保存する
    # with zipfile.ZipFile(input_file, 'r') as z_in:
    #     with zipfile.ZipFile(outputfile, 'w') as z_out:
    #         for item in z_in.infolist():
    #             # chart1.xmlを修正したものに置き換え
    #             if item.filename == 'xl/charts/chart2.xml':
    #                 z_out.writestr(item, new_chart_xml)  # 修正後のXMLを埋め込み
    #             else:
    #                 z_out.writestr(item, z_in.read(item.filename))  # 他のファイルはそのままコピー
