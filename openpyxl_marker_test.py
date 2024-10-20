import openpyxl
import zipfile
import xml.etree.ElementTree as ET


if __name__ == '__main__':
    input_file = "/home/ubuntu/conda_src/excel/!openpyxlでグラフの作成/data/openpyxl_two_type_graph_test.xlsx"
    outputfile = "/home/ubuntu/conda_src/excel/!openpyxlでグラフの作成/data/openpyxl_two_type_graph_test_output.xlsx"
    wb = openpyxl.load_workbook(input_file)

    # for each_ws in wb.worksheets:
    #     charts = each_ws._charts
    #     for chart in charts:
    #         for series in chart.series:
    #             series.marker.symbol = "auto"
    #             series.smooth = False

    # エクセルファイルを保存する
    wb.save(outputfile)

    item_file_name_list = [] # chart_xmlを格納するためのリスト
    new_chart_xml_list = [] # 新たなchart_xmlを格納するためのリスト
    with zipfile.ZipFile(input_file, 'r') as z: # グラフのカスタマイズを行うためにZIPとして開く
        for item in z.infolist():
            targetFileName = 'xl/charts/chart' # chart用xmlファイルを取得するための変数
            if targetFileName in item.filename:
                item_file_name_list.append(item.filename)
                chart_xml = z.read(item.filename) # 'chart○.xml'の読み込み
                
                root = ET.fromstring(chart_xml) # XMLを解析する

                for ser in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/chart}ser'):
                    marker = ET.SubElement(ser, '{http://schemas.openxmlformats.org/drawingml/2006/chart}marker')
                    symbol = ET.SubElement(marker, '{http://schemas.openxmlformats.org/drawingml/2006/chart}symbol')
                    symbol.set('val', 'auto')  # マーカーの形状を設定
                    size = ET.SubElement(marker, '{http://schemas.openxmlformats.org/drawingml/2006/chart}size')
                    size.set('val', '5')  # マーカーのサイズを設定

                # XMLを文字列に変換する
                new_chart_xml = ET.tostring(root, method='xml')
                new_chart_xml_list.append(new_chart_xml)

    # XMLを保存する
    with zipfile.ZipFile(input_file, 'r') as z_in:
        with zipfile.ZipFile(outputfile, 'w') as z_out:
            for item in z_in.infolist(): # z_inのXMLの内容をz_outに書き込む
                z_out.writestr(item, z_in.read(item.filename))  # 他のファイルはそのままコピー
            
            for listNum, item_file_name in enumerate(item_file_name_list): # 変更内容をz_outに書き込む
                for item in z_in.infolist():
                    # chart1.xmlを修正したものに置き換え
                    if item.filename == item_file_name:
                        z_out.writestr(item, new_chart_xml_list[listNum])  # 修正後のXMLを埋め込み

