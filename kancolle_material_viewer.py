import time
import datetime

import fire
import pandas
import plotly.graph_objects as graph_objects
from plotly.subplots import make_subplots
import schedule


def read_excel(file : str, sheet : str) -> pandas.DataFrame:
    dataframe = pandas.read_excel(file, sheet_name = sheet)
    return dataframe


def remove_invalid_record(dataframe : pandas.DataFrame) -> pandas.DataFrame:
    target_column_list = ['日付', '燃料', '弾薬', '鉄鋼', 'ボーキ', 'バケツ']
    dataframe = dataframe.rename(columns = {'#日付' : '日付'})
    dataframe = dataframe[target_column_list]
    dataframe = dataframe.fillna('0')
    # 燃料の記述がないレコードは除外
    dataframe = dataframe[dataframe['燃料'] != '0']
    return dataframe


def plot_material_sub(dataframe : pandas.DataFrame, figure : graph_objects.Figure,
                      row : int, col : int,
                      x_title : str, y_title : str, y_secondary_title : str):
    datetime = dataframe['日付']
    fuel = dataframe['燃料']
    bullet = dataframe['弾薬']
    metal = dataframe['鉄鋼']
    bauxite = dataframe['ボーキ']
    bucket = dataframe['バケツ']

    # line引数にはdict関数での定義が必要 + カラーコードでの定義が必要
    # https://classynode.com/2021/08/plotly_color/
    figure.add_trace(graph_objects.Scattergl(x = datetime, y = fuel, name = '燃料', line = dict(color = "#90ee90")), row = row, col = col)
    figure.add_trace(graph_objects.Scattergl(x = datetime, y = bullet, name = '弾薬', line = dict(color = "#ffd700")), row = row, col = col)
    figure.add_trace(graph_objects.Scattergl(x = datetime, y = metal, name = '鉄鋼', line = dict(color = '#a9a9a9')), row = row, col = col)
    figure.add_trace(graph_objects.Scattergl(x = datetime, y = bauxite, name = 'ボーキ', line = dict(color = '#ffa500')), row = row, col = col)
    figure.add_trace(graph_objects.Scattergl(x = datetime, y = bucket, name = 'バケツ', line = dict(color = '#7fffd4')), row = row, col = col, secondary_y = True)

    figure.update_xaxes(title = x_title,
                        rangeselector = dict(
                            buttons = list([
                                dict(count = 1, label = '1ヶ月', step = 'month', stepmode = 'backward'),
                                dict(count = 6, label = '6ヶ月', step = 'month', stepmode = 'backward'),
                                dict(count = 1, label = '今年から', step = 'year', stepmode = 'todate'),
                                dict(count = 1, label = '1年', step = 'year', stepmode = 'backward'),
                                dict(step = 'all', label = 'すべて'),
                            ])
                        ),
                        rangeslider = dict(visible = True), type = 'date', row = row, col = col)
    figure.update_yaxes(title = y_title, showgrid = False,
                        row = row, col = col)
    figure.update_yaxes(title = y_secondary_title, secondary_y = True, showgrid = False,
                        row = row, col = col)


def plot_material(dataframe : pandas.DataFrame) -> graph_objects.Figure:
    figure = make_subplots(rows = 1, cols = 1, specs = [[{'secondary_y' : True}]])

    plot_material_sub(dataframe, figure, 1, 1, '日付', '総合資源量', '総合資源量(バケツ)')

    figure.update_layout(title = '艦これ資源表',
                         font = {'family' : 'Meiryo', 'size' : 20, 'color' : 'darkgray'}, template = 'plotly_dark')

    # 確認用にfigure.showで確認したくなるが、データ量によってはレスポンスを失うため
    # 確認は出力htmlを推奨

    return figure


def main_job(excel : str, sheet : str, html : str):
    print('------ 定期実行開始 : {}'.format(datetime.datetime.now()))

    print('------ excel読み込み')
    dataframe = read_excel(excel, sheet)
    print(dataframe)

    print('------ 無効レコードを除外')
    dataframe = remove_invalid_record(dataframe)
    print(dataframe)

    print('------ 資源の可視化')
    figure = plot_material(dataframe)

    print('------ 表の出力')
    figure.write_html(html)


def main(excel : str,
         sheet : str = '資源メモ',
         html : str = 'kankolle_material_viewer.html',
         frequency_hour : float = 0):
    """
    艦これ資源ビューア作成スクリプト

    Args:
        excel: エクセルファイルパス
        sheet : 処理対象のシート名
        html : 出力htmlファイルパス
        frequency_hour : 実行周期　時間単位で指定、0指定で定期実行なし、1時間未満を指定する場合は小数点指定
    """
    print('[{}]'.format(__file__))
    print('excel filepath : {}'.format(excel))
    print('excel target sheet : {}'.format(sheet))
    print('output html filepath : {}'.format(html))
    print('frequency execution {}[hours]'.format(frequency_hour))

    # まずは1回
    main_job(excel, sheet, html)
    # frequency_hourが0の場合はここで終わり
    if frequency_hour == 0:
        print('------ 定期実行なしで終了')
        return

    # 定期実行
    schedule.every(frequency_hour).hours.do(main_job, excel = excel, sheet = sheet, html = html)
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    fire.Fire(main)
