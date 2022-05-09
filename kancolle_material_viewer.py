from pprint import pprint

import fire
import pandas


def read_excel(file : str, sheet : str) -> pandas.DataFrame:
    dataframe = pandas.read_excel(file, sheet_name = sheet)
    return dataframe


def remove_invalid_record(dataframe : pandas.DataFrame) -> pandas.DataFrame:
    target_column_list = ['日付', '燃料', '弾薬', '鉄鋼', 'ボーキ', 'バケツ']
    dataframe = dataframe.rename(columns = {'#日付' : '日付'})
    dataframe = dataframe[target_column_list]
    dataframe = dataframe.fillna('')
    # 燃料の記述がないレコードは除外
    dataframe = dataframe[dataframe['燃料'] != '']
    return dataframe


def main(excel : str,
         sheet : str = '資源メモ'):
    """
    艦これ資源ビューア作成スクリプト

    Args:
        excel: エクセルファイルパス
        sheet : 処理対象のシート名
    """
    print('[{}]'.format(__file__))
    print('excel filepath : {}'.format(excel))

    print('------ excel読み込み')
    dataframe = read_excel(excel, sheet)
    print(dataframe)

    print('------ 無効レコードを除外')
    dataframe = remove_invalid_record(dataframe)
    print(dataframe)


if __name__ == '__main__':
    fire.Fire(main)
