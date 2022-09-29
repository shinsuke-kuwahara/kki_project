import datetime
import os
import pandas as pd
import pathlib


# プログラムの実行時間を測定する関数
def log_effect(filename, starttime, cnt):
    """
    :rtype: object
    :param filename:ログ保存用パス 記載例⇒'\フォルダ名\〇〇〇.csv'
    :param starttime:プログラムのスタートタイム 記載例⇒　変数名 = datetime.datetime.now()
    :param cnt: 処理の件数
    :return:
    """
    # 保存先のパス
    path = pathlib.WindowsPath(r"\\10.100.108.150\disk1\テック事業部共有\効果測定用")
    #　ファイルパス
    logfile = f'{path}{filename}'
    # 現在の日付を取得
    date = datetime.datetime.strftime(datetime.datetime.now(), "%Y/%m/%d")
    # ファイルが存在するか確認。無ければ新規作成
    is_file = os.path.isfile(logfile)
    if not is_file:
        df = pd.DataFrame(index=[], columns=[
            '日付', '開始時間', '終了時間', '処理時間', '処理件数'
        ])
    else:
        df = pd.read_csv(logfile, encoding='cp932')
    # 終了時間
    endtime = datetime.datetime.now()
    # 処理時間を算出
    ans = endtime - starttime
    # データフレーム転記用
    count = len(df)
    # logに情報を格納
    log = [date, str(starttime)[11:19], str(endtime)[11:19], str(ans)[:7], cnt]
    # ログファイルにデータを追加
    df.loc[count] = log
    # ログファイルを保存
    df.to_csv(logfile, index=False, encoding='cp932')