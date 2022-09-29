###### tags: `OCR_DEMO`
# 5-2sort_main
# 時間を計測するモジュール
import time
# フォルダを開く為のモジュール（転記してExcelを開いたあと原本と比較して確認できるように）
import subprocess
# Slackへメッセージを飛ばすモジュール
import requests
# ログの吐きだしに使用するモジュール
import logging
# ファイルやフォルダパスの取得に使用するモジュール
import os
# 年数を直打ちでなく変数で取得するためのモジュール
import datetime
# メモリを解放するモジュール
import gc
# セレニウムで読み取り結果を落としてソート・リネームさせるプログラム
from demo_sort_v2 import Sort
# Slackへ通知を行うモジュール
# from　slack_post import demo_slack_post
# フォルダやファイル名に現在日時を使用するので、定義しておく
# 現在時刻を取得
now = datetime.datetime.now()
current_time = now.strftime('%Y-%m-%d-%H-%M-%S')

# ダウンロード結果を格納するフォルダを新規作成
# ここに後々、仕分け結果のトレイ毎のフォルダが入ります。
new_folder_path = './getdata/' + current_time
os.makedirs(new_folder_path, exist_ok = True)

""" ↓logの設定 """
# ベースのログ設定
# INFO以上のレベルの際にlogが書かれる
"""
関数名	    用法
debug( )	問題がないか診断する場合
info( )	    思い通りに動作する場合
warning( )	想定外の事が発生しそうな場合
error( )	重大なエラーにより、一部の機能を実行できない場合
critical( )	プログラム自体が実行を続けられないことを表す、致命的な不具合の場合
"""
logfile = './getlogdate.txt'
formatter = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(level=logging.INFO, format=formatter, filename=logfile)
""" ↑logの設定 """

def sort_main():
    # 処理時間計測の為に処理の始めにstartを定義
    start = time.time()

    # 起動したことを通知
    # コンソール非表示だと動いているかどうか分からない為動き出したことを通知
    # slack_post('<!channel>\nファイルソートのプログラム起動しました')
    print("ファイルソートのプログラム起動しました")
    # Sort()内で取得した読み取り枚数の変数をプリントに入れる
    call_2 = Sort(current_time, new_folder_path)
    # print(call_2)
    # 処理時間計測の為に処理の終わりにend_3を定義
    end_3 = time.time()
    # end_3からstartを引いて処理にかかった時間を求める
    # print('%.3f seconds' % (end_3 - start))
    # メモリ解放
    del end_3
    gc.collect()

    # 仕分け完了のメッセージを追加する
    # slack_post("FAXの仕分けが完了しました")
    print("FAXの仕分けが完了しました")


    # 任意のフォルダを開く(EXCEL転記内容及び仕分け結果確認用）
    subprocess.run('explorer {}'.format(os.getcwd() + "\getdata"))

    # 処理時間計測の為に処理の終わりにend_4を定義
    end_4 = time.time()
    # end_4からstartを引いて処理にかかった時間を求める
    # print('%.3f seconds' % (end_4 - start))
    # メモリ解放
    del end_4
    gc.collect()

if __name__ == '__main__':
    sort_main()
