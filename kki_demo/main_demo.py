###### tags: `OCR_DEMO`
# 1-2メインプログラム：駆動用
""" ↓モジュール """
import time
import subprocess
# import slackweb

import os
import datetime

# DXSLへリクエストする為のプログラム
import logging
from demo_request import D_request
from download_sort import demo_transfer
# Slackへ通知を行うモジュール
# from　slack_post import demo_slack_pos

""" ↑モジュール """

""" ↓logの設定 """
"""
関数名	    用法
debug( )	問題がないか診断する場合
info( )	    思い通りに動作する場合
warning( )	想定外の事が発生しそうな場合
error( )	重大なエラーにより、一部の機能を実行できない場合
critical( )	プログラム自体が実行を続けられないことを表す、致命的な不具合の場合
"""
LOGFILE = r'.\logfile\log.txt'
formatter = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(level=logging.INFO, format=formatter, filename=LOGFILE)
# """ ↑logの設定 """


# プログラムを動作させる関数
def main_request():
    logging.info('リクエスト完了')
    # 起動したことを通知
    # slack_post('<!channel>\nDEMO_１.リクエストのプログラムを起動しました')

    mes, sortid, headers = D_request()
    demo_transfer(sortid, headers)
    # webhook_mesというリスト型変数に格納されているメッセージを最後にまとめて通知する
    # slack_post(str(mes))
    # logを吐きだす
    print(mes)
    logging.info('プログラム終了')


if __name__ == '__main__':
    main_request()
