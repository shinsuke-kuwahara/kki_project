import logging
import requests
import os
import json
import pandas as pd
import pathlib
from datetime import timezone, timedelta, datetime

from kki_func.slack import slack_robo_post
from kki_func.loginfo import loginfo

# 製造部用スリット予定保存先
SEIZOU_DIR = pathlib.WindowsPath(r"\\10.100.108.150\disk1\販売管理データ\製造部用\スリット予定")
# 生産管理用スリット予定保存先
SEISAN_DIR = pathlib.WindowsPath(r"\\10.100.108.150\生産管理\在庫管理表\各種資料\ワイチ納品書")
# ラミ予定用保存先
RAMI_DIR = pathlib.WindowsPath(r"\\10.100.108.150\disk1\販売管理データ\製造部用\ラミ予定")
# 印刷予定表用保存先
PRINT_DIR = pathlib.WindowsPath(r'\\10.100.108.150\disk1\販売管理データ\製造部用\印刷予定表')
# 保存先をリストに格納
SAVE_DIR = [SEIZOU_DIR]
# logファイル保存先
LOG_FILE = r"/logfile.txt"
# ワークスペースK＆Nのボットトークン
TOKEN = 'xoxb-994892783410-3441445086132-WyKdG5wv8l9GsXUsJ7YCfmqe'
# ファイル取得用URL
FILE_URL = 'https://slack.com/api/conversations.history'
# トークン認証用
HEADERS = {"Authorization": "Bearer " + TOKEN}


def slit_download():
    # チャンネルkkiスリット予定表のID
    channel = 'C01126DMRK9'
    # パラメータを設定
    params = {
        'channel': channel,
        'limit': 15
    }
    # リクエストして情報を取得
    res = requests.get(url=FILE_URL, headers=HEADERS, params=params)
    data = res.json()

    data_list = []
    # 1週間前の日付取得用－15日
    before = timedelta(days=30)
    # 1週間前日付取得 -10
    before_date = datetime.strftime(datetime.now() - before, "%Y-%m-%d")
    try:
        # 必要なデータのみ抽出する
        for message in data['messages']:
            if 'files' in message:
                for file in message['files']:
                    if 'timestamp' in file:
                        timestamp = file['timestamp']
                        if timestamp != 0:
                            timestamp = str(datetime.fromtimestamp(timestamp)).replace(":", "_")
                            if timestamp >= before_date and file['filetype'] == 'xlsm':
                                data_list.append([timestamp, file['title'], file['url_private_download']])

        print(data_list)
        # 重複ファイルがないかチェックしなければ保存する
        for file in data_list:
            filename = f'{file[1][:7]} {file[0]}.xlsm'
            for hozon in SAVE_DIR:
                # 保存先に重複ファイルがあるか確認
                if os.path.exists(f"{hozon}\\{filename}"):
                    continue
                elif os.path.exists(f"{hozon}\\処理済みフォルダ\\{filename}"):
                    continue

                url = file[2]
                content = requests.get(
                    url, headers=HEADERS
                ).content
                with open(f"{hozon}\\{filename}", mode="wb") as f:
                    f.write(content)
    except Exception:
        # logging.basicConfig(level=logging.INFO, filename=log_file)
        logging.exception("エラーログ")

    # slack_post('[印刷日報ダウンロード完了]', 2)


def print_download():
    # channel印刷予定表
    channel = "C010VUK5VB6"
    # パラメータを設定
    params = {
        'channel': channel
    }

    # リクエストして情報を取得
    res = requests.get(url=FILE_URL, headers=HEADERS, params=params)
    data = res.json()

    data_list = data_input(data)

    # 重複ファイルがないかチェックしなければ保存する
    for file in data_list:
        if "スラック用" not in file[1]:
            filename = f'{file[1].replace(".xlsm", " ")}{file[0]}.xlsm'
        else:
            continue

        # 保存先に重複ファイルがあるか確認
        if os.path.exists(f"{PRINT_DIR}\\{filename}"):
            continue
        elif os.path.exists(f"{PRINT_DIR}\\処理済みフォルダ\\{filename}"):
            continue

        url = file[2]
        content = requests.get(
            url, headers=HEADERS
        ).content
        with open(f"{PRINT_DIR}\\{filename}", mode="wb") as f:
            f.write(content)


def rami_download():
    # channelラミ納品のチャンネルID
    channel = 'C010YDXUCTG'

    # パラメータを設定
    params = {
        'channel': channel,
        'limit': 15
    }
    # リクエストして情報を取得
    res = requests.get(url=FILE_URL, headers=HEADERS, params=params)
    data = res.json()

    data_list = data_input(data)

    # 重複ファイルがないかチェックしなければ保存する
    for file in data_list:
        filename = f'{file[1].replace(".xlsm", " ")}{file[0]}.xlsm'
        if "敷島" in filename:
            filename = filename.replace(".xls", " ")
            filename = filename.replace("m", ".xls")

        # 保存先に重複ファイルがあるか確認
        if os.path.exists(f"{RAMI_DIR}\\{filename}"):
            continue
        elif os.path.exists(f"{RAMI_DIR}\\処理済みフォルダ\\{filename}"):
            continue

        url = file[2]
        content = requests.get(
            url, headers=HEADERS
        ).content
        with open(f"{RAMI_DIR}\\{filename}", mode="wb") as f:
            f.write(content)


def data_input(data):
    data_list = []
    # 1週間前の日付取得用－15日
    before = timedelta(days=15)
    # 1週間前日付取得 -10
    before_date = datetime.strftime(datetime.now() - before, "%Y-%m-%d")
    try:
        # 必要なデータのみ抽出する
        for message in data['messages']:
            if 'files' not in message:
                continue
            for file in message['files']:
                if 'timestamp' in file:
                    timestamp = file['timestamp']
                    if timestamp != 0:
                        timestamp = str(datetime.fromtimestamp(timestamp)).replace(":", "_")
                        filetype = file['filetype']
                        if timestamp >= before_date and filetype == 'xlsm' or filetype == 'xls':
                            data_list.append([timestamp, file['title'], file['url_private_download']])
        return data_list
    except Exception:
        # logging.basicConfig(level=logging.INFO, filename=log_file)
        logging.exception("エラーログ")


def main():
    slit_download()
    rami_download()
    print_download()


if __name__ == "__main__":
    main()