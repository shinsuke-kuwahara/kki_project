import logging
import requests
import os
import json
import pandas as pd
import pathlib
from datetime import timezone, timedelta, datetime

from kki_function.slack import slack_robo_post
from kki_function.loginfo import loginfo

# 製造部用スリット予定保存先
SEIZOU_DIR = pathlib.WindowsPath(r"\\10.100.108.150\disk1\販売管理データ\製造部用\スリット予定")
# 生産管理用スリット予定保存先
SEISAN_DIR = pathlib.WindowsPath(r"\\10.100.108.150\生産管理\在庫管理表\各種資料\ワイチ納品書")
# 保存先をリストに格納
SAVE_DIR = [SEIZOU_DIR]
# logファイル保存先
LOG_FILE = r"/logfile.txt"
# ワークスペースK＆Nのボットトークン
TOKEN = 'xoxb-994892783410-3441445086132-WyKdG5wv8l9GsXUsJ7YCfmqe'
# ファイル取得用URL
URL = 'https://slack.com/api/files.list'


def main():
    # トークン認証用
    headers = {"Authorization": "Bearer " + TOKEN}
    # チャンネルkkiスリット予定表のID
    channel = 'C01126DMRK9'
    # パラメータを設定
    params = {
        'channel': channel
    }
    # リクエストして情報を取得
    res = requests.get(url=URL, headers=headers, params=params)
    data = res.json()

    data_list = []
    # 1週間前の日付取得用－15日
    before = timedelta(days=15)
    # 1週間前日付取得 -10
    before_date = datetime.strftime(datetime.now() - before, "%Y-%m-%d")
    try:
        # 必要なデータのみ抽出する
        for i in reversed(data['files']):
            i["timestamp"] = str(datetime.fromtimestamp(i["timestamp"])).replace(":", "_")

            if i["timestamp"][:10] <= before_date:
                break

            # 拡張子が".xlsm"以外は処理を飛ばす
            if os.path.splitext(i['title'])[1] != ".xlsm":
                continue

            data_list.append([i['timestamp'], i["title"], i['url_private_download']])

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
                    url, headers=headers
                ).content
                with open(f"{hozon}\\{filename}", mode="wb") as f:
                    f.write(content)
    except Exception:
        # logging.basicConfig(level=logging.INFO, filename=log_file)
        logging.exception("エラーログ")

    # slack_post('[印刷日報ダウンロード完了]', 2)


if __name__ == "__main__":
    main()