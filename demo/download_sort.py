from datetime import datetime, time
import time
import io
import json
import pathlib
import pandas as pd
import requests

# 保存先
PATH = pathlib.WindowsPath(r'\\10.100.108.150\disk1\テック事業部共有\demo')
print(f"{PATH}PATH")

# 仕分け結果をダウンロードしてファイルに保存する
def demo_transfer(sortid, headers):
    uri = 'https://kki-1105.dx-suite.com/Sorter/api/v1/csvDownload'

    data = {
        'sortingUnitId': sortid
    }

    while True:
        time.sleep(10)
        sort = requests.post(uri, headers=headers, data=data)
        try:
            result = sort.json()
            if result['errorCode'] == 106:
                print('ループ中')
        except json.decoder.JSONDecodeError:
            result = sort.text
            break

    df = pd.read_csv(io.StringIO(result))
    df = df.fillna("")
    now_date = datetime.now()
    save_date = datetime.strftime(now_date, '%Y.%m.%d %H-%M')
    filename = f'\demo{save_date}.csv'
    df.to_csv(f'{PATH}{filename}', index=False, encoding='cp932')