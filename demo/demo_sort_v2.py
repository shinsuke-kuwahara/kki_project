###### tags: `OCR_DEMO`
# 5-1demo_sort

# Slackへメッセージを飛ばすモジュール
import requests
# フォルダからワイルドカードで格納されているファイル全てを取り出す為のライブラリ
import glob
# 時間を計測するモジュール
import time
# ログの吐きだしに使用するモジュール
import logging
# ブラウザを自動操作させるモジュール
# seleniumはバージョンが4以上だとfind_by_elementのメソッドが異なるので注意
# バジョーンの確認はpycharmのターミナルでpip listを叩く
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

# ブラウザ自動操作時の待機に使用するモジュール
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Webブラウザを自動操作する（python -m pip install selenium)
# import chromedriver_binary
# ファイルやフォルダパスの取得に使用するモジュール
import os
# zipfileの読み込みに使用
import zipfile
# ファイルの移動に使用
import shutil
# テキストファイルをcsvに変換させるためpandasをインポート
import pandas as pd
# 年数を直打ちでなく変数で取得するためのモジュール
import datetime
# プログラムに何かあった場合終了させるモジュール
import sys
# エラー内容を取得するためのモジュール
import traceback
# メモリを解放するモジュール
import gc
# Slackへ通知を行うモジュール
# from　slack_post import demo_slack_post
# ダウンロードが完了するまで待機させる関数
def dl_wait():
    # 待機タイムアウト時間(秒)設定
    timeout_second = 10
    # ダウンロード先のフォルダパスを変数に入れておく
    # ダウンロード先をデフォルトのダウンロードフォルダから変更する場合はこのパスを変更してください
    dl = os.path.expanduser('~') + "\Downloads"
    # 指定時間分待機
    for i in range(timeout_second + 1):
        # ファイル一覧取得
        download_fileName = glob.glob(f'{dl}\\*.*')
        # ファイルが存在する場合
        if download_fileName:
            # 拡張子の抽出
            extension = os.path.splitext(download_fileName[0])
            # 拡張子が '.crdownload' ではない ダウンロード完了 待機を抜ける
            if ".crdownload" not in extension[1]:
                time.sleep(3)
                break
        # 指定時間待っても .crdownload 以外のファイルが確認できない場合 エラー
        if i >= timeout_second:
            # エラー内容を通知
            # slack_post("時間内にダウンロードが行えなかった為プログラムを終了しました")
            print("時間内にダウンロードが行えなかった為プログラムを終了しました")
            # ブラウザを閉じる
            driver.quit()
            # プログラムを終了する
            sys.exit()
        # 一秒待つ
        time.sleep(1)

    # メモリを解放する為に変数を削除する処理を追加
    del dl, download_fileName, extension
    gc.collect()

# ある要素あ表示するまで待つ関数を定義する
def selector_wait(myClassName):
    # 指定したdriverに対して最大で10秒間待つように設定する
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, myClassName))
    )
    # メモリを解放する為に変数を削除する処理を追加
    del element
    gc.collect()

# フォルダやファイル名に現在日時を使用するので、定義しておく
# 現在時刻を取得
now = datetime.datetime.now()
current_time = now.strftime('%Y-%m-%d-%H-%M-%S')

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
logfile = '格納したいファイルパスをファイル名、拡張子を含めて指定'
formatter = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(level=logging.INFO, format=formatter, filename=logfile)

""" ↑logの設定 """

""" ↓chromeの設定 """
# chromedriverの設定
# ドライバーの位置を明示する1
driver = webdriver.Chrome(ChromeDriverManager().install())
# 指定した要素が見つかるまでの待ち時間を設定する 今回は最大10秒待機する
driver.implicitly_wait(10)

# ヘッドレスChromeでファイルダウンロードするにはここが必要だった
# ここを入れないとhwadlessモードでのダウンロードはデフォルトでは出来ない設定になっているらしい
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
driver.execute("send_command", {
    'cmd': 'Page.setDownloadBehavior',
    'params': {
        'behavior': 'allow',
        'downloadPath': os.path.expanduser('~') + "\Downloads" # ダウンロード先のフォルダパスを書く
     }
})

""" ↑chromeの設定 """

def Sort(current_time, new_folder_path):

    """" ↓ここからseleniumの仕分け処理 """
    # 失敗した場合を想定して最大で3回実行するようにする
    for _ in range(3):
        try:
            # DXSiteを開く
            driver.get('https://kki-1105.dx-suite.com/ConsoleWeb/')  # DXSiteを開く

            # 指定した要素が見つかるまでの待ち時間を設定する 今回は最大10秒待機する
            # driver.implicitly_wait(10)

            # ログインID入力ボックスの要素をid属性値から取得
            element_ID = driver.find_element_by_id("loginId")

            # ログインID入力ボックスにIDを入力
            element_ID.send_keys("kkitech2")

            # メモリ解放のため削除
            del element_ID
            gc.collect()

            # ログインパスワード入力ボックスの要素をid属性値から取得
            element_Pass = driver.find_element_by_id("password")

            # ログインパスワード入力ボックスにIDを入力
            element_Pass.send_keys("kki_1105")

            # メモリ解放のため削除
            del element_Pass
            gc.collect()

            # ログインボタンをクリックする
            login_btn = driver.find_element_by_id('login-submit')
            login_btn.click()

            # メモリ解放のため削除
            del login_btn
            gc.collect()

        # エラーが出た場合はエラー内容をプリントする
        except Exception as e:
            traceback.print_exc()
            # ログにエラーの情報をかかせる
            logging.exception("取得したエラーは下記の内容です")

        # 失敗せずに処理を実行出来た場合はループを抜ける
        else:
            break
    # リトライが全部失敗したときの処理
    else:
        # エラー内容を通知
        # slack_post("ログイン処理が実行出来なかった為プログラムを終了しました")
        print("ログイン処理が実行出来なかった為プログラムを終了しました")
        # ブラウザを閉じる
        driver.quit()
        # プログラムを終了する
        sys.exit()

    """ ↑ここまでがログイン処理 """
    # 失敗した場合を想定して最大で3回実行するようにする
    for _ in range(3):
        try:

            # ElasticSorterボタンをクリックする
            Elastic_btn = driver.find_element_by_xpath("//img[contains(@src,'sorter_enable.png')]")
            Elastic_btn.click()

            # メモリ解放のため削除
            del Elastic_btn
            gc.collect()

        # エラーが出た場合はエラー内容をプリントする
        except Exception as e:
            traceback.print_exc()
            # ログにエラーの情報をかかせる
            logging.exception("取得したエラーは下記の内容です")

        # 失敗せずに処理を実行出来た場合はループを抜ける
        else:
            break
    # リトライが全部失敗したときの処理
    else:
        # エラー内容を通知
        # slack_post("ElasticSorterに入れなかった為プログラムを終了しました")
        print("ElasticSorterに入れなかった為プログラムを終了しました")
        # ブラウザを閉じる
        driver.quit()
        # プログラムを終了する
        sys.exit()

    """ ↑ここまでがElasticSorterに入るところ """
    # 失敗した場合を想定して最大で3回実行するようにする
    for _ in range(3):
        try:
            # ワイチFAXの仕分けルールボタンをクリックする
            waiti_btn = driver.find_element_by_xpath("//img[contains(@src,'Sorter/Rule/Thumbnail?sorter_rule_id=36466')]")
            waiti_btn.click()

            # 要素が表示されるまで待機（最大で10秒待機）
            selector_wait("unit-item-divider")

            time.sleep(5)

            # 最上段の詳細をクリックする btn btn-secondary input-group-btn col-12
            detail_btn = driver.find_elements_by_tag_name('button')[0].click()

            # 要素が表示されるまで待機（最大で10秒待機）
            selector_wait("thumbnail-img")

        # エラーが出た場合はエラー内容をプリントする
        except Exception as e:
            traceback.print_exc()
            # ログにエラーの情報をかかせる
            logging.exception("取得したエラーは下記の内容です")

        # 失敗せずに処理を実行出来た場合はループを抜ける
        else:
            break
    # リトライが全部失敗したときの処理
    else:
        # エラー内容を通知
        # slack_post("詳細を見るに入れなかった為プログラムを終了しました")
        print("詳細を見るに入れなかった為プログラムを終了しました")
        # ブラウザを閉じる
        driver.quit()
        # プログラムを終了する
        sys.exit()

    """ ↑ここまでがElasticSorterの読み取り結果最上段の「詳細を見る」に入るところ """
    # 無限ループが怖いのでcountを設けて20回繰り返してもダメな場合はプログラムを終了させる
    # 変数countを用意する
    count = 0
    # ボタンの要素数が11個か確認
    print(len(driver.find_elements_by_tag_name("button")))
    # 詳細のページが表示されるのをボタンの要素が23個であるか確認することでその間待機させる
    # 要素数はサイト無いのElasticSorterの結果を数えただけなので合わない場合は修正してください
    while not len(driver.find_elements_by_tag_name("button")) == 13 and not count == 20:
        # print("ボタンの要素数が合わない為1秒待機させます")
        time.sleep(1)

        count += 1
        # print("countは" + str(count) + "です")

    else:
        if count == 20:
            print("ElasticSorterの詳細に入れなかった為プログラムを終了しました")
            # エラー内容を通知
            # slack_post("ElasticSorterの詳細に入れなかった為プログラムを終了しました")
            # ブラウザを閉じる
            driver.quit()
            # メモリ解放のため削除
            del count
            gc.collect()
            # プログラムを終了する
            sys.exit()

        else:
            # メモリ解放のため削除
            del count
            gc.collect()
            pass



    # 繰り返しで回してダウンロードしていく仕組みを作成する
    # ダウンロードフォルダまでのパスを定義しておく
    # ダウンロードフォルダまでのパスを変数に入れる
    dl_path = os.path.expanduser('~') + "\Downloads"
    # print(dl_path)
    # 変数iは各トレイに仕分結果があるか判定するのに使用する変数
    i = 4

    try:
        for j in range(4):

            for _ in range(3):  # 最大3回実行
                try:
                    # 各トレイに仕分け結果が入っているか確認す為の変数
                    sort_btn = driver.find_elements_by_tag_name("button")[i]

                except IndexError:
                    # print(str(i) + "番目のボタンの要素が取得できませんでした")
                    # ページを再読み込みするプログラムを追加
                    driver.refresh()
                    # 表示される時間を考慮して3秒待機
                    time.sleep(3)

                else:
                    break

            # リトライが全部失敗した時の処理
            else:
                # エラー内容を通知
                print("３回リトライしましたが" + str(i) + "番目のボタンの要素が取得できませんでした")
                # slack_post("３回リトライしましたが" + str(i) + "番目のボタンの要素が取得できませんでした")
                # ブラウザを閉じる
                driver.quit()
                # プログラムを終了する
                sys.exit()

            # print(len(driver.find_elements_by_tag_name("button")))
            # トレイに結果があればtrue無ければfalseが出力される
            # print(sort_btn.is_enabled())

            if sort_btn.is_enabled() == True:

                # 各トレイのダウンロードボタンをクリック
                driver.find_elements_by_tag_name("button")[i].click()

                # 5秒待機
                time.sleep(5)
                dl_wait()

                # ローカルのダウンロードフォルダ内にある最新ファイルのパスを取得する
                list_of_files = glob.glob(dl_path + "\\*")
                latest_file = max(list_of_files, key=os.path.getctime)
                # print(latest_file)

                # ダウンロードしたzipファイルを冒頭で作成したフォルダに解凍する
                # 解凍したフォルダには各トレイの名称が含まれているため、名称をブラウザ上で取得する必要がなくなった
                zp = zipfile.ZipFile(latest_file, "r")
                zp.extractall(path=new_folder_path)
                zp.close()

            elif sort_btn.is_enabled() == False:

                pass

            # buttonの真偽確認を行う変数リストの値を指定するもの
            i += 2
            # print(i)

    # エラーの場合Slackに通知が行くようにする
    except Exception as e:
        # ログにエラーの情報をかかせる
        logging.exception("取得したエラーは下記の内容です")

        # エラーの情報を出力する
        traceback.print_exc()

        # エラー内容を通知
        # slack_post("ブラウザの自動走査中にエラーが発生しました")
        print(("ブラウザの自動走査中にエラーが発生しました"))
        # ブラウザを閉じる
        driver.quit()
        # プログラムを終了する
        sys.exit()

    time.sleep(5)

    """ ここまでがOCR結果を指定のフォルダへダウンロードさせる部分 """
    # 各トレイ毎に処理を書くと長すぎるので繰り返し処理で回す
    # 各トレイ名が格納されたリストを作成する
    tray_name_list = ["demo_林商店", "demo_桑原農園", "demo_藤田工務店"]

    # 各トレイのIDが格納されたリストを作成する
    tray_id_list = ["document-921671", "document-921670", "document-921706"]

    # 変数tray_name_listに格納された要素数分繰り返し処理を行う
    g = 0
    files = []
    for g in range(3):
        print(g)

        if os.path.exists(new_folder_path + '/' + tray_name_list[g]) == True:
            files.append(glob.glob(new_folder_path + '/' + tray_name_list[g] + "/*"))
            print(tray_name_list[g] + '形式のFAXは' + str(len(files)) + "枚です")

    # 該当の帳票読み取り結果csvをセレニウムの操作で取得しに行く
    # ホームボタンをおしてホームへ戻る
    home_btn = driver.find_element_by_class_name('sorter-icon-button').click()

    # メモリ解放の為、定義した変数の値を削除
    del home_btn
    gc.collect()

    # 3秒待機
    time.sleep(3)

    # csv仕分けデータのDL
    Intelligent_btn = driver.find_element_by_id("unit-download-form")
    Intelligent_btn.click()

    # メモリ解放の為、定義した変数の値を削除
    del Intelligent_btn
    gc.collect()

    # 8秒待機
    time.sleep(8)

    driver.quit()

    # ダウンロード結果を格納するフォルダを新規作成
    # ここに後々、仕分け結果のトレイ毎のフォルダが入ります。
    new_folder_path = './sortresult/' + current_time
    os.makedirs(new_folder_path, exist_ok=True)

    # ローカルのダウンロードフォルダ内にあるcsvファイルのパスを一覧で取得する
    csv_file = glob.glob(dl_path + "\\*.csv")
    print(csv_file)
    # ローカルのダウンロードフォルダ内にある最新ファイルのパスを取得する
    latest_csv = max(csv_file, key=os.path.getctime)
    print(latest_csv)

    # 取得したCSVファイルパスを指定フォルダへ移動させる
    new_csv_path = shutil.move(latest_csv, "./sortresult/" + current_time)
    # print(new_csv_path)

    # 読み取り結果のCSVを読み込み品名を抽出する
    csv_df = pd.read_csv(new_csv_path, encoding="shift-jis", usecols=[4]).values.tolist()
    print(csv_df)

    # 取得したcsvの上から順に品名を取り出す際に使う値の変数
    num = 0
    # リネームするファイル名が重複した場合品名末尾に「_数字」を
    # つける為に使用する変数
    exsist_num = 1
    for file in files:
        try:
            print(file)
            print(files)
            # ファイルの名称に使用できない文字がcsv_df[num][0]に含まれている場合があるのでその場合はリネームして削除する
            if ":" in str(csv_df[num][0]):
                csv_df[num][0] = csv_df[num][0].replace(":", "")

            elif "/" in str(csv_df[num][0]):
                csv_df[num][0] = csv_df[num][0].replace("/", "")

            elif "?" in str(csv_df[num][0]):
                csv_df[num][0] = csv_df[num][0].replace("?", "")

            elif "<" in str(csv_df[num][0]):
                csv_df[num][0] = csv_df[num][0].replace("<", "")

            elif ">" in str(csv_df[num][0]):
                csv_df[num][0] = csv_df[num][0].replace(">", "")

            elif "|" in str(csv_df[num][0]):
                csv_df[num][0] = csv_df[num][0].replace("|", "")

            # 読み取り結果の画像をcsvの品名情報を元に全てリネームする
            os.rename(file[0], './getdata/' + current_time + '/' + tray_name_list[num] + '/' + str(csv_df[num][0]) + '.jpg')
            print(str(tray_name_list[g]))

            num += 1

        # 同一品名が存在していた場合、リネーム時にエラーが発生するので例外処理させる
        except FileExistsError:
            # print(file)

            # 読み取り結果の画像をcsvの品名情報を元に全てリネームする
            os.rename(file[0], './getdata/' + current_time + '/' + tray_name_list[g] + '/' +
                      str(csv_df[num][0]) + "_" + str(exsist_num + 1) + '.jpg')

            # webhookにダブりありましたよメッセージを完了時おくるようにする
            # slack_post("<!channel>\n重複ファイルがありました（" + str(csv_df[num][0]) + "）")
            print("重複ファイルがありました"+str(csv_df[num][0]) )

            num += 1
            exsist_num += 1

    """ ここまでが仕分け結果の各ファイルをcsvの結果を元にリネームする部分　"""

    # サブフォルダのパスを格納するリスト
    manual_input = []

    for fd_path, sb_folder, sb_file in os.walk(new_folder_path):
        for fol in sb_folder:
            sub_fol = fd_path + '\\' + fol
            print(sub_fol)

            manual_input.append(sub_fol)

# Sort(current_time,new_folder_path)
