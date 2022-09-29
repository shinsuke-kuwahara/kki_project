###### tags: `OCR_DEMO`
# 6-1demo_entry
""" ↓モジュール """
# メモ帳を開く為に使用するモジュール
import subprocess
import requests
# ログの吐きだしに使用するモジュール
import logging
# 年数を直打ちでなく変数で取得するためのモジュール
import datetime
# # 時間を計測するモジュール
import time
# プログラムに何かあった場合終了させるモジュール
import sys
# エラー内容を取得するためのモジュール
import traceback
import tkinter as tk
import threading
import pandas as pd
import pyautogui as gui
import pyperclip
# Slackへ通知を行うモジュール
from demo_slack_post import slack_post

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
logfile = ".\logfile\log.txt"
formatter = '%(asctime)s - %(levelname)s - %(message)s'
logging.basicConfig(level=logging.INFO, format=formatter, filename=logfile)
""" ↑logの設定 """

""" ↓グローバル変数 """
# ログの吐きだしに現在日時を使用するので、定義しておく
# 現在時刻を取得
now = datetime.datetime.now()
current_time = now.strftime('%Y-%m-%d-%H-%M-%S')
ymd = current_time[:8]
hm = current_time[8:]

# DFを成形するために必要なファイルパスを変数に入れ込むdatabase
# file_path = "./file_read/MOLD_1.xlsm"
# ローカルファイル動作確認用PATH
# file_path = "./file_read/集計.xlsm"
# disk1のPATH
file_path = r'\\10.100.108.150\disk1\テック事業部共有\demo\集計.xlsm'
# 成形したEXCELからデータフレームを取得する
# df_excel = pd.read_excel(file_path, sheet_name='MOLD_1',
df_excel = pd.read_excel(file_path, sheet_name='database',
                         usecols=["発注No", "発注日", "発注先", "商品コード", "品名",
                                  "数量", "単価", "金額"],
                         dtype=object)
# DFがカラム通り出力されているか確認
print(df_excel)

# 繰り返し処理の回数をだすのにdfの行数が必要な為ここで定義しておく
sum_rows = len(df_excel)
print(sum_rows)

# tkinterを用いてGUIを作成する
root = tk.Tk()
ww = root.winfo_screenwidth()
print(ww)
wh = root.winfo_screenheight()
print(wh)
# 200×90の画面サイズ、表示位置は割り算、掛け算でなんとなくだしてます（多分画面右下に表示されます）
root.geometry("200x90+" + str(int(ww / 4 * 3.5)) + "+" + str(int(wh / 4 * 2.6)))
# ボタンのウィンドウを最前面にしておく為の設定Falseで解除できる
root.attributes("-topmost", True)
# rootの背景色を変更
root.configure(bg='blue')
text_info = tk.StringVar()

# rootを別々の変数で定義することで同一のウィンドウにボタンとメッセージという別々の機能を付与する
# ボタン用
frame1 = tk.Frame(root)
# メッセージ用
frame2 = tk.Frame(root)
# ラベルに表示する文字のフォント設定
font_1 = ("MSゴシック", "9", "bold")

# 保存時に画像認識を使用するのであらかじめ画像のパスを定義しておく
# 今回は仮にデスクトップに保存させる場合に必要な画像を定義
# desktop_icon = r"C:\Users\kenji_karasawa\Desktop\projectDEMO\target\desktop_icon.png"
# file_icon = r"C:\Users\kenji_karasawa\Desktop\projectDEMO\target\file_icon.png"
# 転記完了した際にメモを閉じる為の画像
# close_icon = r"C:\Users\kenji_karasawa\Desktop\projectDEMO\target\close_icon.png"
# 登録ボタンをクリックさせる為の画像
entry_botton = r"C:\Users\kenji_karasawa\Desktop\projectDEMO\target\entry_botton.png"

flg = 0
# 　何件登録したかカウントする為の関数
entry_count = 0
""" ↑グローバル変数 """

""" ↓関数 """


def start():
    thread = threading.Thread(target=entry, daemon=True)
    thread.start()


def end():
    root.destroy()
    slack_post("終了ボタンが押された為\n処理の途中ですがプログラムを終了しました")
    sys.exit()


# entry関数に転記させるための内容を記載してください
def entry():
    global flg
    global df_excel
    global el_flg
    global entry_count
    global all_count

    try:
        # まずはメモを開きます
        # subprocess.Popen(r'C:\Windows\notepad.exe')
        # まずは自作GUIを開きます
        subprocess.Popen(r'C:\Users\kenji_karasawa\Desktop\基幹システムver001.exe')
        # GUIでお知らせ
        info('GUIに登録を開始します')

        # 画面の幅・高さを取得します
        # disp_width = gui.size().width
        # disp_height = gui.size().height

        # 早すぎると飛ばしそうなので２秒待機
        time.sleep(8)

        # 先程求めた幅・高さのそれぞれ半分の座標を指定することで画面中央を指定しクリックさせます
        # gui.click(x=disp_width / 2, y=disp_height / 2)

        # 早すぎると飛ばしそうなので２秒待機
        # time.sleep(2)

        # df_excelの行数分繰り返させる
        for k in range(sum_rows):

            if "nan" in str(df_excel["発注No"][k]):

                pyperclip.copy('発注Noなし')
                time.sleep(0.2)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

                # 発注Noがない時はentry_countを-1する
                if entry_count > 0:
                    entry_count -= 1
                    print("entry_countを-1しました。現在の値は" + str(entry_count) + "です")

            else:
                pyperclip.copy(str(df_excel["発注No"][k]))
                time.sleep(0.2)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "NaT" in str(df_excel["発注日"][k]):

                pyperclip.copy('発注日なし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                hattyuu = str(df_excel["発注日"][k])

                # pyperclip.copy(str(df_excel["発注日"][k]))
                time.sleep(0.1)

                hattyuu.replace('-', '/')

                pyperclip.copy(hattyuu[:10])
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "nan" in str(df_excel["発注先"][k]):

                pyperclip.copy('発注先なし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                pyperclip.copy(str(df_excel["発注先"][k]))
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "nan" in str(df_excel["商品コード"][k]):

                pyperclip.copy('商品コードなし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                pyperclip.copy(str(df_excel["商品コード"][k]))
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "nan" in str(df_excel["品名"][k]):

                pyperclip.copy('品名なし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                pyperclip.copy(str(df_excel["品名"][k]))
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "nan" in str(df_excel["数量"][k]) or "NaN" in str(df_excel["数量"][k]):

                pyperclip.copy('数量なし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                pyperclip.copy(str(df_excel["数量"][k]))
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "nan" in str(df_excel["単価"][k]):

                pyperclip.copy('単価なし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                pyperclip.copy(str(df_excel["単価"][k]))
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで次の項目へ
            gui.press('tab')
            time.sleep(0.1)

            if "nan" in str(df_excel["金額"][k]):

                pyperclip.copy('金額なし')
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            else:
                pyperclip.copy(str(df_excel["金額"][k]))
                time.sleep(0.1)
                # まずはctrl+Vでペーストさせる
                gui.keyDown('ctrl')
                gui.press('v')
                gui.keyUp('ctrl')

            time.sleep(0.1)
            # 項目の入力が終わったらtabで登録ボタンへ
            gui.press('tab')
            time.sleep(0.1)
            # 登録ボタンはクリックでないと押せないので画像判定させる
            if gui.locateOnScreen(entry_botton, confidence=0.9):
                while True:
                    print("表示なし")
                    # 　反応を待つ
                    time.sleep(1)
                    # アイコンの画像が見つかったら対象のアイコンをダブルクリック
                    if gui.locateOnScreen(entry_botton, confidence=0.9):
                        # 画像からx,y座標を割り出す
                        x, y = gui.locateCenterOnScreen(entry_botton, confidence=0.9)
                        # print(x, y)
                        # マウスを先程割り出したx,y座標へ移動させる
                        gui.moveTo(x, y)
                        # gui.doubleClick()
                        gui.click()
                        print("登録ボタンをクリック")
                        break
            time.sleep(0.3)
            # 頭の項目へ戻ります
            gui.press('tab')
            time.sleep(1.5)

            # ここで登録を押してGUIのメッセージ内容を消す
            gui.click()

            # entry_countをインクリメントしてメモ帳に記載できた件数を計算する
            entry_count += 1

            info('メモ帳への記載が' + str(entry_count) + '件完了しました')

        # GUIに終了メッセージを飛ばす
        info('登録が完了しました\nプログラムを終了します')
        time.sleep(2)
        root.destroy()
        sys.exit()

    except Exception as e:
        # エラーが発生した場合logファイルへエラー内容の保存を行わせる
        logging.info('エラーが発生しました')

        # エラーが発生した為、プログラムが終了したことをSlackで通知
        slack_post('エラーが発生した為プログラムを終了しました\n詳細はlogファイルを確認してくだい')

        # プログラムを終了させる為にrootを終わらせる
        root.destroy()
        sys.exit()


# guiにメッセージを表示させる関数
# 変数messageにGUIに表示させたいメッセージをstr型で格納してください
def info(message):
    text_info.set(message)
    label1 = tk.Label(frame2, textvariable=text_info, font=font_1).grid(row=0, column=0)
    frame2.pack()
    print(message)


def main():
    global entry_count
    global all_count
    # 起動したことを通知
    # コンソール非表示だと動いているかどうか分からない為動き出したことを通知
    slack_post('基幹システムver001への登録を開始しました')
    btn1 = tk.Button(frame1, text="開始", command=start).grid(row=0, column=0)
    btn3 = tk.Button(frame1, text="終了", command=end).grid(row=0, column=10)

    frame1.pack()

    # root.mainloop()の処理が完了しないとこの行以降の処理は行われない
    root.mainloop()
    # 正常に終了された場合に通知が飛ぶ
    slack_post('基幹システムver001へ' + str(entry_count) + '件の登録が完了しました')


if __name__ == '__main__':
    main()
