import ast
import configparser
from datetime import datetime
import glob
import logging
import os
import sys
import subprocess
import threading
import time
import tkinter
from tkinter import messagebox
import pandas as pd
import pyautogui

from kki import effect_log


# データ保存用の日付取得
date_now = datetime.now()
save_date = datetime.strftime(date_now, "%Y.%m.%d %H-%M")
# 製造管理入力用日付
input_date = datetime.strftime(date_now, "%Y%m%d")
# ユーザーネーム取得
user_name = '"' + os.environ['USERNAME'] + '"'
# デスクトップパスを取得する
desktop_path = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + "\\Desktop"
# ログファイル
log_file = "log\ログ(" + str(save_date) + ").log"


# インターフェースを画面中央へ
def tk_center(w):
    w.update_idletasks()
    ww = w.winfo_screenwidth()
    lw = w.winfo_width()
    wh = w.winfo_screenheight()
    lh = w.winfo_height()
    w.geometry(str(lw) + "x" + str(lh) + "+" + str(int(ww / 2 - lw / 2)) + "+" + str(int(wh / 2 - lh / 2)))


# インターフェースを画面右上へ
def tk_right(w):
    w.update_idletasks()
    ww = w.winfo_screenwidth()
    lw = w.winfo_width()
    wh = w.winfo_screenheight()
    lh = w.winfo_height()
    w.geometry(str(lw) + "x" + str(lh) + "+" + str(int(ww / 4 * 3 - lw / 2)) + "+" + str(int(wh / 2 - lh / 2)))


def gui_start():
    global root
    root = tkinter.Tk()
    # メイン画面作成
    root.title('プログラム起動')
    root.geometry('300x100')
    # guiを画面センターへ
    tk_center(root)
    # ラベル作成
    label_name = tkinter.Label(text='処理を開始しますか？')
    label_name.place(x=50, y=15)
    # ボタン作成
    button_start = tkinter.Button(text="確定", width=12)
    button_cancel = tkinter.Button(text="取消", width=12)
    button_start.place(x=40, y=55)
    button_cancel.place(x=160, y=55)

    button_start["command"] = start
    button_cancel["command"] = cancel
    root.mainloop()
    del label_name


def gui_stop():
    global win
    win = tkinter.Tk()
    win.title('プログラム実行中')
    win.geometry('300x100')
    win.attributes('-topmost', True)
    tk_right(win)
    label_msg = tkinter.Label(text='処理中です。しばらくお待ちください。')
    label_msg.place(x=65, y=15)
    button_stop = tkinter.Button(text='終了', width=12)
    button_stop.place(x=100, y=55)
    button_stop['command'] = stop
    win.mainloop()


def stop():
    messagebox.showinfo(title='処理終了', message='処理を終了します。', icon='warning')
    sys.exit()


def cancel():
    msg_cancel = messagebox.askyesno(title="アプリケーション終了", message="処理を中止しますか？")
    if msg_cancel:
        sys.exit()


def start():
    global thread
    msg_start = messagebox.askyesno(title="プログラム起動確認", message="処理を開始しますか？")
    if msg_start:
        root.destroy()
        thread = threading.Thread(target=main, daemon=True)
        thread.start()
        gui_stop()


def order_delete():
    """
    date_naw:現在の日付
    save_date:日付の表示方法指定
    file_name:ファイル名
    save_dir:保存先
    return df_orders, path_txt, file_num
    """
    # 受注一覧の参照先
    path_txt = glob.glob(desktop_path + r"\TXT_NGV\受注実績一覧表*.TXT")
    # フォルダ内の受注一覧表の数を取得
    file_num = len(path_txt)
    # 空のデータフレーム生成
    df_orders = pd.DataFrame()
    # 読み込んだテキストファイルを繰り返しdf_ordersに格納
    for i in range(file_num):
        # データ読み込み
        df_order = pd.read_csv(path_txt[i], encoding="cp932", dtype={"商品コード": object})
        # df_ordersに追記
        df_orders = pd.concat([df_orders, df_order], axis=0, ignore_index=True)
        # 受注dataのファイル名を変更
        # os.rename(path_txt[i], path_rename_txt + "(" + str(save_date) + "-" + str(i) + ").TXT")
    # データフレームに相手先ロットＮＯを追加
    df_orders["相手先ロットＮＯ"] = ""
    # 製造管理の入力順に並び替え
    df_orders = df_orders.reindex(columns=["ロットＮＯ", "相手先ロットＮＯ", "納期日付", "商品コード", "商品名", "商品名２", "受注ｍ数", "内容１", "内容２"])
    # ワイチLOTのアルファベット入力用にテキストファイルから読み込み
    config_ini = configparser.ConfigParser()
    config_ini.read(filenames="config.ini", encoding="utf-8")
    # 各変数に読み込み先を代入
    lot_new = ast.literal_eval(config_ini["lot"]["new"])
    print(df_orders.dtypes)
    num = 0
    # 空白があるとstr型にならないので強制的に変更
    df_orders.fillna("", inplace=True)
    for lot in df_orders["ロットＮＯ"]:
        # 日付の"."を削除する
        df_orders["納期日付"][num] = df_orders["納期日付"][num].replace(".", "")
        # KKI加工以外を削除する
        if "KKI" not in df_orders["内容１"][num] and "KKI" not in df_orders["内容２"][num]:
            df_orders.drop(index=num, inplace=True)
        # ワイチロットなら相手先ロットNOに入力
        for i, fuji in lot_new.items():
            if int(str(lot)[:2]) == int(i):
                df_orders["相手先ロットＮＯ"][num] = fuji + str(lot)[2:]
                break
        num += 1
    df_orders.reset_index(drop=True, inplace=True)
    df_orders.to_csv(desktop_path + r"\TXT_NGV\受注一覧表\受注data(txt)\orders_old.csv", index=False, encoding="cp932")
    return df_orders, path_txt, file_num


def entry_product(imagefile, order, path_txt, file_num):
    """
    :param imagefile: 読み込みたい画像ファイルパス
    :param order: 関数order_deleteのdf_orderを受け取る
    :return:
    """
    # 保存先変更パス
    path_rename_txt = desktop_path + r"\TXT_NGV\受注一覧表\受注data(txt)\受注実績一覧表"
    # 入力用のデータ保存先
    save_csv = desktop_path + r"\TXT_NGV\受注一覧表\受注data(csv)"
    # 入力用のファイル名
    file_name_order = r"\NGV受注書(" + str(save_date) + ").csv"
    # 登録なしリストの保存先とファイル名
    save_dir = desktop_path + r"\TXT_NGV\登録なしリスト\登録なしリスト(" + str(save_date) + ").csv"
    # 重複リストの保存先とファイル名
    duplicate_path = ".\重複フォルダ\訂正リスト\訂正リスト(" + str(save_date) + ").csv"
    # 相手先ロット重複ファイル名
    lot_path = f".\重複フォルダ\相手ロット重複\相手先重複({save_date}.csv"
    # 訂正用リスト作成
    df_fix = pd.DataFrame(index=[], columns=order.columns)
    # 転記完了リスト用
    df_entry = pd.DataFrame(index=[], columns=order.columns)
    # 製造管理アイコン
    hankan = desktop_path + "/KKI販売管理.RDP"
    # 商品登録がないと赤くなるのでその画像を取得
    not_product = "image/err_sys.PNG"
    # 更新ロック回避用の画像取得
    lock_sys = "image/sys_lock.PNG"
    # 製造管理メイン画像
    system_main = "image/system.PNG"
    # 入力画面画像
    input_screen = "image/input_screen.PNG"
    # 受注実績一覧表画面の画像
    orders_list = "image/txt_main.PNG"
    # 受注実績一覧表プリント画面
    txt_screen = "image/txt_screen.PNG"
    # 訂正用画像
    sys_fix = "image/sys_fix.PNG"
    # リモートエラー画像
    remote_err = r"image\connect_err.PNG"
    # 相手先ロット重複用画像
    lot_err = r"image\err_lot.PNG"

    # 画面からimageを見つけたらその座標を取得し、アイコンをクリックする
    count = 0
    while True:
        # 製造管理システムアイコンをsys_imageに格納
        sys_image = list(pyautogui.locateAllOnScreen(imagefile, confidence=0.9))
        if sys_image:
            x, y = pyautogui.locateCenterOnScreen(imagefile, confidence=0.9)
            pyautogui.click(x, y, 2)
            break
        else:
            # 3回読み込みに失敗したらシステムを終了する
            if count < 3:
                messagebox.showinfo(title="アイコンの確認", message="製造管理アイコンが見つかりません。\n指定の位置に配置してください。", icon="warning")
                count += 1
            else:
                messagebox.showinfo(title="システムの終了", message="３回失敗しました。システムを終了します。", icon="warning")
                win.destroy()
                sys.exit()

    # システム画面が立ち上がったら処理を開始する
    while True:
        if pyautogui.locateOnScreen(system_main):
            break

    # 販売管理を２重で開いてしまうことがあるのでアラートを出し回避する、3回ミスでプログラム終了
    time.sleep(1)
    if pyautogui.locateOnScreen(remote_err, confidence=0.9):
        messagebox.showinfo(title="リモートデスクトップ接続失敗",
                            message="販売管理が2重で開かれています。リモートデスクトップ接続メッセージを\n閉じてからもう一度プログラムを開始してください。",
                            icon="warning")
        win.destroy()
        sys.exit()

    # 受注入力を選択する
    pyautogui.press("N")
    pyautogui.press("tab")
    pyautogui.press("enter")
    # 入力画面が現れたら処理を開始する
    while True:
        if pyautogui.locateOnScreen(input_screen):
            break

    # 商品登録がない製品リスト用にデータフレームを作成
    df_list = pd.DataFrame(columns=["ロットＮＯ", "相手先ロットＮＯ", "商品コード", "商品名", "商品名２", "受注ｍ数"])
    # 受注件数取得
    num = len(order)
    # 登録件数用カウント
    entry_cnt = 0
    # NGV受注すべて製造管理に登録する
    for i in range(num):
        time.sleep(1.5)
        pyautogui.typewrite("1")
        pyautogui.press("tab")
        pyautogui.typewrite("1")
        pyautogui.press("tab")
        time.sleep(1)
        pyautogui.typewrite(str(order["ロットＮＯ"][i]))
        pyautogui.press("tab")
        # 重複商品、一時保管用
        items = []
        time.sleep(1)
        # 入力内容が訂正ならリストに格納し処理を飛ばす
        if pyautogui.locateOnScreen(sys_fix, confidence=0.9):
            for item in order:
                items.append(order[item][i])
            df_fix.loc[i] = items
            pyautogui.keyDown("shift")
            pyautogui.press("tab", presses=3)
            pyautogui.keyUp("shift")
            continue

        # 相手先ロットに値があれば相手先ロットNOを入力する、なければ処理を飛ばす
        lot_item = []
        df_lot = pd.DataFrame(index=[], columns=order.columns)
        if order["相手先ロットＮＯ"][i] == "":
            pyautogui.press("tab", presses=5)
        else:
            pyautogui.keyDown("shift")
            pyautogui.typewrite(str(order["相手先ロットＮＯ"][i])[0:1])
            pyautogui.keyDown("shift")
            pyautogui.typewrite(str(order["相手先ロットＮＯ"][i])[1:2])
            pyautogui.keyUp("shift")
            pyautogui.typewrite(str(order["相手先ロットＮＯ"][i])[2:6])
            pyautogui.press("tab", presses=5)
            if pyautogui.locateOnScreen(lot_err, confidence=0.9):
                for lot in order:
                    lot_item.append(order[lot][i])
                df_lot.loc[i] = lot_item
                pyautogui.keyDown("shift")
                pyautogui.press("tab", presses=3)
                pyautogui.keyUp("shift")
                continue

        pyautogui.typewrite(order["納期日付"][i])
        time.sleep(2)
        pyautogui.press("tab")
        pyautogui.typewrite(str(order["商品コード"][i]))
        pyautogui.press("tab")

        # 商品登録がなければリストに追加し繰り返しの先頭に戻る
        time.sleep(2)
        if pyautogui.locateOnScreen(not_product):
            df_list = df_list.append({"ロットＮＯ": order["ロットＮＯ"][i], "相手先ロットＮＯ": order["相手先ロットＮＯ"][i],
                                      "商品コード": order["商品コード"][i], "商品名": order["商品名"][i],
                                      "商品名２": order["商品名２"][i], "受注ｍ数": order["受注ｍ数"][i]}, ignore_index=True)
            pyautogui.press("up", presses=9)
            continue

        pyautogui.typewrite(str(order["受注ｍ数"][i]))
        pyautogui.press("F5")
        time.sleep(1)
        pyautogui.press("enter")
        # 更新ロック画面が現れたら更新が終了するまで待つ
        time.sleep(5)
        while True:
            if not pyautogui.locateOnScreen(lock_sys):
                break

        # 製造管理に登録ができたら、転記リストに追加する
        entry = []
        for entered in order:
            entry.append(order[entered][i])
        df_entry.loc[i] = entry
        entry_cnt += 1
    pyautogui.press("F10")
    time.sleep(1)
    # 重複リストがあるか確認用
    no_cnt = len(df_fix)
    # リストに要素があれば商品登録がないと判断しcsvファイルを立ち上げる
    count = len(df_list)
    # 相手先ロット重複用カウント
    lot_cnt = len(df_lot)
    # 訂正+商品登録なしが処理件数と一致しなければ入力データありと判断しプリントアウトする
    if no_cnt + count + lot_cnt != num:
        pyautogui.press("Z")
        pyautogui.press("tab")
        pyautogui.press("enter")
        # 受注一覧表の入力画面が現れるまで待つ
        while True:
            if pyautogui.locateOnScreen(orders_list):
                break

        pyautogui.typewrite(input_date)
        pyautogui.press("tab")
        pyautogui.typewrite(input_date)
        pyautogui.press("tab", presses=3)
        pyautogui.typewrite("1")
        pyautogui.press("tab", 3)
        pyautogui.typewrite("1")
        pyautogui.press("F5")
        pyautogui.press("enter")

        # プレビュー画面が現れるまで待つ
        while True:
            if pyautogui.locateOnScreen(txt_screen):
                break

        pyautogui.press("F5")
        pyautogui.press("tab", presses=2)
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.press("F10")

    # 製造管理を閉じる
    pyautogui.press("F")
    pyautogui.press("tab")
    pyautogui.press("enter")

    # 　商品なし＆重複ありが両方ある場合の処理
    if count > 0 and no_cnt > 0 and lot_cnt > 0:
        messagebox.showinfo(title="要確認",
                            message=f"{entry_cnt}件、受注しました。\n{no_cnt}件の重複ファイル、{count}件の商品登録、{lot_cnt}の相手先重複があります。確認してください。",
                            icon="warning")
        # 重複リストをフォルダに保存する
        df_fix.to_csv(duplicate_path, index=False, encoding="cp932")
        # 登録なしリストをフォルダに保存する
        df_list.to_csv(save_dir, index=False, encoding="cp932")
        subprocess.Popen([save_dir], shell=True)
        subprocess.Popen([duplicate_path], shell=True)
    # 相手先重複のみの処理
    elif lot_cnt > 0:
        messagebox.showinfo(
            title="要確認",
            message=f"{entry_cnt}件、受注しました。\n{lot_cnt}件、相手先の重複を確認しました。確認してください。",
            icon="warning"
        )
        df_lot.to_csv(lot_path, index=False, encoding="cp932")
        subprocess.Popen([lot_path], shell=True)
    # 重複ファイルのみある場合の処理
    elif no_cnt > 0:
        messagebox.showinfo(title="要確認",
                            message=str(entry_cnt) + "件、受注しました。\n" + str(no_cnt) + "件の重複ファイルが存在します。確認してください。",
                            icon="warning")
        df_fix.to_csv(duplicate_path, index=False, encoding="cp932")
        subprocess.Popen([duplicate_path], shell=True)
    # 商品なしファイルのみある場合の処理
    elif count > 0:
        messagebox.showinfo(title="要確認",
                            message=str(entry_cnt) + "件、受注しました。\n" + str(count) + "件の商品登録がありません。確認してください。",
                            icon="warning")
        df_list.to_csv(save_dir, index=False, shell=True)
        subprocess.Popen([save_dir], shell=True)
    else:
        messagebox.showinfo(title="プログラム終了", message=str(entry_cnt) + "件、受注しました。\nプログラムを終了します。")

    # 登録ファイルの格納先変更
    for i in range(file_num):
        os.rename(path_txt[i], path_rename_txt + "(" + str(save_date) + "-" + str(i) + ").TXT")
    # 転記リストを保存
    df_entry.to_csv(save_csv + file_name_order, index=False, encoding="cp932")


def main():
    try:
        start_time = time.time()
        order, path_txt, file_num = order_delete()
        if file_num != 0:
            entry_product("image/manage_sys.PNG", order, path_txt, file_num)
        else:
            messagebox.showinfo(title="テキストファイルの確認", message="フォルダ内にファイルがありません。確認してください。", icon="warning")
    except Exception as e:
            # ログを取得
            messagebox.showinfo(title="システムエラー", message="システムエラーが発生しました。\n管理者へ問い合わせしてください。", icon="warning")
            logging.basicConfig(level=logging.INFO, filename=log_file)
            logging.exception("エラーログ")
    win.destroy()
    end_time = time.time()


if __name__ == '__main__':
    gui_start()
