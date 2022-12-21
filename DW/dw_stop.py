import ast
import configparser
from datetime import timedelta
from datetime import datetime
import os
import pandas as pd
import pathlib
import pyautogui
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions
import time
from webdriver_manager.chrome import ChromeDriverManager
import logging

from kki_function import slack

TOKEN = "FJONUUiXQb7dQafFyi_GtFXc7baNdQpF"
PIN = 'V0'
PIN1 = 'V1'
PIN2 = 'V2'
PIN3 = 'V3'
PIN4 = 'V4'
PIN5 = 'V5'


# chromeを開くための関数
def driver_get():
    chop = webdriver.ChromeOptions()
    # prefs = {"download.default_directory": download_directory}
    # chop.add_experimental_option("prefs", prefs)
    # chop.add_argument('--user-data-dir=C:\\RPA\\Chrome\\User Data')
    chop.add_argument('--profile-directory=Default')
    chop.add_argument('--ignore-certificate-errors')  # SSLエラー対策
    # chop.add_argument('--headless')
    # chromedriverのPATHを指定（Pythonファイルと同じフォルダの場合）
    # driver_path = base_dir + 'driver\\chromedriver.exe'
    # Chrome起動
    driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=chop)
    # 要素がロードされるまでの待ち時間を10秒に設定
    driver.implicitly_wait(10)
    # 画面サイズ最大化
    driver.maximize_window()
    return driver


# 指定したピンの状態を取得
def read(token, pin):
    URI = "https://blynk.cloud/external/api/get"
    params = {
        "token": token,
        "pin": pin
    }

    res = requests.get(url=URI, params=params)
    return res.json()


# 指定した複数のピンの状態を取得
def multi_read(token, pin, pin1, pin2, pin3, pin4, pin5):
    URI = f"https://blynk.cloud/external/api/get?token={token}&{pin}&{pin1}&{pin2}&{pin3}&{pin4}&{pin5}"

    res = requests.get(url=URI)
    return res.json()


# 指定したピンにメッセージを書き込む
def write(token, pin, value):
    api_url = "https://blynk.cloud/external/api/update?token=" + token + "&" + pin + "=" + value
    params = {
        "token": token,
        "pin": pin,
        "value": value
    }

    requests.get(url=api_url, params=params)


def dw_stop_input(driver, loginid, password, gouki, stop_no, start_date, start_time, end_date, end_time):
    """
    :param driver:
    :param loginid: dwのログインID
    :param password: dwのpass
    :param gouki: 対応号機
    :param stop_no:　dwの停止入力用indexNo
    :param start_date:開始日付
    :param start_time:開始時間
    :param end_data:終了日付
    :param end_time:終了時間
    :return:
    """
    dw_url = "https://www.sdknet2.jp/login"
    driver.get(dw_url)
    # 要素が現れるまで待機
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "form-text")))
    time.sleep(1)
    # loginID
    driver.find_element(by=By.ID, value="edit-name").send_keys(loginid)
    # PASSWORD
    driver.find_element(by=By.ID, value="edit-pass").send_keys(password)
    # ログインをクリック
    time.sleep(2)
    dw_login = driver.find_element(by=By.ID, value="edit-submit")
    driver.execute_script('arguments[0].click();', dw_login)
    # # 指定したdriverに対して最大で60秒間待つように設定する
    # wait = WebDriverWait(driver, 60)
    # # 指定したボタンが表示されクリック出来る状態になるまで待機する
    # wait.until(expected_conditions.element_to_be_clickable((By.ID, "breadcrumb-item")))
    time.sleep(15)
    # 設定ボタンを見つける
    set_click = driver.find_element(by=By.ID, value="show-sidebar-right")
    print(set_click)
    driver.execute_script('arguments[0].click();', set_click)

    data_click = "https://www.sdknet2.jp/alter-operation"
    driver.get(data_click)
    input_button = driver.find_elements(by=By.CLASS_NAME, value="material-icons")
    print(input_button)
    driver.execute_script('arguments[0].click();', input_button[2])
    print("test")
    time.sleep(2)
    device = driver.find_element(by=By.NAME, value="gd_machine_id")
    select = Select(device)
    select.select_by_index(gouki)
    time.sleep(1)

    # 各elementを変数に代入
    start_date_elem = driver.find_element(by=By.NAME, value="startdate")
    start_time_elem = driver.find_element(by=By.NAME, value="starttime")
    end_date_elem = driver.find_element(by=By.NAME, value="enddate")
    end_time_elem = driver.find_element(by=By.NAME, value="endtime")
    data_kubun = driver.find_element(by=By.NAME, value="gd_data_class_id")
    new_register = driver.find_element(by=By.NAME, value="op")
    start_result = driver.find_element(by=By.ID, value="edit-startdate")
    end_result = driver.find_element(by=By.ID, value="edit-enddate")
    edit_submit = driver.find_element(by=By.ID, value="edit-submit")

    time.sleep(1)
    start_date_elem.send_keys(Keys.DELETE)
    start_date_elem.send_keys(start_date)
    time.sleep(1)
    start_time_elem.send_keys(Keys.DELETE)
    start_time_elem.send_keys(start_time)
    time.sleep(1)
    end_date_elem.send_keys(Keys.DELETE)
    end_date_elem.send_keys(end_date)
    time.sleep(1)
    end_time_elem.send_keys(Keys.DELETE)
    end_time_elem.send_keys(end_time)
    time.sleep(1)
    data_select = Select(data_kubun)
    data_select.select_by_index(stop_no)
    time.sleep(1)
    new_register.click()
    time.sleep(1)
    start_result.send_keys(Keys.DELETE)
    start_result.send_keys(start_date)
    time.sleep(1)
    end_result.send_keys(Keys.DELETE)
    end_result.send_keys(end_date)
    time.sleep(1)
    # edit_submit.click()
    time.sleep(60)
    print(gouki)
    driver.close()


# 値からキーを取得
def get_keys_from_value(d, val):
    keys = {}
    for k, v in d.items():
        if v == val:
            keys[k] = v
    return keys


def main():
    # slackTOKEN
    slack_token = "xoxb-3109530806049-3413685506163-bot1UXc4Of9aTbY4Ud4LdOii"
    # channel
    channel = "C032X6T56P6"
    # ファイルから情報を取得する
    config_ini = configparser.ConfigParser()
    config_ini.read(filenames="config.ini", encoding="utf-8")
    # ログイン名、パスワード、号機取得、停止要因
    loginid = config_ini["dw"]["id"]
    password = config_ini["dw"]["pw"]
    gouki = config_ini["dw"]["gouki"]
    stop_list = ast.literal_eval(config_ini["stop"]["cause"])
    no_list = ast.literal_eval(config_ini["no"]["stop"])

    flg = 0
    cnt = 0
    while True:
        # 全てのボタンの状態を取得
        buttons = multi_read(TOKEN, PIN, PIN1, PIN2, PIN3, PIN4, PIN5)
        print(f'start---------------------------')
        # ONになっているボタンのキーを取得
        keys = get_keys_from_value(buttons, 1)
        print(f'keys:{keys.keys()}')
        # ONになっているボタンが何個あるか取得
        status = len(keys)

        # ボタンが１つなら
        if status == 1 and cnt == 0:
            flg = 1
            before_key = 0
            for k, v in keys.items():
                stop_name = stop_list[k]
                stop_no = no_list[stop_name]
                before_key = k
                before_val = v
                # print(f'stop_name:{stop_name}stop_no:{stop_no}')
                if stop_name == "荷待ち":
                    slack.slack_post(slack_token, channel, f'{gouki}号機で{stop_name}が発生しています。')
                break
            # write(TOKEN, "V6", "              　   ")
        elif status > 1:
            # write(TOKEN, "V6", "2つ以上選択されています。")
            print('ボタンが2個選択')
            continue
        elif flg == 1 and buttons[before_key] != before_val:
            flg = 2
        elif status == 0:
            write(TOKEN, "V6", "              　   ")
            continue

        if flg == 1 and cnt == 0:
            start_date = datetime.now().strftime("%Y-%m-%d")
            start_time = datetime.now().strftime("%H:%M")
            write(TOKEN, "V6", f'{start_time}:{stop_name}')
            cnt = 1
            print(f'date:{start_date} time:{start_time}:{stop_name}')
        elif flg == 2:
            end_date = datetime.now().strftime("%Y-%m-%d")
            if stop_name == "加工終了":
                end_time = "22:00"
            else:
                end_time = datetime.now().strftime("%H:%M")
            print(f"DW書き込み開始:{end_date},time:{end_time}:{stop_name}")
            write(TOKEN, "V6", "更新中。。。")
            # driver = driver_get()
            # dw_stop_input(driver, loginid, password, gouki, stop_no, start_date, start_time, end_date, end_time)
            flg = 0
            cnt = 0
            print(f"DW書き込み完了")
            write(TOKEN, "V6", "更新完了")
            time.sleep(3)
            write(TOKEN, "V6", "　　　　")

        time.sleep(2)


if __name__ == '__main__':
    main()


