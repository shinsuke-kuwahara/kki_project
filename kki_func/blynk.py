import requests


def pin_read(token, pin):
    """
    :param token:string
    :param pin:string
    :return:
    """
    URI = f'https://blynk.cloud/external/api/get?token={token}&{pin}'

    res = requests.get(url=URI)
    return res.json()


# 指定した複数のピンの状態を取得
def pins_read(token, pins):
    """
    :param token:
    :param pins: list
    :return:
    """
    URI = f"https://blynk.cloud/external/api/get?token={token}"

    for p in pins:
        URI = f'{URI}&{p}'

    res = requests.get(url=URI)
    return res.json()


# 指定したピンにメッセージを書き込む
def pin_write(token, pin, value):
    """
    :param token:
    :param pin:   書き込みたいピンを指定
    :param value: string　”0”　or　”1"
    :return:
    """
    api_url = f"https://blynk.cloud/external/api/update"
    params = {
        "token": token,
        "pin": pin,
        "value": value
    }

    requests.get(url=api_url, params=params)


def pin_clear(token, pin, val):
    """
    :param token:string
    :param pin: string
    :param val: string
    :return:
    """
    URI = f'https://blynk.cloud/external/api/update?token={token}&{pin}={val}'

    res = requests.get(url=URI)


# ピンを全てOFFにする
def pins_clear(token, pins):
    """
    :param token:
    :param pins: list
    :return:
    """
    api_url = f"https://blynk.cloud/external/api/batch/update?token={token}"

    cnt = []
    for _ in pins:
        cnt.append("0")

    input_clear = dict(zip(pins, cnt))

    for p, v in input_clear.items():
        api_url = f'{api_url}&{p}={v}'

    requests.get(url=api_url)


# ピンを全てOFFにする
def pins_reset(token, pins, key):
    """
    :param token:
    :param pins: list
    :return:
    """
    api_url = f"https://blynk.cloud/external/api/batch/update?token={token}"

    cnt = []
    for k in pins:
        if key == k:
            cnt.append("1")
        else:
            cnt.append("0")

    input_clear = dict(zip(pins, cnt))

    for p, v in input_clear.items():
        api_url = f'{api_url}&{p}={v}'

    requests.get(url=api_url)


# LCDに2行書き込む
def lcd_write(token, pins, *args):
    """
    :param token:
    :param pins: list
    :param args: messages　"mes1", "mes2" ※mes2はなくてもよい
    :return:
    """
    clear = " "
    multi_write_api = f'https://blynk.cloud/external/api/batch/update?token={token}'

    cnt = len(args)
    if cnt == 0:
        mes = [clear, clear]
    elif cnt == 1:
        mes = [args[0], clear]
    else:
        mes = [args[0], args[1]]

    dict_mes = dict(zip(pins, mes))
    for p, v in dict_mes.items():
        multi_write_api = f'{multi_write_api}&{p}={v}'

    requests.get(url=multi_write_api)