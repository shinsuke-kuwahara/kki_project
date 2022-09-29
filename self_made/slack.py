import requests


# ワークスペースK＆NのチャンネルロボPC使用状況に投稿する関数
def slack_robo_post(pname, num):
    """
    :param pname:コメント
    :param num: １・・・使用中、２・・・空室
    """
    # slackbotの情報
    bot_token = 'xoxb-994892783410-3441445086132-WyKdG5wv8l9GsXUsJ7YCfmqe'
    URL = 'https://slack.com/api/chat.postMessage'
    headers = {"Authorization": "Bearer " + bot_token}
    channel = "C026UF83FST"
    # K&Nのロボpc使用状況チャンネルに使用中を通知
    use = {
        'channel': channel,
        'text': str(pname) + '\n :使用中:',
        'as_user': True
    }
    # K&Nのロボpc使用状況チャンネルに空室を通知
    disuse = {
        'channel': channel,
        'text': str(pname) + '\n :空室:',
        'as_user': True
    }
    # 通知を分岐させる為の変数aを設ける
    a = num
    if a==1:
        requests.post(URL, headers=headers, data=use)

    elif a==2:
        requests.post(URL, headers=headers, data=disuse)


# slackにメッセージを投稿
def slack_post(token, channel, comment=''):
    """
    :param token: 投稿したいワークスペースのトークン
    :param channel: 投稿したいチャンネルのID
    :param comment: 投稿したい内容　メンションは<
    :return:
    """
    URL = 'https://slack.com/api/chat.postMessage'
    headers = {"Authorization": "Bearer " + token}

    data = {
        'channel': channel,
        'text': comment,
        'as_user': True
    }

    requests.post(url=URL, headers=headers, data=data)


# Slackにメッセージとファイルを投稿する関数
def slack_fileupload(token, channel, file, comment=""):
    """
    :param token:ボットトークンを指定
    :param channel:チャンネル名指定
    :param file:アップしたいファイルパスを指定
    :param comment:投稿したい内容を記載
    :return:
    """
    # ファイルの名前をかえられます
    file_name = "my file"

    url = "https://slack.com/api/files.upload"

    data = {
       "token": token,
       "channels": channel,
       # "title": file_name ,
       "initial_comment": comment
    }

    files = {'file': open(file, 'rb')}

    # 投稿用
    requests.post(url, data=data, files=files)
