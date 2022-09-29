###### tags: `OCR_DEMO`
# 7-1demo_slack_post
# Slackへ通知を行う関数
import requests
# 変数mesにはSlackへ送信したいメッセージをstr型で指定してください
def slack_post(mes):
    # ワークスペース柄澤備忘録(仮)のbot情報
    bot_token = 'xoxb-1372090409680-3526226890113-jaWNBQeuh9kXSLa5ylbqTRXq'
    URL = 'https://slack.com/api/chat.postMessage'
    headers = {"Authorization": "Bearer " + bot_token}
    # チャンネルはbot_message_post
    channel = "C040B44601K"
    # 指定のチャンネルに使用中を通知
    data = {
        'channel': channel,
        'text': mes,
        'as_user': True
    }

    requests.post(URL, headers=headers, data=data)

