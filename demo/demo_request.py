###### tags: `OCR_DEMO`
# 1-1リクエストプログラム
""" ↓モジュール """
import requests
import glob
# 文字列を辞書型に変更するモジュール
import ast
# PDFの１ファイルに何ページあるか読み取るモジュール
import PyPDF2
# ログの吐きだしに使用するモジュール
import logging
# テキストファイルをcsvに変換させるためpandasをインポート
import pandas as pd
# 年数を直打ちでなく変数で取得するためのモジュール
import datetime
# # 時間を計測するモジュール
import time
# プログラムに何かあった場合終了させるモジュール
import sys
# エラー内容を取得するためのモジュール
import traceback
# メモリを解放するモジュール
import gc


# Slackへ通知を行うモジュール
# from　slack_post

# import demo_slack_post

""" ↑モジュール """

# ログの吐きだしに現在日時を使用するので、定義しておく
# 現在時刻を取得
now = datetime.datetime.now()
current_time = now.strftime('%Y-%m-%d-%H-%M-%S')

""" ↓logの設定 """
"""
関数名	    用法
debug( )	問題がないか診断する場合
info( )	    思い通りに動作する場合
warning( )	想定外の事が発生しそうな場合
error( )	重大なエラーにより、一部の機能を実行できない場合
critical( )	プログラム自体が実行を続けられないことを表す、致命的な不具合の場合
"""


# DXSLへリクエストを行う関数
def D_request():
    # 指定のフォルダから全てのPDFファイルのパスをリストに格納させる
    files_list = glob.glob("post_pdf/*.pdf")
    # 中身を確認
    print(files_list)

    # 一枚当たりの処理時間を定義
    # この値は実際に計測してみて最大読み取り項目数の書式よりにしてます。
    ave_time = 18

    # 使いたい機能のエンドポイントが記されているURI（URL）を指定する（仕様書参照）
    # これは仕様書で決まっている為変更不可
    # 継続して処理を行わせたい場合はuriの末尾からクエリストリングの形式で記入する。
    # 「？」を先頭に「名前＝値」の形になるように記載をする
    # 処理パラメータとして設定する値が複数ある場合には「＆」で区切る
    # 仕様書通りに設定しないと動作しない。真は「true」、「True」ではない、偽も同様に「false」でなければならない
    # 月のリクエスト数を部署毎に管理を行いたいのでユーザーの設定を行う。
    uri = 'https://kki-1105.dx-suite.com/Sorter/api/v1/add?runSortingFlag=true&sendOcrFlag=true&userId=80144'

    # これはブラウザ上で設定を行ったものをコピペする長すぎて扱いずらいので変数に代入している
    api_key = "29ade314c7cfbf28aa16e0b91dc9cc007d90860ea9a9ed43019d7deeddc30af2465be8c4ed52a8ad89016e5122c2a725c95e424fa4ed3759c3d0f4142f2fb9c5"

    # ここにAPIキーを指定する
    headers = {
        'X-ConsoleWeb-ApiKey': api_key,
    }

    """ これで一度に複数ファイルのリクエストが可能に(ファイルのリクエストはリスト型で行う) """
    files = [('sorterRuleId', (None, "36466"))]

    # OCRの待ち時間を算出するのに1ファイルに何ページ格納されているかしる必要があるので
    # その値を格納させるリストを容易
    pages_list = []
    try:
        for i in files_list:
            files.append(('file', open(i, 'rb')))
            # PDFの読み込みを定義引数にファイルパスが格納されている変数iを指定
            reader = PyPDF2.PdfFileReader(i)
            # getNumPagesメソッドでPDFファイル内のページ数を出す。引数はなし
            num_pages = reader.getNumPages()
            # ページ数確認の為に出力
            # print(num_pages)
            # for文の外で定義したpages_listという空リストに取得したページ数を追加していく
            pages_list.append(num_pages)
            # 確認の為出力
            # print(pages_list)

            # ループ後のリストの中身を確認
            # print(files)

    # エラーが出た場合はエラー内容をプリントする
    except Exception as e:
        traceback.print_exc()
        # ログにエラーの情報を書かせる
        logging.exception("取得したエラーは下記の内容です。")
        # エラー内容を通知
        # slack_post("PDFが指定フォルダにないか正常に読み取れませんでした")
        # プログラムを終了する
        sys.exit()

    # 関数sumでリストの値を合計する
    pages_list_sum = sum(pages_list)
    # 確認の為出力
    # print(pages_list_sum)

    try:
        # ()内の引数は第一引数がエンドポイント、第二引数がAPIキー、第三引数がアップロードさせるファイルパスと仕分けルールのIDが入る
        res = requests.post(uri, headers=headers, files=files)

        # apiのレスポンスをテキスト化する
        result = res.text

        print(result)

        # messageにはslackに送信する文面を格納する
        # レスポンスに200が含まれていればOK
        if "success" in result:
            message = "リクエストが正常に行われました。"
            sort_json = res.json()
            sortid = str(sort_json['sortingUnitId'])
            print(message)
        # 認証エラーユーザーIDが間違っているかログインできないユーザーが指定されている場合い発生
        elif "101" in result:
            message = "errorCode：101\n認証エラーです。管理者に確認してください"
        # 指定フォルダにPDFが格納されていない場合に発生
        elif "103" and "file" in result:
            message = "errorCode：103\nPDFが指定フォルダに入っているか確認し再度プログラムを実行してください"
        # 上記２件以外のエラーは想定外のエラーなのでレスポンスの内容をそのままSlackに送る
        else:
            message = res.text
            print(message)

    # エラーが出た場合はエラー内容をプリントする
    except Exception as e:
        traceback.print_exc()
        # ログにエラーの情報を書かせる
        logging.exception("取得したエラーは下記の内容です。")
        # エラー内容を通知
        # slack_post("仕分けのリクエストが正常に行えませんでした")
        # プログラムを終了する
        sys.exit()

    # メモリ解放の為削除
    del uri
    gc.collect()


    # logを吐きだす
    logging.info('リクエスト完了')
    return message, sortid, headers


if __name__ == '__main__':
    D_request()