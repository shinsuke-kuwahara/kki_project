import os
from oauth2client.service_account import ServiceAccountCredentials
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


# googledriveにファイルをUploadする関数
def fileupload(json, folder, filepath):
    """
    :param json:認証用トークン
    :param folder: 保存先のフォルダパス
    :param filepath: 参照用のファイルパス
    """
    # ファイル作成用に日付を取得
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive',
             'https://www.googleapis.com/auth/spreadsheets']

    # サービスアカウントキーを読み込む
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json, scope)
    # pydrive用にOAuth認証を行う
    gauth = GoogleAuth()
    gauth.credentials = credentials
    drive = GoogleDrive(gauth)

    # フォルダ内のファイルをドライブへ保存
    for files in filepath:
        # 保存用の名前を取得
        file_name = os.path.basename(files)
        # 指定フォルダにファイルをアップロード
        f = drive.CreateFile({"title": file_name,
                              'parents': [{"id": folder}]})
        f.SetContentFile(files)
        f.Upload()

