# ------------------------------
# 事前設定
#   pip install msal --trusted-host pypi.org --trusted-host files.pythonhosted.org
#   pip install requests --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
# ------------------------------
import os
import re
import base64
from msal import ConfidentialClientApplication
import requests
import json
import traceback
import shutil
import stat
import datetime


# ==============================
# 設定値
# ==============================
# 多重起動防止用ファイル（スクリプトと同一ディレクトリに配置）
PROCESSING_FILE_DIR = os.path.dirname(os.path.abspath(__file__))
PROCESSING_FILE_NAME = '.Processing'

# ログファイル
LOG_FILE_NAME = os.path.basename(__file__).replace('.py', '.log')

# Microsoft Entra ID (Azure AD) アプリケーション登録情報
TENANT_ID = 'XXXXXXXXXX'
CLIENT_ID = 'XXXXXXXXXX'
CLIENT_SECRET = 'XXXXXXXXXX'

# 監視対象メールアドレス (受信トレイを監視)
MONITOR_EMAIL = 'system-rakurakuhanbai@hakuto.co.jp'
# 監視対象メールの差出人
TARGET_SENDER_FOR_MONITOR = 'system@rakurakuhanbai.jp'
# 監視対象メールの件名フィルタ用キーワード (前方一致)
TARGET_KEYWORD = '[見積送付メール]'

# 処理済みメールを移動するOutlookフォルダ名
PROCESSED_FOLDER = 'Processed'
# ダウンロードした添付ファイルの一時保存ディレクトリ
TEMP_ATTACHMENT_DIR = 'temp_attachments'

# エラー通知メールの送信元
ERROR_MAIL_FROM = 'system-rakurakuhanbai@hakuto.co.jp'
# エラー通知メールの送信先
ERROR_MAIL_TO = 'suzuki-hirokazu@hakuto.co.jp'

# Microsoft Graph APIのエンドポイント
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'


# ==============================
# Logger
# ==============================
def print_log(level, msg):
  now = datetime.datetime.now()
  ts = now.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
  print(f'[{ts}][{os.getpid()}][{level}] {msg}')
  
  log_file_path = os.path.join(PROCESSING_FILE_DIR, LOG_FILE_NAME)
  with open(log_file_path, 'a') as f:
    f.write(f'[{ts}][{level}] {msg}\n')

# ==============================
# 多重起動防止処理
# ==============================
def check_and_create_processing_file():
  processing_file_path = os.path.join(PROCESSING_FILE_DIR, PROCESSING_FILE_NAME)
  if os.path.exists(processing_file_path):
    print_log('WARN', 'ロックファイルが存在している（他プロセスで処理が実行中）ため本処理を終了します')
    return False
  else:
    with open(processing_file_path, 'w') as f:
      f.write('') # 0バイトファイルを作成
    print_log('INFO', 'ロックファイルを作成しました')
    return True

def delete_processing_file():
  processing_file_path = os.path.join(PROCESSING_FILE_DIR, PROCESSING_FILE_NAME)
  if os.path.exists(processing_file_path):
    os.remove(processing_file_path)
    print_log('INFO', 'ロックファイルを削除しました')

# ==============================
# Microsoft Graph API のアクセストークン取得 (OAuth2)
# ==============================
def get_access_token():
  authority = f'https://login.microsoftonline.com/{TENANT_ID}'
  app = ConfidentialClientApplication(
    CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
  )
  # Mail.Read, Mail.Send, MailboxSettings.Read (フォルダ作成のため) 権限が含まれる.defaultスコープを指定
  scopes = ['https://graph.microsoft.com/.default'] 
  
  result = app.acquire_token_for_client(scopes=scopes)
  if 'access_token' not in result:
    raise Exception(f'アクセストークンの取得に失敗しました:\n {result.get("error_description", result)}')
  return result['access_token']

# ==============================
# Graph API を使用して対象メールを取得
# ==============================
def fetch_target_mails(access_token):
  headers = {'Authorization': f'Bearer {access_token}'}
  url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/mailFolders/Inbox/messages'
  params = {
    '$orderby': 'receivedDateTime asc',  # 受信日時の古いものから順
    '$top': 20                           # 最大20件を取得
  }
  response = requests.get(url, headers=headers, params=params)
  
  response.raise_for_status() # HTTPエラーが発生した場合は例外を発生させる
  data = response.json()  
  mails = data.get('value', []) 
  
  # 差出人と件名のチェック
  filtered = []
  for m in mails:
    addr = m['from']['emailAddress']['address']
    subj = m['subject'] or ''
    if addr == TARGET_SENDER_FOR_MONITOR and subj.startswith(TARGET_KEYWORD):
      filtered.append(m)  
  
  if filtered:
    print_log('INFO', f'対象メールは[{len(filtered)}]件です')
  else:
    print_log('WARN', '対象メールはありません')
  
  return filtered

# ==============================
# メール本文の1〜4行目からFrom, To, Cc, Bccアドレスを抽出し、それらを除去した残りの本文を返却
# ==============================
def parse_mail_body(body_content):
  lines = body_content.splitlines()
  
  sender_from = ''
  recipients_to = []
  recipients_cc = []
  recipients_bcc = []
  new_body_lines = []

  # 各行を解析
  for i, line in enumerate(lines):
    if i == 0 and line.startswith('From:'):
      sender_from = line[len('From:'):].strip()
    elif i == 1 and line.startswith('To:'):
      # ;で分割し、各アドレスをstripして空でないもののみをリストに入れる
      recipients_to = [addr.strip() for addr in line[len('To:'):].split(';') if addr.strip() and '@' in addr]
    elif i == 2 and line.startswith('Cc:'):
      recipients_cc = [addr.strip() for addr in line[len('Cc:'):].split(';') if addr.strip() and '@' in addr]
    elif i == 3 and line.startswith('Bcc:'):
      recipients_bcc = [addr.strip() for addr in line[len('Bcc:'):].split(';') if addr.strip() and '@' in addr]
    else:
      new_body_lines.append(line)
  
  new_body = '\n'.join(new_body_lines).strip()
  return {
    'from': sender_from,
    'to': recipients_to,
    'cc': recipients_cc,
    'bcc': recipients_bcc,
    'body': new_body
  }

# ==============================
# 指定されたメールIDの添付ファイルを一時ディレクトリにダウンロード
# ==============================
def download_attachments(access_token, mail_id, download_dir):
  headers = {'Authorization': f'Bearer {access_token}'}
  
  # 添付ファイルの一時保存ディレクトリを作成
  os.makedirs(download_dir, exist_ok=True)
  
  attachments_url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/messages/{mail_id}/attachments'
  print_log('INFO', '添付ファイルをダウンロードします')
  response = requests.get(attachments_url, headers=headers)
  response.raise_for_status()
  attachments_data = response.json()
  
  downloaded_paths = []
  if 'value' in attachments_data:
    for attachment in attachments_data['value']:
      # isInlineがFalse (インライン添付ではない) のものをダウンロード対象とする
      if not attachment.get('isInline', False):
        file_name = attachment['name']
        # contentBytesはBase64エンコードされているのでデコードする
        file_content = base64.b64decode(attachment['contentBytes'])
        file_path = os.path.join(download_dir, file_name)
        
        with open(file_path, 'wb') as f:
          f.write(file_content)
        print_log('INFO', f'添付ファイルのダウンロードが完了しました: {file_name}')
        downloaded_paths.append(file_path)
  
  if not downloaded_paths:
    print_log('WARN', '添付ファイルは見つかりません...')
    
  return downloaded_paths

# ==============================
# Microsoft Graph API でメールを再発射
# ==============================
def send_email_graph(access_token, sender_email, recipient_to, subject, body_content, recipient_cc=None, recipient_bcc=None, attachments=None):
  headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
  }
  message = {
    'subject': subject,
    'body': {
      'contentType': 'Text', # テキスト形式の本文
      'content': body_content
    },
    'toRecipients': [{'emailAddress': {'address': addr}} for addr in recipient_to]
  }

  # CCとBCCの追加 (リストが空でない場合のみ)
  if recipient_cc:
    message['ccRecipients'] = [{'emailAddress': {'address': addr}} for addr in recipient_cc]
  if recipient_bcc:
    message['bccRecipients'] = [{'emailAddress': {'address': addr}} for addr in recipient_bcc]

  # 添付ファイルの追加
  if attachments:
    message['attachments'] = []
    for file_path in attachments:
      with open(file_path, 'rb') as f:
        file_content_bytes = f.read()
        # ファイル内容をBase64エンコードしてJSONに含める
        encoded_content = base64.b64encode(file_content_bytes).decode('utf-8')
      
      attachment_name = os.path.basename(file_path)
      message['attachments'].append({
        '@odata.type': '#microsoft.graph.fileAttachment', # ファイル添付のタイプを指定
        'name': attachment_name,
        'contentType': 'application/octet-stream', # 汎用的なMIMEタイプを使用
        'contentBytes': encoded_content
      })

  payload = {
    'message': message,
    'saveToSentItems': 'true' # 送信済みアイテムに保存する
  }
  send_mail_url = f'{GRAPH_ENDPOINT}/users/{sender_email}/sendMail'
  response = requests.post(send_mail_url, headers=headers, data=json.dumps(payload))
  response.raise_for_status() # HTTPエラーが発生した場合は例外を発生させる

# ==============================
# メールを処理済みフォルダに移動
# ==============================
def move_mail_to_processed_folder(access_token, mail_id):
  headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}

  # まずProcessedフォルダのIDを取得
  processed_folder_id = None
  mail_folders_url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/mailFolders'
  response = requests.get(mail_folders_url, headers=headers)
  response.raise_for_status()
  folders = response.json().get('value', [])

  for folder in folders:
    if folder['displayName'] == PROCESSED_FOLDER:
      processed_folder_id = folder['id']
      break
  
  # Processedフォルダが存在しない場合は作成する
  if not processed_folder_id:
    create_folder_url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/mailFolders'
    create_folder_payload = {'displayName': PROCESSED_FOLDER}
    response = requests.post(create_folder_url, headers=headers, data=json.dumps(create_folder_payload))
    response.raise_for_status()
    processed_folder_id = response.json()['id']
    print_log('INFO', f'[{PROCESSED_FOLDER}]フォルダ が存在しないため作成しました')

  # メールを移動する
  move_mail_url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/messages/{mail_id}/move'
  move_payload = {'destinationId': processed_folder_id}
  response = requests.post(move_mail_url, headers=headers, data=json.dumps(move_payload))
  response.raise_for_status()
  print_log('INFO', f'中継処理完了済みメールを[{PROCESSED_FOLDER}]フォルダに移動しました')

# ==============================
# エラー通知メール送信
# ==============================
def send_error_notification(access_token, error_detail):
  subject = '【エラー】楽楽販売の見積メール中継処理にてエラーが発生しました'
  body = f'エラー内容：\n\n{error_detail}'
  
  send_email_graph(
    access_token=access_token,
    sender_email=ERROR_MAIL_FROM,
    recipient_to=[ERROR_MAIL_TO],
    subject=subject,
    body_content=body
  )
  print_log('INFO', f'エラー通知メールを[{ERROR_MAIL_TO}]宛に送信しました')

# ==============================
# 一時フォルダ削除が失敗したとき用（まず属性を「書き込み可」にしてから再実行）
# ==============================
def on_rm_error(func, path, exc_info):
  # func: shutil.rmtree など、path: 削除しようとして失敗したパス、exc_info: sys.exc_info() のタプル
  os.chmod(path, stat.S_IWRITE)
  func(path)



if __name__ == '__main__':
  access_token = None
  print_log('INFO', '処理を開始します')
  try:
    # 多重起動防止（既にファイルが存在する場合は処理を終了）
    if not check_and_create_processing_file():
      os._exit(0)  # os._exitだとfinallyは動かない

    # アクセストークンの取得
    access_token = get_access_token()
    print_log('INFO', 'アクセストークンの取得が成功')

    # 対象メールの取得
    target_mails = fetch_target_mails(access_token)

    if not target_mails:
      print_log('WARN', '対象メールがないため処理を終了します')
      exit()
    else:
      # 添付ファイル一時保存ディレクトリをクリーンアップ
      if os.path.exists(TEMP_ATTACHMENT_DIR):
        shutil.rmtree(TEMP_ATTACHMENT_DIR, onerror=on_rm_error)
      os.makedirs(TEMP_ATTACHMENT_DIR, exist_ok=True) 

      for i, mail in enumerate(target_mails, start=1):
        mail_id = mail['id']
        original_subject = mail['subject']
        # Graph APIのbodyコンテンツはHTML形式の場合があるので、textコンテンツを使用
        original_body = mail['body']['content']
        print_log('INFO', f'{i}通目のメールを処理します')

        # メール本文から必要情報を抽出
        parsed_info = parse_mail_body(original_body)
        extracted_from = parsed_info['from']
        extracted_to = parsed_info['to']
        extracted_cc = parsed_info['cc']
        extracted_bcc = parsed_info['bcc']
        new_subject = original_subject.replace(TARGET_KEYWORD, '', 1).strip()  # 件名からキーワードを削除
        new_body_content = parsed_info['body']
        print_log('INFO', f'抽出したメール情報:\n件名：{new_subject}\nFrom: {extracted_from}\nTo: {extracted_to}\nCc: {extracted_cc}\nBcc: {extracted_bcc}')

        #  添付ファイルをダウンロード
        downloaded_attachments = download_attachments(access_token, mail_id, TEMP_ATTACHMENT_DIR)

        # メールを再発射
        send_email_graph(
          access_token=access_token,
          sender_email=extracted_from,
          recipient_to=extracted_to,
          subject=new_subject,
          body_content=new_body_content,
          recipient_cc=extracted_cc,
          recipient_bcc=extracted_bcc,
          attachments=downloaded_attachments
        )
        print_log('INFO', 'メール中継（再送信）が成功しました!')

        # このメールに紐づく添付ファイルの一時ファイルをクリーンアップ
        for file_path in downloaded_attachments:
          if os.path.exists(file_path):
            os.remove(file_path)
            print_log('INFO', '添付ファイル一時保存ファイルを削除しました')
        
        # 元メールを処理済みフォルダに移動
        move_mail_to_processed_folder(access_token, mail_id)
        
        print_log('INFO', ''.join((f'{i}通目のメール中継処理が完了しました\n', '-'*40)))

    print_log('INFO', '全てのメール中継処理が完了しました')

  except Exception as e:
    error_detail = f'処理中にエラーが発生しました: {e}\n{traceback.format_exc()}'
    print_log('ERROR', error_detail)
    if access_token:  # トークンが取得できていればエラー通知を試みる
      send_error_notification(access_token, error_detail)
  finally:
    # 添付ファイル一時保存ディレクトリ全体を削除
    if os.path.exists(TEMP_ATTACHMENT_DIR):
      shutil.rmtree(TEMP_ATTACHMENT_DIR, onerror=on_rm_error)
      print_log('INFO', '添付ファイル一時保存ディレクトリを削除しました')
    
    # 多重起動防止ファイルを削除
    delete_processing_file()
    print_log('INFO', '処理を終了します')
