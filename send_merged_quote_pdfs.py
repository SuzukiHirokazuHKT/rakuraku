# ------------------------------
# 事前設定
#   pip install msal requests PyPDF2 --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
# ------------------------------
import os
import shutil
import json
import base64
import traceback

# 本体のquote_mail_relay.py の変数/関数を読み込み
import quote_mail_relay as qmr

from PyPDF2 import PdfMerger


# ==============================
# 設定値
# ==============================
# 多重起動防止用ファイル（スクリプトと同一ディレクトリに配置）
PROCESSING_FILE_DIR = os.path.dirname(os.path.abspath(__file__))
PROCESSING_PATH = os.path.join(PROCESSING_FILE_DIR, '.Processing_Pdf-merge')

# ログは既存のprint_logを利用
print_log = qmr.print_log

# ログファイル
LOG_FILE_NAME = os.path.basename(__file__).replace('.py', '.log')
LOG_FILE_PATH = os.path.join(PROCESSING_FILE_DIR, LOG_FILE_NAME)

# 監視対象メールアドレス (受信トレイを監視)
MONITOR_EMAIL = 'system-rakurakuhanbai@hakuto.co.jp'
# 監視対象メールの差出人
TARGET_SENDER_FOR_MONITOR = 'system@rakurakuhanbai.jp'
# 監視対象メールの件名フィルタ用キーワード (前方一致)
TARGET_KEYWORD_FOR_MERGE = '[結合PDF送付メール]'

# 処理済みメールを移動するOutlookフォルダ名
PROCESSED_FOLDER = qmr.PROCESSED_FOLDER
# ダウンロードした添付ファイルの一時保存ディレクトリ
TEMP_ATTACHMENT_DIR = 'temp_pdf-merge_attachments'

# Microsoft Graph APIのエンドポイント
GRAPH_ENDPOINT = qmr.GRAPH_ENDPOINT


# ==============================
# Graph API を使用して対象メールを取得
# ==============================
def fetch_target_mails_for_merge(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/mailFolders/Inbox/messages'
    params = {
        '$orderby': 'receivedDateTime asc',
        '$top': 20
    }
    response = qmr.requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    data = response.json()
    mails = data.get('value', [])
    
    filtered = []
    for m in mails:
        addr = m['from']['emailAddress']['address']
        subj = m['subject'] or ''
        if addr == TARGET_SENDER_FOR_MONITOR and subj.startswith(TARGET_KEYWORD_FOR_MERGE):
            filtered.append(m)
            
    if filtered:
        print_log('INFO', f'対象メール（結合PDF）は[{len(filtered)}]件です', qmr.LOG_FILE_PATH)
    else:
        print_log('WARN', '対象メール（結合PDF）はありません', qmr.LOG_FILE_PATH)
        
    return filtered

# ==============================
# メール本文から必要情報をを抽出・整形して返却
# ==============================
def parse_merge_mail_body(body_content: str):
    lines = body_content.splitlines()
    # 初期値
    to_addresses = []
    to_name = ''
    cover_filename = ''
    quote_filename = ''
    merged_filename = ''
    
    for line in lines:
        if line.startswith('To:'):
            # ;区切りのメールアドレス群
            to_addresses = [addr.strip() for addr in line[len('To:'):].split(';') if addr.strip() and '@' in addr]
        elif line.startswith('To氏名:'):
            to_name = line[len('To氏名:'):].strip()
        elif line.startswith('表紙:'):
            cover_filename = line[len('表紙:'):].strip()
        elif line.startswith('見積書:'):
            quote_filename = line[len('見積書:'):].strip()
        elif line.startswith('結合pdf:'):
            merged_filename = line[len('結合pdf:'):].strip()
            
    return {
        'to_addresses': to_addresses,
        'to_name': to_name,
        'cover_filename': cover_filename,
        'quote_filename': quote_filename,
        'merged_filename': merged_filename
    }

# ==============================
# 指定されたメールIDの添付ファイルを一時ディレクトリにダウンロード
# ==============================
def download_specific_attachments(access_token, mail_id, download_dir, required_names: list[str]):
    """
    メールIDに紐づく添付ファイルのうち、required_namesに一致するファイル名のみダウンロード。
    """
    headers = {'Authorization': f'Bearer {access_token}'}
    os.makedirs(download_dir, exist_ok=True)
    
    attachments_url = f'{GRAPH_ENDPOINT}/users/{MONITOR_EMAIL}/messages/{mail_id}/attachments'
    print_log('INFO', '添付ファイルを確認します', qmr.LOG_FILE_PATH)
    response = qmr.requests.get(attachments_url, headers=headers)
    response.raise_for_status()
    attachments_data = response.json()
    
    downloaded_paths = []
    required_set = set(required_names)
    found_set = set()
    
    if 'value' in attachments_data:
        for attachment in attachments_data['value']:
            if not attachment.get('isInline', False):
                file_name = attachment['name']
                if file_name in required_set:
                    file_content = qmr.base64.b64decode(attachment['contentBytes'])
                    file_path = os.path.join(download_dir, file_name)
                    with open(file_path, 'wb') as f:
                        f.write(file_content)
                    print_log('INFO', f'添付ダウンロード完了: {file_name}', qmr.LOG_FILE_PATH)
                    downloaded_paths.append(file_path)
                    found_set.add(file_name)
                    
    missing = list(required_set - found_set)
    if missing:
        print_log('ERROR', f'必要な添付が不足: {missing}', qmr.LOG_FILE_PATH)
        
    return downloaded_paths, missing


# ==============================
# PDF結合処理
# ==============================
def merge_multiple_pdfs(pdf_paths: list[str], output_path: str) -> None:
    merger = PdfMerger()
    try:
        for p in pdf_paths:
            merger.append(p)
        with open(output_path, "wb") as f:
            merger.write(f)
    finally:
        merger.close()


# ==============================
# ユーザー送付メールの本文を生成
# ==============================
def build_mail_body_for_recipient(to_name: str) -> str:
    lines = (
        '※本メールは見積システム（楽楽販売）からの自動送信メールです※',
        '',
        f'{to_name} 様',
        '',
        '以下①②を結合したPDFを添付ファイルにてお送りいたします。',
        '',
        '①見積システムで作成した表紙pdf（押印済み）',
        '②見積システムにアップロードした見積書pdf',
    )
    # HTMLにする
    return '<br>'.join(lines)



if __name__ == '__main__':
    access_token = None
    print_log('INFO', '結合PDF処理を開始します', LOG_FILE_PATH)
    try:
        # 多重起動防止
        if not qmr.check_and_create_processing_file(PROCESSING_PATH):
            os._exit(0)
            
        # アクセストークン取得
        access_token = qmr.get_access_token()
        print_log('INFO', 'アクセストークンの取得が成功', LOG_FILE_PATH)
        
        # 対象メール取得（結合PDF用フィルタ）
        target_mails = fetch_target_mails_for_merge(access_token)
        if not target_mails:
            print_log('WARN', '対象メールがないため処理を終了します', LOG_FILE_PATH)
            exit()
            
        # 一時ディレクトリ準備
        if os.path.exists(TEMP_ATTACHMENT_DIR):
            shutil.rmtree(TEMP_ATTACHMENT_DIR, onerror=qmr.on_rm_error)
        os.makedirs(TEMP_ATTACHMENT_DIR, exist_ok=True)
        
        for i, mail in enumerate(target_mails, start=1):
            mail_id = mail['id']
            original_subject = mail['subject']
            original_body = mail['body']['content']  # 既存と同様
            
            print_log('INFO', f'{i}通目の結合PDFメールを処理します', LOG_FILE_PATH)
            
            # 本文パース
            parsed = parse_merge_mail_body(original_body)
            to_addresses = parsed['to_addresses']
            to_name = parsed['to_name']
            cover_filename = parsed['cover_filename']
            quote_filename = parsed['quote_filename']
            merged_filename = parsed['merged_filename']
            
            # 必須チェック
            if not to_addresses or not to_name or not cover_filename or not quote_filename or not merged_filename:
                print_log('ERROR', f'本文情報不足: {parsed}', LOG_FILE_PATH)
                continue
                
            # 添付ダウンロード（指定ファイルのみ）
            required_names = [cover_filename, quote_filename]
            downloaded_paths, missing = download_specific_attachments(
                access_token, mail_id, TEMP_ATTACHMENT_DIR, required_names
            )
            if missing:
                print_log('ERROR', f'必要添付不足によりスキップします: {missing}', LOG_FILE_PATH)
                continue
                
            # PDF結合（表紙→見積書の順）
            cover_path = [p for p in downloaded_paths if os.path.basename(p) == cover_filename][0]
            quote_path = [p for p in downloaded_paths if os.path.basename(p) == quote_filename][0]
            output_path = os.path.join(TEMP_ATTACHMENT_DIR, merged_filename)
            try:
                merge_multiple_pdfs([cover_path, quote_path], output_path)
                print_log('INFO', f'PDF結合完了: {output_path}', LOG_FILE_PATH)
            except Exception as e:
                print_log('ERROR', f'PDF結合に失敗: {e}', LOG_FILE_PATH)
                continue
                
            # 送信内容作成
            sender_email = MONITOR_EMAIL
            recipient_to = [to_addresses[0]]
            subject = '【見積システム】結合PDF（表紙＋見積書）をお送りします'
            body_html = build_mail_body_for_recipient(to_name)
            
            # 添付ファイルをBase64化して送信（既存send_email_graphを利用）
            try:
                qmr.send_email_graph(
                    access_token=access_token,
                    sender_email=sender_email,
                    recipient_to=recipient_to,
                    subject=subject,
                    body_content=body_html,
                    recipient_cc=None,
                    recipient_bcc=None,
                    attachments=[output_path],
                    contentType='HTML'
                )
                print_log('INFO', f'結合PDFメール送信成功: To={recipient_to[0]} 件名={subject}', LOG_FILE_PATH)
            except Exception as e:
                print_log('ERROR', f'メール送信失敗: {e}', LOG_FILE_PATH)
                continue
                
            # 後処理：元メールをProcessedへ移動
            try:
                qmr.move_mail_to_processed_folder(access_token, mail_id)
            except Exception as e:
                print_log('ERROR', f'Processedフォルダ移動失敗: {e}', LOG_FILE_PATH)
                
            # 終了ログ
            print_log('INFO', ''.join((f'{i}通目の結合PDF処理が完了しました\n', '-'*40)), LOG_FILE_PATH)
            
        print_log('INFO', '全ての結合PDF処理が完了しました', LOG_FILE_PATH)
        
    except Exception as e:
        error_detail = f'処理中にエラーが発生しました: {e}\n{traceback.format_exc()}'
        print_log('ERROR', error_detail, LOG_FILE_PATH)
        if access_token:
            # 既存のエラー通知機能を利用
            qmr.send_error_notification(access_token, error_detail)
    finally:
        # 一時ディレクトリ削除
        if os.path.exists(TEMP_ATTACHMENT_DIR):
            shutil.rmtree(TEMP_ATTACHMENT_DIR, onerror=qmr.on_rm_error)
            print_log('INFO', '結合PDF用の一時ディレクトリを削除しました', LOG_FILE_PATH)
        
        # ロック解除
        qmr.delete_processing_file(PROCESSING_PATH)
        print_log('INFO', '結合PDF処理を終了します', LOG_FILE_PATH)
