import os
import datetime
import traceback

# 本体のquote_mail_relay.py の変数/関数を読み込み
import quote_mail_relay as qmr

# ==============================
# 設定値
# ==============================
# ログファイル
PROCESSING_FILE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE_NAME = os.path.basename(__file__).replace('.py', '.log')
LOG_FILE_PATH = os.path.join(PROCESSING_FILE_DIR, LOG_FILE_NAME)

# 滞留メール数の閾値（この数を超過した場合、警告メール送信）
MAIL_COUNT_THRESHOLD = -1


if __name__ == '__main__':
  access_token = None
  qmr.print_log('INFO', '処理を開始します', LOG_FILE_PATH)
  try:
    # アクセストークンの取得
    access_token = qmr.get_access_token()
    qmr.print_log('INFO', 'アクセストークンの取得が成功',LOG_FILE_PATH)

    # 対象メールの取得
    target_mails = qmr.fetch_target_mails(access_token)
    stock_mail_count = len(target_mails)
    qmr.print_log('INFO', f'現在の滞留メール数: {stock_mail_count} 件', LOG_FILE_PATH)

    # 閾値未満判定
    if stock_mail_count > MAIL_COUNT_THRESHOLD:
      qmr.print_log('WARN', f'滞留メール数が閾値（{MAIL_COUNT_THRESHOLD}）を超過しています！！！！！！！', LOG_FILE_PATH)
      
      # メール通知
      subject = f'【WARN】楽楽販売の見積メール中継処理にてメール滞留が発生しています'
      body = (
        f'実行時刻：{datetime.datetime.now()}\n\n'
        f'現在のメール滞留数：{stock_mail_count} 件\n\n'
        f'閾値：{MAIL_COUNT_THRESHOLD} 件\n'
      )
      qmr.send_email_graph(
        access_token=access_token,
        sender_email=qmr.ERROR_MAIL_FROM,
        recipient_to=[qmr.ERROR_MAIL_TO],
        subject=subject,
        body_content=body
      )
      qmr.print_log('INFO', 'メール滞留の通知メールを送信しました', LOG_FILE_PATH)
    else:
      qmr.print_log('INFO', f'正常稼働です（メール滞留は発生していません）', LOG_FILE_PATH)
  except Exception as e:
    error_detail = f'処理中にエラーが発生しました: {e}\n{traceback.format_exc()}'
    qmr.print_log('ERROR', error_detail, LOG_FILE_PATH)
    
    if access_token:  # トークンが取得できていればエラー通知を試みる
      qmr.send_error_notification(access_token, error_detail)
