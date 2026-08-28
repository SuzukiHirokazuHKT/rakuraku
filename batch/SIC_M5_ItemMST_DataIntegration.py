#--------------------------------------------------------------------------------------------------------------------------
# 事前：pip install pyodbc --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
#--------------------------------------------------------------------------------------------------------------------------

import pyodbc
import sys
import os
from datetime import datetime as dt
import csv
import io
import requests


# ログ設定
SCRIPT_PATH = os.path.abspath(sys.argv[0])
LOG_FILE = os.path.splitext(SCRIPT_PATH)[0] + '.log'
MAX_LINES = 5000

def log_and_print(msg, error_level='INFO'):

    now = dt.now()
    dt_str = now.strftime('%Y-%m-%d %H:%M:%S.%f')
    
    log_line = f"[{dt_str}][{error_level}] {msg}"
    
    print(log_line)
    
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(log_line + '\n')
        
    # 行数が1000行を超えた場合、古い行を削って上書きする
    try:
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            
        if len(lines) > MAX_LINES:
            with open(LOG_FILE, 'w', encoding='utf-8') as f:
                # 後ろから MAX_LINES 分だけ残して上書き
                f.writelines(lines[-MAX_LINES:])
    except IOError as e:
        log_and_print(f"[Log Error] ログファイルのローテーションに失敗しました: {e}", 'ERROR')


class SQLServerClient:
    # 初期化時にデータベースへの接続を確立しself.connに保持
    def __init__(self):
        
        DRIVER = '{SQL Server}'
        SERVER = 'sh72019'
        DATABASE = 'PartsList'
        USERNAME = 'XXXXXXXXXX'
        PASSWORD = 'XXXXXXXXXX'
        
        conn_str = f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
        
        try:
            self.conn = pyodbc.connect(conn_str)
        except pyodbc.Error as e:
            log_and_print(f'データベース接続エラー: {e}', 'ERROR')
            raise
            
    # 該当テーブルの件数を所得する
    def get_count(self, table_name) -> int:
        query = f"SELECT COUNT(*) FROM {table_name}"
        
        with self.conn.cursor() as cursor:
            try:
                cursor.execute(query)
                result = cursor.fetchone()
                # fetchone() はタプルを返すので、最初の要素(インデックス0)を取得します
                return result[0] if result else 0
            except pyodbc.Error as e:
                log_and_print(f"レコード数取得エラー (テーブル: {table_name}): {e}")
                raise
    
    # SELECTクエリを実行し、データ部のみをCSV形式の文字列として返却する
    def select_data(self, query) -> str:
        
        with self.conn.cursor() as cursor:
            cursor.execute(query)
            rows = cursor.fetchall()

            output = io.StringIO()
            writer = csv.writer(output, lineterminator='\n', quoting=csv.QUOTE_ALL)

            for row in rows:
                writer.writerow(row)

            return output.getvalue()

    # インスタンスが破棄される際にコネクションをクローズする
    def __del__(self):
        if hasattr(self, 'conn'):
            self.conn.close()

# 楽々販売のCSVインポートAPI実行する関数
def exec_csv_import_api(csv_data):

    URL = "https://hnleda.rakurakuhanbai.jp/j6a7kma/api/csvdataimport/version/v1"
    API_TOKEN = 'XXXXXXXXXX'
    DB_SCHEMA_ID = 'XXXXXXXXXX'
    IMPORT_ID = 'XXXXXXXXXX'
    
    headers = {
        'X-HD-apitoken': API_TOKEN,
        'Content-Type': 'multipart/form-data; boundary=boundary;' 
    }
    
    body = '\n'.join([
        '--boundary',
        'Content-Disposition: form-data; name="json"',
        'Content-Type: application/json',
        '',
        '{{',
        '  "dbSchemaId": "{}",',
        '  "importId": "{}"',
        '}}',
        '--boundary',
        'Content-Disposition: form-data; name="uploadFile"; filename="upload.csv"',
        'Content-Type: text/csv',
        '',
        '{}',
        '--boundary--'
    ]).format(
        DB_SCHEMA_ID,
        IMPORT_ID,
        csv_data
    )
    log_and_print(f'body:\n{body}')
    
    try:
        response = requests.post(URL, headers=headers, data=body.encode('utf-8'))
        
        response.raise_for_status()  # HTTPステータスコードがエラー（4xx, 5xx）の場合は例外を発生させる
        
        log_and_print(f'API実行成功: ステータスコード {response.status_code}')
        log_and_print('レスポンス内容:', response.text)
        
        return response.text
    except requests.exceptions.RequestException as e:
        log_and_print(f'API実行時にエラーが発生しました: {e}', 'ERROR')
        if hasattr(e, 'response') and e.response is not None:
             log_and_print('エラー詳細:', e.response.text, 'ERROR')
        return None


if __name__ == '__main__':
    try:
        conn = SQLServerClient()
        
        TARGET_TABLE = '[VW_RAKUHAN_API_SIC-M5_PartsList]'
        
        cnt = conn.get_count(TARGET_TABLE)
        
        log_and_print(f'{TARGET_TABLE}： {cnt}件')
        
        if cnt < 1:
            log_and_print('連携対象がないため処理を終了します')
            sys.exit(0)
        
        sql = f'select top(100) * from {TARGET_TABLE}'
                
        csv_data = conn.select_data(sql)
                
        # 楽々販売のCSVインポートAPIを実行する
        exec_csv_import_api(csv_data)
        
    except Exception as e:
        log_and_print(f'処理中にエラーが発生しました: {e}', 'ERROR')

