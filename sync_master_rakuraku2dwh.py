#--------------------------------------------------------------------------------------------------------------------------
# 事前：pip install pymssql --trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org
#--------------------------------------------------------------------------------------------------------------------------
import os
import requests
import json
import pymssql
import csv


def exec_request(body):
    url = 'https://hnleda.rakurakuhanbai.jp/j6a7kma/api/csvexport/version/v1'
    headers = {
        'Content-Type': 'application/json; charset=utf-8',
        'X-HD-apitoken': 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'
    }
    response = requests.post(url, headers=headers, json=body)
    # print('status:', response.status_code)
    # print(response.text)
    return response.text

def import_csv_to_db(target_table, csv_path, column_count):
    # 設定値
    SERVER = 'XXXX'
    DATABASE = 'XXXX'
    USER = 'XXXX'
    PASSWORD = 'XXXX'
    TABLE = target_table
    CSV_PATH = csv_path
    
    # 接続
    conn = pymssql.connect(
        server=SERVER,
        user=USER,
        password=PASSWORD,
        database=DATABASE,
        charset='UTF-8',  # 文字コードに応じて
    )
    cursor = conn.cursor()
    
    # TRUNCATE
    cursor.execute(f'TRUNCATE TABLE {TABLE}')
    conn.commit()
    
    # CSV取り込み
    with open(CSV_PATH, mode='r', encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        rows = list(reader)  # ヘッダーなし想定
    
    # 列数に応じて'%s'の数を調整（pymssql は %s プレースホルダ）
    placeholders = ', '.join(['%s'] * column_count)
    insert_sql = f'INSERT INTO {TABLE} VALUES ({placeholders})'
    
    cursor.executemany(insert_sql, rows)
    conn.commit()
    
    cursor.close()
    conn.close()


if __name__ == '__main__':
    
    export_dir = os.path.dirname(os.path.abspath(__file__))
    
    #------------------------------
    # 最新レート
    #------------------------------
    # CSV出力
    export_path = os.path.join(export_dir, '最新レート.csv')
    if os.path.exists(export_path):
        os.remove(export_path)
    
    with open(export_path, mode='a', encoding='utf-8') as f:
        # API実行
        body = {
            'dbSchemaId': '101258',
            'listId': '101162',
            'limit': '200',
            'offset': '0'
        }        
        ret = exec_request(body)
        print(ret)
        # 1行目（ヘッダー行）を除いてファイル書き込み
        for i, line in enumerate(ret.splitlines()):
            if i == 0:
                continue
            f.write(f'{line}\n')
    # DBインポート
    target_table = '[SAPIFP].[SAPIFP].[TB_RAKUHAN_EXPORT_LATEST_RATE]'
    column_count = 5
    
    import_csv_to_db(target_table, export_path, column_count)
    
    #------------------------------
    # 社員
    #------------------------------
    # CSV出力
    export_path = os.path.join(export_dir, '社員.csv')
    if os.path.exists(export_path):
        os.remove(export_path)
    
    with open(export_path, mode='a', encoding='utf-8') as f:
        
        # API実行 ※1回のlimitが200レコードのためそれ毎に実行
        for offset in (0, 200, 400, 600, 800):
            body = {
                'dbSchemaId': '101252',
                'searchId': '103967',
                'listId': '101161',
                'limit': '200',
                'offset': f'{offset}'
            }        
            ret = exec_request(body)
            print(ret)
            # 1行目（ヘッダー行）を除いてファイル書き込み
            for i, line in enumerate(ret.splitlines()):
                if i == 0:
                    continue
                f.write(f'{line}\n')
    # DBインポート
    target_table = '[SAPIFP].[SAPIFP].[TB_RAKUHAN_EXPORT_SHAIN]'
    column_count = 8
    
    import_csv_to_db(target_table, export_path, column_count)
