# reports/db_connection.py
import os
import pyodbc
from dotenv import load_dotenv

# загружаем .env
load_dotenv()

def get_mssql_connection():
    conn_str = (
        'DRIVER={%s};'
        'SERVER=%s;'
        'DATABASE=%s;'
        'UID=%s;'
        'PWD=%s;'
    ) % (
        os.getenv('DB_DRIVER'),
        os.getenv('DB_SERVER'),
        os.getenv('DB_NAME'),
        os.getenv('DB_USER'),
        os.getenv('DB_PASSWORD'),
    )

    return pyodbc.connect(conn_str)
