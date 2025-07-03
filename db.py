import os
import mariadb
from dotenv import load_dotenv

# Load env
load_dotenv()

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "port": int(os.getenv("DB_PORT", 3306)),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASSWORD"),
    "database": os.getenv("DB_NAME"),
}

def get_connection(db_override=None):
    try:
        config = DB_CONFIG.copy()
        if db_override:
            config["database"] = db_override
        conn = mariadb.connect(**config)
        return conn
    except mariadb.Error as e:
        print(f"❌ Gagal koneksi ke database: {e}")
        return None
