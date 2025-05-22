import os
import datetime
import subprocess
from dotenv import load_dotenv


load_dotenv()

DB_HOST = os.getenv('DB_HOST')
DB_USER = os.getenv('DB_USER')
DB_PASSWORD = os.getenv('DB_PASSWORD')
DB_NAME = os.getenv('DB_NAME')
BACKUP_FOLDER = os.getenv('BACKUP_FOLDER')

def backup_database():
    os.makedirs(BACKUP_FOLDER, exist_ok=True)

    waktu = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    nama_file = f"{DB_NAME}-{waktu}.sql"
    path_file = os.path.join(BACKUP_FOLDER, nama_file)
    
    env = os.environ.copy()
    env["MYSQL_PWD"] = DB_PASSWORD  

    command = [
        r'C:\xampp\mysql\bin\mysqldump.exe',  
        f'-h{DB_HOST}',
        f'-u{DB_USER}',
        DB_NAME,
        '--no-tablespaces'
    ]

    try:
        with open(path_file, 'w') as file:
            subprocess.run(command, stdout=file, check=True, env=env)
        print(f"✅ Backup berhasil: {path_file}")
    except subprocess.CalledProcessError as e:
        print(f"❌ Gagal melakukan backup: {e}")
    except Exception as e:
        print(f"❌ Terjadi error lain: {e}")

if __name__ == '__main__':
    backup_database()
