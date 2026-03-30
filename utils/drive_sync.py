
import io
import json
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']


def get_drive_service(credentials_json_bytes):
    info = json.load(io.BytesIO(credentials_json_bytes))
    creds = service_account.Credentials.from_service_account_info(
        info,
        scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)


def list_files_with_ext_in_folder(service, folder_id, ext):
    # 先列出所有檔案，顯示 debug 訊息
    debug_results = service.files().list(q=f"'{folder_id}' in parents and trashed=false", fields="files(id, name, mimeType, createdTime)").execute()
    debug_files = debug_results.get('files', [])
    print("[Google Drive DEBUG] 檔案列表：")
    for f in debug_files:
        print(f"  {f['name']} | {f['mimeType']}")

    # 只用副檔名過濾
    files = [f for f in debug_files if f['name'].lower().endswith(ext)]
    files.sort(key=lambda x: x['createdTime'], reverse=True)
    return files


def download_drive_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh.read()


def get_latest_csv_dataframe(credentials_json_bytes, folder_id):
    service = get_drive_service(credentials_json_bytes)
    files = list_files_with_ext_in_folder(service, folder_id, '.csv')
    if not files:
        raise FileNotFoundError('找不到任何 CSV 檔案')
    latest_file = files[0]
    csv_bytes = download_drive_file(service, latest_file['id'])
    df = pd.read_csv(io.BytesIO(csv_bytes), encoding='utf-8-sig')
    return df, latest_file['name']

def get_latest_xml_dataframe(credentials_json_bytes, folder_id, year=None, month=None):
    from .xml_importer import parse_cwmoney_xml
    service = get_drive_service(credentials_json_bytes)
    files = list_files_with_ext_in_folder(service, folder_id, '.xml')
    if not files:
        raise FileNotFoundError('找不到任何 XML 檔案')
    latest_file = files[0]
    xml_bytes = download_drive_file(service, latest_file['id'])
    df = parse_cwmoney_xml(xml_bytes, year=year, month=month)
    return df, latest_file['name']
