import io
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

def authenticate_google_drive():
    try:
        # Conexion con google drive
        creds_json = os.getenv('GOOGLE_DRIVE_CREDENTIALS')
        creds_dict = json.loads(creds_json)
        credentials = service_account.Credentials.from_service_account_info(creds_dict)
        service = build('drive', 'v3', credentials=credentials)

        print("\033[92mConexion con Drive exitosa\033[0m")
        return service
    
    except Exception as e:
        print(f"\033[91mAutenticación con Drive fallida. Verifique sus credenciales. Error: {e}\033[0m")
        return None
    
def download_from_drive(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file = io.BytesIO()
    downloader = MediaIoBaseDownload(file, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"Descargando el template: {service.files().get(fileId=file_id).execute()['name']}: {int(status.progress() * 100)}%")
    file.seek(0)
    return file

def upload_to_drive(service, output_stream, drive_folder_id, file_name):
    try:
        file_metadata = {'name': file_name, 'parents': [drive_folder_id]}
        media = MediaIoBaseUpload(output_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"File ID: {file.get('id')}")
        
    except Exception as e:
        print(f"\033[91mError al subir el archivo a Google Drive: {e}\033[0m")

def list_files_in_folder(service, folder_id):
    try:
        results = service.files().list(
            q=f"'{folder_id}' in parents",
            pageSize=10,
            fields="files(id, name)").execute()
        items = results.get('files', [])
        if not items:
            print("No files found.")
        else:
            print("Files in folder:")
            for item in items:
                print(f"{item['name']} ({item['id']})")
    except Exception as e:
        print(f"Error listing files in folder: {e}")