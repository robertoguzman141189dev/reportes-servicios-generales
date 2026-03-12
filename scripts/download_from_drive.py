#!/usr/bin/env python3
"""
Descarga el archivo Excel más reciente de Google Drive
"""

import os
import io
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Configuración
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
OUTPUT_PATH = 'data/input.xlsx'


def get_credentials():
    """Obtiene las credenciales de Google desde variable de entorno o archivo"""
    creds_path = os.environ.get('GOOGLE_CREDENTIALS', '/tmp/credentials.json')
    
    if os.path.exists(creds_path):
        return service_account.Credentials.from_service_account_file(
            creds_path, scopes=SCOPES
        )
    
    # Intentar desde variable de entorno como JSON string
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        creds_info = json.loads(creds_json)
        return service_account.Credentials.from_service_account_info(
            creds_info, scopes=SCOPES
        )
    
    raise ValueError("No se encontraron credenciales de Google")


def get_latest_excel(service, folder_id):
    """Obtiene el archivo Excel más reciente de la carpeta"""
    
    # Buscar archivos Excel en la carpeta
    query = f"'{folder_id}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false"
    
    results = service.files().list(
        q=query,
        orderBy='modifiedTime desc',
        pageSize=1,
        fields='files(id, name, modifiedTime)'
    ).execute()
    
    files = results.get('files', [])
    
    if not files:
        raise FileNotFoundError("No se encontraron archivos Excel en la carpeta de Drive")
    
    return files[0]


def download_file(service, file_id, output_path):
    """Descarga un archivo de Google Drive"""
    
    # Crear directorio si no existe
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    request = service.files().get_media(fileId=file_id)
    
    with io.FileIO(output_path, 'wb') as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status:
                print(f"Descargando... {int(status.progress() * 100)}%")
    
    print(f"✅ Archivo descargado: {output_path}")


def main():
    # Obtener configuración
    folder_id = os.environ.get('DRIVE_FOLDER_ID')
    specific_file_id = os.environ.get('FILE_ID', '').strip()
    
    if not folder_id and not specific_file_id:
        raise ValueError("Se requiere DRIVE_FOLDER_ID o FILE_ID")
    
    # Autenticar
    print("🔐 Autenticando con Google Drive...")
    credentials = get_credentials()
    service = build('drive', 'v3', credentials=credentials)
    
    # Obtener archivo
    if specific_file_id:
        print(f"📄 Descargando archivo específico: {specific_file_id}")
        file_info = service.files().get(fileId=specific_file_id, fields='id,name').execute()
        file_id = specific_file_id
    else:
        print(f"📁 Buscando Excel más reciente en carpeta: {folder_id}")
        file_info = get_latest_excel(service, folder_id)
        file_id = file_info['id']
    
    print(f"📊 Archivo encontrado: {file_info['name']}")
    
    # Descargar
    download_file(service, file_id, OUTPUT_PATH)
    
    # Guardar metadata
    with open('data/file_info.json', 'w') as f:
        json.dump(file_info, f, indent=2)
    
    print("✅ Descarga completada")


if __name__ == '__main__':
    main()
