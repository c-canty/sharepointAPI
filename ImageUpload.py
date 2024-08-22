import requests
import os
import time
from concurrent.futures import ThreadPoolExecutor

access_token = ''
sheet_id =  
base_download_path = r'C:\\Users\\ChristianCanty\\Documents\\ProdApp\\SavedImages\\'
column_dict = {}

def get_columns_and_data(sheet_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}?include=attachments"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()
    for column in data['columns']:
        column_dict[column['title']] = column['id']
    return data

def get_attachment_details(sheet_id, attachment_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/attachments/{attachment_id}"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def get_cell_value(row_data, column_id): 
    for cell in row_data['cells']:
        if cell['columnId'] == column_id:
            return cell.get('value', None)
    return None

def create_folder(folder_name, base_path):
    folder_path = os.path.join(base_path, folder_name)
    try:
        os.makedirs(folder_path, exist_ok=True)
    except OSError as e:
        print(f"Error creating directory {folder_path}: {e}")
    return folder_path

def change_FilesSynced(sheet_id, row_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows/{row_id}"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    FilesSynced_id = column_dict['FilesSynced']
    payload = {
        'cells': [
            {
                'columnId': FilesSynced_id,
                'value': True
            }
        ]
    }
    response = requests.put(url, headers=headers, json=payload)
    response.raise_for_status()

def download_attachment(attachment, download_path):
    attachment_id = attachment['id']
    attachment_name = attachment['name']
    file_path = os.path.join(download_path, attachment_name)

    # Check if the file already exists
    if os.path.exists(file_path):
        return

    # Get the attachment details to find the download URL
    attachment_details = get_attachment_details(sheet_id, attachment_id)
    download_url = attachment_details.get('url')

    if not download_url:
        return

    try:
        # Download the file content
        response = requests.get(download_url)
        response.raise_for_status()
        
        # Write the file content to the local file system
        with open(file_path, 'wb') as file:
            file.write(response.content)

    except requests.exceptions.RequestException as e:
        print(f"Request error occurred: {e}")

def download_and_save_attachments():
    start_time = time.time()

    # Retrieve sheet data and columns in one API call
    sheet_data = get_columns_and_data(sheet_id)

    with ThreadPoolExecutor(max_workers=10) as executor:
        for row in sheet_data['rows']:
            row_id = row['id']

            # Check if FilesSynced is true; if yes, skip the row
            files_synced = get_cell_value(row, column_dict['FilesSynced'])
            if files_synced == 'true':
                continue

            # Get the folder name based on the cell value (e.g., address)
            folder_name = get_cell_value(row, column_dict['Adresse'])
            if not folder_name:
                continue

            # Create a folder using the cell value (e.g., address)
            download_path = create_folder(folder_name, base_download_path)

            # Ensure the folder was created or exists
            if not os.path.exists(download_path):
                print(f"Failed to create or access the folder: {download_path}")
                continue

            # Download and save each attachment in parallel
            attachments = row.get('attachments', [])
            for attachment in attachments:
                executor.submit(download_attachment, attachment, download_path)

            # Change the FilesSynced cell value to True
            change_FilesSynced(sheet_id, row_id)

    # Stop the timer
    end_time = time.time()

    # Calculate the elapsed time
    elapsed_time = end_time - start_time
    print(f"download_and_save_attachments took {elapsed_time:.4f} seconds to run.")
    

if __name__ == '__main__':
    download_and_save_attachments()
