import requests
import os

access_token = ''
sheet_id = INT
base_download_path = r'C:\\Users\\ChristianCanty\\Documents\\ProdApp\\SavedImages\\'  # base download folder
column_dict = {}

def get_columns(sheet_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()
    columns = data['columns']
    for column in columns:
        column_dict[column['title']] = column['id']


def get_attachments(sheet_id, row_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}?include=attachments"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()
    for row in data['rows']:
        if row['id'] == row_id:
            if 'attachments' in row:
                return row['attachments']
    return []

def get_attachment_details(sheet_id, attachment_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/attachments/{attachment_id}"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def get_sheet_data(sheet_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}?include=attachments"
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def get_cell_value(row_data, column_id=5190131364614020): # Address column ID
    for cell in row_data['cells']:
        if cell['columnId'] == column_id:
            return cell.get('value', None)
    return None

def create_folder(folder_name, base_path):
    folder_path = os.path.join(base_path, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def change_FilesSynced(sheet_id, row_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows/{row_id}"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    get_columns(sheet_id)
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


def download_and_save_attachments():
    sheet_data = get_sheet_data(sheet_id)

    for row in sheet_data['rows']:
        row_id = row['id']
        attachments = get_attachments(sheet_id, row_id)
        
        if not attachments:
            continue

        # Get the folder name based on the cell value (e.g., address)
        folder_name = get_cell_value(row)
        if not folder_name:
            continue

        # Create a folder using the cell value (e.g., address)
        download_path = create_folder(folder_name, base_download_path)

        # Download and save each attachment
        for attachment in attachments:
            attachment_id = attachment['id']
            attachment_name = attachment['name']
            
            # Define the full file path for the download
            file_path = os.path.join(download_path, attachment_name)

            # Check if the file already exists
            if os.path.exists(file_path):
                continue
            
            # Get the attachment details to find the download URL
            attachment_details = get_attachment_details(sheet_id, attachment_id)
            download_url = attachment_details.get('url')

            if not download_url:
                continue
            
            try:
                # Download the file content
                response = requests.get(download_url)
                response.raise_for_status()
                
                # Write the file content to the local file system
                with open(file_path, 'wb') as file:
                    file.write(response.content)
                

            except requests.exceptions.RequestException as e:
                print(f"Request error occurred: {e}")

        # Change the FilesSynced cell value to True
        change_FilesSynced(sheet_id, row_id)
    


download_and_save_attachments()

get_columns(sheet_id)
# print(column_dict)
