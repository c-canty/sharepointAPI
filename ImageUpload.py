import requests
import os

access_token = ''
sheet_id = 6288161969753988  
base_download_path = r'C:\\Users\\ChristianCanty\\Documents\\ProdApp\\SavedImages\\'  # base download folder

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

def get_cell_value(row_data, column_id=5190131364614020):
    for cell in row_data['cells']:
        if cell['columnId'] == column_id:
            return cell.get('value', None)
    return None

def create_folder(folder_name, base_path):
    folder_path = os.path.join(base_path, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def download_and_save_attachments():
    sheet_data = get_sheet_data(sheet_id)

    for row in sheet_data['rows']:
        row_id = row['id']
        attachments = get_attachments(sheet_id, row_id)
        
        if not attachments:
            print(f"No attachments found for row {row_id}.")
            continue

        # Get the folder name based on the cell value (e.g., address)
        folder_name = get_cell_value(row)
        if not folder_name:
            print(f"No valid folder name found for row {row_id}.")
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
                print(f"File already exists: {file_path}. Skipping download.")
                continue
            
            # Get the attachment details to find the download URL
            attachment_details = get_attachment_details(sheet_id, attachment_id)
            download_url = attachment_details.get('url')

            if not download_url:
                print(f"Download URL not found for attachment ID {attachment_id}")
                continue
            
            try:
                # Download the file content
                response = requests.get(download_url)
                response.raise_for_status()
                
                # Write the file content to the local file system
                with open(file_path, 'wb') as file:
                    file.write(response.content)
                
                print(f"Downloaded: {attachment_name} to {file_path}")

            except requests.exceptions.RequestException as e:
                print(f"Request error occurred: {e}")

download_and_save_attachments()
