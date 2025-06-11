import requests
import os
import json
from urllib.parse import urlencode
from dotenv import load_dotenv

load_dotenv()

def get_bearer_token():
    url = "https://lendlease.service-now.com/oauth_token.do"
    payload_dict = {
        'grant_type': 'password',
        'username': os.getenv('SNOW_USERNAME'),
        'password': os.getenv('SNOW_PASSWORD'),
        'client_id': os.getenv('SNOW_CLIENT_ID'),
        'client_secret': os.getenv('SNOW_CLIENT_SECRET')
    }
    payload = urlencode(payload_dict)

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
    }
    response = requests.post(url, data=payload, headers=headers)
    print(f"Response Status Code: {response.status_code}")
    data = response.json()
    print(data)
    return data['access_token'] 

# Your existing API call
url = "https://lendlease.service-now.com/api/now/attachment?sysparm_query=table_sys_id%3D001d5d331bbba010ef4543b4bd4bcb05"
payload = {}

token = get_bearer_token()
headers = {
    
    'Authorization': f'Bearer {token}',
}

 

def download_attachments(headers):
    # Get the attachment list
    response = requests.request("GET", url, headers=headers, data=payload)
    
    if response.status_code == 200:
        data = response.json()
        attachments = data.get('result', [])
        
        # Create downloads directory if it doesn't exist
        table_sys_id=data['result'][0]['table_sys_id']
        print(f"Table Sys ID: {table_sys_id}")
        # exit()
        download_dir = f"attachment_downloads_{table_sys_id}"
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)
        
        print(f"Found {len(attachments)} attachments to download")
        
        for attachment in attachments:
            file_name = attachment.get('file_name')
            download_link = attachment.get('download_link')
            file_size = attachment.get('size_bytes')
            
            if download_link and file_name:
                print(f"\nDownloading: {file_name} ({file_size} bytes)")
                
                try:
                    # Download the file
                    file_response = requests.get(download_link, headers=headers)
                    
                    if file_response.status_code == 200:
                        # Save the file
                        file_path = os.path.join(download_dir, file_name)
                        
                        with open(file_path, 'wb') as f:
                            f.write(file_response.content)
                        
                        print(f"✓ Successfully downloaded: {file_path}")
                    else:
                        print(f"✗ Failed to download {file_name}. Status code: {file_response.status_code}")
                        
                except Exception as e:
                    print(f"✗ Error downloading {file_name}: {str(e)}")
            else:
                print(f"✗ Missing download link or filename for attachment: {attachment.get('sys_id')}")
    
    else:
        print(f"Failed to get attachment list. Status code: {response.status_code}")
        print(f"Response: {response.text}")

# Run the download function
if __name__ == "__main__":
    download_attachments(headers)