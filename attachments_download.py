import requests
import os
import json
 
# Your existing API call
url = "https://lendlease.service-now.com/api/now/attachment?sysparm_query=table_sys_id%3D001d5d331bbba010ef4543b4bd4bcb05"
payload = {}

headers = {
  'Authorization': 'Bearer LFT612CwbhRj2QR_id2LcEojIWraGE9lgpu51BabMHszqcWclNCQIpqHSaDO-pmwGXBowMINr1d_ENTsypdIeQ',
  'Cookie': 'BIGipServerpool_lendlease=c5889ad29f701618e3baa37002034b82; JSESSIONID=3A42AFC667475429E96BBE550763EDE9; glide_node_id_for_js=fc4812175032dd94c0ff92cf846b17cf27f0dce0a6beb49e12e5c7bb0f48d836; glide_session_store=F5FE82552B3D6E50E412F41CD891BF86; glide_user_activity=U0N2M18xOnRMdkppdFlTN2o2cFlnUVdaQ092UjZ6S0pFdXV0dmZBb3BMcGxVa0hrZ1E9OlVBQWc4QWozUERYQi9mVCs2WDRJa0hTRTgwQjkxMGZkMzUrNGxlUXRNUW89; glide_user_route=glide.5a07cc0a1b859ed021434a69d48daaeb'
}

 

def download_attachments():
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
    download_attachments()