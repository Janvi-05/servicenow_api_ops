import requests
from urllib.parse import urlencode
from dotenv import load_dotenv
import os

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

    if response.status_code != 200:
        raise Exception(f"Token request failed: {response.text}")

    data = response.json()
    return data['access_token']


def download_servicenow_pdf(sys_id):   
    bearer_token = get_bearer_token()
    headers = {
        'Authorization': f'Bearer {bearer_token}'
    }
    
    url = f"https://lendlease.service-now.com/sn_hr_core_case.do?PDF&sys_id={sys_id}&sysparm_view=Default%20view"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        filename = f"{sys_id}.pdf"
        with open(filename, "wb") as f:
            f.write(response.content)
        print(f"PDF successfully saved as {filename}")
    else:
        print(f"Failed to download PDF. Status code: {response.status_code}")
        print(response.text)


if __name__ == "__main__":
    sys_id = "00016018db51b450ee773313e29619cc"  # Replace with your actual sys_id
    download_servicenow_pdf(sys_id)
