import os
import requests
import json
import time
import random
import logging

# Configure logging
logging.basicConfig(
    filename='flow_run.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Load credentials from environment variables
refresh_token = os.getenv("REFRESH_TOKEN")
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")

calls = [
    'https://graph.microsoft.com/v1.0/me/drive/root',
    'https://graph.microsoft.com/v1.0/me/drive',
    'https://graph.microsoft.com/v1.0/drive/root',
    'https://graph.microsoft.com/v1.0/users',
    'https://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    'https://graph.microsoft.com/v1.0/me/drive/root/children',
    'https://api.powerbi.com/v1.0/myorg/apps',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',
    'https://graph.microsoft.com/v1.0/applications?$count=true',
    'https://graph.microsoft.com/v1.0/me/?$select=displayName,skills',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    'https://graph.microsoft.com/beta/me/outlook/masterCategories',
    'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top=1',
    'https://graph.microsoft.com/v1.0/sites/root/lists',
    'https://graph.microsoft.com/v1.0/sites/root',
    'https://graph.microsoft.com/v1.0/sites/root/drives'
]

def get_access_token(refresh_token, client_id, client_secret):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': 'http://localhost:53682/'
    }
    response = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(response.text)
    if 'access_token' not in jsontxt:
        logging.error("Failed to retrieve access token. Response from server: %s", json.dumps(jsontxt, indent=2))
        raise KeyError("access_token not found in the response.")
    access_token = jsontxt['access_token']
    return access_token

def main():
    random.shuffle(calls)
    endpoints = calls[random.randint(0, 10)::]
    try:
        access_token = get_access_token(refresh_token, client_id, client_secret)
        logging.info("Access token successfully retrieved.")
    except Exception as e:
        logging.error("Error obtaining access token: %s", str(e))
        return

    session = requests.Session()
    session.headers.update({
        'Authorization': access_token,
        'Content-Type': 'application/json'
    })
    num = 0
    for endpoint in endpoints:
        try:
            response = session.get(endpoint)
            if response.status_code == 200:
                num += 1
                logging.info("Call %d successful: %s", num, endpoint)
            else:
                logging.warning("Call failed with status %d: %s", response.status_code, endpoint)
        except requests.exceptions.RequestException as e:
            logging.error("Request exception for endpoint %s: %s", endpoint, str(e))
            pass
    localtime = time.asctime(time.localtime(time.time()))
    logging.info("End of run: %s", localtime)
    logging.info("Number of calls made: %d", len(endpoints))

for _ in range(3):
    main()
