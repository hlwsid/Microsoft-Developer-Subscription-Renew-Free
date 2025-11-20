import os
import requests
import json
import time
import random

refresh_token = os.getenv("REFRESH_TOKEN")
client_id = os.getenv("CONFIG_ID")
client_secret = os.getenv("CONFIG_KEY")

calls = [
    'https://graph.microsoft.com/v1.0/me/drive/root',
    'https://graph.microsoft.com/v1.0/me/drive',
    'https://graph.microsoft.com/v1.0/drive/root',
    'https://graph.microsoft.com/v1.0/users',
    'https://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    'https://graph.microsoft.com/v1.0/me/drive/root/children',
    'https://api.powerbi.com/v1.0/myorg/apps',
    'https://graph.microsoft.com/v1.0/me/mailFolders',
   count=true',
    'https://graph.microsoft.com/v1.0/me/?$select=displayName,skills',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    'https://graph.microsoft.com/beta/me/outlook/masterCategories',
    'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top=1',
    'https://graph.microsoft.com/v1.0/sites/root/lists',
    'https://graph.microsoft.com/v1.0/sites/root',
    'https://graph.microsoft.com/v1.0/sites/root/drives'
]

def get_access_token(refresh_token, client_id, client_secret):
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret
    }
    response = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = response.json()
    if 'access_token' not in jsontxt:
        raise Exception(f"Failed to refresh token: {jsontxt}")
    # Save new refresh token if provided
    if 'refresh_token' in jsontxt:
        with open("Secret.txt", "w") as f:
            f.write(jsontxt['refresh_token'])
    return jsontxt['access_token']

def main():
    access_token = get_access_token(refresh_token, client_id, client_secret)
    session = requests.Session()
    session.headers.update({
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    })
    random.shuffle(calls)
    endpoints = calls[random.randint(0, 10):]
    num = 0
    for endpoint in endpoints:
        try:
            response = session.get(endpoint)
            if response.status_code == 200:
                num += 1
                print(f'{num}th Call successful: {endpoint}')
            else:
                print(f'Failed call {endpoint}: {response.status_code}')
        except requests.exceptions.RequestException as e:
            print(f'Error calling {endpoint}: {e}')
    print('Run completed at:', time.asctime())
    print('Number of calls:', len(endpoints))

for _ in range(3):
    main()
