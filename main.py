import requests
import json
import time
import random
import logging
import os
import sys

# Configure logging
logging.basicConfig(
    filename='workflow_steps.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Load credentials from environment variables
refresh_token = os.getenv("REFRESH_TOKEN")
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")

# Validate environment variables
if not refresh_token or not client_id or not client_secret:
    logging.error("Missing one or more required environment variables: REFRESH_TOKEN, CLIENT_ID, CLIENT_SECRET")
    sys.exit(1)

# Map endpoints to human-readable task descriptions
endpoint_tasks = {
    'https://graph.microsoft.com/v1.0/me/drive/root': "Retrieve user's root drive information",
    'https://graph.microsoft.com/v1.0/me/drive': "Retrieve user's drive information",
    'https://graph.microsoft.com/v1.0/drive/root': "Retrieve root drive information",
    'https://graph.microsoft.com/v1.0/users': "List users in the organization",
    'https://graph.microsoft.com/v1.0/me/messages': "Retrieve user's email messages",
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules': "Retrieve inbox message rules",
    'https://graph.microsoft.com/v1.0/me/drive/root/children': "List items in user's root drive",
    'https://api.powerbi.com/v1.0/myorg/apps': "List Power BI apps in the organization",
    'https://graph.microsoft.com/v1.0/me/mailFolders': "List user's mail folders",
    'https://graph.microsoft.com/v1.0/me/outlook/masterCategories': "Retrieve user's Outlook master categories",
    'https://graph.microsoft.com/v1.0/applications?$count=true': "Count applications in the directory",
    'https://graph.microsoft.com/v1.0/me/?$select=displayName,skills': "Retrieve user's display name and skills",
    'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta': "Retrieve delta of inbox messages",
    'https://graph.microsoft.com/beta/me/outlook/masterCategories': "Retrieve beta Outlook master categories",
    'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top=1': "Retrieve beta message headers",
    'https://graph.microsoft.com/v1.0/sites/root/lists': "List SharePoint site lists",
    'https://graph.microsoft.com/v1.0/sites/root': "Retrieve SharePoint site root information",
    'https://graph.microsoft.com/v1.0/sites/root/drives': "List SharePoint site drives"
}

calls = list(endpoint_tasks.keys())

def get_access_token(refresh_token, client_id, client_secret):
    logging.info("Step 1: Connecting to Microsoft Graph to obtain access token.")
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
        logging.error("Failed to retrieve access token. Response: %s", json.dumps(jsontxt, indent=2))
        raise KeyError("access_token not found in the response.")
    access_token = jsontxt['access_token']
    logging.info("Access token successfully retrieved.")
    return access_token

def main():
    logging.info("Workflow execution started.")
    random.shuffle(calls)
    endpoints = calls[random.randint(0, 10)::]
    try:
        access_token = get_access_token(refresh_token, client_id, client_secret)
    except Exception as e:
        logging.error("Error obtaining access token: %s", str(e))
        return

    session = requests.Session()
    session.headers.update({
        'Authorization': access_token,
        'Content-Type': 'application/json'
    })

    for i, endpoint in enumerate(endpoints, start=2):
        task_description = endpoint_tasks.get(endpoint, "Unknown task")
        logging.info("Step %d: %s", i, task_description)
        logging.info("Calling endpoint: %s", endpoint)
        try:
            response = session.get(endpoint)
            if response.status_code == 200:
                logging.info("Call successful. Status: %d", response.status_code)
                logging.debug("Response content: %s", response.text)
            else:
                logging.warning("Call failed. Status: %d", response.status_code)
                logging.warning("Response content: %s", response.text)
        except requests.exceptions.RequestException as e:
            logging.error("Request exception for endpoint %s: %s", endpoint, str(e))

    localtime = time.asctime(time.localtime(time.time()))
    logging.info("Workflow execution completed at: %s", localtime)

for _ in range(3):
    main()
