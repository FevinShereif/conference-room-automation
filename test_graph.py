from msal import ConfidentialClientApplication
import requests
from config import *

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

token_result = app.acquire_token_for_client(
    scopes=["https://graph.microsoft.com/.default"]
)

if "access_token" not in token_result:
    print("Failed to get token")
    print(token_result)
    exit()

access_token = token_result["access_token"]

headers = {
    "Authorization": f"Bearer {access_token}"
}

url = f"https://graph.microsoft.com/v1.0/users/{ROOM_EMAIL}/calendar/events"

response = requests.get(url, headers=headers)

print("Status Code:", response.status_code)
print(response.json())