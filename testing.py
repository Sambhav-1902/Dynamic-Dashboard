import requests

# Replace with your actual SharePoint URL
SHAREPOINT_URL = "https://exlservice.sharepoint.com"

try:
    response = requests.get(SHAREPOINT_URL, timeout=10)
    print("Status code:", response.status_code)
    print("SUCCESS — SharePoint is reachable from your machine")
except Exception as e:
    print("FAILED:", e)
