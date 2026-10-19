import urllib.request
import json

url = "http://localhost:8000/webhook"
payload = {
    "type": "message",
    "source": {
        "userId": "test_user_01"
    },
    "createdTime": 1726360000000,
    "content": {
        "type": "file",
        "fileId": "mock_file_id_12345",
        "fileName": "report_sample.xlsx"
    }
}

data = json.dumps(payload).encode("utf-8")
req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})

try:
    with urllib.request.urlopen(req) as res:
        print("Status Code:", res.status)
        print("Response Body:", res.read().decode("utf-8"))
except Exception as e:
    print("Error:", e)