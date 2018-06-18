import sys
import json

class MandrillConfig:
    def __init__(self):
        path = 'config.json'
        with open(path) as f:
            config = f.read()
        self.config = json.loads(config)


class GoogleSheetsConfig:
    def __init__(self):
        scope = ['https://spreadsheets.google.com/feeds']
        cred_json = {
            "type": "service_account",
            "project_id": "abcdef",
            "private_key_id": "abcdef",
            "private_key": "-----BEGIN PRIVATE KEY-----\n-----END PRIVATE KEY-----\n",
            ....
        }
        self.scope = scope
        self.cred_json = cred_json