import pandas as pd
import requests
import msal


def read_credentials(filename):
    with open(filename, "r+") as file:
        text = file.read() 
    lines = text.split("\n") 
    pairs = [line.split(":") for line in lines]

    pairs = {key.strip(): value.strip() for key, value in pairs}
    return pairs

class MicrosoftExcelAPI:

    def __init__(self, credentials_file_path_text, drive_id, file_id):
        credentials_dict = read_credentials(credentials_file_path_text)
        self.client_id = credentials_dict["Application (client) ID"]
        self.client_secret = credentials_dict["Client Secret Value"]
        self.tenant_id = credentials_dict["Directory (tenant) ID"]
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority,
        )
        self.drive_id = drive_id
        self.file_id = file_id

    def connect(self):
        result = self.app.acquire_token_silent(self.scope, account=None)
        if not result:
            result = self.app.acquire_token_for_client(scopes=self.scope)

        self.access_token = result.get("access_token")

    def read_data(self):
        if not self.access_token:
            raise Exception("Access token is not available. Please call connect method to authenticate.")

        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{self.file_id}/workbook/worksheets"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(url, headers=headers)

        data = []
        for sheet in response.json()["value"]:
            sheet_name = sheet["name"]
            sheet_id = sheet["id"]
            sheet_range = sheet["usedRange"]["address"]

            url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{self.file_id}/workbook/worksheets('{sheet_id}')/range(address='{sheet_range}')"
            response = requests.get(url, headers=headers)
            values = response.json()["values"]
            df = pd.DataFrame(values[1:], columns=values[0])
            data.append((sheet_name, df))

        return dict(data)

    def write_data(self, sheet_name, df):
        if not self.access_token:
            raise Exception("Access token is not available. Please call connect method to authenticate.")

        headers = {"Authorization": f"Bearer {self.access_token}", "Content-Type": "application/json"}

        # create new worksheet
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{self.file_id}/workbook/worksheets"
        data = {"name": sheet_name}
        response = requests.post(url, headers=headers, json=data)
        worksheet_id = response.json()["id"]

        # clear cells in the worksheet
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{self.file_id}/workbook/worksheets('{worksheet_id}')/range/clear"
        response = requests.post(url, headers=headers)

        # write dataframe to worksheet
        rows = []
        columns = []
        for col in df.columns:
            column = {"values": [col]}
            columns.append(column)
        rows.append({"index": 0, "values": columns})

        for i, row in df.iterrows():
            cells = []
            for _, value in row.iteritems():
                cells.append({"values": [value]})
            rows.append({"index": i+1, "values": cells})

        data = {"values": rows}
        url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{self.file_id}/workbook/worksheets('{worksheet_id}')/range"
        response = requests.patch(url, headers=headers, json=data)

        if response.status_code == 200:
            print(f"Data written to worksheet '{sheet_name}' successfully.")
        else:
            print(f"Error writing data to worksheet '{sheet_name}'.")