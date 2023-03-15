import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd


class GoogleSheetsAPI:
    def __init__(self, credentials_file_path_json, sheet_id):
        self.credentials_file_path = credentials_file_path_json
        self.sheet_id = sheet_id
        self.credentials = None
        self.gc = None
        self.sh = None
    
    #connects with the google sheets api    
    def connect(self):
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_file_path, scope)
        self.gc = gspread.authorize(self.credentials)
        self.sh = self.gc.open_by_key(self.sheet_id)
    
    #read the data as dataframes    
    def read_data(self, sheet_index):
        worksheet = self.sh.get_worksheet(sheet_index)
        data = worksheet.get_all_values()
        headers = data[0]
        values = data[1:]
        df = pd.DataFrame(values, columns=headers)
        return df

    #write the data this method will not be required for any form as it should not be updated from the program at any cost
    def write_data(self, sheet_index, df):
        worksheet = self.sh.get_worksheet(sheet_index)
        headers = df.columns.tolist()
        values = df.values.tolist()
        worksheet.clear()
        worksheet.append_row(headers)
        for row in values:
            worksheet.append_row(row)
        

        

    

