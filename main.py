from GoogleAuth import *
from MicrosoftAuth import *

GoogleApi = GoogleSheetsAPI('googleauth.json', '18fJgtsbDTYvxFF7JXdJ11hdAcZr6kDKVH2IQs_rwCco') ##my json file is been gitignore please use your creadentials json (readme.md)
GoogleApi.connect()

# Read data from sheet at index 0
df = GoogleApi.read_data(0)
print(df)



##if the sheet is in the microsoft excel

MicrosoftApi = MicrosoftExcelAPI(credentials_file_path_text='microsoftauth.txt', drive_id= '81d7c6d73cdb7803', file_id='81D7C6D73CDB7803!22456')

MicrosoftApi.connect()


#Read data from worksheet
df = MicrosoftApi.read_data()
print(df)
