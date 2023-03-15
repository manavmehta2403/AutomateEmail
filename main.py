from GoogleAuth import *

api = GoogleSheetsAPI('accountsauth.json', '18fJgtsbDTYvxFF7JXdJ11hdAcZr6kDKVH2IQs_rwCco') ##my json file is been gitignore please use your creadentials json (readme.md)
api.connect()

# Read data from sheet at index 0
df = api.read_data(0)
print(df)

