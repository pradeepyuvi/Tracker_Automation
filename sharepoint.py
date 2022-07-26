import io
import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.file import File

url = "https://amedeloitte.sharepoint.com/sites/TR-Clear-DataEngineering"
username = '********'
password = '############'

ctx_auth = AuthenticationContext(url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(url, ctx_auth)
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print("Authentication successful")
    file_path = '/sites/TR-Clear-DataEngineering/Shared%20Documents/General/Daily%20Status%20Trackers/CLEAR-DE%20Daily%20Status%20Tracker.xlsx'
    #file_path= 'https://amedeloitte.sharepoint.com/:x:/r/sites/TR-Clear-DataEngineering/Shared%20Documents/General/Daily%20Status%20Trackers/CLEAR-DE%20Daily%20Status%20Tracker.xlsx?d=w1d6d8959378544e29bb51c5b5aad2e51&csf=1&web=1&e=GfZyof'
    response = File.open_binary(ctx, file_path)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(response.content)
    bytes_file_obj.seek(0)
else:
    print(ctx_auth.get_last_error())