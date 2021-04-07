import os
import numpy as np
import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.client_context import ClientCredential
from office365.sharepoint.files.file import File
from office365.sharepoint.files.file_creation_information import FileCreationInformation
from office365.sharepoint.folders.folder import Folder


# after generating client id and client secret store in a secured way (azure key vault, etc)
config = {
    "url": "your_url",
    "client_id": "your_client_id",
    "client_secret": "your_client_secret",
}

url = config["url"]
client_id = config["client_id"]
client_secret = config["client_secret"]

# client context
credentials = ClientCredential(client_id, client_secret)
ctx = ClientContext(url).with_credentials(credentials)
ctx

# helper functions

def get_file_names(ctx, url, folder):
    """get all file names in folder
    TODO: return file objects to filter, sort, etc instead of just file names
    """
    rel_url = url[url.find("sites") - 1 :]

    spo_folder = ctx.web.get_folder_by_server_relative_url(
        f"/{rel_url}/Shared%20Documents/{folder}"
    )

    ctx.load(spo_folder)
    ctx.execute_query()

    spo_folder_files = spo_folder.files
    ctx.load(spo_folder_files)
    ctx.execute_query()

    files = []
    for file in spo_folder_files:
        file_properties = file.properties["Name"]
        file_time_last_modified = file.properties["TimeLastModified"]
        files.append(file)
    return files

def download_file(ctx, url, source, file, target):
    """download file from SPO to target folder"""
    rel_url = url[url.find('sites') - 1 :]
    if source:
        response_url = f'{rel_url}/Shared%20Documents/{source}/{file}'
    else:
        response_url = f'{rel_url}/Shared%20Documents/{file}'

    local_file_path = os.path.join(target, file)
    response = File.open_binary(ctx, response_url)
    response.raise_for_status()
    with open(local_file_path, 'wb') as local_file:
        local_file.write(response.content)
