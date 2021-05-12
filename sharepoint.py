import subprocess
import sys

def install(package):
    print(f'Installing {package}...')
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    
install('Office365-REST-Python-Client')
install('Office365')

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

def upload_file(ctx, url, target, source, file):
    rel_url = url[url.find("sites") - 1 :]
    path = os.path.join(source, file)
    with open(path, "rb") as content_file:
        file_content = content_file.read()
    target_folder = ctx.web.get_folder_by_server_relative_url(
        f"/{rel_url}/Shared%20Documents/{target}"
    )
    info = FileCreationInformation()
    info.content = file_content
    info.url = os.path.basename(path)
    info.overwrite = True
    target_file = target_folder.files.add(info)
    ctx.execute_query()

    
### sharepoint

# move to processing
def validate_df(df, schema):
    '''validates and splits bad and good date'''
    df_bad = None
    errors = schema.validate(df)
    errors_idxs = [e.row for e in errors]
    if len(errors_idxs) != 0 and errors_idxs != [-1]:
        df_good = df.drop(index=errors_idxs)
        df_bad = df.iloc[errors_idxs]
        df_bad['error'] = errors
        df = df_good
    return df, df_bad
  
# move to processing
def split_and_upload_df_to_spo(df, metadata, existing_files, target_folder, target_spo_folder):
    for name, group in df.groupby(metadata['split_cols']):
        file_name = f'{metadata["entity"]}-{name[0]}-{name[1]}.csv'
        print(file_name)
        if file_name in existing_files:
            spo.download_file(ctx=spo_cxn, url=cxn['spo_url'], 
                              source=target_spo_folder, 
                              file=file_name, target=target_folder)
            df_existing = pd.read_csv(os.path.join(target_folder, file_name))
            df_existing.append(group)
            df_existing = df_existing.drop_duplicates(subset=metadata['unique_id_cols'])
            df_existing.to_csv(os.path.join(target_folder, file_name), index=False)
        else:
            print('new file needs to be created')
            group.to_csv(os.path.join(target_folder, file_name), encoding=metadata['encoding'], index=False)
        spo.upload_file(ctx=spo_cxn, url=cxn['spo_url'], target=target_spo_folder, source=target_folder, file=file_name)
        existing_files.append(file_name)
        time.sleep(1) 
        #os.remove(os.path.join(target_folder, file_name))
        cxn['target_files_to_exclude'].append(file_name)
        return existing_files
   
# move to processing
def process_files(cxn, config, metadata):

    source_files = get_files_from_spo(spo_url=cxn['spo_url'],
                                         spo_folder = cxn['source_spo_folder'], 
                                         files_to_exlude=config['raw_files_to_exclude'])
    if source_files != []:
        for source_file in source_files:
            print(f'processing {source_file}')
            
            spo.download_file(ctx=spo_cxn, url=cxn['spo_url'],
                              source=cxn['source_spo_folder'],
                              file=source_file, target=cxn['source_folder'])
            df = pd.read_csv(os.path.join(cxn['source_folder'], source_file), 
                             header=0, index_col=False, names=list(metadata['data_types'].keys()),
                             encoding=metadata['encoding'])
            
            df, df_bad = validate_df(df=df, schema=metadata['schema'])
            df_clean = clean_df(df)
            
            target_files = get_files_from_spo(spo_url = cxn['spo_url'],
                               spo_folder = cxn['target_spo_folder'],
                               files_to_exlude=[])


            split_and_upload_df_to_spo(df=df_clean, metadata=metadata,
                                       existing_files=target_files, 
                                       target_folder=cxn['target_folder'],
                                       target_spo_folder=cxn['target_spo_folder'])
            
            os.remove(os.path.join(cxn['source_folder'], source_file))
            config['raw_files_to_exclude'].append(source_file)
            config['raw_files_to_exclude'] = list(set(config['raw_files_to_exclude']))
    else:
        print('nothing to process..')
    return config['raw_files_to_exclude']
  
### processing
def fix_user_market(x):
    x = str(x)
    x = x.replace('VF - ', '')
    x = x.replace('United Kingdom', 'UK')
    x = x.replace('VSS', 'VOIS_')
    x = x.replace(' -', '')
    return x
  
# move to processing
def get_files_from_spo(spo_url, spo_folder, files_to_exlude):
    files = []
    spo_files = spo.get_file_names(ctx=spo_cxn, url=spo_url, folder=spo_folder)
    for i, file in enumerate(spo_files):
        files.append(spo_files[i].properties['Name'])
    files = [file for file in files if file not in files_to_exlude]
    return files
  
