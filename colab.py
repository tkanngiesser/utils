import os
from google.colab import files

def get_config_json():
    '''allows to upload config_json, 
    returns dict and removes uploaded file again'''
    file = files.upload()
    file_name = list(file.keys())[0]
    config_json = str(file[file_name])
    config_json = config_json[config_json.find('{'):config_json.find('}')+1]
    os.remove(file_name)
    return config_json
