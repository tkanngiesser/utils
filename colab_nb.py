import os
from google.colab import files
import json

def get_config_json():
    '''allows to upload config_json, 
    returns dict and removes uploaded file again'''
    file = files.upload()
    file_name = list(file.keys())[0]
    with open(file_name) as json_file:
        data = json.load(json_file)
    os.remove(file_name)
    return data
