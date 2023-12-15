import json
from AutoDocumentationCreator.const import *
# https://www.daleseo.com/python-json/

class JsonSerializationHandler:
    def __init__(self):
        pass

    def save_json(self, path, json_object):
        with open(path, 'w') as f:
            json.dump(json_object, f, indent=2)

    def read_json(self, path):
        with open(path) as f:
            json_object = json.load(f)
        return json_object
