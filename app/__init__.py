import json
import logging


def safe_load_json(path_to_json: str) -> dict:
    try:
        return json.load(open(path_to_json, 'r'))
    except Exception as error:
        logging.warning(f'Can\'t load config: {path_to_json}. Error: {error}')
        return {}