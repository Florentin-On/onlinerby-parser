import inspect
import logging
from typing import Any, Optional, Union

import requests
from requests import exceptions


def safe_get_requester(url: str, default_return: Optional[Union[dict, list]] = None, raw_response: bool = False,
                       params: Any = None, **kwargs: Any) -> Optional[Union[dict, list, requests.Response]]:
    frame = inspect.currentframe()

    function_name = module_name = 'unknown'
    if frame and frame.f_back:
        function_name = frame.f_back.f_code.co_name
        module_name = frame.f_back.f_globals["__name__"]
    func_log_message = f'{function_name} from {module_name}:'
    try:
        logging.debug(f'{func_log_message} sending GET request with parameters: {kwargs}')
        response = requests.get(url, params=params, **kwargs)
        response.raise_for_status()
        if raw_response:
            return response
        data = response.json()
        return data

    except exceptions.HTTPError as err_http:
        logging.warning(f'{func_log_message} HTTP Error. Parameters: {kwargs}\n{err_http}')

    except exceptions.ConnectionError as err_connect:
        logging.warning(f'{func_log_message} Error Connecting. Parameters: {kwargs}\n{err_connect}')

    except exceptions.Timeout as err_timeout:
        logging.warning(f'{func_log_message} Timeout Error. Parameters: {kwargs}\n{err_timeout}')

    except exceptions.RequestException as err_request:
        logging.warning(f'{func_log_message} Request error has occurred. Parameters: {kwargs}\n{err_request}')

    except BaseException as err_other:
        logging.warning(f'{func_log_message} occurred an unexpected error. Parameters: {kwargs}\n{err_other}')

    logging.warning(f'An error has occurred, using default data: {default_return}')

    return default_return
