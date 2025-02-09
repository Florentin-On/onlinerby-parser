"""
Cache organization for the app.

Intersession cache (UiCache) commonly used for UI. Save settings end customise UI on start
Option for import:
ui_cache

Session cache (AppCache) should be used for temp data and will be lost on app closing.
"""
import json
import logging
import os
import pickle
from pathlib import Path
from typing import Any

from app import safe_load_json
from app.common.constants import APPDATA_PATH


class CoreCache(object):
    """
    Import an instance of this class (always use g_cache from this module)
    where you need to load data from the cache to your panel.
    """

    def __init__(self) -> None:
        self.__cache_path = os.path.join(APPDATA_PATH, 'cache.json')
        self.__data: dict = {}
        self.__load_cache()
        self._sync_to_file()

    def __load_cache(self) -> None:
        """
        Load data from cache file if it exists.
        """
        if os.path.isfile(self.__cache_path):
            self.__data = safe_load_json(self.__cache_path)

    def _sync_to_file(self) -> None:
        """
        Save data cache to file.
        """
        if not os.path.isdir(APPDATA_PATH):
            logging.warning('Cache folder not found. Creating OnlinerParser folder for cache file.')
            os.makedirs(APPDATA_PATH)
        with open(self.__cache_path, 'w') as cache_file:
            cache_file.write(json.dumps(self.__data, indent=4))

    def get_from_cache(self, key: str) -> Any:
        """
        The method returns data from the cache according to the received key.
        """
        return self.__data.get(key)

    def set_to_cache(self, key: str, value: Any) -> None:
        """
        The method saves the data from the panel to the cache. Format:
        :param key: name panel
        :param value: data
        {
        name_panel: {'login': 'login', 'password': 111111}
        }
        """
        self.__data.update({key: value})
        self._sync_to_file()

    def remove_from_cache(self, key: str) -> None:
        """
        The method removed from cache key
        """
        if key in self.__data:
            self.__data.pop(key)
            self._sync_to_file()


g_cache = CoreCache()


class AppCache(object):
    """
    Temp data for current session
    app_cache format - dict
    {
        KEY: VALUE
    }
    KEY - any type
    VALUE - any type
    """

    def __init__(self):
        self.__app_cache = {}

    def update_cache(self, key, value):
        """
        add or update data in cache
        :param key: str: name of parameter in cache
        :param value: any: any type of data for caching
        :return: None
        """

        self.__app_cache.update({key: value})

    def get_from_cache(self, key):
        """
        get data from cache by given key
        :param key: str: key for part of cache
        :return: saved value for this key or None if key not found
        """

        if key in self.__app_cache:
            return self.__app_cache[key]

    def remove_from_cache(self, key):
        """
        remove data from cache
        :param key: str: key for part of cache
        :return: None
        """
        try:
            del self.__app_cache[key]
        except KeyError:
            pass

    def _get_app_cache(self):
        return self.__app_cache


# init cache
app_cache = AppCache()


class UiCache(object):
    """
    Cache format - dict:
    {
        'CONTROLLER NAME': {'PARAM': VALUE},
        'Common': {'PARAM': VALUE},
    }
    !!! Has 'Common' section that not connected to any controller.
    Controller that use this cache should have self.name, and it would be the key in dict for part of cache
    for this controller.
    Use PARAM with type str and VALUE with any format.
    """

    def __init__(self):
        self.cache_path = self._get_cache_path()

    def load_from_ui_cache(self, section, controller):
        """
        This method get cache for given section from ui_cache,
        and it it's not empty call method "customize_from_ui_cache"
        for a given controller with selected cache as parameter

        Usually called on panel initialising for updating some fields with cached values.

        :param section: str: section name
        :param controller: wx.Panel object: panel for calling to update
        :return: dict - part of cache for this controller
        """
        saved_cache = self._get_ui_cache()

        if section in saved_cache:
            controller.customize_from_ui_cache(saved_cache[section])

    def update_ui_cache(self, key, param):
        """
        add or update data in cache for given key
        :param key: str: section name
        :param param: dict: {str key: any value}
        :return: None
        """
        saved_cache = self._get_ui_cache()

        if key not in saved_cache:
            saved_cache[key] = {}

        saved_cache[key].update(param)

    def get_from_ui_cache(self, section):
        """
        get cache for given section
        :param section: str:
        :return: dict -> part of cache for this section or empty dict if section is not existed
        """
        saved_cache = self._get_ui_cache()

        if section in saved_cache:
            return saved_cache[section]

        return {}

    def remove_from_ui_cache(self, section, param):
        """
        remove param from cache for given section
        :param section: str:
        :param param: dict: {str key: any value}
        :return: None
        """
        saved_cache = self._get_ui_cache()
        try:
            del saved_cache[section][param]
        except KeyError:
            return

    def save_ui_cache(self, cache):

        with open(self.cache_path, 'wb') as handle:
            pickle.dump(cache, handle)

    def _get_ui_cache(self):

        loaded_ui_cache = app_cache.get_from_cache(key='ui_cache')
        if loaded_ui_cache is not None:
            return loaded_ui_cache

        if not self._check_cache_file():
            self.save_ui_cache({})

        with open(self.cache_path, 'rb') as handle:
            saved_cache = pickle.load(handle)

        if not isinstance(saved_cache, dict):
            app_cache.update_cache(key='ui_cache', value={})
            return {}

        app_cache.update_cache(key='ui_cache', value=saved_cache)
        return saved_cache

    @staticmethod
    def _get_cache_path():
        appdata_path = os.getenv('APPDATA')
        file_path = os.path.join(appdata_path, 'OnlinerParser', 'ui_cache.pickle')
        return file_path

    def _check_cache_file(self):
        data_file = Path(self.cache_path)

        if not data_file.exists():
            return False

        return True


# init ui cache
ui_cache = UiCache()
