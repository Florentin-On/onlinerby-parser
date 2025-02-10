"""
Main file for program starting.
All panels are initiating here.
Set your menu elements here.
"""
import logging
import traceback
import webbrowser
from functools import partial

import requests
import wx

from app.common.cache import app_cache, ui_cache
from app.common.constants import MAIN_TITLE, MIN_SIZE, MAX_SIZE, MAIN_SIZE, ABOUT_URL, SETTINGS
from app.common.log_collector import PATH_TO_ROOT_LOG
from app.common.safe_requesters import safe_get_requester
from app.common_ui.dialogs import confirmation_dialog
from app.multiparse.multiparse import Multiparse
from app.multiparse.multiparse_ui import MultiparsePanel
from app.source.welcome_screen import WelcomePanel


class OnlinerParserApp(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, wx.ID_ANY, title=MAIN_TITLE, size=MAIN_SIZE)

        result = wx.YES
        while result == wx.YES:
            result = safe_get_requester('https://google.com/', raw_response=True)
            if result is None:
                result = confirmation_dialog('Ошибка подключения к интернету',
                                             'Невозможно установить соединение с Интернетом. Полный '
                                             'текст ошибки доступен в логах приложения.\nПопробовать установить '
                                             'соединение снова? В случае отказа приложение будет закрыто')
        if result == wx.NO:
            self.on_close_on_start()
            return
        # ----------------------------------------------------------------------------------------------------------
        # Make your changes according to the steps below
        # ------------------
        # STEP 1 - Add your panel to the panels_list
        # ------------------
        self.panels_list = [
            'Welcome screen',
            'Parse Onliner Catalogue'
        ]

        self.panels = {
            panel: {
                'panel': None,
                'controller': None,
                'is_initialized': False,
                'is_shown': False,
            }
            for panel in self.panels_list
        }
        # ------------------
        # STEP 2 - Set up your menu section here
        # ------------------
        help_menu = wx.Menu()
        welcome_screen_point = help_menu.Append(wx.ID_ANY, 'Приветственный экран')
        multiparse_point = help_menu.Append(wx.ID_ANY, 'Обработать каталог Onliner')
        about_point = help_menu.Append(wx.ID_ANY, 'Об авторе')

        menu_bar = wx.MenuBar()
        menu_bar.Append(help_menu, 'Меню')

        # Add the MenuBar to the Frame content
        self.SetMenuBar(menu_bar)

        # ------------------
        # STEP 3 - Add your binds here for Menu sections
        # ------------------
        self.Bind(wx.EVT_MENU, partial(self.switch, 'Welcome screen'), welcome_screen_point)
        self.Bind(wx.EVT_MENU, partial(self.switch, 'Parse Onliner Catalogue'), multiparse_point)
        self.Bind(wx.EVT_MENU, self.on_about, about_point)

        # ------------------
        # STEP 4
        # ------------------
        # add your panel to the "init_panel" method below
        # ------------------
        # END of your changes
        # ------------------

        # ------------------------------------------------------------------------------------------------------------
        # technical part

        # set app size
        self.SetSizeHints(minSize=MIN_SIZE, maxSize=MAX_SIZE)

        # center app on the screen
        self.Centre(direction=wx.BOTH)
        self.ToggleWindowStyle(wx.STAY_ON_TOP)

        # make icon
        icon = wx.Icon()
        # There will normally be a log message if a non-existent file or png file with empty pixels
        # is loaded into a wx.Bitmap.
        # It can be suppressed with wx.LogNull
        no_log = wx.LogNull()
        icon.CopyFromBitmap(wx.Bitmap("app/source/onliner_parser.ico", wx.BITMAP_TYPE_ANY))
        # when noLog is destroyed the old log sink is restored
        del no_log
        self.SetIcon(icon)

        # set main sizer
        self.sizer = wx.BoxSizer(wx.VERTICAL)

        # setup start panel based on settings
        self.SetSizer(self.sizer)
        self.setup_start_panel()

        # bind app closing
        self.Bind(wx.EVT_CLOSE, self.on_close)
        # logging
        logging.info('Onliner Parser loaded')

    # --------------------------------------------------------------------------------------------------------------
    # methods

    def setup_start_panel(self):
        # setup start panel based on settings
        settings_cache = ui_cache.get_from_ui_cache('Settings')
        panel_on_start = None
        panel_on_start_selection = settings_cache.get('panel_on_start_selection')

        common_ui_cache = ui_cache.get_from_ui_cache('Common')
        welcome_screen_shown = common_ui_cache.get('welcome_screen_shown', None)

        if panel_on_start_selection is not None:
            if panel_on_start_selection == SETTINGS['panel_selection_default']:
                panel_on_start = settings_cache.get('panel_on_close')
            else:
                panel_on_start = panel_on_start_selection

        if panel_on_start is not None and welcome_screen_shown:
            self.switch(panel_on_start, True)
        else:
            self.switch('Welcome screen', True)
            ui_cache.update_ui_cache('Common', {'welcome_screen_shown': True})

    def switch(self, panel_for_showing, event):

        if self.panels[panel_for_showing]['is_shown']:
            return

        self.SetTitle(MAIN_TITLE + ' - ' + panel_for_showing)
        shown_panel = None

        for panel in self.panels:
            if self.panels[panel]['is_shown']:
                shown_panel = panel

        if not self.panels[panel_for_showing].get('is_initialized'):
            self.init_panel(panel_for_showing, event)

        if event:
            self.panels[panel_for_showing]['panel'].Show()
            self.panels[panel_for_showing]['panel'].Layout()
            self.panels[panel_for_showing]['is_shown'] = True
            ui_cache.update_ui_cache('Settings', {'panel_on_start_selection': panel_for_showing})
            logging.debug('"{}" shown'.format(panel_for_showing))
            if shown_panel is not None:
                self.panels[shown_panel]['is_shown'] = False
                self.panels[shown_panel]['panel'].Hide()
                logging.debug('"{}" hidden'.format(shown_panel))

        self.sizer.Clear()
        self.sizer.Add(window=self.panels[panel_for_showing]['panel'], proportion=1, flag=wx.EXPAND)
        self.Layout()

    def init_panel(self, panel_for_init, event):
        if panel_for_init == 'Welcome screen':
            logging.info('Initialising "{}"'.format(panel_for_init))
            panel = WelcomePanel(parent=self, size=MAIN_SIZE)

            self.panels[panel_for_init].update(
                {
                    'panel': panel,
                    'controller': None,
                    'is_initialized': True,
                    'is_shown': True,
                }
            )
        if panel_for_init == 'Parse Onliner Catalogue':
            logging.info('Initialising "{}"'.format(panel_for_init))
            panel = MultiparsePanel(parent=self, size=MAIN_SIZE)
            controller = Multiparse(panel)

            self.panels[panel_for_init].update(
                {
                    'panel': panel,
                    'controller': controller,
                    'is_initialized': True,
                    'is_shown': True,
                }
            )

    @staticmethod
    def on_about(event):
        webbrowser.get('windows-default').open(ABOUT_URL)

    def on_close(self, event):
        # save panel on close
        for panel in self.panels:
            if self.panels[panel]['is_shown']:
                ui_cache.update_ui_cache('Settings', {'panel_on_close': panel})
                break

        # save ui cache
        loaded_ui_cache = app_cache.get_from_cache(key='ui_cache')

        if loaded_ui_cache is None:
            ui_cache.save_ui_cache(cache={})

        ui_cache.save_ui_cache(cache=loaded_ui_cache)

        # log
        logging.info('Onliner Parser closed')
        with open(PATH_TO_ROOT_LOG, 'a') as logfile:
            logfile.write('\\----------------------------------------------------------------------------------/\n')
        self.Destroy()

    def on_close_on_start(self):
        logging.warning('Onliner Parser app was closed due error in establishment Internet connection')
        with open(PATH_TO_ROOT_LOG, 'a') as logfile:
            logfile.write('\\----------------------------------------------------------------------------------/\n')
        self.Destroy()
