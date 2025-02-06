"""
Collecting logs from module 'logging'.
Use available options for your logging:
    logging.debug('Your text for logging')
    logging.info('Your text for logging')
    logging.warning('Your text for logging')
    logging.error('Your text for logging')
    logging.critical('Your text for logging')

Detailed information about logging levels: https://docs.python.org/2/howto/logging.html

Also, catching missed exceptions from stderr and writing to log file and log window
"""
import logging
import os
import sys
import traceback
from logging.handlers import RotatingFileHandler

import wx
import wx.lib.agw.genericmessagedialog as GMD

DATEFORMAT = '%Y-%m-%d %H:%M:%S'
CONSOLE_FORMATTER = ROOT_FORMATTER = logging.Formatter(
    fmt="%(asctime)s > %(filename)s > %(lineno)s > %(levelname)s : %(message)s", datefmt=DATEFORMAT)
LOGS_FOLDER = os.path.join(os.getcwd(), 'onliner_parser_logs')
ROOT_LOG = 'common.log'
PATH_TO_ROOT_LOG = os.path.join(LOGS_FOLDER, ROOT_LOG)

if not os.path.exists(LOGS_FOLDER):
    os.mkdir(LOGS_FOLDER)


def get_console_handler() -> logging.StreamHandler:
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(CONSOLE_FORMATTER)
    console_handler.setLevel(logging.DEBUG)
    return console_handler


def get_file_handler() -> RotatingFileHandler:
    file_handler = RotatingFileHandler(PATH_TO_ROOT_LOG, maxBytes=1096, encoding='utf-8')
    file_handler.setFormatter(ROOT_FORMATTER)
    file_handler.setLevel(logging.DEBUG)
    return file_handler


root = logging.getLogger()
root.setLevel(logging.DEBUG)
root.addHandler(get_console_handler())
root.addHandler(get_file_handler())


def exception_hook(etype, value, trace) -> None:
    """
    Handler for all unhandled exceptions.
    """
    tmp = traceback.format_exception(etype, value, trace)
    logging.error("Uncaught exception", exc_info=(etype, value, trace))
    exception = ''.join(tmp)
    dlg = ExceptionDialog(msg=exception)
    dlg.ShowModal()


class ExceptionDialog(GMD.GenericMessageDialog):
    """
    Dialogue for output a message with an exception in case of a crash.
    """

    def __init__(self, msg: str) -> None:
        GMD.GenericMessageDialog.__init__(self, None, msg, "Exception!", wx.ICON_ERROR | wx.OK)
        self.base_message = msg

    def OnOk(self, event: wx.Event) -> None:
        """
        Ok button override to close the window and copy to the buffer of the exception message.
        """
        data_obj = wx.TextDataObject()
        data_obj.SetText(str(self.base_message))
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(data_obj)
            wx.TheClipboard.Close()
        wx.MessageBox('Ошибка скопирована в буфер обмена')
        self.Destroy()
