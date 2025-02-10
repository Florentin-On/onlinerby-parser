import logging
import sys

from app.common.log_collector import exception_hook, PATH_TO_ROOT_LOG

sys.excepthook = exception_hook


def run():
    # logging
    with open(PATH_TO_ROOT_LOG, 'a') as logfile:
        logfile.write('\n/----------------------------------------------------------------------------------\\\n')
    logging.debug('Start Onliner Parser')

    import os
    logging.debug(f'Running program from: {os.getcwd()}')

    import wx
    from app.source.onliner_parser_core import OnlinerParserApp
    app = wx.App(False)
    frame = OnlinerParserApp()
    frame.Show(True)
    frame.ToggleWindowStyle(wx.STAY_ON_TOP)
    frame.SetFocus()

    app.MainLoop()


if __name__ == '__main__':
    run()
