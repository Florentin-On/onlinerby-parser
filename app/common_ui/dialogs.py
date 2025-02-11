import wx
import pyperclip
from app.common.cache import ui_cache


def dialog(caption, message, style=wx.OK):
    """
    Information dialog
    :param caption: str: caption of dialog
    :param message: str: message of dialog
    :param style: long: wx styles
                  if add to style wx.ID_HELP "Copy" button be displayed
                  when button pushed the dialog's message be copied to clipboard
    :return: None
    """
    dlg = wx.MessageDialog(parent=None, message=message, caption=caption, style=style)
    dlg.SetHelpLabel('Copy')
    if dlg.ShowModal() == wx.ID_HELP:
        pyperclip.copy(message)


def dialog_with_checkbox(cache_section, caption, message, cache_item, style=wx.OK):
    """
    Information dialog
    :param cache_section: str: section in cache for saving "Don't show again"
    :param caption: str: caption of dialog
    :param message: str: message of dialog
    :param cache_item: str: name of key for cache of this panel to find if "don't show" was checked
    :param style: long: wx styles
    :return: None
    """

    cache = ui_cache.get_from_ui_cache(cache_section)
    dont_show = cache.get(cache_item, False)
    if dont_show:
        return

    dlg = wx.RichMessageDialog(parent=None, message=message, caption=caption, style=style)
    dlg.ShowCheckBox("Don't show again")
    dlg.ShowModal()  # return value ignored as we have "Ok" only anyhow

    if dlg.IsCheckBoxChecked():
        # ... make sure we won't show it again the next time ...
        ui_cache.update_ui_cache(key=cache_section, param={cache_item: True})


def confirmation_with_cancel_dialog(caption, message, style=wx.YES_NO | wx.CANCEL | wx.ICON_EXCLAMATION):
    """
    Confirmation dialog
    By default with YES/NO/CANCEL buttons. User forsed to make selection
    :param caption: str: caption of dialog
    :param message: str: message of dialog
    :param style: long: wx styles
    :return: bool: True if Yes pushed; False if No pushed
    """
    dlg = wx.MessageBox(message=message, caption=caption, style=style)
    return dlg


def confirmation_dialog(caption, message, style=wx.YES_NO | wx.ICON_EXCLAMATION):
    """
    Confirmation dialog
    By default with YES/NO buttons. User forced to make selection
    :param caption: str: caption of dialog
    :param message: str: message of dialog
    :param style: long: wx styles
    :return: bool: True if Yes pushed; False if No pushed
    """
    return wx.MessageBox(message=message, caption=caption, style=style)


def select_file(message, style=wx.FD_DEFAULT_STYLE | wx.FD_FILE_MUST_EXIST):
    dlg = wx.FileDialog(None, message=message,style=style)

    if dlg.ShowModal() == wx.ID_OK:
        return dlg.GetPath()

    dlg.Destroy()


def select_dir(message, style=wx.DD_DIR_MUST_EXIST):
    dlg = wx.DirDialog(None, message=message, style=style)

    if dlg.ShowModal() == wx.ID_OK:
        return dlg.GetPath()

    dlg.Destroy()
