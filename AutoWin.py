from pywinauto.application import Application
import pywinauto
# appIE=Application().connect(class_name="IEFrame",visible_only='true',top_level_only=False)
# appIE.window().minimize()
# from win32com.client import Dispatch
# функция, которая определяет URL открытого IE
from win32com.client import Dispatch
SHELL = Dispatch("Shell.Application")
def get_ie(shell):
    for win in shell.Windows():
        if win.Name == "Internet Explorer":
            return win
    return None

def GetUrl():
    ie = get_ie(SHELL)
    if ie:
        return ie.LocationURL

    else:
        return "error"

