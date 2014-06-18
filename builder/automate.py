"""
Automation for Adaptrade Builder
Expects default layout. To reset layout, regedit.exe -> remove this:
HKEY_CURRENT_USER\Software\Adaptrade Software\Builder.
"""

__author__ = 'Chris Degiere (chris@degiere.net)'

import time
import os

from pywinauto import application, WindowNotFoundError

default_path = 'C:\\Program Files\\Adaptrade Software\\Adaptrade Builder 1.6\\'
match_untitled = ".*Untitled.*Builder.*"
match_builder = ".*gpstrat.*Builder"


def wait_existing(ws, tries=5):
    """
    Wait 1 second until window exists or timeout after N tries
    :param ws: WindowSpecification
    """
    for i in range(0, tries):
        if ws.Exists():
            return
        else:
            time.sleep(1)
    raise Exception("timed out after %s tries waiting for condition to be true" % tries)


def exe(path=default_path):
    """
    :param path: fully qualified path to the directory where Builder is installed
    :return: fully qualified path to executable localized for machine architecture
    """
    bit32 = path + 'Builder.exe'
    bit64 = path + 'Builder64.exe'
    if os.path.isfile(bit32):
        return bit32
    else:
        return bit64


def connect():
    """
    Connect to an existing running named or untitled application instance
    :return: pywinauto Application instance
    """
    app = application.Application()
    try:
        return app.connect_(title_re=match_builder)
    except WindowNotFoundError:
        pass
    return app.connect_(title_re=match_untitled)


def start(exe_path):
    """ Start Builder and wait until fully loaded
    :param exe_path: fully qualified path to executable
    :return: pywinauto application instance
    """
    app = application.Application().start(exe_path, timeout=10)
    wait_existing(app.window_(title_re=match_untitled), 30)
    return app


class Builder:
    """ Abstraction of builder application and GUI functionality """
    def __init__(self, path=default_path):
        self.exe = exe(path)
        try:
            self.app = connect()
        except WindowNotFoundError:
            self.app = start(self.exe)

    def main(self):
        return self.app.top_window_()

    def exit(self):
        main = self.main()
        main.TypeKeys("%f")
        main.TypeKeys("{UP 1}{ENTER}")
