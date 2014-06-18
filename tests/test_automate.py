from nose.tools import *
import builder.automate as ba
from pywinauto import WindowNotFoundError, application


def test_exe():
    exe = ba.exe("/")
    assert exe.split('.')[1] == 'exe'


@raises(WindowNotFoundError)
def test_connect_not_running():
    ba.connect()


def test_connect():
    ba.start(ba.exe(ba.default_path))
    app = ba.connect()
    assert type(app) == application.Application
    app.kill_()


def test_start():
    app = ba.start(ba.exe(ba.default_path))
    assert type(app) == application.Application
    app.kill_()


def test_builder():
    bd = ba.Builder()
    assert type(bd.app) == application.Application
    assert bd.app.window_(title_re=ba.match_untitled).Exists()
    bd.app.kill_()


def test_builder_exit():
    bd = ba.Builder()
    bd.exit()
    assert not bd.app.window_(title_re=ba.match_untitled).Exists()


def test_builder_main():
    bd = ba.Builder()
    main = bd.main()
    assert type(main) == application.WindowSpecification
    bd.app.kill_()


if __name__ == '__main__':
    import nose

    nose.runmodule(argv=[__file__, '-v'])
