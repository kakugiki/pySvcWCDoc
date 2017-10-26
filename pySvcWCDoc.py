import win32service
import win32serviceutil
import win32event
import configparser
import logging
import time
import datetime
import socket
import os
import inspect
import logging.handlers
from distutils.dir_util import copy_tree


class PySvcWCDoc(win32serviceutil.ServiceFramework):
    _svc_name_ = 'PySvcWCDoc'
    _svc_display_name_ = 'Python Service - WC Doc'
    _svc_description_ = 'This service, written in Python, copies WC files to DocRec'

    _config = configparser.ConfigParser()
    _path = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
    _config.read(_path + '/app.ini')
    _srcpi = _config["default"]["srcpi"]
    _srcwc = _config["default"]["srcwc"]
    _tgtpi = _config["default"]["tgtpi"]
    _tgtwc = _config["default"]["tgtwc"]
    _start = _config["default"]["start"]
    _end = _config["default"]["end"]
    _smtp = _config["default"]["smtp"]
    _sender = _config["default"]["sender"]
    _recipient = _config["default"]["recipient"]
    _userid = _config["default"]["userid"]
    _userkey = _config["default"]["userkey"]
    _port = _config["default"]["port"]
    _cell = _config["default"]["cell"]
    _sqlcon = _config["connection"]["sqlcon"]
    _now = time.strftime("%H:%M")
    _today = datetime.datetime.now()
    _logpath = _config["default"]["logpath"] + 'pyLog' + _today.strftime('%Y%m%d') + '.txt'

    handler = logging.handlers.WatchedFileHandler(os.environ.get("LOGFILE", _logpath))
    formatter = logging.Formatter(logging.BASIC_FORMAT)
    handler.setFormatter(formatter)
    root = logging.getLogger()
    root.setLevel(os.environ.get("LOGLEVEL", "INFO"))
    root.addHandler(handler)


    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)

        # create an event to listen for stop requests on
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)


        # core logic of the service
    def SvcDoRun(self):
        import servicemanager

        # p = usp_appsettings.appConfig.getSetting(self, "default", "path")
        #p = self._config["default"]["path"]
        #f = open(p, 'w+')
        rc = None

        # if the stop event hasn't been fired keep looping
        while rc != win32event.WAIT_OBJECT_0:
            self.CopyPIDocs(self._srcpi, self._tgtpi)

            #f.write(self._now.strftime('%Y%m%d') + '\n')
            #f.flush()
            # block for 5 seconds and listen for a stop event
            rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)

        #f.write('SHUTTING DOWN\n')
        #f.close()


        # called when we're being shut down
    def SvcStop(self):
        # tell the SCM we're shutting down
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # fire the stop event
        win32event.SetEvent(self.hWaitStop)


    def CopyPIDocs(self, src, dst):
        try:
            copy_tree(src, dst)
            logging.info("file copied " + src + " to " + dst)
            # parallel ??
        except Exception as e:
            logging.exception("Error " + e)


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(PySvcWCDoc)
