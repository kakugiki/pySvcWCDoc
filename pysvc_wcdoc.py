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
import multiprocessing as mp
from distutils.dir_util import copy_tree


class PySvcWCDoc(win32serviceutil.ServiceFramework):
    #region Default
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
    _today = datetime.datetime.now()
    _logpath = _config["default"]["logpath"] + 'pyLog' + _today.strftime('%Y%m%d') + '.txt'

    handler = logging.handlers.WatchedFileHandler(os.environ.get("LOGFILE", _logpath))
    handler = logging.FileHandler(_logpath)
    formatter = logging.Formatter(logging.BASIC_FORMAT)
    handler.setFormatter(formatter)
    root = logging.getLogger()
    root.setLevel(os.environ.get("LOGLEVEL", "INFO"))
    root.addHandler(handler)
    #endregion


    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)

        # create an event to listen for stop requests on
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)


    # core logic of the service
    def SvcDoRun(self):
        import servicemanager

        rc = None

        # if the stop event hasn't been fired keep looping
        while rc != win32event.WAIT_OBJECT_0:
            #self.CopyPIDocsTree(self._srcpi, self._tgtpi)
            # self.ParallelCopy(self._srcpi, self._tgtpi)
            logging.info(self.isCopied(self._srcpi, self._tgtpi))

            # block for 5 seconds and listen for a stop event
            rc = win32event.WaitForSingleObject(self.hWaitStop, 5 * 60 * 1000)

        #logging.debug(self._svc_name_)


    # called when we're being shut down
    def SvcStop(self):
        # tell the SCM we're shutting down
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # fire the stop event
        win32event.SetEvent(self.hWaitStop)


    # copy_tree more useful for exe not windows service
    def CopyPIDocsTree(self, src, dst):
        _now = time.strftime("%H:%M")

        try:
            if _now > self._start and _now < self._end:
                copy_tree(src, dst)

                logging.info(self._today.strftime('%Y-%m-%d %H:%M:%S ') + self._svc_name_
                             + " file(s) copied from " + src + " to " + dst)

            # parallel ??

            # check if a file is copied?
        except Exception:
            logging.exception(self._today.strftime('%Y-%m-%d %H:%M:%S ') + self._svc_name_)


    def ParallelCopy(self, src, dst):
        '''
            allfiles = os.listdir(self._srcpi)
            # only list the directories and files, but not subdir, and files within
            allfiles = next(os.walk(self._srcpi))[2] # [] only?
        '''
        try:
            allfiles = self.isCopied(self._srcpi, self._tgtpi)

            logging.info(allfiles)
        except Exception:
            logging.exception(self._today.strftime('%Y-%m-%d %H:%M:%S ') + self._svc_name_)


    def getFilePaths(self, directory):
        """
        This function will generate the file names in a directory
        tree by walking the tree either top-down or bottom-up. For each
        directory in the tree rooted at directory top (including top itself),
        it yields a 3-tuple (dirpath, dirnames, filenames).
        """
        file_paths = []  # List which will store all of the full filepaths.

        # Walk the tree.
        try:
            for root, directories, files in os.walk(directory):
                for filename in files:
                    # Join the two strings in order to form the full filepath.
                    filepath = os.path.join(root, filename)
                    file_paths.append(filepath)  # Add it to the list.

                    logging.info(self._today.strftime('%Y-%m-%d %H:%M:%S ') + filepath)

            return file_paths
        except Exception:
            logging.exception(self._today.strftime('%Y-%m-%d %H:%M:%S ') + 'getFilePaths')


    def isCopied(self, src, dst):
        try:
            for root, directories, files in os.walk(src):
                for filename in files:
                    # srcpath = os.path.join(root, filename)
                    tgtpath = os.path.join(root.replace(src, dst), filename)
                    if os.path.exists(tgtpath):
                        return True
            logging.info(self._today.strftime('%Y-%m-%d %H:%M:%S ') + tgtpath)
        except Exception:
            logging.exception(self._today.strftime('%Y-%m-%d %H:%M:%S ') + 'isCopied')


if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(PySvcWCDoc)
