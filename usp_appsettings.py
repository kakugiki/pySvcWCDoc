import configparser
import os
import inspect

class appConfig:
    def getSetting(self, config, key):
        self._config = configparser.ConfigParser()
        self._path = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))

        self._config.read(self._path + "/app.ini")

        return self._config.get(config, key)
