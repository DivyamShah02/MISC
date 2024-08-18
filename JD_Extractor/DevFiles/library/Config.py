import configparser


class Config:

    def __init__(self, filename):
        config = configparser.ConfigParser()
        config.read(filename)

        for section in config.sections():
            setattr(self, section, ConfigSection(config[section]))


class ConfigSection:

    def __init__(self, section):
        for key, value in section.items():
            setattr(self, key, value)
