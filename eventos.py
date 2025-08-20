import xmltodict, os, configparser


def GetPathDeDownload():
    # Cria uma instância do ConfigParser
    config = configparser.ConfigParser()
    # Lê o arquivo .ini
    config.read('config.ini')
    PathDeDownload = config['Dados-Para-Download-xmls']['path-dos-xmls']
    return PathDeDownload


