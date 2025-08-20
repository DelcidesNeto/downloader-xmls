import os, configparser
def GetPathDeDownload():
    # Cria uma instância do ConfigParser
    config = configparser.ConfigParser()

    # Lê o arquivo .ini
    config.read('config.ini')
    PathDeDownload = config['Dados-Para-Download-xmls']['path-dos-xmls']
    return PathDeDownload
PathDeDownload = GetPathDeDownload()
PastaCnpj = os.listdir(PathDeDownload)
try:
    for Cnpj in PastaCnpj:
        xmls = os.listdir(f'{PathDeDownload}/{Cnpj}')
        for xml in xmls:
            os.rename(f'{PathDeDownload}/{Cnpj}/{xml}', f'{PathDeDownload}/{Cnpj}/{xml}.zip')
except:
    pass
input('Todos os Arquvos foram renomeados com Sucesso!')