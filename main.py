from playwright.sync_api import sync_playwright
import calendar
from datetime import datetime
import pytz
from collections import namedtuple
from pyautogui import press
import asyncio
import configparser
import subprocess
import re
import xmltodict
import os, configparser


def GetPathDeDownload():
    # Cria uma instância do ConfigParser
    config = configparser.ConfigParser()
    # Lê o arquivo .ini
    config.read('config.ini')
    PathDeDownload = config['Dados-Para-Download-xmls']['path-dos-xmls']
    return PathDeDownload


def MoverEventos(): #No dia que o Migrador de XML mudar por algum motivo, e for preciso inutilizar os eventos, basta você descomentar a linha onde essa função é chamada
    PathDeDownload = GetPathDeDownload()
    PastaCnpj = os.listdir(PathDeDownload)
    for Cnpj in PastaCnpj:
        PastaXmls = os.listdir(f'{PathDeDownload}/{Cnpj}')
        for ArquivoXml in PastaXmls:
            if 'xml' in ArquivoXml.lower():
                xml = xmltodict.parse(open(f'{PathDeDownload}/{Cnpj}/{ArquivoXml}', 'rt').read())
                try:
                    if xml['procEventoNFe']:
                        try:
                            os.makedirs(f'{PathDeDownload}/{Cnpj}/Eventos')
                        except:
                            pass
                        os.rename(f'{PathDeDownload}/{Cnpj}/{ArquivoXml}', f'{PathDeDownload}/{Cnpj}/Eventos/{ArquivoXml}')
                except:
                    pass


def CriarArqIni():

    # Cria uma instância do ConfigParser
    config = configparser.ConfigParser()

    # Adiciona uma seção e define valores para chaves
    config['Dados-Para-Download-xmls'] = {
        'path-dos-xmls': '',
        'mes/ano-inicial': '',
        'mes/ano-final': ''
    }

    # Escreve os dados no arquivo INI
    with open('config.ini', 'wt+') as configfile:
        config.write(configfile)

    print("Arquivo .ini criado com sucesso!")

def GetDadosArqIni():
    # Cria uma instância do ConfigParser
    config = configparser.ConfigParser()

    # Lê o arquivo .ini
    config.read('config.ini')


    # Acessa uma seção e lê uma chave
    MesAnoInicial = config['Dados-Para-Download-xmls']['mes/ano-inicial']
    MesAnoFinal = config['Dados-Para-Download-xmls']['mes/ano-final']
    PathDeDownload = config['Dados-Para-Download-xmls']['path-dos-xmls']
    SelecionarCertificadoAutomaticamente = config['Dados-Para-Download-xmls']['selecionar-certificado-automaticamente']
    ModeloDoDocumento = config['Dados-Para-Download-xmls']['modelo-do-documento']
    AutoSelectCert = False
    if SelecionarCertificadoAutomaticamente == 'True':
        AutoSelectCert = True
    MyRecord = namedtuple('Record', ['MesAnoInicial', 'MesAnoFinal', 'PathDeDownload', 'AutoSelectCert', 'ModeloDoDocumento'])
    return MyRecord(MesAnoInicial=MesAnoInicial, MesAnoFinal=MesAnoFinal, PathDeDownload=PathDeDownload, AutoSelectCert=AutoSelectCert, ModeloDoDocumento=ModeloDoDocumento)

def GetMes(AMes: str):
    pos = AMes.find('/')
    return int(AMes[:pos])

def GetAno(AAno: str):
    pos = AAno.find('/')
    return int(AAno[pos+1:])


import subprocess
import re

def PegarCnpjManual(CertificadoNome):
    try:
        # Listar certificados instalados
        resultado = subprocess.check_output(
            'certutil -store -user My', 
            shell=True, 
            text=True
        )

        # Encontra a seção correspondente ao certificado desejado
        certificados = resultado.split("\n")
        capturando = False
        certificado_dados = ""

        for linha in certificados:
            if CertificadoNome.lower() in linha.lower():
                capturando = True  # Começa a capturar as informações do certificado
            if capturando:
                certificado_dados += linha + "\n"
            if capturando and "Certificado" in linha:
                break  # Para a captura quando termina o bloco do certificado

        # Procurar por CNPJ formatado
        cnpj = re.search(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", certificado_dados)
        if cnpj:
            return cnpj.group()

        # Procurar por CNPJ sem formatação
        cnpj_sem_formato = re.search(r"\b\d{14}\b", certificado_dados)
        if cnpj_sem_formato:
            cnpj = cnpj_sem_formato.group()
            return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}".replace(".", "").replace("/", "").replace("-", "")
        
        print("Nenhum CNPJ encontrado no certificado")
        return None
    except Exception as e:
        print(f"Erro ao buscar CNPJ: {e}")
        return None


def PegarCnpjAuto():
    try:
        # Listar certificados instalados
        resultado = subprocess.check_output(
            'certutil -store -user My', 
            shell=True, 
            text=True
        )
        
        # Procurar por padrões de CNPJ (14 dígitos) no resultado
        cnpjs = re.findall(r'\b\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}\b', resultado)
        if cnpjs:
            return cnpjs[0]  # Retorna o primeiro CNPJ encontrado
        else:
            # Procurar por padrão numérico que pode ser um CNPJ sem formatação
            cnpjs_sem_formato = re.findall(r'\b\d{14}\b', resultado)
            if cnpjs_sem_formato:
                cnpj = cnpjs_sem_formato[0]
                # Formatar o CNPJ
                return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}".replace(".", "").replace("/", "").replace("-", "")
            
        print("Nenhum CNPJ encontrado nos certificados")
        return None
    except Exception as e:
        print(f"Erro ao buscar CNPJ: {e}")
        return None
    

async def EsperarParaApertarTab():
    # Dá um tempo para o popup de certificado aparecer
    await asyncio.sleep(2)
    # Pressiona tab algumas vezes e depois Enter para selecionar o certificado
    for c in range(0, 3):
        press('tab')
        await asyncio.sleep(0.5)
    press('enter')

def calcular_diferenca_em_meses(mes_ano_inicial, mes_ano_final):
    # Calcula a diferença entre os meses
    ano_inicial = GetAno(mes_ano_inicial)
    mes_inicial = GetMes(mes_ano_inicial)
    ano_final = GetAno(mes_ano_final)
    mes_final = GetMes(mes_ano_final)

    # Calculando a diferença em meses
    diff_anos = ano_final - ano_inicial
    diff_meses = mes_final - mes_inicial
    return diff_anos * 12 + diff_meses

def GetLastDay(mes_ano):
    """
    Recebe uma string no formato 'MM/YYYY' e retorna o último dia daquele mês.
    Exemplo: '03/2025' retorna 31 (porque março tem 31 dias)
    """
    try:
        # Separar mês e ano
        mes, ano = mes_ano.split('/')
        mes = int(mes)
        ano = int(ano)
        # Usar calendar para obter o último dia do mês
        ultimo_dia = calendar.monthrange(ano, mes)[1]
        if ultimo_dia < 10:
            ultimo_dia = f'0{ultimo_dia}'
        return ultimo_dia
    except Exception as e:
        print(f"Erro ao calcular último dia: {e}")
        return None
    

def AdicionarLog(AText=''):
    with open('log.txt', 'at') as Arquivo:
        Arquivo.write(f'{AText}\n')

def ZerarLog():
    with open('log.txt', 'wt+') as Arquivo:
        pass
def CriarLogDev():
    with open('dev.txt', 'wt+') as Arquivo:
        pass
def AdicionarLogDev(AText=''):
    with open('dev.txt', 'at') as Arquivo:
        Arquivo.write(f'{AText}\n')


def FazerPesquisa(Site, Data, ModeloDoDocumento: str):
    AdicionarLogDev('Entrou na func Fazer Pesquisa')
    if not Site.locator('//*[@id="cmpDataInicial"]').is_visible():
        AdicionarLogDev('func Fazer Pesquisa Caiu no 1 if, tentando clicar realizar cliques')
        Site.locator('//*[@id="content"]/div[2]/div/a/button').click()
        AdicionarLogDev('func Fazer Pesquisa Caiu no 1 if, conseguiu realizar cliques')
    AdicionarLogDev('func Fazer Pesquisa T1')
    Site.locator('//*[@id="cmpDataInicial"]').click()
    AdicionarLogDev('func Fazer Pesquisa T2')
    if Site.locator('//*[@id="cmpDataInicial"]').is_visible():
        AdicionarLogDev('func Fazer Pesquisa T3')
        Site.locator('//*[@id="cmpDataInicial"]').type(f'01/{Data}', delay=50)
        AdicionarLogDev('func Fazer Pesquisa T4')
    Site.locator('//*[@id="cmpDataFinal"]').click()
    AdicionarLogDev('func Fazer Pesquisa T5')
    if Site.locator('//*[@id="cmpDataFinal"]').is_visible():
        AdicionarLogDev('func Fazer Pesquisa T6')
        Site.locator('//*[@id="cmpDataFinal"]').type(f'{GetLastDay(Data)}/{Data}', delay=50)
        AdicionarLogDev('func Fazer Pesquisa T7')

    AdicionarLogDev('func Fazer Pesquisa T8')
    Site.locator('//*[@id="filtro"]').click()
    AdicionarLogDev('func Fazer Pesquisa T9')
    Site.locator('//*[@id="cmpModelo"]').select_option(value=ModeloDoDocumento)
    AdicionarLogDev('func Fazer Pesquisa T10')

    Site.locator('//*[@id="btnPesquisar"]').click()
    AdicionarLogDev('func Fazer Pesquisa T11')
    Site.wait_for_event('load')
    AdicionarLogDev('func Fazer Pesquisa T12')
    Status = ''
    AdicionarLogDev('func Fazer Pesquisa T13')
    if Site.locator('//*[@id="message-containter"]/div').is_visible():
        AdicionarLogDev('func Fazer Pesquisa T14')
        Status = Site.locator('//*[@id="message-containter"]/div').inner_text()
        AdicionarLogDev('func Fazer Pesquisa T15')
    # while True:
    #     try:
    #         url = Site.evaluate("window.location.href")
    #         if 'ErrorMsg' in url:
    #             Status = Site.locator('//*[@id="message-containter"]/div').inner_text()
    #             break
    #         elif 'recaptcha' in url:
    #             break
    #     except:
    #         continue
    return Status


try:
    ZerarLog()
    CriarLogDev()
    with sync_playwright() as pw:
        DadosArqIni = GetDadosArqIni()
        if DadosArqIni.AutoSelectCert:
            CNPJ = PegarCnpjAuto()
        else:
            CNPJ = PegarCnpjManual(str(input('Nome/CPF/CNPJ do Certificado: ')).replace('.', '').replace('/', '').replace('-', ''))
        Navegador = pw.chromium.launch(headless=False, channel="msedge", downloads_path=DadosArqIni.PathDeDownload+f'/{CNPJ}')
        Site = Navegador.new_page()
        Site.set_default_navigation_timeout = 120000
        Site.set_default_timeout = 120000
        if DadosArqIni.AutoSelectCert:
            asyncio.create_task(EsperarParaApertarTab())
        Site.goto('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica/principal')
        Site.locator('//*[@id="panel"]/div[2]/div[3]/div/a/button').click()
        CnpjSite = Site.locator('//*[@id="cmpCnpj"]').inner_text()
        if CNPJ != CnpjSite:
            Navegador.close()
            #input("CPF/CNPJ do certificado não confere com o CNPJ do site\n")
        else:
            AnoInical = GetAno(DadosArqIni.MesAnoInicial)
            MesInicial = GetMes(DadosArqIni.MesAnoInicial)
            AnoFinal = GetAno(DadosArqIni.MesAnoFinal)
            MesFinal =  GetMes(DadosArqIni.MesAnoFinal)
            QuantidadeDeMeses = calcular_diferenca_em_meses(DadosArqIni.MesAnoInicial, DadosArqIni.MesAnoFinal)+1
            PrimeiroLoop = True
            for c in range(0, QuantidadeDeMeses):
                EncontrouDocoumentos = True
                DataLog = datetime.today().astimezone(pytz.timezone('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M:%S')
                PrimeiroLoop = False
                Data = f'0{MesInicial}/{AnoInical}' if MesInicial < 10 else f'{MesInicial}/{AnoInical}'
                AdicionarLogDev('func main T1')
                if not Site.locator('//*[@id="cmpDataInicial"]').is_visible():
                    AdicionarLogDev('func main T2')
                    try:
                        AdicionarLogDev('func main try T1')
                        Site.locator('//*[@id="content"]/div[2]/div/a/button').click()
                        AdicionarLogDev('func main try T1 ok!')
                    except:
                        AdicionarLogDev('func main try except T1')
                        Site.goto('https://nfeweb.sefaz.go.gov.br/nfeweb/sites/nfe/consulta-publica/principal')
                        AdicionarLogDev('func main try except T2')
                    AdicionarLogDev('func main T3')
                AdicionarLogDev('func main T4')
                CnpjSite = Site.locator('//*[@id="cmpCnpj"]').inner_text()
                AdicionarLogDev('func main T5')
                if CNPJ != CnpjSite:
                    Navegador.close()
                    AdicionarLog(f'{DataLog} - O CPF/CNPJ do certificado não confere com o CPF/CNPJ do site, caso queira continuar de onde parou configure o mes/ano-inicial = {Data}')
                    exit()
                #Data = '02/2000'
                Erro1 = FazerPesquisa(Site=Site, Data=Data, ModeloDoDocumento=DadosArqIni.ModeloDoDocumento)
                if Erro1 != '':         
                    Erro2 = FazerPesquisa(Site=Site, Data=  Data, ModeloDoDocumento=DadosArqIni.ModeloDoDocumento)
                    if Erro2 != '':
                        Erro3 = FazerPesquisa(Site=Site, Data=Data, ModeloDoDocumento=DadosArqIni.ModeloDoDocumento)
                        if Erro3 != '':
                            AdicionarLog(f'{DataLog} - Erro ao pesquisar a data {Data}, mês sem resultados para o CNPJ {CNPJ}!')
                            EncontrouDocoumentos = False
                if EncontrouDocoumentos:
                    AdicionarLogDev('func main T6')
                    Site.locator('//*[@id="content"]/div[2]/div/button').click()  
                    AdicionarLogDev('func main T7')
                    Site.locator('//*[@id="dnwld-all-btn-ok"]').click() 
                    AdicionarLogDev('func main T8')  
                    while True:
                        try:
                            AdicionarLogDev('func main, while True T1')
                            Site.locator('//*[@id="myModalLabel"]').wait_for(timeout=300000)
                            AdicionarLogDev('func main, while True T2')
                            if 'Concluído' in Site.locator('//*[@id="myModalLabel"]').inner_text():
                                AdicionarLogDev('func main, while True T3')
                                Site.locator('//*[@id="content"]/div[2]/div/a/button').click()
                                AdicionarLogDev('func main, while True T4')
                                AdicionarLog(f'{DataLog} - {Data} baixado com sucesso!')
                                AdicionarLogDev('func main, while True T5')
                                break
                        except:
                            continue
                if not PrimeiroLoop:
                    if MesInicial == MesFinal and AnoInical == AnoFinal:
                        Navegador.close()
                        input("Fim das pesquisas\n")
                if MesInicial < 13:
                    MesInicial += 1
                if MesInicial == 13:
                    MesInicial = 1
                    AnoInical += 1
except Exception as E:
    input(f'Ocorreu o seguinte erro: {E}\n')
else:
    #MoverEventos()#Descomente quando for necessário mover os eventos
    pass
