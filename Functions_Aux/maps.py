import requests
import json
import datetime
from datetime import date
from datetime import datetime
from pathlib import Path
import time
from bs4 import BeautifulSoup

def Composicao_MAPS(data, cnpj):

    headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.54 Safari/537.36'}

    login_data = {
        'username': 'integracao_ger1_ot',
        'password': 'Desenvolvimento@ger1_9'}

    dados_consulta = {
        "idc_hf_0": "",
        "entity:entityInfo": json.dumps({"id": int(fundos['ID MAPS'][fundos[fundos.CNPJ == cnpj].index[0]]), "naturalKey":
            fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]],"persistableClass" : "jmine.biz.invest.operacional.carteira.domain.Carteira"}),
        "entity:autoComplete": fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]],
        "data": f"{data}",
        "idioma": "0",
        "pageCommands:elements:0:cell": "Pesquisar"}

    with requests.Session() as s:
        # Encontrando a autenticação para entrar no Pegasus Maps
        url = 'https://ot.cloud.mapsfinancial.com/pegasus/main'
        r = s.get(url, headers=headers)

        # Extraindo o url em html e usando o parser para dividir o código e depois obter o url do form de autenticação
        soup = BeautifulSoup(r.content, 'html.parser')
        x = soup.find('form')['action']

        # Fazendo o requests.post com os parametros do usuário (credenciais)
        r = s.post(x, data=login_data, headers=headers)

        # Entrnando na url conde segue para colocação dos parâmetros da consulta
        carteira = 'https://ot.cloud.mapsfinancial.com/pegasus/main/wicket/bookmarkable/web.pages.consulta.carteira.composicao.carteira.ComposicaoCarteiraPage'
        r = s.get(carteira, headers=headers)

        # fazendo o requests.post com os paramêtros da consulta
        consulta_carteira = 'https://ot.cloud.mapsfinancial.com/pegasus/main/wicket/page?1-1.IFormSubmitListener-mainForm'
        r = s.post(consulta_carteira, data=dados_consulta, headers=headers)

        if r.status_code == 200:

            if datetime.strptime(data, '%d/%m/%Y').date() == date.today():

                # Extraindo EXCEL em bytes para EXCEL
                excel_get = 'https://ot.cloud.mapsfinancial.com/pegasus/main/wicket/page?2-IResourceListener-report-reportTable-exporter_panel-elements-1-cell'
                r = s.get(excel_get, headers=headers).content
                with open(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/'
                          + 'COMPOSICAO_' + fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ', '_') + '_' +
                          datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y') + '_A' + '.xlsx', 'wb') as output:
                    output.write(r)
                    output.close()

                # Extraindo PDF em bytes para arquivo PDF
                pdf_get = 'https://ot.cloud.mapsfinancial.com/pegasus/main/wicket/page?2-IResourceListener-report-reportTable-exporter_panel-elements-0-cell'
                r = s.get(pdf_get, headers=headers).content
                with open(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/'
                          + 'COMPOSICAO_' + fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ', '_') + '_' +
                          datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y') + '_A' +'.pdf', 'wb') as output:
                    output.write(r)
                    output.close()
            else:

                # Extraindo EXCEL em bytes para EXCEL
                excel_get = 'https://ot.cloud.mapsfinancial.com/pegasus/main/wicket/page?2-IResourceListener-report-reportTable-exporter_panel-elements-1-cell'
                r = s.get(excel_get, headers=headers).content
                with open(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/'
                          + 'COMPOSICAO_' + fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ','_') + '_' +
                          datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y') + '.xlsx', 'wb') as output:
                    output.write(r)
                    output.close()

                # Extraindo PDF em bytes para arquivo PDF
                pdf_get = 'https://ot.cloud.mapsfinancial.com/pegasus/main/wicket/page?2-IResourceListener-report-reportTable-exporter_panel-elements-0-cell'
                r = s.get(pdf_get, headers=headers).content
                with open(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/'
                          + 'COMPOSICAO_' + fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ','_') + '_' +
                          datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y') + '.pdf', 'wb') as output:
                    output.write(r)
                    output.close()

        s.close()

def demonstrativo_caixa(data, cnpj):

    '''Obter o TOKEN'''

    url_token = "https://ot.cloud.mapsfinancial.com/auth/realms/mapscloud/protocol/openid-connect/token"

    dados = {"client_id": "integracao_ger1_ot", "client_secret": "34950f60-1f3c-491d-a04b-abbaeccfdd5a",
            "grant_type": "client_credentials"}

    response_token = requests.post(url_token, data=dados)

    token = json.loads(response_token.content)["access_token"]

    '''Baixar efetivamente o demonstrativo de caixa em PDF'''

    params = {'tipoArquivo': 'PDF', 'carteira': fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]],
              'dataInicial': datetime.strptime(data, '%d/%m/%Y').strftime('%Y-%m-%d'),
              'dataFinal': datetime.strptime(data, '%d/%m/%Y').strftime('%Y-%m-%d')}

    url_api = "https://ot.cloud.mapsfinancial.com/pegasus/api/demonstrativoCaixa"

    headers = {"Accept": "application/octet-stream", "Authorization": "Bearer {0}".format(token)}

    response = requests.get(url_api, headers=headers, params=params)

    filename = Path(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/' + 'DemonstrativoCaixa_'
                    + fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ', '_') + '_' +  datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y') + '.pdf')

    filename.write_bytes(response.content)

    '''Baixar efetivamente o demonstrativo de caixa em XLSX'''

    params = {'tipoArquivo': 'XLSX', 'carteira': fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]],
              'dataInicial': datetime.strptime(data, '%d/%m/%Y').strftime('%Y-%m-%d'),
              'dataFinal': datetime.strptime(data, '%d/%m/%Y').strftime('%Y-%m-%d')}

    url_api = "https://ot.cloud.mapsfinancial.com/pegasus/api/demonstrativoCaixa"

    headers = {"Accept": "application/octet-stream", "Authorization": "Bearer {0}".format(token)}

    response = requests.get(url_api, headers=headers, params=params)

    filename = Path(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/' + 'DemonstrativoCaixa_'
                    + fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ', '_') + '_' + datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y')  + '.xlsx')

    filename.write_bytes(response.content)

def xml_5(data, cnpj):

    '''Obter o TOKEN'''

    url_token = "https://ot.cloud.mapsfinancial.com/auth/realms/mapscloud/protocol/openid-connect/token"

    dados = {"client_id": "integracao_ger1_ot", "client_secret": "34950f60-1f3c-491d-a04b-abbaeccfdd5a",
            "grant_type": "client_credentials"}

    response_token = requests.post(url_token, data = dados)

    token = json.loads(response_token.content)["access_token"]

    '''Baixar efetivamente a carteira em XML'''

    params = {'carteira': fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]],
              'data': datetime.strptime(data, '%d/%m/%Y').strftime('%Y-%m-%d')}

    url_api = "https://ot.cloud.mapsfinancial.com/pegasus/api/posicaoXmlAnbima5/"

    headers = {"Accept": "application/xml;charset=UTF-8", "Authorization": "Bearer {0}".format(token)}

    response = requests.get(url_api, headers=headers, params=params)

    filename = Path(fundos['CAMINHO DA PASTA'][fundos[fundos.CNPJ == cnpj].index[0]] + '/Relatorios MAPS/'
                    +  fundos['MNEUMÔNICO MAPS'][fundos[fundos.CNPJ == cnpj].index[0]].replace(' ', '_') + '_' +  datetime.strptime(data, '%d/%m/%Y').strftime('%d%m%Y') + '.xml')

    filename.write_bytes(response.content)

def checar_status(id_fundo, data, token):

    idCarteira = id_fundo
    data_aux = datetime.strptime(data, '%d/%m/%Y').strftime('%Y-%m-%d')
    params = {'abertura': False, 'idioma': 'PORTUGUES'}
    url_api = "https://ot.cloud.mapsfinancial.com/pegasus/api/composicao/cabecalho/{0}/{1}".format(idCarteira, data_aux)
    headers = {"Accept": "application/json;charset=UTF-8", "Authorization": "Bearer {0}".format(token)}

    statusCota = json.loads(requests.get(url_api, headers = headers, params = params).content)['statusCota']

    if statusCota == 'Liberada':
        if fundos[fundos['ID MAPS'] == id_fundo]['coluna auxiliar'].isnull().values[0]:
            with open(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\Historico.csv', 'a') as file:
                file.write(f"{id_fundo},{fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]},{statusCota},{datetime.now().time().strftime('%H:%M:%S')}\n")
                file.close()
            with open(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\Historico_COMPLETO.csv','a') as file:
                file.write(f"{id_fundo},{fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]},{statusCota},{time.ctime()}\n")
                file.close()
            fundos.loc[fundos['ID MAPS'] == id_fundo, 'coluna auxiliar'] = datetime.now().time().strftime('%H:%M:%S')
            print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} LIBERADO")
            message = f"Fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} Liberado"
            # bot.send_message(chat_id=destination, text=message)

            return 'V'