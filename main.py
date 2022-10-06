from Functions_Aux.feriado import *
from Functions_Aux.mail_sending import *
from Functions_Aux.maps import *

aux_anbima = 0
conn = mysql.connector.connect(user = 'fidc_consulta', password = 'FSenFDS65#', host = '172.16.0.36', database = 'oliveir_scot', auth_plugin = 'mysql_native_password')

#Inicio do código para checagem automática

#Criar CSV do Histórico do dia (apaga o do dia anterior)

with open(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\Historico.csv', 'w') as file:
    file.write("id fundo,mneumônico maps,status,horário\n")
    file.close()

fundos = pd.read_excel(r"\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Fundos.xlsx", dtype='string')
fundos.CNPJ = fundos.CNPJ.str.replace('.', '', regex=False)
fundos.CNPJ = fundos.CNPJ.str.replace('/', '', regex=False)
fundos.CNPJ = fundos.CNPJ.str.replace('-', '', regex=False)
fundos['coluna auxiliar'] = np.nan

agora = datetime.now().time()

print('Processo Iniciado por horário')

bot.send_message(chat_id=destination, text = "O robô foi iniciado no dia " + date.today().strftime('%d/%m/%Y'))

# Amortizações de D+1

if eh_feriado((date.today() + BDay(1)).strftime('%Y-%m-%d')) == 1:
    if eh_feriado((date.today() + BDay(2)).strftime('%Y-%m-%d')) == 1:
        D1 = (date.today() + BDay(3)).strftime('%Y-%m-%d')
    else:
        D1 = (date.today() + BDay(2)).strftime('%Y-%m-%d')
else:
    D1 = (date.today() + BDay(1)).strftime('%Y-%m-%d')

query1 = "SELECT cod_operacao, tit, data, valor, percentual, descricao FROM vw_amort_fds WHERE data = " + "'" + D1 + "'"
cursor = conn.cursor(buffered=True)
cursor.execute(query1)
Amortizacoes = pd.DataFrame(cursor.fetchall())

cotas = pd.read_excel(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\cotas_oper.xlsx')

# Checando quais amortizações são de FIDC's

if Amortizacoes.shape[0] != 0:
    auxamort = Amortizacoes[Amortizacoes[1].isin(cotas['Cota'])]
    if auxamort.shape[0] != 0:
        auxamort = auxamort.rename(columns = {1: 'Cota'})
        # Novo dataframe com nomes das cotas
        novo = auxamort.merge(cotas, how = 'left', on = 'Cota')

        bot.send_message(chat_id=destination, text="Amortizações de D+1 (" + datetime.strptime(D1, '%Y-%m-%d').strftime('%d/%m/%Y') + "): ")
        message1 = ""
        time.sleep(3)
        for amort in novo['Cota - Fundo']:
            message1 = message1 + amort + ';' + '\n'

        bot.send_message(chat_id=destination, text= message1)

    else:
        bot.send_message(chat_id=destination, text="Sem amortizaçoes para o dia")
else:
    bot.send_message(chat_id=destination,
                     text="Amortizações de D+1 (" + datetime.strptime(D1, '%Y-%m-%d').strftime('%d/%m/%Y') + "): ")
    time.sleep(3)
    bot.send_message(chat_id=destination, text="Sem amortizaçoes para o dia")

# Amortizações de D+2

if eh_feriado((date.today() + BDay(2)).strftime('%Y-%m-%d')) == 1:
    if eh_feriado((date.today() + BDay(3)).strftime('%Y-%m-%d')) == 1:
        D2 = (date.today() + BDay(4)).strftime('%Y-%m-%d')
    else:
        D2 = (date.today() + BDay(3)).strftime('%Y-%m-%d')
else:
    D2 = (date.today() + BDay(2)).strftime('%Y-%m-%d')

query2 = "SELECT cod_operacao, tit, data, valor, percentual, descricao FROM vw_amort_fds WHERE data = " + "'" + D2 + "'"
cursor = conn.cursor(buffered=True)
cursor.execute(query2)
Amortizacoes = pd.DataFrame(cursor.fetchall())

# Checando quais amortizações são de FIDC's

if Amortizacoes.shape[0] != 0:
    auxamort = Amortizacoes[Amortizacoes[1].isin(cotas['Cota'])]
    if auxamort.shape[0] != 0:
        auxamort = auxamort.rename(columns = {1: 'Cota'})
        # Novo dataframe com nomes das cotas
        novo = auxamort.merge(cotas, how = 'left', on = 'Cota')

        message2 = ""

        bot.send_message(chat_id=destination, text="Amortizações de D+2 (" + datetime.strptime(D2, '%Y-%m-%d').strftime('%d/%m/%Y') + "): ")
        time.sleep(3)
        for amort in novo['Cota - Fundo']:
            message2 = message2 + amort + ';' + '\n'

        bot.send_message(chat_id=destination, text= message2)

    else:
        bot.send_message(chat_id=destination, text="Sem amortizaçoes para o dia")
else:
    bot.send_message(chat_id=destination,
                     text="Amortizações de D+2 (" + datetime.strptime(D2, '%Y-%m-%d').strftime('%d/%m/%Y') + "): ")
    time.sleep(3)
    bot.send_message(chat_id=destination, text="Sem amortizaçoes para o dia")


while agora > datetime.time(datetime.strptime('07:45:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('22:00:00', '%H:%M:%S')):

    for id_fundo in fundos['ID MAPS']:

        # Obter o TOKEN de autorização
        try:
            url_token = "https://ot.cloud.mapsfinancial.com/auth/realms/mapscloud/protocol/openid-connect/token"

            credenciais = {"client_id": "integracao_ger1_ot", "client_secret": "34950f60-1f3c-491d-a04b-abbaeccfdd5a",
                       "grant_type": "client_credentials"}

            response_token = requests.post(url_token, data=credenciais)

            token = json.loads(response_token.content)["access_token"]
        except Exception:
            pass
        try:
            if fundos[fundos['ID MAPS'] == id_fundo]['CALCULO COTA SUB'].iloc[0][0] == 'A':
                try:
                    data = date.today().strftime('%d/%m/%Y')

                    if eh_feriado((date.today() - BDay(1)).strftime('%d/%m/%Y')) == 1:
                        if eh_feriado((date.today() - BDay(2)).strftime('%d/%m/%Y')) == 1:
                            d1 = (date.today() - BDay(3)).strftime('%d/%m/%Y')
                        else:
                            d1 = (date.today() - BDay(2)).strftime('%d/%m/%Y')
                    else:
                        d1 = (date.today() - BDay(1)).strftime('%d/%m/%Y')

                    if checar_status(id_fundo, data, token) == 'V':

                        Composicao_MAPS(data, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        Composicao_MAPS(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} CARTEIRA BAIXADA EM D0 E D-1")
                        demonstrativo_caixa(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} DEMONSTRATIVO BAIXADO EM D-1")
                        xml_5(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} XML BAIXADO EM D-1")

                except Exception:
                    pass

            elif fundos[fundos['ID MAPS'] == id_fundo]['CALCULO COTA SUB'].iloc[0][0] == 'F' and fundos[fundos['ID MAPS'] == id_fundo]['REF CALC COTA SUB'].iloc[0] == 'D0':
                try:
                    data = date.today().strftime('%d/%m/%Y')

                    if checar_status(id_fundo, data, token) == 'V':
                        Composicao_MAPS(data, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} CARTEIRA BAIXADA EM D-1")
                        demonstrativo_caixa(data, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} DEMONSTRATIVO BAIXADO EM D-1")
                        xml_5(data, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} XML BAIXADO EM D-1")
                except Exception:
                    pass

            elif fundos[fundos['ID MAPS'] == id_fundo]['CALCULO COTA SUB'].iloc[0][0] == 'F' and fundos[fundos['ID MAPS'] == id_fundo]['REF CALC COTA SUB'].iloc[0] == 'D-1':
                try:

                    if eh_feriado((date.today() - BDay(1)).strftime('%d/%m/%Y')) == 1:
                        if eh_feriado((date.today() - BDay(2)).strftime('%d/%m/%Y')) == 1:
                            d1 = (date.today() - BDay(3)).strftime('%d/%m/%Y')
                        else:
                            d1 = (date.today() - BDay(2)).strftime('%d/%m/%Y')
                    else:
                        d1 = (date.today() - BDay(1)).strftime('%d/%m/%Y')

                    if checar_status(id_fundo, d1, token) == 'V':
                        Composicao_MAPS(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} CARTEIRA BAIXADA EM D-1")
                        demonstrativo_caixa(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} DEMONSTRATIVO BAIXADO EM D-1")
                        xml_5(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} XML BAIXADO EM D-1")
                except Exception:
                    pass
            elif fundos[fundos['ID MAPS'] == id_fundo]['CALCULO COTA SUB'].iloc[0][0] == 'F' and fundos[fundos['ID MAPS'] == id_fundo]['REF CALC COTA SUB'].iloc[0] == 'D-2':
                try:

                    if eh_feriado((date.today() - BDay(2)).strftime('%d/%m/%Y')) == 1:
                        if eh_feriado((date.today() - BDay(3)).strftime('%d/%m/%Y')) == 1:
                            d1 = (date.today() - BDay(4)).strftime('%d/%m/%Y')
                        else:
                            d1 = (date.today() - BDay(3)).strftime('%d/%m/%Y')
                    else:
                        d1 = (date.today() - BDay(2)).strftime('%d/%m/%Y')

                    if checar_status(id_fundo, d1, token) == 'V':
                        Composicao_MAPS(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} CARTEIRA BAIXADA EM D-1")
                        demonstrativo_caixa(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} DEMONSTRATIVO BAIXADO EM D-1")
                        xml_5(d1, fundos[fundos['ID MAPS'] == id_fundo]['CNPJ'].iloc[0])
                        print(f"fundo {fundos[fundos['ID MAPS'] == id_fundo]['MNEUMÔNICO MAPS'].iloc[0]} XML BAIXADO EM D-1")
                except Exception:
                    pass

        except Exception:
            continue
    try:

        if agora > datetime.time(datetime.strptime('11:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('14:00:00', '%H:%M:%S')):
            fundos.to_excel(r'\\scot\h\FUNDOS\GER1\Colaboradores - GER1\Luiz Araujo\Log\Log.xlsx')
        if agora > datetime.time(datetime.strptime('15:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('16:00:00', '%H:%M:%S')):
            fundos.to_excel(r'\\scot\h\FUNDOS\GER1\Colaboradores - GER1\Luiz Araujo\Log\Log.xlsx')
        if agora > datetime.time(datetime.strptime('18:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('19:30:00', '%H:%M:%S')):
            fundos.to_excel(r'\\scot\h\FUNDOS\GER1\Colaboradores - GER1\Luiz Araujo\Log\Log.xlsx')

        # PROJEÇÃO IPCA ANBIMA

        if agora > datetime.time(datetime.strptime('11:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('12:00:00', '%H:%M:%S')) and aux_anbima != 1:

            url_token = "https://api.anbima.com.br/oauth/access-token"

            credentials = "0RpMSlIVJKuI" + ':' + "XFSonYn3HpoI"
            message_bytes = credentials.encode('ascii')
            base64_bytes = base64.b64encode(message_bytes)
            base64_credentials = base64_bytes.decode('ascii')

            header = {'Content-Type': 'application/json', "Authorization": "Basic " + base64_credentials}

            dados = {"grant_type": "client_credentials"}
            dados = json.dumps(dados)

            response_token = requests.post(url_token, headers=header, data=dados)

            token = json.loads(response_token.content)["access_token"]

            url_api = "https://api.anbima.com.br/feed/precos-indices/v1/titulos-publicos/projecoes"

            headers = {"Accept": "application/json", "client_id": "0RpMSlIVJKuI", "access_token": token}

            # params = {'mes': '09', 'ano': '2019'}

            Response = json.loads(requests.get(url_api, headers=headers).content)
            Response = pd.DataFrame(Response)
            Response.drop(index=Response[Response.indice == 'IGP-M'].index, inplace=True)

            Response.drop(index=Response[Response.tipo_projecao == 'PROJEÇÕES PARA O MÊS POSTERIOR'].index, inplace=True)

            aux = pd.read_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', dtype='string')

            for n in Response['data_validade']:
                if aux[aux['data_validade'] == n].shape[0] == 0:
                    aux = aux.append(Response[Response['data_validade'] == n])

            aux.to_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', index=None)

            aux_anbima += 1

        if agora > datetime.time(datetime.strptime('13:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('14:00:00', '%H:%M:%S')) and aux_anbima != 2:

            url_token = "https://api.anbima.com.br/oauth/access-token"

            credentials = "0RpMSlIVJKuI" + ':' + "XFSonYn3HpoI"
            message_bytes = credentials.encode('ascii')
            base64_bytes = base64.b64encode(message_bytes)
            base64_credentials = base64_bytes.decode('ascii')

            header = {'Content-Type': 'application/json', "Authorization": "Basic " + base64_credentials}

            dados = {"grant_type": "client_credentials"}
            dados = json.dumps(dados)

            response_token = requests.post(url_token, headers=header, data=dados)

            token = json.loads(response_token.content)["access_token"]

            url_api = "https://api.anbima.com.br/feed/precos-indices/v1/titulos-publicos/projecoes"

            headers = {"Accept": "application/json", "client_id": "0RpMSlIVJKuI", "access_token": token}

            # params = {'mes': '09', 'ano': '2019'}

            Response = json.loads(requests.get(url_api, headers=headers).content)
            Response = pd.DataFrame(Response)
            Response.drop(index=Response[Response.indice == 'IGP-M'].index, inplace=True)

            Response.drop(index=Response[Response.tipo_projecao == 'PROJEÇÕES PARA O MÊS POSTERIOR'].index, inplace=True)

            aux = pd.read_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', dtype='string')

            for n in Response['data_validade']:
                if aux[aux['data_validade'] == n].shape[0] == 0:
                    aux = aux.append(Response[Response['data_validade'] == n])

            aux.to_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', index=None)

            aux_anbima += 1

        if agora > datetime.time(datetime.strptime('16:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('17:00:00', '%H:%M:%S')) and aux_anbima != 3:
            url_token = "https://api.anbima.com.br/oauth/access-token"

            credentials = "0RpMSlIVJKuI" + ':' + "XFSonYn3HpoI"
            message_bytes = credentials.encode('ascii')
            base64_bytes = base64.b64encode(message_bytes)
            base64_credentials = base64_bytes.decode('ascii')

            header = {'Content-Type': 'application/json', "Authorization": "Basic " + base64_credentials}

            dados = {"grant_type": "client_credentials"}
            dados = json.dumps(dados)

            response_token = requests.post(url_token, headers=header, data=dados)

            token = json.loads(response_token.content)["access_token"]

            url_api = "https://api.anbima.com.br/feed/precos-indices/v1/titulos-publicos/projecoes"

            headers = {"Accept": "application/json", "client_id": "0RpMSlIVJKuI", "access_token": token}

            # params = {'mes': '09', 'ano': '2019'}

            Response = json.loads(requests.get(url_api, headers=headers).content)
            Response = pd.DataFrame(Response)
            Response.drop(index=Response[Response.indice == 'IGP-M'].index, inplace=True)

            Response.drop(index=Response[Response.tipo_projecao == 'PROJEÇÕES PARA O MÊS POSTERIOR'].index, inplace=True)

            aux = pd.read_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', dtype='string')

            for n in Response['data_validade']:
                if aux[aux['data_validade'] == n].shape[0] == 0:
                    aux = aux.append(Response[Response['data_validade'] == n])

            aux.to_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', index=None)

            aux_anbima += 1

        if agora > datetime.time(datetime.strptime('18:00:00', '%H:%M:%S')) and agora < datetime.time(datetime.strptime('19:00:00', '%H:%M:%S')) and aux_anbima != 4:

            url_token = "https://api.anbima.com.br/oauth/access-token"

            credentials = "0RpMSlIVJKuI" + ':' + "XFSonYn3HpoI"
            message_bytes = credentials.encode('ascii')
            base64_bytes = base64.b64encode(message_bytes)
            base64_credentials = base64_bytes.decode('ascii')

            header = {'Content-Type': 'application/json', "Authorization": "Basic " + base64_credentials}

            dados = {"grant_type": "client_credentials"}
            dados = json.dumps(dados)

            response_token = requests.post(url_token, headers=header, data=dados)

            token = json.loads(response_token.content)["access_token"]

            url_api = "https://api.anbima.com.br/feed/precos-indices/v1/titulos-publicos/projecoes"

            headers = {"Accept": "application/json", "client_id": "0RpMSlIVJKuI", "access_token": token}

            # params = {'mes': '09', 'ano': '2019'}

            Response = json.loads(requests.get(url_api, headers=headers).content)
            Response = pd.DataFrame(Response)
            Response.drop(index=Response[Response.indice == 'IGP-M'].index, inplace=True)

            Response.drop(index=Response[Response.tipo_projecao == 'PROJEÇÕES PARA O MÊS POSTERIOR'].index,inplace=True)

            aux = pd.read_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', dtype='string')

            for n in Response['data_validade']:
                if aux[aux['data_validade'] == n].shape[0] == 0:
                    aux = aux.append(Response[Response['data_validade'] == n])

            aux.to_csv(r'H:\FUNDOS\GER1\Requisição de Baixa de Arquivos\Projecoes_Anbima_Robo\Anbima.csv', sep=';', decimal=',', encoding='ISO-8859-1', index=None)

            aux_anbima += 1

    except Exception:
        continue

    time.sleep(180)
    agora = datetime.now().time()

# Enviar email do relatório (e salvar relatório de atrasados)

try:

    atrasados_aux = pd.read_csv(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\Historico.csv', encoding = 'ISO-8859-1', dtype = 'string').rename(columns = {'id fundo': 'ID MAPS'})
    fundos = pd.read_excel(r"\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Fundos.xlsx", dtype='string')
    fundos.index = fundos['ID MAPS']
    atrasados = atrasados_aux.merge(fundos['HORÁRIO PREVISTO'].to_frame(), how = 'left', on = 'ID MAPS')
    atrasados['horário em segundos'] = pd.to_datetime(atrasados['horário'], format = '%H:%M:%S').dt.hour*3600 + \
                        pd.to_datetime(atrasados['horário'], format = '%H:%M:%S').dt.minute*60 + \
                                    pd.to_datetime(atrasados['horário'], format = '%H:%M:%S').dt.second
    atrasados['HORÁRIO PREVISTO em Segundos'] = pd.to_datetime(atrasados['HORÁRIO PREVISTO'], format = '%H:%M:%S').dt.hour*3600 + \
                                    pd.to_datetime(atrasados['HORÁRIO PREVISTO'], format = '%H:%M:%S').dt.minute*60 + \
                                    pd.to_datetime(atrasados['HORÁRIO PREVISTO'], format = '%H:%M:%S').dt.second
    atrasados['Diferença'] = atrasados['horário em segundos'] - atrasados['HORÁRIO PREVISTO em Segundos']

    atrasados = atrasados.rename(columns = {'mneumônico maps': 'Fundo', 'horário': 'Horário', 'HORÁRIO PREVISTO': 'Horário Previsto', 'status': 'Status'})

    Maior_Atrasado = atrasados[atrasados['Diferença'] == atrasados['Diferença'].nlargest(3).iloc[0]]['Fundo'].iloc[0]
    Maior_Atrasado_Tempo = round(atrasados['Diferença'].nlargest(3).iloc[0] / 3600, 2)
    Segundo_Maior_Atrasado = atrasados[atrasados['Diferença'] == atrasados['Diferença'].nlargest(3).iloc[1]]['Fundo'].iloc[0]
    Segundo_Maior_Atrasado_Tempo = round(atrasados['Diferença'].nlargest(3).iloc[1] / 3600, 2)
    Terceiro_Maior_Atrasado = atrasados[atrasados['Diferença'] == atrasados['Diferença'].nlargest(3).iloc[2]]['Fundo'].iloc[0]
    Terceiro_Maior_Atrasado_Tempo = round(atrasados['Diferença'].nlargest(3).iloc[2] / 3600, 2)
    Percentual_Atraso = round((atrasados[atrasados['Diferença'] > 0].shape[0] / fundos.shape[0])*100, 2)
    Percentual_nao_liberado = round(100 - (atrasados.shape[0] / fundos.shape[0])*100, 2)
    Percentual_Sem_Atraso = round((atrasados[atrasados['Diferença'] <= 0].shape[0] / fundos.shape[0])*100, 2)

    auxiliar = atrasados[['Fundo', 'Status', 'Horário', 'Horário Previsto', 'Diferença']]
    auxiliar = auxiliar[auxiliar['Diferença'] > 0]
    auxiliar = auxiliar.sort_values(by = ['Diferença'], ascending = False)
    auxiliar = auxiliar.drop(columns = 'Diferença')
    auxiliar['Status'] = 'Liberado com Atraso'
    auxiliar.index = range(auxiliar.shape[0])

    #Incluindo os não liberados

    for aux in fundos['MNEUMÔNICO MAPS']:
        aux3 = 'NÃO TEM'
        for aux2 in atrasados['Fundo']:
            if aux == aux2:
                aux3 = 'TEM'
        if aux3 == 'NÃO TEM':
            auxiliar.loc[auxiliar.shape[0], 'Fundo'] = aux

    auxiliar['Status'].fillna('Não Liberado', inplace = True)
    auxiliar['Horário'].fillna('Não Liberado', inplace = True)
    auxiliar['Horário Previsto'].fillna('Não Liberado', inplace = True)

    auxiliar.to_excel(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\HISTORICO_EMAIL\Historico_Email_' + datetime.strftime(date.today(), '%d%m%Y') + '.xlsx', index = None)

#    enviar_email_com_anexo(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\auxiliar.xlsx', 'luiz.araujo@oliveiratrust.com.br, lucas.mattos@oliveiratrust.com.br, alan.najman@oliveiratrust.com.br, ger1.fundos@oliveiratrust.com.br, lucas.torres@oliveiratrust.com.br, leonardo.goulart@oliveiratrust.com.br, raphael.morgado@oliveiratrust.com.br',
#                Percentual_Atraso, Percentual_nao_liberado, Percentual_Sem_Atraso, Maior_Atrasado, Maior_Atrasado_Tempo,
#                Segundo_Maior_Atrasado, Segundo_Maior_Atrasado_Tempo, Terceiro_Maior_Atrasado, Terceiro_Maior_Atrasado_Tempo)
#    os.remove(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Liberação_Maps\auxiliar.xlsx')

except Exception:
    pass


print('PROCESSO FINALIZADO PELO HORÁRIO')
