from datetime import datetime
import pandas as pd
import numpy as np

def eh_feriado(data):
    try:
        data_aux = datetime.strptime(data, '%d/%m/%Y')
    except Exception:
        data_aux = datetime.strptime(data, '%Y-%m-%d')
    feriados = pd.read_excel(r'\\scot\h\FUNDOS\GER1\Requisição de Baixa de Arquivos\Feriados.xlsx')
    return np.where(feriados.Data == data_aux)[0].shape[0]

