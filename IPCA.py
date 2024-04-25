import pandas as pd
import swifter
import datetime as dt
import math
import matplotlib.pyplot as plt
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta
from dateutil.rrule import *
from dateutil.parser import *
from datetime import *
import datetime as dt
import pprint
import sys
from pandas.tseries.offsets import *
import numpy as np
from io import BytesIO
import requests
import openpyxl

def get_feriados_anbima():
    response = requests.get('https://www.anbima.com.br/feriados/arqs/feriados_nacionais.xls')
    df = pd.read_excel(BytesIO(response.content))

    # Limita o DataFrame até a linha 1265
    df = df.iloc[:1264]

    # Converte as colunas para datetime e dia da semana
    df['Data'] = pd.to_datetime(df['Data'])
    df['Dia da Semana'] = df['Data'].dt.weekday

    # Ignora os feriados que caem nos finais de semana
    df = df[~df['Dia da Semana'].isin([5,6])]
    
    return df
    
def calc_business_days(date1, date2):
    return len(pd.date_range(start=date1, end=date2, freq=BDay()))

# Função para truncar valores
def truncate(num, n):
    integer = int(num * (10**n))/(10**n)
    return float(integer)

# Função para calcular o desconto no fluxo de caixa -> desconto = valor_cupom / (1+taxa_desconto)**du
def calc_discount(yield_rate, du): 
    return 1 / (1 + ((yield_rate/2)/100))**(du/180)

# Função para calcular a cotação da amortização
def calc_amortization(yield_rate, last_value):
    return 100 / (1 + ((yield_rate)/100))**(last_value/180)/100
    
    # 1.1 Inputs

    # yield rate: Taxa de desconto negociada para o título
    # settlement: Data de liquidação do título
    # maturity: Data de vencimento do título
    # VNA: VNA atualizados para a operação

    # 1.1.1 Obs:
    # Yield rate deve ser digitado sem % e.g. 6.32 para uma taxa de 6.32%
    # Para datas, utilizar modelo ISO e.g. "2022-07-27"
    # O VNA atualizado pode ser encontrado no site da ANBIMA

def execute_PU_NTN_B(yield_rate, settlement, maturity, VNA):
    # Gera as datas de pagamento dos cupons
    year, month, date = maturity.split('-')
    month = int(month)
    if month > 5: # Regra para datas no cashflow vencendo em Maio
        df = pd.DataFrame(rrule(MONTHLY, bymonth=(5,11), bymonthday=15,
                                    dtstart=parse(settlement), until=parse(maturity)), columns=["CUPOM_Date"])
    else: # Regra para datas no cashflow vencendo em Agosto
        df = pd.DataFrame(rrule(MONTHLY, bymonth=(2,8), bymonthday=15,
                                    dtstart=parse(settlement), until=parse(maturity)), columns=["CUPOM_Date"])
    
    feriados = get_feriados_anbima()

    # Ajusta as datas de pagamento dos cupons com base nos feriados
    df['CUPOM_Date'] = df['CUPOM_Date'].apply(lambda x: x if x not in feriados['Data'].values else x + pd.DateOffset(days=1))

    # Calcula a diferença de dias úteis entre a data de liquidação e a data de vencimento
    df["du_entre_datas"] = df["CUPOM_Date"].apply(lambda x: calc_business_days(settlement, x.strftime("%Y-%m-%d")))

    # Calcula o desconto no fluxo de caixa
    df["%_CUPOM"] = df["du_entre_datas"].apply(lambda x: calc_discount(yield_rate, x))

    # Calcula a cotação da amortização
    amort = calc_amortization(yield_rate, df.loc[df.index[-1], 'du_entre_datas'])

    # Adiciona a cotação da amortização ao último valor de cupom
    df.loc[df.index[-1], '%_CUPOM'] += amort

    # Calcula o valor dos pagamentos de cupons
    df["Valor_CUPOM"] = df["%_CUPOM"] * VNA

    # Calcula o valor presente de cada pagamento de cupom
    df["Valor_Presente_CUPOM"] = df["Valor_CUPOM"] * df["%_CUPOM"]

    # Calcula o valor presente total do título
    valor_presente_total = df["Valor_Presente_CUPOM"].sum()

    # Calcula o preço unitário (PU) do título
    cotação_perc = truncate(df['%_CUPOM'].sum(),6)
    
    df.to_excel('x.xlsx')
    
    return truncate((cotação_perc * VNA),6), valor_presente_total

execute_PU_NTN_B(6.25, '2022-07-28','2028-08-15', 3987.04)
