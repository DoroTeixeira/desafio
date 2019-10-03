import xlwings as xw
import pandas as pd

from datetime import datetime
from google.cloud import bigquery


#arquivo
fname = ".../Vendas_de_Combustiveis_m3.xls"


#lista de parametros
list_uf = [
'ACRE',
'ALAGOAS',
'AMAPÁ',
'AMAZONAS',
'BAHIA',
'CEARÁ',
'DISTRITO FEDERAL',
'ESPÍRITO SANTO',
'GOIÁS',
'MARANHÃO',
'MATO GROSSO',
'MATO GROSSO DO SUL',
'MINAS GERAIS',
'PARÁ',
'PARAÍBA',
'PARANÁ',
'PERNAMBUCO',
'PIAUÍ',
'RIO DE JANEIRO',
'RIO GRANDE DO NORTE',
'RIO GRANDE DO SUL',
'RONDÔNIA',
'RORAIMA',
'SANTA CATARINA',
'SÃO PAULO',
'SERGIPE',
'TOCANTINS'
]


list_prd = [
'ÓLEO DIESEL (OUTROS ) (m3)',
'ÓLEO DIESEL MARÍTIMO (m3)',
'ÓLEO DIESEL S-10 (m3)',
'ÓLEO DIESEL S-1800 (m3)',
'ÓLEO DIESEL S-500 (m3)'
]


def parse_mes(mes):

    list_mes = [
    ['Janeiro',1],
    ['Fevereiro',2],
    ['Março',3],
    ['Abril',4],
    ['Maio',5],
    ['Junho',6],
    ['Julho',7],
    ['Agosto',8],
    ['Setembro',9],
    ['Outubro',10],
    ['Novembro',11],
    ['Dezembro',12]
    ]

    for m in list_mes:
        if (m[0] == mes):
            return m[1] 
    
    return m[0]


#define data frame 
df1 = pd.DataFrame()


#coleta os dados da pivot
for uf in list_uf:
    wbook = xw.Book( fname )
    wbook.sheets['Plan1'].range('C129').value = uf

    for prd in list_prd:
        wbook.sheets['Plan1'].range('C130').value = prd
        wbook.save()
        
        df = pd.read_excel(fname,sheet_name="Plan1",skiprows=132,header=0, nrows=12, index_col=0)
        
        df = df.loc[:, ['Dados',2013,2014,2015,'2016',2017,2018,2019]]
        df['estado'] = uf
        df['produto'] = prd
        df['unidade'] = 'm3'

        df1 = df1.append(df)


df1['timestamp_captura'] = datetime.now()
df1 = pd.melt(df1, id_vars=['Dados','estado','produto','unidade','timestamp_captura'])
df1 = df1.dropna(thresh=7)
df1 = df1.rename(columns={'Dados':'mes','variable':'ano','value':'vol_demanda_m3'})


for i, row in df1.iterrows():
    df1.at[i,'mes'] = parse_mes(df1.at[i,'mes'])


#load no GCP bigquery
client = bigquery.Client.from_service_account_json("...")
dataset_ref = client.dataset('data_set_name')
table_ref = dataset_ref.table('table_name')
client.load_table_from_dataframe(df1, table_ref).result()


print('Fim')