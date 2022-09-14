from mkt import open_table_mkt
from open_table import open_table_relatorio
import pandas as pd
import xlsxwriter
from config import *

config = CatalogConfig()
config.read()

df_periodo = open_table_mkt()
df_open = open_table_relatorio()

def dfs_tabs(df_list, sheet_list, file_name):
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')
    for dataframe, sheet in zip(df_list, sheet_list):
        dataframe.to_excel(writer, sheet_name=sheet, startrow=0 , startcol=0)
    writer.save()

# list of dataframes and sheet names
dfs = [df_open, df_periodo]
sheets = ['Open Table','Periodo']

# run function
path = config['PATH']['OPEN']
dfs_tabs(dfs, sheets, path)