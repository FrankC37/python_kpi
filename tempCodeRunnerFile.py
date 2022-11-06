from calendar import week
import smtplib
import ssl
from email.message import EmailMessage
from datetime import date
import os
from tokenize import group
from turtle import home
import pandas as pd
from pathlib import Path
import locale
import datetime as DT



# Data Work
### ----- Setting Currency Locale

locale.setlocale(locale.LC_ALL,'')

### ----- Path Definitions
sp_prefix = r"C:/Users/Frank/USA Sealing. INC/USA Sealing. INC Team Site - Documents/pythonData/"
quote_kpi_file = sp_prefix + 'customsearch_usas_python_quote_summary.csv'



### ----- Checking Paths


df = pd.read_csv(quote_kpi_file)

num_of_quotes = int(df['Document Number'].count())
value_of_quotes = float(df['Amount'].sum())

quotes_currency = locale.currency(value_of_quotes,grouping=True)

#finding the top quotes
top_5_quotes = df.sort_values(by=['Amount'], ascending=False)
top_5_quotes = top_5_quotes.drop(['Date'], axis=1)
top_5_quotes = top_5_quotes.reset_index(drop=True)

#finding top quotes for unique customers
ranked_df = df.sort_values(by=['Amount'], ascending=False)
ranked_df = ranked_df.drop_duplicates(subset=['Name'])

### -- Needs to be updated with new logic to use dicts and loop these for cleaner code
### ----- Quotes Rankings for Email Table

docNumbers = top_5_quotes['Document Number'].values.tolist()

print(docNumbers[1])