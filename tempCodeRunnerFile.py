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
# sp_prefix = r"C:/Users/Frank/USA Sealing. INC/USA Sealing. INC Team Site - Documents/pythonData/"
spPathFromEnv = os.environ.get('OneDriveCommercial')
# quote_kpi_file = sp_prefix + 'customsearch_usas_python_quote_summary.csv'
quote_kpi_file = spPathFromEnv + r'\USA Sealing. INC Team Site - Documents\pythonData\customsearch_usas_python_quote_summary.csv'


print(quote_kpi_file)
