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
sp_prefix = r"C:/Users/fconiglio/USA Sealing. INC/USA Sealing. INC Team Site - Documents/pythonData/"
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


### ----- Quotes Rankings for Email Table

doc_1 = top_5_quotes.iloc[0]['Document Number']
doc_2 = top_5_quotes.iloc[1]['Document Number']
doc_3 = top_5_quotes.iloc[2]['Document Number']
doc_4 = top_5_quotes.iloc[3]['Document Number']
doc_5 = top_5_quotes.iloc[4]['Document Number']

cust_1 = top_5_quotes.iloc[0]['Name']
cust_2 = top_5_quotes.iloc[1]['Name']
cust_3 = top_5_quotes.iloc[2]['Name']
cust_4 = top_5_quotes.iloc[3]['Name']
cust_5 = top_5_quotes.iloc[4]['Name'] 

# includes conversions to show as USD
raw_amt_1 = top_5_quotes.iloc[0]['Amount']
amt_1 = locale.currency(raw_amt_1,grouping=True)
raw_amt_2 = top_5_quotes.iloc[1]['Amount']
amt_2 = locale.currency(raw_amt_2,grouping=True)
raw_amt_3 = top_5_quotes.iloc[2]['Amount']
amt_3 = locale.currency(raw_amt_3,grouping=True)
raw_amt_4 = top_5_quotes.iloc[3]['Amount']
amt_4 = locale.currency(raw_amt_4,grouping=True)
raw_amt_5 = top_5_quotes.iloc[4]['Amount']
amt_5 = locale.currency(raw_amt_5,grouping=True)

sr_1 = top_5_quotes.iloc[0]['Sales Rep']
sr_2 = top_5_quotes.iloc[1]['Sales Rep']
sr_3 = top_5_quotes.iloc[2]['Sales Rep']
sr_4 = top_5_quotes.iloc[3]['Sales Rep']
sr_5 = top_5_quotes.iloc[4]['Sales Rep']

### ----- Customer Rankings for Email Table

unq_doc_1 = ranked_df.iloc[0]['Document Number']
unq_doc_2 = ranked_df.iloc[1]['Document Number']
unq_doc_3 = ranked_df.iloc[2]['Document Number']
unq_doc_4 = ranked_df.iloc[3]['Document Number']
unq_doc_5 = ranked_df.iloc[4]['Document Number']

unq_cust_1 = ranked_df.iloc[0]['Name']
unq_cust_2 = ranked_df.iloc[1]['Name']
unq_cust_3 = ranked_df.iloc[2]['Name']
unq_cust_4 = ranked_df.iloc[3]['Name']
unq_cust_5 = ranked_df.iloc[4]['Name'] 

# includes conversions to show as USD
unq_raw_amt_1 = ranked_df.iloc[0]['Amount']
unq_amt_1 = locale.currency(raw_amt_1,grouping=True)
unq_raw_amt_2 = ranked_df.iloc[1]['Amount']
unq_amt_2 = locale.currency(raw_amt_2,grouping=True)
unq_raw_amt_3 = ranked_df.iloc[2]['Amount']
unq_amt_3 = locale.currency(raw_amt_3,grouping=True)
unq_raw_amt_4 = ranked_df.iloc[3]['Amount']
unq_amt_4 = locale.currency(raw_amt_4,grouping=True)
unq_raw_amt_5 = ranked_df.iloc[4]['Amount']
unq_amt_5 = locale.currency(raw_amt_5,grouping=True)

unq_sr_1 = ranked_df.iloc[0]['Sales Rep']
unq_sr_2 = ranked_df.iloc[1]['Sales Rep']
unq_sr_3 = ranked_df.iloc[2]['Sales Rep']
unq_sr_4 = ranked_df.iloc[3]['Sales Rep']
unq_sr_5 = ranked_df.iloc[4]['Sales Rep']


### ----- Email Variables
# make sure your system has the eviroment variables setup
email_sender = os.environ.get('usas_email')
email_password = os.environ.get('usas_gmail_app_pw')
email_receiver = ['frank.coniglio@usasealing.com']


### ----- Time Variables
week_of = date.today()
adjusted_week_of = week_of -DT.timedelta(days=8)


### ----- Email Content
subject = f"USAS Quotes KPI for the week of {adjusted_week_of}."
body = """
Still in Testing,looking for feedback =)

"""
### ----- Email Class Definitions
em = EmailMessage()
em['From'] = email_sender
em['To'] = email_receiver
em['Subject'] = subject
em.set_content(body)
em.add_alternative(
    f"""\
<html>
<h1 align='center'>USAS Sales KPI</h1>
<table id="kpis" width="100%">
  <thead>
   <tr>
    <th scope="col"># of Quotes</th>
    <th scope="col">$ of Quotes</th>
   </tr>
  </thead>

  <tbody>
  <tr>
    <td align='center'>{num_of_quotes}</td>
    <td align='center'>{quotes_currency}</td>
  </tr>

  </tbody>
</table>
<body>
<h2 align='center'>Top 5 Quotes</h2>
<table id="quotes" width="100%">
  <thead>
   <tr>
    <th scope="col">Quote #</th>
    <th scope="col">Customer</th>
    <th scope="col">Amount</th>
    <th scope="col">Sales Rep</th>
  </thead>

  <tbody>
  <tr>
    <td align='center'>{doc_1}</td>
    <td align='center'>{cust_1}</td>
    <td align='center'>{amt_1}</td>
    <td align='center'>{sr_1}</td>
  </tr>
  <tr>
    <td align='center'>{doc_2}</td>
    <td align='center'>{cust_2}</td>
    <td align='center'>{amt_2}</td>
    <td align='center'>{sr_2}</td>
  </tr>
    <tr>
    <td align='center'>{doc_3}</td>
    <td align='center'>{cust_3}</td>
    <td align='center'>{amt_3}</td>
    <td align='center'>{sr_3}</td>
  </tr>
    <tr>
    <td align='center'>{doc_4}</td>
    <td align='center'>{cust_4}</td>
    <td align='center'>{amt_4}</td>
    <td align='center'>{sr_4}</td>
  </tr>
    <tr>
    <td align='center'>{doc_5}</td>
    <td align='center'>{cust_5}</td>
    <td align='center'>{amt_5}</td>
    <td align='center'>{sr_5}</td>
  </tr>
  </tbody>
</table>
</br>
<h2 align='center'>Top 5 Quotes per Unique Customer</h2>
<table id="unq_quotes" width="100%">
  <thead>
   <tr>
    <th scope="col">Quote #</th>
    <th scope="col">Customer</th>
    <th scope="col">Amount</th>
    <th scope="col">Sales Rep</th>
  </thead>

  <tbody>
  <tr>
    <td align='center'>{unq_doc_1}</td>
    <td align='center'>{unq_cust_1}</td>
    <td align='center'>{unq_amt_1}</td>
    <td align='center'>{unq_sr_1}</td>
  </tr>
  <tr>
    <td align='center'>{unq_doc_2}</td>
    <td align='center'>{unq_cust_2}</td>
    <td align='center'>{unq_amt_2}</td>
    <td align='center'>{unq_sr_2}</td>
  </tr>
    <tr>
    <td align='center'>{unq_doc_3}</td>
    <td align='center'>{unq_cust_3}</td>
    <td align='center'>{unq_amt_3}</td>
    <td align='center'>{unq_sr_3}</td>
  </tr>
    <tr>
    <td align='center'>{unq_doc_4}</td>
    <td align='center'>{unq_cust_4}</td>
    <td align='center'>{unq_amt_4}</td>
    <td align='center'>{unq_sr_4}</td>
  </tr>
    <tr>
    <td align='center'>{unq_doc_5}</td>
    <td align='center'>{unq_cust_5}</td>
    <td align='center'>{unq_amt_5}</td>
    <td align='center'>{unq_sr_5}</td>
  </tr>
  </tbody>
</table>
</br>
</br>
</br>
<p>This KPI utilizes Netsuite at the source of truth. </br><a href="https://4535487.app.netsuite.com/app/common/search/searchresults.nl?searchid=3080&whence=">USAS Quotes this week</a><p>
</body>
</html>
""",
    subtype='html',
)

### ----- Securing the Message vis SSL
context = ssl.create_default_context()


### ----- Login to Google and send message
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    smtp.login(email_sender, email_password)
    smtp.sendmail(email_sender, email_receiver, em.as_string())