import email
from pickle import FALSE
from re import M, X
import re
from types import LambdaType
from typing import Union
from openpyxl.styles.borders import Border
import pandas as pd
from pandas.core.indexes.api import union_indexes
from pandas.core.reshape.pivot import crosstab, pivot
from pandas.io import html
from pandas.io.formats import style
from pandas import ExcelWriter
import xlsxwriter as xl
from openpyxl import workbook
import openpyxl as op
import numpy as np
import matplotlib.pyplot as pit
import seaborn as sns
import mysql.connector
import datetime
from seaborn import load_dataset
from tabulate import tabulate
import os
import os.path
import shutil
from pathlib import Path
# import win32com.client as win32
import glob
import shutil
import openpyxl
from openpyxl.styles import Border, Side, colors
import locale
import decimal
import sys
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pathlib
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
#######poner en report######
from io import BytesIO
import xlsxwriter

try:
    from urllib.request import urlopen
except ImportError:
    from urllib import urlopen


# today = date.today()

# Conexion a Base de datos
connection = mysql.connector.connect(
    host="",
    user="",
    password="",
    database=""

)
cursor = connection.cursor(prepared=True)


today = datetime.date.today()
print(today)

 # dia anterior
fecha = today - datetime.timedelta(days=1)
print(str(fecha))


#fecha = '2022-06-9'

# fecha = '2021-07-8'
# def fecha1():
# 	today=datetime.date.today()
# 	oneday=datetime.timedelta(days=1)
# 	yesterday=today-oneday
# 	return yesterday

# fecha = fecha1()
# print(fecha)
# hoy = datetime.date.today()
# dia = datetime.timedelta(days=1)
# ayer = hoy-dia
# fecha = ayer
# fecha1=datetime.datetime.today().strftime('%Y-%m-%d')-1
# print(fecha1)

today_string = datetime.datetime.today().strftime('%m%d%Y')
today_string2 = datetime.datetime.today().strftime('%b %d, %Y')

attachment_path = pathlib.Path.home() / 'data/attachments'

# attachment_path = pathlib.Path.home()   /"data"/"attachments" /"home/support/var/www/html/application/data/attachments/"

# archive_dir = Path.cwd() / 'archive'
###############################################################################################


###########################################################

 # Query para Db
sql = "SELECT distinct t.retailerid FROM mireporting.transactions t inner join mireporting.tran_correo c on t.retailerid = c.retailerid WHERE t.fecha_transaccion= '" + \
    str(fecha)+"' and t.codigo_resp =001  and c.correo  is not null  "
cursor.execute(sql)
resultado = cursor.fetchall()
for registro in resultado:
    print(*registro)
# dft =pd.read_sql_query("SELECT t.retailerName as SAB,t.tipo_tx as TIPO_TX,count(t.tipo_tx) as TX,sum(t.monto) as MONTOS  FROM mireporting.transactions t inner join mireporting.tran_correo c on t.retailerid  = c.retailerid   WHERE t.retailerid ='"+str(*registro)+"'  group by t.tipo_tx order by t.retailerName",connection)
    # cursor.execute(dft)
    # resultado1=cursor.fetchall()
    # print(resultado1)
#   for result in dft:
#     #  print (*result)
    # print (dft)

    # WHERE t.retailerid ='"+str(*registro)+"'
   # TABLA PIVOTE

    pvtt = pd.read_sql_query("SELECT distinct t.tipo_tx as COD_TX,tt.nombre as NOMBRE_TX, count(t.tipo_tx) as TOTAL_COD_TX ,sum(t.monto) AS MONTO,t.retailerid as CUC,t.retailerName as SAB,t.fecha_transaccion as FECHA FROM mireporting.transactions t inner join mireporting.tran_correo c on t.retailerid = c.retailerid inner join mireporting.nombres_tx tt on tt.tipo_tx = t.tipo_tx  WHERE t.retailerid ='" +
                             str(*registro)+"' and t.fecha_transaccion= '"+str(fecha)+"'  and  t.codigo_resp =001 group by t.tipo_tx,t.retailerid,t.retailerName,t.tipo_tx,tt.nombre", connection)
    # where t.fecha_transaccion ='"+fecha+"'

    df1 = pd.DataFrame(pvtt)

    # pvtt1 = pd.pivot_table(pvtt,index=['SAB', 'COD_TX','NOMBRE_TX','TOTAL_COD_TX'],values=['MONTO'],aggfunc='sum')
    GD1 = df1.groupby('COD_TX')
    # print(df1)

    # print(format_dict)

# este scripts lo que hace es dividir los archivos por Sab
    attachments_pivot = []
    for id, group_df in GD1:
        attachment = attachment_path / f' {id}_{today_string}.xlsx'
        group_df.to_excel(
            attachment, sheet_name='Reporte consolidado', index=False)
        attachments_pivot.append((id, str(attachment)))

        df9 = pd.DataFrame(attachments_pivot, columns=['COD_TX', 'FILE'])
    # print(df9)

        # email_marge = pd.merge(df1,df9, how="left")
        # combinar = email_marge[['SAB','CORREO','FILE']].drop_duplicates()
        # print(attachments_pivot)

    # print(combinar)
###############################################################################

    # Reporte general

    df = pd.read_sql_query("SELECT t.retailerid as CUC,t.retailerName as SAB,t.terminalid as TERMINAL,t.tipo_tx as COD_TX ,t.monto AS MONTO,t.fecha_transaccion as FECHA,date_format(t.hora_transaccion,'%H:%i:%s') as HORA,t.codigo_resp COD_RESP, t.codigo_convenio AS COD_CONVENIO,nb.nombre as Banco,c.correo as CORREO FROM mireporting.transactions t  inner join Nombre_bancos nb on t.fiid_autoriza = nb.fiid_autoriza inner join tran_correo c on t.retailerid = c.retailerid inner join nombres_tx tn on t.tipo_tx = tn.tipo_tx  where  t.retailerid='" +
                           str(*registro)+"'  and t.fecha_transaccion ='"+str(fecha)+"' and t.codigo_resp ='001'  order by t.retailerid,t.retailerName,t.terminalid,t.tipo_tx,t.monto,t.fecha_transaccion,t.hora_transaccion, t.codigo_resp,t.codigo_convenio,c.correo ", connection)
    #print('Esta es la consulta', df)
    df3 = pd.DataFrame(df)

    GT = df.groupby('Banco')
    # print(df3)

    # este scripts lo que hace es dividir los archivo por CUC
    attachments = []
    for id, group_df in GT:
        attachment = attachment_path / f' {id}_{today_string}.xlsx'
        group_df.to_excel(
            attachment, sheet_name='Reporte General', index=False)
        attachments.append((id, str(attachment)))

    # DataFrame de CUC
        df2 = pd.DataFrame(attachments, columns=['Banco', 'FILE'])

    # este script envia los archivos divididos por CUC
        email_marge1 = pd.merge(df3, df2, how='left')
        combinar1 = email_marge1[['Banco', 'CORREO', 'FILE']].drop_duplicates()
        print(combinar1)


########################################################################################
    attachments_principal = []
    for id, group_df in GD1:
        for id, group_df in GT:
            # with ExcelWriter(attachment_path /f' {id}_{today_string}.xlsx') as writer:
            writer = pd.ExcelWriter(
                attachment_path / f' {id}_{today_string}.xlsx')

            df1.to_excel(writer, sheet_name='Reporte consolidado',
                         startcol=1, startrow=10, index=False)
            df.to_excel(writer, sheet_name='Reporte general',
                        startrow=6, index=False)

        #########poner report######
        workbook = xlsxwriter.Workbook('images_bytesio.xlsx')

        workbook = writer.book
        worksheet = writer.sheets['Reporte consolidado']
        worksheet1 = writer.sheets['Reporte general']
        header_format = workbook.add_format()
        header_format.set_bg_color('#B8CCE4')
        header_format1 = workbook.add_format()
        header_format1.set_font_color('#FFFFFF')
        header_format2 = workbook.add_format()
        header_format2.set_border(1)
        header_format.set_border(1)
        percent_format = workbook.add_format({'num_format': '$0,0.0'})

        header_format3 = workbook.add_format()
        header_format3.set_font_color('#FFFFFF')

        # FORMAT PARA JUNSTIFICAR

        header_format.set_align('center')
        header_format.set_align('vcenter')

        plantilla = workbook.add_format({'font_color': '#ffffff'})

        bold5 = workbook.add_format({'bold': True, 'size': 14})
        bold6 = workbook.add_format({'bold': True, 'size': 11})
        bold7 = workbook.add_format(
            {'size': 11, 'align': 'center', 'font_color': '#9391D0'})
        bold2 = workbook.add_format({'bold': True})
        bold1 = workbook.add_format(
            {'bold': True, 'fg_color': '#1E1E1E', 'font_color': '#ffffff'})
        bold = workbook.add_format(
            {'bold': True, 'num_format': '$0,0.0', 'fg_color': '#1E1E1E', 'font_color': '#ffffff'})
        css1 = workbook.add_format({'bold': True, 'align': 'center'})
        cont = workbook.add_format(
            {'border': 1, 'fg_color': '#B8CCE4'})
        format3 = workbook.add_format(
            {'num_format': 'mm/dd/yy', 'fg_color': '#B8CCE4', 'border': 1, 'align': 'left', })
        # fecha = workbook.add_format({'default_date_format':'yy/mm/dd','fg_color': '#B8CCE4','border': 1,'align': 'center', 'valign': 'top'})
        cont2 = workbook.add_format({'bold': True})

        format4 = workbook.add_format({'font_color': '#ffffff'})

        format2 = workbook.add_format(
            {'bold': False, 'align': 'center', 'valign': 'top', 'text_wrap': True, 'border': 1})

        worksheet.set_column('A:H', 5)
        worksheet.set_row(3, 17)
        worksheet.set_row(6, 17)
        worksheet.set_row(7, 17)

        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#B8CCE4', 'size': 12})

        worksheet.merge_range('B10:E10', 'TRANSACCIONES', merge_format)

        #####################CAMBIO REPORTE################################

        # worksheet.merge_range('D4:D4','')

        # FORMATO Tabla pivote:
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 35)
        worksheet.set_column('D:D', 13)
        worksheet.set_column('E:E', 35)

        #############################
       # worksheet.set_row(2,15)
        # worksheet.set_column('A:A',header_format2)
        #############################
        worksheet.write(10, 1, 'CODIGO', header_format)
        worksheet.write(
            10, 2, 'NOMBRE DE LA TRANSACCION', header_format)
        worksheet.write(10, 3, 'CANTIDAD', header_format)
        worksheet.write(10, 4, 'VOLUMEN PROCESADO', header_format)

        worksheet.write(10, 5, 'CUC')
        worksheet.write(10, 6, 'SAB')
        worksheet.write(10, 7, 'FECHA')

        cell_format = workbook.add_format()

        cell_format.set_font_color('red')

        worksheet.set_column('F:F', 0, header_format1)
        worksheet.set_column('G:G', 0, header_format1)
        # worksheet.set_column('H12:H17',10,header_format1)
        # worksheet.set_column('H12:H45',10,plantilla)
        # worksheet.conditional_format('H12:H1048576', {'type': '#ffffff','font_color': '#ffffff'})

        efg = workbook.add_format({'font_color': '#ffffff'})
        due_range = 'H1:H50'
        worksheet.conditional_format(due_range,
                                     {'type': 'no_blanks',
                                      'format':    efg})

        # worksheet.set_column('A1:Z500', 20,plantilla)
        # worksheet.set_column('H1:H30',12, plantilla)
        #
        worksheet.set_column('E12:E50', 20, percent_format)
        # worksheet.set_column('C6:C50', 20, percent_format)

        row = 0
        col = 0
    # for k,v,p,t,u in df1:#Traverse our data, then fill the cell ([row, col], [row, col + 1]), and then row increment row + = 1 to enter the next cycle, get the next row of data to fill
    #         worksheet.write(row,col,k)
    #         worksheet.write(row,col+20,v)
        row += 26+1
        # Last row and one column negotiation name total
        worksheet.write(row, 2, 'Total', bold2)

        worksheet.write(row, 4, '=SUM(E12:D{})'.format(row), bold)
        worksheet.write(row, 3, '=SUM(D12:C{})'.format(row), bold1)

        # FORMATO Reporte general:

        worksheet1.set_column('A:A', 15)
        worksheet1.set_column('B:B', 20)
        worksheet1.set_column('C:C', 20)
        worksheet1.set_column('D:D', 20)
        worksheet1.set_column('E:E', 20)
        worksheet1.set_column('F:F', 20)
        worksheet1.set_column('G:G', 20)
        worksheet1.set_column('H:H', 20)
        worksheet1.set_column('I:I', 20)

        ###########BORDER##########################################

        #########CAMBIO PARA EL REPORTE##############################################

        worksheet1.write(
            'A6', 'DETALLE DE OPERACIONES PROCESADAS', bold5)
        worksheet.write(
            'C4', 'ESTADO DE TRANSACCIONES PROCESADAS ', bold5)

        worksheet.write(
            'B35:D35', 'En caso de identificar alguna diferencia o descuadre, favor contactar a MiRed a los numeros 809 530 3995, que con gusto le atenderemos.', bold6)
        worksheet.write(
            'B36:D36', ' Este mensaje contiene información legalmente considerada confidencial y privilegiada, con la intención de que sea utilizada', bold6)
        worksheet.write(
            'B37:D37', ' exclusivamente por las personas u organizaciones a quienes está dirigido. De haber recibido este mensaje por error, debes eliminarlo e', bold6)
        worksheet.write(
            'B38:D38', 'informarnos de inmediato al número indicado más arriba.', bold6)
        worksheet.write_url('C39:C39', 'www.mired.com.do', bold7)

        # worksheet1.set_column('A2:I',1,header_format2)

        # url = 'logomired.jpg'

        # image_data = BytesIO(urlopen(url).read())

       # worksheet.insert_image('B2', './images/logomired.jpg', {'x_scale': 0.8, 'y_scale': 0.8})

        # worksheet1.insert_image('A3', filename, {'images': image_data})

        #############################################################

        worksheet1.write(6, 0, 'CUC', header_format)
        worksheet1.write(6, 1, 'SAB', header_format)
        worksheet1.write(6, 2, 'TERMINAL', header_format)
        worksheet1.write(6, 3, 'COD_TX', header_format)
        worksheet1.write(6, 4, 'MONTO', header_format)
        worksheet1.write(6, 5, 'FECHA', header_format)
        worksheet1.write(6, 6, 'HORA', header_format)
        worksheet1.write(6, 7, 'COD_RESP', header_format)
        worksheet1.write(6, 8, 'COD_CONVENIO', header_format)
        worksheet1.set_column('J:J', 9, header_format1)
        worksheet1.set_column('K:K', 10, header_format1)
        worksheet1.write(6, 9, 'Banco', header_format1)
        worksheet1.write(6, 10, 'CORREO', header_format1)
        worksheet1.set_column('E2:E50', 20, percent_format)

        worksheet.insert_image(
            'B2', '/images/logoM.jpg', {'x_scale': 0.7, 'y_scale': 0.7})
        worksheet1.insert_image(
            'A2', '/images/logoM.jpg', {'x_scale': 0.7, 'y_scale': 0.7})

        ###################CAMBIO########################################
        worksheet.write(5, 1, 'FECHA:', cont2)
        worksheet.write(6, 1, 'ESTABLECIMIENTO:', cont2)
        worksheet.write(7, 1, 'CUC:', cont2)
        # worksheet.write(3,1, 'TRANSACCIONES', header_format)
        # worksheet.write(3,0, '', header_format)
        # worksheet.write(3,2, '', header_format)
        # worksheet.write(3,3, '', header_format)
        fe = 0
        cuc = 0
        est = 0
        fe = 5
        cuc += 7+0
        est += 6+0
        worksheet.set_column('B:B', 25)

        worksheet.write(cuc, 2, '=+(F12)'.format(row), cont)
        worksheet.write(est, 2, '=+(G12)'.format(row), cont)
        worksheet.write(fe, 2, '=+(H12)'.format(row), format3)

        worksheet.hide_gridlines(2)
        worksheet1.hide_gridlines(2)

        writer.close()

        # our_list = [(str(*registro), str(fecha))]
        # sqlcon = "INSERT INTO  auditoria_enviados(retailerid,fecha_enviado) VALUES(%s,%s)"

        # cursor.executemany(sqlcon, our_list)
        # connection.commit()

        # connection.rollback()

    for correito in combinar1['CORREO']:
        for archivito in combinar1['FILE']:
            print("Correo: " + correito)
            print("Archivo: " + archivito)

            def send_email():
                print(correito)
                print(archivito)

                email_sender = 'notificaciones@mired.com.do'
                email_recipient = correito

                msg = MIMEMultipart()
                msg['From'] = email_sender
                msg['To'] = email_recipient
                msg['Subject'] = 'Reporte de Transacciones'
                attachment_location = archivito
                email_message = 'Reporte de Transacciones'
                # f = open(each, 'rb')

                msg.attach(MIMEText(email_message, 'plain'))

                if attachment_location != '':
                    filename = os.path.basename(attachment_location)
                    attachment = open(attachment_location, "rb")
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition',
                                    "attachment; filename= %s" % filename)
                    msg.attach(part)
                # f.close()

                try:
                    server = smtplib.SMTP('smtp.office365.com', 587)
                    server.ehlo()
                    server.starttls()
                    server.login('notificaciones@mired.com.do',
                                 '123456789NOT@')
                    text = msg.as_string()
                    server.sendmail(email_sender, email_recipient, text)
                    print('Reporte enviado')
                    server.quit()

                except:
                    print("Error de conexion")
                return True

            send_email()

            our_list = [(str(*registro), str(fecha),
                         str(correito), str(archivito))]
        sqlcon = "INSERT INTO  auditoria_enviados(retailerid,fecha_enviado,correo,archivo) VALUES(%s,%s,%s,%s)"

        cursor.executemany(sqlcon, our_list)
        connection.commit()

        connection.rollback()


for root, dirs, files in os.walk('data/attachments/'):
    for f in files:
        os.unlink(os.path.join(root, f))
        for d in dirs:
            shutil.rmtree(os.path.join(root, d))
