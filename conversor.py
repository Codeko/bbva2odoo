#!/usr/bin/env python

import sys
import os
import csv
import datetime
import codecs

reload(sys)
sys.setdefaultencoding('utf-8')

def convert_number(number):
    result = str(number)
    result = result.replace('.', '')
    result = result.replace(',', '.')
    return result

def convert_date(date):
    day, month, year = date.split('/')
    return year + '-' + month + '-' + day

def is_date(date):
    result = True
    try:
        day, month, year = date.split('/')
        datetime.datetime(int(year), int(month), int(day))
    except ValueError:
        result = False
    return result

def join_texts(t1,t2):
    lista = []
    t1 = str(t1)
    t2= str(t2)
    if 'nan' != t1:
        lista.append(t1)
    if 'nan' != t2:
        lista.append(t2)
    result = ' - '.join(lista)
    result = result.replace(',', ' ')
    return result

def get_name(date):
    mes = ['na',
            'Enero',
            'Febrero',
            'Marzo',
            'Abril',
            'Mayo',
            'Junio',
            'julio',
            'Agosto',
            'Septiembre',
            'Octubre',
            'Noviembre',
            'Diciembre']
    chan = date.split('-')
    return mes[int(chan[1])] + ' ' + chan[0]

def final_date(date):
    chan = date.split('-')
    mes = str(int(chan[1]) + 1).zfill(2) 
    return chan[0] + '-' + mes  + '-01'

def cabeceras():
    return ['date',
            'name',
            'balance_start',
            'balance_end_real',
            'line_ids/date',
            'line_ids/name',
            'line_ids/amount']

def save_csv(campos, filename):
    with codecs.open(filename, 'w', 'utf-8') as out:
        csv_writer = csv.writer(out, delimiter=',', quotechar='"')
        csv_writer.writerow(cabeceras())
        for line in campos:
            csv_writer.writerow(line)

def extract_csv(filename):
    campos = []
    saldo_inicial = ''
    saldo_final = ''
    fecha_inicial = ''
    try:
        with codecs.open(filename, 'r', encoding='iso-8859-1', errors='replace') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter='\t')
            for row in csv_reader:
                if len(row) > 6:
                    if row[0] != '' and is_date(row[0]):
                        date = convert_date(row[0])
                        if '' == fecha_inicial:
                            fecha_inicial = date 
                        cantidad = convert_number(row[6])
                        texto = join_texts(row[3].strip(), row[4].strip("\'"))
                        campos.append([None, None, None, None, date, texto, cantidad])
                    else:
                        if 'Saldo inicial' in row[3]:
                            saldo_inicial = row[6]
                        elif 'Saldo inicial' in row[4]:
                            saldo_inicial = row[7]
                        elif 'Saldo final' in row[3]:
                            saldo_final = row[6]
                        elif 'Saldo final' in row[4]:
                            saldo_final = row[7]
            campos.reverse()
            campos[0][0] = final_date(fecha_inicial)
            campos[0][1] = get_name(fecha_inicial)
            campos[0][2] = convert_number(saldo_inicial)
            campos[0][3] = convert_number(saldo_final)
    except IOError as e:
        err_string = 'Ops! No se ha encontrado el archivo ' + filename
        print(err_string)
    except:
        err_string = 'Ops! Error inesperado' 
        print(err_string)
    return campos

def main():
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        if len(sys.argv) > 2:
            output_file = sys.argv[2]
        else:
            output_file = os.path.splitext(input_file)[0] + '.csv'
        campos = extract_csv(input_file)
        if campos:
            save_csv(campos, output_file)
        else:
            err_string = 'No se han obtenido datos' 
            print(err_string)

    else:
        help_string = 'Uso: '+  sys.argv[0] + ' input.xls [output.csv]'
        print(help_string)

main()
