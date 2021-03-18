#!/usr/bin/env python

import pandas as pd
import xlrd
import sys
import os

reload(sys)
sys.setdefaultencoding('utf-8')
pd.set_option('mode.chained_assignment', None)


def convert_number(number):
    result = str(number)
    # result = result.replace('.', '')
    # result = result.replace(',', '.')
    return result


def convert_date(date):
    chan = date.split('/')
    result = chan[2] + '-' + chan[1] + '-' + chan[0]
    return result


def join_texts(t1, t2):
    lista = []
    t1 = str(t1)
    t2 = str(t2)
    if 'nan' != t1:
        lista.append(t1)
    if 'nan' != t2:
        lista.append(t2)
    result = ' - '.join(lista)
    result = result.replace(',', ' ')
    return result


def get_name(date):
    mes = ['na', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
           'julio', 'Agosto', 'Septiebre', 'Octubre', 'Noviembre', 'Diciembre']
    chan = date.split('-')
    return mes[int(chan[1])] + ' ' + chan[0]


def final_date(date):
    chan = date.split('-')
    mes = str(int(chan[1]) + 1).zfill(2)
    return chan[0] + '-' + mes + '-01'


def main():
    if len(sys.argv) > 1:
        input_file = sys.argv[1]

        if len(sys.argv) > 2:
            output_file = sys.argv[2]
        else:
            output_file = os.path.splitext(input_file)[0] + '.csv'

        try:
            xl = pd.ExcelFile(input_file)
        except IOError:
            help_string = 'El archivo "' + input_file + '" no se ha podido abrir'
            print(help_string)
            sys.exit()
        except xlrd.XLRDError:
            help_string = 'El archivo "' + input_file + '" no parece una hoja de Excell'
            print(help_string)
            sys.exit()

        lista_hojas = xl.sheet_names
        df = xl.parse(lista_hojas[0])
        ndf = df.iloc[15:]

        ndf.columns = ["nada", "fecha", "nada", "nada", "concepto",
                       "descripcion", "importe", "saldo", "nada", "nada",
                       "nada"]

        saldo_primero = float(convert_number(ndf['saldo'].iloc[-1]))
        mov_ini = float(convert_number(ndf['importe'].iloc[-1]))
        saldo_inicial = saldo_primero - mov_ini

        ndf.loc[:, 'importe'] = ndf.loc[:, 'importe'].apply(convert_number)
        ndf.loc[:, 'saldo'] = ndf.loc[:, 'saldo'].apply(convert_number)
        ndf.loc[:, 'fecha'] = ndf.loc[:, 'fecha'].map(convert_date)
        ndf['descripcion'] = ndf['descripcion'].combine(ndf['concepto'],
                                                        join_texts)
        ndf.loc[:, 'vacio'] = ''
        final = pd.DataFrame(ndf[['vacio', 'vacio', 'vacio', 'vacio', 'fecha',
                                  'descripcion', 'importe']])
        final.columns = ['date', 'name', 'balance_start', 'balance_end_real',
                         'line_ids/date', 'line_ids/name', 'line_ids/amount']

        fecha_inicial = final['line_ids/date'].iloc[-1]

        final['date'].iloc[0] = final_date(fecha_inicial)
        final['balance_end_real'].iloc[0] = ndf['saldo'].iloc[0]
        final['balance_start'].iloc[0] = saldo_inicial
        final['name'].iloc[0] = get_name(fecha_inicial)

        try:
            final.to_csv(output_file, sep=',', encoding='utf-8', index=False)
        except IOError:
            help_string = 'No se ha podido guardar el archivo ' + output_file
            print(help_string)

    else:
        help_string = 'Uso: ' + sys.argv[0] + ' input.xls [output.csv]'
        print(help_string)


main()