import os
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog

import datetime
import locale

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

root = tk.Tk()
root.withdraw()

try:
    print('********************************************************************\nINICIANDO...')
    archivo_excel = filedialog.askopenfilename(title='Selecciona un archivo Excel', filetypes=[('Archivos Excel', '*.xlsx')])
    libro = load_workbook(archivo_excel)
    hoja = libro.active
    
    hoy = datetime.date.today()
    prox_sem = hoy + datetime.timedelta(days=(7 - hoy.weekday()))
    fin_sem = prox_sem + datetime.timedelta(days=6)
    get_valor_desc = hoja['A1'].value
    sedes = f"{get_valor_desc.split(':')[0]}"
    set_valor_desc = f"{sedes}: PROGRAMACION DE MENU DEL {prox_sem.day} al {fin_sem.day} de {prox_sem.strftime('%B')}"
    print('********************************************************************\n',set_valor_desc)
    hoja['A1'].value = set_valor_desc
    
    ult_palabra_sede = sedes.split()[-1].lower()
    
    fila = 3  # Comienza en la primera fila
    fila_max = (lambda: 19 if ult_palabra_sede == 'incubacion' else 21 if  ult_palabra_sede == 'gabriela' else None)()
    
    columna = 'C'  # Comienza en la columna C
    seccion_a_ignorar = ['COMPLEMENTO', 'FRUTA', 'PAN']
    if ult_palabra_sede == 'incubacion':
        seccion_a_ignorar.append('BEBIDA')
        
    def cap_valor(valor):
        return valor[0].upper() + valor[1:].lower()
    
    while True:
        
        if columna == 'J':
            print('Se ha terminado de modificar el archivo.')
            libro.save('videogamesales.xlsx')
            break
        else:
            # Muestra el valor actual de la columna B en la misma fila
            valor_columna_b = hoja[f'B{fila}'].value
            
            if valor_columna_b not in seccion_a_ignorar:
            
                valor_dia = hoja[f'{columna}2'].value
                print(f'********************************************************************\nIngresa en: columna {columna} - fila {fila} - {valor_columna_b} - {valor_dia}')
                
                # Solicita al usuario el nuevo valor para la columna C
                valor = input(f'Escribe "fin" para terminar o "back" para corregir la celda anterior: ')
                if valor.lower() == 'fin':
                    break
                elif valor.lower() == 'back':
                    if fila == 3:
                        if columna == 'C':
                            print('********************************************************************\nYa estás en la primera columna y fila. No puedes retroceder más.')
                            continue
                        else:
                            columna = chr(ord(columna) - 1)  # Retrocede a la columna anterior
                            fila = fila_max  # Establece la fila al final del rango
                    fila -= 1  # Retrocede a la celda anterior
                    continue
                if valor == '':
                    valor = str('-')
                    print('Se registró por defecto "-".')
                hoja[f'{columna}{fila}'].value = cap_valor(valor)
            else:
                print(f'********************************************************************\n{valor_columna_b} en la columna {columna} - fila {fila} no se modificará.')
                
            fila += 1  # Incrementa la fila para continuar con la siguiente
            if fila > fila_max:
                fila = 3
                columna = chr(ord(columna) + 1)
            
            libro.save('videogamesales.xlsx')
    
except FileNotFoundError:
    print(f'El archivo "{libro}" no se encontró. Asegúrate de que el archivo exista en la ubicación especificada.')
except Exception as e:
    print(f'Ocurrió un error: {e}')
