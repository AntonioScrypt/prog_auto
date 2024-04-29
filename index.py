import os
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog
import datetime
import locale

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

root = tk.Tk()
root.withdraw()

cadena = '********************************************************************\n'

count = 0
section_eat = 'DESAYUNO'

try:
    print(cadena,'INICIANDO...')
    file_src = filedialog.askopenfilename(title='Selecciona un archivo Excel', filetypes=[('Archivos Excel', '*.xlsx')])
    file_name = os.path.basename(file_src)
    
    file_exc = load_workbook(file_src)
    file_doc = file_exc.active
    
    hoy = datetime.date.today()
    prox_sem = hoy + datetime.timedelta(days=(7 - hoy.weekday()))
    fin_sem = prox_sem + datetime.timedelta(days=6)
    get_valor_desc = file_doc['A1'].value
    sedes = f"{get_valor_desc.split(':')[0]}"
    set_valor_desc = (f"{sedes}: PROGRAMACION DE MENU DEL {prox_sem.day} AL {fin_sem.day} DE {fin_sem.strftime('%B')}").upper()
    print(cadena + set_valor_desc)
    file_doc['A1'].value = set_valor_desc
    
    ult_palabra_sede = sedes.split()[-1].lower()
    
    fila = 3  # Comienza en la primera fila
    fila_max = (lambda: 19 if ult_palabra_sede == 'incubacion' else 21 if  ult_palabra_sede == 'gabriela' else None)()
    
    columna = 'C'  # Comienza en la columna C
    seccion_a_ignorar = ['COMPLEMENTO', 'FRUTA', 'PAN']
    if ult_palabra_sede == 'incubacion':
        seccion_a_ignorar.append('BEBIDA')
        
    def cap_valor(valor):
        return valor[0].upper() + valor[1:].lower()
    
    arr_head_sect = ['BEBIDA CALIENTE', 'SOPA/ENTRADA', 'SOPA']
    
    while True:
        if columna == 'J':
            print(f'{cadena}Se ha terminado de modificar el archivo.{cadena}')
            file_exc.save(f'{file_name}')
            break
        else:
            # Muestra el valor actual de la columna B en la misma fila
            valor_columna_b = file_doc[f'B{fila}'].value
            
            if valor_columna_b in arr_head_sect:
               count += 1
               if count == 1 : section_eat = 'DESAYUNO'
               elif count == 2 : section_eat = 'ALMUERZO'
               elif count == 3 : section_eat = 'CENA'
               elif count == 0 : section_eat = 'CENA'
                
            
            
            #print(count, section_eat, valor_columna_b)
            
            if valor_columna_b not in seccion_a_ignorar:
            
                valor_dia = file_doc[f'{columna}2'].value
                #print(f'{cadena}Ingresa en: columna {columna} - fila {fila} - {valor_columna_b} - {valor_dia} - {section_eat}')
                print(f'{cadena}{valor_dia} ||  {section_eat} || {valor_columna_b}')
                
                # Solicita al usuario el nuevo valor para la columna C
                valor = input(f'{cadena}Escribe "fin" para terminar o "back" para corregir la celda anterior: ')
                if valor.lower() == 'fin':
                    print(f'{cadena}Se cerró el registro.')
                    break
                elif valor.lower() == 'back':
                    if fila == 3:
                        if columna == 'C':
                            print(f'{cadena}Ya estás en {section_eat} del {valor_dia}. No puedes retroceder más.')
                            count = 0
                            continue
                        else:
                            columna = chr(ord(columna) - 1)
                            fila = fila_max  # Establece la fila al final del rango
                            
                    #if valor_columna_b in seccion_a_ignorar:
                    #    fila -= 2
                    #    if count == 1 : count = 0
                    #    elif count == 2 : count = 1
                    #    elif count == 3 : count = 2
                    #    elif count == 0 : count = 0
                    #else:
                    fila -= 1    
                    continue
                if valor == '':
                    valor = str('-')
                    print(cadena, 'Se registró por defecto "-".')
                file_doc[f'{columna}{fila}'].value = cap_valor(valor)
            else:
                print(f'{cadena}{valor_dia} || {valor_columna_b} de {section_eat} no se modificará.')
                
            fila += 1  # Incrementa la fila para continuar con la siguiente
            if fila > fila_max:
                fila = 3
                columna = chr(ord(columna) + 1)
                count = 0
            
            file_exc.save(f'{file_name}')
    
except FileNotFoundError:
    print(f'El sheet "{file_exc}" no se encontró. Asegúrate de que el archivo exista en la ubicación especificada.')
except Exception as e:
    print(f'Ocurrió un error: {e}')
