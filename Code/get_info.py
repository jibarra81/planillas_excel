import os
import os.path
import re


def get_file():
    """
        Devuelve el path del archivo .xlsx y el nombre del archivo
        Ej: 'C:/Documentos/archivos_externos/Dominios/archivo.xlsx'
        Ej: 'archivo.xlsx'
    """
    path = '../Dominios'
    type_file = ".xlsx"
    listado = list()
    names = list()
    for m_fullpath, m_dirs, m_files in os.walk(path):
        for m_file in m_files:
            match = re.search(type_file, m_file)
            if match is not None:
                filepath = os.path.join(m_fullpath, m_file)
                listado.append(filepath)
                filename = os.path.split(filepath)[1]
                names.append(filename)
    return listado, names


def create_sheet(wb, filesheet):
    """Funcion que crea la planilla para guardar los datos de los dominios"""
    return wb.save(filesheet)


def get_is_current(filas, sh):
    """
        Devuelve la cantidad de planillas que estan current
    """
    current = 0
    for i in range(1, filas + 1):
        celda = sh.cell(i, 3)
        if celda.value == 'X':
            current += 1
    return current


def get_is_proposed(filas, sh):
    """
        Devuelve la cantidad de planillas que estan propused
    """
    return (filas - 1) - get_is_current(filas, sh)
