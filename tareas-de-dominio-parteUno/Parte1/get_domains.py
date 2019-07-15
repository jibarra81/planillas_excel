

def get_domains(filas, sh):
    """ Devuelve una lista con los dominios que estan para trabajar"""
    listado = list()
    for i in range(1, filas + 1):
        celda = sh.cell(i, 2)
        if celda.value == 'X':
            dominio = sh.cell(i, 1).value
            listado.append(dominio.strip())
    return listado
