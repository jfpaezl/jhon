""" buscar la columna de un encabezado"""
def buscar_columna(fila : int, encabezado : str, hoja):
    # Identificar en qué fila se encuentra el encabezado
    fila_encabezado = fila  # Por ejemplo, si el encabezado está en la primera fila

    # Recorrer las celdas de la fila buscando el encabezado con el nombre específico
    encabezado_buscado = encabezado  # Por ejemplo, buscamos el encabezado 'Nombre'
    columna_encabezado = None
    for columna in range(1, hoja.max_column + 1):
        celda = hoja.cell(row=fila_encabezado, column=columna)
        if celda.value == encabezado_buscado:
            columna_encabezado = columna
            break


"""crear los encabezados de la hoja de reporte"""
def crear_encabezado(hoja : str, tabla : list):
    #crear el titulo de la tabla donde se va a comparar
    columna = hoja.max_column
    for encabezado in tabla:
        # print(encabezado)
        hoja.cell(row=1, column= columna, value=encabezado)
        columna+=1





