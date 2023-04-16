""" buscar la columna de un encabezado"""
def buscar_columna(fila : int, encabezado : str, hoja):
    # Recorrer las celdas de la fila buscando el encabezado con el nombre específico
    encabezado_buscado = encabezado  # Por ejemplo, buscamos el encabezado 'Cliente'
    columna_encabezado = None
    for columna in range(1, hoja.max_column + 1):
        celda = hoja.cell(row=fila, column=columna)
        if celda.value == encabezado_buscado:
            columna_encabezado = columna
            break
    return columna_encabezado


"""crear los encabezados de la hoja de reporte"""
def crear_encabezado(hoja : str, lista : list):
    #crear el titulo de la lista donde se va a comparar
    columna = hoja.max_column
    for encabezado in lista:
        # print(encabezado)
        hoja.cell(row=1, column= columna, value=encabezado)
        columna+=1


def rellenar_tabla (hoja : str, lista : list, diccionario):
    crear_encabezado(hoja, lista)
    
    """crear variables que se van a usar"""
    columna_total = buscar_columna(1, 'Total', hoja) #seleccionar la posicion de la columna
    columna_clientes = buscar_columna(1, 'Clientes', hoja) #seleccionando la columna doonde se ingresan los nombres de los clientes
    num = 2 #creando una variable para ir aumentando el
    
    """rellenar la columna de clientes en la hoja de Reporte"""
    for clave in diccionario.keys(): #iterar sobre las keys del diccionario
        persona = str(clave) #comvertir la key en un string
        cliente = hoja.cell(row=num, column=columna_clientes) #seleccionar la casilla que se desea rellenar
        cliente.value = persona
        # print(num)
        num = num + 1
    
    """rellenar la columna de los totales"""
    for fila in range(2, hoja.max_row +1 ):
        cliente = hoja.cell(row=fila, column=columna_clientes).value
        total = hoja.cell(row=fila, column=columna_total)

        total.value = diccionario[cliente]
        
        


