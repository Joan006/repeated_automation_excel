import pandas as pd 

# Ruta del archivo de Excel
archivo_excel = str(input("Coloca el pad del archivo (directorio/archivo.xlsx): "))

# Nombre de la hoja en Excel que contiene los datos
nombre_hoja = str(input("Coloca el nombre de la hoja donde esta la tabla: "))

# columna para eliminar los duplicados
col_guia = str(input("Coloca exactamente el nombre de la columna guía : "))

# columna de conteo de items 
col_items = str(input("Coloca exactamente el nombre de la columna del conteo : "))

col_indice = str(input("Coloca exactamente el nombre de la columna que lleva el índice : "))

# Lee toda la tabla y conviértela en un DataFrame
df_tabla = pd.read_excel(archivo_excel, sheet_name=nombre_hoja)

# obteniendo los items repetidos agrupándolos
agrupacion = df_tabla.groupby(col_guia)

# obteniendo las veces que se repite el item
lista_orden = []
lista_conteo = []
for grupo, indices in agrupacion.groups.items():
    cantidad_elementos = len(indices)
    lista_conteo.append(cantidad_elementos)
    # creamos una lista de los grupos que nos servirá para el nuevo orden
    lista_orden.append(grupo)
    print(f"{grupo}: {cantidad_elementos}")

# Nuevo orden del df
df_nuevo_orden = df_tabla.drop_duplicates(subset=[col_guia]).set_index(col_guia).reindex(lista_orden).reset_index()

# Insertar la columna de conteos
df_nuevo_orden[col_items] = lista_conteo

# Eliminar la columna de índice
df_tabla_sin_columna = df_nuevo_orden.drop(columns=[col_indice])

# Crear la lista de índices
index_lista = list(range(len(df_tabla_sin_columna)))

# Insertar la columna de índice al principio del DataFrame
df_tabla_sin_columna.insert(0, col_indice, index_lista)

# Guardar el DataFrame modificado en un nuevo archivo Excel
nombre_nuevo_excel = 'nuevo_archivo.xlsx'
df_tabla_sin_columna.to_excel(nombre_nuevo_excel, index=False)

