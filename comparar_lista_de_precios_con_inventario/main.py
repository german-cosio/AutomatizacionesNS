import src.write_service as write_service

print("---------------------------------------------")
print("Proceso iniciado")

# Rutas de los archivos de entrada y salida
path_inventario = 'comparar_lista_de_precios_con_inventario/src/Listas/Conteo Inventario 01.05.24.xlsx'
path_lista_de_precios = 'comparar_lista_de_precios_con_inventario/src/Listas/Precios_Naval_2024 (1).xlsx'
output_file_path = 'comparar_lista_de_precios_con_inventario/output/Productos_faltantes_en_lista_de_precios.xlsx'

# Extraer valores de nombre y referencia interna del conteo de inventario
inventario = write_service.extract_column_values(path_inventario, 'C')
inventario_nombres = write_service.extract_column_values(path_inventario, 'A')
print(f"{len(inventario)} productos encontrados en Inventario")

# Extraer valores de referencia interna de la lista de precios
lista_de_precios = write_service.extract_column_values(path_lista_de_precios, 'B')
print(f"{len(lista_de_precios)} productos encontrados en Lista de precios")

# Escribir los resultados en un nuevo archivo de Excel
write_service.write_to_new_excel(inventario, inventario_nombres, lista_de_precios, output_file_path)
print("---------------------------------------------")