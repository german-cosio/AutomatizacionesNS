import io
import xlwings as xw
import re
import tempfile
import shutil
import os

def writeXlsx(invoices, template_file):
    # Copy the template file to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_template:
        with open(template_file, 'rb') as f:
            shutil.copyfileobj(f, temp_template)
        temp_template_path = temp_template.name

    # Open the new Excel workbook and sheet
    app = xw.App(visible=False, add_book=False)
    workbook = app.books.open(temp_template_path)
    raw_data_ventas = workbook.sheets['Raw_data_ventas']
    exclusiones = workbook.sheets['Exclusiones']

    # Write the data
    row = 2
    row_ex = 2
    row_pointer = row
    regex_envio = re.compile(r'env[ií]o', re.IGNORECASE)

    for invoice in invoices:
        for product in invoice['products']:
            if (product['name'].lower().startswith('[mdm]') 
                or 'manejo de materiales' in product['name'].lower() 
                or invoice['invoice_number'].startswith('RINV')
                or regex_envio.search(product['name'])):
                worksheet = exclusiones
                row_pointer = row_ex
            else: 
                worksheet = raw_data_ventas
                row_pointer = row
            
            worksheet.range(f'A{row_pointer}').value = invoice['invoice_number']
            worksheet.range(f'B{row_pointer}').value = invoice['invoice_origin']
            worksheet.range(f'C{row_pointer}').value = invoice['invoice_date']
            worksheet.range(f'D{row_pointer}').value = invoice['invoice_exchange_rate']
            purchase_order = invoice['purchase_order']
            if isinstance(purchase_order, (list, tuple)):
                purchase_order = ', '.join(purchase_order)  # Concatenate the purchase_order values into a string
            else:
                purchase_order = str(purchase_order)  # Convert to a string if it's not iterable
            worksheet.range(f'E{row_pointer}').value = purchase_order
            worksheet.range(f'F{row_pointer}').value = invoice['invoice_medium']
            worksheet.range(f'G{row_pointer}').value = invoice['invoice_partner']
            worksheet.range(f'H{row_pointer}').value = product['name']
            worksheet.range(f'I{row_pointer}').value = product['quantity_sale']
            worksheet.range(f'J{row_pointer}').value = product['sale_price']
            worksheet.range(f'K{row_pointer}').value = product['sale_currency']
            worksheet.range(f'L{row_pointer}').value = invoice['paid_through_stripe']
            worksheet.range(f'M{row_pointer}').value = product['purchase_order_price']
            worksheet.range(f'N{row_pointer}').value = product['purchase_order_currency']
            worksheet.range(f'O{row_pointer}').value = product['price_stock']
            worksheet.range(f'P{row_pointer}').value = product['stock_currency']
            worksheet.range(f'Q{row_pointer}').value = product['discount']

            if worksheet == exclusiones:
                row_ex += 1
            else:
                row += 1

    worksheet = raw_data_ventas
    worksheet.range(f'A{row+1}:A{500}').api.EntireRow.Delete()

    worksheet = workbook.sheets['Comision_ventas']
    worksheet.range('B2').formula2 = "=IFERROR(IF(UNIQUE(Raw_data_ventas!W1:W499)=0,"",UNIQUE(Raw_data_ventas!W1:W499)),"")"
    worksheet.range('B56').formula2 = "=IFERROR(IF(UNIQUE(Raw_data_ventas!X1:X499)=0,"",UNIQUE(Raw_data_ventas!X1:X499)),"")"

    # Semaforo para resaltar los porcentajes de ganancia fuera de rango
    for row in range(3, 52):
        porcentaje_de_ganancia = worksheet.range(f'N{row}')
        porcentaje_de_ganancia2 = worksheet.range(f'O{row+54}')

        if not porcentaje_de_ganancia.value and not porcentaje_de_ganancia2.value:
            break

        if porcentaje_de_ganancia.value is not None:
            if 0 < porcentaje_de_ganancia.value < 0.30:
                porcentaje_de_ganancia.color = (255, 255, 185)
            elif porcentaje_de_ganancia.value <= 0 or porcentaje_de_ganancia.value >= 2:
                porcentaje_de_ganancia.color = (255, 100, 100)

        if porcentaje_de_ganancia2.value is not None:
            if 0 < porcentaje_de_ganancia2.value < 0.30:
                porcentaje_de_ganancia2.color = (255, 255, 185)
            elif porcentaje_de_ganancia2.value <= 0 or porcentaje_de_ganancia2.value >= 2:
                porcentaje_de_ganancia2.color = (255, 100, 100)
        
    delete_extra_rows(worksheet, 57, 106)
    delete_extra_rows(worksheet, 3, 52)

    # Crear un archivo temporal de salida
    temp_output_path = tempfile.mktemp(suffix='.xlsx')
    workbook.save(temp_output_path)
    workbook.close()
    app.quit()

    # Leer el archivo guardado en memoria
    output_stream = io.BytesIO()
    with open(temp_output_path, "rb") as f:
        output_stream.write(f.read())
    
    output_stream.seek(0)

    # Eliminar archivos temporales
    os.remove(temp_output_path)

    return output_stream

def write_instalaciones_Xlsx(output_stream):
    # Guardar el archivo en memoria a un archivo temporal para que xlwings pueda abrirlo
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_output:
        temp_output.write(output_stream.read())
        temp_output_path = temp_output.name
    
    try:
        app = xw.App(visible=False, add_book=False)
        workbook = app.books.open(temp_output_path)
        
        # Seleccionar las hojas de Raw_data
        raw_data_ventas = workbook.sheets["Raw_data_ventas"]
        raw_data_instalaciones = workbook.sheets["Raw_data_instalaciones"]

        # Buscar celdas que contienen el texto "instalación" o "diagnóstico" entre []
        # pattern = re.compile(r'^\[.*?(instalaci[oó]n|diagn[oó]stico)(es)?.*?\]', re.IGNORECASE)
        pattern = re.compile(r'(instalaci[oó]n|diagn[oó]stico)(es)?', re.IGNORECASE)
        row_idx = 2  # Comenzar en la fila 2 de la nueva hoja

        last_row = raw_data_ventas.range('H' + str(raw_data_ventas.cells.last_cell.row)).end('up').row

        # Iterar de abajo hacia arriba para evitar problemas al eliminar filas
        for i in range(last_row, 0, -1):
            cell_value = str(raw_data_ventas.range('H' + str(i)).value)
            if pattern.search(cell_value):
                # Copiar toda la fila a la nueva hoja
                raw_data_ventas.range(i, 1).expand('right').copy(raw_data_instalaciones.range(row_idx, 1))
                row_idx += 1
                # Eliminar la fila de raw_data_ventas
                raw_data_ventas.range(f'{i}:{i}').delete()

        raw_data_instalaciones.range(f'A{row_idx+1}:A{100}').api.EntireRow.Delete()

        worksheet = workbook.sheets["Comision_instalaciones"]
        worksheet.range('B2').formula2 = "=IFERROR(IF(UNIQUE(Raw_data_instalaciones!U1:U100)=0,"",UNIQUE(Raw_data_instalaciones!U1:U100)),"")"
        worksheet.range('B14').formula2 = "=IFERROR(IF(UNIQUE(Raw_data_instalaciones!V1:V100)=0,"",UNIQUE(Raw_data_instalaciones!V1:V100)),"")"

        delete_extra_rows(worksheet, 15, 22)
        delete_extra_rows(worksheet, 3, 10)
        
        # Guardar el libro de trabajo en un archivo temporal
        temp_output_path_final = tempfile.mktemp(suffix='.xlsx')
        workbook.save(temp_output_path_final)
        workbook.close()
        app.quit()

        # Leer el archivo guardado en memoria
        output_stream_final = io.BytesIO()
        with open(temp_output_path_final, "rb") as f:
            output_stream_final.write(f.read())
        
        output_stream_final.seek(0)

    finally:
        # Eliminar archivos temporales
        os.remove(temp_output_path)
        if os.path.exists(temp_output_path_final):
            os.remove(temp_output_path_final)

    return output_stream_final

def delete_extra_rows(worksheet, start_row, end_row):
    for row in range(start_row, end_row):
        if not worksheet.range(f'A{row}').value:
            worksheet.range(f'A{row}:A{end_row}').api.EntireRow.Delete()
            break