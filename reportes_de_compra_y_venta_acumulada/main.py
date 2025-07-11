import os
import sys
import logging
from datetime import datetime
from dotenv import load_dotenv
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.odoo_service.odoo_service import OdooService
from reportes_de_compra_y_venta_acumulada.src.write_service import write_products_to_excel, write_invoices_to_excel

load_dotenv()

# Set up log
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def main():
    try:
        # Crear una instancia de OdooService
        odoo_service = OdooService()

        year = datetime.now().year
        start_date = f'{year}-01-01'
        end_date = f'{year}-12-31'
        products_data = []
        output_file = os.path.join(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output'), f'Resumen_anual_de_ventas_{year}.xlsx')

        invoices = odoo_service.get_entire_year_invoices(start_date, end_date)
        write_invoices_to_excel(output_file, invoices)
        
        total_invoices = len(invoices)
        for i, invoice in enumerate(invoices, start=1):
            print(f"\rProcessing invoice {i}/{total_invoices}", end='')
            if(i >= 10): break

            line_ids = invoice['line_ids']
            if line_ids:
                product_lines = odoo_service.request_manager('account.move.line', 'read', [line_ids, ['product_id', 'quantity', 'price_unit']])
                product_lines2 = odoo_service.request_manager('account.move.line', 'read', [line_ids, ['product_id', 'quantity', 'price_unit']])
                for line in product_lines:
                    if line['product_id']:
                        product = odoo_service.request_manager('product.product', 'read', [line['product_id'][0], ['name', 'default_code', 'categ_id']])

                        products_data.append([product[0]['name'], line['quantity'], line['price_unit'], invoice['invoice_date'], product[0]['categ_id'][1], invoice['display_name']])
        print("")

        # Llamar a la función para escribir en Excel
        write_products_to_excel(output_file, products_data)

    except Exception as e:
        print("------------------------------------------------------------")
        print(f"\033[91mOcurrio un error: \033[0m{e}")

if __name__ == "__main__":
    main()