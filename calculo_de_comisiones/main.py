import sys
import os
from datetime import datetime
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import utils.odoo_service.odoo_service as odoo_Service
import utils.google_service.google_service as google_service
import utils.helpers as helpers
import calculo_de_comisiones.src.write_service as xlsx_services
import calculo_de_comisiones.src.select_user as select_user

def main():
    print("------------------------------------------------------------")
    print("\033[94m --> Proceso Iniciado <-- \033[0m")
    print("------------------------------------------------------------")

    try:
        odoo_service = odoo_Service.OdooService()
        drive_service = google_service.authenticate_google_drive()
        drive_folder_id = '16ePynRNqNhWSqspQ3tMHyTJPbFCvPllr'

        month = 6 #datetime.now().month
        month_start, month_end = helpers.getMonthRange(month) # Get the month range

        print("------------------------------------------------------------")

        for seller in range(1, len(select_user.get_sellers()) + 1):
            seller_name, template_file = select_user.select_user(seller)

            # Get the published invoices for the seller in the given month
            invoices = odoo_service.getInvoiceFields(seller_name, month_start, month_end, ['id', 'invoice_date', 'name', 'invoice_origin', 'medium_id', 'partner_id', 'invoice_line_ids'])
            
            print(f"Invoices del vendedor en el mes {month}: {len(invoices)}")

            # Si no hay facturas, continuar con el siguiente vendedor
            if len(invoices) <= 0:
                print(f"0 facturas encontradas para {seller_name} en el mes {month}.")
                print("------------------------------------------------------------")
                continue

            # Aggregate data to invoices
            invoices = odoo_service.agregateDataToInvoices(invoices)
        
            # Write the report
            report_stream = xlsx_services.writeXlsx(invoices, template_file) 
            if seller <= 2: 
                report_stream = xlsx_services.write_instalaciones_Xlsx(report_stream)

            # Subir a Google Drive
            now = datetime.now()
            timestamp = now.strftime("%Y/%m/%d_%H:%M:%S")
            file_name = f'cc_report_{seller_name.replace(" ", "_")}_{timestamp}.xlsx'
            google_service.upload_to_drive(drive_service, report_stream, drive_folder_id, file_name)

            print(f"\033[92mReporte del mes - {month} - para: {seller_name} cargado con extio en DRIVE\033[0m")
            print("------------------------------------------------------------")

    except Exception as e:
        print("------------------------------------------------------------")
        print(f"\033[91mOcurrio un error: \033[0m{e}")
    
    print("------------------------------------------------------------")
    print("\033[94m --> Proceso Finalizado <-- \033[0m")
    print("------------------------------------------------------------")

if __name__ == "__main__":
    main()