import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from stock_valuer.src import peps
from utils.odoo_service.odoo_service import OdooService

def main():
    try:
        # Crear una instancia de OdooService
        odoo_service = OdooService()

        peps.get_peps()
        
        pass
    
    except Exception as e:
        print("------------------------------------------------------------")
        print(f"\033[91mOcurrio un error: \033[0m{e}")

if __name__ == "__main__":
    main()