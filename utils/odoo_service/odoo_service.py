import xmlrpc.client
import os
from dotenv import load_dotenv
import time
load_dotenv()

class OdooService:
    def __init__(self):
        self.url = os.getenv('url')
        self.db = os.getenv('db')
        self.odoo_username = os.getenv('odoo_username')
        self.password = os.getenv('password')
        self.api_key = os.getenv('api_key')

        # Conexión común
        self.common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(self.url))

        self.uid = self.common.authenticate(self.db, self.odoo_username, self.password, {})

        if not self.uid:
            print("\033[91mAutenticación con Odoo fallida. Verifique sus credenciales.\033[0m")
        else:
            print("\033[92mConexion con Odoo exitosa\033[0m")
        
        # Conexión a modelos
        self.models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(self.url))

    # ---------------------------------------------------------------------------------
    # -- Calculo de comisiones --------------------------------------------------------
    # --------------------------------------------------------------------------------- 

    def getInvoiceFields(self, seller_name, month_start, month_end, fields):
        try:
            invoice_ids = self.request_manager('account.move', 'search', [[['invoice_user_id', '=', seller_name], ['invoice_date', '>=', month_start], ['invoice_date', '<=', month_end], ['type_name', '=', 'invoice'], ['state', '=', 'posted']]])
            invoices = self.request_manager('account.move', 'read', [invoice_ids], {'fields': fields})
            
            return invoices
        
        except xmlrpc.client.Fault as e:
            print(f"\033[91mError al obtener las facturas: \033[0m{e}")
            print("------------------------------------------------------------")
            return []

    def agregateDataToInvoices(self, invoices):
        invoices_agregated = []
        cont = 0
        for invoice in invoices:
            cont += 1

            # if cont >= 3:
            #     break

            print(f" - Cargando datos de factucras: {cont} de {len(invoices)}")
            invoices_agregated.append(self.get_invoice_info(invoice))
            time.sleep(2)  # wait 2 seconds to avoid rate limiting
        return invoices_agregated
    
    def get_invoice_info(self, invoice):
        invoice_info = self.get_invoice_basic_info(invoice)
        invoice_info = self.get_invoice_order_info(invoice, invoice_info)
        invoice_info = self.get_invoice_products(invoice, invoice_info)
        return invoice_info
    
    def get_invoice_basic_info(self, invoice):
        invoice_info = {}
        invoice_info['invoice_number'] = invoice['name']
        invoice_info['invoice_origin'] = invoice['invoice_origin']
        invoice_info['invoice_date'] = invoice['invoice_date']
        invoice_info['invoice_exchange_rate'] = self.getExchangeRate(invoice_info['invoice_date'])
        invoice_info['invoice_medium'] = self.get_invoice_medium(invoice)
        invoice_info['invoice_partner'] = invoice['partner_id'][1]
        return invoice_info
    
    def getExchangeRate(self, date):
        records = self.request_manager('res.currency.rate', 'search_read', [[['rate', '<', 1], ['name', '>=', date]]], {'fields': ['inverse_company_rate'], 'order': 'name asc', 'limit': 1})
        
        # Verificar si records no está vacío antes de acceder al índice 0
        if records:
            return records[0]['inverse_company_rate']
        else:
            # Manejar el caso donde no hay registros disponibles
            print(f"\033[91mNo exchange rate found for date: {date}\033[0m")
            return None  # O un valor por defecto adecuado para tu caso
    
    def get_invoice_medium(self, invoice):
        if invoice['medium_id']:
            return invoice['medium_id'][1]
        else:
            return 'Not Set'
    
    def get_invoice_order_info(self, invoice, invoice_info):
        invoice_info['purchase_order'] = self.get_purchase_order_display_name(invoice['invoice_origin'])
        invoice_info['paid_through_stripe'] = self.has_order_been_paid_via_stripe(invoice['invoice_origin'])
        return invoice_info
    
    def get_purchase_order_display_name(self, origin):
        # Find the purchase order ID based on the invoice origin
        purchase_id = self.request_manager('purchase.order', 'search', [[['origin', '=', origin]]])

        if not purchase_id:
            # print(f"No purchase order found with invoice origin {origin}")
            return

        # Read the display name of the purchase order
        purchase = self.request_manager('purchase.order', 'read', [purchase_id], {'fields': ['display_name']})

        if not purchase:
            # print(f"No purchase order found with ID {purchase_id[0]}")
            return
        
        display_names = [item['display_name'] for item in purchase]

        return display_names
    
    def has_order_been_paid_via_stripe(self, order_name):
        # Retrieve the sale order ID based on the name
        sale_order_id = self.request_manager('sale.order', 'search', [[['name', '=', order_name]]], {'limit': 1})

        if not sale_order_id:
            print(f"\033[91mNo sale order found with name {order_name}\033[0m")
            return False

        # Retrieve the payment transactions associated with the sale order
        payment_transactions = self.request_manager('payment.transaction', 'search_read', [[['sale_order_ids', '=', sale_order_id[0]]]])

        if not payment_transactions:
            return False

        # Check if any of the payment transactions were processed by Stripe and marked as successful
        for transaction in payment_transactions:
            if transaction['acquirer_id'][1] == 'Stripe' and transaction['state'] == 'done':
                return True

        return False
    
    def get_invoice_products(self, invoice, invoice_info):
        line_ids = invoice['invoice_line_ids']
        products = []
        discount = 0
        sale_order_lines = self.request_manager('sale.order.line', 'search_read', [[['order_id.name', '=', invoice_info['invoice_origin']]]])

        for line_id in line_ids:
            product = self.process_invoice_line(line_id, invoice_info)
            if product:
                product['discount'] = 0
                for sale_order_line in sale_order_lines:
                    if line_id in sale_order_line['invoice_lines']:
                        discount = sale_order_line.get('discount', 0) / 100
                        product['discount'] = discount if discount else 0
                        break
                products.append(product)
            
        invoice_info['products'] = products
        return invoice_info
    
    def process_invoice_line(self, line_id, invoice_info):
        line = self.request_manager('account.move.line', 'read', [line_id], {'fields': ['product_id', 'name', 'quantity', 'price_unit', 'currency_id']})
        if line and line[0]['product_id']:
            product = self.build_product_info(line_id, line[0], invoice_info)
            return product
        return None
    
    def build_product_info(self, line_id, line, invoice_info):
        product = {}
        product_id = line['product_id'][0]
        product['line_id'] = line_id
        product['name'] = line['name']
        product['quantity_sale'] = line['quantity']
        product['sale_price'] = line['price_unit']
        product['sale_currency'] = line['currency_id'][1]
        product['product_reference'] = self.request_manager('product.product', 'read', [product_id], {'fields': ['default_code']})[0]['default_code']
        product['stock_info'] = self.get_stock_valuation_layers(product_id)
        self.set_product_stock_info(product)
        self.set_purchase_order_info(product, invoice_info)
        return product
    
    def get_stock_valuation_layers(self, product_id):
        records = self.request_manager('stock.valuation.layer', 'search_read', [[['product_id', '=', product_id],['x_studio_costo_movimiento', '>', 0]]], {'fields': ['x_studio_costo_movimiento','x_studio_divisa'], 'order': 'create_date desc', 'limit': 1})

        if not records:
            # get the product cost from the purchase order line
            records = self.request_manager('purchase.order.line', 'search_read', [[['product_id', '=', product_id]]], {'fields': ['price_unit','currency_id'], 'order': 'create_date desc', 'limit': 1})

            if not records:
                return []
            else:
                records[0]['x_studio_costo_movimiento'] = records[0]['price_unit']
                records[0]['x_studio_divisa'] = records[0]['currency_id']
        return records
    
    def set_product_stock_info(self, product):
        if product['stock_info'] == []:
            product['price_stock'] = None
            product['stock_currency'] = None
        else:
            product['price_stock'] = product['stock_info'][0]['x_studio_costo_movimiento']
            product['stock_currency'] = product['stock_info'][0]['x_studio_divisa'][1]
    
    def set_purchase_order_info(self, product, invoice_info):
        product['purchase_order_price'] = None
        product['purchase_order_currency'] = None

        if not invoice_info['purchase_order']:
            return

        purchase_orders = invoice_info['purchase_order']

        # Extrae el valor entre corchetes de product['name']
        # Extrae el valor entre corchetes de product['name']
        if '[' in product['name'] and ']' in product['name']:
            product_reference = product['name'].split('[')[1].split(']')[0]
        else:
            return

        for order in purchase_orders:
            # Obtén el ID de la orden de compra desde Odoo
            purchase_order_ids = self.request_manager('purchase.order', 'search', [[['name', '=', order]]])

            if not purchase_order_ids:
                continue

            # Obtén los productos de la orden de compra
            order_line_ids = self.request_manager('purchase.order.line', 'search', [[['order_id', 'in', purchase_order_ids]]])
            order_products = self.request_manager('purchase.order.line', 'read', [order_line_ids], {'fields': ['name', 'price_unit', 'currency_id', 'product_qty']})

            # Itera sobre los productos de la orden de compra
            for order_product in order_products:
                if 'name' in order_product:
                    if '[' in order_product['name'] and ']' in order_product['name']:
                        order_product_reference = order_product['name'].split('[')[1].split(']')[0]
                    else:
                        continue

                    # Compara las referencias de los productos
                    if product_reference == order_product_reference:
                        product['purchase_order_price'] = order_product['price_unit'] + self.get_flete(order_products)
                        product['purchase_order_currency'] = order_product['currency_id'][1]  # assuming currency_id is a Many2one field, and you need the currency name
                        break

    def get_flete(self, order_products):
        flete = 0
        cantidad = 0

        for order_product in order_products:
            if "flete" in order_product['name'].lower():
                flete += order_product['price_unit']
            
            else:
                cantidad += order_product['product_qty']
        
        flete = flete / cantidad

        return flete
    
    # ---------------------------------------------------------------------------------
    # -- Reportes de compra y venta acumulada -----------------------------------------
    # --------------------------------------------------------------------------------- 

    def get_entire_year_invoices(self, start_date, end_date):
        move_ids = self.request_manager('account.move', 'search', [[['invoice_date', '>=', start_date], ['invoice_date', '<=', end_date], ['move_type', 'in', ['out_invoice']]]])
        
        if move_ids:
            fields = ['display_name', 'invoice_date', 'amount_total', 'line_ids', 'sequence_prefix', 'state', 'move_type', 'currency_id', 'user_id', 'payment_state', 'invoice_origin', 'invoice_partner_display_name']
            invoices = self.request_manager('account.move', 'read', [move_ids, fields])
            return invoices

    def request_manager(self, model, method, *args, **kwargs):
        max_retries = 5
        wait_time = 2 

        for attempt in range(max_retries):
            try:
                return self.models.execute_kw(self.db, self.uid, self.password, model, method, *args, **kwargs)
            except xmlrpc.client.ProtocolError as e:
                if e.errcode == 429:
                    # print(f"Too many requests, retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                    wait_time *= 2  # Exponential backoff
                else:
                    raise e

        raise Exception("Max retries exceeded")