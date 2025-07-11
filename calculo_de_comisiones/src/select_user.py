def get_sellers():
    return {
        1: "MARIO CASTRO",
        2: "EMMANUEL MURILLO",
        3: "CARLOS CASTILLO",
        4: "BRAULIO LUCERO",
        5: "CARLOS LUCERO"
    }

def select_user(seller):
    sellers = get_sellers()

    sources = {
        1: 'calculo_de_comisiones\src\Templates\cc_report_template - Mario.xlsx',
        2: 'calculo_de_comisiones\src\Templates\cc_report_template - Emmanuel.xlsx'
    }

    if seller in sellers: 
        seller_name = sellers[seller]
        print(f"Usuario seleccionado: {seller_name}")
    else:
        raise ValueError("\033[91mNo. de usuario no válido\033[0m")
    
    if seller in sources:
        source = sources[seller]
    else:
        source = 'calculo_de_comisiones\src\Templates\cc_report_template - Vendedores.xlsx'

    return seller_name, source