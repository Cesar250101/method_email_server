# --- IMPORTS ---- Librerías importadas

import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import xmltodict
import pprint
import xmlrpc.client
import json
import csv
import base64
import time

### --------------- DATA --------------- ###

## -------- MAIL DATA -------- ##

# - Server to connect -- Servidor de correo al cual conectarse y su información de logueo
mail_server = ''
email_mail = ''
email_password = ''

# - Casilla de Inbox a leer 
inbox_mail = ''
imap_server = None

## -------- MAIL DATA-------- ##


## -------- ODOO DATA -------- ##

# - Credenciales para conectarse a odoo
url = ''
db = ''
username = ''
password = ''

# - Modelos en los que se van a realizar las búsquedas
model_message_dte_docuemnt = 'mail.message.dte.document'
model_res_partner = 'res.partner'
model_message_dte = 'mail.message.dte'
model_ir_attachment = 'ir.attachment'
model_sii_document_class = 'sii.document_class'
model_res_company = 'res.company'
model_mail_message_dte_document_line = 'mail.message.dte.document.line'
model_product_product = 'product.product'


## -------- ODOO DATA -------- ##

### --------------- DATA --------------- ###



### --------------- METODO MAIN --------------- ###

def main():

    seconds_wait = 10

    while True:
        print("Process starting....")
        # Cantidad de mensajes a leer
        messages_to_read = 15
        # Variable para decidir si leo todos los emails o solo los no leidoas
        read_all = False
        # Proceso mails del servidor y obtengo la data que necesito para enviar a odoo
        mail_data_all = process_mails(messages_to_read, read_all)

        if mail_data_all:

            # Por cada archivo debo procesar los datos
            for mail_data in mail_data_all:
                process_odoo_data(mail_data)


        else:
            print("NO HAY MAILS SIN LEER")

        print("Process done....")
        print(f"Waiting {seconds_wait} seconds....")
        time.sleep(seconds_wait) # 2 second delay

### --------------- METODO MAIN --------------- ###


### ------- METODOS -------- ###

# Método para buscar un producto por nombre
def search_get_prodcut_id_by_product_name(models, credentials, uid, model_product_product, limit, nombre_producto):
    id_result = search_odoo_get_ids_by(models, credentials, uid, model_product_product, limit, 'name', "=", nombre_producto)
    result = read_odoo_ids(models, credentials, uid, model_product_product, id_result)
    #print(f"RESULT PRODUCTS:: {result}")
    if result:
        return result[0]['id']
    else:
        return None

# Método para obtener el partner id
def get_dte_partner_id(models, credentials, uid, model_res_partner, limit, rut_id_receptor):
    id_result = get_partner_by_rut_id(models, credentials, uid, model_res_partner, limit, rut_id_receptor)
    result = read_odoo_ids(models, credentials, uid, model_res_partner, id_result)
    #print(f"RESULT PARTNER:: {result}")
    return result[0]['id']

# Método para obtener o crear para cada producto la linea de producto
def get_item_line(item, models, credentials, uid, model_mail_message_dte_document_line, model_product_product, limit, i):
    new_product = item["NmbItem"]
    product_id = search_get_prodcut_id_by_product_name(models, credentials, uid, model_product_product, limit, new_product)
    if product_id:
        payload = [{'product_id' : product_id, 'sequence': i}]
    else:
        quantity = item["QtyItem"]
        price_unit = item["PrcItem"]
        price_subtotal = item["MontoItem"] 
        payload = [{'new_product' : new_product, 'quantity' : quantity, 'price_unit' : price_unit, 'price_subtotal' : price_subtotal, 'sequence': i }]
    id_result = create_document(models, credentials, uid, model_mail_message_dte_document_line, payload)
    result = read_odoo_ids(models, credentials, uid, model_mail_message_dte_document_line, id_result)
    #print(f"RESULT LINEA:: {result}")
    return result

# Método obtener los ids de las lineas de producto para agregar a la tabla
def create_get_product_line_ids(models, credentials, uid, model_mail_message_dte_document_line, model_product_product, limit, mail_data):
    line_ids = []
    line_id = ""
    i = 1
    leen = type(mail_data["Items"]) is dict
    if type(mail_data["Items"]) is not dict:
        for item in mail_data["Items"]:
            result = get_item_line(item, models, credentials, uid, model_mail_message_dte_document_line, model_product_product, limit, i)
            #print(f"RESULT LINEA:: {result}")
            line_id = result[0]['id']
            line_ids.append(line_id)
            i = i +1
    else:
        item = mail_data["Items"]
        result = get_item_line(item, models, credentials, uid, model_mail_message_dte_document_line, model_product_product, limit, i)
        line_id = result[0]['id']
        line_ids.append(line_id)

    return line_ids

# Método para crear y obtener el DTE id que va en la tabla modelo model_message_dte_docuemnt
def create_get_dte_id(models, credentials, uid, model_message_dte, limit, attachment_id, xml_name):
    payload = [{'message_main_attachment_id' : attachment_id, 'name' : xml_name }]
    id_result = create_document(models, credentials, uid, model_message_dte, payload)
    result = read_odoo_ids(models, credentials, uid, model_message_dte, id_result)
    #print(f"RESULT DTEE:: {result}")
    return result[0]['id']

# Método para obtener el attachment id que va en la tabla modelo model_message_dte_docuemnt
def create_get_attachment_id(models, credentials, uid, model_ir_attachment, limit, mail_data, xml_name, raw_xml):
    ir_attachment = search_odoo_get_all_ids(models, credentials, uid, model_ir_attachment, limit)
    #print(f"ir_attachment:: {ir_attachment}")
    result = read_odoo_ids(models, credentials, uid, model_ir_attachment, ir_attachment)
    xml = str(raw_xml) 

    sample_string_bytes = xml.encode("ascii")
  
    base64_bytes = base64.b64encode(sample_string_bytes)
    base64_string = base64_bytes.decode("ascii")
    xml = base64_bytes.decode("ascii")
    payload = [{"name" : xml_name, "type": "url", 'datas': str(xml) }]
    #print(f"PAYLOAD:: {payload}")
    id_result = create_document(models, credentials, uid, model_ir_attachment, payload)
    result = read_odoo_ids(models, credentials, uid, model_ir_attachment, id_result)
    #print(f"RESULT:: {result}")
    return result[0]['id']

# Método para obtener el company_id que va en la tabla modelo model_message_dte_docuemnt
def get_company_id(models, credentials, uid, model_res_partner, limit, rut_id_receptor):
    global model_res_company

    partner_id = get_partner_by_rut_id(models, credentials, uid, model_res_partner, limit, rut_id_receptor)
    print(f"PARTNER ID:: {partner_id}")
    result = read_odoo_ids(models, credentials, uid, model_res_partner, 1) 
    #result = read_odoo_ids(models, credentials, uid, model_res_partner, partner_id) 
    #print(f"RESULT PARTNER:: {result}")
    commercial_partner_id = result[0]['commercial_partner_id']
    #commercial_partner_id = [1, 'YourCompany']
    #print(f"Commercial partner id:: {commercial_partner_id}")    
    result = read_odoo_ids(models, credentials, uid, model_res_company, 1)
    #result = read_odoo_ids(models, credentials, uid, model_res_company, commercial_partner_id)
    #print(f"RESULT COMPANY:: {result}")
    partner_id = result[0]['partner_id'][0]
    #print(f"Partner id:: {partner_id}")
    company_id = get_company_by_partner_id(models, credentials, uid, model_res_company, limit, partner_id)
    #print(f"COMPANY ID:: {company_id}")

    return company_id


# Método que recibe los datos a armar para la tabla modelo model_message_dte_docuemnt y arma el payload
def get_dte_document_payload(mail_data, company_id, dte_id, raw_xml, invoice_line_ids, partner_id, razon_social, date, folio, id_invoice):
    payload = {}
    monto_total = float(mail_data["Totales"]["MntTotal"])
    payload['amount'] = monto_total
    payload['company_id'] = company_id[0]
    payload['dte_id'] = dte_id
    payload['xml'] = raw_xml
    payload['invoice_line_ids'] = invoice_line_ids
    payload['partner_id'] = partner_id
    payload['new_partner'] = razon_social
    payload['date'] = date
    payload['number'] = folio
    payload['document_class_id'] = id_invoice
    return [payload]

# Procesar la data de odoo para obtener todo lo necesario y armar lo que se va a grabar en la tabla 
def process_odoo_data(mail_data):
    global model_res_partner
    global model_ir_attachment
    global model_message_dte
    global model_mail_message_dte_document_line
    global model_product_product
    global model_sii_document_class
    global model_message_dte_docuemnt

    # Obtengo las credenciales de acuerdo al rut ID del receptor 
    rut_id_receptor = get_rut_receptor(mail_data)
    #print(f"RUT ID RECEPTOR:: {rut_id_receptor}")
    credentials = get_odoo_credentials_by_rut_id(rut_id_receptor)

    if not credentials:
        return

    # Me conecto a odoo y valido acceso a los modelos
    models, credentials, uid = connect_and_validate_odoo(credentials)

    # Busco el company ID por partner ID
    limit = 3
    xml_name = mail_data["XML_name"]
    raw_xml = mail_data["Raw_xml"]
    company_id = get_company_id(models, credentials, uid, model_res_partner, limit, rut_id_receptor)
    
    # Obtengo el id de la tabla ir.attachment creada con los datos del xml
    attachment_id = create_get_attachment_id(models, credentials, uid, model_ir_attachment, limit, mail_data, xml_name, raw_xml)
    #print(attachment_id)

    # Obtendo el dte id del documento message.dte para asignar 
    dte_id = create_get_dte_id(models, credentials, uid, model_message_dte, limit, attachment_id, xml_name)
    #print(dte_id)

    # Obtengo cada producto y armo las lineas de producto. Obtengo el listado de productos
    product_line_ids = create_get_product_line_ids(models, credentials, uid, model_mail_message_dte_document_line, model_product_product, limit, mail_data)
    #print(product_line_ids)
    
    # Armo la relación one to many para indexar en la tabla
    invoice_line_ids = complex_relation(6, False, product_line_ids)
    #print(invoice_line_ids)

    # Obtengo el partner id a partir del rut del emisor
    rut_id_emisor = get_rut_emisor(mail_data)
    partner_id = get_dte_partner_id(models, credentials, uid, model_res_partner, limit, rut_id_emisor) #Cambiar por el del emisor en prod
    #print(partner_id)

    #Razon social del emisor a new partner --> RznSoc --> new_partner
    razon_social = mail_data['Emisor']['RznSoc']

    # Fecha de emisión del docuemnto en date --> FchEmis --> date
    date = mail_data['IdDoc']['FchEmis']

    #Folio en number
    folio = mail_data['IdDoc']['Folio']

    #TipoDTE
    # Obtener la factura por el id de factura
    tipo_dte = mail_data["IdDoc"]["TipoDTE"]
    limit = 1
    #tipo_dte = '636'
    #ids = search_odoo_get_all_ids(models, credentials, uid, model_sii_document_class, limit)
    ids = get_invoice_by_tipo_dte(models, credentials, uid, model_sii_document_class, limit, tipo_dte)
    id_invoice = ids[0]
    result = read_odoo_ids(models, credentials, uid, model_sii_document_class, ids)

    payload = get_dte_document_payload(mail_data, company_id, dte_id, raw_xml, invoice_line_ids, partner_id, razon_social, date, folio, id_invoice)
    #print(f"PAYLOAD:: {payload}")
    id_result = create_document(models, credentials, uid, model_message_dte_docuemnt, payload)
    result = read_odoo_ids(models, credentials, uid, model_message_dte_docuemnt, id_result)
    if result:
        print(f"Se ha creado correctamente un documento con el ID: {result[0]['id']} ")
    else:
        print(f"Se ha producido un error al intentar crear el documento")

# Método para obtener las credenciales a las cuales conectarse a odoo a partir del rut id del receptor
def get_odoo_credentials_by_rut_id(rut_id_receptor):
    credentials = {}
    rut = ""
    with open('db-rut-url.csv') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                #print(f'Column names are {", ".join(row)}')
                line_count += 1
            else:
                rut = row[4]
                if(str(rut) == str(rut_id_receptor)):
                    credentials['url'] = row[0]
                    credentials['db'] = row[1]
                    credentials['username'] = row[2]
                    credentials['password'] = row[3]
                    break
                line_count += 1
        if credentials:
            return credentials
        else:
            print(f"Error, no hay servidor de odoo definido para el RUT del emisor del docuemnto adjunto: {rut}")
        #print(f'Processed {line_count} lines.')

# Método para conectarme a odoo y validar acceso a los modelos a utilizar
def connect_and_validate_odoo(credentials):
    set_odoo_credentials(credentials['url'], credentials['db'], credentials['username'], credentials['password'])
    # Me conecto con odoo
    models, credentials, uid = connect_odoo()
    # Verifico que tengo acceso a los modelos que quiera acceder
    global model_message_dte_docuemnt
    #global model_res_partner
    #global model_message_dte 
    #global model_ir_attachment
    #global model_sii_document_class
    #global model_res_company 
    validate_access_right(models, credentials, uid, model_message_dte_docuemnt)
    #validate_access_right(models, credentials, uid, model_res_partner)
    #validate_access_right(models, credentials, uid, model_message_dte)
    #validate_access_right(models, credentials, uid, model_ir_attachment)
    #validate_access_right(models, credentials, uid, model_sii_document_class)

    return models, credentials, uid

# Método que devuelve los datos de las credenciales de odoo de test
def get_odoo_credentials_test():
    url = "http://erp.method.cl"
    db = "desarrollo"
    username = 'valentina.eli.88@gmail.com'
    password = '1234'
    return { 'url': url, 'db': db, 'username' : username, 'password' : password}

# Método que devuelve los datos de las credenciales de odoo de producción
def get_odoo_credentials_prod():
    url = "http://desarrollo.method.cl"
    db = "method"
    username = 'valentina.eli.88@gmail.com'
    password = '123456Ve!'
    return { 'url': url, 'db': db, 'username' : username, 'password' : password}


# Método para setear las credenciales de odoo enviadas por parámetro
def set_odoo_credentials(url_local, db_local, username_local, password_local):
    global url 
    global db 
    global username
    global password
    url = url_local
    db = db_local
    username = username_local
    password = password_local

# Método para obtener las credenciales de odoo para conectarse (preciamente seteadas)
def get_odoo_credentials():
    global url 
    global db 
    global username
    global password
    return { 'url': url, 'db': db, 'username' : username, 'password' : password}


# Método para conectarse a odoo con las credenciales 
def connect_odoo():
        credentials = get_odoo_credentials()
        #print(credentials)
        url = credentials["url"]
        username = credentials["username"]
        password = credentials["password"]
        db = credentials["db"]
        common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        # Autenticación del usuario
        uid = common.authenticate(db, username, password, {})
         
        #print(f"Version del servidor: {common.version()}")
        #print(f"User ID: {uid}")
        print(f"User Credentials: URL --> {url}, DB --> {db}, USERNAME --> {username}, PASSWORD --> {password}")

        # Me conecto con el servidor para poder hacer luego las consultas
        models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(credentials["url"]))
        #print(models)
        return models, credentials, uid

# Método para validar permiso de acceso al modelo enviado por parámetro
def validate_access_right(models, credentials, uid, model_to_check):
    result = models.execute_kw(credentials["db"], uid, credentials["password"], model_to_check, 'check_access_rights', ['read'], {'raise_exception': False})
    print(f"Access to model - {model_to_check} - : {result}")

# Método para obtener por búsqueda todos los datos de un modelo (con un límite enviado por parámetro pero sin filtros)
def search_odoo_get_all_ids(models, credentials, uid, model_to_search, limit): 
    # Obtengo los ids de la busqueda por modelo (el límite es la cantidad de registros que se obtienen)
    ids = models.execute_kw(credentials["db"], uid, credentials["password"], model_to_search, 'search',[[]], {'limit': limit})
    return ids

# Método para traer los datos de un modelo enviado por parámetro
def read_odoo_ids(models, credentials, uid, model_to_search, ids): 
    result = models.execute_kw(credentials["db"], uid, credentials["password"], model_to_search, 'read', [ids])
    # Si obtengo un resultado de la búsqueda lo devuelvo, sino mensaje de error
    if result:
        return result
    else:
        if (model_to_search == 'product.product'):
             print(f"No se encontró producto para el ID: {ids}, se creará la linea con la data correspondiente")
             return None
        else:
            print(f"Error - No hay resultado. Pruebe con otro modelo :{model_to_search}, IDS: {ids}")
            return None

# Método que itera sobre el resultado de una busqueda y muestra el resultado en consola
def iterate_print_result(result):
    # Para el resultado de cada id muestro los valores 
    if result:
        for res in result:
            print("----------------------------")
            print_result(res)

# Método que imprime el resultado completo 
def print_result(result):
    print("Resultado Compleo ")
    print(result)
    #print_result_data(result)


# Método que imprime los valores del resultado 
def print_result_data(result):
    print(result['number']) # --> Folio
    print(result['new_partner']) # --> RE + RS
    print(result['dte_id']) # --> [algo, nombre del archivo]
    print(f"partner_id --> {result['partner_id']}") # --> [algo, razon social]
    print(result['document_class_id']) # --> PREGUNTAR
    print(result['amount']) # --> MntTotal
    print(result['invoice_line_ids']) # -->  PREGUNTAR
    print(result['company_id']) # -->  RznSocRecep
    print(result['state']) # --> MONEDA, PREGUNTAR
    print(result['invoice_id']) # --> MONEDA, PREGUNTAR
    print(result['xml']) # --> MONEDA, PREGUNTAR
    print(result['purchase_to_done']) # --> MONEDA, PREGUNTAR
    print(result['claim']) # --> MONEDA, PREGUNTAR
    print(result['claim_description']) # --> MONEDA, PREGUNTAR
    print(result['claim_ids']) # --> MONEDA, PREGUNTAR


# Método para obtener el rut del receptor en formato correcto para odoo
def get_rut_receptor(mail_data):
    rut_id = mail_data["Receptor"]["RUTRecep"]
    rut_id = rut_id[0:2]+ '.' + rut_id[2:5] + '.' + rut_id[5:10]
    #print(f"RUT Receptor: {rut_id}")
    return rut_id

# Método para obtener el rut del emisor en formato correcto para odoo
def get_rut_emisor(mail_data):
    rut_id = mail_data["Emisor"]["RUTEmisor"]
    rut_id = rut_id[0:2]+ '.' + rut_id[2:5] + '.' + rut_id[5:10]
    #print(f"RUTEmisor: {rut_id}")
    return rut_id

# Método para buscar el partner proveedor por rut id 
def get_partner_by_rut_id(models, credentials, uid, model_res_partner, limit, rut_id):
    search_parameter = 'document_number'
    search_operator = '='
    search_value = rut_id 
    return search_odoo_get_ids_by(models, credentials, uid, model_res_partner, limit, search_parameter, search_operator, search_value)

# Método para buscar la factura por tipo DTE
def get_invoice_by_tipo_dte(models, credentials, uid, model_sii_document_class, limit, tipo_dte):
    search_parameter = 'id'
    search_operator = '='
    search_value = tipo_dte 
    return search_odoo_get_ids_by(models, credentials, uid, model_sii_document_class, limit, search_parameter, search_operator, search_value)

# Método para buscar el company id por el partner id
def get_company_by_partner_id(models, credentials, uid, model_ir_attachment, limit, partner_id):
    search_parameter = 'partner_id'
    search_operator = '='
    search_value = partner_id 
    return search_odoo_get_ids_by(models, credentials, uid, model_ir_attachment, limit, search_parameter, search_operator, search_value)


# Método para buscar ids por un criterio dado
def search_odoo_get_ids_by(models, credentials, uid, model_to_search, limit, search_parameter, search_operator, search_value): 
    ids = models.execute_kw(credentials["db"], uid, credentials["password"], model_to_search, 'search',[[[search_parameter, search_operator, search_value]]], {'limit': limit})
    #print(ids)
    return ids

# Método para crear una nueva entrada en el modelo indicado por parámetros y con los datos indicados en el payloas también enviado por parámetro
def create_document(models, credentials, uid, model_to_write, payload): 
    #Payload ejemplo --> [{'dte_id': [190, "EnvioDTE-96806980-2-77216294-4-20220413-111131-58842180.xml"]}]
    id = models.execute_kw(credentials["db"], uid, credentials["password"], model_to_write, 'create', payload)
    #print(f" New ID --> {id}")
    return id


# Método para crear una nueva entrada en el modelo indicado por parámetros y con los datos indicados en el payloads también enviado por parámetro
def update_document(models, credentials, uid, model_to_write, payload): 
    #Payload ejemplo --> [[2], {'dte_id': [190, "EnvioDTE-96806980-2-77216294-4-20220413-111131-58842180.xml"]}]
    #Donde el primer valor es el ID al cual se va a modificar
    models.execute_kw(credentials["db"], uid, credentials["password"], model_to_write, 'write', payload)

# Método para generar el tipo de operación necesaria en una realción compleja de one2many o many2many.
def complex_relation(operation, id, values = 0):
    result = [(operation, id, values)]
    return result
    #Para las relaciones mas complejas 
    #(0, 0, values) Agregar un nuevo registro, donde value es un dict con los datos a crear. (Puede ir en un create)
    #(1, id, values) Modificar un registro enviando dicho id, donde value es un dict con los datos a crear. (No puede ir en un create)
    #(2, id, 0) Elimina un registro enviando dicho id. (No puede ir en un create)
    #(3, id, 0) Elimina un registro, pero solo el vínculo, no el registro en sí (No puede ir en un create)
    #(4, id, 0) Linkea o agrega un registro ya existente a la relación
    #(5, 0, 0) Elimina todos los registros. (No puede ir en un create)
    #(6, 0, ids) Reemplaza todos los registros existentes por los enviados por listado de ids


# Método para procesar los mails y obtener toda la data necesaria para enviar a odoo
def process_mails(messages_to_read, read_all):
    # Seteo la información del servidor enviada por parámetros
    #mail_server_local = 'mail.openmethod.cl'
    #email_mail_local = 'dte.method@openmethod.cl'
    #email_password_local = '2010626Ab'

    mail_server_local = 'mail.openmethod.cl'
    email_mail_local = 'dte.pitbull@openmethod.cl'
    email_password_local = '2010626Ab@'
    set_server_information(mail_server_local, email_mail_local, email_password_local)
    # Seteo el servidor para la información previamente enviada
    set_imap_server()
    # Me logueo al servidor de correo para leer los emails
    login_email_server()
    # Seteo la casilla de INBOX la cual voy a leer
    set_box("INBOX")

    if read_all:
        #Obtengo y muestro la cantidad de mensajes en el buzón de entrada
        message_quantity = get_messages_quantity()
        # Leo los archivos adjuntos para la cantidad de mensajes enviada por parámetro
        print(f"Cantidad de mensajes a leer: {messages_to_read}")
        email_data = read_all_messages(message_quantity, messages_to_read)

    else:
        print("Leer mensajes no leidos")
        email_data = read_unseen_messages()

    # Para cada Email imprimo la información 
    #print_data(email_data)
    if email_data:
        # Paso la data de XML a un diccionario en Python
        all_data = set_xml_data_to_dic(email_data)
        #print(all_data)
    else: 
        all_data = ""

    return all_data


# Método para setear la información de inicio de sesión del servidor de correos.
def set_server_information(mail_server_local, email_mail_local, email_password_local):
    global mail_server
    global email_mail
    global email_password
    mail_server = mail_server_local
    email_mail = email_mail_local
    email_password = email_password_local

# Método para obtener la información de inicio de sesión del servidor de correos.
def get_server_information():
    global mail_server
    global email_mail
    global email_password
    return { "server": mail_server,"email": email_mail, "pass": email_password}


# Método para obtener el servidor del correo (global)
def get_imap_server():
    global imap_server
    return imap_server


# Método para setear el servidor del correo (global)
def set_imap_server():
    global imap_server
    mail_server = get_server_information()["server"]
    imap_server = imaplib.IMAP4_SSL(mail_server)


# Método para loguear con usuario y contraseña especificada al servidor de correo.
def login_email_server():
    imap_server = get_imap_server()
    login_info = get_server_information()
    user = get_server_information()["email"]
    passw = get_server_information()["pass"]
    imap_server.login(user, passw)
    print(f"Sesión iniciada para el usuario: {user}, y contraseña: {passw}")


# Método para setear el nombre de la casilla de mensajes que se quiere leer
def set_box(box_name):
    global inbox_mail
    inbox_mail = box_name

# Método para obtener el nombre de la casilla de mensajes que se quiere leer
def get_box():
    global inbox_mail
    return inbox_mail


# Método de obtención de cantidad de mensajes de la casilla de entrada del correo especificado.
def get_messages_quantity():
    imap_server = get_imap_server()
    status, messages = imap_server.select(get_box())
    messages = int(messages[0])
    print(f"Cantidad de mensajes totales de la casilla {get_box()} : {messages}")
    return messages

# Método para imprimir datos de los mensaje leídos, Subject, From, XML (si no tiene archivo adjunto xml lo informa)
def print_data(email_data):
     for data in email_data:
        print("------------------------------------------------------------")
        print(f"{data}")


# Método para leer un archivo adjunto de un email
def read_attachment(msg):
    data = {}
    # Obtenfo el subject del mensaje
    subject, encoding = email.header.decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        if encoding:
            subject = subject.decode(encoding)

    if not isinstance(subject, str):
        subject = str(subject)

    data["Subject"] = subject

    # Filtro por los mensajes donde se recibe una factura DTE
    if "Envio de DTEs" not in subject: 
        data["xml"] = "No XML" 

    # Si tiene factura, y es un mje múltiple, leo las partes
    elif msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            try:
                body = part.get_payload(decode=True).decode()
            except:
                pass
            
            # Si tiene archivo adjunto, lo leo y o obtengo.
        if "attachment" in content_disposition:
            #print("Attachments found")
            filename = part.get_filename()
            
            if filename:
                my_xml = part.get_payload(decode=True)
                data["xml"] = my_xml 
                data["xml_name"] = filename 
        else:
            data["xml"] = "No XML" 

                #set_xml_data(my_xml) 
    return data


# Método que lee los mensajes no leidos
def read_unseen_messages():
    global imap_server
    email_data = []
    data = {}

    box = get_box()
    imap_server.select(mailbox=box, readonly=False)

    response, emails = imap_server.search(None, '(UNSEEN)')
    emails = emails[0].split()
    print(f"Cantidad de emails a analizar: {len(emails)}")
    if emails:
        for mail_id in emails:
            #print(f"Analizando Email ID: {mail_id}...")
            #print(mail_id)
            res, msg = imap_server.fetch(mail_id, "(RFC822)")
            for response in msg:
                if isinstance(response, tuple):
                    msg = email.message_from_bytes(response[1])
                    data = read_attachment(msg)
            
            data_copy = data.copy()
            email_data.append(data_copy)

    else:
        email_data = ""
        
    return email_data


# Método para leer archivos adjuntos dentro de N cantidad de mensajes.
def read_all_messages(total_messages, messages_to_read):
    global imap_server
    messages = total_messages
    N = messages_to_read
    email_data = []
    data = {}

    # Recorro N mensajes
    for i in range(messages, messages-N, -1):
        res, msg = imap_server.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                data = read_attachment(msg)
                
        data_copy = data.copy()
        email_data.append(data_copy)

    return email_data


# Mapear los datos del xml a un dictionary de python para obtener los datos a ingresar en las tablas Odoo
def set_xml_data_to_dic(my_data):

    data_all = []
    data_for_odoo = {}

    for data in my_data:
        my_xml = data["xml"]
        if my_xml != "No XML":
            #print(my_xml)
            my_dict = xmltodict.parse(my_xml)
            my_dict = json.loads(json.dumps(my_dict))

            # Acá se hace el mapeo de lo que viene en el XML y al dato en el que se va a convertir en la tabla para simplificar
            data_for_odoo["Raw_xml"] = my_xml
            data_for_odoo["XML_name"] = data["xml_name"]
            data_for_odoo["Emisor"] = my_dict['EnvioDTE']['SetDTE']['DTE']['Documento']['Encabezado']['Emisor']
            data_for_odoo["IdDoc"] = my_dict['EnvioDTE']['SetDTE']['DTE']['Documento']['Encabezado']['IdDoc']
            data_for_odoo["Receptor"] = my_dict['EnvioDTE']['SetDTE']['DTE']['Documento']['Encabezado']['Receptor']
            data_for_odoo["Totales"] = my_dict['EnvioDTE']['SetDTE']['DTE']['Documento']['Encabezado']['Totales']
            data_for_odoo["Items"] = my_dict['EnvioDTE']['SetDTE']['DTE']['Documento']['Detalle']
            data_for_odoo["TED-DA"] = my_dict['EnvioDTE']['SetDTE']['DTE']['Documento']['TED']['DD']['CAF']['DA']
            
            #print("----------------------------")
            #print(f"Emisor: {data_for_odoo['Emisor']}")
            #print("")
            #print(f"Receptor: {data_for_odoo['Receptor']}")
            #print("")
            #print(f"Items: {data_for_odoo['Items']}")
            #print("")
            #print(f"XML name: {data_for_odoo['XML_name']}")
            data_for_odoo_copy = data_for_odoo.copy()
            data_all.append(data_for_odoo_copy)
        else:
            print("No hay factura")

    return data_all


### ------- METODOS -------- ###
            

if __name__ == "__main__":
    main()