# -*- coding: utf-8 -*--
import pyodbc
import serial
import datetime
from notion_client import Client
import requests
import pyodbc
import time

# Datos de Servidor
SERVER = '192.168.193.204'
DATABASE = 'OPTemplados'
USERNAME = 'sa'
PASSWORD = 'admin9233'

# Notion configuration
NOTION_TOKEN = "secret_DJgeLF4gqSFAci7vmdSje18Wjrh4iOKebK3WlUekpmv"
new_state = True

# Configuracion del Puerto Salida de Templados
puerto_1 = 'COM6'
baudios = 9600

# Deshabilitar la validación del certificado SSL temporalmente
connectionString = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};TrustServerCertificate=yes;'
# Conexion a la base de datos 
try:
    conn = pyodbc.connect(connectionString)
    print("Conexión exitosa a base de datos")
except Exception as e:
    print("Error al conectar a base de datos:", e)    
    
def timestamp():
    # Obtener la marca de tiempo actual
    marca_de_tiempo_actual = datetime.datetime.now()
    # Convertir la marca de tiempo a formato de texto
    marca_de_tiempo_texto = marca_de_tiempo_actual.strftime('%Y-%m-%d %H:%M:%S')
    # Imprimir la marca de tiempo
    print("Marca de tiempo actual:", marca_de_tiempo_texto)
    return marca_de_tiempo_texto
    
def write_db(informacion):
    codesplit = informacion.split("-")
    print("OP:", codesplit[1])
    print("ITEM:", codesplit[2])
    itemsplit = codesplit[2].split("T")
    OP = codesplit[1]
    ITEM = itemsplit[1]
    cursor = conn.cursor()
    print("Datos recibidos:", codebar)
    print("Inserting a new row into table")
    # Insert Query
    tsql = "INSERT INTO OUT_TEMPLADOS (OP,ITEM,FECHA) VALUES (?,?,?);"
    with cursor.execute(tsql,OP,ITEM,timestamp()):
        print("Successfully Inserted!")
    return OP, ITEM

# Filtro para que solo guarde informacion de la etiqueta de templados
def filtro_Data(informacion):
    codesplit = informacion.split("-")
    if codesplit[0] == "001":
        return True
    else:
        print("Dato no permitido")

# Filtro para que no guarde doble data en Sql Server
def filtro_double_Data(informacion):
    codesplit = informacion.split("-")
    itemsplit = codesplit[2].split("T")
    cursor = conn.cursor()
    consulta = "SELECT COUNT(*) FROM OUT_TEMPLADOS WHERE OP = ? AND  ITEM = ?"
    cursor.execute(consulta, codesplit[1], itemsplit[1])
    cantidad_filas = cursor.fetchone()[0]
    if cantidad_filas == 0:
        return True
    else:
        print("Los datos ya existen en la tabla. Evitando la inserción.")


#-----------------------------------------------------------------------------------------------------------
#FUNCIONES ----------------------------------------------------------------------------------------------
#-SQL READ DATA---------------------------------------------------------------------------------
def leer_Tabla_ID(OP):
    try:
        DETECTOR_cursor = conn.cursor()
        DETECTOR_cursor.execute("SELECT @@version;")
        DETECTOR_query = 'SELECT * FROM Tabla_ID where OP = '+str(OP)
        DETECTOR_cursor.execute(DETECTOR_query)
        DETECTOR_FILAS = DETECTOR_cursor.fetchall()
        filas_OP = len(DETECTOR_FILAS)
        DATA_ID = []
        if DETECTOR_FILAS != [] :
            #print("DATA SQL EXISTENTE PARA OP = ",OP)
            for i in reversed(range(filas_OP)):
                DATA_ID.append(DETECTOR_FILAS[i][4])
        else:
            print("DATA SQL NO EXISTENTE PARA OP =",OP)
    except:
        print("F_DETECTOR_CURSOR")
    finally:
        return DATA_ID[0]

def leer_filas_OP(OP):
    filas_OP = 0
    try:
        DETECTOR_cursor = conn.cursor()
        DETECTOR_cursor.execute("SELECT @@version;")
        DETECTOR_query = 'SELECT * FROM PT_V2 where OP = '+str(OP)
        DETECTOR_cursor.execute(DETECTOR_query)
        DETECTOR_FILAS = DETECTOR_cursor.fetchall()
        filas_OP = len(DETECTOR_FILAS)
    except:
        print("F_DETECTOR_CURSOR")
    finally:
        return filas_OP
    
#-NOTION UPDATE---------------------------------------------
def actualizar_datos_database(ID,ITEM):
    notion = Client(auth=NOTION_TOKEN)
    # Obtener el contenido de la base de datos
    database_content = get_database_content(ID)
    # Obtener page ID filtrado por ITEM en especifico
    ID_DB4 = get_page_ID_F_ITEM(database_content, ITEM)
    #print("ID_DB4:",ID_DB4)
    headers = {
        "Authorization": "Bearer " + NOTION_TOKEN ,
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    }
    url = f"https://api.notion.com/v1/pages/{ID_DB4}"
    new_properties = {"properties": {"9_TEMPLADO": {"checkbox": new_state}}}
    try:
        requests.patch(url, json=new_properties, headers=headers)
    except:
        print("F_UPDATE")

#-NOTION ---------------------------------------------------------------------------------
# Función para obtener todas las páginas de la base de datos
def get_database_content_antiguo(ID):
    notion = Client(auth=NOTION_TOKEN)
    query_results = notion.databases.query(database_id = ID)
    return query_results.get('results', [])

def get_database_content(ID):
    notion = Client(auth=NOTION_TOKEN)
    all_results = []
    has_more = True
    start_cursor = None

    while has_more:
        # Realiza la solicitud a la API de Notion
        query_response = notion.databases.query(
            database_id=ID,
            start_cursor=start_cursor
        )

        # Añade los resultados actuales a la lista total
        all_results.extend(query_response.get('results', []))

        # Verifica si hay más resultados y actualiza el cursor de inicio
        has_more = query_response.get('has_more', False)
        start_cursor = query_response.get('next_cursor')

    return all_results

# Función para extraer la OP
def get_page_ITEM(page):
    property_ITEM = page['properties']['1_ITEM']['title'][0]['text']['content']
    if property_ITEM is not None:
        return property_ITEM
    return 'Sin ITEM'

# Función para extraer la ID
def get_page_ID(page):
    property_ID = page['id']
    if property_ID is not None:
        return property_ID
    return 'Sin ID'

# Función para extraer la page ID filtrado por ITEM en especifico
def get_page_ID_F_ITEM(page,ITEM):
    for i in range(len(page)):
        block = page[i]
        page_ITEM = get_page_ITEM(block)
        if(page_ITEM == str(ITEM)):
            page_ID = get_page_ID(block)
            return page_ID
    return "Sin ID"
#-----------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------

while True:
    try:
        conexion = serial.Serial(puerto_1, baudios, timeout=1)
        print("Conexión establecida con", puerto_1 , "Salida de templados")
        break
    except serial.SerialException:
        print("No se pudo establecer la conexión con", puerto_1)
        time.sleep(15)

while True:
        # Conexion de puerto Serial a puerto 1 templados 
    try:
        # Leer una línea desde el puerto serial
        codebar = conexion.readline().decode('utf-8').strip()
        if codebar:
            informacion = codebar
            if filtro_Data(informacion):
                if filtro_double_Data(codebar):
                    OP,ITEM =write_db(informacion)
                    #-----------------------------------
                    ID_DB3 = leer_Tabla_ID(OP)
                    actualizar_datos_database(ID_DB3,ITEM)
                                
    except KeyboardInterrupt:
        print("Programa interrumpido por el usuario")
        break

# Cierra la conexión del puerto serie al finalizar
conexion.close()
conn.close()
print("Conexión cerrada")