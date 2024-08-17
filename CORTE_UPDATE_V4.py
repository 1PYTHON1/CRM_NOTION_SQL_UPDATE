from notion_client import Client
import requests
import pyodbc

NOTION_TOKEN = "secret_QsHgUe4EMkRVgfDrd4CvhvFdwcImB0GnaUD8htBAYjw"
SERVER = '192.168.193.204'
DATABASE = 'OPTemplados'
USERNAME = 'sa'
PASSWORD = 'admin9233'
new_state = True

OP = 163
ITEM = 5

#FUNCIONES ----------------------------------------------------------------------------------------------
#-SQL READ DATA---------------------------------------------------------------------------------
def leer_Tabla_ID(OP):
    try: 
        connection_string = f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
        connection = pyodbc.connect(connection_string)
    except:
        print("F_COMUNICACION_SQL")
    try:
        DETECTOR_cursor = connection.cursor()
        DETECTOR_cursor.execute("SELECT @@version;")
        DETECTOR_query = 'SELECT * FROM Tabla_ID where OP = '+str(OP)
        DETECTOR_cursor.execute(DETECTOR_query)
        DETECTOR_FILAS = DETECTOR_cursor.fetchall()
        filas_OP = len(DETECTOR_FILAS)
        DATA_ID = []
        if DETECTOR_FILAS != [] :
            print("DATA SQL EXISTENTE PARA OP = ",OP)
            for i in reversed(range(filas_OP)):
                DATA_ID.append(DETECTOR_FILAS[i][4])
        else:
            print("DATA SQL NO EXISTENTE PARA OP =",OP)
    except:
        print("F_DETECTOR_CURSOR")
    finally:
        connection.close() # Cerrar la conexión
        return DATA_ID[0]

def leer_filas_OP(OP):
    filas_OP = 0
    try: 
        connection_string = f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
        connection = pyodbc.connect(connection_string)
    except:
        print("F_COMUNICACION_SQL")
    try:
        DETECTOR_cursor = connection.cursor()
        DETECTOR_cursor.execute("SELECT @@version;")
        DETECTOR_query = 'SELECT * FROM PT_V2 where OP = '+str(OP)
        DETECTOR_cursor.execute(DETECTOR_query)
        DETECTOR_FILAS = DETECTOR_cursor.fetchall()
        filas_OP = len(DETECTOR_FILAS)
    except:
        print("F_DETECTOR_CURSOR")
    finally:
        connection.close() # Cerrar la conexión
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
    new_properties = {"properties": {"8_CORTE": {"checkbox": new_state}}}
    try:
        requests.patch(url, json=new_properties, headers=headers)
    except:
        print("F_UPDATE")

#-NOTION ---------------------------------------------------------------------------------
# Función para obtener todas las páginas de la base de datos
def get_database_content(ID):
    notion = Client(auth=NOTION_TOKEN)
    query_results = notion.databases.query(database_id = ID)
    return query_results.get('results', [])

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


#-----------------------------------
ID_DB3 = leer_Tabla_ID(OP)
actualizar_datos_database(ID_DB3,ITEM)