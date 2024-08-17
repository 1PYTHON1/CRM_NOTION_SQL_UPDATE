from notion_client import Client
import requests
import pyodbc

NOTION_TOKEN = "secret_QsHgUe4EMkRVgfDrd4CvhvFdwcImB0GnaUD8htBAYjw"
SERVER = '192.168.193.204'
DATABASE = 'OPTemplados'
USERNAME = 'sa'
PASSWORD = 'admin9233'
ID_DB1 = "9742942ce3be45a7bfe38ddfb45a162f" # ID de la base de datos "LISTA DE PEDIDOS"

OP = 163
DATA_ITEM = []
DATA_BASE = []
DATA_ALTURA = []
DATA_ENTALLE= []
DATA_LAMINADO= []
DATA_PINTADO= []
DATA_INSULADO = []
DATA_COMENTARIO = []
DATA_CORTE = []
DATA_TEMPLADO = []

#-FUNCIONES-----------------------------------------------------------------------
#Funcion para leer valores de ID de SQL server
def leer_PT_V2(OP):
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
        if DETECTOR_FILAS != [] :
            print("DATA SQL EXISTENTE PARA OP =",OP)
            for i in range(filas_OP):
                DATA_ITEM.append(DETECTOR_FILAS[i][1])
                DATA_BASE.append(DETECTOR_FILAS[i][2])
                DATA_ALTURA.append(DETECTOR_FILAS[i][3])
                DATA_ENTALLE.append(DETECTOR_FILAS[i][4])
                DATA_LAMINADO.append(DETECTOR_FILAS[i][5])
                DATA_PINTADO.append(DETECTOR_FILAS[i][6])
                DATA_INSULADO.append(DETECTOR_FILAS[i][7])
                DATA_CORTE.append(DETECTOR_FILAS[i][8])
                DATA_TEMPLADO.append(DETECTOR_FILAS[i][9])
        else:
            print("DATA SQL NO EXISTENTE PARA OP =",OP)
    except:
        print("F_DETECTOR_CURSOR")
    finally:
        connection.close() # Cerrar la conexión
        return filas_OP

#-NOTION ---------------------------------------------------------------------------------
# Función para obtener todas las páginas de la base de datos
def get_database_content(ID):
    notion = Client(auth=NOTION_TOKEN)
    query_results = notion.databases.query(database_id=ID)
    return query_results.get('results', [])

# Función para extraer la OP
def get_page_OP(page):
    property_OP = page['properties']['OP']['unique_id']['number']
    if property_OP is not None:
        return property_OP
    return 'Sin OP'

# Función para extraer la ID
def get_page_ID(page):
    property_ID = page['id']
    if property_ID is not None:
        return property_ID
    return 'Sin ID'

# Función para extraer la page ID filtrado por OP en especifico
def get_ID_DB2_F_OP(ID_DB1,OP):
    content = get_database_content(ID_DB1)
    for i in range(len(content)):
        block = content[i]
        page_OP = get_page_OP(block)
        if(page_OP == OP):
            page_ID = get_page_ID(block)
            return page_ID
    return "Sin ID"

#-NOTION CREAR NUEVA DATABASE VACIA---------------------------------------------------------------------------------
# Crear la nueva base de datos vacia
def create_DB3_vacia(ID_DB2,OP):
    # Definir las estructura de DB3
    database_properties = {
        "1_ITEM": {
            "title":{}
        },
        "2_BASE": {
            "number":{}
        },
        "3_ALTURA": {
            "number":{}
        },
        "4_ENTALLE": {
            "type": "checkbox",
            "checkbox": {}
        },
        "5_LAMINADO": {
            "type": "checkbox",
            "checkbox": {}
        },
        "6_PINTADO": {
            "type": "checkbox",
            "checkbox": {}
        },
        "7_INSULADO": {
            "type": "checkbox",
            "checkbox": {}
        },
        "8_CORTE": {
            "type": "checkbox",
            "checkbox": {}
        },
        "9_TEMPLADO": {
            "type": "checkbox",
            "checkbox": {}
        }
    }
    new_database = {
        "parent": {"type": "page_id", "page_id": ID_DB2},
        "title": [{"type": "text","text": {"content": "Tabla OP "+ str(OP)}}],
        "properties": database_properties
    }
    try:
        notion = Client(auth=NOTION_TOKEN)
        response = notion.databases.create(**new_database)
        ID_DB3 = response["id"]
        return ID_DB3
    except:
        print("F_create_DB3")

#-NOTION LLENAR NUEVA DATABASE---------------------------------------------------------------------------------
# Función para llenar de datos a la nueva base de datos
def llenar_notion_DB3(filas_OP, ID):
    headers = {
        "Authorization": "Bearer " + NOTION_TOKEN ,
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28",
    }
    create_url = "https://api.notion.com/v1/pages"
    lista = list(range(filas_OP))
    try:
        for i in reversed(lista):
            payload = {"parent": {"database_id": ID},
                       "properties": {
                                "1_ITEM": {"title": [{"text": {"content": str(DATA_ITEM[i])}}]},
                                "2_BASE": {"number": DATA_BASE[i]},
                                "3_ALTURA": {"number": DATA_ALTURA[i]},
                                "4_ENTALLE":{"checkbox":False_o_True(DATA_ENTALLE[i])},
                                "5_LAMINADO":{"checkbox":False_o_True(DATA_LAMINADO[i])},
                                "6_PINTADO":{"checkbox":False_o_True(DATA_PINTADO[i])},
                                "7_INSULADO":{"checkbox":False_o_True(DATA_INSULADO[i])},
                                "8_CORTE":{"checkbox":False_o_True(DATA_CORTE[i])},
                                "9_TEMPLADO":{"checkbox":False_o_True(DATA_TEMPLADO[i])}
                        }
                    }
            requests.post(create_url, headers=headers, json=payload)
    except:
        print("F_LLENAR_DB3")

def False_o_True(x):
    if(x == "Sí"):
        return False #True
    elif(x == "No"):
        return False
    else:
        return False

#-SQL LLENAR TABLA_ID---------------------------------------------------------------------------------
def llenar_sql_Tabla_ID(OP, ID_DB1, ID_DB2, ID_DB3):
    try:
        connection_string = f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
        connection = pyodbc.connect(connection_string)
        DETECTOR_cursor = connection.cursor()
        DETECTOR_cursor.execute('''
            INSERT INTO Tabla_ID (OP, ID_DB1, ID_DB2, ID_DB3)
            VALUES (?, ?, ?, ?)
        ''', (OP, ID_DB1, ID_DB2, ID_DB3))
        # Confirmar la transacción
        connection.commit()
    except:
        print("F_TABLA_ID")
    finally:
        # Cerrar la conexión
        connection.close()

        
#---------------------------------------------
filas_OP = leer_PT_V2(OP)
ID_DB2 = get_ID_DB2_F_OP(ID_DB1,OP)
print("ID_DB2:",ID_DB2)
ID_DB3 = create_DB3_vacia(ID_DB2,OP)
print("ID_DB3:", ID_DB3)
llenar_notion_DB3(filas_OP, ID_DB3)
llenar_sql_Tabla_ID(OP, ID_DB1, ID_DB2, ID_DB3)


