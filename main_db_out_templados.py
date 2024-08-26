# -*- coding: utf-8 -*-
import pyodbc
import serial
import datetime

# Datos de Servidor
SERVER = '192.168.193.204'
DATABASE = 'OPTemplados'
USERNAME = 'sa'
PASSWORD = 'admin9233'

# Configuracion del Puerto Salida de Templados
puerto_1 = 'COM6'
baudios = 9600

# Conexion de puerto Serial a puerto 1 templados 
try:
    conexion = serial.Serial(puerto_1, baudios, timeout=1)
    print("Conexión establecida con", puerto_1 , "Salida de templados")
except serial.SerialException:
    print("No se pudo establecer la conexión con", puerto_1)

    
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

while True:
    try:
        # Leer una línea desde el puerto serial
        codebar = conexion.readline().decode('utf-8').strip()
        if codebar:
            informacion = codebar
            if filtro_Data(informacion):
                if filtro_double_Data(codebar):
                    write_db(informacion)
                
    except KeyboardInterrupt:
        print("Programa interrumpido por el usuario")
        break
    

    
# Cierra la conexión del puerto serie al finalizar
conexion.close()
conn.close()
print("Conexión cerrada")