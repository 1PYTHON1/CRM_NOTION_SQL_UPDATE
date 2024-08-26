# -*- coding: utf-8 -*-
"""
Created on Fri May 24 10:51:11 2024

@author: facun
"""

import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton,
                             QTextEdit, QGridLayout, QComboBox, QMessageBox)
import pyodbc
from datetime import datetime
import qtmodern.styles
import qtmodern.windows
from qtawesome import icon  # Importamos qtawesome

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Registrador de Reproceso Templados")
        self.setGeometry(100, 100, 600, 350)

        self.layout = QGridLayout()
        self.setLayout(self.layout)

        self.cliente_data = None
        self.pedido_data = None

        self.create_widgets()

        # Agregar indicador de conexión
        self.connection_label = QLabel("Conexión a Servidor 192.168.193.204: Desconectado")
        self.layout.addWidget(self.connection_label, 0,0,1,4)

    def update_connection_status(self, connected):
        if connected:
            self.connection_label.setText("Conexión a Servidor 192.168.193.204: Conectado")
            self.connection_label.setStyleSheet("color: green;")
        else:
            self.connection_label.setText("Conexión a Servidor 192.168.193.204: Desconectado")
            self.connection_label.setStyleSheet("color: red;")

    def create_widgets(self):
        self.op_label = QLabel('OP:')
        self.op_edit = QLineEdit()
        self.layout.addWidget(self.op_label, 2, 0)
        self.layout.addWidget(self.op_edit, 2, 1)

        self.item_label = QLabel('ITEM:')
        self.item_edit = QLineEdit()
        self.layout.addWidget(self.item_label, 2, 2)
        self.layout.addWidget(self.item_edit, 2, 3)

        self.cliente_label = QLabel('Cliente:')
        self.cliente_edit = QLineEdit()
        self.cliente_edit.setReadOnly(True)
        self.layout.addWidget(self.cliente_label, 3, 0)
        self.layout.addWidget(self.cliente_edit, 3, 1)

        self.vidrio_label = QLabel('Vidrio:')
        self.vidrio_edit = QLineEdit()
        self.vidrio_edit.setReadOnly(True)
        self.layout.addWidget(self.vidrio_label, 3, 2)
        self.layout.addWidget(self.vidrio_edit, 3, 3)

        self.espesor_label = QLabel('Espesor:')
        self.espesor_edit = QLineEdit()
        self.espesor_edit.setReadOnly(True)
        self.layout.addWidget(self.espesor_label, 4, 0)
        self.layout.addWidget(self.espesor_edit, 4, 1)

        self.base_label = QLabel('Base:')
        self.base_edit = QLineEdit()
        self.base_edit.setReadOnly(True)
        self.layout.addWidget(self.base_label, 4, 2)
        self.layout.addWidget(self.base_edit, 4, 3)

        self.altura_label = QLabel('Altura:')
        self.altura_edit = QLineEdit()
        self.altura_edit.setReadOnly(True)
        self.layout.addWidget(self.altura_label, 5, 0)
        self.layout.addWidget(self.altura_edit, 5, 1)

        self.motivo_reproceso_label = QLabel('Motivo:')
        self.motivo_reproceso_combo = QComboBox()
        self.motivo_reproceso_combo.addItems(['', 'Corte', 'Canteadora', 'Entalle', 'Templadora',
                                              'Lavadora', 'Despacho', 'Templado Pulido', 'Caballete', 'Calidad'])
        self.layout.addWidget(self.motivo_reproceso_label, 5, 2)
        self.layout.addWidget(self.motivo_reproceso_combo, 5, 3)

        self.comentario_label = QLabel('Comentario:')
        self.comentario_edit = QLineEdit()
        self.layout.addWidget(self.comentario_label, 6, 0)
        self.layout.addWidget(self.comentario_edit, 6, 1, 1, 3)

        self.search_button = QPushButton(icon('fa.search' , color='black'), "Buscar")
        self.search_button.clicked.connect(self.search_data)
        self.layout.addWidget(self.search_button, 7, 0, 1, 2)

        self.save_button = QPushButton(icon('fa.save', color='blue' ), "Guardar")
        self.save_button.clicked.connect(self.confirm_save)
        self.layout.addWidget(self.save_button, 7, 2, 1, 2)

        self.clear_button = QPushButton(icon('fa.times', color='blue'), "Limpiar")
        self.clear_button.clicked.connect(self.clear_fields)
        self.layout.addWidget(self.clear_button, 8, 0, 1, 4)

    def get_db_connection(self):
        SERVER = '192.168.193.204'
        DATABASE = 'OPTemplados'
        USERNAME = 'sa'
        PASSWORD = 'admin9233'
        connectionString = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};TrustServerCertificate=yes;'

        try:
            conn = pyodbc.connect(connectionString)
            self.update_connection_status(True)  # Actualizar indicador de conexión
            return conn
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"Error al conectar a la base de datos: {e}")
            self.update_connection_status(False)  # Actualizar indicador de conexión
            return None

    def search_data(self):
        op = self.op_edit.text()
        item = self.item_edit.text()

        if not op or not item:
            QMessageBox.warning(self, "Advertencia", "Por favor ingrese OP y ITEM.")
            return

        conn = self.get_db_connection()
        if not conn:
            return

        try:
            cursor = conn.cursor()

            cursor.execute("SELECT cliente, vidrio, espesor FROM CLIENTES_TEMPLADOS WHERE OP = ?", (op,))
            self.cliente_data = cursor.fetchone()

            cursor.execute("SELECT BASE, ALTURA FROM PEDIDOS_TEMPLADOS WHERE OP = ? AND ITEM = ?", (op, item))
            self.pedido_data = cursor.fetchone()

            if self.cliente_data:
                self.cliente_edit.setText(self.cliente_data[0])
                self.vidrio_edit.setText(self.cliente_data[1])
                self.espesor_edit.setText(str(self.cliente_data[2]))  # Convertir a cadena si es necesario
            else:
                self.cliente_edit.setText("No se encontraron datos")
                self.vidrio_edit.setText("No se encontraron datos")
                self.espesor_edit.setText("No se encontraron datos")

            if self.pedido_data:
                self.base_edit.setText(str(self.pedido_data[0]))  # Convertir a cadena si es necesario
                self.altura_edit.setText(str(self.pedido_data[1]))  # Convertir a cadena si es necesario
            else:
                self.base_edit.setText("No se encontraron datos")
                self.altura_edit.setText("No se encontraron datos")

            conn.close()
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"Error al ejecutar la consulta: {e}")
            self.update_connection_status(False)  # Actualizar indicador de conexión
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error: {e}")
            self.update_connection_status(False)  # Actualizar indicador de conexión

    def confirm_save(self):
        if not self.cliente_data or not self.pedido_data:
            QMessageBox.warning(self, "Advertencia", "No hay datos para guardar. Realice una búsqueda primero.")
            return

        reply = QMessageBox.question(self, 'Confirmar Guardado', 
                                     '¿Está seguro de que desea guardar los datos?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.save_data()

    def save_data(self):
        op = self.op_edit.text()
        item = self.item_edit.text()
        motivo = self.motivo_reproceso_combo.currentText()
        comentario = self.comentario_edit.text()
        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        conn = self.get_db_connection()
        if not conn:
            return

        try:
            cursor = conn.cursor()

            cursor.execute("""
                INSERT INTO REPROCESO (OP, ITEM, CLIENTE, VIDRIO, ESPESOR, BASE, ALTURA, MOTIVO, COMENTARIO, FECHA)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (op, item, self.cliente_data[0], self.cliente_data[1], self.cliente_data[2],
                  self.pedido_data[0], self.pedido_data[1], motivo, comentario, fecha_actual))

            conn.commit()
            conn.close()

            QMessageBox.information(self, "Éxito", "Datos guardados exitosamente.")
            self.clear_fields()
        except pyodbc.Error as e:
            QMessageBox.critical(self, "Error", f"Error al ejecutar la inserción: {e}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error: {e}")

    def clear_fields(self):
        self.op_edit.clear()
        self.item_edit.clear()
        self.cliente_edit.clear()
        self.vidrio_edit.clear()
        self.espesor_edit.clear()
        self.base_edit.clear()
        self.altura_edit.clear()
        self.motivo_reproceso_combo.setCurrentIndex(0)
        self.comentario_edit.clear()
        self.cliente_data = None
        self.pedido_data = None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    qtmodern.styles.light(app)
    window = MainWindow()
    mw = qtmodern.windows.ModernWindow(window)
    mw.show()
    sys.exit(app.exec_())

