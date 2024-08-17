# -*- coding: utf-8 -*-
"""
Created on Thu Apr 25 09:13:34 2024
@author: facun
"""

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QGridLayout, QLineEdit, QPushButton, QLabel, QListWidget, QListWidgetItem, QMessageBox, QComboBox, QDateEdit, QInputDialog, QAction, QMainWindow, QFileDialog
from PyQt5.QtGui import QPixmap  # Importar QPixmap
from PyQt5.QtCore import Qt, QDate
import pyodbc
from datetime import datetime
from qtmodern.styles import dark, light
from qtmodern.windows import ModernWindow
from qtawesome import icon
from PyQt5.QtWidgets import QMessageBox
import openpyxl

class DataEntryWidget(QWidget):
    def __init__(self):
        super().__init__()
        # Datos de Servidor SQL
        SERVER = '192.168.193.204'
        DATABASE = 'OPTemplados'
        USERNAME = 'sa'
        PASSWORD = 'admin9233'

        # Deshabilitar la validación del certificado SSL temporalmente
        connectionString = f'DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};TrustServerCertificate=yes;'

        # Establecer la conexión con SQL Server
        try:
            self.conn = pyodbc.connect(connectionString)
            self.cursor = self.conn.cursor()
            print("Conexión exitosa a base de datos")
        except Exception as e:
            print("Error al conectar a base de datos:", e)

        # Inicializar el contador de ítems
        self.item_counter = 1

        # Inicializar la interfaz de usuario
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Registrador Operacion de Producción')
        self.setMinimumSize(820, 920)


        self.layout = QGridLayout()

        # Agregar imagen en la parte superior
        pixmap = QPixmap("ILLATUPAC-OFICIAL-SINFONDO.png")
        self.image_label = QLabel()
        self.image_label.setPixmap(pixmap)
        self.layout.addWidget(self.image_label, 0, 0, 1, 4)
        self.image_label.setAlignment(Qt.AlignCenter)

        # Sección de datos del cliente
        self.cliente_label = QLabel('<h3><b>DATOS DEL CLIENTE</b></h3>')
        self.cliente_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.cliente_label, 1, 0, 1, 2)

        self.op_label = QLabel('OP:')
        self.op_edit = QLineEdit()
        self.layout.addWidget(self.op_label, 2, 0)
        self.layout.addWidget(self.op_edit, 2, 1)

        self.cliente_name_label = QLabel('Nombre:')
        self.cliente_name_edit = QLineEdit()
        self.layout.addWidget(self.cliente_name_label, 3, 0)
        self.layout.addWidget(self.cliente_name_edit, 3, 1)

        self.cliente_vidrio_label = QLabel('Tipo de Vidrio:')
        self.cliente_vidrio_combo = QComboBox()
        self.cliente_vidrio_combo.addItems(['Incoloro', 'Bronce', 'Gris', 'Reflejante Bronce', 'Reflejante Azul', 'Lake Blue'])
        self.layout.addWidget(self.cliente_vidrio_label, 4, 0)
        self.layout.addWidget(self.cliente_vidrio_combo, 4, 1)

        self.cliente_espesor_label = QLabel('Espesor:')
        self.cliente_espesor_combo = QComboBox()
        self.cliente_espesor_combo.addItems(['4mm', '5.5mm', '6mm', '8mm', '10mm', '12mm'])
        self.layout.addWidget(self.cliente_espesor_label, 5, 0)
        self.layout.addWidget(self.cliente_espesor_combo, 5, 1)

        self.cliente_fecha_label = QLabel('Fecha:')
        self.cliente_fecha_edit = QLineEdit()
        self.cliente_fecha_edit.setText(datetime.now().strftime('%Y-%m-%d'))
        self.cliente_fecha_edit.setReadOnly(True)
        self.layout.addWidget(self.cliente_fecha_label, 6, 0)
        self.layout.addWidget(self.cliente_fecha_edit, 6, 1)

        self.cliente_fecha_entrega_label = QLabel('Fecha de Entrega:')
        self.cliente_fecha_entrega_edit = QDateEdit()
        self.cliente_fecha_entrega_edit.setCalendarPopup(True)
        self.cliente_fecha_entrega_edit.setDate(QDate.currentDate())
        self.layout.addWidget(self.cliente_fecha_entrega_label, 7, 0)
        self.layout.addWidget(self.cliente_fecha_entrega_edit, 7, 1)

        self.cliente_total_label = QLabel('Total:')
        self.cliente_total_edit = QLineEdit()
        self.layout.addWidget(self.cliente_total_label, 8, 0)
        self.layout.addWidget(self.cliente_total_edit, 8, 1)

        self.cliente_enjabado_label = QLabel('Enjabado:')
        self.cliente_enjabado_combo = QComboBox()
        self.cliente_enjabado_combo.addItems(['No', 'Sí'])
        self.layout.addWidget(self.cliente_enjabado_label, 9, 0)
        self.layout.addWidget(self.cliente_enjabado_combo, 9, 1)

        self.submit_cliente_button = QPushButton('Agregar Cliente')
        self.submit_cliente_button.clicked.connect(self.add_cliente)
        self.layout.addWidget(self.submit_cliente_button, 10, 0, 1, 2)

        # Sección de datos del pedido
        self.pedido_label = QLabel('<h3><b>DATOS DEL PEDIDO</b></h3>')
        self.pedido_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.pedido_label, 1, 2, 1, 2)

        self.item_label = QLabel('Item:')
        self.item_edit = QLineEdit()
        self.item_edit.setText(str(self.item_counter))  # Configurar el ítem para comenzar en 1 por defecto
        self.item_edit.setReadOnly(True)
        self.layout.addWidget(self.item_label, 2, 2)
        self.layout.addWidget(self.item_edit, 2, 3)

        self.base_label = QLabel('Base:')
        self.base_edit = QLineEdit()
        self.layout.addWidget(self.base_label, 3, 2)
        self.layout.addWidget(self.base_edit, 3, 3)

        self.altura_label = QLabel('Altura:')
        self.altura_edit = QLineEdit()
        self.layout.addWidget(self.altura_label, 4, 2)
        self.layout.addWidget(self.altura_edit, 4, 3)

        self.entalle_label = QLabel('Entalle:')
        self.entalle_combo = QComboBox()
        self.entalle_combo.addItems(['No', 'Sí'])
        self.layout.addWidget(self.entalle_label, 5, 2)
        self.layout.addWidget(self.entalle_combo, 5, 3)

        self.laminado_label = QLabel('Laminado:')
        self.laminado_combo = QComboBox()
        self.laminado_combo.addItems(['No', 'Sí'])
        self.layout.addWidget(self.laminado_label, 6, 2)
        self.layout.addWidget(self.laminado_combo, 6, 3)

        self.pintado_label = QLabel('Pintado:')
        self.pintado_combo = QComboBox()
        self.pintado_combo.addItems(['No', 'Sí'])
        self.layout.addWidget(self.pintado_label, 7, 2)
        self.layout.addWidget(self.pintado_combo, 7, 3)

        self.insulado_label = QLabel('Insulado:')
        self.insulado_combo = QComboBox()
        self.insulado_combo.addItems(['No', 'Sí'])
        self.layout.addWidget(self.insulado_label, 8, 2)
        self.layout.addWidget(self.insulado_combo, 8, 3)

        self.comentario_label = QLabel('Comentario:')
        self.comentario_edit = QLineEdit()
        self.layout.addWidget(self.comentario_label, 9, 2)
        self.layout.addWidget(self.comentario_edit, 9, 3)

        self.submit_pedido_button = QPushButton('Agregar a la Lista')
        self.submit_pedido_button.clicked.connect(self.add_to_list)
        self.layout.addWidget(self.submit_pedido_button, 10, 2, 1, 2)

        self.data_list_widget = QListWidget()
        self.data_list_widget.setEditTriggers(QListWidget.DoubleClicked | QListWidget.SelectedClicked)
        self.layout.addWidget(self.data_list_widget, 11, 0, 1, 4)

        self.submit_all_button = QPushButton(icon("ei.download-alt"), 'Enviar Todos los Datos al Servidor')
        self.submit_all_button.clicked.connect(self.submit_all_data)
        self.layout.addWidget(self.submit_all_button, 12, 0, 1, 4)

        # Botón para borrar todos los datos de la lista
        self.clear_list_button = QPushButton(icon("ph.x-circle-bold"), 'Borrar Lista')
        self.clear_list_button.clicked.connect(self.clear_list)
        self.layout.addWidget(self.clear_list_button, 13, 2, 1, 2)

        # Botón para limpiar los campos de entrada del cliente
        self.clear_client_button = QPushButton(icon("mdi.account-multiple-remove"), 'Limpiar Datos del Cliente')
        self.clear_client_button.clicked.connect(self.clear_client_data)
        self.layout.addWidget(self.clear_client_button, 13, 0, 1, 2)

        self.setLayout(self.layout)

        # Conectar la señal returnPressed() de los campos de entrada de datos del pedido al método add_to_list()
        self.base_edit.returnPressed.connect(self.add_to_list)
        self.altura_edit.returnPressed.connect(self.add_to_list)
        self.comentario_edit.returnPressed.connect(self.add_to_list)
        
        self.export_excel_button = QPushButton(icon("mdi.microsoft-excel"),"Exportar a Excel")
        self.export_excel_button.clicked.connect(self.export_to_excel)
        self.layout.addWidget(self.export_excel_button, 14, 0, 1, 4)  # Ajusta las coordenadas según tu diseño

        
        
    def add_cliente(self):
        OP = self.op_edit.text().strip()
        cliente = self.cliente_name_edit.text().strip()
        vidrio = self.cliente_vidrio_combo.currentText()
        espesor = self.cliente_espesor_combo.currentText()
        fecha = self.cliente_fecha_edit.text().strip()
        fecha_entrega = self.cliente_fecha_entrega_edit.date().toString(Qt.ISODate)
        total = self.cliente_total_edit.text().strip()
        enjabado = self.cliente_enjabado_combo.currentText()
    
        if not OP or not cliente or not vidrio or not espesor or not fecha or not fecha_entrega or not total:
            QMessageBox.warning(self, 'Advertencia', 'Por favor, ingresa todos los datos para el cliente.')
            return
    
        try:
            # Verificar si la OP ya existe en la tabla CLIENTES
            self.cursor.execute("SELECT COUNT(*) FROM CLIENTES WHERE OP = ?", (OP,))
            op_count = self.cursor.fetchone()[0]
    
            if op_count > 0:
                QMessageBox.warning(self, 'Advertencia', 'La OP ya existe en la base de datos.')
                return
    
            # Insertar el nuevo cliente si la OP no existe
            self.cursor.execute("INSERT INTO CLIENTES VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'agregado')", (OP, cliente, vidrio, espesor, fecha, fecha_entrega, total, enjabado))
            self.conn.commit()
            QMessageBox.information(self, 'Éxito', 'Datos del cliente agregados correctamente.')
        except Exception as e:
            QMessageBox.warning(self, 'Error', f'Error al agregar datos del cliente: {e}')


    def add_to_list(self):
        op = self.op_edit.text().strip()
        base = self.base_edit.text().strip()
        altura = self.altura_edit.text().strip()
        entalle = self.entalle_combo.currentText()
        laminado = self.laminado_combo.currentText()
        pintado = self.pintado_combo.currentText()
        insulado = self.insulado_combo.currentText()
        comentario = self.comentario_edit.text().strip()
    
        if not op or not base or not altura:
            QMessageBox.warning(self, 'Advertencia', 'Por favor, ingresa todos los datos obligatorios.')
            return
    
        try:
            # Obtener la cantidad utilizando QInputDialog
            cantidad, ok = QInputDialog.getInt(self, 'Cantidad', 'Ingrese la cantidad:', 1, 1)
            if not ok:
                return  # El usuario canceló la entrada de cantidad
    
            for _ in range(cantidad):
                item_text = f'OP: {op}, Item: {self.item_counter}, Base: {base}, Altura: {altura}, Entalle: {entalle}, Laminado: {laminado}, Pintado: {pintado}, Insulado: {insulado}, Comentario: {comentario}'
                item = QListWidgetItem(item_text)
                item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.data_list_widget.addItem(item)
    
                # Incrementar automáticamente el número de ítems
                self.item_counter += 1
                self.item_edit.setText(str(self.item_counter))
    
            # Agregar el último valor del ítem al campo de total del cliente
            self.cliente_total_edit.setText(str(self.item_counter - 1))
    
        except ValueError:
            QMessageBox.warning(self, 'Advertencia', 'La cantidad ingresada no es válida.')
        finally:
            # Limpiar los campos de entrada después de agregar a la lista
            self.clear_inputs()
    
        # Establecer el foco en el campo "Base" después de limpiar los campos de entrada
        self.base_edit.setFocus()
    
    def submit_all_data(self):
        # Obtener los datos del cliente
        op = self.op_edit.text().strip()
        cliente = self.cliente_name_edit.text().strip()
        vidrio = self.cliente_vidrio_combo.currentText()
        espesor = self.cliente_espesor_combo.currentText()
        fecha = self.cliente_fecha_edit.text().strip()
        fecha_entrega = self.cliente_fecha_entrega_edit.date().toString(Qt.ISODate)
        total = self.cliente_total_edit.text().strip()
        enjabado = self.cliente_enjabado_combo.currentText()
        
        reply = QMessageBox.question(self, 'Confirmación', '¿Estás seguro de que deseas enviar todos los datos al servidor?',QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
       
                # Verificar que todos los campos del cliente estén completos
                if not op or not cliente or not vidrio or not espesor or not fecha or not fecha_entrega or not total:
                    QMessageBox.warning(self, 'Advertencia', 'Por favor, ingresa todos los datos para el cliente.')
                    return
            
                try:
                    # Verificar si la OP ya existe en la tabla CLIENTES
                    self.cursor.execute("SELECT COUNT(*) FROM CLIENTES WHERE OP = ?", (op,))
                    op_count = self.cursor.fetchone()[0]
            
                    if op_count > 0:
                        QMessageBox.warning(self, 'Advertencia', 'La OP ya existe en la base de datos.')
                        return
            
                    # Insertar los datos del cliente en la tabla CLIENTES
                    self.cursor.execute("INSERT INTO CLIENTES VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'agregado')",
                                        (op, cliente, vidrio, espesor, fecha, fecha_entrega, total, enjabado))
                    self.conn.commit()
            
                    # Obtener todos los elementos de la lista
                    items = [self.data_list_widget.item(i) for i in range(self.data_list_widget.count())]
            
                    # Verificar si hay entradas duplicadas antes de enviar los datos
                    for item in items:
                        item_text = item.text()
                        values = item_text.split(',')
                        op = values[0].split(':')[1].strip()
                        item_num = values[1].split(':')[1].strip()
                        self.cursor.execute("SELECT COUNT(*) FROM PEDIDOS WHERE OP = ? AND ITEM = ?", (op, item_num))
                        duplicate_count = self.cursor.fetchone()[0]
                        if duplicate_count > 0:
                            QMessageBox.warning(self, 'Error', 'Ya existe una entrada con la misma OP y el mismo item en la tabla PEDIDOS.')
                            return
            
                    # Insertar todos los datos en la tabla de PEDIDOS
                    for item in items:
                        item_text = item.text()
                        values = item_text.split(',')
                        op = values[0].split(':')[1].strip()
                        item_num = values[1].split(':')[1].strip()
                        base = values[2].split(':')[1].strip()
                        altura = values[3].split(':')[1].strip()
                        entalle = values[4].split(':')[1].strip()
                        laminado = values[5].split(':')[1].strip()
                        pintado = values[6].split(':')[1].strip()
                        insulado = values[7].split(':')[1].strip()
                        comentario = values[8].split(':')[1].strip()
            
                        # Insertar los valores en la tabla de PEDIDOS
                        self.cursor.execute("INSERT INTO PEDIDOS (OP, ITEM, BASE, ALTURA, ENTALLE, LAMINADO, PINTADO, INSULADO, COMENTARIO) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                            (op, item_num, base, altura, entalle, laminado, pintado, insulado, comentario))
                        self.conn.commit()
            
                    QMessageBox.information(self, 'Éxito', 'Todos los datos se han enviado correctamente a las tablas de CLIENTES y PEDIDOS.')
            
                except Exception as e:
                    QMessageBox.warning(self, 'Error', f'Error al enviar todos los datos: {e}')
            except Exception as e:
                QMessageBox.warning(self, 'Error', f'Error al enviar todos los datos: {e}')
        else:
            return


    def clear_list(self):
        self.data_list_widget.clear()

    def clear_client_data(self):
        self.op_edit.clear()
        self.cliente_name_edit.clear()
        self.cliente_vidrio_combo.setCurrentIndex(0)
        self.cliente_espesor_combo.setCurrentIndex(0)
        self.cliente_fecha_edit.setText(datetime.now().strftime('%Y-%m-%d'))
        self.cliente_fecha_entrega_edit.setDate(QDate.currentDate())
        self.cliente_total_edit.clear()
        self.cliente_enjabado_combo.setCurrentIndex(0)

        # Restablecer el contador de ítems
        self.item_counter = 1
        self.item_edit.setText(str(self.item_counter))       
            

    def clear_inputs(self):
        self.base_edit.clear()
        self.altura_edit.clear()
        self.comentario_edit.clear()
    
    def export_to_excel(self):
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Datos"

            # Escribir encabezados
            headers = ["OP", "Item", "Base", "Altura", "Entalle", "Laminado", "Pintado", "Insulado", "Comentario"]
            sheet.append(headers)

            # Escribir datos de la lista
            for row in range(self.data_list_widget.count()):
                item_text = self.data_list_widget.item(row).text()
                values = item_text.split(',')
                data_row = [value.split(':')[1].strip() for value in values]
                sheet.append(data_row)

            # Guardar el archivo Excel
            file_name, _ = QFileDialog.getSaveFileName(self, "Guardar archivo Excel", "", "Archivos Excel (*.xlsx)")
            if file_name:
                workbook.save(file_name)
                QMessageBox.information(self, "Éxito", "Los datos se han exportado correctamente a Excel.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error al exportar a Excel: {e}")


def set_application_style(app):
    # Establecer la hoja de estilo para la aplicación
    app.setStyleSheet("QWidget { font-size: 16px; }")  # Ajusta el tamaño de la fuente según tus preferencias

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Crear la ventana principal y establecer el widget central como DataEntryWidget
    main_window = QMainWindow()
    main_window.setCentralWidget(DataEntryWidget())
    
    # Crear una instancia de ModernWindow con la ventana principal
    window = ModernWindow(main_window)

    # Establecer el título de la ventana principal
    window.setWindowTitle('Registrador Operacion de Producción TEMPLADOS')

    # Establecer el tema claro por defecto
    light(app)

    # Establecer la hoja de estilo al iniciar la aplicación
    set_application_style(app)

    # Crear acciones de menú para cambiar el tema
    dark_theme_action = QAction("Tema Oscuro", window)
    light_theme_action = QAction("Tema Claro", window)

    # Funciones para cambiar el tema
    def change_theme_dark():
        dark(app)
        set_application_style(app)  # Llamar a la función para restablecer el estilo de la aplicación

    def change_theme_light():
        light(app)
        set_application_style(app)  # Llamar a la función para restablecer el estilo de la aplicación

    # Conectar las acciones de menú a las funciones de cambio de tema
    dark_theme_action.triggered.connect(change_theme_dark)
    light_theme_action.triggered.connect(change_theme_light)

    # Crear menú Tema y agregar acciones
    theme_menu = main_window.menuBar().addMenu("Tema")
    theme_menu.addAction(dark_theme_action)
    theme_menu.addAction(light_theme_action)

    window.show()
    sys.exit(app.exec_())


