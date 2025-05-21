import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView, QLineEdit,
    QStyledItemDelegate, QFileDialog
)
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtCore import Qt
from balance_engine import calcular_balance, calcular_ratios, exportar_balance_profesional, generar_diagnostico

class CustomDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setStyleSheet("""
            QLineEdit {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                padding: 4px;
            }
            QLineEdit:focus {
                border: 1px solid #007bff;
            }
        """)
        return editor

class CoreXUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CoreX - Balance Contable")
        self.setGeometry(150, 60, 1000, 700)
        self.setStyleSheet("""
            QWidget {
                background-color: #f8f7f6;
                border-radius: 15px;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 10px;
            }
            QTableWidget::item {
                padding: 6px;
            }
            QTableWidget::item:disabled {
                color: #666666;
                background-color: #f5f5f5;
            }
            QTableWidget::item:selected {
                background-color: #e6f3ff;
            }
            QLineEdit {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        self.manual_data = pd.DataFrame(columns=["categoria", "tipo", "valor"], dtype=object)
        self.manual_data["valor"] = self.manual_data["valor"].astype("float64")
        self.balance_final = None
        self.init_ui()

    def init_ui(self):
        font_title = QFont("Arial", 24, QFont.Bold)
        font_button = QFont("Arial", 14)
        font_table = QFont("Arial", 12)
        font_label = QFont("Arial", 12)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(20, 20, 20, 20)

        title_label = QLabel("CoreX - Balance Contable")
        title_label.setFont(font_title)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #1a1a1a;")
        layout.addWidget(title_label)

        input_layout = QHBoxLayout()
        self.input_empresa = QLineEdit()
        self.input_empresa.setPlaceholderText("Nombre de la Empresa")
        self.input_empresa.setFont(font_label)
        self.input_empresa.setToolTip("Ingrese el nombre de la empresa")
        input_layout.addWidget(self.input_empresa)

        self.input_fecha = QLineEdit()
        self.input_fecha.setPlaceholderText("Fecha del Balance (ej: 15/05/2025)")
        self.input_fecha.setFont(font_label)
        self.input_fecha.setToolTip("Ingrese la fecha del balance")
        input_layout.addWidget(self.input_fecha)
        layout.addLayout(input_layout)

        button_layout = QHBoxLayout()
        self.btn_generar = self.crear_boton("📊 Generar Balance", "#17a2b8", font_button, self.generar_balance)
        self.btn_generar.setToolTip("Calcular balance, totales y análisis financiero")
        self.btn_guardar = self.crear_boton("💾 Guardar en Excel", "#dc3545", font_button, self.guardar_balance)
        self.btn_guardar.setToolTip("Exportar balance a Excel")

        button_layout.addWidget(self.btn_generar)
        button_layout.addWidget(self.btn_guardar)
        layout.addLayout(button_layout)

        self.table = QTableWidget()
        self.table.setFont(font_table)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Categoría", "Tipo", "Valor"])
        self.table.setStyleSheet("""
            QHeaderView::section {
                background-color: #f1f1f1;
                color: #333333;
                padding: 8px;
                font-weight: bold;
                border: 1px solid #e0e0e0;
            }
            QTableWidget {
                gridline-color: #e0e0e0;
                selection-background-color: #e6f3ff;
            }
        """)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setFixedHeight(40)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.DoubleClicked)
        self.table.setItemDelegateForColumn(2, CustomDelegate(self))
        self.table.itemChanged.connect(self.on_item_changed)
        layout.addWidget(self.table)

        self.diagnostico_label = QLabel("Diagnóstico financiero: (pendiente)")
        self.diagnostico_label.setFont(font_label)
        self.diagnostico_label.setAlignment(Qt.AlignLeft)
        self.diagnostico_label.setStyleSheet("color: #333333; padding: 10px; background-color: #f1f1f1; border-radius: 8px;")
        layout.addWidget(self.diagnostico_label)

        self.setLayout(layout)
        self.init_table()

    def crear_boton(self, texto, color, fuente, accion):
        boton = QPushButton(texto)
        boton.setFont(fuente)
        boton.setMinimumHeight(50)
        boton.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 12px;
            }}
            QPushButton:hover {{
                background-color: {self.oscurecer_color(color)};
            }}
            QPushButton:pressed {{
                background-color: {self.oscurecer_color(color, 0.7)};
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
                color: #666666;
            }}
        """)
        boton.clicked.connect(accion)
        return boton

    def oscurecer_color(self, hex_color, factor=0.85):
        color = QColor(hex_color)
        return color.darker(int(100 / factor)).name()

    def init_table(self):
        data = [
            ("ACTIVOS", "", ""),
            ("Activos Corrientes", "", "0.00"),
            ("", "Efectivo", "0.00"),
            ("", "Cuentas por Cobrar", "0.00"),
            ("", "Bancos", "0.00"),
            ("", "Inventarios", "0.00"),
            ("", "Pagos Anticipados", "0.00"),
            ("Activos No Corrientes", "", "0.00"),
            ("", "Equipo de Oficina", "0.00"),
            ("", "Activos Intangibles (Licencias)", "0.00"),
            ("PASIVOS", "", ""),
            ("Pasivos Corrientes", "", "0.00"),
            ("", "Cuentas por Pagar", "0.00"),
            ("", "Proveedores", "0.00"),
            ("", "Impuestos por Pagar", "0.00"),
            ("", "Acreedores Diversos", "0.00"),
            ("Pasivos No Corrientes", "", "0.00"),
            ("", "Bonos por Pagar", "0.00"),
            ("", "Provisiones Laborales (Pensiones, Indemnizaciones)", "0.00"),
            ("PATRIMONIO", "", ""),
            ("Patrimonio", "", "0.00"),
            ("", "Capital Social", "0.00"),
            ("", "Reserva Legal", "0.00"),
            ("", "Resultados Acumulados", "0.00"),
            ("", "Resultado del Ejercicio", "0.00"),
            ("TOTALES", "", ""),
            ("", "Total Activos", "0.00"),
            ("", "Total Pasivos", "0.00"),
            ("", "Total Patrimonio", "0.00"),
            ("", "Total Pasivos + Patrimonio", "0.00")
        ]

        self.table.setRowCount(len(data))
        for row, (categoria, tipo, valor) in enumerate(data):
            item_categoria = QTableWidgetItem(categoria if categoria else "")
            item_tipo = QTableWidgetItem(tipo if tipo else "")
            item_valor = QTableWidgetItem(f"${float(valor):,.2f}" if valor else "$0.00")

            item_categoria.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            item_tipo.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            item_valor.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)

            item_categoria.setFlags(item_categoria.flags() & ~Qt.ItemIsEditable)
            item_tipo.setFlags(item_tipo.flags() & ~Qt.ItemIsEditable)
            if tipo and "Total" not in tipo and valor == "0.00":
                item_valor.setFlags(item_valor.flags() | Qt.ItemIsEditable)
                item_valor.setBackground(QColor("#d4edda"))
            else:
                item_valor.setFlags(item_valor.flags() & ~Qt.ItemIsEditable)
                item_valor.setFlags(item_valor.flags() & ~Qt.ItemIsEnabled)
                item_valor.setBackground(QColor("#f5f5f5"))

            if categoria in ["ACTIVOS", "PASIVOS", "PATRIMONIO", "TOTALES"]:
                item_categoria.setBackground(QColor("#e9ecef"))
                item_categoria.setFont(QFont("Arial", 12, QFont.Bold))

            self.table.setItem(row, 0, item_categoria)
            self.table.setItem(row, 1, item_tipo)
            self.table.setItem(row, 2, item_valor)

        for row in range(self.table.rowCount()):
            if not self.table.item(row, 1).text():
                self.table.setRowHeight(row, 30)
            else:
                self.table.setRowHeight(row, 35)

    def on_item_changed(self, item):
        if item.column() == 2:
            try:
                row = item.row()
                text = item.text().strip().replace('$', '').replace(',', '')
                
                # Store values before any modifications
                categoria = self.table.item(row, 0).text().lower() if self.table.item(row, 0) else ""
                tipo = self.table.item(row, 1).text().lower() if self.table.item(row, 1) else ""
                
                if text == "":
                    self.table.blockSignals(True)
                    item.setText("$0.00")
                    item.setBackground(QColor("#ffffff"))
                    item.setToolTip("Valor vacío, establecido a $0.00")
                    self.table.blockSignals(False)
                    self.update_manual_data(row, categoria, tipo, 0.0)
                    self.update_table_and_totals()
                    return
                    
                try:
                    valor = float(text)
                    if valor < 0:
                        raise ValueError("Los valores no pueden ser negativos.")
                        
                    self.table.blockSignals(True)
                    item.setText(f"${valor:,.2f}")
                    item.setBackground(QColor("#d4edda"))
                    item.setToolTip("Valor válido")
                    self.table.blockSignals(False)
                    
                    self.update_manual_data(row, categoria, tipo, valor)
                    self.update_table_and_totals()
                except ValueError:
                    self.table.blockSignals(True)
                    item.setBackground(QColor("#f8d7da"))
                    item.setToolTip("Error: Ingrese un valor numérico válido y no negativo")
                    QMessageBox.warning(self, "Error", "Por favor, ingrese un valor numérico válido y no negativo.")
                    item.setText("$0.00")
                    item.setBackground(QColor("#ffffff"))
                    item.setToolTip("Valor corregido a $0.00")
                    self.table.blockSignals(False)
                    
                    self.update_manual_data(row, categoria, tipo, 0.0)
                    self.update_table_and_totals()
            except Exception as e:
                print(f"Error in on_item_changed: {str(e)}")
                self.table.blockSignals(False)

    def update_manual_data(self, row, categoria, tipo, valor):
        mask = (self.manual_data["categoria"] == categoria) & (self.manual_data["tipo"] == tipo)
        if mask.any():
            self.manual_data.loc[mask, "valor"] = valor
        else:
            new_row = pd.DataFrame({
                "categoria": [categoria],
                "tipo": [tipo],
                "valor": [valor]
            }, dtype=object)
            new_row["valor"] = new_row["valor"].astype("float64")
            if self.manual_data.empty:
                self.manual_data = new_row
            else:
                self.manual_data = pd.concat([self.manual_data, new_row], ignore_index=True)

    def update_table_and_totals(self):
        data = []
        current_category = ""
        for row in range(self.table.rowCount()):
            categoria_item = self.table.item(row, 0)
            tipo_item = self.table.item(row, 1)
            valor_item = self.table.item(row, 2)

            if not categoria_item or not tipo_item or not valor_item:
                continue

            categoria = categoria_item.text().strip()
            tipo = tipo_item.text().strip()
            valor_text = valor_item.text().strip().replace('$', '').replace(',', '')

            if categoria and categoria not in ["ACTIVOS", "PASIVOS", "PATRIMONIO", "TOTALES"]:
                current_category = categoria.lower()

            if tipo and "Total" not in tipo and valor_item.flags() & Qt.ItemIsEditable:
                try:
                    valor = float(valor_text)
                except ValueError:
                    valor = 0.0
                data.append({
                    "categoria": current_category,
                    "tipo": tipo.lower(),
                    "valor": valor
                })

        self.manual_data = pd.DataFrame(data) if data else pd.DataFrame(columns=["categoria", "tipo", "valor"], dtype=object)
        self.manual_data["valor"] = self.manual_data["valor"].astype("float64")

        try:
            df_combined = self.manual_data.copy()
            if df_combined.empty:
                df_combined = pd.DataFrame(columns=["categoria", "tipo", "valor"], dtype=object)
                df_combined["valor"] = df_combined["valor"].astype("float64")

            self.balance_final = calcular_balance(df_combined)
            ratios = calcular_ratios(self.balance_final)

            self.table.blockSignals(True)
            total_rows = {
                "Total Activos": 0.0,
                "Total Pasivos": 0.0,
                "Total Patrimonio": 0.0,
                "Total Pasivos + Patrimonio": 0.0
            }
            for _, row in self.balance_final[self.balance_final["categoria"] == "TOTALES"].iterrows():
                if row["tipo"] == "Total Activos":
                    total_rows["Total Activos"] = row["valor"]
                elif row["tipo"] == "Total Pasivos":
                    total_rows["Total Pasivos"] = row["valor"]

            # Calcular Patrimonio como Total Activos - Total Pasivos
            total_rows["Total Patrimonio"] = total_rows["Total Activos"] - total_rows["Total Pasivos"]
            # Verificar que la ecuación contable se cumpla
            total_rows["Total Pasivos + Patrimonio"] = total_rows["Total Pasivos"] + total_rows["Total Patrimonio"]

            category_totals = {
                "activos corrientes": 0.0,
                "activos no corrientes": 0.0,
                "pasivos corrientes": 0.0,
                "pasivos no corrientes": 0.0,
                "patrimonio": 0.0
            }
            for _, row in self.balance_final.iterrows():
                categoria = row["categoria"].lower()
                if categoria in category_totals:
                    category_totals[categoria] += row["valor"]

            for row in range(self.table.rowCount()):
                table_categoria = self.table.item(row, 0)
                table_tipo = self.table.item(row, 1)
                if table_categoria and table_tipo and not table_tipo.text():
                    if table_categoria.text().lower() in category_totals:
                        item_valor = QTableWidgetItem(f"${category_totals[table_categoria.text().lower()]:,.2f}")
                        item_valor.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                        item_valor.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                        item_valor.setBackground(QColor("#f5f5f5"))
                        self.table.setItem(row, 2, item_valor)
                elif table_categoria and table_tipo and table_tipo.text() in total_rows:
                    valor = total_rows[table_tipo.text()]
                    item = self.table.item(row, 2)
                    if item:
                        item.setText(f"${valor:,.2f}")
                        item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                        item.setBackground(QColor("#f5f5f5"))

            self.table.blockSignals(False)

            diagnostico = generar_diagnostico(ratios, total_rows)
            self.diagnostico_label.setText(diagnostico)
            self.diagnostico_label.setTextFormat(Qt.RichText)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al actualizar totales: {str(e)}")

    def generar_balance(self):
        try:
            df_combined = self.manual_data.copy()
            if df_combined.empty:
                df_combined = pd.DataFrame(columns=["categoria", "tipo", "valor"], dtype=object)
                df_combined["valor"] = df_combined["valor"].astype("float64")

            self.balance_final = calcular_balance(df_combined)
            ratios = calcular_ratios(self.balance_final)
            total_rows = {
                "Total Activos": 0.0,
                "Total Pasivos": 0.0,
                "Total Patrimonio": 0.0,
                "Total Pasivos + Patrimonio": 0.0
            }
            for _, row in self.balance_final[self.balance_final["categoria"] == "TOTALES"].iterrows():
                if row["tipo"] == "Total Activos":
                    total_rows["Total Activos"] = row["valor"]
                elif row["tipo"] == "Total Pasivos":
                    total_rows["Total Pasivos"] = row["valor"]

            # Calcular Patrimonio como Total Activos - Total Pasivos
            total_rows["Total Patrimonio"] = total_rows["Total Activos"] - total_rows["Total Pasivos"]
            # Verificar que la ecuación contable se cumpla
            total_rows["Total Pasivos + Patrimonio"] = total_rows["Total Pasivos"] + total_rows["Total Patrimonio"]

            self.table.blockSignals(True)
            category_totals = {
                "activos corrientes": 0.0,
                "activos no corrientes": 0.0,
                "pasivos corrientes": 0.0,
                "pasivos no corrientes": 0.0,
                "patrimonio": 0.0
            }
            for _, row in self.balance_final.iterrows():
                categoria = row["categoria"].lower()
                tipo = row["tipo"].lower()
                valor = row["valor"]
                for table_row in range(self.table.rowCount()):
                    table_categoria = self.table.item(table_row, 0)
                    table_tipo = self.table.item(table_row, 1)
                    if table_categoria and table_tipo and table_categoria.text().lower() == categoria and table_tipo.text().lower() == tipo:
                        item_valor = QTableWidgetItem(f"${valor:,.2f}")
                        item_valor.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                        if "total" not in tipo:
                            item_valor.setFlags(Qt.ItemIsEditable | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                            item_valor.setBackground(QColor("#d4edda"))
                        else:
                            item_valor.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                            item_valor.setBackground(QColor("#f5f5f5"))
                        self.table.setItem(table_row, 2, item_valor)
                        if table_categoria.text().lower() in category_totals:
                            category_totals[table_categoria.text().lower()] += valor
                        break
            for table_row in range(self.table.rowCount()):
                table_categoria = self.table.item(table_row, 0)
                table_tipo = self.table.item(table_row, 1)
                if table_categoria and table_tipo and not table_tipo.text():
                    if table_categoria.text().lower() in category_totals:
                        item_valor = QTableWidgetItem(f"${category_totals[table_categoria.text().lower()]:,.2f}")
                        item_valor.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                        item_valor.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                        item_valor.setBackground(QColor("#f5f5f5"))
                        self.table.setItem(table_row, 2, item_valor)

            self.table.blockSignals(False)

            diagnostico = generar_diagnostico(ratios, total_rows)
            self.diagnostico_label.setText(diagnostico)
            self.diagnostico_label.setTextFormat(Qt.RichText)
            QMessageBox.information(self, "Éxito", "Balance generado correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar balance: {str(e)}")

    def guardar_balance(self):
        nombre_empresa = self.input_empresa.text().strip()
        fecha_balance = self.input_fecha.text().strip()

        if not nombre_empresa or not fecha_balance:
            QMessageBox.warning(self, "Campo requerido", "Por favor ingresa el nombre de la empresa y la fecha del balance.")
            return

        try:
            df_combined = self.manual_data.copy()
            if df_combined.empty:
                df_combined = pd.DataFrame(columns=["categoria", "tipo", "valor"], dtype=object)
                df_combined["valor"] = df_combined["valor"].astype("float64")

            self.balance_final = calcular_balance(df_combined)
            ratios = calcular_ratios(self.balance_final)
            diagnostico = generar_diagnostico(ratios)

            archivo, _ = QFileDialog.getSaveFileName(self, "Guardar Balance", "", "Excel (*.xlsx)")
            if archivo:
                exportar_balance_profesional(
                    nombre_empresa,
                    fecha_balance,
                    self.balance_final,
                    archivo,
                    ratios,
                    diagnostico
                )
                QMessageBox.information(self, "Éxito", f"Balance guardado en:\n{archivo}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = CoreXUI()
    ventana.show()
    sys.exit(app.exec_())