import sys
import pandas as pd
import subprocess
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QLineEdit, QComboBox, QListWidget, QListWidgetItem, QProgressBar, QTextEdit, QCheckBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal


class Worker(QThread):
    status = pyqtSignal(str)

    def __init__(self, file_path, device_col, part_col, appr_col, columns, is_one_way, lsl, usl):
        super().__init__()
        self.file_path = file_path
        self.device_col = device_col
        self.part_col = part_col
        self.appr_col = appr_col
        self.columns = columns
        self.is_one_way = is_one_way
        self.lsl = lsl
        self.usl = usl

    def run(self):
        self.status.emit("Gerando relatório, por favor aguarde...")
        script_path = os.path.join(
            os.path.dirname(__file__), 'gage_rr_analysis.R')
        try:
            result = subprocess.run(
                ["Rscript", script_path, self.file_path, self.device_col,
                 self.part_col, self.appr_col, str(self.is_one_way),
                 self.lsl, self.usl] + self.columns,
                capture_output=True, text=True
            )
            if result.returncode == 0:
                self.status.emit("Relatório gerado com sucesso!")
            else:
                self.status.emit(
                    f"Erro ao gerar o relatório:\n{result.stderr}")
        except Exception as e:
            self.status.emit(f"Erro ao executar o script R: {e}")


class App(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Gerador de Relatórios Gage R&R")
        self.setGeometry(100, 100, 1000, 600)

        self.initUI()

    def initUI(self):
        mainLayout = QHBoxLayout()

        # Lateral esquerda (ocupada totalmente pelo QListWidget)
        leftLayout = QVBoxLayout()

        self.columnsLabel = QLabel("Selecione as colunas para análise:")
        leftLayout.addWidget(self.columnsLabel)

        self.searchLabel = QLabel("Filtrar colunas:")
        leftLayout.addWidget(self.searchLabel)

        self.searchBox = QLineEdit()
        self.searchBox.textChanged.connect(self.filterColumns)
        leftLayout.addWidget(self.searchBox)

        self.columnsListWidget = QListWidget()
        self.columnsListWidget.setSelectionMode(QListWidget.MultiSelection)
        leftLayout.addWidget(self.columnsListWidget)

        mainLayout.addLayout(leftLayout)

        # Lateral direita (todos os outros componentes)
        rightLayout = QVBoxLayout()

        self.fileLabel = QLabel("Selecione o arquivo de entrada:")
        rightLayout.addWidget(self.fileLabel)

        self.filePath = QLineEdit()
        rightLayout.addWidget(self.filePath)

        self.browseButton = QPushButton("Procurar")
        self.browseButton.clicked.connect(self.browseFile)
        rightLayout.addWidget(self.browseButton)

        self.deviceLabel = QLabel("Selecione a coluna do dispositivo:")
        rightLayout.addWidget(self.deviceLabel)

        self.deviceComboBox = QComboBox()
        rightLayout.addWidget(self.deviceComboBox)

        self.partLabel = QLabel("Selecione a coluna do part (peça):")
        rightLayout.addWidget(self.partLabel)

        self.partComboBox = QComboBox()
        rightLayout.addWidget(self.partComboBox)

        self.apprLabel = QLabel("Selecione a coluna do appr (operador):")
        rightLayout.addWidget(self.apprLabel)

        self.apprComboBox = QComboBox()
        rightLayout.addWidget(self.apprComboBox)

        self.lslLabel = QLabel("Valor LSL (Limite Inferior Especificado):")
        rightLayout.addWidget(self.lslLabel)

        self.lslInput = QLineEdit()
        rightLayout.addWidget(self.lslInput)

        self.uslLabel = QLabel("Valor USL (Limite Superior Especificado):")
        rightLayout.addWidget(self.uslLabel)

        self.uslInput = QLineEdit()
        rightLayout.addWidget(self.uslInput)

        self.oneWayCheckBox = QCheckBox(
            "Modo One-Way (Sem variedade de operadores)")
        rightLayout.addWidget(self.oneWayCheckBox)

        self.generateButton = QPushButton("Gerar Relatório")
        self.generateButton.clicked.connect(self.startReportGeneration)
        rightLayout.addWidget(self.generateButton)

        self.statusLabel = QLabel("Status:")
        rightLayout.addWidget(self.statusLabel)

        self.statusTextEdit = QTextEdit()
        self.statusTextEdit.setReadOnly(True)
        rightLayout.addWidget(self.statusTextEdit)

        self.progressBar = QProgressBar()
        self.progressBar.setRange(0, 0)
        self.progressBar.setTextVisible(False)
        rightLayout.addWidget(self.progressBar)
        self.progressBar.hide()

        mainLayout.addLayout(rightLayout)

        container = QWidget()
        container.setLayout(mainLayout)
        self.setCentralWidget(container)

        self.filePath.textChanged.connect(self.updateColumnSelectors)

    def browseFile(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo", "", "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)", options=options)
        if filePath:
            self.filePath.setText(filePath)
            self.updateColumnSelectors()

    def updateColumnSelectors(self):
        file_path = self.filePath.text()
        if file_path:
            try:
                df = pd.read_excel(file_path)
                columns = df.columns.tolist()

                self.deviceComboBox.clear()
                self.deviceComboBox.addItems(columns)

                self.partComboBox.clear()
                self.partComboBox.addItems(columns)

                self.apprComboBox.clear()
                self.apprComboBox.addItems(columns)

                self.columnsListWidget.clear()
                for col in columns:
                    self.columnsListWidget.addItem(QListWidgetItem(col))

            except Exception as e:
                self.statusTextEdit.setText(
                    f"Erro ao carregar o arquivo ou ler as colunas: {e}")

    def filterColumns(self):
        filter_text = self.searchBox.text().lower()
        for i in range(self.columnsListWidget.count()):
            item = self.columnsListWidget.item(i)
            item.setHidden(filter_text not in item.text().lower())

    def startReportGeneration(self):
        file_path = self.filePath.text()
        device_col = self.deviceComboBox.currentText()
        part_col = self.partComboBox.currentText()
        appr_col = self.apprComboBox.currentText()
        lsl = self.lslInput.text() if self.lslInput.text() else "NA"
        usl = self.uslInput.text() if self.uslInput.text() else "NA"

        selected_items = self.columnsListWidget.selectedItems()
        columns = [item.text() for item in selected_items]

        is_one_way = self.oneWayCheckBox.isChecked()

        self.worker = Worker(file_path, device_col, part_col,
                             appr_col, columns, is_one_way, lsl, usl)
        self.worker.status.connect(self.updateStatus)
        self.worker.start()
        self.progressBar.show()

    def updateStatus(self, message):
        self.statusTextEdit.append(message)
        self.progressBar.hide()


def main():
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
