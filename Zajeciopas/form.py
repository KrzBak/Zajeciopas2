from qgis.PyQt.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QFileDialog,
    QComboBox,
    QHBoxLayout,
    QSpacerItem,
    QSizePolicy
)
from qgis.gui import QgsMapLayerComboBox
from qgis.core import QgsProject, QgsMapLayerProxyModel


class message(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Przycinanie kabli do działek")
        self.resize(1000, 700)

        layout = QVBoxLayout(self)

        # Działki
        layout.addWidget(QLabel("Wybierz działki:"))
        self.cmbDzialki = QgsMapLayerComboBox()
        self.cmbDzialki.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        self.cmbDzialki.setProject(QgsProject.instance())
        layout.addWidget(self.cmbDzialki)

        # Kable
        layout.addWidget(QLabel("Wybierz kable:"))
        self.cmbKable = QgsMapLayerComboBox()
        self.cmbKable.setFilters(QgsMapLayerProxyModel.LineLayer)
        self.cmbKable.setProject(QgsProject.instance())
        layout.addWidget(self.cmbKable)

        # Kanalizacja istniejąca
        layout.addWidget(QLabel("Wybierz kanalizację istniejącą (opcjonalnie):"))
        self.cmbKanalizacja = QgsMapLayerComboBox()
        self.cmbKanalizacja.setFilters(QgsMapLayerProxyModel.LineLayer)
        self.cmbKanalizacja.setProject(QgsProject.instance())
        layout.addWidget(self.cmbKanalizacja)

        # Kanalizacja projektowana
        layout.addWidget(QLabel("Wybierz kanalizację projektowaną (opcjonalnie):"))
        self.cmbKanalizacjaProj = QgsMapLayerComboBox()
        self.cmbKanalizacjaProj.setFilters(QgsMapLayerProxyModel.LineLayer)
        self.cmbKanalizacjaProj.setProject(QgsProject.instance())
        layout.addWidget(self.cmbKanalizacjaProj)

        # Przycisk Przytnij
        self.btnUruchom = QPushButton("Przytnij")
        layout.addWidget(self.btnUruchom)

        # Przełącznik tabel
        layout.addWidget(QLabel("Wybierz tabelę:"))
        self.cmbTabela = QComboBox()
        self.cmbTabela.addItems(["Tabela kabli", "Tabela kanalizacji projektowanej"])
        layout.addWidget(self.cmbTabela)

        # Tabela 1
        layout.addWidget(QLabel("Wyniki:"))
        self.tblWyniki = QTableWidget()
        layout.addWidget(self.tblWyniki)

        # Tabela 2
        self.tblKanalizacja = QTableWidget()
        self.tblKanalizacja.hide()
        layout.addWidget(self.tblKanalizacja)

        # === DOLNY PANEL Z PRZYCISKAMI ===

        dolny_layout = QHBoxLayout()

        # Spacer wypychający przyciski w prawo
        spacer = QSpacerItem(
            40,
            20,
            QSizePolicy.Expanding,
            QSizePolicy.Minimum
        )
        dolny_layout.addItem(spacer)

        # Eksport
        self.btnExcel = QPushButton("Eksport do Excel")
        dolny_layout.addWidget(self.btnExcel)

        # Agregacja
        self.btnAgreguj = QPushButton("Agreguj")
        dolny_layout.addWidget(self.btnAgreguj)

        layout.addLayout(dolny_layout)