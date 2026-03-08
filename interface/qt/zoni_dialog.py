# -*- coding: utf-8 -*-
from qgis.PyQt.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QLabel,
    QPushButton,
    QGroupBox,
    QGridLayout,
    QMessageBox,
    QFrame,
)

from qgis.PyQt.QtCore import pyqtSignal
from qgis.gui import QgsMapLayerComboBox
from qgis.core import QgsMapLayerProxyModel

from ...interface.qt.filtro_eventos import EnterKeyFilter
from ...infraestrutura.espacial.config_camadas import detectar_camada_inteligente


class ZoniDialog(QDialog):

    sinal_iniciar_selecao = pyqtSignal()
    sinal_executar_analise = pyqtSignal()
    sinal_dialogo_fechado = pyqtSignal()

    def __init__(self, iface):
        super().__init__(iface.mainWindow())
        self.iface = iface
        self._montar_ui()

    # -------------------------------------------------
    # UI
    # -------------------------------------------------

    def _montar_ui(self):
        self.setWindowTitle("Zôni v2 – Seleção de Camadas")
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        hero = QLabel(
            "<b>Zôni v2</b><br>"
            "Selecione as camadas-chave para análise urbanística, ambiental e de risco. "
            "Os campos são detectados automaticamente, mas você pode substituir a qualquer momento."
        )
        hero.setWordWrap(True)
        hero.setStyleSheet("font-size:11pt;color:#1f2430; margin-bottom:6px;")
        layout.addWidget(hero)

        card = QFrame()
        card.setObjectName("cardMain")
        card.setStyleSheet(
            "QFrame#cardMain{border:1px solid #d0d7e2;border-radius:8px;"
            "background: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 #f7f9fc, stop:1 #ffffff);}"
        )
        grid = QGridLayout(card)
        grid.setHorizontalSpacing(14)
        grid.setVerticalSpacing(10)

        # LOTES
        lbl_lotes = QLabel("Lotes (polígonos)")
        self.combo_lotes = QgsMapLayerComboBox()
        self.combo_lotes.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        self.combo_lotes.setToolTip("Camada poligonal de lotes/parcelas usada para seleção e análise.")
        grid.addWidget(lbl_lotes, 0, 0)
        grid.addWidget(self.combo_lotes, 1, 0)

        # ZONEAMENTO
        lbl_zon = QLabel("Zoneamento (polígonos)")
        self.combo_zoneamento = QgsMapLayerComboBox()
        self.combo_zoneamento.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        self.combo_zoneamento.setToolTip("Camada poligonal com códigos de zoneamento da LC 275/2025.")
        grid.addWidget(lbl_zon, 0, 1)
        grid.addWidget(self.combo_zoneamento, 1, 1)

        # LOGRADOUROS
        lbl_log = QLabel("Logradouros (linhas)")
        self.combo_logradouros = QgsMapLayerComboBox()
        self.combo_logradouros.setFilters(QgsMapLayerProxyModel.LineLayer)
        self.combo_logradouros.setToolTip("Camada linear de vias para cálculo de testadas e notas especiais.")
        grid.addWidget(lbl_log, 2, 0)
        grid.addWidget(self.combo_logradouros, 3, 0)

        # Status
        status = QLabel("Detectado automaticamente: ajuste manual se necessário.")
        status.setWordWrap(True)
        status.setStyleSheet("color:#4d5870;font-size:9pt;")
        grid.addWidget(status, 3, 1)

        layout.addWidget(card)

        # APP
        app_group = QGroupBox("Camadas Ambientais (APP)")
        app_layout = QGridLayout()
        app_group.setLayout(app_layout)

        self.combo_app_nuic = QgsMapLayerComboBox()
        self.combo_app_nuic.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        app_layout.addWidget(QLabel("Faixa APP - NUIC:"), 0, 0)
        app_layout.addWidget(self.combo_app_nuic, 0, 1)
        self.combo_app_nuic.setToolTip("Faixa marginal de curso d'água (NUIC).")

        self.combo_app_manguezal = QgsMapLayerComboBox()
        self.combo_app_manguezal.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        app_layout.addWidget(QLabel("APP - Manguezais:"), 1, 0)
        app_layout.addWidget(self.combo_app_manguezal, 1, 1)
        self.combo_app_manguezal.setToolTip("Camada de APP em áreas de manguezal.")

        self.combo_app_inclinacao = QgsMapLayerComboBox()
        self.combo_app_inclinacao.setFilters(QgsMapLayerProxyModel.RasterLayer)
        app_layout.addWidget(QLabel("APP - Inclinação:"), 2, 0)
        app_layout.addWidget(self.combo_app_inclinacao, 2, 1)
        self.combo_app_inclinacao.setToolTip("Raster de declividade para APP por inclinação (>45°).")

        layout.addWidget(app_group)

        # RISCO
        risco_group = QGroupBox("Camadas de Risco")
        risco_layout = QGridLayout()
        risco_group.setLayout(risco_layout)

        self.combo_risco_geo = QgsMapLayerComboBox()
        self.combo_risco_geo.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        risco_layout.addWidget(QLabel("Movimento de Massa:"), 0, 0)
        risco_layout.addWidget(self.combo_risco_geo, 0, 1)
        self.combo_risco_geo.setToolTip("Mapa de suscetibilidade a movimento de massa.")

        self.combo_risco_inun = QgsMapLayerComboBox()
        self.combo_risco_inun.setFilters(QgsMapLayerProxyModel.PolygonLayer)
        risco_layout.addWidget(QLabel("Inundação:"), 1, 0)
        risco_layout.addWidget(self.combo_risco_inun, 1, 1)
        self.combo_risco_inun.setToolTip("Mapa de suscetibilidade a inundação.")

        layout.addWidget(risco_group)

        # BOTÕES
        self.botao_selecionar = QPushButton("Selecionar lote(s)")
        self.botao_selecionar.clicked.connect(self.sinal_iniciar_selecao.emit)
        layout.addWidget(self.botao_selecionar)
        self.botao_selecionar.setToolTip("Ativa ferramenta de seleção retangular no mapa para escolher lotes.")

        self.botao_analisar = QPushButton("Analisar")
        self.botao_analisar.clicked.connect(self.sinal_executar_analise.emit)
        layout.addWidget(self.botao_analisar)
        self.botao_analisar.setToolTip("Executa a análise urbanística/ambiental sobre os lotes selecionados.")

        status = QLabel("Dica: os combos tentam detectar automaticamente as camadas mais prováveis; ajuste manual se necessário.")
        status.setWordWrap(True)
        status.setStyleSheet("color:#444; font-size:9pt;")
        status.setFrameShape(QFrame.NoFrame)
        layout.addWidget(status)

        self.setStyleSheet("""
            QDialog { background-color: #e9ecf2; }
            QLabel { color: #1f2430; }
            QGroupBox {
                border: 1px solid #d0d7e2;
                border-radius: 6px;
                margin-top: 10px;
                padding: 10px;
                font-weight: 600;
                color: #2c3444;
            }
            QGroupBox::title { subcontrol-origin: margin; left: 8px; padding: 0 4px; }
            QPushButton {
                background-color: #4472C4;
                color: #fff;
                border: none;
                border-radius: 4px;
                padding: 8px 12px;
                font-weight: 600;
            }
            QPushButton:hover { background-color: #365a97; }
        """)

        self.resize(520, 720)

        # ----------------------------
        # SELEÇÃO AUTOMÁTICA INTELIGENTE
        # ----------------------------

    def aplicar_selecao_automatica(self):
        from ...infraestrutura.espacial.config_camadas import detectar_camada_inteligente

        mapa_combos = {
            "lotes": self.combo_lotes,
            "zoneamento": self.combo_zoneamento,
            "logradouros": self.combo_logradouros,
            "faixa_app_nuic": self.combo_app_nuic,
            "app_manguezal": self.combo_app_manguezal,
            "app_inclinacao": self.combo_app_inclinacao,
            "susc_mov_massa": self.combo_risco_geo,
            "susc_inundacao": self.combo_risco_inun,
        }

        for chave, combo in mapa_combos.items():
            camada_auto = detectar_camada_inteligente(chave)
            if camada_auto:
                combo.setLayer(camada_auto)

    # -------------------------------------------------
    # MÉTODOS EXPOSTOS AO PRESENTER
    # -------------------------------------------------

    def set_layer(self, chave, layer):
        mapa = {
            "lotes": self.combo_lotes,
            "zoneamento": self.combo_zoneamento,
            "logradouros": self.combo_logradouros,
            "faixa_app_nuic": self.combo_app_nuic,
            "app_manguezal": self.combo_app_manguezal,
            "app_inclinacao": self.combo_app_inclinacao,
            "susc_mov_massa": self.combo_risco_geo,
            "susc_inundacao": self.combo_risco_inun,
        }
        combo = mapa.get(chave)
        if combo and layer:
            combo.setLayer(layer)

    def get_camada(self, chave):
        mapa = {
            "lotes": self.combo_lotes,
            "zoneamento": self.combo_zoneamento,
            "logradouros": self.combo_logradouros,
            "faixa_app_nuic": self.combo_app_nuic,
            "app_manguezal": self.combo_app_manguezal,
            "app_inclinacao": self.combo_app_inclinacao,
            "susc_mov_massa": self.combo_risco_geo,
            "susc_inundacao": self.combo_risco_inun,
        }
        combo = mapa.get(chave)
        return combo.currentLayer() if combo else None

    def mostrar_erro(self, msg):
        QMessageBox.critical(self, "Erro", msg)

    def perguntar_sim_nao(self, titulo, msg):
        return QMessageBox.question(self, titulo, msg) == QMessageBox.Yes

    # -------------------------------------------------
    # CONTRATO COM O PRESENTER
    # -------------------------------------------------

    def on_layer_changed(self, callback):
        self.combo_lotes.layerChanged.connect(callback)

    def habilitar_botao_analisar(self, ativo: bool):
        self.botao_analisar.setEnabled(ativo)

    def ocultar(self):
        self.hide()

    def mostrar(self):
        self.show()

    def trazer_para_frente(self):
        self.raise_()
        self.activateWindow()

    def confirmar_nota10(self, nome_via):
        return QMessageBox.question(
            self,
            "Nota 10",
            f"Aplicar regra de acesso único para {nome_via}?"
        ) == QMessageBox.Yes

    def closeEvent(self, event):
        self.sinal_dialogo_fechado.emit()
        super().closeEvent(event)
