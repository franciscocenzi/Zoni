# -*- coding: utf-8 -*-
import os
from datetime import datetime
from pathlib import Path

from qgis.PyQt.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QTextBrowser,
    QPushButton,
    QMessageBox,
    QFileDialog,
    QHBoxLayout,
)
from qgis.PyQt.QtCore import QTimer
from qgis.core import QgsProject

# IMPORTS CORRETOS
from ...dominio.motores.motor_analise_lote import CenarioEdificacao, analisar_lote
from ...infraestrutura.relatorios.construtor_relatorio import construir_contexto_relatorio
from ...infraestrutura.relatorios.renderizador_docx import RenderizadorDOCX
from ...infraestrutura.espacial.config_camadas import registrar_camada
from ...infraestrutura.espacial.validadores import lotes_sao_contiguos
from ...infraestrutura.espacial.geometrias import unir_geometrias
from ...infraestrutura.espacial.lote_utils import extrair_dados_cadastrais
from ...compartilhado.caminhos import obter_caminho_parametros


class ZoniPresenter:
    def __init__(self, view, iface):
        self.view = view
        self.iface = iface

        # Estado
        self.lotes_selecionados = []
        self.conexao_selecao = None
        self.timer_selecao = None
        self.camada_lotes_atual = None
        self._enter_filter = None
        self._event_filter_target = None
        base_plugin = Path(__file__).resolve().parents[2]
        self.caminho_parametros = obter_caminho_parametros(str(base_plugin))

        # Conectar sinais da View
        self.view.sinal_iniciar_selecao.connect(self.iniciar_selecao_lotes)
        self.view.sinal_executar_analise.connect(self.executar_analise_zoni_v2)
        self.view.on_layer_changed(self._on_camada_lotes_changed)

        self.view.aplicar_selecao_automatica()

        # Configurar monitoramento inicial
        self._configurar_monitor_selecao()

    # ------------------------------------------------------------------ #
    # MONITORAMENTO DE SELEÇÃO                                           #
    # ------------------------------------------------------------------ #
    def _configurar_monitor_selecao(self):
        self._desconectar_monitor_selecao()

        camada_lotes = self.view.get_camada("lotes")
        if camada_lotes:
            registrar_camada("lotes", camada_lotes)
            self.camada_lotes_atual = camada_lotes
            camada_lotes = self.camada_lotes_atual
        else:
            camada_lotes = None

        if camada_lotes:
            self.conexao_selecao = camada_lotes.selectionChanged.connect(
                self._atualizar_selecao_lotes
            )
            self._atualizar_selecao_lotes()

    def _desconectar_monitor_selecao(self):
        if self.conexao_selecao:
            try:
                self.conexao_selecao.disconnect()
            except Exception:
                pass
            self.conexao_selecao = None

    def _atualizar_selecao_lotes(self):
        if self.timer_selecao:
            self.timer_selecao.stop()

        self.timer_selecao = QTimer()
        self.timer_selecao.setSingleShot(True)
        self.timer_selecao.timeout.connect(self._processar_atualizacao_selecao)
        self.timer_selecao.start(100)

    def _processar_atualizacao_selecao(self):
        camada_lotes = self.view.get_camada("lotes")

        if not camada_lotes:
            self.timer_selecao = None
            return

        if camada_lotes != self.camada_lotes_atual:
            self.camada_lotes_atual = camada_lotes
            self._configurar_monitor_selecao()
            return

        selecionadas = list(camada_lotes.getSelectedFeatures())

        if selecionadas:
            self.lotes_selecionados = selecionadas
            self.view.habilitar_botao_analisar(True)
            mensagem = f"{len(selecionadas)} lote(s) selecionado(s) na camada '{camada_lotes.name()}'"
            self.iface.messageBar().pushInfo("Zôni v2", mensagem)
        else:
            self.lotes_selecionados = []
            self.view.habilitar_botao_analisar(False)

        self.timer_selecao = None

    def _on_camada_lotes_changed(self):
        camada = self.view.get_camada("lotes")
        if camada:
            registrar_camada("lotes", camada)
        else:
            self.iface.messageBar().pushInfo(
                "Zôni v2",
                "Nenhuma camada de lotes detectada automaticamente; selecione manualmente."
            )
        self._configurar_monitor_selecao()

    # ------------------------------------------------------------------ #
    # SELEÇÃO DE LOTES                                                   #
    # ------------------------------------------------------------------ #
    def iniciar_selecao_lotes(self):
        camada_lotes = self.view.get_camada("lotes")

        if camada_lotes is None:
            self.iface.messageBar().pushWarning(
                "Zôni v2",
                "Nenhuma camada de lotes selecionada no menu. Escolha uma camada.",
            )
            return

        node = QgsProject.instance().layerTreeRoot().findLayer(camada_lotes.id())
        if node is not None:
            node.setItemVisibilityChecked(True)

        self.iface.layerTreeView().setCurrentLayer(camada_lotes)
        self.iface.setActiveLayer(camada_lotes)

        self.view.ocultar()
        self.iface.actionSelectRectangle().trigger()

        # Import corrigido com caminho absoluto
        from ...interface.qt.filtro_eventos import EnterKeyFilter

        self._enter_filter = EnterKeyFilter(self.finalizar_selecao_lotes)
        alvo = self.iface.mapCanvas()
        alvo.installEventFilter(self._enter_filter)
        self._event_filter_target = alvo

    def finalizar_selecao_lotes(self):
        if self._event_filter_target and self._enter_filter:
            self._event_filter_target.removeEventFilter(self._enter_filter)

        from qgis.PyQt.QtWidgets import QApplication

        QApplication.processEvents()
        self.iface.mapCanvas().refresh()
        QApplication.processEvents()

        camada_lotes = self.view.get_camada("lotes")

        if camada_lotes is None:
            self.iface.messageBar().pushWarning(
                "Zôni v2",
                "Camada de lotes não encontrada ao finalizar seleção.",
            )
            self.view.mostrar()
            return

        selecionados = list(camada_lotes.getSelectedFeatures())

        if not selecionados:
            self.iface.messageBar().pushWarning(
                "Zôni v2",
                "Nenhum lote selecionado.",
            )
            self.view.mostrar()
            return

        self.lotes_selecionados = selecionados
        self.view.habilitar_botao_analisar(True)

        self.view.mostrar()
        self.view.trazer_para_frente()

        self._configurar_monitor_selecao()

#    # ------------------------------------------------------------------ #
#    # HELPERS DE CAMADA                                                  #
#    # ------------------------------------------------------------------ #
#    def _layer(self, combo, chave):
#        if combo is not None:
#            lyr = combo.currentLayer()
#            if lyr:
#                return lyr
#        return MAPA_CAMADAS.get(chave)

    def _obter_camada_lotes_atual(self):
        return self.view.get_camada("lotes")

    # ------------------------------------------------------------------ #
    # HELPERS DE ANÁLISE                                                 #
    # ------------------------------------------------------------------ #
    def _obter_area_lote(self, feature):
        nomes = feature.fields().names()
        if "área" in nomes:
            return feature["área"]
        if "area" in nomes:
            return feature["area"]
        return feature.geometry().area() if feature.geometry() else 0

    def _registrar_camadas(self):
        registrar_camada("lotes", self.view.get_camada("lotes"))
        registrar_camada("zoneamento", self.view.get_camada("zoneamento"))
        registrar_camada("logradouros", self.view.get_camada("logradouros"))
        registrar_camada("faixa_app_nuic", self.view.get_camada("faixa_app_nuic"))
        registrar_camada("app_manguezal", self.view.get_camada("app_manguezal"))
        registrar_camada("app_inclinacao", self.view.get_camada("app_inclinacao"))
        registrar_camada("susc_inundacao", self.view.get_camada("susc_inundacao"))
        registrar_camada("susc_mov_massa", self.view.get_camada("susc_mov_massa"))

    def _rodar_analise(self, geom_lote, cenario, nota10=False, nota37=False):
        return analisar_lote(
            geom_lote=geom_lote,
            cenario=cenario,
            caminho_parametros_zoneamento=self.caminho_parametros,
            nota10_ativada=nota10,
            nota37_ativada=nota37,
        )

    def _resolver_notas(self, analise, geom_lote, cenario):
        nota37_detectada = bool(getattr(analise, "detectou_frente_nota_37", False))
        nota10_detectada = bool(getattr(analise, "detectou_frente_nota_10", False))
        nome_via = getattr(analise, "nome_via_nota_10", "logradouro detectado")

        # Aplica Nota 37 automaticamente
        if nota37_detectada and not getattr(getattr(analise, "zona_resolvida", None), "notas_ativas", []):
            analise = self._rodar_analise(
                geom_lote=geom_lote,
                cenario=cenario,
                nota10=False,
                nota37=True,
            )
            analise.detectou_frente_nota_37 = True

        # Nota 10 depende de confirmação
        if nota10_detectada:
            aplicar = self.view.confirmar_nota10(nome_via)
            if aplicar:
                analise = self._rodar_analise(
                    geom_lote=geom_lote,
                    cenario=cenario,
                    nota10=True,
                    nota37=nota37_detectada,
                )
                analise.detectou_frente_nota_10 = True
                analise.detectou_frente_nota_37 = nota37_detectada
                analise.nome_via_nota_10 = nome_via

        return analise

    def _gerar_relatorio(self, analise, dados_lotes, titulo):
        contexto = construir_contexto_relatorio(dados_lotes, analise)
        
        from qgis.PyQt.QtWidgets import QFileDialog, QMessageBox
        import os
        
        data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_sugerido = f"Zoni_v2_{data_hora}.docx"
        file_path, _ = QFileDialog.getSaveFileName(
            self.iface.mainWindow(),
            "Salvar Relatório DOCX Oficial",
            nome_sugerido,
            "Documentos do Word (*.docx);;Todos os arquivos (*)",
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".docx"):
            file_path += ".docx"
            
        renderizador = RenderizadorDOCX()
        sucesso = renderizador.renderizar_e_salvar(contexto, file_path)
        
        if sucesso:
            res = QMessageBox.question(
                self.iface.mainWindow(),
                "Sucesso",
                f"Relatório gerado em:\n{file_path}\n\nDeseja abrir o arquivo agora?",
                QMessageBox.Yes | QMessageBox.No
            )
            if res == QMessageBox.Yes:
                os.startfile(file_path)

    def _debug_app_faixa(self, camada_app, geom, etiqueta):
        if not (camada_app and geom):
            print(f"🔵 Camada de APP faixa não disponível para {etiqueta}")
            return

        from qgis.core import QgsFeatureRequest

        bbox = geom.boundingBox()
        request = QgsFeatureRequest().setFilterRect(bbox).setLimit(100)
        features = list(camada_app.getFeatures(request))
        print(f"🔵 Feições da APP faixa na área do {etiqueta}: {len(features)}")
        for feat in features:
            geom_app = feat.geometry()
            if geom_app.intersects(geom):
                print("  ✅ Interseção encontrada!")
                attrs = feat.attributes()
                fields = feat.fields().names()
                for i, f in enumerate(fields):
                    print(f"    {f}: {attrs[i]}")
            else:
                print("  ❌ Feição dentro do bbox mas não intersecta")
    # ------------------------------------------------------------------ #
    # EXECUÇÃO DE ANÁLISE                                                #
    # ------------------------------------------------------------------ #
    def executar_analise_zoni_v2(self):
        """Executa a análise completa do lote/gleba."""
        camada_app_faixa = self.view.get_camada("faixa_app_nuic")
        if not self.lotes_selecionados:
            camada_lotes = self._obter_camada_lotes_atual()
            if camada_lotes:
                self.lotes_selecionados = list(camada_lotes.getSelectedFeatures())

        camada_lotes = self._obter_camada_lotes_atual()
        if camada_lotes is None:
            self.iface.messageBar().pushWarning(
                "Zôni v2",
                "Camada de lotes não encontrada. Selecione uma camada no dropdown.",
            )
            return

        if not self.lotes_selecionados:
            self.iface.messageBar().pushWarning(
                "Zôni v2",
                "Nenhum lote foi selecionado. Selecione lotes na camada 'Lotes' ou use o botão 'Selecionar lote(s)'.",
            )
            return

        try:
            self._registrar_camadas()
        except Exception as e:
            self.iface.messageBar().pushWarning("Zôni v2", f"Erro ao registrar camadas: {e}")
            return

        if not os.path.exists(self.caminho_parametros):
            self.iface.messageBar().pushWarning("Zôni v2", f"Arquivo não encontrado:\n{self.caminho_parametros}")
            return

        # ============================================================
        # CASO 1 — GLEBA (múltiplos lotes contíguos)
        # ============================================================
        if len(self.lotes_selecionados) > 1:
            if not lotes_sao_contiguos(self.lotes_selecionados):
                QMessageBox.warning(
                    self.iface.mainWindow(),
                    "Zôni v2",
                    "Os lotes selecionados não são contíguos.\n\n"
                    "Selecione apenas lotes adjacentes para análise conjunta."
                )
                return

            geom_unificada = unir_geometrias(self.lotes_selecionados)
            if geom_unificada is None or geom_unificada.isEmpty():
                self.iface.messageBar().pushCritical(
                    "Zôni v2",
                    "Erro ao unir geometrias dos lotes selecionados."
                )
                return

            self._debug_app_faixa(camada_app_faixa, geom_unificada, "gleba")

            area_gleba = geom_unificada.area()
            cenario = CenarioEdificacao(area_lote_m2=area_gleba)

            analise = self._rodar_analise(
                geom_lote=geom_unificada,
                cenario=cenario,
            )
            analise.area_gleba_unificada = area_gleba
            analise = self._resolver_notas(analise, geom_unificada, cenario)
            analise.area_gleba_unificada = area_gleba

            lista_dados_lote = [extrair_dados_cadastrais(f) for f in self.lotes_selecionados]
            self._gerar_relatorio(analise, lista_dados_lote, "Relatório Zôni v2 – Gleba Unificada")
            return

        # ============================================================
        # CASO 2 — APENAS 1 LOTE
        # ============================================================
        feat_lote = self.lotes_selecionados[0]
        geom_lote = feat_lote.geometry()

        if geom_lote is None or geom_lote.isEmpty():
            self.iface.messageBar().pushWarning("Zôni v2", "Geometria do lote inválida.")
            return

        self._debug_app_faixa(camada_app_faixa, geom_lote, "lote")

        area_lote = self._obter_area_lote(feat_lote)

        cenario = CenarioEdificacao(
            area_lote_m2=area_lote,
            area_construida_total_m2=None,
            area_ocupada_projecao_m2=None,
            area_permeavel_m2=None,
            altura_maxima_m=None,
            numero_pavimentos=None,
        )

        analise = self._rodar_analise(
            geom_lote=geom_lote,
            cenario=cenario,
        )

        analise = self._resolver_notas(analise, geom_lote, cenario)

        lista_dados_lote = [extrair_dados_cadastrais(feat_lote)]
        self._gerar_relatorio(analise, lista_dados_lote, "Relatório Zôni v2 – Lote")

    # ------------------------------------------------------------------ #
    # EXIBIÇÃO DO RELATÓRIO                                              #
    # ------------------------------------------------------------------ #
    def _mostrar_relatorio_html(self, html: str, titulo: str):
        pass  # Depreciado pelo fluxo DOCX nativo

    def _salvar_como_pdf(self, html: str, titulo: str):
        """Salva o relatório HTML como arquivo PDF."""
        from PyQt5.QtPrintSupport import QPrinter
        from PyQt5.QtGui import QTextDocument
        from PyQt5.QtWidgets import QFileDialog, QMessageBox

        data_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_sugerido = f"Zoni_v2_{data_hora}.pdf"

        file_path, _ = QFileDialog.getSaveFileName(
            self.iface.mainWindow(),
            "Salvar Relatório como PDF",
            nome_sugerido,
            "Arquivos PDF (*.pdf);;Todos os arquivos (*)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".pdf"):
            file_path += ".pdf"

        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        printer.setPageSize(QPrinter.A4)
        printer.setOrientation(QPrinter.Portrait)

        document = QTextDocument()
        document.setHtml(html)
        document.print_(printer)

        QMessageBox.information(
            self.iface.mainWindow(),
            "PDF Salvo com Sucesso",
            f"Relatório salvo como:\n{file_path}",
        )

    def _imprimir_html(self, html: str):
        """Imprime o relatório HTML."""
        from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
        from PyQt5.QtGui import QTextDocument
        from PyQt5.QtWidgets import QMessageBox

        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self.iface.mainWindow())

        if dialog.exec_() == QPrintDialog.Accepted:
            document = QTextDocument()
            document.setHtml(html)
            document.print_(printer)
            QMessageBox.information(
                self.iface.mainWindow(),
                "Impressão",
                "Relatório enviado para a impressora.",
            )
