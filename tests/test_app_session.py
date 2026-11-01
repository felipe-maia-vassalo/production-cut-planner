"""Streamlit session checks for both languages and Excel interchange."""
from io import BytesIO
from pathlib import Path
import unittest
from unittest.mock import patch

import pandas as pd
from pandas.testing import assert_frame_equal
from streamlit.testing.v1 import AppTest

from test_workflows import (
    cuts_frame, mapping_workbook, structure_workbook,
)
from calcular_cortes_streamlit import calcular_cortes_por_listas
from nome_V6_adaptado import buscar_produtos


APP = Path(__file__).resolve().parents[1] / "app.py"
PAGES = [
    "📖 Instruções de uso", "🔎 Buscar códigos dos produtos",
    "✂️ Planejamento de cortes", "📅 Programação das máquinas",
]
TITLES = {
    "en": ["🏭 Production Cut Planner — How to use", "🏭 Production Cut Planner", "✂️ Cut Planning", "📅 Machine Scheduling"],
    "pt": ["🏭 Corte de Produção — Como usar", "🏭 Corte de Produção", "✂️ Planejamento de Cortes", "📅 Programação das Máquinas"],
}


def cut_result():
    return calcular_cortes_por_listas(["101"], [50], arquivo_estrutura=structure_workbook())


def new_session(language="en", page=PAGES[0], with_results=False):
    app = AppTest.from_file(str(APP), default_timeout=10)
    app.session_state["language"] = language
    app.session_state["pagina"] = page
    if with_results:
        app.session_state["resultado_planejamento_cortes"] = cut_result()
        app.session_state["resultado_busca_produtos"] = buscar_produtos(
            ["CHAIR ALIAS", "MISSING ITEM"],
            arquivo_estrutura=structure_workbook(), arquivo_de_para=mapping_workbook(),
        )
    return app.run()


class AppSessionTests(unittest.TestCase):
    def assert_no_app_error(self, app):
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])
        self.assertEqual(len(app.error), 0, [item.value for item in app.error])

    def test_all_pages_render_in_both_languages_with_results(self):
        for language in ("en", "pt"):
            with self.subTest(language=language):
                app = new_session(language, with_results=True)
                for page, title in zip(PAGES, TITLES[language]):
                    app.radio(key="pagina").set_value(page).run()
                    self.assert_no_app_error(app)
                    self.assertEqual(app.title[0].value, title)

    def test_editor_changes_and_schedule_survive_language_switch(self):
        app = new_session("en", PAGES[3], with_results=True)
        changes = {
            "editor_capacidade": {2: {"PERDAS (min)": 15.0}},
            "editor_metas_gerais": {2: {"META/H": 640.0}},
            "editor_excecoes": {0: {"META/H": 150.0}},
        }
        for key, edited_rows in changes.items():
            app.session_state[key] = {"edited_rows": edited_rows, "added_rows": [], "deleted_rows": []}
        app.run()
        before = {key: app.session_state["tabela_" + key].copy() for key in changes}
        self.assertEqual(before["editor_capacidade"].iloc[2]["PERDAS (min)"], 15)
        self.assertEqual(before["editor_metas_gerais"].iloc[2]["META/H"], 640)
        self.assertEqual(before["editor_excecoes"].iloc[0]["META/H"], 150)
        app.button(key="botao_organizar_programacao").click().run()
        self.assert_no_app_error(app)
        scheduled = app.session_state["resultado_programacao_maquinas"]["programacao"].copy()
        self.assertFalse(scheduled.empty)
        for language in ("pt", "en"):
            app.selectbox(key="language").set_value(language).run()
            self.assert_no_app_error(app)
            self.assertEqual(app.radio(key="pagina").value, PAGES[3])
            for key, expected in before.items():
                assert_frame_equal(app.session_state["tabela_" + key], expected)
            assert_frame_equal(app.session_state["resultado_programacao_maquinas"]["programacao"], scheduled)

    def test_english_cut_download_can_be_uploaded_in_either_language(self):
        downloaded = {}

        def capture_download(*args, **kwargs):
            if kwargs.get("key") == "download_planejamento_cortes":
                downloaded["data"] = kwargs["data"].getvalue()
                downloaded["name"] = kwargs["file_name"]
            return False

        with patch("streamlit.download_button", side_effect=capture_download):
            app = new_session("en", PAGES[2], with_results=True)
        self.assert_no_app_error(app)
        self.assertTrue(downloaded["name"].endswith(".xlsx"))
        with pd.ExcelFile(BytesIO(downloaded["data"])) as workbook:
            self.assertEqual(workbook.sheet_names, ["Cuts"])
            self.assertIn("Total quantity", pd.read_excel(workbook).columns)
        for language in ("en", "pt"):
            with self.subTest(language=language):
                imported = AppTest.from_file(str(APP), default_timeout=10)
                imported.session_state["language"] = language
                imported.session_state["pagina"] = PAGES[3]
                imported.session_state["origem_cortes"] = "Enviar planejamento em Excel"
                imported.session_state["fonte_upload_planejamento"] = {
                    "nome": downloaded["name"], "dados": downloaded["data"],
                }
                imported.run()
                self.assert_no_app_error(imported)
                imported.button(key="botao_organizar_programacao").click().run()
                self.assert_no_app_error(imported)
                schedule = imported.session_state["resultado_programacao_maquinas"]["programacao"]
                self.assertEqual(int(schedule["QUANT. PROGRAMADA"].sum()), 150)
                self.assertEqual(set(schedule["CÓD."]), {"0000000007", "0000000008"})

    def test_english_quantities_reject_ambiguous_separators_before_calculation(self):
        app = new_session("en", PAGES[2])
        app.session_state["fonte_upload_estrutura_cortes"] = {
            "nome": "synthetic_structures.xlsx", "dados": structure_workbook().getvalue(),
        }
        app.run()
        app.text_area(key="texto_codigos_cortes").set_value("101").run()
        for invalid in ("1,000", "1.000", "1.5"):
            with self.subTest(quantity=invalid):
                app.text_area(key="texto_quantidades_cortes").set_value(invalid).run()
                app.button(key="botao_calcular_cortes").click().run()
                self.assertEqual(len(app.exception), 0)
                self.assertGreater(len(app.error), 0)
                self.assertNotIn("resultado_planejamento_cortes", app.session_state)
        app.text_area(key="texto_quantidades_cortes").set_value("1000").run()
        app.button(key="botao_calcular_cortes").click().run()
        self.assert_no_app_error(app)
        self.assertEqual(app.session_state["resultado_planejamento_cortes"]["resumo"]["total_de_pecas_cortadas"], 3000)

    def test_product_inputs_survive_language_switch(self):
        app = new_session("en", PAGES[1])
        app.text_area(key="texto_produtos_busca").set_value("Custom product\nAnother product").run()
        app.selectbox(key="language").set_value("pt").run()
        self.assert_no_app_error(app)
        self.assertEqual(app.text_area(key="texto_produtos_busca").value, "Custom product\nAnother product")


if __name__ == "__main__":
    unittest.main()
