"""Check the downloadable fictional workbooks and the complete sample workflow."""
from io import BytesIO
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

import pandas as pd
from pandas.testing import assert_frame_equal
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from calcular_cortes_streamlit import calcular_cortes_por_listas
from i18n import source_frame
from nome_V6_adaptado import buscar_produtos
from organizar_programacao_turnos import preparar_entrada


SAMPLES = ROOT / "samples"
STRUCTURES = SAMPLES / "sample_structures.xlsx"
MAPPING = SAMPLES / "sample_product_mapping.xlsx"
CUTS = SAMPLES / "sample_cut_plan.xlsx"
PAGES = [
    "📖 Instruções de uso", "🔎 Buscar códigos dos produtos",
    "✂️ Planejamento de cortes", "📅 Programação das máquinas",
]
PRODUCTS = ["DEMO CHAIR", "DEMO TABLE", "DEMO FRAME"]
CODES = ["900001", "900002", "900003"]
QUANTITIES = [100, 50, 25]
CUT_QUANTITIES = {"9900000001": 400, "9900000002": 200,
                  "9900000003": 100, "9900000004": 100}


def calculated_samples():
    return calcular_cortes_por_listas(CODES, QUANTITIES, arquivo_estrutura=STRUCTURES)


def sample_app(language="en"):
    app = AppTest.from_file(str(ROOT / "app.py"), default_timeout=15)
    app.session_state["language"] = language
    return app.run()


class SampleWorkbookTests(unittest.TestCase):
    def test_lookup_uses_fictional_aliases_and_rejects_discontinued_product(self):
        result = buscar_produtos(
            PRODUCTS + ["DEMO LEGACY CHAIR", "DEMO OLD CHAIR", "MISSING SAMPLE PRODUCT"],
            arquivo_estrutura=STRUCTURES, arquivo_de_para=MAPPING,
        )
        self.assertEqual(result["CÓDIGO"].iloc[:3].tolist(), CODES)
        self.assertEqual(result["PRODUTO ENCONTRADO"].iloc[:3].tolist(),
                         ["DEMO CHAIR A", "DEMO TABLE B", "DEMO FRAME C"])
        self.assertTrue(result["STATUS"].iloc[:3].str.startswith("Encontrado").all())
        self.assertTrue(result["CÓDIGO"].iloc[3:].isin(["Verificar", "Não encontrado"]).all())

    def test_downloadable_cut_plan_matches_calculation_from_structures(self):
        result = calculated_samples()
        self.assertEqual(result["nao_encontrados"], [])
        expected = preparar_entrada(result["cortes"])
        with pd.ExcelFile(CUTS, engine="openpyxl") as workbook:
            self.assertIn("Cuts", workbook.sheet_names)
            imported = preparar_entrada(source_frame(pd.read_excel(workbook, sheet_name="Cuts")))
        self.assertEqual(dict(zip(expected["CÓD."], expected["QUANT."])), CUT_QUANTITIES)
        self.assertEqual(int(expected["QUANT."].sum()), 800)
        columns = ["PRODUTO", "CÓD.", "TUBO", "VARA", "CORTE", "DIAG.",
                   "DESCRIÇÃO", "DESCRIÇÃO DO PRODUTO", "QUANT.", "QUANT"]
        # Excel may infer dimension strings as numbers; compare their normalized
        # values while retaining exact identifiers and descriptions.
        for frame in (expected, imported):
            for name in ("VARA", "CORTE", "QUANT.", "QUANT"):
                frame[name] = pd.to_numeric(frame[name])
            frame["DIAG."] = frame["DIAG."].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
        assert_frame_equal(
            expected[columns].sort_values("CÓD.").reset_index(drop=True),
            imported[columns].sort_values("CÓD.").reset_index(drop=True),
            check_dtype=False,
        )


class SampleAppTests(unittest.TestCase):
    def assert_healthy(self, app):
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])
        self.assertEqual(len(app.error), 0, [item.value for item in app.error])

    def assert_complete_schedule(self, app):
        self.assert_healthy(app)
        result = app.session_state["resultado_programacao_maquinas"]
        schedule = result["programacao"]
        self.assertEqual(int(schedule["QUANT. PROGRAMADA"].sum()), 800)
        self.assertTrue(result["nao_alocados"].empty)
        self.assertEqual(set(schedule["CÓD."]), set(CUT_QUANTITIES))
        # The example covers an ordinary tube, a mandatory saw cut, and an
        # angled cut. Allocation must retain those actual production rules.
        self.assertTrue(schedule.loc[schedule["CÓD."] == "9900000003", "MÁQUINA"]
                        .isin(["SERRA 1", "SERRA 2"]).all())
        self.assertTrue(schedule.loc[schedule["CÓD."] == "9900000004", "MÁQUINA"]
                        .isin(["SIMEC 1", "SIMEC 2", "SERRA 1", "SERRA 2"]).all())

    def test_sample_download_buttons_serve_the_bundled_workbooks(self):
        downloads = {}

        def capture(*args, **kwargs):
            if kwargs.get("key", "").startswith("sample_download_"):
                content = kwargs["data"]
                downloads[kwargs["key"]] = content.getvalue() if hasattr(content, "getvalue") else content
            return False

        for language in ("en", "pt"):
            with self.subTest(language=language), patch(
                "streamlit.delta_generator.DeltaGenerator.download_button", autospec=True, side_effect=capture
            ):
                app = sample_app(language)
                self.assert_healthy(app)
                self.assertEqual(downloads["sample_download_structures"], STRUCTURES.read_bytes())
                self.assertEqual(downloads["sample_download_mapping"], MAPPING.read_bytes())
                self.assertEqual(downloads["sample_download_cuts"], CUTS.read_bytes())

    def test_loaded_samples_complete_lookup_cutting_and_scheduling_in_both_languages(self):
        for language in ("en", "pt"):
            with self.subTest(language=language):
                app = sample_app(language)
                app.button(key="carregar_dados_exemplo").click().run()
                self.assert_healthy(app)
                self.assertEqual(app.radio(key="pagina").value, PAGES[0])
                for key, path in (
                    ("fonte_upload_estrutura_busca", STRUCTURES),
                    ("fonte_upload_estrutura_cortes", STRUCTURES),
                    ("fonte_upload_de_para", MAPPING),
                    ("fonte_upload_planejamento", CUTS),
                ):
                    self.assertEqual(app.session_state[key]["dados"], path.read_bytes())
                # A visitor can schedule immediately, before running lookup or
                # calculating cuts, using the included ready-made plan.
                app.radio(key="pagina").set_value(PAGES[3]).run()
                self.assertEqual(app.radio(key="origem_cortes").value, "Enviar planejamento em Excel")
                app.button(key="botao_organizar_programacao").click().run()
                self.assert_complete_schedule(app)
                app.radio(key="pagina").set_value(PAGES[1]).run()
                self.assertEqual(app.text_area(key="texto_produtos_busca").value, "\n".join(PRODUCTS))
                app.button(key="botao_processar_produtos").click().run()
                self.assert_healthy(app)
                self.assertEqual(app.session_state["resultado_busca_produtos"]["CÓDIGO"].tolist(), CODES)

                app.radio(key="pagina").set_value(PAGES[2]).run()
                self.assertEqual(app.text_area(key="texto_codigos_cortes").value, "\n".join(CODES))
                self.assertEqual(app.text_area(key="texto_quantidades_cortes").value, "100\n50\n25")
                app.button(key="botao_calcular_cortes").click().run()
                self.assert_healthy(app)
                self.assertEqual(app.session_state["resultado_planejamento_cortes"]["resumo"]
                                 ["total_de_pecas_cortadas"], 800)

                app.radio(key="pagina").set_value(PAGES[3]).run()
                app.radio(key="origem_cortes").set_value("Cortes calculados nesta sessão").run()
                app.button(key="botao_organizar_programacao").click().run()
                self.assert_complete_schedule(app)

                # The ready-made English plan also works through the upload
                # branch in either interface language.
                app.radio(key="origem_cortes").set_value("Enviar planejamento em Excel").run()
                app.button(key="botao_organizar_programacao").click().run()
                self.assert_complete_schedule(app)

    def test_reloading_samples_replaces_inputs_and_invalidates_previous_results(self):
        app = sample_app()
        app.button(key="carregar_dados_exemplo").click().run()
        app.radio(key="pagina").set_value(PAGES[2]).run()
        app.text_area(key="texto_quantidades_cortes").set_value("1\n1\n1").run()
        app.button(key="botao_calcular_cortes").click().run()
        self.assert_healthy(app)
        self.assertEqual(app.session_state["resultado_planejamento_cortes"]["resumo"]
                         ["total_de_pecas_cortadas"], 12)
        app.radio(key="pagina").set_value(PAGES[3]).run()
        app.session_state["editor_metas_gerais"] = {
            "edited_rows": {2: {"META/H": 123.0}}, "added_rows": [], "deleted_rows": [],
        }
        app.run()
        self.assertEqual(app.session_state["tabela_editor_metas_gerais"].iloc[2]["META/H"], 123)
        app.button(key="botao_organizar_programacao").click().run()
        self.assert_healthy(app)
        self.assertIn("resultado_programacao_maquinas", app.session_state)

        app.radio(key="pagina").set_value(PAGES[0]).run()
        app.button(key="carregar_dados_exemplo").click().run()
        self.assert_healthy(app)
        for key in ("resultado_busca_produtos", "resultado_planejamento_cortes", "resultado_programacao_maquinas"):
            self.assertNotIn(key, app.session_state)
        app.radio(key="pagina").set_value(PAGES[2]).run()
        self.assertEqual(app.text_area(key="texto_quantidades_cortes").value, "100\n50\n25")
        app.radio(key="pagina").set_value(PAGES[3]).run()
        self.assertEqual(app.session_state["tabela_editor_metas_gerais"]["META/H"].tolist(), [320] * 7)


if __name__ == "__main__":
    unittest.main()
