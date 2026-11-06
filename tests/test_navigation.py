"""Previous/next navigation keeps workflow data and sidebar selection aligned."""
from pathlib import Path
from datetime import time
import unittest

from pandas.testing import assert_frame_equal
from streamlit.testing.v1 import AppTest

from test_samples import calculated_samples


APP = Path(__file__).resolve().parents[1] / "app.py"
PAGES = [
    "📖 Instruções de uso", "🔎 Buscar códigos dos produtos",
    "✂️ Planejamento de cortes", "📅 Programação das máquinas",
]


def session(language="en", page=PAGES[0]):
    app = AppTest.from_file(str(APP), default_timeout=15)
    app.session_state["language"] = language
    app.session_state["pagina"] = page
    return app.run()


class NavigationTests(unittest.TestCase):
    def assert_healthy(self, app):
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])
        self.assertEqual(len(app.error), 0, [item.value for item in app.error])

    def click_page(self, app, key, expected_page):
        app.button(key=key).click().run()
        self.assert_healthy(app)
        self.assertEqual(app.radio(key="pagina").value, expected_page)
        self.assertEqual(app.session_state["pagina"], expected_page)

    def test_back_navigation_is_available_before_any_activity_is_completed(self):
        for language in ("en", "pt"):
            with self.subTest(language=language):
                app = session(language)
                self.click_page(app, "nav_instrucoes_seguinte", PAGES[1])
                self.assertTrue(app.button(key="nav_busca_seguinte").disabled)
                self.click_page(app, "nav_busca_anterior", PAGES[0])

                app.radio(key="pagina").set_value(PAGES[2]).run()
                self.assertTrue(app.button(key="nav_cortes_seguinte").disabled)
                self.click_page(app, "nav_cortes_anterior", PAGES[1])

                # Scheduling returns early when no workbook/cuts are supplied;
                # its footer must still let the visitor return to the prior page.
                app.radio(key="pagina").set_value(PAGES[3]).run()
                self.click_page(app, "nav_programacao_anterior", PAGES[2])
                app.radio(key="pagina").set_value(PAGES[3]).run()
                self.click_page(app, "nav_programacao_inicio", PAGES[0])

    def test_empty_results_do_not_enable_next_step(self):
        for language in ("en", "pt"):
            with self.subTest(language=language):
                app = session(language, PAGES[2])
                empty_result = calculated_samples()
                empty_result["cortes"] = empty_result["cortes"].iloc[:0].copy()
                empty_result["resumo"]["total_de_pecas_cortadas"] = 0
                empty_result["resumo"]["codigos_de_corte_distintos"] = 0
                app.session_state["resultado_planejamento_cortes"] = empty_result
                app.run()
                self.assert_healthy(app)
                self.assertTrue(app.button(key="nav_cortes_seguinte").disabled)
                self.click_page(app, "nav_cortes_anterior", PAGES[1])

    def test_sample_workflow_uses_next_and_back_without_losing_inputs_or_results(self):
        for language in ("en", "pt"):
            with self.subTest(language=language):
                app = session(language)
                app.button(key="carregar_dados_exemplo").click().run()
                self.click_page(app, "nav_instrucoes_seguinte", PAGES[1])
                products = app.text_area(key="texto_produtos_busca").value
                app.button(key="botao_processar_produtos").click().run()
                self.assert_healthy(app)
                lookup = app.session_state["resultado_busca_produtos"].copy()
                self.assertFalse(app.button(key="nav_busca_seguinte").disabled)
                self.click_page(app, "nav_busca_seguinte", PAGES[2])

                codes = app.text_area(key="texto_codigos_cortes").value
                quantities = app.text_area(key="texto_quantidades_cortes").value
                app.button(key="botao_calcular_cortes").click().run()
                self.assert_healthy(app)
                cuts = app.session_state["resultado_planejamento_cortes"]["cortes"].copy()
                self.assertEqual(int(cuts["QUANT."].sum()), 800)
                self.assertFalse(app.button(key="nav_cortes_seguinte").disabled)
                # Simulate a prior uploaded plan selection. Next must use the
                # newly calculated cuts, rather than a previously uploaded plan.
                app.session_state["origem_cortes"] = "Enviar planejamento em Excel"
                self.click_page(app, "nav_cortes_seguinte", PAGES[3])
                self.assertEqual(app.radio(key="origem_cortes").value,
                                 "Cortes calculados nesta sessão")
                app.button(key="botao_organizar_programacao").click().run()
                self.assert_healthy(app)
                schedule = app.session_state["resultado_programacao_maquinas"]["programacao"].copy()
                self.assertEqual(int(schedule["QUANT. PROGRAMADA"].sum()), 800)
                self.assertTrue(app.session_state["resultado_programacao_maquinas"]["nao_alocados"].empty)

                # Browser cleanup removes widget keys for hidden pages. The
                # durable input copies must be restored by Back navigation.
                for key in ("texto_produtos_busca", "texto_codigos_cortes", "texto_quantidades_cortes"):
                    if key in app.session_state:
                        del app.session_state[key]
                self.click_page(app, "nav_programacao_anterior", PAGES[2])
                self.assertEqual(app.text_area(key="texto_codigos_cortes").value, codes)
                self.assertEqual(app.text_area(key="texto_quantidades_cortes").value, quantities)
                assert_frame_equal(app.session_state["resultado_planejamento_cortes"]["cortes"], cuts)
                assert_frame_equal(app.session_state["resultado_programacao_maquinas"]["programacao"], schedule)

                self.click_page(app, "nav_cortes_anterior", PAGES[1])
                self.assertEqual(app.text_area(key="texto_produtos_busca").value, products)
                assert_frame_equal(app.session_state["resultado_busca_produtos"], lookup)
                self.click_page(app, "nav_busca_anterior", PAGES[0])
                self.assertIn("resultado_programacao_maquinas", app.session_state)

    def test_navigation_button_labels_follow_the_selected_language(self):
        labels = {}
        page_keys = [
            (PAGES[0], ["nav_instrucoes_seguinte"]),
            (PAGES[1], ["nav_busca_anterior", "nav_busca_seguinte"]),
            (PAGES[2], ["nav_cortes_anterior", "nav_cortes_seguinte"]),
            (PAGES[3], ["nav_programacao_anterior", "nav_programacao_inicio"]),
        ]
        for language in ("en", "pt"):
            app = session(language)
            labels[language] = {}
            for page, keys in page_keys:
                app.radio(key="pagina").set_value(page).run()
                self.assert_healthy(app)
                for key in keys:
                    labels[language][key] = app.button(key=key).label
        self.assertEqual(set(labels["en"]), set(labels["pt"]))
        for key in labels["en"]:
            with self.subTest(button=key):
                self.assertNotEqual(labels["en"][key], labels["pt"][key])

    def test_nondefault_scheduling_settings_and_results_survive_back_and_next(self):
        app = session()
        app.button(key="carregar_dados_exemplo").click().run()
        self.click_page(app, "nav_instrucoes_seguinte", PAGES[1])
        app.button(key="botao_processar_produtos").click().run()
        self.click_page(app, "nav_busca_seguinte", PAGES[2])
        app.button(key="botao_calcular_cortes").click().run()
        self.click_page(app, "nav_cortes_seguinte", PAGES[3])

        app.number_input(key="quantidade_turnos_programacao").set_value(2).run()
        controls = {
            "inicio_turno_1": time(6), "fim_turno_1": time(14),
            "inicio_turno_2": time(14), "fim_turno_2": time(22),
        }
        for key, value in controls.items():
            app.time_input(key=key).set_value(value).run()
        for number in (1, 2):
            app.multiselect(key=f"maquinas_participantes_turno_{number}").set_value(["EMT"]).run()
        app.multiselect(key="maquinas_gargalo").set_value(["EMT"]).run()
        app.session_state["editor_capacidade"] = {
            "edited_rows": {2: {"PERDAS (min)": 10.0, "HORA EXTRA (min)": 15.0}},
            "added_rows": [], "deleted_rows": [],
        }
        app.run()
        self.assert_healthy(app)
        app.button(key="botao_organizar_programacao").click().run()
        self.assert_healthy(app)
        result = app.session_state["resultado_programacao_maquinas"]["programacao"].copy()
        self.assertFalse(result.empty)
        capacity = app.session_state["tabela_editor_capacidade"].copy()
        self.click_page(app, "nav_programacao_anterior", PAGES[2])
        hidden_controls = list(controls) + [
            "quantidade_turnos_programacao", "maquinas_gargalo", "origem_cortes",
            "maquinas_participantes_turno_1", "maquinas_participantes_turno_2",
        ]
        for key in hidden_controls:
            if key in app.session_state:
                del app.session_state[key]
        app.run()
        self.click_page(app, "nav_cortes_seguinte", PAGES[3])
        self.assertEqual(app.number_input(key="quantidade_turnos_programacao").value, 2)
        for key, expected in controls.items():
            self.assertEqual(app.time_input(key=key).value, expected)
        for number in (1, 2):
            self.assertEqual(app.multiselect(key=f"maquinas_participantes_turno_{number}").value, ["EMT"])
        self.assertEqual(app.multiselect(key="maquinas_gargalo").value, ["EMT"])
        assert_frame_equal(app.session_state["tabela_editor_capacidade"], capacity)
        assert_frame_equal(app.session_state["resultado_programacao_maquinas"]["programacao"], result)


if __name__ == "__main__":
    unittest.main()
