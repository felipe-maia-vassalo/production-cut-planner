"""Synthetic workflow regression checks; no company workbooks are required.

Run with: python -m unittest discover -s tests -v
"""
from io import BytesIO
from pathlib import Path
import sys
import unittest

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from pandas.testing import assert_frame_equal

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from calcular_cortes_streamlit import calcular_cortes_por_listas
from nome_V6_adaptado import buscar_produtos
from organizar_programacao_turnos import (
    construir_linha_do_tempo_turnos,
    organizar_programacao,
    resumir_turnos,
)


HEADERS = [
    "PRODUTO", "CÓD.", "TUBO", "VARA", "CORTE", "DIAG.",
    "DESCRIÇÃO", "DESCRIÇÃO DO PRODUTO", "QUANT",
]


def structure_workbook():
    book = Workbook()
    sheet = book.active
    sheet.title = "RELAÇÃO DE ESTRUTURAS"
    # An explanatory row verifies that header discovery is used.
    sheet.append(["Synthetic fixture — no production data"])
    sheet.append(HEADERS)
    sheet.append([101, 7, "20X20X1,2", 6000, 500, "", "Perna", "CADEIRA ATIVA", 2])
    # Blank product cells continue the preceding structure.
    sheet.append([None, 8, "20X20X1,2", 6000, 600, "", "Travessa", "CADEIRA ATIVA", 1])
    sheet.append([202, 9, "30X50", 6000, 2000, "", "Base", "MESA ANTIGA", 1])
    sheet.cell(5, HEADERS.index("DESCRIÇÃO DO PRODUTO") + 1).font = Font(color="FFFF0000")
    stream = BytesIO()
    book.save(stream)
    stream.seek(0)
    return stream


def mapping_workbook():
    stream = BytesIO()
    pd.DataFrame(
        [["CHAIR ALIAS", "CADEIRA ATIVA", 101], ["OLD TABLE ALIAS", "MESA ANTIGA", 202]],
        columns=["DESCRIÇÃO PROGRAMAÇÃO", "DESCRIÇÃO SIMPLIFICADA", "COD_PRODUTO"],
    ).to_excel(stream, index=False)
    stream.seek(0)
    return stream


def cuts_frame(quantity=1000, tube="20X20X1,2", cut=500):
    return pd.DataFrame([{
        "PRODUTO": "101", "CÓD.": "0000000007", "TUBO": tube,
        "VARA": 6000, "CORTE": cut, "DIAG.": "",
        "DESCRIÇÃO": "Perna", "DESCRIÇÃO DO PRODUTO": "CADEIRA ATIVA",
        "QUANT.": quantity, "QUANT": 2,
    }])


def cuts_workbook(frame):
    stream = BytesIO()
    frame.to_excel(stream, sheet_name="Cortes", index=False)
    stream.seek(0)
    return stream


class ProductLookupTests(unittest.TestCase):
    def test_exact_and_mapping_lookup_validate_active_structure(self):
        result = buscar_produtos(
            ["CADEIRA ATIVA", "CHAIR ALIAS"],
            arquivo_estrutura=structure_workbook(), arquivo_de_para=mapping_workbook(),
        )
        self.assertEqual(result["CÓDIGO"].tolist(), ["101", "101"])
        self.assertTrue(result["STATUS"].str.startswith("Encontrado").all())

    def test_discontinued_and_missing_products_are_not_accepted(self):
        result = buscar_produtos(
            ["MESA ANTIGA", "OLD TABLE ALIAS", "NONEXISTENT WIDGET"],
            arquivo_estrutura=structure_workbook(), arquivo_de_para=mapping_workbook(),
        )
        self.assertEqual(result["CÓDIGO"].tolist(), ["Não encontrado", "Verificar", "Não encontrado"])

    def test_missing_upload_does_not_fall_back_to_local_private_files(self):
        with self.assertRaises(ValueError):
            buscar_produtos(["CADEIRA ATIVA"], arquivo_de_para=mapping_workbook())


class CutCalculationTests(unittest.TestCase):
    def test_zero_padded_duplicate_products_and_continued_structure(self):
        result = calcular_cortes_por_listas(
            ["0000000101", "101.0", "999"], [3, 2, 4],
            arquivo_estrutura=structure_workbook(),
        )
        self.assertEqual(result["pedidos"].to_dict("records"), [
            {"PRODUTO": "101", "QUANTIDADE": 5},
            {"PRODUTO": "999", "QUANTIDADE": 4},
        ])
        self.assertEqual(result["cortes"]["CÓD."].tolist(), ["0000000007", "0000000008"])
        self.assertEqual(result["cortes"]["QUANT."].tolist(), [10, 5])
        self.assertEqual(result["nao_encontrados"], ["999"])
        self.assertEqual(result["resumo"]["total_de_pecas_cortadas"], 15)

    def test_list_mismatch_and_invalid_quantities_fail_before_calculation(self):
        with self.assertRaises(ValueError):
            calcular_cortes_por_listas(["101", "202"], [1], structure_workbook())
        for quantity in [0, -1, "1.5", "NaN"]:
            with self.subTest(quantity=quantity), self.assertRaises(ValueError):
                calcular_cortes_por_listas(["101"], [quantity], structure_workbook())

    def test_missing_structure_headers_are_reported(self):
        stream = BytesIO()
        pd.DataFrame({"wrong": [1]}).to_excel(stream, sheet_name="RELAÇÃO DE ESTRUTURAS", index=False)
        stream.seek(0)
        with self.assertRaisesRegex(ValueError, "Cabeçalho"):
            calcular_cortes_por_listas(["101"], [1], stream)


class SchedulingTests(unittest.TestCase):
    def test_capacity_respects_participation_losses_and_overtime(self):
        shifts = [{"nome": "Turno 1", "inicio": "07:00", "fim": "09:00", "maquinas": ["EMT"]}]
        scheduled, unallocated, summary = organizar_programacao(
            cuts_workbook(cuts_frame()), turnos=shifts,
            metas={"EMT": {"PADRAO": 60}},
            hora_extra_min={"EMT": 30}, perdas_min={"EMT": 15},
        )
        # Two hours minus one-hour break, plus 30 overtime minus 15 losses = 75 pieces.
        self.assertEqual(int(scheduled["QUANT. PROGRAMADA"].sum()), 75)
        self.assertEqual(int(unallocated["QUANT. RESTANTE"].sum()), 925)
        self.assertEqual(set(scheduled["MÁQUINA"]), {"EMT"})
        self.assertEqual(scheduled.iloc[0]["INÍCIO"], "07:15:00")
        self.assertEqual(scheduled.iloc[-1]["TÉRMINO"], "09:30:00")
        self.assertLessEqual(summary["UTILIZAÇÃO %"].max(), 100)

    def test_bottleneck_uses_full_shift_without_lunch_break(self):
        shifts = [{"nome": "Turno 1", "inicio": "07:00", "fim": "09:00", "maquinas": ["EMT"]}]
        scheduled, unallocated, _ = organizar_programacao(
            cuts_workbook(cuts_frame(100)), turnos=shifts,
            metas={"EMT": {"PADRAO": 60}}, maquinas_gargalo=["EMT"],
        )
        self.assertEqual(int(scheduled["QUANT. PROGRAMADA"].sum()), 100)
        self.assertTrue(unallocated.empty)
        self.assertEqual(scheduled.iloc[0]["TÉRMINO"], "08:40:00")

    def test_forced_saw_rule_cannot_be_bypassed_by_available_tube_machine(self):
        shifts = [{"nome": "Turno 1", "inicio": "07:00", "fim": "09:00", "maquinas": ["EMT"]}]
        scheduled, unallocated, _ = organizar_programacao(
            cuts_workbook(cuts_frame(20, cut=2000)), turnos=shifts,
        )
        self.assertTrue(scheduled.empty)
        self.assertEqual(int(unallocated["QUANT. RESTANTE"].sum()), 20)
        self.assertIn("não participam", unallocated.iloc[0]["MOTIVO"])

    def test_overnight_shift_and_overlap_validation(self):
        shifts = [
            {"nome": "Turno 1", "inicio": "22:00", "fim": "06:00"},
            {"nome": "Turno 2", "inicio": "06:00", "fim": "14:00"},
        ]
        timeline = construir_linha_do_tempo_turnos(shifts, maquina="EMT")
        self.assertEqual(timeline["minutos_produtivos"], 14 * 60)
        self.assertGreater(timeline["turnos"][1]["inicio_dt"], timeline["turnos"][0]["inicio_dt"])
        with self.assertRaises(ValueError):
            construir_linha_do_tempo_turnos([
                {"nome": "Turno 1", "inicio": "07:00", "fim": "17:00"},
                {"nome": "Turno 2", "inicio": "16:00", "fim": "20:00"},
            ])

    def test_unavailable_last_shift_cannot_receive_overtime(self):
        with self.assertRaises(ValueError):
            construir_linha_do_tempo_turnos(
                [{"inicio": "07:00", "fim": "17:00", "maquinas": ["BLM"]}],
                maquina="EMT", hora_extra_min=30,
            )


class LocalizationTests(unittest.TestCase):
    def setUp(self):
        import i18n
        self.i18n = i18n
        self.i18n.set_language("en")

    def tearDown(self):
        self.i18n.set_language("pt")

    def test_english_cut_export_roundtrip_preserves_engine_inputs(self):
        original = cuts_frame()
        english = self.i18n.localized_frame(original)
        self.assertNotEqual(list(english.columns), list(original.columns))
        assert_frame_equal(original, self.i18n.source_frame(english))
        self.assertEqual(self.i18n.excel_sheet("Cortes"), "Cuts")
        # Data is localized for display on a copy, never the canonical source.
        self.assertIn("QUANT.", original.columns)

    def test_validation_errors_are_translated_without_losing_product_code(self):
        portuguese = "A quantidade do produto 101 deve ser maior que zero."
        english = self.i18n.localized_error(ValueError(portuguese))
        self.assertNotEqual(english, portuguese)
        self.assertIn("101", english)
        self.i18n.set_language("pt")
        self.assertEqual(self.i18n.localized_error(ValueError(portuguese)), portuguese)

    def test_language_selection_preserves_user_data_and_changes_owned_text(self):
        original = pd.DataFrame({
            "CÓDIGO": ["Não encontrado"], "STATUS": ["Não encontrado"],
            "DESCRIÇÃO DO PRODUTO": ["Não encontrado"], "PRODUTO OP": ["Custom product"],
        })
        english = self.i18n.localized_frame(original)
        self.assertNotEqual(english.iloc[0, 0], "Não encontrado")
        self.assertEqual(english.iloc[0, 2], "Não encontrado")
        self.assertEqual(self.i18n.tr("Processar", "Process"), "Process")
        self.i18n.set_language("pt")
        assert_frame_equal(original, self.i18n.localized_frame(original))
        self.assertEqual(self.i18n.tr("Processar", "Process"), "Processar")


if __name__ == "__main__":
    unittest.main()
