# Production Cut Planner / Corte de Produção

**[Open the live app](https://factory-cut-planner.streamlit.app/)** · [GitHub repository](https://github.com/felipe-maia-vassalo/production-cut-planner)

Streamlit application for finding product codes, calculating tube cuts from a bill of materials, and scheduling production across machines and shifts. Choose **English** or **Português (Brasil)** in the sidebar. English is the default language.

The interface, instructions, validation messages, result tables, and Excel downloads follow the selected language. Product names, codes, machine names, and uploaded business descriptions retain their original values. The original Portuguese Excel input templates remain supported. Cut plans exported in either language can be uploaded again for scheduling.

## Run locally

Use Python 3.12 or 3.13. From the project folder on macOS or Linux:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

On Windows:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Workflow

1. **Find product codes:** upload the structures workbook and the product mapping workbook, paste one product name per line, and process. Review ambiguous or missing matches.
2. **Cut planning:** upload the structures workbook and enter one product code and one quantity per line. The lists must have the same length. In English, enter quantities as digits without separators (for example `1000`). Calculate and review the cuts.
3. **Machine scheduling:** use cuts from the current session or upload an exported cut plan. Configure shifts, participating machines, bottlenecks, production targets, expected losses, and authorized overtime. Generate the schedule and review any unallocated pieces.

For detailed operating instructions, open the app's **Instructions** page.

## Excel input formats

Keep these original sheet and header names even when the interface is in English:

| Workbook | Required sheet / headers |
| --- | --- |
| Structures / bill of materials | Sheet `RELAÇÃO DE ESTRUTURAS`; columns `PRODUTO`, `CÓD.`, `QUANT`. Scheduling also uses `TUBO`, `VARA`, `CORTE`, `DIAG.`, `DESCRIÇÃO`, `DESCRIÇÃO DO PRODUTO`. |
| Product mapping / DE-PARA | First sheet; columns `DESCRIÇÃO PROGRAMAÇÃO`, `DESCRIÇÃO SIMPLIFICADA`, `COD_PRODUTO`. |
| Exported cut plan | Sheet `Cortes` (Portuguese) or `Cuts` (English); retain the exported headers. |

Supported uploads are `.xlsx` and `.xlsm`, up to 50 MB per file. Save older `.xls` files as `.xlsx` first. Macros are not executed.

## Files checked against the original app

| Original file | Role in this release |
| --- | --- |
| `app_V4.py` | Current interface, published as `app.py` with language support. |
| `calcular_cortes_streamlit.py` | Actual imported cutting engine. Includes normalization that treats zero-padded and unpadded product codes as equivalent. |
| `nome_V6_adaptado.py` | Product lookup engine. |
| `organizar_programacao_turnos.py` | Scheduling engine and shift configuration. |
| `calcular_cortes_streamlit_V1.py` | Older copy; not imported by `app_V4.py` and not needed in this release. |

The original Desktop files are unchanged. The processing engines retain the original production rules; the public copy removes the machine-specific local path and personal example filename. `i18n.py` handles presentation translations without changing internal calculation fields.

## Session data

No company spreadsheets are bundled in this repository. Uploads and results are held in the current Streamlit session; download results before closing or starting a new session. This app does not write uploaded workbooks to the repository. On a hosted deployment, workbook contents are processed on the hosting server. Use suitable test data when sharing a public demo.

## Tests

```sh
python -m unittest discover -s tests -v
```

All **19 automated tests passed**. Tests generate synthetic Excel workbooks in memory and cover lookup, cutting, scheduling, language switching with edited settings, and translation compatibility. No company data is included.

Live cloud verification on 1 November 2026 (Python 3.13): synthetic workbook uploads; two found product names and one missing product; 1,000 products producing 3,000 cut pieces; all 3,000 pieces scheduled across two machines; Portuguese/English switching preserved the schedule; downloaded English cut/schedule workbooks reopened successfully with correct totals and zero-padded cut codes. Load testing and production-data validation were not performed.

## Deploy on Streamlit Community Cloud

Connect this GitHub repository, select branch `main`, entry point `app.py`, and Python 3.13. No secrets or paid APIs are required. See the [official deployment instructions](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy).

## Português

Aplicativo para buscar códigos, calcular cortes e programar máquinas por turno. Selecione **Português (Brasil)** na barra lateral. As planilhas originais continuam compatíveis; mantenha os nomes das abas e dos cabeçalhos indicados acima. Os arquivos originais do computador foram preservados. Não há planilhas da empresa neste repositório. Para executar, siga os comandos de instalação acima e abra `python -m streamlit run app.py`.
