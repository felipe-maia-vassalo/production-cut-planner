from hashlib import sha256
import re
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta, time
from i18n import (
    set_language, tr, localized_frame, source_frame, localized_error,
    excel_sheet, format_integer, column_label,
)

from nome_V6_adaptado import buscar_produtos
from calcular_cortes_streamlit import calcular_cortes_por_listas
from organizar_programacao_turnos import (
    organizar_programacao, resumir_turnos, MAQUINAS, MAQUINAS_GARGALO_PADRAO,
    metas_padrao, validar_configuracao, familia_tubo, preparar_entrada,
)


# ==========================================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================================
def preservar_editores_ao_trocar_idioma():
    """Restore the latest edited tables when translated column labels change."""
    for chave in ("editor_capacidade", "editor_metas_gerais", "editor_excecoes"):
        ultima = st.session_state.get("tabela_" + chave)
        if ultima is not None:
            st.session_state["base_" + chave] = ultima.copy()
        st.session_state.pop(chave, None)


set_language(st.session_state.get("language", "en"))
st.set_page_config(
    page_title=tr("Corte de Produção", 'Production Cut Planner'),
    page_icon="🏭",
    layout="wide",
)

language = st.sidebar.selectbox(
    "Language / Idioma", options=["en", "pt"],
    format_func=lambda value: {"pt": "Português (Brasil)", "en": "English"}[value],
    key="language",
    on_change=preservar_editores_ao_trocar_idioma,
)
set_language(language)


# ==========================================================
# FUNÇÕES AUXILIARES
# ==========================================================
def gerar_excel(dataframe, nome_aba="Resultado", localizar=True):
    """Transforma um DataFrame em arquivo Excel em memória."""
    arquivo_excel = BytesIO()

    with pd.ExcelWriter(
        arquivo_excel,
        engine="openpyxl",
    ) as writer:
        (localized_frame(dataframe) if localizar else dataframe).to_excel(
            writer,
            index=False,
            sheet_name=excel_sheet(nome_aba) if localizar else nome_aba,
        )

    arquivo_excel.seek(0)
    return arquivo_excel


def gerar_excel_programacao(programacao, nao_alocados, resumo, turnos_df=None):
    """Gera o arquivo final da programação em memória."""
    arquivo_excel = BytesIO()

    with pd.ExcelWriter(
        arquivo_excel,
        engine="openpyxl",
    ) as writer:
        localized_frame(programacao).to_excel(
            writer,
            index=False,
            sheet_name=excel_sheet("Programacao"),
        )
        localized_frame(nao_alocados).to_excel(
            writer,
            index=False,
            sheet_name=excel_sheet("Nao_alocados"),
        )
        localized_frame(resumo).to_excel(
            writer,
            index=False,
            sheet_name=excel_sheet("Resumo"),
        )

        nomes_abas = [excel_sheet(nome) for nome in ("Programacao", "Nao_alocados", "Resumo")]

        if turnos_df is not None and not turnos_df.empty:
            localized_frame(turnos_df).to_excel(
                writer,
                index=False,
                sheet_name=excel_sheet("Turnos"),
            )
            nomes_abas.append(excel_sheet("Turnos"))

        for nome_aba in nomes_abas:
            ws = writer.book[nome_aba]
            ws.freeze_panes = "A2"

            for coluna in ws.columns:
                letra = coluna[0].column_letter
                maior = max(
                    len("" if celula.value is None else str(celula.value))
                    for celula in coluna
                )
                ws.column_dimensions[letra].width = min(
                    max(maior + 2, 10),
                    32,
                )

    arquivo_excel.seek(0)
    return arquivo_excel



def somar_horas_horario(horario, horas=10):
    """Soma horas a um datetime.time, permitindo virar a meia-noite."""
    base = datetime(2000, 1, 1, horario.hour, horario.minute, horario.second)
    return (base + timedelta(hours=horas)).time()


def linhas_preenchidas(texto):
    """Converte um texto com um item por linha em uma lista limpa."""
    return [
        linha.strip()
        for linha in texto.splitlines()
        if linha.strip()
    ]


def atualizar_upload(chave, resultados):
    """Mantém uploads na sessão mesmo ao navegar e invalida resultados antigos."""
    enviado = st.session_state.get(chave)
    fonte = None if enviado is None else {"nome": enviado.name, "dados": enviado.getvalue()}
    anterior = st.session_state.get("fonte_" + chave)
    if fonte != anterior:
        st.session_state["fonte_" + chave] = fonte
        for resultado in resultados:
            st.session_state.pop(resultado, None)


def upload_planilha(titulo, descricao, chave, resultados):
    with st.container(border=True):
        st.subheader(titulo)
        st.write(descricao)
        st.file_uploader(
            tr("Selecione a planilha Excel", 'Select an Excel spreadsheet'), type=["xlsx", "xlsm"], key=chave,
            on_change=atualizar_upload, args=(chave, resultados))
        fonte = st.session_state.get("fonte_" + chave)
        if fonte:
            st.success(tr("Arquivo em uso nesta sessão: " + fonte["nome"], 'File in use this session: ' + fonte['nome']))
            return BytesIO(fonte["dados"])
        st.info(tr("Envie este arquivo para continuar nesta atividade.", 'Upload this file to continue with this task.'))
    return None


def editor_persistente(tabela, **opcoes):
    """Mantém os dados editados ao navegar, sem atribuir valor ao widget."""
    chave = opcoes["key"]
    base, ultima = "base_" + chave, "tabela_" + chave
    if base not in st.session_state or chave not in st.session_state:
        st.session_state[base] = st.session_state.get(ultima, tabela).copy()
    resultado = st.data_editor(st.session_state[base], **opcoes)
    st.session_state[ultima] = resultado.copy()
    return resultado


def pagina_instrucoes():
    st.title(tr("🏭 Corte de Produção — Como usar", '🏭 Production Cut Planner — How to use'))
    st.write(tr("Use o menu à esquerda para abrir cada atividade. Envie as planilhas nos campos indicados; você não precisa organizar pastas no computador.", 'Use the menu on the left to open each task. Upload your spreadsheets in the indicated fields.'))
    st.subheader(tr("Antes de começar", 'Before you start'))
    st.markdown(tr("""
- Tenha as planilhas atualizadas em **Excel (.xlsx ou .xlsm)**. Mantenha os nomes das abas e os cabeçalhos originais.
- A **Relação Simplificada de Estruturas** deve ter a aba **RELAÇÃO DE ESTRUTURAS**.
- O **DE-PARA** deve ter, na primeira aba, as colunas **DESCRIÇÃO PROGRAMAÇÃO**, **DESCRIÇÃO SIMPLIFICADA** e **COD_PRODUTO**.
- Cada campo aceita um arquivo. Confira o nome mostrado em **Arquivo em uso nesta sessão**.
""", '\n- Use up-to-date **Excel (.xlsx or .xlsm)** files. Keep the original sheet names and column headers in your source workbooks.\n- The **bill of materials** must contain the **RELAÇÃO DE ESTRUTURAS** sheet.\n- The **product name mapping** must have **DESCRIÇÃO PROGRAMAÇÃO**, **DESCRIÇÃO SIMPLIFICADA** and **COD_PRODUTO** in its first sheet.\n- Each upload field accepts one file. Check the name shown under **File in use this session**.\n'))
    st.subheader(tr("1. Buscar códigos dos produtos", '1. Find product codes'))
    st.markdown(tr("""
1. Envie a **Relação de Estruturas** no primeiro campo.
2. Envie o **DE-PARA** no segundo campo. Exemplo de nome: de_para_produtos.xlsx.
3. Cole os nomes dos produtos da ordem de produção, **um por linha**, e clique em **Processar**.
4. Confira os resultados. Produtos marcados como **Verificar** ou **Não encontrado** precisam de conferência antes de seguir. Estruturas descontinuadas, identificadas pela cor vermelha, não são usadas na busca.
5. Copie os códigos ou baixe o resultado em Excel.
""", '\n1. Upload the **bill of materials** in the first field.\n2. Upload the **product name mapping** in the second field, for example product_mapping.xlsx.\n3. Paste the production order product names, **one per line**, and click **Process**.\n4. Review the results. Products marked **Check** or **Not found** need checking before you continue. Discontinued structures, marked in red in the source workbook, are excluded from the search.\n5. Copy the codes or download the results as Excel.\n'))
    st.subheader(tr("2. Planejamento de cortes", '2. Cut planning'))
    st.markdown(tr("""
1. Envie a **Relação de Estruturas** no campo desta página. Pode ser a mesma planilha usada na busca.
2. Cole os códigos e as quantidades, **um por linha**, na mesma ordem. A quantidade da primeira linha corresponde ao código da primeira linha, e assim por diante.
3. Informe quantidades inteiras maiores que zero e clique em **Calcular cortes**.
4. Confira os produtos não encontrados e os cortes calculados. **QUANT.** é o total de peças a produzir; **QUANT** é o número de peças daquele corte por produto.
5. Baixe o planejamento em Excel, se desejar. Os cortes também ficam disponíveis diretamente na página de programação durante esta sessão.

A aba de estruturas deve conter **PRODUTO, CÓD. e QUANT**, além de **TUBO, VARA, CORTE, DIAG., DESCRIÇÃO e DESCRIÇÃO DO PRODUTO** para a programação das máquinas.
""", '\n1. Upload the **bill of materials** on this page. You may use the same spreadsheet as in product search.\n2. Paste product codes and quantities, **one per line**, in matching order. The first quantity belongs to the first code, and so on.\n3. Enter positive whole-number quantities and click **Calculate cuts**.\n4. Review missing products and calculated cuts. **Total quantity** is the total to produce; **Quantity per product** is the number of pieces of that cut per product.\n5. Download the cut plan as Excel if needed. The calculated cuts are also available directly on the scheduling page during this session.\n\nThe source structures sheet must contain **PRODUTO, CÓD. and QUANT**, plus **TUBO, VARA, CORTE, DIAG., DESCRIÇÃO and DESCRIÇÃO DO PRODUTO** for machine scheduling. Keep these source headers in Portuguese.\n'))
    st.subheader(tr("3. Programação das máquinas", '3. Machine scheduling'))
    st.markdown(tr("""
1. Escolha **Cortes calculados nesta sessão** ou **Enviar planejamento em Excel**. Se enviar um Excel, use a aba **Cortes**, com as colunas do planejamento exportado pelo app.
2. Defina a quantidade de turnos e seus horários. Em cada turno, remova da seleção as máquinas que não participarão, por exemplo, por estarem paradas. O fim antes do início indica término no dia seguinte. Os turnos não podem se sobrepor; turnos com pausa precisam ter pelo menos duas horas.
3. Selecione as **máquinas gargalo** que trabalham durante o almoço. Nenhuma vem selecionada. As demais fazem uma pausa de uma hora no meio de cada turno; por exemplo, 07:00–17:00 gera pausa de 12:00–13:00.
4. Informe **perdas previstas** e **hora extra autorizada**, em minutos, por máquina. Use zero quando não houver. As perdas reservam tempo no início da programação; a hora extra fica disponível após o último turno e só é usada quando necessária.
5. Edite a **meta geral por máquina** diretamente na tabela. A tabela de **exceções por tubo** substitui a meta geral somente para a combinação informada. Você pode adicionar ou excluir exceções. Metas devem ser maiores que zero.
6. Clique em **Organizar programação das máquinas**. Confira o resumo, as abas por máquina e as peças não alocadas. Baixe o resultado em Excel.

As metas iniciais são **320 peças/hora**; nas SIMECs, **20X48 = 200/h** e **40X77 = 110/h**. O tubo 7/8 está liberado para BLM. Pedidos com mesmo tubo, corte e grau podem compartilhar a mesma máquina, incluindo PDV, respeitando as regras de máquina. O setup já está incluído em H/MÁQ.
""", '\n1. Choose **Cuts calculated this session** or **Upload Excel cut plan**. Uploaded plans must have a **Cortes** or **Cuts** sheet with the columns exported by this app. Both Portuguese and English exports are accepted.\n2. Set the number of shifts and their times. Remove machines that will not participate in each shift, for example due to downtime. An end time before the start time means the shift ends the next day. Shifts must not overlap; shifts with a break must last at least two hours.\n3. Select the **bottleneck machines** that run through lunch. None are selected initially. Other machines take a one-hour break starting at the midpoint of each shift; for example, 07:00–17:00 gives a 12:00–13:00 break.\n4. Enter **expected downtime** and **authorized overtime**, in minutes, for each machine. Use zero when none applies. Downtime reserves time at the beginning of the schedule; overtime is available after the last shift and used only when needed.\n5. Edit the **general production target per machine** in the table. The **tube exceptions** table overrides the general target only for the specified machine and tube. You can add or delete exceptions. Targets must be greater than zero.\n6. Click **Schedule machines**. Review the summary, machine tabs and unallocated pieces. Download the results as Excel.\n\nInitial targets are **320 pieces/hour**; for SIMEC machines, **20X48 = 200/hour** and **40X77 = 110/hour**. Tube 7/8 is allowed on BLM. Orders with the same tube, cut and angle may share a machine, including PDV, subject to machine rules. Setup time is already included in machine hours.\n'))
    st.subheader(tr("Conferência e arquivos da sessão", 'Review and session files'))
    st.markdown(tr("""
- **Peças não alocadas** podem indicar falta de capacidade ou ausência de regra para o tubo. Confira antes de liberar a produção.
- Ao trocar uma planilha, refaça a atividade. Mudanças nos cortes, turnos, máquinas participantes, metas, perdas ou gargalos exigem recalcular a programação.
- Os uploads, resultados e metas editadas ficam na **sessão atual**. Ao iniciar uma nova sessão, envie novamente as planilhas e ajuste as metas. Baixe os resultados que precisar guardar.
- Se uma planilha der erro, confira o campo de envio, a aba e os cabeçalhos indicados. Para .xls antigo, salve uma cópia em .xlsx no Excel.
""", '\n- **Unallocated pieces** may indicate insufficient capacity or a missing tube rule. Check them before releasing production.\n- Run the task again after changing a source workbook. Changes to cuts, shifts, participating machines, targets, downtime or bottlenecks require recalculating the schedule.\n- Uploads, results and edited targets remain in the **current session**. Upload files and set targets again in a new session. Download any results you need to keep.\n- If a spreadsheet fails, check the upload field, sheet name and required headers. Save legacy .xls files as .xlsx in Excel first.\n'))


# ==========================================================
# PÁGINA 1 - BUSCAR CÓDIGOS DOS PRODUTOS
# ==========================================================
def pagina_buscar_codigos():
    st.title(tr("🏭 Corte de Produção", '🏭 Production Cut Planner'))

    st.write(
        tr("Cole abaixo os produtos da Ordem de Produção.", 'Paste the production order product names below.')
    )

    arquivo_estrutura = upload_planilha(
        tr("1. Relação de Estruturas", '1. Bill of materials'), tr("Envie a Relação Simplificada de Estruturas (.xlsm ou .xlsx), com a aba RELAÇÃO DE ESTRUTURAS.", 'Upload the bill of materials (.xlsm or .xlsx), with the RELAÇÃO DE ESTRUTURAS sheet.'),
        "upload_estrutura_busca", ["resultado_busca_produtos"])
    arquivo_de_para = upload_planilha(
        tr("2. DE-PARA dos produtos", '2. Product name mapping'), tr("Envie o DE-PARA que relaciona os nomes da programação aos nomes simplificados. Exemplo: de_para_produtos.xlsx.", 'Upload the mapping between production order names and simplified product names, for example product_mapping.xlsx.'),
        "upload_de_para", ["resultado_busca_produtos"])

    texto_produtos = st.text_area(
        tr("Produtos", 'Products'),
        height=300,
        placeholder=(
            "MP-CT-3\n"
            "P-3 T/R\n"
            "P-6 T/R\n"
            "EST-5 18L\n"
            "MCPE-3 T/R"
        ),
        key="texto_produtos_busca",
    )

    if st.button(
        tr("Processar", 'Process'),
        type="primary",
        key="botao_processar_produtos",
        disabled=arquivo_estrutura is None or arquivo_de_para is None,
    ):
        lista_produtos = linhas_preenchidas(texto_produtos)

        if not lista_produtos:
            st.warning(tr("Cole pelo menos um produto.", 'Paste at least one product.'))

        else:
            try:
                with st.spinner(tr("Buscando produtos...", 'Searching for products...')):
                    resultado = buscar_produtos(lista_produtos, arquivo_estrutura=arquivo_estrutura, arquivo_de_para=arquivo_de_para)

                st.session_state["resultado_busca_produtos"] = resultado

            except Exception as erro:
                st.error(tr(f"Erro ao processar: {erro}", f'Processing error: {localized_error(erro)}'))

    # ======================================================
    # RESULTADO DA BUSCA
    # ======================================================
    if "resultado_busca_produtos" in st.session_state:
        resultado = st.session_state["resultado_busca_produtos"]

        st.divider()
        st.subheader(tr("Resultado", 'Results'))

        total = len(resultado)

        encontrados = (
            ~resultado["CÓDIGO"]
            .astype(str)
            .isin(["Verificar", "Não encontrado"])
        ).sum()

        verificar = (
            resultado["CÓDIGO"]
            .astype(str)
            .eq("Verificar")
            .sum()
        )

        nao_encontrados = (
            resultado["CÓDIGO"]
            .astype(str)
            .eq("Não encontrado")
            .sum()
        )

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Total", total)
        col2.metric(tr("Encontrados", 'Found'), encontrados)
        col3.metric(tr("Verificar", 'Review'), verificar)
        col4.metric(tr("Não encontrados", 'Not found'), nao_encontrados)

        colunas_principais = [
            coluna
            for coluna in [
                "PRODUTO OP",
                "PRODUTO ENCONTRADO",
                "CÓDIGO",
            ]
            if coluna in resultado.columns
        ]

        st.dataframe(
            localized_frame(resultado[colunas_principais]),
            use_container_width=True,
            hide_index=True,
        )

        problemas = resultado[
            resultado["CÓDIGO"]
            .astype(str)
            .isin(["Verificar", "Não encontrado"])
        ]

        if not problemas.empty:
            st.warning(
                tr("Existem produtos que precisam de conferência manual.", 'Some products need manual review.')
            )

            st.dataframe(
                localized_frame(problemas),
                use_container_width=True,
                hide_index=True,
            )

        st.divider()
        st.subheader(tr("Códigos para copiar", 'Codes to copy'))

        codigos = "\n".join(
            localized_frame(resultado[["CÓDIGO"]]).iloc[:, 0]
            .astype(str)
            .tolist()
        )

        st.session_state["codigos_para_copiar"] = codigos
        st.text_area(
            tr("Copie e cole no Excel:", 'Copy and paste into Excel:'),
            height=250,
            key="codigos_para_copiar",
        )

        arquivo_excel = gerar_excel(
            resultado,
            nome_aba="Resultado",
        )

        st.download_button(
            tr("Baixar resultado em Excel", 'Download results as Excel'),
            data=arquivo_excel,
            file_name=tr("resultado_produtos.xlsx", 'product_results.xlsx'),
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            key="download_resultado_produtos",
        )


# ==========================================================
# PÁGINA 2 - PLANEJAMENTO DE CORTES
# ==========================================================
def pagina_planejamento_cortes():
    st.title(tr("✂️ Planejamento de Cortes", '✂️ Cut Planning'))

    st.write(
        tr("Cole os códigos dos produtos e as respectivas quantidades. "
        "Cada linha de quantidade deve corresponder à mesma linha do código.", 'Paste the product codes and their quantities. Each quantity must be on the same line as its corresponding code.')
    )

    arquivo_estrutura = upload_planilha(
        tr("Relação de Estruturas para calcular os cortes", 'Bill of materials for cut calculation'), tr("Envie a Relação Simplificada de Estruturas, com a aba RELAÇÃO DE ESTRUTURAS. Não envie o DE-PARA neste campo.", 'Upload the bill of materials, with the RELAÇÃO DE ESTRUTURAS sheet. Use the product mapping upload on the product search page.'),
        "upload_estrutura_cortes", ["resultado_planejamento_cortes", "resultado_programacao_maquinas"])

    coluna_codigos, coluna_quantidades = st.columns(2)

    with coluna_codigos:
        texto_codigos = st.text_area(
            tr("Códigos dos produtos", 'Product codes'),
            height=300,
            placeholder=(
                "17448\n"
                "51270\n"
                "51332\n"
                "51530\n"
                "47976"
            ),
            key="texto_codigos_cortes",
        )

    with coluna_quantidades:
        texto_quantidades = st.text_area(
            tr("Quantidades", 'Quantities (whole numbers, no separators)'),
            height=300,
            placeholder=(
                "800\n"
                "320\n"
                "168\n"
                "164\n"
                "51"
            ),
            key="texto_quantidades_cortes",
            help=tr("Informe uma quantidade por linha.", "Enter one positive whole number per line using digits only, for example 1000. Omit thousands separators and decimal marks."),
        )

    if st.button(
        tr("Calcular cortes", 'Calculate cuts'),
        type="primary",
        key="botao_calcular_cortes",
        disabled=arquivo_estrutura is None,
    ):
        codigos = linhas_preenchidas(texto_codigos)
        quantidades = linhas_preenchidas(texto_quantidades)

        if not codigos:
            st.warning(tr("Cole pelo menos um código de produto.", 'Paste at least one product code.'))

        elif not quantidades:
            st.warning(tr("Informe as quantidades dos produtos.", 'Enter the product quantities.'))

        elif len(codigos) != len(quantidades):
            st.error(
                tr(f"Você informou {len(codigos)} códigos e "
                f"{len(quantidades)} quantidades. As duas listas precisam "
                "ter o mesmo número de linhas.", f'You entered {len(codigos)} codes and {len(quantidades)} quantities. Both lists must have the same number of lines.')
            )

        elif language == "en" and any(re.fullmatch(r"[0-9]+", quantidade) is None for quantidade in quantidades):
            st.error("Enter quantities using digits only, for example 1000. Remove thousands separators and decimal marks.")

        else:
            try:
                with st.spinner(tr("Calculando cortes...", 'Calculating cuts...')):
                    resultado = calcular_cortes_por_listas(
                        codigos=codigos,
                        quantidades=quantidades,
                        arquivo_estrutura=arquivo_estrutura,
                    )

                st.session_state["resultado_planejamento_cortes"] = resultado
                st.session_state.pop("resultado_programacao_maquinas", None)

            except Exception as erro:
                st.error(tr(f"Erro ao calcular os cortes: {erro}", f'Error calculating cuts: {localized_error(erro)}'))

    # ======================================================
    # RESULTADO DO PLANEJAMENTO
    # ======================================================
    if "resultado_planejamento_cortes" in st.session_state:
        resultado = st.session_state["resultado_planejamento_cortes"]
        df_cortes = resultado["cortes"]
        resumo = resultado["resumo"]
        nao_encontrados = resultado["nao_encontrados"]

        st.divider()
        st.subheader(tr("Cortes necessários", 'Required cuts'))

        col1, col2, col3, col4, col5 = st.columns(5)

        col1.metric(
            tr("Produtos informados", 'Products entered'),
            resumo["produtos_informados"],
        )

        col2.metric(
            tr("Produtos encontrados", 'Products found'),
            resumo["produtos_encontrados"],
        )

        col3.metric(
            tr("Não encontrados", 'Not found'),
            resumo["produtos_nao_encontrados"],
        )

        col4.metric(
            tr("Cortes diferentes", 'Distinct cuts'),
            resumo["codigos_de_corte_distintos"],
        )

        col5.metric(
            tr("Peças a cortar", 'Pieces to cut'),
            resumo["total_de_pecas_cortadas"],
        )

        if nao_encontrados:
            st.warning(
                tr("Não foi encontrada estrutura para: "
                + ", ".join(nao_encontrados), 'No bill of materials was found for: ' + ', '.join(nao_encontrados))
            )

        if df_cortes.empty:
            st.info(tr("Nenhum corte foi encontrado para os produtos informados.", 'No cuts were found for the products entered.'))

        else:
            st.dataframe(
                localized_frame(df_cortes),
                use_container_width=True,
                hide_index=True,
            )

            st.caption(
                tr("QUANT. = quantidade total daquele corte a produzir. "
                "QUANT = quantidade daquele corte utilizada em cada produto.", 'Total quantity = total number of pieces of this cut to produce. Quantity per product = number of pieces of this cut used in each product.')
            )

            st.divider()

            with st.expander(tr("Ver pedidos consolidados", 'View consolidated orders')):
                st.dataframe(
                    localized_frame(resultado["pedidos"]),
                    use_container_width=True,
                    hide_index=True,
                )

            arquivo_excel = gerar_excel(
                df_cortes,
                nome_aba="Cortes",
            )

            st.download_button(
                tr("Baixar planejamento de cortes em Excel", 'Download cut plan as Excel'),
                data=arquivo_excel,
                file_name=tr("planejamento_cortes.xlsx", 'cut_plan.xlsx'),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                key="download_planejamento_cortes",
            )



# ==========================================================
# PÁGINA 3 - PROGRAMAÇÃO DAS MÁQUINAS
# ==========================================================
def pagina_programacao_maquinas():
    st.title(tr("📅 Programação das Máquinas", '📅 Machine Scheduling'))

    st.write(tr("Use os cortes calculados no app ou envie um planejamento já salvo em Excel.", 'Use cuts calculated in this app or upload a saved Excel cut plan.'))
    fonte_cortes = st.radio(
        tr("Origem dos cortes", 'Cut source'), ["Cortes calculados nesta sessão", "Enviar planejamento em Excel"],
        index=0 if "resultado_planejamento_cortes" in st.session_state else 1,
        format_func={
            "Cortes calculados nesta sessão": tr("Cortes calculados nesta sessão", "Cuts calculated this session"),
            "Enviar planejamento em Excel": tr("Enviar planejamento em Excel", "Upload Excel cut plan"),
        }.__getitem__,
        key="origem_cortes")
    if fonte_cortes == "Enviar planejamento em Excel":
        arquivo_cortes = upload_planilha(
            tr("Planejamento de cortes", 'Cut plan'), tr("Envie o planejamento em Excel com a aba Cortes. Este campo não recebe a Relação de Estruturas nem o DE-PARA.", 'Upload an Excel cut plan with a Cortes or Cuts sheet, as exported by this app.'),
            "upload_planejamento", ["resultado_programacao_maquinas"])
        if arquivo_cortes is None:
            st.session_state.pop("resultado_programacao_maquinas", None)
            return
        try:
            with pd.ExcelFile(arquivo_cortes, engine="openpyxl") as planilha:
                aba_cortes = next((nome for nome in ("Cortes", "Cuts") if nome in planilha.sheet_names), None)
                if aba_cortes is None:
                    raise ValueError(tr("A planilha deve conter a aba Cortes ou Cuts.", "The workbook must contain a Cortes or Cuts sheet."))
                df_cortes = preparar_entrada(source_frame(pd.read_excel(planilha, sheet_name=aba_cortes)))
        except Exception as erro:
            st.session_state.pop("resultado_programacao_maquinas", None)
            st.error(tr(f"Não foi possível ler o planejamento: {erro}", f'Could not read the cut plan: {localized_error(erro)}'))
            return
    else:
        if "resultado_planejamento_cortes" not in st.session_state:
            st.session_state.pop("resultado_programacao_maquinas", None)
            st.info(tr("Abra Planejamento de cortes, envie a Relação de Estruturas e clique em Calcular cortes, ou escolha Enviar planejamento em Excel acima.", 'Open Cut planning, upload the bill of materials and click Calculate cuts, or select Upload Excel cut plan above.'))
            return
        df_cortes = st.session_state["resultado_planejamento_cortes"]["cortes"]

    if df_cortes.empty:
        st.info(tr("O planejamento atual não possui cortes para organizar.", 'The current plan has no cuts to schedule.'))
        return

    col1, col2 = st.columns(2)
    col1.metric(tr("Linhas de corte recebidas", 'Cut rows received'), len(df_cortes))

    if "QUANT." in df_cortes.columns:
        total_pecas = int(
            pd.to_numeric(
                df_cortes["QUANT."],
                errors="coerce",
            ).fillna(0).sum()
        )
        col2.metric(
            tr("Peças a organizar", 'Pieces to schedule'),
            format_integer(total_pecas),
        )
    else:
        col2.metric(tr("Peças a organizar", 'Pieces to schedule'), "-")

    st.divider()
    st.subheader(tr("Configuração dos turnos", 'Shift configuration'))

    quantidade_turnos = int(
        st.number_input(
            tr("Quantidade de turnos", 'Number of shifts'),
            min_value=1,
            max_value=6,
            value=1,
            step=1,
            key="quantidade_turnos_programacao",
        )
    )

    st.caption(
        tr("Máquinas gargalo trabalham sem pausa de almoço. As demais possuem pausa de 1 hora. "
        "A pausa começa no meio do horário informado. "
        "Exemplo: 07:00–17:00 gera pausa de 12:00–13:00. "
        "Se o término for menor que o início, o sistema entende que o turno termina no dia seguinte.", 'Bottleneck machines run through lunch. Other machines take a one-hour break starting at the midpoint of the shift. For example, 07:00–17:00 gives a 12:00–13:00 break. An end time before the start time means the shift ends the following day.')
    )

    turnos = []
    inicio_sugerido = time(7, 0)
    fim_sugerido = time(17, 0)

    for indice_turno in range(1, quantidade_turnos + 1):
        st.markdown(tr(f"**Turno {indice_turno}**", f'**Shift {indice_turno}**'))
        coluna_inicio, coluna_fim = st.columns(2)

        with coluna_inicio:
            inicio_turno = st.time_input(
                tr(f"Início do turno {indice_turno}", f'Shift start {indice_turno}'),
                value=inicio_sugerido,
                step=15 * 60,
                key=f"inicio_turno_{indice_turno}",
            )

        with coluna_fim:
            fim_turno = st.time_input(
                tr(f"Término do turno {indice_turno}", f'Shift end {indice_turno}'),
                value=fim_sugerido,
                step=15 * 60,
                key=f"fim_turno_{indice_turno}",
            )

        maquinas_turno = st.multiselect(
            tr(f"Máquinas que participarão do turno {indice_turno}", f'Machines participating in shift {indice_turno}'),
            options=MAQUINAS, default=list(MAQUINAS),
            key=f"maquinas_participantes_turno_{indice_turno}",
            help=tr("Remova as máquinas paradas ou indisponíveis neste turno. Elas não receberão cortes nesse período.", 'Remove machines that are stopped or unavailable for this shift. No cuts will be allocated to them during this period.'),
        )
        if not maquinas_turno:
            st.info(tr(f"Nenhuma máquina participará do turno {indice_turno}.", f'No machines will participate in shift {indice_turno}.'))

        turnos.append(
            {
                "nome": f"Turno {indice_turno}",
                "inicio": inicio_turno,
                "fim": fim_turno,
                "maquinas": maquinas_turno,
            }
        )

        # Para um novo turno, sugere continuidade a partir do fim do anterior.
        inicio_sugerido = fim_turno
        fim_sugerido = somar_horas_horario(fim_turno, 10)

    st.subheader(tr("Capacidade e metas das máquinas", 'Machine capacity and production targets'))
    gargalos = st.multiselect(
        tr("Máquinas gargalo — sem pausa de almoço", 'Bottleneck machines — no lunch break'), MAQUINAS,
        default=list(MAQUINAS_GARGALO_PADRAO), key="maquinas_gargalo")
    st.caption(
        tr("Informe as perdas previstas em minutos e a hora extra autorizada por máquina. "
        "As perdas reservam tempo no início da programação; a hora extra fica disponível "
        "após o último turno. Se a máquina não participar do último turno, deixe sua hora extra em zero. Para máquinas fora de todos os turnos, deixe também as perdas em zero. Setup já está embutido em H/MÁQ.", 'Enter expected downtime and authorized overtime in minutes for each machine. Downtime reserves time at the beginning of the schedule; overtime is available after the last shift. Set overtime to zero for machines absent from the last shift. Set downtime to zero for machines absent from all shifts. Setup time is already included in machine hours.'))
    capacidades = editor_persistente(
        pd.DataFrame({"MÁQUINA": MAQUINAS, "PERDAS (min)": [0.0] * len(MAQUINAS),
                      "HORA EXTRA (min)": [0.0] * len(MAQUINAS)}),
        disabled=["MÁQUINA"], hide_index=True, key="editor_capacidade",
        column_config={
            "MÁQUINA": column_label("MÁQUINA"),
            "PERDAS (min)": st.column_config.NumberColumn(column_label("PERDAS (min)"), min_value=0, required=True),
            "HORA EXTRA (min)": st.column_config.NumberColumn(column_label("HORA EXTRA (min)"), min_value=0, required=True)})

    metas_iniciais = metas_padrao()
    gerais = editor_persistente(
        pd.DataFrame([{"MÁQUINA": m, "META/H": metas_iniciais[m]["PADRAO"]} for m in MAQUINAS]),
        disabled=["MÁQUINA"], hide_index=True, key="editor_metas_gerais",
        column_config={
            "MÁQUINA": column_label("MÁQUINA"),
            "META/H": st.column_config.NumberColumn(column_label("META/H"), min_value=0.01, required=True)})
    st.caption(tr("Meta geral por máquina. A tabela abaixo substitui essa meta somente para o tubo informado.", 'General production target for each machine. The table below overrides this target only for the specified tube.'))
    excecoes = editor_persistente(
        pd.DataFrame([
            {"MÁQUINA": m, "TUBO": tubo, "META/H": taxa}
            for m, valores in metas_iniciais.items() for tubo, taxa in valores.items() if tubo != "PADRAO"
        ], columns=["MÁQUINA", "TUBO", "META/H"]),
        num_rows="dynamic", hide_index=True, key="editor_excecoes",
        column_config={
            "MÁQUINA": st.column_config.SelectboxColumn(column_label("MÁQUINA"), options=MAQUINAS, required=True),
            "TUBO": st.column_config.TextColumn(column_label("TUBO"), required=True),
            "META/H": st.column_config.NumberColumn(column_label("META/H"), min_value=0.01, required=True)})
    st.caption(tr("As metas iniciais de 20X48 = 200/h e 40X77 = 110/h se aplicam somente às SIMECs.", 'The initial targets of 20X48 = 200/hour and 40X77 = 110/hour apply only to SIMEC machines.'))

    configuracao_turnos_valida = True
    resumo_turnos_atual = pd.DataFrame()
    metas = {}
    extras = {}
    perdas = {}
    try:
        metas = {r["MÁQUINA"]: {"PADRAO": float(r["META/H"])} for _, r in gerais.iterrows()}
        # As exceções apagadas no app devem deixar de valer, inclusive as iniciais.
        for _, r in excecoes.iterrows():
            if pd.isna(r["MÁQUINA"]) or pd.isna(r["TUBO"]) or not str(r["TUBO"]).strip():
                raise ValueError("Preencha máquina, tubo e meta de cada exceção.")
            if r["MÁQUINA"] not in MAQUINAS:
                raise ValueError("Máquina inválida na tabela de exceções.")
            metas[r["MÁQUINA"]][familia_tubo(r["TUBO"])] = float(r["META/H"])
        extras = dict(zip(capacidades["MÁQUINA"], capacidades["HORA EXTRA (min)"]))
        perdas = dict(zip(capacidades["MÁQUINA"], capacidades["PERDAS (min)"]))
        metas, _, extras, perdas = validar_configuracao(metas, gargalos, extras, perdas)
        tabelas_turnos = []
        for m in MAQUINAS:
            tabela = resumir_turnos(turnos, sem_pausa=m in gargalos, hora_extra_min=extras[m], maquina=m)
            tabela.insert(0, "MÁQUINA", m)
            tabelas_turnos.append(tabela)
        resumo_turnos_atual = pd.concat(tabelas_turnos, ignore_index=True)
        # Mesma validação de capacidade usada no cálculo.
        from organizar_programacao_turnos import construir_linha_do_tempo_turnos
        for m in MAQUINAS:
            if perdas[m] > construir_linha_do_tempo_turnos(turnos, m in gargalos, extras[m], maquina=m)["minutos_produtivos"]:
                raise ValueError(f"Perdas de {m} excedem o tempo disponível.")
        st.write(tr("**Como os turnos serão considerados por máquina:**", '**Shift allocation for each machine:**'))
        st.dataframe(localized_frame(resumo_turnos_atual), use_container_width=True, hide_index=True)
    except Exception as erro_turnos:
        configuracao_turnos_valida = False
        st.error(tr(f"Configuração inválida: {erro_turnos}", f'Invalid configuration: {localized_error(erro_turnos)}'))

    assinatura_turnos = repr({
        "fonte_cortes": sha256(df_cortes.to_csv(index=False).encode()).hexdigest(),
        "origem": fonte_cortes,
        "turnos": [{"nome": t["nome"], "inicio": t["inicio"].isoformat(), "fim": t["fim"].isoformat(), "maquinas": sorted(t["maquinas"])} for t in turnos],
        "metas": metas, "gargalos": sorted(gargalos), "extras": extras, "perdas": perdas,
    })
    resultado_anterior = st.session_state.get("resultado_programacao_maquinas")
    if resultado_anterior is not None and (
        not configuracao_turnos_valida or resultado_anterior.get("assinatura_turnos") != assinatura_turnos
    ):
        st.session_state.pop("resultado_programacao_maquinas", None)
        st.info(tr("A configuração foi alterada. Clique novamente em 'Organizar programação das máquinas' para recalcular.", 'The configuration has changed. Click Schedule machines again to recalculate.'))

    st.divider()

    with st.expander(tr("Ver cortes que serão organizados", 'View cuts to be scheduled')):
        st.dataframe(
            localized_frame(df_cortes),
            use_container_width=True,
            hide_index=True,
        )

    if st.button(
        tr("Organizar programação das máquinas", 'Schedule machines'),
        type="primary",
        key="botao_organizar_programacao",
    ):
        if not configuracao_turnos_valida:
            st.error(tr("Corrija a configuração antes de organizar a programação.", 'Correct the configuration before scheduling machines.'))
            return

        try:
            with st.spinner(tr("Distribuindo os cortes entre as máquinas...", 'Allocating cuts to machines...')):
                # Reaproveita a função já existente no projeto.
                # O Excel é criado apenas em memória.
                arquivo_em_memoria = gerar_excel(
                    df_cortes,
                    nome_aba="Cortes",
                    localizar=False,
                )

                programacao, nao_alocados, resumo = organizar_programacao(
                    arquivo_em_memoria,
                    aba="Cortes",
                    turnos=turnos,
                    metas=metas, maquinas_gargalo=gargalos,
                    hora_extra_min=extras, perdas_min=perdas,
                )

                st.session_state["resultado_programacao_maquinas"] = {
                    "programacao": programacao,
                    "nao_alocados": nao_alocados,
                    "resumo": resumo,
                    "turnos_df": resumo_turnos_atual,
                    "assinatura_turnos": assinatura_turnos,
                }

        except Exception as erro:
            st.error(tr(f"Erro ao organizar a programação: {erro}", f'Error scheduling machines: {localized_error(erro)}'))

    if "resultado_programacao_maquinas" not in st.session_state:
        return

    resultado = st.session_state["resultado_programacao_maquinas"]
    programacao = resultado["programacao"]
    nao_alocados = resultado["nao_alocados"]
    resumo = resultado["resumo"]
    turnos_df = resultado.get("turnos_df", pd.DataFrame())

    st.divider()
    st.subheader(tr("Resumo da programação", 'Schedule summary'))

    pecas_programadas = 0
    if not programacao.empty and "QUANT. PROGRAMADA" in programacao.columns:
        pecas_programadas = int(
            pd.to_numeric(
                programacao["QUANT. PROGRAMADA"],
                errors="coerce",
            ).fillna(0).sum()
        )

    maquinas_utilizadas = 0
    if not resumo.empty and "LOTES" in resumo.columns:
        maquinas_utilizadas = int(
            (
                pd.to_numeric(
                    resumo["LOTES"],
                    errors="coerce",
                ).fillna(0)
                > 0
            ).sum()
        )

    pecas_nao_alocadas = 0
    if not nao_alocados.empty and "QUANT. RESTANTE" in nao_alocados.columns:
        pecas_nao_alocadas = int(
            pd.to_numeric(
                nao_alocados["QUANT. RESTANTE"],
                errors="coerce",
            ).fillna(0).sum()
        )

    m1, m2, m3 = st.columns(3)

    m1.metric(
        tr("Peças programadas", 'Scheduled pieces'),
        format_integer(pecas_programadas),
    )
    m2.metric(
        tr("Máquinas utilizadas", 'Machines used'),
        maquinas_utilizadas,
    )
    m3.metric(
        tr("Peças não alocadas", 'Unallocated pieces'),
        format_integer(pecas_nao_alocadas),
    )

    st.dataframe(
        localized_frame(resumo),
        use_container_width=True,
        hide_index=True,
    )

    if not turnos_df.empty:
        with st.expander(tr("Ver turnos usados nesta programação", 'View shifts used in this schedule')):
            st.dataframe(
                localized_frame(turnos_df),
                use_container_width=True,
                hide_index=True,
            )

    st.divider()
    st.subheader(tr("Programação por máquina", 'Schedule by machine'))

    if programacao.empty:
        st.info(tr("Nenhum lote foi programado.", 'No batches were scheduled.'))
    else:
        maquinas = resumo["MÁQUINA"].astype(str).tolist()
        abas = st.tabs(maquinas)

        for aba_maquina, maquina in zip(abas, maquinas):
            with aba_maquina:
                df_maquina = programacao[
                    programacao["MÁQUINA"].astype(str).eq(maquina)
                ].copy()

                if df_maquina.empty:
                    st.info(tr("Nenhum corte programado para esta máquina.", 'No cuts were scheduled for this machine.'))
                else:
                    colunas_exibir = [
                        coluna
                        for coluna in [
                            "TURNO",
                            "CÓD.",
                            "TUBO",
                            "CORTE",
                            "DIAG.",
                            "DESCRIÇÃO",
                            "DESCRIÇÃO DO PRODUTO",
                            "QUANT. PROGRAMADA",
                            "META/H",
                            "H/MÁQ",
                            "INÍCIO",
                            "TÉRMINO",
                            "QUANT. VARAS",
                            "REGRA",
                        ]
                        if coluna in df_maquina.columns
                    ]

                    st.dataframe(
                        localized_frame(df_maquina[colunas_exibir]),
                        use_container_width=True,
                        hide_index=True,
                    )

    st.divider()

    if nao_alocados.empty:
        st.success(
            tr("Todos os cortes foram alocados dentro dos turnos configurados.", 'All cuts were allocated within the configured shifts.')
        )
    else:
        st.warning(
            tr("Existem cortes que não couberam na programação ou não possuem "
            "máquina elegível pelas regras atuais.", 'Some cuts exceeded available capacity or have no eligible machine under the current rules.')
        )

        st.dataframe(
            localized_frame(nao_alocados),
            use_container_width=True,
            hide_index=True,
        )

    arquivo_final = gerar_excel_programacao(
        programacao,
        nao_alocados,
        resumo,
        turnos_df=turnos_df,
    )

    st.download_button(
        tr("Baixar programação das máquinas em Excel", 'Download machine schedule as Excel'),
        data=arquivo_final,
        file_name=tr("programacao_maquinas.xlsx", 'machine_schedule.xlsx'),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        key="download_programacao_maquinas",
    )



# ==========================================================
# NAVEGAÇÃO
# ==========================================================
st.sidebar.title(tr("Navegação", 'Navigation'))

pagina = st.sidebar.radio(
    tr("Página", 'Page'),
    [
        "📖 Instruções de uso",
        "🔎 Buscar códigos dos produtos",
        "✂️ Planejamento de cortes",
        "📅 Programação das máquinas",
    ],
    key="pagina",
    format_func={
        "📖 Instruções de uso": tr("📖 Instruções de uso", "📖 How to use"),
        "🔎 Buscar códigos dos produtos": tr("🔎 Buscar códigos dos produtos", "🔎 Find product codes"),
        "✂️ Planejamento de cortes": tr("✂️ Planejamento de cortes", "✂️ Cut planning"),
        "📅 Programação das máquinas": tr("📅 Programação das máquinas", "📅 Machine scheduling"),
    }.__getitem__,
)

if pagina == "📖 Instruções de uso":
    pagina_instrucoes()
elif pagina == "🔎 Buscar códigos dos produtos":
    pagina_buscar_codigos()
elif pagina == "✂️ Planejamento de cortes":
    pagina_planejamento_cortes()
else:
    pagina_programacao_maquinas()
