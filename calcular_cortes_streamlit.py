"""Motor de cálculo de cortes, preparado para ser usado em Streamlit.

Este arquivo NÃO cria a página do Streamlit.
Ele contém apenas a lógica que poderá ser importada pelo app depois.

Exemplo futuro no Streamlit:

    from calcular_cortes_streamlit import calcular_cortes_por_listas

    resultado = calcular_cortes_por_listas(
        codigos=["17448", "51270", "47976"],
        quantidades=[800, 320, 51],
    )

    st.dataframe(resultado["cortes"], use_container_width=True)

Também é possível passar um arquivo enviado pelo st.file_uploader em
arquivo_estrutura=arquivo_enviado.
"""

from collections import OrderedDict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata

import pandas as pd
from openpyxl import load_workbook


# ----------------------------------------------------------
# RELAÇÃO DE ESTRUTURAS
# ----------------------------------------------------------
ARQUIVO_ESTRUTURA = "RELAÇÃO SIMPLIFICADA DE ESTRUTURAS 2026.xlsm"
ABA_ESTRUTURA = "RELAÇÃO DE ESTRUTURAS"
COLUNA_PRODUTO_ESTRUTURA = "DESCRIÇÃO DO PRODUTO"
COLUNA_CODIGO_ESTRUTURA = "PRODUTO"

# Se os códigos numéricos da planilha tiverem 10 dígitos, manter este valor.
DIGITOS_CODIGO_CORTE = 10

COLUNAS_SAIDA = [
    "PRODUTO",
    "CÓD.",
    "TUBO",
    "VARA",
    "CORTE",
    "DIAG.",
    "DESCRIÇÃO",
    "DESCRIÇÃO DO PRODUTO",
    "QUANT.",
    "QUANT",
]


# ----------------------------------------------------------
# FUNÇÕES AUXILIARES
# ----------------------------------------------------------
def normalizar_cabecalho(valor):
    """Ignora acentos, espaços extras e quebras de linha no cabeçalho."""
    texto = " ".join(str(valor or "").upper().split())
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def normalizar_codigo_produto(valor):
    """Normaliza códigos de produto ignorando zeros à esquerda.

    Exemplos equivalentes:
        47891
        47891.0
        "47891"
        "0000047891"

    Todos são convertidos para "47891".
    """
    if valor is None:
        return ""

    texto = str(valor).strip()

    if not texto:
        return ""

    # Quando o Excel entrega um código numérico como 47891.0.
    if re.fullmatch(r"\d+\.0+", texto):
        texto = texto.split(".")[0]

    # Para códigos exclusivamente numéricos, remove zeros à esquerda.
    # Assim 0000047891 e 47891 passam a ser tratados como o mesmo produto.
    if texto.isdigit():
        return str(int(texto))

    return texto.upper()


def quantidade_inteira(valor, campo="Quantidade"):
    """Converte a quantidade para inteiro e valida o valor.

    Aceita, por exemplo:
      4
      4.0
      "4"
      "4,0"
      "2.000"
    """
    if valor is None:
        raise ValueError(f"{campo} está vazia.")

    texto = str(valor).strip().replace(" ", "")

    if not texto:
        raise ValueError(f"{campo} está vazia.")

    # Formato brasileiro com separador de milhar: 2.000 ou 2.000,00
    if re.fullmatch(r"\d{1,3}(?:\.\d{3})+(?:,\d+)?", texto):
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        numero = Decimal(texto)
    except InvalidOperation as exc:
        raise ValueError(f"{campo} inválida: {valor!r}") from exc

    if not numero.is_finite() or numero < 0:
        raise ValueError(f"{campo} deve ser um número não negativo: {valor!r}")

    if numero != numero.to_integral_value():
        raise ValueError(f"{campo} deve ser um número inteiro: {valor!r}")

    return int(numero)


def valor_para_texto(valor):
    if valor is None:
        return ""

    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))

    return str(valor).strip()


def formatar_codigo_corte(celula):
    """Mantém códigos de corte como texto e preserva zeros à esquerda."""
    valor = celula.value

    if valor is None:
        return ""

    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if int(valor) != valor:
            raise ValueError(f"Código de corte não inteiro: {valor}")
        return str(int(valor)).zfill(DIGITOS_CODIGO_CORTE)

    texto = str(valor).strip()
    return texto.zfill(DIGITOS_CODIGO_CORTE) if texto.isdigit() else texto


def localizar_cabecalho(aba):
    """Encontra automaticamente a linha do cabeçalho."""
    obrigatorias = {"PRODUTO", "COD.", "QUANT"}

    for numero_linha, linha in enumerate(
        aba.iter_rows(min_row=1, max_row=min(50, aba.max_row)), start=1
    ):
        colunas = {}

        for indice, celula in enumerate(linha):
            nome = normalizar_cabecalho(celula.value)
            if nome:
                colunas[nome] = indice

        if obrigatorias.issubset(colunas):
            return numero_linha, colunas

    raise ValueError(
        "Cabeçalho não encontrado. Verifique se a aba contém "
        "PRODUTO, CÓD. e QUANT."
    )


def consolidar_pedidos(pedidos):
    """Soma produtos repetidos mantendo a ordem da primeira ocorrência.

    Exemplo:
        [("47976", 51), ("47976", 1)]

    Resultado:
        {"47976": 52}
    """
    consolidados = OrderedDict()

    for numero, pedido in enumerate(pedidos, start=1):
        if len(pedido) != 2:
            raise ValueError(
                f"Pedido da posição {numero} deve conter código e quantidade."
            )

        codigo, quantidade = pedido
        codigo = normalizar_codigo_produto(codigo)

        if not codigo:
            raise ValueError(f"O código do produto na posição {numero} está vazio.")

        quantidade = quantidade_inteira(
            quantidade,
            f"Quantidade do produto {codigo}",
        )

        if quantidade == 0:
            raise ValueError(f"A quantidade do produto {codigo} deve ser maior que zero.")

        consolidados[codigo] = consolidados.get(codigo, 0) + quantidade

    if not consolidados:
        raise ValueError("Informe pelo menos um produto.")

    return consolidados


def _preparar_fonte_excel(arquivo_estrutura):
    """Prepara caminho local ou arquivo enviado pelo Streamlit."""
    if arquivo_estrutura is None:
        arquivo_estrutura = ARQUIVO_ESTRUTURA

    # Caminho local: planilha incluída junto do projeto/app.
    if isinstance(arquivo_estrutura, (str, Path)):
        arquivo = Path(arquivo_estrutura)

        if not arquivo.is_absolute():
            arquivo = Path(__file__).resolve().parent / arquivo

        if not arquivo.is_file():
            raise FileNotFoundError(
                f"Não encontrei a planilha: {arquivo}\n"
                "Coloque-a na mesma pasta deste arquivo Python ou informe outro caminho."
            )

        return arquivo

    # st.file_uploader retorna um objeto semelhante a arquivo.
    if hasattr(arquivo_estrutura, "seek"):
        arquivo_estrutura.seek(0)

    return arquivo_estrutura


# ----------------------------------------------------------
# CÁLCULO PRINCIPAL
# ----------------------------------------------------------
def calcular_cortes_lote(
    pedidos,
    arquivo_estrutura=None,
    aba_estrutura=ABA_ESTRUTURA,
):
    """Calcula os cortes necessários para vários produtos.

    Parâmetros
    ----------
    pedidos:
        Lista de pares (codigo_produto, quantidade).
        Exemplo:
            [("17448", 800), ("51270", 320)]

    arquivo_estrutura:
        Pode ser:
        - None: usa ARQUIVO_ESTRUTURA;
        - caminho str/Path;
        - arquivo vindo de st.file_uploader.

    aba_estrutura:
        Nome da aba da relação de estruturas.

    Retorno
    -------
    dict com:
        "cortes": DataFrame com os cortes calculados;
        "pedidos": DataFrame com os pedidos consolidados;
        "nao_encontrados": lista de produtos sem estrutura localizada;
        "resumo": dicionário com totais úteis para a futura tela.
    """
    pedidos_consolidados = consolidar_pedidos(pedidos)
    fonte = _preparar_fonte_excel(arquivo_estrutura)

    pasta = load_workbook(
        fonte,
        read_only=True,
        data_only=True,
    )

    try:
        if aba_estrutura not in pasta.sheetnames:
            raise ValueError(
                f"A aba '{aba_estrutura}' não existe. "
                f"Abas disponíveis: {pasta.sheetnames}"
            )

        aba = pasta[aba_estrutura]
        linha_cabecalho, colunas = localizar_cabecalho(aba)

        def obter(linha, nome_coluna):
            indice = colunas.get(normalizar_cabecalho(nome_coluna))

            if indice is None or indice >= len(linha):
                return None

            return linha[indice]

        # Mantém a saída na mesma ordem em que os produtos foram informados.
        cortes_por_produto = {
            codigo: []
            for codigo in pedidos_consolidados
        }

        ultimo_produto = ""

        for numero_linha, linha in enumerate(
            aba.iter_rows(min_row=linha_cabecalho + 1),
            start=linha_cabecalho + 1,
        ):
            celula_produto = obter(linha, COLUNA_CODIGO_ESTRUTURA)

            # Ajuda quando PRODUTO aparece apenas na primeira linha ou em célula mesclada.
            if celula_produto is not None and celula_produto.value not in (None, ""):
                ultimo_produto = normalizar_codigo_produto(celula_produto.value)

            if ultimo_produto not in pedidos_consolidados:
                continue

            celula_codigo_corte = obter(linha, "CÓD.")

            # Linha sem código de corte não entra no resultado.
            if (
                celula_codigo_corte is None
                or celula_codigo_corte.value in (None, "")
            ):
                continue

            celula_quant = obter(linha, "QUANT")

            if celula_quant is None or celula_quant.value in (None, ""):
                raise ValueError(
                    f"Linha {numero_linha}, produto {ultimo_produto}, "
                    f"corte {celula_codigo_corte.value}: QUANT está vazio."
                )

            quant_por_produto = quantidade_inteira(
                celula_quant.value,
                f"QUANT na linha {numero_linha}",
            )

            quantidade_pedido = pedidos_consolidados[ultimo_produto]
            quantidade_total_corte = quantidade_pedido * quant_por_produto

            def texto(nome_coluna):
                celula = obter(linha, nome_coluna)
                return valor_para_texto(celula.value) if celula is not None else ""

            cortes_por_produto[ultimo_produto].append(
                {
                    "PRODUTO": ultimo_produto,
                    "CÓD.": formatar_codigo_corte(celula_codigo_corte),
                    "TUBO": texto("TUBO"),
                    "VARA": texto("VARA"),
                    "CORTE": texto("CORTE"),
                    "DIAG.": texto("DIAG."),
                    "DESCRIÇÃO": texto("DESCRIÇÃO"),
                    "DESCRIÇÃO DO PRODUTO": texto(COLUNA_PRODUTO_ESTRUTURA),
                    # QUANT. = quantidade a fabricar do corte.
                    "QUANT.": quantidade_total_corte,
                    # QUANT = quantidade desse corte necessária por produto.
                    "QUANT": quant_por_produto,
                }
            )

        linhas_cortes = [
            corte
            for codigo in pedidos_consolidados
            for corte in cortes_por_produto[codigo]
        ]

        nao_encontrados = [
            codigo
            for codigo in pedidos_consolidados
            if not cortes_por_produto[codigo]
        ]

        df_cortes = pd.DataFrame(linhas_cortes, columns=COLUNAS_SAIDA)

        df_pedidos = pd.DataFrame(
            [
                {
                    "PRODUTO": codigo,
                    "QUANTIDADE": quantidade,
                }
                for codigo, quantidade in pedidos_consolidados.items()
            ]
        )

        resumo = {
            "produtos_informados": len(pedidos_consolidados),
            "produtos_encontrados": len(pedidos_consolidados) - len(nao_encontrados),
            "produtos_nao_encontrados": len(nao_encontrados),
            "linhas_de_corte": len(df_cortes),
            "codigos_de_corte_distintos": (
                int(df_cortes["CÓD."].nunique()) if not df_cortes.empty else 0
            ),
            "total_de_pecas_cortadas": (
                int(df_cortes["QUANT."].sum()) if not df_cortes.empty else 0
            ),
        }

        return {
            "cortes": df_cortes,
            "pedidos": df_pedidos,
            "nao_encontrados": nao_encontrados,
            "resumo": resumo,
        }

    finally:
        pasta.close()


def calcular_cortes_por_listas(
    codigos,
    quantidades,
    arquivo_estrutura=None,
    aba_estrutura=ABA_ESTRUTURA,
):
    """Atalho ideal para a futura tela do Streamlit.

    Recebe duas listas na mesma ordem:

        codigos = ["17448", "51270", "47976"]
        quantidades = [800, 320, 51]

    e devolve o mesmo dicionário de calcular_cortes_lote().
    """
    codigos = list(codigos)
    quantidades = list(quantidades)

    if len(codigos) != len(quantidades):
        raise ValueError(
            f"Foram informados {len(codigos)} códigos e "
            f"{len(quantidades)} quantidades. "
            "As listas precisam ter o mesmo número de itens."
        )

    pedidos = list(zip(codigos, quantidades))

    return calcular_cortes_lote(
        pedidos=pedidos,
        arquivo_estrutura=arquivo_estrutura,
        aba_estrutura=aba_estrutura,
    )


def calcular_cortes_dataframe(
    tabela_pedidos,
    coluna_codigo="PRODUTO",
    coluna_quantidade="QUANTIDADE",
    arquivo_estrutura=None,
    aba_estrutura=ABA_ESTRUTURA,
):
    """Recebe diretamente um DataFrame editado no Streamlit.

    Isto será útil se depois usarmos st.data_editor(), por exemplo:

        pedidos_editados = st.data_editor(df_pedidos)
        resultado = calcular_cortes_dataframe(pedidos_editados)
    """
    if not isinstance(tabela_pedidos, pd.DataFrame):
        tabela_pedidos = pd.DataFrame(tabela_pedidos)

    if coluna_codigo not in tabela_pedidos.columns:
        raise ValueError(
            f"A tabela de pedidos não possui a coluna '{coluna_codigo}'."
        )

    if coluna_quantidade not in tabela_pedidos.columns:
        raise ValueError(
            f"A tabela de pedidos não possui a coluna '{coluna_quantidade}'."
        )

    tabela = tabela_pedidos[[coluna_codigo, coluna_quantidade]].copy()

    # Remove somente linhas completamente vazias, comuns em st.data_editor.
    tabela = tabela.dropna(how="all")

    pedidos = list(
        tabela[[coluna_codigo, coluna_quantidade]].itertuples(
            index=False,
            name=None,
        )
    )

    return calcular_cortes_lote(
        pedidos=pedidos,
        arquivo_estrutura=arquivo_estrutura,
        aba_estrutura=aba_estrutura,
    )
