import pandas as pd
import unicodedata
import re
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles.colors import COLOR_INDEX


from io import BytesIO

# ==========================================================
# LOCALIZAÇÃO DOS ARQUIVOS
# ==========================================================

# Identifica a pasta em que está o código Python
BASE_DIR = Path(__file__).resolve().parent


# ==========================================================
# RELAÇÃO DE ESTRUTURAS
# ==========================================================

ARQUIVO_ESTRUTURA = (
    BASE_DIR / "RELAÇÃO SIMPLIFICADA DE ESTRUTURAS 2026.xlsm"
)

ABA_ESTRUTURA = "RELAÇÃO DE ESTRUTURAS"

COLUNA_PRODUTO_ESTRUTURA = "DESCRIÇÃO DO PRODUTO"

COLUNA_CODIGO_ESTRUTURA = "PRODUTO"


# ==========================================================
# PLANILHA DE-PARA
# ==========================================================

ARQUIVO_DE_PARA = (
    BASE_DIR / "product_mapping.xlsx"
)

ABA_DE_PARA = 0

COLUNA_PRODUTO_PROGRAMACAO = "DESCRIÇÃO PROGRAMAÇÃO"

COLUNA_PRODUTO_SIMPLIFICADO = "DESCRIÇÃO SIMPLIFICADA"

COLUNA_CODIGO_DE_PARA = "COD_PRODUTO"
# ----------------------------------------------------------
# PREFIXOS DE FAMÍLIA
# ----------------------------------------------------------

PREFIXOS_SEPARADOS = {
    "M",
    "MP",
    "P"
}


# ==========================================================
# VERIFICAR ARQUIVOS
# ==========================================================

def preparar_fonte(arquivo, nome):
    """Usa uma cópia em memória do upload; nenhuma base global é compartilhada."""
    if arquivo is None:
        raise ValueError(f"Envie a planilha {nome} na tela do aplicativo.")
    if hasattr(arquivo, "getvalue"):
        return BytesIO(arquivo.getvalue())
    if hasattr(arquivo, "read"):
        arquivo.seek(0)
        return BytesIO(arquivo.read())
    # Caminhos explícitos continuam disponíveis para chamadas externas ao app.
    caminho = Path(arquivo)
    if not caminho.is_file():
        raise FileNotFoundError(f"Planilha {nome} não encontrada.")
    return BytesIO(caminho.read_bytes())


# ==========================================================
# NORMALIZAR TEXTO
# ==========================================================

def normalizar(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto)

    # Espaço especial do Excel
    texto = texto.replace("\xa0", " ")

    # Diferentes tipos de hífen
    for hifen in [
        "–",
        "—",
        "−",
        "‐",
        "-"
    ]:
        texto = texto.replace(hifen, "-")

    # Diferentes tipos de barra
    texto = texto.replace("⁄", "/")
    texto = texto.replace("∕", "/")

    texto = texto.upper().strip()

    # Separa acentos
    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    # Remove acentos
    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    # Remove espaços duplicados
    texto = " ".join(
        texto.split()
    )

    return texto


# ==========================================================
# COMPACTAR TEXTO
# ==========================================================

def compactar(texto):
    texto = normalizar(texto)

    # Remove todos os caracteres que não sejam
    # letras ou números.
    #
    # MCPE-3 T/R -> MCPE3TR
    texto = re.sub(
        r"[^A-Z0-9]",
        "",
        texto
    )

    return texto


# ==========================================================
# NORMALIZAR CÓDIGO
# ==========================================================

def normalizar_codigo(codigo):
    if pd.isna(codigo):
        return ""

    codigo = str(codigo).strip()

    # Valores que não representam código
    if codigo.upper() in {
        "",
        "-",
        "NAN",
        "NONE"
    }:
        return ""

    # 56818.0 -> 56818
    if codigo.endswith(".0"):
        codigo = codigo[:-2]

    return codigo


# ==========================================================
# ENCONTRAR CABEÇALHO
# ==========================================================

def encontrar_cabecalho(
    arquivo,
    aba,
    colunas_procuradas
):
    bruto = pd.read_excel(
        arquivo,
        sheet_name=aba,
        engine="openpyxl",
        header=None
    )

    procuradas = [
        normalizar(coluna)
        for coluna in colunas_procuradas
    ]

    for i, linha in bruto.iterrows():
        valores = [
            normalizar(valor)
            for valor in linha.tolist()
        ]

        if all(
            coluna in valores
            for coluna in procuradas
        ):
            return i

    raise ValueError(
        "Cabeçalho não encontrado: "
        f"{colunas_procuradas}"
    )


# ==========================================================
# ENCONTRAR COLUNA NO EXCEL
# ==========================================================

def encontrar_coluna_excel(
    worksheet,
    linha_cabecalho,
    nome_coluna
):
    procurado = normalizar(
        nome_coluna
    )

    for celula in worksheet[
        linha_cabecalho
    ]:
        if normalizar(
            celula.value
        ) == procurado:
            return celula.column

    raise ValueError(
        f"Coluna '{nome_coluna}' não encontrada."
    )


# ==========================================================
# VERIFICAR SE A CÉLULA ESTÁ VERMELHA
# ==========================================================

def celula_vermelha(celula):
    cor = celula.font.color

    if cor is None:
        return False

    rgb = None

    # RGB comum
    if (
        cor.type == "rgb"
        and cor.rgb
    ):
        rgb = str(
            cor.rgb
        )[-6:]

    # Cor indexada
    elif (
        cor.type == "indexed"
        and cor.indexed is not None
    ):
        try:
            rgb = COLOR_INDEX[
                cor.indexed
            ][-6:]
        except Exception:
            return False

    if rgb is None:
        return False

    rgb = rgb.upper()

    try:
        vermelho = int(
            rgb[0:2],
            16
        )

        verde = int(
            rgb[2:4],
            16
        )

        azul = int(
            rgb[4:6],
            16
        )

    except Exception:
        return False

    return (
        vermelho >= 180
        and verde <= 100
        and azul <= 100
    )


# ==========================================================
# CÓDIGO ÚNICO
# ==========================================================

def obter_codigo_unico(
    dataframe,
    coluna_codigo
):
    if dataframe.empty:
        return None

    codigos = (
        dataframe[
            coluna_codigo
        ]
        .apply(normalizar_codigo)
        .tolist()
    )

    codigos = [
        codigo
        for codigo in codigos
        if codigo != ""
    ]

    # Remove repetidos mantendo ordem
    codigos = list(
        dict.fromkeys(codigos)
    )

    if len(codigos) == 1:
        return codigos[0]

    return None


# ==========================================================
# SEPARAR TOKENS
# ==========================================================

def separar_tokens(texto):
    texto = normalizar(texto)

    texto = texto.replace(",", ".")

    return re.findall(
        r"[A-Z]+|\d+(?:\.\d+)?",
        texto
    )


# ==========================================================
# EXTRAIR FAMÍLIA
# ==========================================================

def extrair_chaves_familia(texto):
    tokens = separar_tokens(
        texto
    )

    if not tokens:
        return set()

    chaves = set()

    indice = 0

    # Ignorar MP, M, P quando estiverem separados
    while (
        indice < len(tokens)
        and tokens[indice] in PREFIXOS_SEPARADOS
    ):
        indice += 1

    if indice >= len(tokens):
        return set()

    familia = tokens[indice]

    if not familia.isalpha():
        return set()

    numero = None

    for token in tokens[
        indice + 1:
    ]:
        if re.fullmatch(
            r"\d+(?:\.\d+)?",
            token
        ):
            numero = token
            break

    if numero is None:
        return set()

    # Família original
    chaves.add(
        familia + numero
    )

    # MCPE -> CPE
    if (
        familia.startswith("M")
        and len(familia) >= 4
    ):
        chaves.add(
            familia[1:]
            + numero
        )

    # MPCPE -> CPE
    if (
        familia.startswith("MP")
        and len(familia) >= 4
    ):
        chaves.add(
            familia[2:]
            + numero
        )

    return chaves


# ==========================================================
# MESMA FAMÍLIA
# ==========================================================

def mesma_familia(
    produto_1,
    produto_2
):
    familia_1 = extrair_chaves_familia(
        produto_1
    )

    familia_2 = extrair_chaves_familia(
        produto_2
    )

    if (
        not familia_1
        or not familia_2
    ):
        return False

    return not familia_1.isdisjoint(
        familia_2
    )


# ==========================================================
# EXTRAIR CARACTERÍSTICAS
# ==========================================================

def extrair_caracteristicas(texto):
    texto = normalizar(texto)

    texto = texto.replace(
        ",",
        "."
    )

    caracteristicas = set()

    # T/R
    if re.search(
        r"\bT\s*/\s*R\b",
        texto
    ):
        caracteristicas.add("TR")

    # DIREITA / ESQUERDA
    if re.search(
        r"\bDIR\b",
        texto
    ):
        caracteristicas.add("DIR")

    if re.search(
        r"\bESQ\b",
        texto
    ):
        caracteristicas.add("ESQ")

    # TAMPO
    if "TAMPO" in texto:
        caracteristicas.add(
            "TAMPO"
        )

    # ENDLESS
    if "ENDLESS" in texto:
        caracteristicas.add(
            "ENDLESS"
        )

    # MEDIDAS
    medidas = re.findall(
        r"\b\d+(?:\.\d+)?\s*M\b",
        texto
    )

    for medida in medidas:
        caracteristicas.add(
            medida.replace(
                " ",
                ""
            )
        )

    # LUGARES
    lugares = re.findall(
        r"\b\d+\s*LUG\b",
        texto
    )

    for lugar in lugares:
        caracteristicas.add(
            lugar.replace(
                " ",
                ""
            )
        )

    # POSIÇÕES
    posicoes = re.findall(
        r"\b\d+\s*P\b",
        texto
    )

    for posicao in posicoes:
        caracteristicas.add(
            posicao.replace(
                " ",
                ""
            )
        )

    # SEM BRAÇO
    if re.search(
        r"S\s*/\s*BRACO",
        texto
    ):
        caracteristicas.add(
            "SEM_BRACO"
        )

    # OUTRAS PALAVRAS
    palavras = [
        "ABS",
        "GR",
        "GD",
        "EMP",
        "ASSENTO",
        "ENCOSTO"
    ]

    for palavra in palavras:
        if palavra in texto:
            caracteristicas.add(
                palavra
            )

    return caracteristicas


# ==========================================================
# DESEMPATAR FAMÍLIA
# ==========================================================

def desempatar_familia(
    candidatos,
    produto_op,
    coluna_nome
):
    if candidatos.empty:
        return candidatos

    caracteristicas_op = (
        extrair_caracteristicas(
            produto_op
        )
    )

    if not caracteristicas_op:
        return candidatos

    candidatos = candidatos.copy()

    def calcular_score(nome):
        caracteristicas_candidato = (
            extrair_caracteristicas(
                nome
            )
        )

        acertos = len(
            caracteristicas_op
            &
            caracteristicas_candidato
        )

        faltantes = len(
            caracteristicas_op
            -
            caracteristicas_candidato
        )

        extras = len(
            caracteristicas_candidato
            -
            caracteristicas_op
        )

        return (
            acertos * 10
            - faltantes * 5
            - extras * 2
        )

    candidatos[
        "SCORE_DESEMPATE"
    ] = candidatos[
        coluna_nome
    ].apply(
        calcular_score
    )

    melhor_score = candidatos[
        "SCORE_DESEMPATE"
    ].max()

    if (
        pd.isna(melhor_score)
        or melhor_score <= 0
    ):
        return candidatos

    return candidatos[
        candidatos[
            "SCORE_DESEMPATE"
        ] == melhor_score
    ]


# ==========================================================
# BUSCA EXATA + EXATA COMPACTA
# ==========================================================

def buscar_exato(
    dataframe,
    produto,
    coluna_normalizada,
    coluna_compacta
):
    produto_normalizado = normalizar(
        produto
    )

    produto_compacto = compactar(
        produto
    )

    # EXATO NORMAL
    encontrados = dataframe[
        dataframe[
            coluna_normalizada
        ] == produto_normalizado
    ]

    if not encontrados.empty:
        return (
            encontrados,
            "exato"
        )

    # EXATO COMPACTO
    encontrados = dataframe[
        dataframe[
            coluna_compacta
        ] == produto_compacto
    ]

    if not encontrados.empty:
        return (
            encontrados,
            "exato compacto"
        )

    return (
        dataframe.iloc[0:0],
        None
    )


# ==========================================================
# BUSCA INCOMPLETA
# ==========================================================

def buscar_incompleto(
    dataframe,
    produto,
    coluna_nome,
    coluna_compacta
):
    consulta = compactar(
        produto
    )

    if consulta == "":
        return (
            dataframe.iloc[0:0],
            None
        )

    # PARCIAL LITERAL
    encontrados = dataframe[
        dataframe[
            coluna_compacta
        ].apply(
            lambda nome:
            consulta in nome
            if nome != ""
            else False
        )
    ]

    if not encontrados.empty:
        return (
            encontrados,
            "parcial literal"
        )

    # FAMÍLIA
    encontrados = dataframe[
        dataframe[
            coluna_nome
        ].apply(
            lambda nome:
            mesma_familia(
                produto,
                nome
            )
            if not pd.isna(nome)
            else False
        )
    ]

    if not encontrados.empty:
        encontrados = (
            desempatar_familia(
                encontrados,
                produto,
                coluna_nome
            )
        )

        return (
            encontrados,
            "família + características"
        )

    return (
        dataframe.iloc[0:0],
        None
    )


# ==========================================================
# CARREGAR E PREPARAR AS BASES
# ==========================================================

def carregar_bases(arquivo_estrutura, arquivo_de_para):
    arquivo_estrutura = preparar_fonte(arquivo_estrutura, "Relação de Estruturas")
    arquivo_de_para = preparar_fonte(arquivo_de_para, "DE-PARA")

    # ------------------------------------------------------
    # RELAÇÃO DE ESTRUTURAS
    # ------------------------------------------------------

    linha_cabecalho_estrutura = (
        encontrar_cabecalho(
            arquivo_estrutura,
            ABA_ESTRUTURA,
            [
                COLUNA_PRODUTO_ESTRUTURA,
                COLUNA_CODIGO_ESTRUTURA
            ]
        )
    )

    estrutura = pd.read_excel(
        arquivo_estrutura,
        sheet_name=ABA_ESTRUTURA,
        engine="openpyxl",
        header=linha_cabecalho_estrutura
    )

    estrutura.columns = (
        estrutura.columns
        .astype(str)
        .str.strip()
    )

    # ------------------------------------------------------
    # IDENTIFICAR LINHAS VERMELHAS
    # ------------------------------------------------------

    workbook = load_workbook(
        arquivo_estrutura,
        data_only=False,
        keep_vba=False
    )

    worksheet = workbook[
        ABA_ESTRUTURA
    ]

    linha_cabecalho_excel = (
        linha_cabecalho_estrutura
        + 1
    )

    coluna_produto_excel = (
        encontrar_coluna_excel(
            worksheet,
            linha_cabecalho_excel,
            COLUNA_PRODUTO_ESTRUTURA
        )
    )

    descontinuados = []

    for indice in estrutura.index:
        linha_excel = (
            linha_cabecalho_estrutura
            + 2
            + indice
        )

        celula = worksheet.cell(
            row=linha_excel,
            column=coluna_produto_excel
        )

        descontinuados.append(
            celula_vermelha(
                celula
            )
        )

    workbook.close()

    estrutura[
        "DESCONTINUADO"
    ] = descontinuados

    # ------------------------------------------------------
    # PREPARAR ESTRUTURA
    # ------------------------------------------------------

    estrutura = estrutura[
        [
            COLUNA_PRODUTO_ESTRUTURA,
            COLUNA_CODIGO_ESTRUTURA,
            "DESCONTINUADO"
        ]
    ].copy()

    estrutura = estrutura.dropna(
        subset=[
            COLUNA_PRODUTO_ESTRUTURA,
            COLUNA_CODIGO_ESTRUTURA
        ]
    )

    estrutura[
        "PRODUTO_NORMALIZADO"
    ] = estrutura[
        COLUNA_PRODUTO_ESTRUTURA
    ].apply(
        normalizar
    )

    estrutura[
        "PRODUTO_COMPACTO"
    ] = estrutura[
        COLUNA_PRODUTO_ESTRUTURA
    ].apply(
        compactar
    )

    estrutura[
        "CODIGO_NORMALIZADO"
    ] = estrutura[
        COLUNA_CODIGO_ESTRUTURA
    ].apply(
        normalizar_codigo
    )

    # Somente estruturas ativas
    estrutura_ativa = estrutura[
        estrutura[
            "DESCONTINUADO"
        ] == False
    ].copy()

    estrutura_ativa = (
        estrutura_ativa
        .drop_duplicates(
            subset=[
                "PRODUTO_NORMALIZADO",
                "CODIGO_NORMALIZADO"
            ]
        )
    )

    # ------------------------------------------------------
    # DE-PARA
    # ------------------------------------------------------

    linha_cabecalho_de_para = (
        encontrar_cabecalho(
            arquivo_de_para,
            ABA_DE_PARA,
            [
                COLUNA_PRODUTO_PROGRAMACAO,
                COLUNA_PRODUTO_SIMPLIFICADO,
                COLUNA_CODIGO_DE_PARA
            ]
        )
    )

    de_para = pd.read_excel(
        arquivo_de_para,
        sheet_name=ABA_DE_PARA,
        engine="openpyxl",
        header=linha_cabecalho_de_para
    )

    de_para.columns = (
        de_para.columns
        .astype(str)
        .str.strip()
    )

    de_para[
        "PROGRAMACAO_NORMALIZADA"
    ] = de_para[
        COLUNA_PRODUTO_PROGRAMACAO
    ].apply(
        normalizar
    )

    de_para[
        "PROGRAMACAO_COMPACTA"
    ] = de_para[
        COLUNA_PRODUTO_PROGRAMACAO
    ].apply(
        compactar
    )

    de_para[
        "SIMPLIFICADA_NORMALIZADA"
    ] = de_para[
        COLUNA_PRODUTO_SIMPLIFICADO
    ].apply(
        normalizar
    )

    de_para[
        "SIMPLIFICADA_COMPACTA"
    ] = de_para[
        COLUNA_PRODUTO_SIMPLIFICADO
    ].apply(
        compactar
    )

    de_para[
        "CODIGO_NORMALIZADO"
    ] = de_para[
        COLUNA_CODIGO_DE_PARA
    ].apply(
        normalizar_codigo
    )

    return estrutura_ativa, de_para


# As bases são carregadas por chamada, com os arquivos da sessão do usuário.


# ==========================================================
# VALIDAR CÓDIGO NA ESTRUTURA
# ==========================================================

def validar_codigo_estrutura(
    codigo, estrutura_ativa
):
    encontrado = estrutura_ativa[
        estrutura_ativa[
            "CODIGO_NORMALIZADO"
        ] == codigo
    ]

    if encontrado.empty:
        return None

    return encontrado.iloc[0]


# ==========================================================
# VALIDAR PELO NOME SIMPLIFICADO
# ==========================================================

def validar_por_simplificada(
    candidatos_de_para, estrutura_ativa
):
    encontrados_estrutura = []

    for _, linha in candidatos_de_para.iterrows():
        descricao = linha[
            COLUNA_PRODUTO_SIMPLIFICADO
        ]

        if pd.isna(descricao):
            continue

        # Primeiro exato / compacto
        candidatos, _ = buscar_exato(
            estrutura_ativa,
            descricao,
            "PRODUTO_NORMALIZADO",
            "PRODUTO_COMPACTO"
        )

        # Se não encontrou, parcial
        if candidatos.empty:
            candidatos, _ = buscar_incompleto(
                estrutura_ativa,
                descricao,
                COLUNA_PRODUTO_ESTRUTURA,
                "PRODUTO_COMPACTO"
            )

        if not candidatos.empty:
            encontrados_estrutura.append(
                candidatos
            )

    if not encontrados_estrutura:
        return None

    todos = pd.concat(
        encontrados_estrutura,
        ignore_index=True
    )

    codigo = obter_codigo_unico(
        todos,
        "CODIGO_NORMALIZADO"
    )

    if codigo is None:
        return None

    return todos[
        todos[
            "CODIGO_NORMALIZADO"
        ] == codigo
    ].iloc[0]


# ==========================================================
# RESOLVER RESULTADO DO DE-PARA
# ==========================================================

def resolver_de_para(
    candidatos, estrutura_ativa
):
    if candidatos.empty:
        return None

    codigo = obter_codigo_unico(
        candidatos,
        "CODIGO_NORMALIZADO"
    )

    # Se há um único código
    if codigo is not None:
        # Tenta primeiro validar diretamente pelo código
        linha = validar_codigo_estrutura(
            codigo, estrutura_ativa
        )

        if linha is not None:
            return {
                "codigo": codigo,
                "linha": linha,
                "validacao": "código"
            }

        # Se o código não existe entre estruturas ativas,
        # tenta pela descrição simplificada
        linha = validar_por_simplificada(
            candidatos, estrutura_ativa
        )

        if linha is not None:
            return {
                "codigo":
                    linha[
                        "CODIGO_NORMALIZADO"
                    ],

                "linha":
                    linha,

                "validacao":
                    "descrição simplificada"
            }

    return None


# ==========================================================
# FUNÇÃO PRINCIPAL PARA O STREAMLIT
# ==========================================================

def buscar_produtos(lista_produtos, arquivo_estrutura=None, arquivo_de_para=None):
    """
    Recebe uma lista de nomes da Ordem de Produção
    e devolve um DataFrame com:

    - PRODUTO OP
    - PRODUTO ENCONTRADO
    - CÓDIGO
    - MÉTODO
    - STATUS

    Esta função não usa input(), print() nem clipboard.
    Ela foi preparada para ser chamada pelo app.py.
    """

    estrutura_ativa, de_para = carregar_bases(arquivo_estrutura, arquivo_de_para)
    resultados = []

    for produto_op in lista_produtos:
        produto_op = str(produto_op).strip()

        # Ignora itens vazios
        if produto_op == "":
            continue

        # ==================================================
        # 1. ESTRUTURA EXATA / EXATA COMPACTA
        # ==================================================

        candidatos, metodo = buscar_exato(
            estrutura_ativa,
            produto_op,
            "PRODUTO_NORMALIZADO",
            "PRODUTO_COMPACTO"
        )

        codigo = obter_codigo_unico(
            candidatos,
            "CODIGO_NORMALIZADO"
        )

        if codigo is not None:
            linha = candidatos[
                candidatos[
                    "CODIGO_NORMALIZADO"
                ] == codigo
            ].iloc[0]

            resultados.append({
                "PRODUTO OP":
                    produto_op,

                "PRODUTO ENCONTRADO":
                    linha[
                        COLUNA_PRODUTO_ESTRUTURA
                    ],

                "CÓDIGO":
                    codigo,

                "MÉTODO":
                    f"Estrutura - {metodo}",

                "STATUS":
                    "Encontrado"
            })

            continue

        # ==================================================
        # 2. DE-PARA EXATO / EXATO COMPACTO
        # ==================================================

        candidatos_de_para, metodo = buscar_exato(
            de_para,
            produto_op,
            "PROGRAMACAO_NORMALIZADA",
            "PROGRAMACAO_COMPACTA"
        )

        if not candidatos_de_para.empty:
            resolvido = resolver_de_para(
                candidatos_de_para, estrutura_ativa
            )

            if resolvido is not None:
                linha = resolvido[
                    "linha"
                ]

                resultados.append({
                    "PRODUTO OP":
                        produto_op,

                    "PRODUTO ENCONTRADO":
                        linha[
                            COLUNA_PRODUTO_ESTRUTURA
                        ],

                    "CÓDIGO":
                        resolvido[
                            "codigo"
                        ],

                    "MÉTODO":
                        f"DE-PARA - {metodo}",

                    "STATUS":
                        (
                            "Encontrado e validado por "
                            + resolvido[
                                "validacao"
                            ]
                        )
                })

                continue

        # ==================================================
        # 3. DE-PARA INCOMPLETO
        # ==================================================

        candidatos_de_para, metodo = (
            buscar_incompleto(
                de_para,
                produto_op,
                COLUNA_PRODUTO_PROGRAMACAO,
                "PROGRAMACAO_COMPACTA"
            )
        )

        if not candidatos_de_para.empty:
            resolvido = resolver_de_para(
                candidatos_de_para, estrutura_ativa
            )

            if resolvido is not None:
                linha = resolvido[
                    "linha"
                ]

                resultados.append({
                    "PRODUTO OP":
                        produto_op,

                    "PRODUTO ENCONTRADO":
                        linha[
                            COLUNA_PRODUTO_ESTRUTURA
                        ],

                    "CÓDIGO":
                        resolvido[
                            "codigo"
                        ],

                    "MÉTODO":
                        f"DE-PARA - {metodo}",

                    "STATUS":
                        (
                            "Encontrado e validado por "
                            + resolvido[
                                "validacao"
                            ]
                        )
                })

                continue

            # DE-PARA encontrou várias possibilidades
            codigos_possiveis = (
                candidatos_de_para[
                    "CODIGO_NORMALIZADO"
                ]
                .apply(normalizar_codigo)
                .unique()
                .tolist()
            )

            codigos_possiveis = [
                codigo
                for codigo in codigos_possiveis
                if codigo != ""
            ]

            resultados.append({
                "PRODUTO OP":
                    produto_op,

                "PRODUTO ENCONTRADO":
                    "Mais de uma possibilidade",

                "CÓDIGO":
                    "Verificar",

                "MÉTODO":
                    f"DE-PARA - {metodo}",

                "STATUS":
                    (
                        "Possíveis códigos: "
                        + ", ".join(
                            codigos_possiveis
                        )
                    )
            })

            continue

        # ==================================================
        # 4. ESTRUTURA INCOMPLETA
        # ==================================================

        candidatos, metodo = buscar_incompleto(
            estrutura_ativa,
            produto_op,
            COLUNA_PRODUTO_ESTRUTURA,
            "PRODUTO_COMPACTO"
        )

        codigo = obter_codigo_unico(
            candidatos,
            "CODIGO_NORMALIZADO"
        )

        if codigo is not None:
            linha = candidatos[
                candidatos[
                    "CODIGO_NORMALIZADO"
                ] == codigo
            ].iloc[0]

            resultados.append({
                "PRODUTO OP":
                    produto_op,

                "PRODUTO ENCONTRADO":
                    linha[
                        COLUNA_PRODUTO_ESTRUTURA
                    ],

                "CÓDIGO":
                    codigo,

                "MÉTODO":
                    f"Estrutura - {metodo}",

                "STATUS":
                    "Encontrado"
            })

            continue

        # ==================================================
        # 5. NÃO ENCONTRADO
        # ==================================================

        resultados.append({
            "PRODUTO OP":
                produto_op,

            "PRODUTO ENCONTRADO":
                "Não encontrado",

            "CÓDIGO":
                "Não encontrado",

            "MÉTODO":
                "Não encontrado",

            "STATUS":
                "Não encontrado"
        })

    return pd.DataFrame(
        resultados,
        columns=[
            "PRODUTO OP",
            "PRODUTO ENCONTRADO",
            "CÓDIGO",
            "MÉTODO",
            "STATUS"
        ]
    )
