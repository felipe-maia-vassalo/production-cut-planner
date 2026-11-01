from __future__ import annotations

import math
import os
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd


# ============================================================
# CONFIGURAÇÕES
# ============================================================

MAQUINAS = [
    "SIMEC 1",
    "SIMEC 2",
    "EMT",
    "BLM",
    "TECNOMAQ",
    "SERRA 1",
    "SERRA 2",
]

# Máquinas que não são gargalo têm uma pausa de 1 hora por turno.
# O horário da pausa é calculado automaticamente a partir do horário do turno.
# Exemplo: 07:00 -> 17:00 gera pausa de 12:00 -> 13:00.
PAUSA_TURNO_MIN = 60

# Configuração usada quando a função for chamada sem informar turnos.
TURNOS_PADRAO = [
    {"nome": "Turno 1", "inicio": "07:00", "fim": "17:00"},
]

# Setup já está embutido em H/MÁQ; não há intervalo adicional entre lotes.
META_H_PADRAO = 320
META_H_POR_TUBO = {"20X48": 200, "40X77": 110}  # somente SIMEC
MAQUINAS_GARGALO_PADRAO = ()


def metas_padrao() -> dict:
    """Cada máquina tem uma meta geral e exceções por família de tubo."""
    return {
        m: {"PADRAO": META_H_PADRAO,
            **(META_H_POR_TUBO if m.startswith("SIMEC") else {})}
        for m in MAQUINAS
    }


def validar_configuracao(metas=None, maquinas_gargalo=None,
                         hora_extra_min=None, perdas_min=None):
    configuracao = metas_padrao()
    for maquina, valores in (metas or {}).items():
        if maquina not in MAQUINAS:
            raise ValueError(f"Máquina desconhecida: {maquina}")
        # A configuração informada da máquina substitui suas exceções iniciais.
        # Assim, excluir uma exceção no app também a exclui no cálculo e no JSON.
        configuracao[maquina] = {"PADRAO": META_H_PADRAO}
        for tubo, valor in valores.items():
            taxa = float(valor)
            if not math.isfinite(taxa) or taxa <= 0:
                raise ValueError(f"Meta de {maquina}/{tubo} deve ser positiva e finita.")
            chave = "PADRAO" if normalizar_texto(tubo) == "PADRAO" else familia_tubo(tubo)
            configuracao[maquina][chave] = taxa
    gargalos = set(MAQUINAS_GARGALO_PADRAO if maquinas_gargalo is None else maquinas_gargalo)
    if gargalos - set(MAQUINAS):
        raise ValueError("Há máquinas gargalo desconhecidas.")
    valores_normalizados = []
    for nome, valores in (("Hora extra", hora_extra_min), ("Perdas", perdas_min)):
        resultado = {m: 0.0 for m in MAQUINAS}
        for maquina, valor in (valores or {}).items():
            if maquina not in MAQUINAS:
                raise ValueError(f"Máquina desconhecida: {maquina}")
            minutos = float(valor)
            if not math.isfinite(minutos) or minutos < 0:
                raise ValueError(f"{nome} de {maquina} deve ser não negativa e finita.")
            resultado[maquina] = minutos
        valores_normalizados.append(resultado)
    return configuracao, gargalos, *valores_normalizados

# Regras de tubo por máquina.
TUBOS_EMT = {
    "3/4",
    "7/8",
    "20X48",
    "29X58",
    "1.1/2",
    "20X20",
    "20X30",
    "20X40",
    "25X25",
}

# 7/8 liberado para BLM.
TUBOS_BLM = {
    "7/8",
    "16X30",
    "5/8",
    "20X20",
    "20X30",
    "20X40",
    "25X25",
}

TUBOS_TECNOMAQ = {
    "20X48",
    "16X30",
    "25X25",
    "20X20",
}

TUBOS_SERRA = {
    "BARRA CHATA",
    "30X50",
    "GONZO",
    "CHAPA",
    "25X50",
    "40X80",
    "40X77",
}

COLUNAS_OBRIGATORIAS = [
    "PRODUTO",
    "CÓD.",
    "TUBO",
    "VARA",
    "CORTE",
    "DIAG.",
    "DESCRIÇÃO",
    "DESCRIÇÃO DO PRODUTO",
    "QUANT.",
]



# Pasta padrão do projeto.
# Para testar em outra pasta sem editar o código, defina a variável de ambiente
# CORTE_PRODUCAO_DIR antes de executar.
PASTA_PROJETO = Path(
    os.getenv(
        "CORTE_PRODUCAO_DIR",
        str(Path(__file__).resolve().parent),
    )
)


def localizar_arquivo_planejamento(pasta: str | Path) -> Path:
    """
    Localiza automaticamente o arquivo de cortes dentro da pasta do projeto.

    Procura arquivos como:
      - planejamento_cortes.xlsx
      - planejamento_cortes (1).xlsx
      - planejamento_cortes (1)(1).xlsx

    Se houver mais de um, usa o modificado mais recentemente.
    Arquivos temporários do Excel (~$...) e a saída da programação são ignorados.
    """
    pasta = Path(pasta).expanduser()

    if not pasta.exists():
        raise FileNotFoundError(
            f"A pasta do projeto não foi encontrada: {pasta}\n"
            "Confira se o caminho está correto."
        )

    candidatos = [
        arq
        for arq in pasta.glob("planejamento_cortes*.xlsx")
        if arq.is_file()
        and not arq.name.startswith("~$")
        and "programacao_maquinas" not in arq.name.lower()
    ]

    if not candidatos:
        raise FileNotFoundError(
            "Não encontrei nenhum arquivo 'planejamento_cortes*.xlsx' em:\n"
            f"{pasta}"
        )

    return max(candidatos, key=lambda arq: arq.stat().st_mtime)

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    texto = texto.replace("×", "X").replace(" ", "").replace(",", ".")
    return texto


def para_float(valor, padrao=0.0) -> float:
    if pd.isna(valor):
        return float(padrao)
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return float(padrao)


def familia_tubo(tubo) -> str:
    """Reduz 20X30X1,2 -> 20X30; 7/8X1,2 -> 7/8 etc."""
    s = normalizar_texto(tubo)

    retangular = re.match(r"(\d+X\d+)", s)
    if retangular:
        return retangular.group(1)

    fracoes = ["1.1/8", "1.1/2", "1-1/8", "1-1/2", "7/8", "5/8", "3/4"]
    for fracao in fracoes:
        if s.startswith(fracao):
            return fracao.replace("-", ".")

    if "BARRACHATA" in s:
        return "BARRA CHATA"
    if "GONZO" in s:
        return "GONZO"
    if "CHAPA" in s:
        return "CHAPA"

    return s


def tem_grau(linha: pd.Series) -> bool:
    diag = normalizar_texto(linha.get("DIAG."))
    descricao_original = str(linha.get("DESCRIÇÃO") or "")

    if diag not in {"", "-", "NAO", "N", "0"}:
        return True

    # Há linhas em que o grau aparece em DESCRIÇÃO em vez de DIAG.
    if re.search(r"\d+\s*[°º]", descricao_original):
        return True

    return False


def eh_pdv(linha: pd.Series) -> bool:
    produto = normalizar_texto(linha.get("DESCRIÇÃO DO PRODUTO"))
    palavras = ["PDV", "PRODUTODIVERSO", "PRODUTOSDIVERSOS", "DIVERSOS"]
    return any(p in produto for p in palavras)


def meta_hora(linha: pd.Series, maquina: str = "SIMEC 1", metas=None) -> float:
    configuracao = metas_padrao() if metas is None else metas
    valores = configuracao[maquina]
    return float(valores.get(familia_tubo(linha.get("TUBO")), valores["PADRAO"]))


def chave_lote(linha: pd.Series) -> tuple:
    """Não mistura espessuras nem ângulos distintos; PDV não faz parte da chave."""
    diag = normalizar_texto(linha.get("DIAG."))
    if diag in {"", "-", "NAO", "N", "0"}:
        diag = ""
    # Também reconhece o grau escrito na descrição. Normaliza 45, 45.0 e 45°.
    angulos = tuple(float(v.replace(",", ".")) for v in re.findall(
        r"(\d+(?:[.,]\d+)?)\s*[°º]", str(linha.get("DESCRIÇÃO", ""))))
    diag_numerico = str(linha.get("DIAG.", "")).strip().replace(",", ".").rstrip("°º")
    if diag and re.fullmatch(r"\d+(?:\.\d+)?", diag_numerico):
        assinatura_grau = (float(diag_numerico), angulos or (float(diag_numerico),))
    elif not diag and angulos:
        assinatura_grau = (angulos[0], angulos)
    else:
        assinatura_grau = (diag, angulos)
    return normalizar_texto(linha.get("TUBO")), para_float(linha.get("CORTE")), assinatura_grau


def pecas_por_vara(linha: pd.Series) -> int:
    vara = para_float(linha.get("VARA"))
    corte = para_float(linha.get("CORTE"))
    if vara <= 0 or corte <= 0:
        return 1
    return max(1, math.floor(vara / corte))


def _hora_em_segundos(valor) -> int:
    """Converte str ou datetime.time em segundos desde 00:00."""
    if hasattr(valor, "hour") and hasattr(valor, "minute"):
        return int(valor.hour) * 3600 + int(valor.minute) * 60 + int(getattr(valor, "second", 0))

    texto = str(valor).strip()
    formatos = ["%H:%M:%S", "%H:%M"]
    for formato in formatos:
        try:
            h = datetime.strptime(texto, formato)
            return h.hour * 3600 + h.minute * 60 + h.second
        except ValueError:
            pass

    raise ValueError(f"Horário inválido: {valor}")


def construir_linha_do_tempo_turnos(turnos=None, sem_pausa=False, hora_extra_min=0, maquina=None) -> dict:
    """
    Monta a linha do tempo produtiva usada pelo balanceamento.

    Regras:
      - cada turno recebe uma pausa de 1 hora, exceto com sem_pausa=True;
      - hora extra prolonga somente o último turno, sem alterar a pausa regular;
      - a pausa começa no meio do intervalo total do turno;
      - término <= início significa término no dia seguinte;
      - os turnos são considerados na ordem em que foram informados;
      - turnos sobrepostos não são permitidos.
    """
    if not math.isfinite(float(hora_extra_min)) or hora_extra_min < 0:
        raise ValueError("Hora extra deve ser não negativa e finita.")
    if not turnos:
        turnos = TURNOS_PADRAO

    # Ausência da seleção mantém todas as máquinas disponíveis (compatibilidade).
    if maquina is not None and maquina not in MAQUINAS:
        raise ValueError("Máquina desconhecida na configuração dos turnos.")
    for turno in turnos:
        if set(turno.get("maquinas", MAQUINAS)) - set(MAQUINAS):
            raise ValueError("Há máquinas desconhecidas na seleção do turno.")
    if (maquina is not None and hora_extra_min
            and maquina not in turnos[-1].get("maquinas", MAQUINAS)):
        raise ValueError(f"{maquina}: zere a hora extra, pois a máquina não participa do último turno.")

    base = datetime(2000, 1, 1, 0, 0, 0)
    turnos_normalizados = []
    segmentos = []
    acumulado = 0.0

    inicio_anterior_seg = None
    fim_anterior_dt = None
    dia_inicio = 0

    for indice, configuracao in enumerate(turnos, start=1):
        participa = maquina is None or maquina in configuracao.get("maquinas", MAQUINAS)
        nome = str(configuracao.get("nome") or f"Turno {indice}")
        inicio_seg = _hora_em_segundos(configuracao.get("inicio"))
        fim_seg = _hora_em_segundos(configuracao.get("fim"))

        if inicio_seg == fim_seg:
            raise ValueError(
                f"{nome}: início e término não podem ser iguais."
            )

        if indice == 1:
            dia_inicio = 0
        elif inicio_seg <= inicio_anterior_seg:
            dia_inicio += 1

        inicio_dt = base + timedelta(days=dia_inicio, seconds=inicio_seg)

        if fim_anterior_dt is not None and inicio_dt < fim_anterior_dt:
            raise ValueError(
                f"{nome} começa antes do término do turno anterior. "
                "Ajuste os horários para que os turnos não se sobreponham."
            )

        fim_dt = base + timedelta(days=dia_inicio, seconds=fim_seg)
        if fim_seg <= inicio_seg:
            fim_dt += timedelta(days=1)

        duracao_total_min = (fim_dt - inicio_dt).total_seconds() / 60
        if participa and not sem_pausa and duracao_total_min < 120:
            raise ValueError(
                f"{nome} precisa ter pelo menos 2 horas, pois possui uma pausa obrigatória de 1 hora."
            )

        # Mantém o comportamento atual: 07:00-17:00 -> pausa 12:00-13:00.
        pausa_inicio_dt = inicio_dt + timedelta(minutes=duracao_total_min / 2)
        pausa_fim_dt = pausa_inicio_dt + timedelta(minutes=PAUSA_TURNO_MIN)

        if participa and not sem_pausa and pausa_fim_dt > fim_dt:
            raise ValueError(f"{nome}: a pausa de 1 hora não cabe dentro do turno informado.")

        blocos = [(inicio_dt, fim_dt)] if sem_pausa else [
            (inicio_dt, pausa_inicio_dt), (pausa_fim_dt, fim_dt)]
        if not participa:
            blocos = []
        # Hora extra após o último turno, sem mudar o almoço do turno regular.
        if participa and indice == len(turnos) and hora_extra_min:
            fim_dt += timedelta(minutes=hora_extra_min)
            blocos.append((blocos[-1][1], fim_dt))

        for bloco_inicio, bloco_fim in blocos:
            duracao_bloco = (bloco_fim - bloco_inicio).total_seconds() / 60
            if duracao_bloco <= 1e-9:
                continue

            segmentos.append(
                {
                    "turno": nome,
                    "turno_indice": indice,
                    "inicio_dt": bloco_inicio,
                    "fim_dt": bloco_fim,
                    "prod_inicio": acumulado,
                    "prod_fim": acumulado + duracao_bloco,
                }
            )
            acumulado += duracao_bloco

        turnos_normalizados.append(
            {
                "turno": nome,
                "indice": indice,
                "participa": participa,
                "inicio_dt": inicio_dt,
                "pausa_inicio_dt": pausa_inicio_dt,
                "pausa_fim_dt": pausa_fim_dt,
                "fim_dt": fim_dt,
                "duracao_total_min": duracao_total_min,
                "duracao_produtiva_min": (duracao_total_min - (0 if sem_pausa else PAUSA_TURNO_MIN)
                + (hora_extra_min if indice == len(turnos) else 0)) if participa else 0,
            }
        )

        inicio_anterior_seg = inicio_seg
        fim_anterior_dt = fim_dt

    if maquina is None and (not segmentos or acumulado <= 0):
        raise ValueError("Os turnos informados não possuem tempo produtivo disponível.")

    return {
        "turnos": turnos_normalizados,
        "segmentos": segmentos,
        "minutos_produtivos": acumulado,
        "data_base": turnos_normalizados[0]["inicio_dt"].date(),
    }


def _resolver_minuto_produtivo(minuto: float, linha_tempo: dict, eh_inicio: bool = False):
    """Transforma um minuto produtivo em datetime real, pulando pausas e intervalos entre turnos."""
    segmentos = linha_tempo["segmentos"]
    total = float(linha_tempo["minutos_produtivos"])
    minuto = max(0.0, min(float(minuto), total))
    eps = 1e-7

    if eh_inicio:
        for posicao, segmento in enumerate(segmentos):
            if minuto < segmento["prod_fim"] - eps:
                if minuto >= segmento["prod_inicio"] - eps:
                    deslocamento = max(0.0, minuto - segmento["prod_inicio"])
                    return segmento["inicio_dt"] + timedelta(minutes=deslocamento), segmento

            if abs(minuto - segmento["prod_fim"]) <= eps:
                if posicao + 1 < len(segmentos):
                    proximo = segmentos[posicao + 1]
                    return proximo["inicio_dt"], proximo
                return segmento["fim_dt"], segmento

    else:
        for segmento in segmentos:
            if minuto <= segmento["prod_fim"] + eps and minuto >= segmento["prod_inicio"] - eps:
                deslocamento = min(
                    max(0.0, minuto - segmento["prod_inicio"]),
                    segmento["prod_fim"] - segmento["prod_inicio"],
                )
                return segmento["inicio_dt"] + timedelta(minutes=deslocamento), segmento

    ultimo = segmentos[-1]
    return ultimo["fim_dt"], ultimo


def _formatar_horario_linha_tempo(instante: datetime, linha_tempo: dict) -> str:
    """Mostra D+N quando a programação atravessa a meia-noite."""
    deslocamento_dias = (instante.date() - linha_tempo["data_base"]).days
    prefixo = "" if deslocamento_dias == 0 else f"D+{deslocamento_dias} "
    return prefixo + instante.strftime("%H:%M:%S")


def minuto_produtivo_para_horario(minuto: float, linha_tempo: dict, eh_inicio: bool = False) -> str:
    instante, _ = _resolver_minuto_produtivo(minuto, linha_tempo, eh_inicio=eh_inicio)
    return _formatar_horario_linha_tempo(instante, linha_tempo)


def turno_do_intervalo(inicio_min: float, termino_min: float, linha_tempo: dict) -> str:
    _, seg_inicio = _resolver_minuto_produtivo(inicio_min, linha_tempo, eh_inicio=True)
    _, seg_fim = _resolver_minuto_produtivo(termino_min, linha_tempo, eh_inicio=False)

    if seg_inicio["turno"] == seg_fim["turno"]:
        return seg_inicio["turno"]
    return f'{seg_inicio["turno"]} -> {seg_fim["turno"]}'


def resumir_turnos(turnos=None, sem_pausa=False, hora_extra_min=0, maquina=None) -> pd.DataFrame:
    """Retorna uma tabela amigável para exibir no Streamlit e no Excel."""
    linha_tempo = construir_linha_do_tempo_turnos(turnos, sem_pausa, hora_extra_min, maquina)
    linhas = []

    for turno in linha_tempo["turnos"]:
        linhas.append(
            {
                "TURNO": turno["turno"],
                "INÍCIO": _formatar_horario_linha_tempo(turno["inicio_dt"], linha_tempo),
                **({"PARTICIPA": "Sim" if turno["participa"] else "Não"} if maquina is not None else {}),
                "PAUSA": "—" if not turno["participa"] else "Sem pausa" if sem_pausa else _formatar_horario_linha_tempo(turno["pausa_inicio_dt"], linha_tempo),
                "RETORNO": "—" if not turno["participa"] or sem_pausa else _formatar_horario_linha_tempo(turno["pausa_fim_dt"], linha_tempo),
                "TÉRMINO": _formatar_horario_linha_tempo(turno["fim_dt"], linha_tempo),
                "HORAS PRODUTIVAS": minutos_para_duracao(turno["duracao_produtiva_min"]),
            }
        )

    return pd.DataFrame(linhas)

def minutos_para_duracao(minutos: float) -> str:
    segundos = int(round(minutos * 60))
    horas = segundos // 3600
    resto = segundos % 3600
    mins = resto // 60
    segs = resto % 60
    return f"{horas:02d}:{mins:02d}:{segs:02d}"


def maquinas_elegiveis(linha: pd.Series, consolidado: bool = False) -> tuple[list[str], str]:
    fam = familia_tubo(linha.get("TUBO"))
    corte = para_float(linha.get("CORTE"))
    quantidade = int(round(para_float(linha.get("QUANT."))))
    grau = tem_grau(linha)
    produto = normalizar_texto(linha.get("DESCRIÇÃO DO PRODUTO"))
    descricao = normalizar_texto(linha.get("DESCRIÇÃO"))

    # 1) Regras que obrigam a serra
    if corte > 1812:
        return ["SERRA 1", "SERRA 2"], "CORTE > 1812 mm"

    if fam == "20X48" and abs(corte - 1690) < 0.5:
        return ["SERRA 1", "SERRA 2"], "20X48 COM CORTE 1690 mm"

    if "MCP-2-DABS1.20MTPNOVO" in produto and ("TAMPO" in descricao or "BALIZA" in descricao):
        return ["SERRA 1", "SERRA 2"], "TAMPO/BALIZA MCP-2-D ABS"

    # 2) PDV com menos de 100 peças -> serra, exceto se tiver grau
    if not consolidado and eh_pdv(linha) and quantidade < 100 and not grau:
        return ["SERRA 1", "SERRA 2"], "PDV < 100 PEÇAS"

    # 3) Grau -> prioriza as SIMEC; serra fica como fallback
    if grau:
        return ["SIMEC 1", "SIMEC 2", "SERRA 1", "SERRA 2"], "CORTE COM GRAU"

    # 4) Cortes menores que 100 mm -> SIMEC
    if corte < 100:
        return ["SIMEC 1", "SIMEC 2"], "CORTE < 100 mm"

    # 5) Regras por família do tubo
    candidatas: list[str] = []

    if fam in TUBOS_EMT:
        candidatas.append("EMT")
    if fam in TUBOS_BLM:
        candidatas.append("BLM")
    if fam in TUBOS_TECNOMAQ:
        candidatas.append("TECNOMAQ")
    if fam in TUBOS_SERRA:
        candidatas.extend(["SERRA 1", "SERRA 2"])

    if candidatas:
        return candidatas, "REGRA POR TUBO"

    return [], "SEM REGRA CADASTRADA"


def preparar_entrada(df: pd.DataFrame) -> pd.DataFrame:
    faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Colunas obrigatórias ausentes: {faltantes}")

    df = df.copy()
    df = df.dropna(how="all")

    # Mantém código como texto, preservando zeros à esquerda.
    df["CÓD."] = (
        df["CÓD."]
        .fillna("")
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.zfill(10)
    )
    df["PRODUTO"] = df["PRODUTO"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)

    df["QUANT."] = pd.to_numeric(df["QUANT."], errors="coerce").fillna(0).astype(int)
    df = df[df["QUANT."] > 0].reset_index(drop=True)

    return df


# ============================================================
# ORGANIZAÇÃO / BALANCEAMENTO
# ============================================================

def organizar_programacao(
    arquivo_entrada: str | Path | BinaryIO,
    aba: str = "Cortes",
    turnos: list[dict] | None = None,
    metas: dict | None = None,
    maquinas_gargalo: Iterable[str] | None = None,
    hora_extra_min: dict | None = None,
    perdas_min: dict | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      1) programacao_df
      2) nao_alocados_df
      3) resumo_df

    O algoritmo tenta:
      - respeitar as máquinas permitidas e a participação em cada turno;
      - priorizar SIMEC para grau;
      - usar serra como fallback para grau;
      - equilibrar máquinas equivalentes;
      - aplicar pausa de 1 hora apenas às máquinas que não são gargalo;
      - reservar perdas no início e disponibilizar hora extra após o último turno;
      - usar metas por máquina e consolidar tubo+corte+grau;
      - respeitar os horários de início e término configurados;
      - fracionar um lote apenas quando ele não cabe inteiro em uma máquina.
    """

    metas, gargalos, extras, perdas = validar_configuracao(
        metas, maquinas_gargalo, hora_extra_min, perdas_min)
    tempos = {m: construir_linha_do_tempo_turnos(turnos, m in gargalos, extras[m], maquina=m)
              for m in MAQUINAS}
    limites = {m: float(tempos[m]["minutos_produtivos"]) for m in MAQUINAS}
    if any(perdas[m] > limites[m] for m in MAQUINAS):
        raise ValueError("As perdas não podem exceder o tempo disponível da máquina.")

    df = preparar_entrada(pd.read_excel(arquivo_entrada, sheet_name=aba))
    agrupadas = {}
    for indice, linha in df.iterrows():
        agrupadas.setdefault(chave_lote(linha), []).append((indice, linha))

    grupos_compativeis = []
    for membros in agrupadas.values():
        # PDV pequeno pode compartilhar um lote igual; regras obrigatórias permanecem.
        regras = [maquinas_elegiveis(linha, consolidado=len(membros) > 1)
                  for _, linha in membros]
        candidatas = [m for m in regras[0][0] if all(m in r[0] for r in regras)]
        if not candidatas and len(membros) > 1:
            # Se outra regra obriga máquinas incompatíveis, preserva cada pedido.
            for membro in membros:
                regra = maquinas_elegiveis(membro[1])
                grupos_compativeis.append(([membro], [regra], regra[0]))
        else:
            grupos_compativeis.append((membros, regras, candidatas))

    tarefas = []
    for grupo_id, (membros, regras, candidatas) in enumerate(grupos_compativeis):
        for indice, linha in membros:
            taxa = min((meta_hora(linha, m, metas) for m in candidatas), default=META_H_PADRAO)
            qtd = int(linha["QUANT."])
            ppv = pecas_por_vara(linha)
            vara, corte = para_float(linha.get("VARA")), para_float(linha.get("CORTE"))
            tarefas.append({
                "indice_origem": indice, "linha": linha, "grupo_id": grupo_id,
                "familia": familia_tubo(linha.get("TUBO")), "grau": tem_grau(linha),
                "quantidade": qtd, "maquinas": candidatas,
                "regra": regras[0][1] if len(membros) == 1 else "LOTE CONSOLIDADO",
                "regra_prioridade": regras[0][1],
                "pecas_por_vara": ppv,
                "sobra_por_vara": max(0.0, vara - ppv * corte) if vara > 0 and corte > 0 else 0.0,
                "duracao_total_min": qtd / taxa * 60,
            })

    # Perdas são reservadas no início da janela produtiva da máquina.
    carga = dict(perdas)
    agenda = {m: [] for m in MAQUINAS}
    nao_alocados = []

    def taxa_maquina(tarefa, maquina):
        return meta_hora(tarefa["linha"], maquina, metas)

    def capacidade_em_pecas(maquina, tarefa):
        disponivel = limites[maquina] - carga[maquina]
        return max(0, math.floor(disponivel * taxa_maquina(tarefa, maquina) / 60 + 1e-9))

    def adicionar_lote(tarefa: dict, maquina: str, quantidade: int) -> None:
        taxa = taxa_maquina(tarefa, maquina)
        duracao = quantidade / taxa * 60
        inicio_min = carga[maquina]
        termino_min = inicio_min + duracao

        if termino_min > limites[maquina] + 1e-7:
            raise RuntimeError(f"Tentativa de ultrapassar a capacidade dos turnos configurados em {maquina}")

        agenda[maquina].append(
            {
                "tarefa": tarefa,
                "quantidade_programada": quantidade,
                "inicio_min": inicio_min,
                "termino_min": termino_min,
                "duracao_min": duracao,
                "meta_h": taxa,
                "quant_varas": math.ceil(quantidade / tarefa["pecas_por_vara"]),
            }
        )
        carga[maquina] = termino_min

    def dividir_equilibrado(tarefa: dict, maquinas: Iterable[str], quantidade: int) -> list[tuple[str, int]]:
        """
        Quando um lote não cabe inteiro em nenhuma máquina do grupo,
        divide entre as máquinas tentando terminar com cargas semelhantes.
        """
        candidatas = []
        for m in maquinas:
            cap = capacidade_em_pecas(m, tarefa)
            if cap > 0:
                base = carga[m]
                candidatas.append((m, base, cap))

        if not candidatas or quantidade <= 0:
            return []

        qtd_alvo = min(quantidade, sum(cap for _, _, cap in candidatas))

        # Procura por um horário final comum (water filling).
        baixo = min(base for _, base, _ in candidatas)
        alto = max(limites[m] for m, _, _ in candidatas)

        for _ in range(50):
            meio = (baixo + alto) / 2
            total = 0
            for m, base, cap in candidatas:
                total += min(cap, max(0, math.floor((meio - base) * taxa_maquina(tarefa, m) / 60)))
            if total >= qtd_alvo:
                alto = meio
            else:
                baixo = meio

        alvo_tempo = alto
        alocacao = {}
        total = 0

        for m, base, cap in candidatas:
            qtd = min(cap, max(0, math.floor((alvo_tempo - base) * taxa_maquina(tarefa, m) / 60)))
            alocacao[m] = qtd
            total += qtd

        # Ajuste fino para bater exatamente qtd_alvo.
        while total > qtd_alvo:
            maquinas_com_qtd = [m for m, qtd in alocacao.items() if qtd > 0]
            m = max(
                maquinas_com_qtd,
                key=lambda x: next(base for mm, base, _ in candidatas if mm == x)
                + alocacao[x] / taxa_maquina(tarefa, x) * 60,
            )
            alocacao[m] -= 1
            total -= 1

        while total < qtd_alvo:
            opcoes = []
            for m, base, cap in candidatas:
                if alocacao[m] < cap:
                    novo_fim = base + (alocacao[m] + 1) / taxa_maquina(tarefa, m) * 60
                    opcoes.append((novo_fim, m))
            if not opcoes:
                break
            _, m = min(opcoes)
            alocacao[m] += 1
            total += 1

        return [(m, qtd) for m, qtd in alocacao.items() if qtd > 0]

    def chave_prioridade(tarefa: dict):
        sem_regra = 1 if not tarefa["maquinas"] else 0
        regra_forcada = 0 if tarefa["regra_prioridade"] not in {"REGRA POR TUBO", "CORTE COM GRAU"} else 1
        numero_opcoes = len(tarefa["maquinas"]) if tarefa["maquinas"] else 99

        # Para tarefas flexíveis, a sobra maior vem antes. Isso ajuda a preservar
        # a lógica de manter os cortes de maior sobra nas máquinas mais adequadas.
        return (
            sem_regra,
            regra_forcada,
            numero_opcoes,
            -tarefa["duracao_total_min"],
            -tarefa["sobra_por_vara"],
        )

    grupos_lotes = {}
    for tarefa in tarefas:
        grupos_lotes.setdefault(tarefa["grupo_id"], []).append(tarefa)
    grupos_ordenados = sorted(grupos_lotes.values(), key=lambda g: chave_prioridade({
        **g[0], "duracao_total_min": sum(t["duracao_total_min"] for t in g)}))

    for membros in grupos_ordenados:
        tarefa = membros[0]
        restante = sum(t["quantidade"] for t in membros)
        pendentes = [[t, t["quantidade"]] for t in membros]

        def registrar(maquina, quantidade):
            for item in pendentes:
                qtd = min(item[1], quantidade)
                if qtd:
                    adicionar_lote(item[0], maquina, qtd)
                    item[1] -= qtd
                    quantidade -= qtd
                if quantidade == 0:
                    break

        if tarefa["grau"]:
            grupos = [[m for m in tarefa["maquinas"] if m.startswith("SIMEC")],
                      [m for m in tarefa["maquinas"] if m.startswith("SERRA")]]
        else:
            grupos = [tarefa["maquinas"]]

        for grupo in grupos:
            if restante <= 0 or not grupo:
                continue
            opcoes_inteiras = []
            for maquina in grupo:
                if capacidade_em_pecas(maquina, tarefa) >= restante:
                    fim_projetado = carga[maquina] + restante / taxa_maquina(tarefa, maquina) * 60
                    opcoes_inteiras.append((fim_projetado, carga[maquina], maquina))
            if opcoes_inteiras:
                escolhida = min(opcoes_inteiras)[2]
                registrar(escolhida, restante)
                restante = 0
                break
            for maquina, qtd in dividir_equilibrado(tarefa, grupo, restante):
                registrar(maquina, qtd)
                restante -= qtd

        for t, qtd in pendentes:
            if qtd:
                if not t["maquinas"]:
                    motivo = "Sem máquina elegível pela regra atual"
                elif all(limites[m] == 0 for m in t["maquinas"]):
                    motivo = "Máquinas elegíveis não participam dos turnos configurados"
                else:
                    motivo = "Capacidade diária insuficiente"
                nao_alocados.append((t, qtd, motivo))

    # ========================================================
    # MONTA DATAFRAMES DE SAÍDA
    # ========================================================
    linhas_programacao = []

    for maquina in MAQUINAS:
        linha_tempo = tempos[maquina]
        for lote in agenda[maquina]:
            tarefa = lote["tarefa"]
            linha = tarefa["linha"]

            linhas_programacao.append(
                {
                    "MÁQUINA": maquina,
                    "TURNO": turno_do_intervalo(
                        lote["inicio_min"], lote["termino_min"], linha_tempo
                    ),
                    "CÓD.": linha.get("CÓD."),
                    "PRODUTO": linha.get("PRODUTO"),
                    "TUBO": linha.get("TUBO"),
                    "VARA": linha.get("VARA"),
                    "CORTE": linha.get("CORTE"),
                    "DIAG.": linha.get("DIAG."),
                    "DESCRIÇÃO": linha.get("DESCRIÇÃO"),
                    "DESCRIÇÃO DO PRODUTO": linha.get("DESCRIÇÃO DO PRODUTO"),
                    "QUANT. ORIGINAL": tarefa["quantidade"],
                    "QUANT. PROGRAMADA": lote["quantidade_programada"],
                    "META/H": lote["meta_h"],
                    "H/MÁQ": minutos_para_duracao(lote["duracao_min"]),
                    "INÍCIO": minuto_produtivo_para_horario(
                        lote["inicio_min"], linha_tempo, eh_inicio=True
                    ),
                    "TÉRMINO": minuto_produtivo_para_horario(
                        lote["termino_min"], linha_tempo, eh_inicio=False
                    ),
                    "QUANT. VARAS": lote["quant_varas"],
                    "PEÇAS/VARA": tarefa["pecas_por_vara"],
                    "SOBRA/VARA (mm)": round(tarefa["sobra_por_vara"], 2),
                    "REGRA": tarefa["regra"],
                }
            )

    programacao_df = pd.DataFrame(linhas_programacao)

    linhas_nao_alocados = []
    for tarefa, restante, motivo in nao_alocados:
        linha = tarefa["linha"]
        linhas_nao_alocados.append(
            {
                "CÓD.": linha.get("CÓD."),
                "PRODUTO": linha.get("PRODUTO"),
                "TUBO": linha.get("TUBO"),
                "VARA": linha.get("VARA"),
                "CORTE": linha.get("CORTE"),
                "DIAG.": linha.get("DIAG."),
                "DESCRIÇÃO": linha.get("DESCRIÇÃO"),
                "DESCRIÇÃO DO PRODUTO": linha.get("DESCRIÇÃO DO PRODUTO"),
                "QUANT. ORIGINAL": tarefa["quantidade"],
                "QUANT. RESTANTE": restante,
                "FAMÍLIA": tarefa["familia"],
                "REGRA": tarefa["regra"],
                "MOTIVO": motivo,
            }
        )

    nao_alocados_df = pd.DataFrame(linhas_nao_alocados)

    linhas_resumo = []
    for maquina in MAQUINAS:
        lotes = agenda[maquina]
        linhas_resumo.append(
            {
                "MÁQUINA": maquina,
                "TEMPO OCUPADO": minutos_para_duracao(carga[maquina] - perdas[maquina]),
                "PERDAS": minutos_para_duracao(perdas[maquina]),
                "HORA EXTRA DISPONÍVEL": minutos_para_duracao(extras[maquina]),
                "PAUSA ALMOÇO": "Não participa" if limites[maquina] == 0 else "Sem pausa" if maquina in gargalos else "01:00:00 por turno",
                "TEMPO DISPONÍVEL": minutos_para_duracao(limites[maquina] - perdas[maquina]),
                "UTILIZAÇÃO %": round((carga[maquina] - perdas[maquina]) / (limites[maquina] - perdas[maquina]) * 100, 1) if limites[maquina] > perdas[maquina] else 0.0,
                "PEÇAS PROGRAMADAS": sum(l["quantidade_programada"] for l in lotes),
                "QUANT. VARAS": sum(l["quant_varas"] for l in lotes),
                "LOTES": len(lotes),
            }
        )

    resumo_df = pd.DataFrame(linhas_resumo)

    return programacao_df, nao_alocados_df, resumo_df


# ============================================================
# EXPORTAÇÃO PARA EXCEL
# ============================================================

def exportar_programacao(
    programacao_df: pd.DataFrame,
    nao_alocados_df: pd.DataFrame,
    resumo_df: pd.DataFrame,
    caminho_saida: str | Path,
) -> None:
    caminho_saida = Path(caminho_saida)

    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        programacao_df.to_excel(writer, sheet_name="Programacao", index=False)
        nao_alocados_df.to_excel(writer, sheet_name="Nao_alocados", index=False)
        resumo_df.to_excel(writer, sheet_name="Resumo", index=False)

        # Ajuste simples de larguras para facilitar o teste.
        for nome_aba in ["Programacao", "Nao_alocados", "Resumo"]:
            ws = writer.book[nome_aba]
            ws.freeze_panes = "A2"
            for coluna in ws.columns:
                letra = coluna[0].column_letter
                maior = 0
                for celula in coluna:
                    valor = "" if celula.value is None else str(celula.value)
                    maior = max(maior, len(valor))
                ws.column_dimensions[letra].width = min(max(maior + 2, 10), 32)


def main():
    # Busca automaticamente a base de cortes dentro da pasta do projeto.
    entrada = localizar_arquivo_planejamento(PASTA_PROJETO)
    saida = PASTA_PROJETO / "programacao_maquinas.xlsx"

    print(f"\nPasta do projeto: {PASTA_PROJETO}")
    print(f"Arquivo de cortes encontrado: {entrada.name}")

    programacao, nao_alocados, resumo = organizar_programacao(entrada, aba="Cortes")
    exportar_programacao(programacao, nao_alocados, resumo, saida)

    print("\nPROGRAMAÇÃO GERADA:")
    print(programacao.to_string(index=False))

    print("\nRESUMO:")
    print(resumo.to_string(index=False))

    if nao_alocados.empty:
        print("\nTodos os cortes foram alocados dentro do turno.")
    else:
        print("\nATENÇÃO - CORTES NÃO ALOCADOS:")
        print(nao_alocados.to_string(index=False))

    print(f"\nArquivo salvo em: {saida.resolve()}")


if __name__ == "__main__":
    main()
