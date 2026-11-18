# -*- coding: utf-8 -*-
"""
painel_core.py — Núcleo de Estatística de Dados em Painel e Geração de Documentos.
Autor: Amarildo Fernandes - amarildo@ufrj.br

Módulo com funções de formatação, leitura de planilhas, preparação de dados em
painel (formato longo balanceado), estimadores (MQO agrupado, efeitos fixos
within e LSDV, efeitos aleatórios Swamy-Arora), testes (F, LM de Breusch-Pagan,
Hausman via Mundlak), erros padrão agrupados, renderização de equações LaTeX e
geração de relatório em .docx (python-docx) e .pdf (reportlab).
"""
from __future__ import annotations
 
import math
import os
import warnings
from typing import Any, Optional
 
VERSAO = "6.0"   # versão do módulo
 
import numpy as np
import pandas as pd
from scipy import linalg  # noqa: F401 — disponibilizado para cálculos matriciais
from scipy import stats
 
LIMITE_WITHIN_LSDV = 1e-6
 
# ----------------------------------------------------------------------------
# Formatação
# ----------------------------------------------------------------------------
 
def fmt(v: Any, casas: int = 3) -> str:
    """Formata número com vírgula decimal brasileira; '—' para None/NaN."""
    if v is None:
        return "—"
    try:
        if isinstance(v, float) and math.isnan(v):
            return "—"
        if pd.isna(v):
            return "—"
    except (TypeError, ValueError):
        pass
    s = f"{float(v):,.{casas}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")
 
 
def fmt_p(p: Optional[float]) -> str:
    """Formata valor-p: '< 0,0001' se p < 0.0001; senão 4 casas com vírgula."""
    if p is None or (isinstance(p, float) and math.isnan(p)):
        return "—"
    if p < 0.0001:
        return "< 0,0001"
    return f"{p:.4f}".replace(".", ",")
 
 
def estrelas(p: Optional[float]) -> str:
    """Significância em estrelas: ***/**/* para p < 0.01/0.05/0.10."""
    if p is None or (isinstance(p, float) and math.isnan(p)):
        return ""
    if p < 0.01:
        return "***"
    if p < 0.05:
        return "**"
    if p < 0.10:
        return "*"
    return ""
 
# ----------------------------------------------------------------------------
# Leitura e perfilamento
# ----------------------------------------------------------------------------
 
def ler_planilha(arquivo: Any) -> dict[str, pd.DataFrame]:
    """Lê todas as abas de uma planilha Excel em um dicionário de DataFrames."""
    return pd.read_excel(arquivo, sheet_name=None)
 
 
def perfilar_base(df: pd.DataFrame) -> dict:
    """Perfil básico da base: linhas, colunas, tipos e problemas de dados."""
    numericas = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    categoricas = [c for c in df.columns if c not in numericas]
    return {
        "n_linhas": int(df.shape[0]),
        "n_colunas": int(df.shape[1]),
        "colunas_numericas": numericas,
        "colunas_categoricas": categoricas,
        "com_ausentes": [c for c in df.columns if df[c].isna().any()],
        "com_valor_unico": [c for c in df.columns if df[c].nunique(dropna=False) <= 1],
    }
 
# ----------------------------------------------------------------------------
# Preparação dos dados
# ----------------------------------------------------------------------------
 
def preparar_dados(df: pd.DataFrame, cfg: dict):
    """
    Prepara os dados em painel: converte tipos, ordena por id e tempo, cria
    dummies de tempo (drop_first) e de invariantes (drop_first), remove
    ausentes e calcula N (unidades), T (períodos), n (observações).
    """
    dados = df.copy()
    id_c, tempo_c, y_c = cfg["id"], cfg["tempo"], cfg["y"]
    interesses = list(cfg.get("interesse", []))
    controles = list(cfg.get("controles", []))
    invariantes = list(cfg.get("invariantes", []))
 
    dados[id_c] = dados[id_c].astype(str)
    dados[tempo_c] = pd.to_numeric(dados[tempo_c], errors="coerce")
    dados[y_c] = pd.to_numeric(dados[y_c], errors="coerce")
    for c in interesses + controles:
        dados[c] = pd.to_numeric(dados[c], errors="coerce")
 
    n_antes = len(dados)
    dados = dados.sort_values([id_c, tempo_c]).reset_index(drop=True)
 
    # Dummies de tempo (drop_first) — colunas de controle que variam no tempo
    dummies_tempo = pd.get_dummies(dados[tempo_c].astype(str), prefix="__t", drop_first=True, dtype=float)
 
    # Dummies das invariantes (drop_first)
    if invariantes:
        dummies_inv = pd.get_dummies(dados[invariantes].astype(str), prefix="__inv", drop_first=True, dtype=float)
        nomes_invariantes = list(dummies_inv.columns)
    else:
        dummies_inv = pd.DataFrame(index=dados.index)
        nomes_invariantes = []
 
    controles_t = controles  # controles que variam no tempo
    X_tempo_df = pd.concat([dados[interesses + controles_t], dummies_tempo], axis=1)
    nomes_tempo = list(X_tempo_df.columns)
 
    df_modelo = pd.concat([dados[[id_c, tempo_c, y_c]], X_tempo_df, dummies_inv], axis=1)
    antes_modelo = len(df_modelo)
    df_modelo = df_modelo.dropna().reset_index(drop=True)
    removidas = antes_modelo - len(df_modelo)
 
    id_arr = df_modelo[id_c].to_numpy()
    tempo_arr = df_modelo[tempo_c].to_numpy()
    y = df_modelo[y_c].to_numpy(dtype=float)
    X_tempo = df_modelo[nomes_tempo].to_numpy(dtype=float) if nomes_tempo else np.empty((len(y), 0))
    X_invariantes = df_modelo[nomes_invariantes].to_numpy(dtype=float) if nomes_invariantes else np.empty((len(y), 0))
    X_completo = np.hstack([X_tempo, X_invariantes])
    nomes_completo = nomes_tempo + nomes_invariantes
 
    unidades = pd.unique(id_arr)
    N = len(unidades)
    contagens = df_modelo.groupby(id_c).size()
    T = int(contagens.max())
    balanceado = bool((contagens == T).all())
    diagnostico = {
        "y": y, "X_tempo": X_tempo, "nomes_tempo": nomes_tempo,
        "X_invariantes": X_invariantes, "nomes_invariantes": nomes_invariantes,
        "X_completo": X_completo, "nomes_completo": nomes_completo,
        "id_arr": id_arr, "tempo_arr": tempo_arr,
        "N": N, "T": T, "n": len(y), "balanceado": balanceado, "removidas": int(removidas),
    }
    return df_modelo, diagnostico
 
# ----------------------------------------------------------------------------
# Estimadores
# ----------------------------------------------------------------------------
 
def estimar_mqo_agrupado(y: np.ndarray, X: np.ndarray) -> dict:
    """MQO agrupado (pooled OLS) com intercepto: coef, ep, R2, R2_aj, SSR, gl."""
    n, k = X.shape[0], X.shape[1]
    Xc = np.hstack([np.ones((n, 1)), X])
    beta, *_ = np.linalg.lstsq(Xc, y, rcond=None)
    resid = y - Xc @ beta
    SSR = float(resid @ resid)
    ybar = y.mean()
    SST = float(((y - ybar) ** 2).sum())
    R2 = 1.0 - SSR / SST if SST > 0 else float("nan")
    gl = n - (k + 1)
    R2_aj = 1.0 - (1.0 - R2) * (n - 1) / gl if gl > 0 else float("nan")
    sigma2 = SSR / gl if gl > 0 else float("nan")
    covb = sigma2 * np.linalg.pinv(Xc.T @ Xc)
    ep = np.sqrt(np.diag(covb))
    return {"coef": beta[1:], "ep": ep[1:], "R2": R2, "R2_aj": R2_aj, "SSR": SSR, "gl": gl, "n": n, "k": k}
 
 
def estimar_efeitos_fixos(y: np.ndarray, X_tempo: np.ndarray, id_arr: np.ndarray) -> dict:
    """
    Efeitos fixos by within (demeaning por unidade) em y e em TODAS as colunas
    de X_tempo; MQO sem intercepto. Os resíduos são calculados SOBRE O y
    TRANSFORMADO (yw = y - média de y por unidade) — ponto crítico para o R2
    within e para o Swamy-Arora.
    """
    df = pd.DataFrame(X_tempo)
    df["__id"] = id_arr
    df["__y"] = y
    medias = df.groupby("__id").transform("mean")
    yw = df["__y"].to_numpy(float) - medias["__y"].to_numpy(float)
    Xw = df.drop(columns="__y").drop(columns="__id").to_numpy(float) - medias.drop(columns="__y").to_numpy(float)
    n, k = Xw.shape
    beta, *_ = np.linalg.lstsq(Xw, yw, rcond=None)
    residuos = yw - Xw @ beta          # resíduos sobre o y TRANSFORMADO
    SSR = float(residuos @ residuos)
    SST = float(((yw - yw.mean()) ** 2).sum())
    R2_within = 1.0 - SSR / SST if SST > 0 else float("nan")
    N = len(pd.unique(id_arr))
    gl = n - N - k
    sigma2 = SSR / gl if gl > 0 else float("nan")
    covb = sigma2 * np.linalg.pinv(Xw.T @ Xw)
    ep = np.sqrt(np.diag(covb))
    fitted = Xw @ beta
    return {"coef": beta, "ep": ep, "R2_within": R2_within, "residuos": residuos,
            "fitted": fitted, "gl": gl, "SSR": SSR, "N": N, "k": k}
 
 
def efeitos_fixos_lsdv(y: np.ndarray, X_tempo: np.ndarray, id_arr: np.ndarray) -> dict:
    """Efeitos fixos por LSDV: dummies de unidade + X_tempo, MQO com intercepto."""
    dummies = pd.get_dummies(pd.Series(id_arr, dtype=str), prefix="__u", drop_first=True, dtype=float).to_numpy(float)
    Xc = np.hstack([np.ones((len(y), 1)), X_tempo, dummies])
    beta, *_ = np.linalg.lstsq(Xc, y, rcond=None)
    return {"coef": beta[1:1 + X_tempo.shape[1]], "beta_completo": beta, "k_intercepto": 1}
 
 
def verificar_within_lsdv(res_within: dict, res_lsdv: dict) -> dict:
    """Compara os coeficientes within com os do LSDV (diferença máxima < 1e-6)."""
    a, b = np.asarray(res_within["coef"], float), np.asarray(res_lsdv["coef"], float)
    if a.shape != b.shape:
        return {"diferenca_maxima": float("nan"), "aprovado": False}
    dif = float(np.max(np.abs(a - b)))
    return {"diferenca_maxima": dif, "aprovado": bool(dif < LIMITE_WITHIN_LSDV)}
 
 
def estimar_efeitos_aleatorios(y: np.ndarray, X_completo: np.ndarray, id_arr: np.ndarray, T: int) -> dict:
    """
    Efeitos aleatórios (Swamy-Arora): sigma2_e dos resíduos do within com gl
    corretos; sigma2_u da regressão between descontando sigma2_e/T; quasi-
    demeaning com theta; MQO sobre Xq, yq; R2 overall.
    """
    try:
        from painel_core import estimar_efeitos_fixos as _ef  # uso normal (módulo importado)
    except ModuleNotFoundError:
        _ef = globals()["estimar_efeitos_fixos"]              # módulo carregado diretamente
    res_w = _ef(y, X_completo, id_arr)
    n, k = X_completo.shape
    sigma2_e = res_w["SSR"] / res_w["gl"] if res_w["gl"] > 0 else 0.0
 
    # Between: médias por unidade
    df = pd.DataFrame(X_completo); df["__id"] = id_arr; df["__y"] = y
    g = df.groupby("__id").mean()
    ybar = g["__y"].to_numpy(float)
    Xbar = g.drop(columns="__y").to_numpy(float)
    Xb = np.hstack([np.ones((Xbar.shape[0], 1)), Xbar])
    beta_b, *_ = np.linalg.lstsq(Xb, ybar, rcond=None)
    res_b = ybar - Xb @ beta_b
    Nb, kb = Xb.shape
    ssr_b = float(res_b @ res_b)
    gl_b = Nb - kb
    sigma2_1 = ssr_b / gl_b if gl_b > 0 else 0.0
    sigma2_u = sigma2_1 - sigma2_e / T if T > 0 else 0.0
    sigma2_u = max(sigma2_u, 0.0)
 
    denom = T * sigma2_u + sigma2_e
    theta = 1.0 - math.sqrt(sigma2_e / denom) if denom > 0 else 0.0
 
    med = df.groupby("__id").mean()
    Xbar_linhas = med.drop(columns="__y").reindex(pd.Series(id_arr)).to_numpy(float)
    ybar_linhas = med["__y"].reindex(pd.Series(id_arr)).to_numpy(float)
    Xq = X_completo - theta * Xbar_linhas
    yq = y - theta * ybar_linhas
    Xc = np.hstack([np.ones((n, 1)), Xq])
    beta, *_ = np.linalg.lstsq(Xc, yq, rcond=None)
    resid = yq - Xc @ beta
    SSR = float(resid @ resid)
    SST = float(((yq - yq.mean()) ** 2).sum())
    R2 = 1.0 - SSR / SST if SST > 0 else float("nan")
    gl = n - (k + 1)
    sigma2 = SSR / gl if gl > 0 else float("nan")
    covb = sigma2 * np.linalg.pinv(Xc.T @ Xc)
    ep = np.sqrt(np.diag(covb))
    return {"coef": beta[1:], "ep": ep[1:], "theta": float(theta), "sigma2_e": float(sigma2_e),
            "sigma2_u": float(sigma2_u), "R2_overall": R2}
 
# ----------------------------------------------------------------------------
# Testes
# ----------------------------------------------------------------------------
 
def teste_f(ssr_restrito: float, ssr_completo: float, n: int, N: int, k: int) -> dict:
    """Teste F dos efeitos fixos (efeitos nulos): F, gls e valor-p."""
    gl_num = N - 1
    gl_den = n - N - k
    if gl_den <= 0 or gl_num <= 0:
        return {"F": float("nan"), "gl_num": gl_num, "gl_den": gl_den, "p": float("nan")}
    F = ((ssr_restrito - ssr_completo) / gl_num) / (ssr_completo / gl_den)
    p = float(stats.f.sf(F, gl_num, gl_den))
    return {"F": float(F), "gl_num": gl_num, "gl_den": gl_den, "p": p}
 
 
def teste_lm_breusch_pagan(residuos: np.ndarray, id_arr: np.ndarray, T: int, n: int) -> dict:
    """Teste LM de Breusch-Pagan para variância dos efeitos aleatórios."""
    df = pd.DataFrame({"__id": id_arr, "__e": np.asarray(residuos, float)})
    soma_i = df.groupby("__id")["__e"].sum()
    soma2_i = float((soma_i ** 2).sum())
    soma_e2 = float((df["__e"] ** 2).sum())
    if soma_e2 <= 0 or T <= 1:
        return {"LM": float("nan"), "p": float("nan")}
    LM = (n * T / (2.0 * (T - 1))) * ((soma2_i / soma_e2) - 1.0) ** 2
    p = float(stats.chi2.sf(LM, 1))
    return {"LM": float(LM), "p": p}
 
 
def teste_hausman_mundlak(y: np.ndarray, X_tempo: np.ndarray, X_invariantes: np.ndarray,
                          id_arr: np.ndarray, T: int) -> dict:
    """
    Teste de Hausman via Mundlak: modelo quasi-demeaned acrescido das médias por
    unidade das variáveis que variam no tempo; Wald conjunto b' V^-1 b sobre os
    coeficientes das médias. Observação: o Hausman clássico pode não ser
    definido positivo em amostras pequenas; Mundlak é mais estável.
    """
    k = X_tempo.shape[1]
    if k == 0:
        return {"Wald": float("nan"), "gl": 0, "p": float("nan"), "coef_medias": np.array([])}
    X_completo = np.hstack([X_tempo, X_invariantes])
    df = pd.DataFrame(X_tempo); df["__id"] = id_arr; df["__y"] = y
    med = df.drop(columns="__y").groupby("__id").mean()
    med_linhas = med.reindex(pd.Series(id_arr)).to_numpy(float)
    Xm = np.hstack([X_completo, med_linhas])
    # reusa quasi-demeaning com theta = 1 (aproximação within + médias) sobre demeaned
    res = estimar_efeitos_aleatorios(y, Xm, id_arr, T)
    b = np.asarray(res["coef"], float)[-k:]
    # covariância aproximada: usa a matriz de variâncias dos coeficientes das médias
    # (bloco diagonal conservador a partir dos ep)
    epm = np.asarray(res["ep"], float)[-k:]
    V = np.diag(epm ** 2)
    try:
        Wald = float(b @ np.linalg.pinv(V) @ b)
    except np.linalg.LinAlgError:
        Wald = float("nan")
    p = float(stats.chi2.sf(Wald, k)) if np.isfinite(Wald) else float("nan")
    return {"Wald": Wald, "gl": k, "p": p, "coef_medias": b}
 
# ----------------------------------------------------------------------------
# Erros padrão agrupados
# ----------------------------------------------------------------------------
 
def ep_agrupado(y: np.ndarray, X_tempo: np.ndarray, id_arr: np.ndarray, idx: dict) -> dict:
    """
    Erros padrão cluster-robust por unidade para o MQO agrupado (idx: resultado
    de estimar_mqo_agrupado), com correção G/(G-1) * (n-1)/(n-k). t bilateral, gl = N-1.
    """
    n, k = X_tempo.shape
    G = len(pd.unique(id_arr))
    Xc = np.hstack([np.ones((n, 1)), X_tempo])
    beta = np.hstack([0.0, np.asarray(idx["coef"], float)])  # intercepto 0 não usado; refit abaixo
    beta, *_ = np.linalg.lstsq(Xc, y, rcond=None)
    u = y - Xc @ beta
    # meat por cluster
    meat = np.zeros((k + 1, k + 1))
    dfX = pd.DataFrame(Xc); dfX["__id"] = id_arr; dfX["__u"] = u
    for _, g in dfX.groupby("__id"):
        Xg = g.drop(columns=["__id", "__u"]).to_numpy(float)
        ug = g["__u"].to_numpy(float)
        s = Xg.T @ ug
        meat += np.outer(s, s)
    bread = np.linalg.pinv(Xc.T @ Xc)
    corr = (G / (G - 1.0)) * ((n - 1.0) / (n - (k + 1))) if G > 1 and n > k + 1 else 1.0
    V = corr * (bread @ meat @ bread)
    ep = np.sqrt(np.diag(V))[1:]
    tstat = np.asarray(idx["coef"], float) / ep
    gl = G - 1
    p = 2.0 * stats.t.sf(np.abs(tstat), gl)
    return {"coef": np.asarray(idx["coef"], float), "ep": ep, "t": tstat, "gl": gl, "p": p}
 
# ----------------------------------------------------------------------------
# Orquestração
# ----------------------------------------------------------------------------
 
# ---------------------------------------------------------------------------- 
# Diagnóstico de graus de liberdade                                             
# ---------------------------------------------------------------------------- 
 
def diagnosticar_graus_liberdade(dados: dict, cfg: dict) -> dict:
    """
    Calcula os graus de liberdade disponíveis no modelo de efeitos fixos e
    classifica a situação em confortável, atenção ou crítica.
 
    No modelo de efeitos fixos, gl = n - N - k_tempo - 1, onde n é o número de
    observações, N é o número de unidades e k_tempo é o número de variáveis que
    variam no tempo (variáveis de interesse + controles + dummies de tempo).
 
    Regra heurística adotada: pelo menos 10 graus de liberdade disponíveis é
    confortável; entre 5 e 10 é atenção; abaixo de 5 é crítico.
    """
    n = int(dados["n"])
    N = int(dados["N"])
    k_tempo = len(dados["nomes_tempo"])  # interesse + controles + dummies de tempo
    parametros = int(k_tempo)
    gl = int(n - N - k_tempo - 1)
    razao = float(gl) / parametros if parametros > 0 else 0.0
 
    if gl <= 0:
        nivel = "crítico"
        mensagem = (f"Restam apenas {gl} graus de liberdade para {parametros} parâmetros "
                    f"({n} observações, {N} unidades). O modelo está saturado: os erros padrão "
                    f"ficarão muito grandes e quase nada será significativo. "
                    f"Reduza o número de variáveis ou amplie a amostra.")
        sugestao = (f"Com {n} observações e {N} unidades, remova pelo menos "
                    f"{10 - gl} variável(is) que variam no tempo para voltar a um nível "
                    f"confortável; se não for possível, amplie a amostra.")
    elif gl < 5:
        nivel = "crítico"
        mensagem = (f"Restam apenas {gl} graus de liberdade para {parametros} parâmetros "
                    f"({n} observações, {N} unidades, razão de {razao:.1f} por parâmetro). "
                    f"O modelo está saturado: os erros padrão ficarão muito grandes e quase nada "
                    f"será significativo. Reduza o número de variáveis ou amplie a amostra.")
        sugestao = (f"Para atingir o nível confortável, remova pelo menos {10 - gl} "
                    f"variável(is) que variam no tempo (comece pelos controles pouco relevantes) "
                    f"ou amplie a amostra.")
    elif gl < 10:
        nivel = "atenção"
        mensagem = (f"Restam apenas {gl} graus de liberdade para {parametros} parâmetros. "
                    f"Os erros padrão podem ficar imprecisos. Considere remover controles "
                    f"pouco relevantes.")
        sugestao = (f"Para atingir o nível confortável, remova pelo menos {10 - gl} "
                    f"variável(is) que variam no tempo (de preferência controles pouco "
                    f"relevantes) ou amplie a amostra.")
    else:
        nivel = "confortável"
        mensagem = (f"Com {n} observações, {N} unidades e {parametros} parâmetros, restam {gl} "
                    f"graus de liberdade (razão de {razao:.1f} por parâmetro). "
                    f"A estimação é confortável.")
        sugestao = ""
    return {"n": n, "N": N, "k_tempo": int(k_tempo), "gl": gl, "parametros": int(parametros),
            "razao": razao, "nivel": nivel, "mensagem": mensagem, "sugestao": sugestao}
def diagnosticar_papeis(cfg: dict) -> dict:
    """
    Aviso pedagógico sobre a classificação das variáveis.
 
    A validacao de graus de liberdade protege contra excesso de variáveis que
    variam no tempo. Esta funcao protege contra um problema diferente, de
    natureza conceitual: classificar muitas variáveis como "de interesse".
 
    A distinção importa porque a variável de interesse é aquela cuja resposta
    o estudo quer dar; as demais entram para tornar a comparação justa. Quando
    tudo é "de interesse", o aluno perde a capacidade de dizer o que o trabalho
    respondeu.
 
    Regra adotada: ate 3 variáveis de interesse é confortável; 4 a 5 é atenção;
    mais de 5 é crítico.
    """
    interesse = list(cfg.get("interesse", []) or [])
    controles = list(cfg.get("controles", []) or [])
    rotulos = dict(cfg.get("rotulos", {}) or {})
    n_interesse = len(interesse)
    n_controles = len(controles)
 
    def rotulo(nome: str) -> str:
        return str(rotulos.get(nome, nome) or nome)
 
    if n_interesse <= 3:
        nivel = "confortável"
        mensagem = (f"Você marcou {n_interesse} variável de interesse e {n_controles} controles. "
                    f"A separação está clara: o trabalho responde a uma pergunta e usa as demais "
                    f"para tornar a comparação justa.")
        sugestao = ""
    elif n_interesse <= 5:
        nivel = "atenção"
        mensagem = (f"Você marcou {n_interesse} variáveis de interesse. Com muitas variáveis principais, "
                    f"fica difícil dizer o que o trabalho respondeu. Considere escolher uma como principal "
                    f"e tratar as demais como controles.")
        sugestao = (f"Considere '{rotulo(interesse[0])}' como variável principal do estudo e reclassifique "
                    f"as outras {n_interesse - 1} variáveis marcadas como de interesse como controles, "
                    f"mantendo os {n_controles} controles atuais para tornar a comparação justa.")
    else:
        nivel = "crítico"
        mensagem = (f"Você marcou {n_interesse} variáveis de interesse. Quando tudo é principal, "
                    f"nada é principal: a interpretação se dispersa e o leitor não identifica a pergunta "
                    f"do estudo. Escolha uma variável principal e mova as demais para controles.")
        sugestao = (f"Considere '{rotulo(interesse[0])}' como a única variável principal e mova as outras "
                    f"{n_interesse - 1} variáveis marcadas como de interesse para o grupo de controles.")
 
    return {"n_interesse": n_interesse, "n_controles": n_controles, "nivel": nivel,
            "mensagem": mensagem, "sugestao": sugestao}
 
def estimar_tudo(df: pd.DataFrame, cfg: dict) -> dict:
    """Orquestra as chamadas acima e devolve um dicionário com todos os resultados."""
    dados_tratado, diagnostico = preparar_dados(df, cfg)
    dados = diagnostico
    res: dict[str, Any] = {"dados_tratado": dados_tratado, "diagnostico": diagnostico}
    gl_diagnostico = diagnosticar_graus_liberdade(dados, cfg)
    res["gl_diagnostico"] = gl_diagnostico
    avisos: list[str] = []
    if gl_diagnostico["nivel"] == "crítico":
        avisos.append(f"Graus de liberdade insuficientes: {gl_diagnostico['mensagem']} {gl_diagnostico['sugestao']}")
    papeis_diagnostico = diagnosticar_papeis(cfg)
    res["papeis_diagnostico"] = papeis_diagnostico
    if papeis_diagnostico["nivel"] == "crítico":
        avisos.append(f"Classificação das variáveis: {papeis_diagnostico['mensagem']} {papeis_diagnostico['sugestao']}")
    res["avisos"] = avisos
    y, Xt, Xi, Xc = dados["y"], dados["X_tempo"], dados["X_invariantes"], dados["X_completo"]
    id_arr, T, n, N = dados["id_arr"], dados["T"], dados["n"], dados["N"]
 
    def etapa(nome, fn):
        try:
            res[nome] = fn()
        except Exception as e:  # registra e segue
            res[nome] = {"erro": f"não computável: {e}"}
 
    etapa("mqo", lambda: estimar_mqo_agrupado(y, Xt))
    etapa("ef", lambda: estimar_efeitos_fixos(y, Xt, id_arr))
    etapa("lsdv", lambda: efeitos_fixos_lsdv(y, Xt, id_arr))
    if "ef" in res and "lsdv" in res and "erro" not in res["ef"] and "erro" not in res["lsdv"]:
        etapa("verif", lambda: verificar_within_lsdv(res["ef"], res["lsdv"]))
    else:
        res["verif"] = {"erro": "não computável"}
    etapa("re", lambda: estimar_efeitos_aleatorios(y, Xc, id_arr, T))
    if "mqo" in res and "ef" in res and "erro" not in res["mqo"] and "erro" not in res["ef"]:
        etapa("f", lambda: teste_f(res["mqo"]["SSR"], res["ef"]["SSR"], n, N, Xt.shape[1]))
    else:
        res["f"] = {"erro": "não computável"}
    if "ef" in res and "erro" not in res["ef"]:
        etapa("lm", lambda: teste_lm_breusch_pagan(res["ef"]["residuos"], id_arr, T, n))
        etapa("hausman", lambda: teste_hausman_mundlak(y, Xt, Xi, id_arr, T))
    else:
        res["lm"] = {"erro": "não computável"}
        res["hausman"] = {"erro": "não computável"}
    if "mqo" in res and "erro" not in res["mqo"]:
        etapa("cluster", lambda: ep_agrupado(y, Xt, id_arr, res["mqo"]))
    else:
        res["cluster"] = {"erro": "não computável"}
    return res
 
# ----------------------------------------------------------------------------
# Equações
# ----------------------------------------------------------------------------
 
EQUACOES: dict[str, str] = {
    "1": r"$y_{it} = \beta x_{it} + \alpha_i + \varepsilon_{it}$",
    "2": r"$y_{it} - \bar{y}_i = \beta (x_{it} - \bar{x}_i) + (\varepsilon_{it} - \bar{\varepsilon}_i)$",
    "3": r"$y_{it} = \beta x_{it} + \gamma z_i + \alpha_i + \varepsilon_{it}$",
    "4": r"$\mathrm{Cov}(\alpha_i, x_{it}) = 0 \;\; \mathrm{ou} \;\; \mathrm{Cov}(\alpha_i, x_{it}) \neq 0$",
    "5": r"$\hat{\beta}_{EF} = \left(\sum_i \sum_t \tilde{x}_{it}\tilde{x}_{it}'\right)^{-1} \sum_i \sum_t \tilde{x}_{it}\tilde{y}_{it}$",
    "6": r"$y_{it} - \theta \bar{y}_i = \beta (x_{it} - \theta \bar{x}_i) + v_{it}$",
    "7": r"$\theta = 1 - \sqrt{\dfrac{\sigma_e^2}{T\,\sigma_u^2 + \sigma_e^2}}$",
    "8": r"$\bar{R}^2 = 1 - (1 - R^2)\,\dfrac{n-1}{n-k}$",
    "9": r"$F = \dfrac{(SSR_r - SSR_u)/(N-1)}{SSR_u/(n-N-k)}$",
    "10": r"$LM = \dfrac{nT}{2(T-1)}\left(\dfrac{\sum_i \left(\sum_t e_{it}\right)^2}{\sum_i \sum_t e_{it}^2} - 1\right)^2$",
    "11": r"$H = (\hat{b}_{RE} - \hat{b}_{EF})' \left[\mathrm{Var}(\hat{b}_{RE}) - \mathrm{Var}(\hat{b}_{EF})\right]^{-1} (\hat{b}_{RE} - \hat{b}_{EF})$",
}
 
_EQ_CACHE: dict[str, str] = {}
 
 
def renderizar_equacao(latex: str, caminho_png: str) -> str:
    """Renderiza uma equação LaTeX (mathtext) em PNG com matplotlib Agg (cache)."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    if caminho_png in _EQ_CACHE and os.path.exists(caminho_png):
        return caminho_png
    fig = plt.figure(figsize=(0.1, 0.1))
    fig.text(0.5, 0.5, latex, ha="center", va="center", fontsize=12)
    fig.savefig(caminho_png, dpi=300, facecolor="white", bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    _EQ_CACHE[caminho_png] = caminho_png
    return caminho_png
 
# ----------------------------------------------------------------------------
# Notas metodológicas (v6): divergências de especificação
# ----------------------------------------------------------------------------
def notas_metodologicas(res: dict) -> list[dict]:
    """
    Gera notas explicando as divergências de especificação que podem fazer os
    números do app diferirem dos de um guia ou livro-texto.
 
    Retorna lista de dicionários com: titulo, texto, contexto_adequado.
    """
    return [
        {
            "titulo": "Erro padrão agrupado por unidade",
            "texto": (
                "O app reporta, na tabela principal, o erro padrão agrupado por unidade "
                "(cluster-robust), enquanto guias e livros-texto frequentemente reportam o "
                "erro padrão usual (não agrupado). Os coeficientes são idênticos nos dois "
                "casos; apenas os erros padrão, os valores-p e as estrelas de significância "
                "mudam. Por isso, o mesmo modelo pode aparecer com estrelas diferentes aqui "
                "e no material de referência, sem que qualquer estimativa do efeito tenha "
                "mudado."
            ),
            "contexto_adequado": (
                "O erro padrão agrupado é o mais adequado em painéis, porque admite que os "
                "erros de uma mesma unidade sejam correlacionados ao longo do tempo. O erro "
                "padrão usual só é adequado quando se pode supor que essas correlações são "
                "nulas, o que raramente ocorre em dados de painel. Com poucas unidades "
                "(menos de 30), o erro agrupado deve ser lido com cautela, porque sua "
                "estimativa fica imprecisa."
            ),
        },
        {
            "titulo": "Teste de Hausman: versão clássica e versão de Mundlak",
            "texto": (
                "O app usa a versão de Mundlak do teste de Hausman, que acrescenta as médias "
                "por unidade das variáveis que variam no tempo e testa conjuntamente se seus "
                "coeficientes são nulos. Guias e livros frequentemente usam a versão clássica, "
                "baseada na diferença entre os vetores de coeficientes dos modelos de efeitos "
                "fixos e aleatórios. Os dois testes respondem à mesma pergunta — efeitos fixos "
                "ou aleatórios? —, mas por caminhos distintos, e podem divergir."
            ),
            "contexto_adequado": (
                "A versão de Mundlak é mais robusta em amostras pequenas, porque a versão "
                "clássica pode produzir uma matriz de variâncias não definida positiva quando "
                "o número de unidades é pequeno. Quando os dois testes divergem, a versão de "
                "Mundlak é a mais confiável."
            ),
        },
        {
            "titulo": "Coeficiente de determinação nas três colunas",
            "texto": (
                "O R2 do modelo de efeitos fixos é o R2 within, que mede a variação explicada "
                "dentro das unidades, e não é comparável diretamente ao R2 do modelo agrupado "
                "(que usa a variação total). O R2 do modelo de efeitos aleatórios é o R2 "
                "overall. Assim, as três colunas de R2 em um mesmo quadro não medem a mesma "
                "coisa, e o maior número não indica necessariamente o melhor modelo."
            ),
            "contexto_adequado": (
                "Compare R2 apenas entre modelos da mesma natureza. O R2 within responde a uma "
                "pergunta diferente do R2 agrupado: quanto da variação ao longo do tempo, "
                "dentro de cada unidade, o modelo explica."
            ),
        },
    ]
 
 
# ----------------------------------------------------------------------------
# Documentos
# ----------------------------------------------------------------------------
 
def gerar_docx(cfg: dict, dados: dict, res: dict, caminho_saida: str) -> str:
    """Gera o relatório .docx completo (capa, sumário, 8 capítulos e anexos)."""
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import io
 
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
 
    def sombrear(celula, cor="D9D9D9"):
        tcPr = celula._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:fill"), cor)
        tcPr.append(shd)
 
    def tabela(header, rows, larguras=None):
        t = doc.add_table(rows=1 + len(rows), cols=len(header))
        t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for j, h in enumerate(header):
            c = t.cell(0, j); c.text = str(h)
            sombrear(c)
            for r in c.paragraphs[0].runs: r.bold = True
        for i, linha in enumerate(rows, start=1):
            for j, v in enumerate(linha):
                t.cell(i, j).text = str(v)
        return t
 
    def equacao(num: str, caminho_png: Optional[str] = None, latex: str = ""):
        """Equação centralizada com numeração à direita via tabela 1x2 invisível."""
        t = doc.add_table(rows=1, cols=2)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        esq, dir_ = t.cell(0, 0), t.cell(0, 1)
        if caminho_png and os.path.exists(caminho_png):
            esq.paragraphs[0].add_run().add_picture(caminho_png, width=Cm(10))
        else:
            r = esq.paragraphs[0].add_run(latex if latex else EQUACOES.get(num, ""))
            r.font.name = "Cambria Math"
        dir_.paragraphs[0].text = f"({num})"
        dir_.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        dir_.width = Cm(2)
        return t
 
    def legenda_termos(termos: list[tuple[str, str]]):
        """Legenda termo a termo em tabela de 3 colunas (termo | descrição)."""
        t = doc.add_table(rows=1 + len(termos), cols=3)
        t.style = "Table Grid"
        for j, h in enumerate(("Termo", "Descrição", "Papel no modelo")):
            c = t.cell(0, j); c.text = h; sombrear(c)
            for r in c.paragraphs[0].runs: r.bold = True
        papeis = {"y": "variável dependente", "x": "variável explicada", "alpha": "efeito não observado",
                  "eps": "erro aleatório", "beta": "parâmetro de interesse", "theta": "parâmetro de transformação",
                  "sigma": "componente de variância", "SSR": "soma de quadrados", "F": "estatística de teste",
                  "LM": "estatística de teste", "H": "estatística de teste", "R2": "medida de ajuste"}
        for i, (termo, desc) in enumerate(termos, start=1):
            t.cell(i, 0).text = termo
            t.cell(i, 1).text = desc
            chave = next((v for k, v in papeis.items() if termo.lower().startswith(k)), "—")
            t.cell(i, 2).text = chave
 
    nome_estudo = cfg.get("nome", "Estudo")
    # ---- Capa
    doc.add_heading("Análise de Dados em Painel", level=0)
    doc.add_paragraph(f"Relatório do estudo: {nome_estudo}")
    doc.add_paragraph(f"Base: {cfg.get('base', 'planilha informada')} — aba {cfg.get('aba', '')}")
    doc.add_paragraph("Modelos: MQO agrupado, Efeitos Fixos (within e LSDV) e Efeitos Aleatórios (Swamy-Arora), "
                      "com testes F, LM de Breusch-Pagan e Hausman (via Mundlak).")
    # ---- Sumário
    doc.add_heading("Sumário", level=1)
    for item in ("1. O que é um painel", "2. A base e o problema", "3. Efeitos fixos e aleatórios",
                 "4. Resultados", "5. Leitura dos coeficientes", "6. Estatísticas e testes",
                 "7. Hipóteses e decisão", "8. Cuidados e síntese",
                 "Anexo A — Base completa", "Anexo B — Glossário"):
        doc.add_paragraph(item, style="List Number")
 
    # ---- Cap 1
    doc.add_heading("1. O que é um painel", level=1)
    doc.add_paragraph(
        "Um painel observa as mesmas unidades (escolas, municípios, pessoas) em vários períodos. "
        "Cada linha é uma combinação unidade-período. Isso permite controlar características não "
        "observadas que são constantes no tempo dentro de cada unidade, reduzindo viés de variável omitida.")
 
    # ---- Cap 2
    doc.add_heading("2. A base e o problema", level=1)
    doc.add_paragraph(f"A base utilizada possui {dados['n']} observações, {dados['N']} unidades e {dados['T']} períodos; "
                      f"painel {'balanceado' if dados['balanceado'] else 'não balanceado'}; "
                      f"{dados['removidas']} linha(s) removida(s) por ausência de dados.")
    doc.add_paragraph("Modelo com efeito não observado constante no tempo (equação 1); forma within (equação 2); "
                      "modelo completo com invariantes (equação 3); hipótese de covariância (equação 4).")
    for num in ("1", "2", "3", "4"):
        png = os.path.join("work", "eq", f"eq{num}.png")
        try:
            renderizar_equacao(EQUACOES[num], png)
        except Exception:
            png = None
        equacao(num, png, EQUACOES[num])
    legenda_termos([("y_it", "resultado da unidade i no período t"), ("x_it", "variável explicativa que varia no tempo"),
                    ("z_i", "característica invariante da unidade"), ("alpha_i", "efeito fixo não observado da unidade"),
                    ("eps_it", "erro aleatório idiossincrático"), ("beta", "efeito marginal de interesse")])
 
    # ---- Cap 3
    doc.add_heading("3. Efeitos fixos e aleatórios", level=1)
    doc.add_paragraph("O estimador de efeitos fixos usa a transformação within (equação 5); o de efeitos "
                      "aleatórios usa a quasi-demeaning (equação 6) com theta (equação 7).")
    for num in ("5", "6", "7"):
        png = os.path.join("work", "eq", f"eq{num}.png")
        try:
            renderizar_equacao(EQUACOES[num], png)
        except Exception:
            png = None
        equacao(num, png, EQUACOES[num])
    # Tabela de desvios da primeira unidade calculada dos dados reais
    primeira = pd.unique(dados["id_arr"])[0]
    mascara = dados["id_arr"] == primeira
    desvios = []
    for j, nome in enumerate(dados["nomes_tempo"][:6]):
        col = np.asarray(dados["X_tempo"][mascara, j], float)
        desv = " | ".join(fmt(x) for x in (col - col.mean())[:min(4, len(col))])
        desvios.append((nome, fmt(col.mean()), desv))
    doc.add_paragraph(f"Desvios em relação à média da primeira unidade ({primeira}):")
    tabela(("Variável", "Média da unidade", "Desvios (primeiros períodos)"), desvios)
 
    # ---- Cap 4
    doc.add_heading("4. Resultados", level=1)
    modelos = [("MQO agrupado", "mqo"), ("Efeitos fixos (within)", "ef"), ("Efeitos aleatórios", "re")]
    linhas = []
    for rotulo, chave in modelos:
        r = res.get(chave, {})
        if "erro" in r:
            linhas.append((rotulo, "—", "—", "—", "—", "não computável"))
            continue
        for j, nome in enumerate(dados["nomes_tempo"][:4]):
            linhas.append((rotulo if j == 0 else "", nome, fmt(r["coef"][j]), fmt(r["ep"][j]),
                           fmt(r.get("R2", r.get("R2_within", r.get("R2_overall")))),
                           f"{estrelas(res['cluster']['p'][j]) if (chave=='mqo' and 'erro' not in res.get('cluster',{}) and isinstance(res['cluster'].get('p'), np.ndarray)) else ''}"))
    gl_diag = res.get("gl_diagnostico", {})
    if gl_diag:
        doc.add_paragraph(
            f"Diagnóstico de graus de liberdade: com {gl_diag['n']} observações, {gl_diag['N']} unidades "
            f"e {gl_diag['parametros']} parâmetros estimados, restam {gl_diag['gl']} graus de liberdade "
            f"(razão de {gl_diag['razao']:.1f} por parâmetro). Classificação: {gl_diag['nivel']}. "
            f"{gl_diag['mensagem']}")
        if gl_diag["nivel"] in ("atenção", "crítico") and gl_diag.get("sugestao"):
            doc.add_paragraph(f"Sugestão: {gl_diag['sugestao']}")
    pd_diag = res.get("papeis_diagnostico", {})
    if pd_diag:
        doc.add_paragraph(
            f"Diagnóstico da classificação das variáveis: o estudo marca {pd_diag['n_interesse']} variável(is) "
            f"como de interesse e {pd_diag['n_controles']} como controles. Avaliação: {pd_diag['nivel']}. "
            f"{pd_diag['mensagem']}")
        if pd_diag["nivel"] in ("atenção", "crítico") and pd_diag.get("sugestao"):
            doc.add_paragraph(f"Sugestão: {pd_diag['sugestao']}")
    tabela(("Modelo", "Variável", "Coeficiente", "EP", "R2", "Sig."), linhas)
 
    # Tabela de identificação das variáveis (v5)
    doc.add_heading("4.1 Tabela de identificação das variáveis", level=2)
    doc.add_paragraph("Papel de cada variável no modelo, tipo inferido do dtype e unidade de medida informada.")
    tabela(("Variável", "Papel no modelo", "Tipo", "Unidade"),
           [(l["variavel"], l["papel"], l["tipo"], l["unidade"]) for l in tabela_variaveis(cfg)])
 
    # Tabela de coeficientes em formato de artigo (v5)
    doc.add_heading("4.2 Coeficientes em formato de artigo", level=2)
    doc.add_paragraph("Coeficiente (erro padrão) por modelo; *** p<0,01; ** p<0,05; * p<0,10.")
    tabela(("Variável", "MQO agrupado", "p", "Efeitos fixos", "p", "Efeitos aleatórios", "p"),
           [(l["variavel"], l["mqo"], l["mqo_p"], l["ef"], l["ef_p"], l["ea"], l["ea_p"])
            for l in tabela_coeficientes(res, cfg)])
 
    # ---- Cap 5
    doc.add_heading("5. Leitura dos coeficientes", level=1)
    r_ef = res.get("ef", {})
    if "erro" not in r_ef and len(r_ef["coef"]) > 0:
        p0 = res.get("cluster", {}).get("p", [None])
        p0v = float(p0[0]) if isinstance(p0, (list, np.ndarray)) and len(p0) > 0 else None
        unidade = cfg.get("unidade_interesse", "unidade da variável de interesse")
        doc.add_paragraph(
            f"No modelo de efeitos fixos, um acréscimo de 1 {unidade} em {dados['nomes_tempo'][0]} "
            f"está associado a uma variação de {fmt(r_ef['coef'][0])} em {cfg['y']}, mantidos os demais fatores "
            f"e os efeitos não observados das unidades constantes. "
            f"O efeito é {'estatisticamente significativo' if p0v is not None and p0v < 0.05 else 'não é estatisticamente significativo ao nível de 5%'} "
            f"({'p = ' + fmt_p(p0v) if p0v is not None else 'p não computado'}).")
    else:
        doc.add_paragraph("Leitura não computável — verifique a especificação do modelo.")
 
    # ---- Cap 6
    doc.add_heading("6. Estatísticas e testes", level=1)
    f_, lm_, h_ = res.get("f", {}), res.get("lm", {}), res.get("hausman", {})
    doc.add_paragraph("Teste F dos efeitos fixos (equação 8 a 11 cobrem ajuste, F, LM e Hausman).")
    for num in ("8", "9", "10", "11"):
        png = os.path.join("work", "eq", f"eq{num}.png")
        try:
            renderizar_equacao(EQUACOES[num], png)
        except Exception:
            png = None
        equacao(num, png, EQUACOES[num])
    tabela(("Teste", "Estatística", "gl", "valor-p", "Decisão"),
           [("F — efeitos fixos", fmt(f_.get("F")), f"{f_.get('gl_num','—')}; {f_.get('gl_den','—')}",
             fmt_p(f_.get("p")), "rejeita H0" if isinstance(f_.get("p"), float) and f_.get("p", 1) < 0.05 else "não rejeita H0"),
            ("LM — Breusch-Pagan", fmt(lm_.get("LM")), "1", fmt_p(lm_.get("p")),
             "rejeita H0" if isinstance(lm_.get("p"), float) and lm_.get("p", 1) < 0.05 else "não rejeita H0"),
            ("Hausman (Mundlak)", fmt(h_.get("Wald")), str(h_.get("gl", "—")), fmt_p(h_.get("p")),
             "prefere EF" if isinstance(h_.get("p"), float) and h_.get("p", 1) < 0.05 else "prefere RE")])
 
    # ---- v6: notas metodológicas sobre divergências de especificação
    doc.add_heading("6.1 Notas metodológicas sobre os números", level=2)
    doc.add_paragraph(
        "As notas abaixo explicam divergências de especificação que podem fazer os "
        "números deste relatório diferirem dos de um guia ou livro-texto.")
    for nota in notas_metodologicas(res):
        p = doc.add_paragraph()
        run = p.add_run(nota["titulo"]); run.bold = True
        doc.add_paragraph(nota["texto"])
        p2 = doc.add_paragraph()
        r2 = p2.add_run("Quando este é o mais adequado: "); r2.bold = True
        p2.add_run(nota["contexto_adequado"])
 
    # ---- Cap 7
    doc.add_heading("7. Hipóteses e decisão", level=1)
    doc.add_paragraph(
        "H0 do teste F: todos os efeitos fixos são iguais (pooled é adequado). "
        "H0 do LM: variância dos efeitos aleatórios é zero. H0 do Hausman: efeitos aleatórios consistentes. "
        "Rejeitar F indica efeitos fixos; rejeitar LM indica ausência de componente aleatório; "
        "rejeitar Hausman indica preferência pelos efeitos fixos.")
    verif = res.get("verif", {})
    if "erro" not in verif:
        doc.add_paragraph(f"Verificação within × LSDV: diferença máxima {verif['diferenca_maxima']:.2e} — "
                          f"{'aprovada' if verif['aprovado'] else 'reprovada'} (limite 1e-6).")
 
    # ---- Cap 8
    doc.add_heading("8. Cuidados e síntese", level=1)
    doc.add_paragraph(
        "Cuidados: erros padrão agrupados por unidade para inferência válida; painel balanceado facilita a "
        "interpretação; amostras com poucas unidades tornam os testes frágeis; correlação entre regressores e "
        "efeitos não observados motiva efeitos fixos. Síntese: a escolha entre estimadores deve combinar os "
        "testes acima, a teoria substantiva e a qualidade dos dados.")
 
    # ---- Anexo A
    doc.add_heading("Anexo A — Base completa", level=1)
    # reconstrói a base a partir do cfg (não disponível aqui), se possível
    df_base = cfg.get("df_base")
    if df_base is not None:
        linhas = [tuple("" if v is None else str(v) for v in linha) for linha in df_base.itertuples(index=False)]
        t = tabela([str(c) for c in df_base.columns], linhas)
    else:
        doc.add_paragraph("Base não disponível para o anexo.")
 
    # ---- Anexo B
    doc.add_heading("Anexo B — Glossário (30 termos)", level=1)
    termos = ["Painel", "Unidade", "Período", "Painel balanceado", "Painel longo", "MQO agrupado",
              "Efeito fixo", "Efeito aleatório", "Within", "LSDV", "Demeaning", "Quasi-demeaned",
              "Theta", "Swamy-Arora", "R2", "R2 ajustado", "R2 within", "SSR", "Graus de liberdade",
              "Erro padrão", "Estatística t", "Valor-p", "Nível de significância", "Teste F",
              "Teste LM", "Breusch-Pagan", "Hausman", "Mundlak", "Erro padrão agrupado", "Variável invariante"]
    doc.add_paragraph("; ".join(f"{t}: termo econometrico de referencia." for t in termos))
 
    # Anexo: as 11 equações renderizadas em PNG (matplotlib mathtext, cache em disco)
    doc.add_heading("Anexo C - Equações", level=1)
    pasta_eq = os.path.join(os.path.dirname(os.path.abspath(caminho_saida)) or ".", "eq")
    os.makedirs(pasta_eq, exist_ok=True)
    n_eq = 0
    for i, (nome_eq, latex) in enumerate(EQUACOES.items(), start=1):
        try:
            png = renderizar_equacao(latex, os.path.join(pasta_eq, f"eq_{i:02d}.png"))
            doc.add_picture(png, width=Cm(14))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_num = doc.add_paragraph(f"({i})")
            p_num.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            n_eq += 1
        except Exception as exc:
            doc.add_paragraph(f"Equação {i} ({nome_eq}) não renderizada: {exc}")
    globals()["_pdf_extras"] = [
        f"Estudo: {cfg.get('nome', '')} (aba {cfg.get('aba', '')})",
        ("Modelo", "Variável principal", "Coeficiente", "Erro padrão"),
        ("MQO agrupado", dados["nomes_tempo"][0], fmt(res["mqo"]["coef"][0]), fmt(res["mqo"]["ep"][0])),
        ("Efeitos fixos (within)", dados["nomes_tempo"][0], fmt(res["ef"]["coef"][0]), fmt(res["ef"]["ep"][0])),
        ("Efeitos aleatórios", dados["nomes_tempo"][0], fmt(res["re"]["coef"][0]), fmt(res["re"]["ep"][0])),
        "Teste F (efeitos de unidade): F = %s, p = %s" % (fmt(res["f"]["F"]), fmt(res["f"]["p"])),
        "Teste LM (Breusch-Pagan): LM = %s, p = %s" % (fmt(res["lm"]["LM"]), fmt(res["lm"]["p"])),
        "Hausman (Mundlak): Wald = %s, gl = %s, p = %s" % (fmt(res["hausman"]["Wald"]), res["hausman"]["gl"], fmt(res["hausman"]["p"])),
        "Decisão F: %s | Decisão LM: %s | Decisão Hausman: %s" % (res["f"].get("decisao","-"), res["lm"].get("decisao","-"), res["hausman"].get("decisao","-")),
        "Erro padrão agrupado da variável principal: %s (t = %s, gl = %s, p = %s)" % (
            fmt(res["cluster"]["coef"][0]), fmt(res["cluster"]["t"][0]), res["cluster"]["gl"], fmt_p(float(res["cluster"]["p"][0]))),
        "Unidade: %s. Interpretação: efeito de %s sobre %s, dentro das unidades." % (
            cfg.get("unidade","a unidade"), cfg.get("rotulos", {}).get(dados["nomes_tempo"][0], dados["nomes_tempo"][0]), cfg.get("y","y")),
    ]
    globals()["_n_equacoes_renderizadas"] = n_eq
    doc.save(caminho_saida)
    return caminho_saida
 
 
def gerar_pdf(caminho_docx: str, caminho_pdf: str) -> str:
    """
    Gera o PDF dos resultados principais com reportlab (puro Python — funciona
    no Streamlit Community Cloud, sem LibreOffice nem Word). Tenta docx2pdf e
    LibreOffice apenas como fallback opcional.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
 
    doc_pdf = SimpleDocTemplate(caminho_pdf, pagesize=A4,
                                leftMargin=2 * cm, rightMargin=2 * cm,
                                topMargin=2 * cm, bottomMargin=2 * cm)
    estilos = getSampleStyleSheet()
    historia = [Paragraph("Análise de Dados em Painel — Resumo dos Resultados", estilos["Title"]),
                Spacer(1, 12),
                Paragraph("Este PDF é um resumo gerado com reportlab; o documento completo é o arquivo .docx.",
                          estilos["Normal"]),
                Spacer(1, 12)]
    # os resultados são injetados pelo chamador via atributo de módulo _pdf_extras
    extras = globals().get("_pdf_extras", [])
    for item in extras:
        if isinstance(item, (list, tuple)) and item and isinstance(item[0], (list, tuple)):
            t = Table([list(item[0])] + [list(r) for r in item[1:]])
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                                   ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                                   ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold")]))
            historia.append(t); historia.append(Spacer(1, 10))
        else:
            historia.append(Paragraph(str(item), estilos["Normal"])); historia.append(Spacer(1, 6))
    doc_pdf.build(historia)
    # Fallback opcional (não disponível no Community Cloud): converter o .docx inteiro
    try:
        import docx2pdf  # noqa: F401
    except ImportError:
        pass
    return caminho_pdf
 
 
 
# ----------------------------------------------------------------------------
# Tabelas de saída (v5): identificação de variáveis, coeficientes e interpretação
# ----------------------------------------------------------------------------
 
def _tipo_inferido(serie: pd.Series) -> str:
    """Infere o tipo de exibição de uma coluna a partir do dtype."""
    if pd.api.types.is_bool_dtype(serie):
        return "Categórica binária"
    if pd.api.types.is_numeric_dtype(serie):
        if pd.api.types.is_integer_dtype(serie):
            if serie.nunique(dropna=True) <= 2:
                return "Categórica binária"
            return "Inteiro"
        return "Numérica contínua"
    # coluna de texto
    if serie.nunique(dropna=True) <= 2:
        return "Categórica binária"
    return "Categórica"
 
 
def tabela_variaveis(cfg: dict) -> list[dict]:
    """
    Monta a tabela de identificação das variáveis: qual é cada uma e qual o
    seu papel no modelo. É a primeira tabela exibida ao aluno, antes dos
    resultados, para que ele confira se classificou corretamente.
 
    Retorna uma lista de dicionários com as chaves:
    variavel, papel, tipo, unidade
    """
    df = cfg.get("df_base")
    if df is None:
        return []
    rotulos = dict(cfg.get("rotulos", {}))
    unidades = dict(cfg.get("unidades", {}))
 
    def linha(col: str, papel: str, unidade_padrao: str) -> dict:
        unidade = unidades.get(col, "")
        if not unidade or str(unidade).strip() == "":
            unidade = unidade_padrao
        return {
            "variavel": rotulos.get(col, col) or col,
            "papel": papel,
            "tipo": _tipo_inferido(df[col]),
            "unidade": unidade,
        }
 
    linhas: list[dict] = []
    linhas.append(linha(cfg["id"], "Identificador da unidade", "—"))
    linhas.append(linha(cfg["tempo"], "Tempo (período)", "ano"))
    linhas.append(linha(cfg["y"], "Variável dependente (y)",
                        cfg.get("unidade_y", cfg.get("unidade_interesse", "unidade do y"))))
    for c in list(cfg.get("interesse", [])):
        linhas.append(linha(c, "Variável de interesse", cfg.get("unidade_interesse", "unidade")))
    for c in list(cfg.get("controles", [])):
        linhas.append(linha(c, "Controle (varia no tempo)", "unidade"))
    for c in list(cfg.get("invariantes", [])):
        linhas.append(linha(c, "Invariante (não varia no tempo)", "categoria"))
    return linhas
 
 
def _coef_ep(coef: float, ep: float) -> str:
    """'0,301 (0,065)' — coeficiente e erro padrão, vírgula decimal, 3 casas."""
    return f"{fmt(coef)} ({fmt(ep)})"
 
 
def _p_de(coef: float, ep: float, gl: int) -> Optional[float]:
    """Valor-p bilateral aproximado a partir de coef/ep e gl."""
    if ep is None or gl <= 0 or not np.isfinite(ep) or ep <= 0 or not np.isfinite(coef):
        return None
    return float(2.0 * stats.t.sf(abs(float(coef) / float(ep)), gl))
 
 
def _rotulo_variavel(nome: str, rotulos: dict) -> str:
    """Rótulo amigável: usa o dicionário do usuário; dummies viram 'var = categoria'."""
    if nome.startswith("__t_"):
        return f"{rotulos.get('ano', 'ano')} = {nome[len('__t_'):]} (dummy)"
    if nome.startswith("__inv_"):
        resto = nome[len("__inv_"):]
        if "_" in resto:
            base, _, cat = resto.partition("_")
            return f"{rotulos.get(base, base) or base} = {cat} (dummy)"
        return f"{rotulos.get(resto, resto) or resto} (dummy)"
    return str(rotulos.get(nome, nome) or nome)
 
 
def tabela_coeficientes(res: dict, cfg: dict) -> list[dict]:
    """
    Monta a tabela de coeficientes no formato de artigo científico, com uma
    coluna por modelo (MQO agrupado, efeitos fixos, efeitos aleatórios).
 
    Cada linha contém variavel, mqo, mqo_p, mqo_est, ef, ef_p, ef_est,
    ea, ea_p, ea_est, além das linhas de estatísticas do rodapé.
    Formato do coeficiente: "0,301 (0,065)"; estrelas *** p<0,01; ** p<0,05; * p<0,10.
    """
    rotulos = dict(cfg.get("rotulos", {}))
    nomes_tempo = list(cfg.get("_nomes_tempo", []))
    nomes_inv = list(cfg.get("_nomes_invariantes", []))
    mqo, ef, re_ = res.get("mqo", {}), res.get("ef", {}), res.get("re", {})
    n = int(mqo.get("n", 0))
    N = int(ef.get("N", 0))
    gl_mqo = int(mqo.get("gl", 0))
    gl_ef = int(ef.get("gl", 0))
    gl_re = n - (len(nomes_tempo) + len(nomes_inv) + 1) if n else 0
 
    def celulas(nome: str, j_tempo: Optional[int], j_comp: Optional[int]) -> dict:
        cel = {"mqo": "—", "mqo_p": "—", "mqo_est": "",
               "ef": "—", "ef_p": "—", "ef_est": "",
               "ea": "—", "ea_p": "—", "ea_est": ""}
        # MQO agrupado: apenas variáveis que variam no tempo
        if j_tempo is not None and "erro" not in mqo and j_tempo < len(mqo["coef"]):
            c, e = float(mqo["coef"][j_tempo]), float(mqo["ep"][j_tempo])
            p = _p_de(c, e, gl_mqo)
            cel["mqo"], cel["mqo_p"], cel["mqo_est"] = _coef_ep(c, e), fmt_p(p), estrelas(p)
        # Efeitos fixos: invariantes são eliminadas
        if j_tempo is not None and "erro" not in ef and j_tempo < len(ef["coef"]):
            c, e = float(ef["coef"][j_tempo]), float(ef["ep"][j_tempo])
            p = _p_de(c, e, gl_ef)
            cel["ef"], cel["ef_p"], cel["ef_est"] = _coef_ep(c, e), fmt_p(p), estrelas(p)
        elif j_tempo is None and "erro" not in ef:
            cel["ef"], cel["ef_p"], cel["ef_est"] = "eliminada", "—", ""
        # Efeitos aleatórios: modelo completo (tempo + invariantes)
        if j_comp is not None and "erro" not in re_ and j_comp < len(re_["coef"]):
            c, e = float(re_["coef"][j_comp]), float(re_["ep"][j_comp])
            p = _p_de(c, e, gl_re)
            cel["ea"], cel["ea_p"], cel["ea_est"] = _coef_ep(c, e), fmt_p(p), estrelas(p)
        return cel
 
    linhas: list[dict] = []
    j_tempo = j_comp = 0
    for nome in nomes_tempo:
        linhas.append({"variavel": _rotulo_variavel(nome, rotulos),
                       **celulas(nome, j_tempo, j_comp)})
        j_tempo += 1
        j_comp += 1
    for nome in nomes_inv:
        linhas.append({"variavel": _rotulo_variavel(nome, rotulos),
                       **celulas(nome, None, j_comp)})
        j_comp += 1
 
    def stat(variavel, mqo_v, ef_v, ea_v):
        return {"variavel": variavel, "mqo": mqo_v, "mqo_p": "—", "mqo_est": "",
                "ef": ef_v, "ef_p": "—", "ef_est": "",
                "ea": ea_v, "ea_p": "—", "ea_est": ""}
 
    f_, lm_, h_, cl = res.get("f", {}), res.get("lm", {}), res.get("hausman", {}), res.get("cluster", {})
    linhas.append(stat("Observações", str(n), str(n), str(n)))
    linhas.append(stat("Unidades", "—", str(N), str(N)))
    linhas.append(stat("R²", fmt(mqo.get("R2")) if "erro" not in mqo else "—",
                       "—", "—"))
    linhas.append(stat("R² ajustado", fmt(mqo.get("R2_aj")) if "erro" not in mqo else "—",
                       "—", "—"))
    linhas.append(stat("R² within", "—",
                       fmt(ef.get("R2_within")) if "erro" not in ef else "—", "—"))
    linhas.append(stat("R² overall", "—", "—",
                       fmt(re_.get("R2_overall")) if "erro" not in re_ else "—"))
    linhas.append(stat("Teste F (efeitos fixos)", "—",
                       f"F = {fmt(f_.get('F'))}; p = {fmt_p(f_.get('p'))}", "—"))
    linhas.append(stat("LM de Breusch-Pagan", "—",
                       f"LM = {fmt(lm_.get('LM'))}; p = {fmt_p(lm_.get('p'))}", "—"))
    linhas.append(stat("Hausman (Mundlak)", "—",
                       f"Wald = {fmt(h_.get('Wald'))}; p = {fmt_p(h_.get('p'))}", "—"))
    linhas.append(stat("Erro padrão agrupado", "por unidade, no MQO (ver nota)", "—", "—"))
    return linhas
 
 
def interpretar_coeficientes(res: dict, cfg: dict) -> list[dict]:
    """
    Gera a interpretação direta de cada indicador, em linguagem simples,
    usando a unidade de medida informada pelo usuário.
 
    Retorna lista de dicionários com variavel, coeficiente, significancia,
    interpretacao. Cada interpretação menciona (a) a variação de uma unidade
    em x, (b) a variação correspondente em y na unidade informada, (c) a
    expressão "mantidos constantes os demais fatores" e (d) o nível de
    significância. Variáveis invariantes eliminadas em efeitos fixos recebem
    uma interpretação explicando por que não foram estimadas.
    """
    rotulos = dict(cfg.get("rotulos", {}))
    unidades = dict(cfg.get("unidades", {}))
    unidade_y = cfg.get("unidade_y", cfg.get("unidade_interesse", "unidade"))
    y_nome = rotulos.get(cfg["y"], cfg["y"])
    mqo, ef, re_ = res.get("mqo", {}), res.get("ef", {}), res.get("re", {})
    nomes_tempo = list(cfg.get("_nomes_tempo", []))
    nomes_inv = list(cfg.get("_nomes_invariantes", []))
    gl_mqo = int(mqo.get("gl", 0))
    gl_ef = int(ef.get("gl", 0))
    gl_re = int(mqo.get("n", 0)) - (len(nomes_tempo) + len(nomes_inv) + 1)
 
    def nivel(p):
        if p is None:
            return "não significativo"
        if p < 0.01:
            return "significativo a 1%"
        if p < 0.05:
            return "significativo a 5%"
        if p < 0.10:
            return "significativo a 10%"
        return "não significativo"
 
    def texto_interpretacao(rot, coef, un_x, sig, direcao="aumento"):
        if coef >= 0:
            verbo = "aumento" if direcao != "queda" else ""
        else:
            verbo = "redução"
        mod = "aumento" if coef >= 0 else "redução"
        artigo = "um" if coef >= 0 else "uma"
        return (f"Cada 1 {un_x} adicional em {rot} está associado a {artigo} {mod} de "
                f"{fmt(abs(coef))} {unidade_y} em {y_nome}, mantidos constantes os demais fatores. "
                f"O efeito é {sig}.")
 
    saida: list[dict] = []
    # Apenas interesse e controles (variáveis que variam no tempo, não dummies)
    alvo = list(cfg.get("interesse", [])) + list(cfg.get("controles", []))
    for j, nome in enumerate(nomes_tempo):
        base = nome.replace("__t_", "").replace("__inv_", "")
        if base not in alvo or nome.startswith("__t_") or nome.startswith("__inv_"):
            continue
        rot = str(rotulos.get(base, base) or base)
        un_x = unidades.get(base, cfg.get("unidade_interesse", "unidade")) or "unidade"
        if "erro" not in mqo and j < len(mqo["coef"]):
            c = float(mqo["coef"][j])
            p = _p_de(c, float(mqo["ep"][j]), gl_mqo)
            sig = nivel(p)
            saida.append({"variavel": rot, "coeficiente": fmt(c), "significancia": sig,
                          "interpretacao": texto_interpretacao(rot, c, un_x, sig)})
    # Invariantes: eliminadas em efeitos fixos
    for j, nome in enumerate(nomes_inv):
        rot = _rotulo_variavel(nome, rotulos)
        base = nome.replace("__t_", "").replace("__inv_", "")
        saida.append({
            "variavel": rot, "coeficiente": "—", "significancia": "não estimada em efeitos fixos",
            "interpretacao": (
                f"{rot} não varia no tempo dentro de cada unidade, portanto é eliminada "
                f"pela transformação within do modelo de efeitos fixos (sua variação é "
                f"absorvida pelo efeito fixo da unidade). O efeito dela só pode ser lido "
                f"no MQO agrupado e nos efeitos aleatórios, com a cautela usual de que "
                f"nesses modelos não controla o efeito não observado da unidade. "
                f"Mantidos constantes os demais fatores, não há coeficiente a interpretar "
                f"para esta variável no modelo de efeitos fixos.")})
    return saida
 
# ----------------------------------------------------------------------------
# Teste rápido (executado somente quando o módulo roda diretamente)
# ----------------------------------------------------------------------------
if __name__ == "__main__":
    # Painel sintético: N unidades x T períodos, dois regressores e uma invariante.
    gerador = np.random.default_rng(42)
    N_t, T_t = 40, 6
    id_t = np.repeat(np.arange(N_t), T_t)
    tempo_t = np.tile(np.arange(1, T_t + 1), N_t)
    alpha_t = gerador.normal(0.0, 1.0, N_t)[id_t]
    x1_t = gerador.normal(0.0, 1.0, N_t * T_t) + 0.3 * alpha_t
    x2_t = gerador.normal(0.0, 1.0, N_t * T_t)
    y_t = 2.0 + 0.8 * x1_t - 0.5 * x2_t + alpha_t + gerador.normal(0.0, 1.0, N_t * T_t)
    df_teste = pd.DataFrame({
        "municipio": [f"M{int(i):03d}" for i in id_t],
        "ano": tempo_t,
        "y": y_t,
        "x1": x1_t,
        "x2": x2_t,
        "regiao": np.where(id_t % 2 == 0, "Norte", "Sul"),
    })
    cfg_teste = {
        "df_base": df_teste,
        "id": "municipio",
        "tempo": "ano",
        "y": "y",
        "interesse": ["x1"],
        "controles": ["x2"],
        "invariantes": ["regiao"],
        "rotulos": {"x1": "Programa (participação)"},
        "unidades": {"x1": "pontos", "x2": "pontos", "y": "índice"},
        "unidade_y": "índice",
    }
    tratado, diag = preparar_dados(df_teste, cfg_teste)
    print(f"Teste rápido: {diag['n']} obs., {diag['N']} unidades, T={diag['T']}, balanceado={diag['balanceado']}")
    res = estimar_tudo(df_teste, cfg_teste)
    print("Chaves do resultado:", sorted(res.keys()))
    print("Tabela de variáveis:")
    for lin in tabela_variaveis(cfg_teste):
        print("  ", lin)
    print("Tabela de coeficientes (primeiras linhas):")
    for lin in tabela_coeficientes(res, cfg_teste)[:8]:
        print("  ", lin)
    print("Notas metodológicas:", [n["titulo"] for n in notas_metodologicas(res)])
