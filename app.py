# -*- coding: utf-8 -*-
"""
app.py — Interface Streamlit para análise de dados em painel.

Executa localmente com: streamlit run app.py
"""
import io

import pandas as pd
import streamlit as st

import core as pc

# Configuração da página
st.set_page_config(page_title="Análise de Dados em Painel", page_icon="📈", layout="wide")

# Cabeçalho
st.title("Análise de Dados em Painel")
st.caption("Estimadores de efeitos fixos, aleatórios e MQO agrupado, com testes e relatório em Word e PDF.")

# --- Barra lateral: upload e controles
with st.sidebar:
    st.header("Dados")
    arquivo = st.file_uploader("Carregue a planilha (.xlsx/.xls)", type=["xlsx", "xls"])
    aba = None
    if arquivo is not None:
        abas = pc.ler_planilha(arquivo)
        if len(abas) > 1:
            aba = st.selectbox("Escolha a aba", list(abas.keys()))
        else:
            aba = list(abas.keys())[0]
    if st.button("Recomeçar"):
        st.session_state.clear()
        st.rerun()

if arquivo is None or aba is None:
    st.info("Carregue uma planilha na barra lateral para começar.")
    st.stop()

df = abas[aba]

# --- Prévia e métricas
st.subheader(f"Prévia da aba: {aba}")
st.dataframe(df.head(10), use_container_width=True)
colunas = list(df.columns)
numericas = [c for c in colunas if pd.api.types.is_numeric_dtype(df[c])]
st.metric("Linhas", df.shape[0])
c1, c2, c3, c4 = st.columns(4)
c1.metric("Colunas", df.shape[1])
c2.metric("Colunas numéricas", len(numericas))

# Estado do formulário em session_state (não reestimar a cada interação)
if "estimado" not in st.session_state:
    st.session_state["estimado"] = None
if "resultado" not in st.session_state:
    st.session_state["resultado"] = None

with st.form("config"):
    st.subheader("Configuração do painel")
    c1, c2, c3 = st.columns(3)
    cfg_id = c1.selectbox("Coluna de identificação (unidade)", colunas, index=0)
    cfg_tempo = c2.selectbox("Coluna de tempo", colunas, index=min(2, len(colunas) - 1))
    cfg_y = c3.selectbox("Variável dependente (y)", numericas, index=len(numericas) - 1 if numericas else 0)
    cfg_interesse = st.multiselect("Variáveis de interesse", numericas)
    cfg_controles = st.multiselect("Controles (variam no tempo)", [c for c in numericas if c not in cfg_interesse])
    cfg_invariantes = st.multiselect("Invariantes (não variam no tempo)",
                                     [c for c in colunas if c not in (cfg_interesse + cfg_controles + [cfg_y])])
    cfg_nome = st.text_input("Nome do estudo", value="Meu estudo")
    # Rótulos e unidades por variável escolhida
    rotulos, unidades = {}, {}
    for v in cfg_interesse + cfg_controles:
        cc1, cc2 = st.columns(2)
        rotulos[v] = cc1.text_input(f"Rótulo de {v}", value=v, key=f"rot_{v}")
        unidades[v] = cc2.text_input(f"Unidade de {v}", value="", key=f"uni_{v}")
    enviar = st.form_submit_button("Estimar e gerar documento", type="primary")
def mostrar_diagnostico(gl_diag):
    """Exibe o diagnóstico de graus de liberdade com o widget adequado ao nível."""
    if not gl_diag:
        return
    fn = {"confortável": st.info, "atenção": st.warning, "crítico": st.error}.get(gl_diag["nivel"], st.info)
    texto = gl_diag["mensagem"]
    if gl_diag.get("sugestao"):
        texto += " " + gl_diag["sugestao"]
    fn(texto)

def mostrar_papeis(p_diag):
    """Exibe o diagnóstico da classificação das variáveis (segundo aviso)."""
    if not p_diag:
        return
    if p_diag["nivel"] == "confortável":
        st.info(f"**Classificação das variáveis — confortável.** {p_diag['mensagem']}")
    elif p_diag["nivel"] == "atenção":
        st.warning(f"**Classificação das variáveis — atenção.** {p_diag['mensagem']} {p_diag.get('sugestao', '')}")
    else:
        st.warning(f"**Classificação das variáveis — crítico.** {p_diag['mensagem']} {p_diag.get('sugestao', '')}")

# --- Diagnóstico de graus de liberdade em tempo real
erros = []
if enviar:
    papéis = {}
    for papel, cols in (("id", [cfg_id]), ("tempo", [cfg_tempo]), ("y", [cfg_y]),
                        ("interesse", cfg_interesse), ("controles", cfg_controles),
                        ("invariantes", cfg_invariantes)):
        for c in cols:
            papéis.setdefault(c, []).append(papel)
    duplicados = {c: ps for c, ps in papéis.items() if len(ps) > 1}
    if duplicados:
        erros.append(f"Mesma coluna em dois papéis: {duplicados}")
    if not cfg_interesse:
        erros.append("Selecione ao menos uma variável de interesse.")
    if cfg_y not in numericas:
        erros.append("A variável dependente (y) precisa ser numérica.")
    if erros:
        for e in erros:
            st.error(e)

# --- Diagnóstico de papéis em tempo real (segundo aviso, independente do de graus de liberdade)
if enviar and cfg_interesse:
    papeis_tempo_real = pc.diagnosticar_papeis({"interesse": cfg_interesse, "controles": cfg_controles,
                                                "rotulos": rotulos})
    mostrar_papeis(papeis_tempo_real)

if enviar and not erros:
    cfg = {"id": cfg_id, "tempo": cfg_tempo, "y": cfg_y, "interesse": cfg_interesse,
           "controles": cfg_controles, "invariantes": cfg_invariantes, "nome": cfg_nome,
           "aba": aba, "base": getattr(arquivo, "name", "planilha"),
           "unidade_interesse": unidades.get(cfg_interesse[0], "unidade") if cfg_interesse else "unidade",
           "rotulos": dict(rotulos),
           "df_base": df}
    try:
        with st.spinner("Estimando modelos..."):
            dados = pc.preparar_dados(df, cfg)
            res = pc.estimar_tudo(dados, cfg)
            st.session_state["resultado"] = (dados, res, cfg)
            st.session_state["estimado"] = True
    except Exception as exc:
        st.session_state["resultado"] = None
        st.session_state["estimado"] = False
        with st.expander("Detalhes do erro"):
            st.exception(exc)

# --- Exibição dos resultados
if st.session_state.get("resultado"):
    dados, res, cfg = st.session_state["resultado"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unidades (N)", dados["N"])
    c2.metric("Períodos (T)", dados["T"])
    c3.metric("Observações (n)", dados["n"])
    c4.metric("Balanceado", "Sim" if dados["balanceado"] else "Não")

    r_ef = res.get("ef", {})
    if "erro" not in r_ef:
        st.metric(f"Coeficiente principal (EF): {dados['nomes_tempo'][0]}", pc.fmt(r_ef["coef"][0]))
        p0 = res.get("cluster", {}).get("p")
        if isinstance(p0, (list, pd.Series)) and len(p0) > 0:
            st.caption(f"EP agrupado: {pc.fmt(res['cluster']['ep'][0])} | p = {pc.fmt_p(float(p0[0]))} {pc.estrelas(float(p0[0]))}")

    # Diagnóstico de graus de liberdade e verificação within x LSDV
    st.subheader("Diagnóstico de graus de liberdade")
    mostrar_diagnostico(res.get("gl_diagnostico"))
    if res.get("avisos"):
        for av in res["avisos"]:
            st.warning(av)

    # Diagnóstico da classificação das variáveis (segundo aviso, problema conceitual distinto)
    st.subheader("Diagnóstico da classificação das variáveis")
    mostrar_papeis(res.get("papeis_diagnostico"))

    # Tabela de 3 colunas (modelo, coeficiente, EP)
    linhas = []
    for rotulo, chave in (("MQO agrupado", "mqo"), ("Efeitos fixos", "ef"), ("Efeitos aleatórios", "re")):
        r = res.get(chave, {})
        if "erro" in r:
            continue
        for j, nome in enumerate(dados["nomes_tempo"]):
            linhas.append({"Modelo": rotulo if j == 0 else "", "Variável": nome,
                           "Coeficiente": pc.fmt(r["coef"][j]), "EP": pc.fmt(r["ep"][j])})
    if res.get("gl_diagnostico", {}).get("nivel") == "crítico":
        st.warning("Graus de liberdade insuficientes: os erros padrão podem estar inflados "
                   "e os resultados devem ser interpretados com cautela.")
    st.subheader("Tabela de coeficientes")
    st.dataframe(pd.DataFrame(linhas), use_container_width=True)

    # Testes
    st.subheader("Testes")
    f_, lm_, h_ = res.get("f", {}), res.get("lm", {}), res.get("hausman", {})
    decisao_f = "rejeita H0 — efeitos fixos" if isinstance(f_.get("p"), float) and f_.get("p", 1) < 0.05 else "não rejeita H0"
    decisao_lm = "rejeita H0 — há efeito aleatório" if isinstance(lm_.get("p"), float) and lm_.get("p", 1) < 0.05 else "não rejeita H0"
    decisao_h = "prefere efeitos fixos" if isinstance(h_.get("p"), float) and h_.get("p", 1) < 0.05 else "prefere efeitos aleatórios"
    st.write(f"**Teste F** (efeitos fixos): F = {pc.fmt(f_.get('F'))}, gl = {f_.get('gl_num')} ; {f_.get('gl_den')}, p = {pc.fmt_p(f_.get('p'))} — {decisao_f}")
    st.write(f"**LM de Breusch-Pagan**: LM = {pc.fmt(lm_.get('LM'))}, gl = 1, p = {pc.fmt_p(lm_.get('p'))} — {decisao_lm}")
    st.write(f"**Hausman (Mundlak)**: Wald = {pc.fmt(h_.get('Wald'))}, gl = {h_.get('gl')}, p = {pc.fmt_p(h_.get('p'))} — {decisao_h}")

    verif = res.get("verif", {})
    if "erro" not in verif:
        if verif["aprovado"]:
            st.success(f"Verificação within × LSDV aprovada (diferença máxima {verif['diferenca_maxima']:.2e}).")
        else:
            st.warning(f"Verificação within × LSDV reprovada (diferença máxima {verif['diferenca_maxima']:.2e}).")

    # Download buttons gerados em memória
    cdoc, cpdf = st.columns(2)
    try:
        buf_doc = io.BytesIO()
        pc.gerar_docx(cfg, dados, res, buf_doc if False else "/tmp/rel.docx")
        with open("/tmp/rel.docx", "rb") as fh:
            buf_doc.write(fh.read())
        cdoc.download_button("Baixar relatório (.docx)", buf_doc.getvalue(),
                             "relatorio_painel.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as exc:
        with st.expander("Erro ao gerar .docx"):
            st.exception(exc)
    try:
        buf_pdf = io.BytesIO()
        # PDF com reportlab, resultados principais
        pc._pdf_extras = [
            ("Resumo do painel",),
            f"Nome do estudo: {cfg['nome']}",
            f"Unidades: {dados['N']} | Períodos: {dados['T']} | Observações: {dados['n']} | Balanceado: {'sim' if dados['balanceado'] else 'não'}",
            ("Modelo", "Variável", "Coeficiente", "EP"),
            *[(rot, nome, pc.fmt(r["coef"][j]), pc.fmt(r["ep"][j]))
              for rot, chave in (("MQO agrupado", "mqo"), ("Efeitos fixos", "ef"), ("Efeitos aleatórios", "re"))
              if "erro" not in (r := res.get(chave, {}))
              for j, nome in enumerate(dados["nomes_tempo"])],
            ("Teste", "Estatística", "gl", "p", "Decisão"),
            ("F", pc.fmt(f_.get("F")), f"{f_.get('gl_num')};{f_.get('gl_den')}", pc.fmt_p(f_.get("p")), decisao_f),
            ("LM", pc.fmt(lm_.get("LM")), "1", pc.fmt_p(lm_.get("p")), decisao_lm),
            ("Hausman (Mundlak)", pc.fmt(h_.get("Wald")), str(h_.get("gl")), pc.fmt_p(h_.get("p")), decisao_h),
            "Interpretação: os coeficientes de efeitos fixos medem o efeito dentro das unidades, "
            "controlando características não observadas constantes no tempo.",
        ]
        pc.gerar_pdf("/tmp/rel.docx", "/tmp/rel.pdf")
        with open("/tmp/rel.pdf", "rb") as fh:
            buf_pdf.write(fh.read())
        cpdf.download_button("Baixar resumo (.pdf)", buf_pdf.getvalue(), "resumo_painel.pdf", "application/pdf")
    except Exception as exc:
        with st.expander("Erro ao gerar .pdf"):
            st.exception(exc)

st.caption("Painel em formato longo e balanceado é o cenário de referência. Interprete com cautela quando N for pequeno.")
