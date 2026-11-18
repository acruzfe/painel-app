# -*- coding: utf-8 -*-
"""
app.py — Interface Streamlit para análise de dados em painel (versão 6).

Propósito: este módulo contém APENAS a interface do app. Ele monta o
formulário de configuração, chama o motor de cálculo (painel_core.py) e exibe
as três tabelas de saída na tela. Os relatórios em Word e PDF ficam recolhidos
em uma seção separada, ao final.

Correção crítica da v6: na v5, a lógica de exclusão de colunas já atribuídas
removia colunas indevidamente e o campo de controles podia mostrar apenas duas
opções. A função opcoes_para corrige isso: cada campo oferece todas as colunas
da base, descontando apenas as atribuídas a OUTROS campos.

Executa localmente com: streamlit run app.py
"""
import io
import os
import tempfile

import pandas as pd
import streamlit as st

import painel_core as pc

VERSAO_APP = "6.0"

# Configuração da página
st.set_page_config(page_title="Análise de Dados em Painel", page_icon=":bar_chart:", layout="wide")


def detectar_colunas_id(df):
    """
    Detecta automaticamente colunas de texto que funcionam como identificador
    ou rótulo (ex.: id_municipio, municipio).

    Uma coluna é tratada como identificadora quando:
      - tem dtype object (texto), e
      - tem poucos valores distintos em relação ao total de linhas (nomes
        repetidos ao longo do tempo, um por unidade do painel).

    Colunas numéricas nunca são consideradas identificadoras.
    """
    colunas_id = []
    for coluna in df.columns:
        if pd.api.types.is_numeric_dtype(df[coluna]):
            continue
        serie = df[coluna].astype(str)
        n_distintos = serie.nunique()
        e_texto = serie.dtype == object
        if e_texto and n_distintos <= df.shape[0]:
            colunas_id.append(coluna)
    return colunas_id


def opcoes_para(campo, cfg_atual, colunas, colunas_id):
    """
    Retorna as colunas disponíveis para um campo do formulário.

    CORREÇÃO CRÍTICA DA v6: na v5, a lógica de exclusão removia colunas
    indevidamente e o campo de controles chegou a mostrar apenas duas opções,
    porque uma coluna atribuída ao próprio campo também era excluída da sua
    própria lista, impedindo o usuário de trocar a seleção.

    Regra: cada campo oferece TODAS as colunas da base, descontando apenas as
    que já foram atribuídas a OUTROS campos. O próprio campo nunca é excluído
    da sua lista. Colunas identificadoras de texto (ex.: id_municipio,
    municipio) são oferecidas apenas nos campos de identificação. A ordem
    original das colunas é preservada. Nunca retorna lista vazia: se a
    filtragem esvaziar a lista, retorna todas as colunas não identificadoras.
    """
    usados_outros = set()
    for outro, valor in cfg_atual.items():
        if outro == campo:
            continue
        if isinstance(valor, (list, tuple)):
            usados_outros.update(valor)
        elif valor is not None:
            usados_outros.add(valor)
    if campo == "identificacao":
        return [c for c in colunas if c in colunas_id]
    disponiveis = [c for c in colunas if c not in colunas_id and c not in usados_outros]
    if not disponiveis:
        disponiveis = [c for c in colunas if c not in colunas_id]
    return disponiveis


def validar_configuracao(cfg):
    """
    Impede a mesma coluna em dois campos do formulário.

    Retorna lista de mensagens de erro; lista vazia significa configuração
    válida. Mensagem exata: "A coluna X já está atribuída ao campo Y.
    Escolha outra coluna ou libere-a primeiro."
    """
    erros = []
    papeis = {}
    pares = (
        ("identificação", [cfg.get("id")]),
        ("tempo", [cfg.get("tempo")]),
        ("variável dependente", [cfg.get("y")]),
        ("variáveis de interesse", list(cfg.get("interesse") or [])),
        ("controles", list(cfg.get("controles") or [])),
        ("invariantes", list(cfg.get("invariantes") or [])),
    )
    for nome_campo, cols in pares:
        for c in cols:
            if c is None:
                continue
            papeis.setdefault(c, []).append(nome_campo)
    for c, campos in papeis.items():
        if len(campos) > 1:
            erros.append(
                f"A coluna {c} já está atribuída ao campo {campos[1]}. "
                "Escolha outra coluna ou libere-a primeiro."
            )
    return erros


def montar_formulario(df):
    """
    Monta os seis campos do formulário de configuração do painel.

    Os campos são: coluna de identificação (unidade), coluna de tempo,
    variável dependente (y), variáveis de interesse (multiselect), controles
    (multiselect) e invariantes (multiselect). Inclui também o nome do estudo,
    a unidade de medida do y e as unidades das variáveis de interesse.

    As listas de opções vêm de opcoes_para, que corrige o bug da v5.
    Retorna o dicionário cfg preenchido e um booleano indicando se o usuário
    clicou em "Estimar".
    """
    colunas = list(df.columns)
    colunas_id = detectar_colunas_id(df)
    cfg_atual = {}
    enviar = False

    with st.form("config"):
        st.subheader("Configuração do painel")
        c1, c2, c3 = st.columns(3)
        cfg_id = c1.selectbox(
            "Coluna de identificação (unidade)",
            opcoes_para("identificacao", cfg_atual, colunas, colunas_id), index=0)
        cfg_atual["identificacao"] = cfg_id
        opts_tempo = opcoes_para("tempo", cfg_atual, colunas, colunas_id)
        cfg_tempo = c2.selectbox("Coluna de tempo", opts_tempo,
                                 index=min(1, len(opts_tempo) - 1))
        cfg_atual["tempo"] = cfg_tempo
        opts_y = opcoes_para("dependente", cfg_atual, colunas, colunas_id)
        cfg_y = c3.selectbox("Variável dependente (y)", opts_y,
                             index=max(0, len(opts_y) - 1))
        cfg_atual["dependente"] = cfg_y
        cfg_interesse = st.multiselect(
            "Variáveis de interesse",
            opcoes_para("interesse", cfg_atual, colunas, colunas_id))
        cfg_atual["interesse"] = list(cfg_interesse)
        cfg_controles = st.multiselect(
            "Controles (variam no tempo)",
            opcoes_para("controles", cfg_atual, colunas, colunas_id))
        cfg_atual["controles"] = list(cfg_controles)
        cfg_invariantes = st.multiselect(
            "Invariantes (não variam no tempo)",
            opcoes_para("invariantes", cfg_atual, colunas, colunas_id))
        cfg_atual["invariantes"] = list(cfg_invariantes)
        cfg_nome = st.text_input("Nome do estudo", value="Meu estudo")
        unidade_y = st.text_input("Unidade de medida do y", value="")
        # Unidade de medida e rótulo de cada variável de interesse
        unidades = {}
        rotulos = {}
        for v in cfg_interesse:
            unidades[v] = st.text_input(f"Unidade de {v}", value="", key=f"uni_{v}")
            rotulos[v] = st.text_input(f"Rótulo de {v}", value=v, key=f"rot_{v}")
        enviar = st.form_submit_button("Estimar e gerar documento", type="primary")

    cfg = {
        "id": cfg_id, "tempo": cfg_tempo, "y": cfg_y,
        "interesse": list(cfg_interesse), "controles": list(cfg_controles),
        "invariantes": list(cfg_invariantes), "nome": cfg_nome,
        "unidade_y": unidade_y, "unidades": dict(unidades),
        "rotulos": dict(rotulos),
    }
    return cfg, enviar


def exibir_tabela_variaveis(cfg):
    """
    Exibe a seção 1. Variáveis do modelo: tabela com o papel de cada
    variável no modelo e uma legenda curta explicando os papéis.
    """
    st.subheader("1. Variáveis do modelo")
    st.caption("Papéis: o identificador marca a unidade; o tempo marca o período; "
               "a variável dependente (y) é o resultado explicado; as variáveis de "
               "interesse respondem à pergunta do estudo; os controles tornam a "
               "comparação justa; as invariantes não variam no tempo dentro da unidade.")
    tv = pc.tabela_variaveis(cfg)
    df_tv = pd.DataFrame(tv).rename(columns={"variavel": "Variável", "papel": "Papel no modelo",
                                             "tipo": "Tipo", "unidade": "Unidade"})
    st.dataframe(df_tv, use_container_width=True, hide_index=True)


def exibir_tabela_coeficientes(res, cfg):
    """
    Exibe a seção 2. Coeficientes estimados: tabela no formato de artigo
    científico, com uma coluna por modelo, e a nota de rodapé com as estrelas.
    """
    st.subheader("2. Coeficientes estimados")
    tc = pc.tabela_coeficientes(res, cfg)
    df_tc = pd.DataFrame(tc).rename(columns={
        "variavel": "Variável", "mqo": "MQO agrupado", "mqo_p": "p", "mqo_est": " ",
        "ef": "Efeitos fixos", "ef_p": "p ", "ef_est": "  ",
        "ea": "Efeitos aleatórios", "ea_p": "p  ", "ea_est": "   "})
    st.dataframe(df_tc, use_container_width=True, hide_index=True,
                 column_config={
                     "Variável": st.column_config.TextColumn("Variável", width="large"),
                     "MQO agrupado": st.column_config.TextColumn("MQO agrupado", width="medium"),
                     "Efeitos fixos": st.column_config.TextColumn("Efeitos fixos", width="medium"),
                     "Efeitos aleatórios": st.column_config.TextColumn("Efeitos aleatórios", width="medium"),
                 })
    st.caption("Erro padrão entre parênteses. *** p<0,01; ** p<0,05; * p<0,10.")


def exibir_notas_metodologicas(res):
    """
    Exibe o expander de notas metodológicas sobre divergências de
    especificação: três notas, cada uma com título em negrito, o texto e
    o campo "Quando este é o mais adequado".
    """
    with st.expander("Notas metodológicas sobre os números"):
        for nota in pc.notas_metodologicas(res):
            st.markdown(f"**{nota['titulo']}**")
            st.markdown(nota["texto"])
            st.markdown(f"**Quando este é o mais adequado:** {nota['contexto_adequado']}")


def exibir_interpretacoes(res, cfg):
    """
    Exibe a seção 3. Interpretação dos coeficientes: um bloco por variável,
    com nome em negrito, coeficiente e parágrafo de interpretação.
    """
    st.subheader("3. Interpretação dos coeficientes")
    for item in pc.interpretar_coeficientes(res, cfg):
        st.markdown(f"**{item['variavel']}** — coeficiente {item['coeficiente']}")
        st.markdown(item["interpretacao"])


def exibir_relatorios(res, cfg, dados):
    """
    Exibe o expander de Relatórios, com os botões de download do .docx e do
    .pdf. Os arquivos são gerados em arquivos temporários e carregados em
    memória com io.BytesIO, sem tocar nos arquivos do usuário.
    """
    with st.expander("Relatórios"):
        cdoc, cpdf = st.columns(2)
        tmp_doc = os.path.join(tempfile.gettempdir(), "relatorio_painel.docx")
        tmp_pdf = os.path.join(tempfile.gettempdir(), "relatorio_painel.pdf")
        try:
            pc.gerar_docx(cfg, dados, res, tmp_doc)
            buf_doc = io.BytesIO()
            with open(tmp_doc, "rb") as fh:
                buf_doc.write(fh.read())
            cdoc.download_button(
                "Baixar relatório (.docx)", buf_doc.getvalue(),
                "relatorio_painel.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as exc:
            with cdoc.expander("Erro ao gerar .docx"):
                st.exception(exc)
        try:
            pc.gerar_pdf(tmp_doc, tmp_pdf)
            buf_pdf = io.BytesIO()
            with open(tmp_pdf, "rb") as fh:
                buf_pdf.write(fh.read())
            cpdf.download_button("Baixar resumo (.pdf)", buf_pdf.getvalue(),
                                 "relatorio_painel.pdf", "application/pdf")
        except Exception as exc:
            with cpdf.expander("Erro ao gerar .pdf"):
                st.exception(exc)


def main():
    """Orquestra o app: upload, formulário, estimação e exibição dos resultados."""
    st.title("Análise de Dados em Painel")
    st.caption("Estimadores de efeitos fixos, aleatórios e MQO agrupado, com testes "
               "e relatório em Word e PDF. Versão " + VERSAO_APP)

    # --- Barra lateral: upload e controles
    with st.sidebar:
        st.header("Dados")
        arquivo = st.file_uploader("Carregue a planilha (.xlsx/.xls)", type=["xlsx", "xls"])
        aba = None
        abas = {}
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
    c1, c2, c3 = st.columns(3)
    c1.metric("Linhas", df.shape[0])
    c2.metric("Colunas", df.shape[1])
    c3.metric("Colunas numéricas",
              len([c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]))

    # --- Formulário e validação
    cfg, enviar = montar_formulario(df)
    if "resultado" not in st.session_state:
        st.session_state["resultado"] = None

    erros = validar_configuracao(cfg)
    if enviar and erros:
        for e in erros:
            st.error(e)
    if enviar and not cfg["interesse"]:
        st.error("Selecione ao menos uma variável de interesse.")
        erros = erros or ["sem interesse"]

    if enviar and not erros and cfg["interesse"]:
        try:
            with st.spinner("Estimando modelos..."):
                cfg["aba"] = aba
                cfg["base"] = getattr(arquivo, "name", "planilha")
                cfg["df_base"] = df
                dados = pc.preparar_dados(df, cfg)
                res = pc.estimar_tudo(dados, cfg)
                st.session_state["resultado"] = (dados, res, cfg)
        except Exception as exc:
            st.session_state["resultado"] = None
            with st.expander("Detalhes do erro"):
                st.exception(exc)

    # --- Exibição dos resultados, na ordem pedida
    if st.session_state.get("resultado"):
        dados, res, cfg = st.session_state["resultado"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Unidades (N)", dados["N"])
        c2.metric("Períodos (T)", dados["T"])
        c3.metric("Observações (n)", dados["n"])
        c4.metric("Balanceado", "Sim" if dados["balanceado"] else "Não")

        st.subheader("Diagnóstico de graus de liberdade")
        gl_diag = res.get("gl_diagnostico")
        if gl_diag:
            fn = {"confortável": st.info, "atenção": st.warning,
                  "crítico": st.error}.get(gl_diag.get("nivel"), st.info)
            fn(gl_diag.get("mensagem", ""))
        p_diag = res.get("papeis_diagnostico")
        if p_diag:
            st.info(f"**Classificação das variáveis.** {p_diag.get('mensagem', '')} "
                    f"{p_diag.get('sugestao', '')}")

        exibir_tabela_variaveis(cfg)
        exibir_tabela_coeficientes(res, cfg)
        exibir_notas_metodologicas(res)
        exibir_interpretacoes(res, cfg)
        exibir_relatorios(res, cfg, dados)


if __name__ == "__main__":
    main()
