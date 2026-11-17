# painel-app
Painel Data Analises
## O que é
 
Aplicativo em Streamlit para análise de dados em painel (formato longo, mesmas unidades
acompanhadas ao longo do tempo). O app lê uma planilha Excel, permite escolher as colunas
de unidade, tempo, variável dependente, variáveis de interesse, controles e invariantes,
e estima três modelos: **MQO agrupado**, **Efeitos Fixos** (within e verificação via LSDV)
e **Efeitos Aleatórios** (Swamy-Arora). Calcula o teste **F** de efeitos fixos, o teste
**LM de Breusch-Pagan** e o teste de **Hausman** (via Mundlak), além de erros padrão
agrupados por unidade. Gera relatório completo em **.docx** e resumo em **.pdf** para download.
 
## Diagnósticos pedagógicos
 
A partir da v3, o aplicativo exibe **dois diagnósticos pedagógicos independentes**, antes e depois da estimação:
 
- **Diagnóstico de graus de liberdade** (`diagnosticar_graus_liberdade`): verifica se há observações suficientes em relação ao número de parâmetros do modelo de efeitos fixos (n - N - k > 0). Protege contra excesso de variáveis que variam no tempo.
- **Diagnóstico da classificação das variáveis** (`diagnosticar_papeis`): verifica quantas variáveis foram marcadas como de interesse (até 3: confortável; 4 a 5: atenção; mais de 5: crítico). Protege contra o problema conceitual de tratar tudo como variável principal. Os dois avisos aparecem de forma claramente distinta na interface e, quando o nível é crítico, entram na lista de avisos dos resultados.
 
## Rodar localmente
 
```bash
pip install -r requirements.txt
streamlit run app.py
```
 
O navegador abre automaticamente em http://localhost:8501.
 
## Publicar no Streamlit Community Cloud
 
1. Suba os arquivos (`app.py`, `painel_core.py`, `requirements.txt`) em um repositório **público** no GitHub.
2. Acesse https://share.streamlit.io e faça login com a conta GitHub.
3. Clique em **New app**.
4. Escolha o repositório e o branch (geralmente `main`).
5. Em **Main file path**, informe `app.py`.
6. Clique em **Deploy**.
7. Copie a URL gerada (algo como `https://usuario-app-abc123.streamlit.app`).
 
## Usar no Google Classroom
 
Publique a URL do app como **material** (link externo) na turma. Os alunos acessam o painel
pelo navegador, carregam a planilha da aula e seguem o roteiro de análise.
 
## Limitações
 
- Os dados devem estar em **formato longo** e, preferencialmente, **balanceados** (mesmo número de períodos por unidade).
- Amostras com **poucas unidades** exigem cautela: os erros padrão agrupados e os testes ficam frágeis com N pequeno.
- O **PDF** é gerado com **reportlab** (puro Python) e contém um **resumo** dos resultados (tabela de coeficientes, testes e decisões); ele **não é réplica** do documento .docx completo.
 
 
## Novidade da v3 — Validação de graus de liberdade
 
Antes de estimar, o aplicativo calcula os graus de liberdade do modelo de efeitos fixos (gl = n - N - k_tempo - 1) e classifica a situação em confortavel, atencao ou critico, avisando o aluno quando os erros padrao podem ficar imprecisos ou inflados. O diagnostico tambem entra no relatorio Word (capitulo 4) e na chave 'gl_diagnostico' do resultado.
 
 
## Novidades da v5
 
A tela de resultados agora exibe, antes dos relatórios: (1) a tabela de identificação das variáveis, (2) a tabela de coeficientes em formato de artigo científico e (3) a interpretação dos coeficientes em linguagem simples. Os downloads dos relatórios .docx e .pdf ficam em um expander "Relatórios" ao final.
      
