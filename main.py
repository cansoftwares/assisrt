import io
import json
import os
import pandas as pd
import streamlit as st
from openai import OpenAI

# Configuração da Página e do Título da Aba do Navegador
st.set_page_config(
    page_title="Assistente NBS & Reforma Tributária", page_icon="⚖️"
)

# Estilo CSS otimizado com correções de margem superior e responsividade para o cabeçalho
st.markdown(
    """
<style>
    /* Expande a área central com padding superior seguro para o cabeçalho não encostar na borda */
    [data-testid="stMainBlockContainer"] {
        max-width: 92% !important;
        width: 92% !important;
        border: 1px solid rgba(49, 51, 63, 0.2) !important;
        border-radius: 16px !important;
        padding: 3rem 2.5rem 2.5rem 2.5rem !important;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.06);
        background-color: transparent !important;
        margin: auto !important;
    }

    /* Estilização refinada para a caixa de input flutuante */
    [data-testid="stChatInput"] {
        border-radius: 12px !important;
    }

    /* Cabeçalho principal com espaçamento interno adequado e quebra limpa */
    .cabecalho-principal {
        font-size: 1.45rem !important;
        font-weight: 600;
        margin-top: 0.5rem;
        margin-bottom: 0.75rem;
        color: inherit;
        word-break: break-word;
        line-height: 1.3 !important;
    }

    /* Balões de chat ocupando 100% da largura alinhados */
    .stChatMessage {
        max-width: 100% !important;
        width: 100% !important;
        border-radius: 12px;
        padding: 12px 18px;
        margin-bottom: 12px;
        border: 1px solid rgba(49, 51, 63, 0.1);
    }

    /* Mensagem do Utilizador (Alinhada à Direita com destaque corporativo) */
    [data-testid="stChatMessage-user"] {
        background-color: #f0f2f6 !important; 
        margin-left: auto !important;
        margin-right: 0px !important;
        flex-direction: row-reverse;
    }
    
    /* Inverte a ordem do avatar do utilizador para ficar na direita */
    [data-testid="stChatMessage-user"] > div:first-child {
        flex-direction: row-reverse;
    }

    /* Mensagem do Assistente (Alinhada à Esquerda) */
    [data-testid="stChatMessage-assistant"] {
        background-color: #ffffff !important; 
        margin-left: 0px !important;
        margin-right: auto !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03);
    }

    /* Adaptação e rolagem fluida para tabelas */
    table {
        width: 100% !important;
    }
    
    /* Media Query para Celulares e Telas Pequenas */
    @media (max-width: 768px) {
        [data-testid="stMainBlockContainer"] {
            max-width: 100% !important;
            width: 100% !important;
            padding: 1.5rem 1rem 1rem 1rem !important;
            border: none !important;
            border-radius: 0px !important;
            box-shadow: none !important;
        }
        .cabecalho-principal {
            font-size: 1.2rem !important;
        }
    }
</style>
""",
    unsafe_allow_html=True,
)

# Componente dedicado para forçar o foco inicial no input principal da aplicação ao carregar a página
st.components.v1.html(
    """
    <script>
        function forcarFocoInput() {
            const doc = window.parent.document;
            const chatInput = doc.querySelector('[data-testid="stChatInput"] textarea');
            if (chatInput) {
                chatInput.focus();
                return true;
            }
            return false;
        }

        let tentativas = 0;
        const intervalo = setInterval(function() {
            if (forcarFocoInput() || tentativas > 25) {
                clearInterval(intervalo);
            }
            tentativas++;
        }, 150);
    </script>
""",
    height=0,
)

# Cabeçalho visual principal e subtítulo na medida exata de duas linhas elegantes
st.markdown(
    '<p class="cabecalho-principal">⚖️ Assistente Especialista em NBS e Reforma'
    " Tributária</p>",
    unsafe_allow_html=True,
)
st.markdown(
    "Consulte códigos de serviços da LC 116/2003, descrições normativas"
    " oficiais e correspondências detalhadas de equivalência NBS para o"
    " ecossistema tributário."
)


# Função para carregar o Anexo VIII do Excel mapeando corretamente a base oficial
@st.cache_data
def carregar_base_lc116():
  caminho_excel = (
      "AnexoVIII-CorrelacaoItemNBSIndOpCClassTrib_IBSCBS_V1.00.00.xlsx"
  )
  if os.path.exists(caminho_excel):
    try:
      df = pd.read_excel(caminho_excel, sheet_name="tabela geral", dtype=str)
      df.columns = [str(col).strip() for col in df.columns]

      col_item_lc = df.columns[0]
      col_desc_lc = df.columns[1]
      col_nbs = df.columns[2]
      col_desc_nbs = df.columns[3]

      df[col_item_lc] = df[col_item_lc].ffill()
      df[col_desc_lc] = df[col_desc_lc].ffill()

      base_mapeada = {}
      for _, row in df.iterrows():
        subitem = str(row[col_item_lc]).strip()
        desc_lc = str(row[col_desc_lc]).strip()
        cod_nbs = str(row[col_nbs]).strip()
        desc_nbs = str(row[col_desc_nbs]).strip()

        if subitem and subitem != "nan":
          if subitem not in base_mapeada:
            base_mapeada[subitem] = {
                "descricao_lc": desc_lc,
                "nbs_oficiais": [],
            }

          if cod_nbs and cod_nbs != "nan":
            base_mapeada[subitem]["nbs_oficiais"].append(
                {"codigo": cod_nbs, "descricao": desc_nbs}
            )

      return base_mapeada
    except Exception as e:
      return {}
  return {}


dicionario_lc116 = carregar_base_lc116()

# Montar um sumário estruturado de toda a base oficial para a IA ter contexto completo de termos, subitens e descrições
resumo_base_texto = ""
for subitem_k, info_v in dicionario_lc116.items():
  resumo_base_texto += (
      f"Subitem LC 116: {subitem_k} - Descrição: {info_v['descricao_lc']}\n"
  )
  for nbs_item in info_v["nbs_oficiais"]:
    resumo_base_texto += (
        f"   -> NBS: {nbs_item['codigo']} | Descrição NBS:"
        f" {nbs_item['descricao']}\n"
    )

# Instrução de Sistema com tom em 1ª pessoa do singular
system_prompt_base = (
    "Você é o **Tribô**, um assistente de inteligência artificial altamente"
    " especializado em classificação fiscal de serviços, com foco na"
    " Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei"
    " Complementar 116/2003 e ao ecossistema da Reforma Tributária.\n\n"
    "🚨 **DIRETRIZES CRÍTICAS DE PREENCHIMENTO E ESCOPO:**\n"
    "1. **Restrição Absoluta de Tema:** Você foi criado exclusivamente para"
    " auxiliar em dúvidas sobre a Reforma Tributária, LC 116/2003, NBS e"
    " classificação fiscal de serviços. Se o usuário perguntar sobre assuntos"
    " alheios ao tema (como futebol, política partidária, BBB, entretenimento,"
    " culinária ou qualquer outro assunto fora do escopo profissional), recuse"
    " de forma educada e elegante.\n"
    "2. **Tom em Primeira Pessoa do Singular:** Responda SEMPRE em **primeira"
    " pessoa do singular** (utilize 'identifiquei', 'apresento', 'consultei',"
    " 'analisei', 'encontrei'). É expressamente proibido o uso de pronomes ou"
    " verbos no plural (como 'identificamos', 'apresentamos').\n"
    "3. **Interpretação Ampla e Natural:** O usuário pode fazer perguntas em"
    " linguagem natural (ex: 'qual o nbs para contabilidade?', 'serviços de"
    " informática', 'código 17.02'). Você deve interpretar o tema, buscar na"
    " tabela de referência abaixo o subitem e os códigos NBS mais"
    " compatíveis, e responder de forma fluida e direta.\n"
    "4. **PROIBIÇÃO DE DUPLICAÇÃO NA COLUNA DE EXEMPLO PRÁTICO:** A coluna"
    " `Área de Atuação com Exemplo Prático` **NUNCA** pode ser cópia ou repetição"
    " da coluna 'Descrição NBS'. Enquanto a 'Descrição NBS' traz o texto"
    " normativo oficial, a coluna de **Exemplo Prático** deve descrever um"
    " **caso real de mercado ou operação empresarial concreta** que se encaixe"
    " naquele código.\n"
    "5. **Estrutura Obrigatória da Resposta:** Apresente o texto introdutório"
    " (em 1ª pessoa), a tabela Markdown e, logo abaixo dela, inclua"
    " obrigatoriamente o seguinte parágrafo exato: \n"
    "   *Importante*: A seleção precisa do código NBS é de suma importância"
    " na Reforma Tributária. A correta classificação fiscal garante a"
    " aplicação adequada das regras de incidência, não cumulatividade e"
    " eventuais alíquotas diferenciadas, mitigando riscos de bitributação ou"
    " autuações fiscais. Esta ferramenta atua como um suporte estratégico e"
    " inteligente de alto nível, mas não tem o objetivo de substituir seu"
    " contador — **Valorize esse profissional!**\n"
    "6. **Formato Técnico Oculto:** Forneça a resposta em texto corrido e"
    " formatado para leitura no chat, e ao final insira obrigatoriamente um"
    " bloco de código JSON isolado contendo a lista estruturada exata com a"
    " chave seguinte: \n"
    '     `{"dados_tabela": [{"subitem": "...", "codigo_nbs": "...",'
    ' "descricao_nbs": "...", "exemplo_pratico": "..."}, ...]}`\n'
    "7. **Disclaimer Legal:** Insira o aviso de rodapé padrão no final (sem"
    " citar termos técnicos internos como 'base de dados local').\n"
    "8. **O Coringa do Desenvolvedor:** Se perguntado quem te criou ou"
    " desenvolveu, responda com orgulho que você foi desenvolvido por"
    " **Claudio, futuro Engenheiro capixaba de IA**.\n\n"
    "### TABELA DE REFERÊNCIA OFICIAL (LC 116 / NBS):\n"
    f"{resumo_base_texto}"
)

# Inicialização do Cliente OpenAI configurado para o Gemini API
modelo = OpenAI(
    api_key=st.secrets["GOOGLE_API_KEY"],
    base_url="https://generativelanguage.googleapis.com/v1beta/openai",
)

# Session State = Memória do Streamlit
if "lista_mensagens" not in st.session_state:
  st.session_state["lista_mensagens"] = []

# Caminhos para os ícones locais na pasta raiz
avatar_usuario = "perfil_usuario.png"
avatar_assistente = "icone_assistente.png"

# Controlador booleano para garantir que o script de foco seja injetado imediatamente após renderizar a nova resposta
deve_focar_input = False

# Exibir o histórico de mensagens limpo
for idx, mensagem in enumerate(st.session_state["lista_mensagens"]):
  role = mensagem["role"]
  content = mensagem["content"]

  if role == "user":
    with st.chat_message("user", avatar=avatar_usuario):
      st.markdown(f"**Você**\n\n{content}")
  elif role == "assistant":
    with st.chat_message("assistant", avatar=avatar_assistente):
      st.markdown(
          f"**Tribô – Seu assistente na Reforma Tributária**\n\n{content}"
      )

      tabela_para_baixar = mensagem.get("tabela_dados", [])
      if tabela_para_baixar:
        df_resposta = pd.DataFrame(tabela_para_baixar)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_resposta.to_excel(writer, index=False, sheet_name="Enquadramento")
        excel_data = output.getvalue()

        # Nome profissional atualizado conforme solicitado
        subitem_referencia = mensagem.get("subitem_ref", "Geral")
        nome_arquivo_excel = (
            f"Relatorio_NBS_Inteligente_-_Subitem_{subitem_referencia}.xlsx"
        )

        st.download_button(
            label="📥 Baixar Relatório em Excel (.xlsx)",
            data=excel_data,
            file_name=nome_arquivo_excel,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_xlsx_{idx}",
        )

# Entrada do utilizador
mensagem_usuario = st.chat_input(
    "Escreva sua dúvida ou código (ex: 17.02, contabilidade...)"
)

if mensagem_usuario:
  with st.chat_message("user", avatar=avatar_usuario):
    st.markdown(f"**Você**\n\n{mensagem_usuario}")

  st.session_state["lista_mensagens"].append(
      {"role": "user", "content": mensagem_usuario}
  )

  texto_processado = mensagem_usuario.strip()
  dados_tabela_estruturados = []
  subitem_identificado_cache = "Geral"

  # Verificação flexível de subitem direto
  subitem_encontrado_direto = None
  for sub in dicionario_lc116.keys():
    if sub.lower() in texto_processado.lower():
      subitem_encontrado_direto = sub
      break

  if subitem_encontrado_direto:
    subitem_identificado_cache = subitem_encontrado_direto
    info_sub = dicionario_lc116[subitem_encontrado_direto]
    instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador mencionou diretamente ou indiretamente o subitem '{subitem_encontrado_direto}' ({info_sub['descricao_lc']}).
Gere a resposta em PRIMEIRA PESSOA DO SINGULAR (ex: 'identifiquei', 'apresento'), apresentando a tabela com os códigos NBS oficiais, exemplos práticos reais, o parágrafo humano de valorização do contador e o bloco JSON oculto para o Excel.
"""
  else:
    instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador fez uma consulta em linguagem natural: '{texto_processado}'.
Analise a Tabela de Referência Oficial fornecida acima, identifique em primeira pessoa do singular o(s) subitem(ns) da LC 116/2003 e os códigos NBS mais adequados. Apresente a tabela correspondente, os exemplos práticos reais, finalize com o parágrafo humano de valorização do contador e forneça o bloco JSON oculto. Se o texto for completamente fora do tema de tributação ou LC 116, aplique a diretriz de recusa educada.
"""

  system_proxy_final = {
      "role": "system",
      "content": system_prompt_base + "\n\n" + instrucao_especifica,
  }

  historico_chat = [
      m
      for m in st.session_state["lista_mensagens"]
      if m["role"] in ["user", "assistant"]
  ]
  mensagens_para_ia = [system_proxy_final] + historico_chat

  try:
    resposta_modelo = modelo.chat.completions.create(
        messages=mensagens_para_ia, model="gemini-flash-lite-latest"
    )

    resposta_ia = resposta_modelo.choices[0].message.content

    # Extração inteligente do JSON gerado pela IA para popular o Excel
    try:
      if "```json" in resposta_ia:
        json_str = resposta_ia.split("```json")[1].split("```")[0].strip()
      elif "```" in resposta_ia:
        json_str = resposta_ia.split("```")[1].split("```")[0].strip()
      else:
        json_str = ""

      dados_json = json.loads(json_str)
      if "dados_tabela" in dados_json:
        for item in dados_json["dados_tabela"]:
          sub_val = item.get("subitem", "")
          if sub_val and subitem_identificado_cache == "Geral":
            subitem_identificado_cache = sub_val

          dados_tabela_estruturados.append({
              "Subitem LC 116": sub_val,
              "Código NBS": item.get("codigo_nbs", ""),
              "Descrição Oficial da NBS": item.get("descricao_nbs", ""),
              "Área de Atuação com Exemplo Prático": item.get(
                  "exemplo_pratico", ""
              ),
          })
    except Exception:
      pass

    # Remove o bloco JSON da visualização do chat para mantê-lo limpo e elegante
    if "```json" in resposta_ia:
      resposta_ia_exibicao = resposta_ia.split("```json")[0].strip()
    else:
      resposta_ia_exibicao = resposta_ia

    mensagem_ia = {
        "role": "assistant",
        "content": resposta_ia_exibicao,
        "tabela_dados": dados_tabela_estruturados,
        "subitem_ref": subitem_identificado_cache,
    }
    st.session_state["lista_mensagens"].append(mensagem_ia)

    deve_focar_input = True
    st.rerun()

  except Exception as e:
    st.error(f"Ocorreu um erro ao consultar a IA: {e}")

# Script disparado dinamicamente para garantir que o input recupere o foco logo após a resposta ser exibida
if deve_focar_input:
  st.components.v1.html(
      """
        <script>
            function focarNovamente() {
                const doc = window.parent.document;
                const chatInput = doc.querySelector('[data-testid="stChatInput"] textarea');
                if (chatInput) {
                    chatInput.focus();
                    return true;
                }
                return false;
            }
            let t = 0;
            const iv = setInterval(function() {
                if (focarNovamente() || t > 30) {
                    clearInterval(iv);
                }
                t++;
            }, 100);
        </script>
    """,
      height=0,
  )
