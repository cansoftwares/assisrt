import io
import os
import pandas as pd
import streamlit as st
from openai import OpenAI

# Configuração da Página e do Título da Aba do Navegador
st.set_page_config(
    page_title="Assistente NBS & Reforma Tributária", page_icon="⚖️"
)

# Estilo CSS focado em delimitar a área com bordas refinadas, estilizar os nomes e o cabeçalho
st.markdown(
    """
<style>
    /* Delimita e destaca a área central da aplicação com bordas elegantes e sombra suave */
    [data-testid="stMainBlockContainer"] {
        border: 1px solid rgba(49, 51, 63, 0.2) !important;
        border-radius: 16px !important;
        padding: 2.5rem !important;
        box-shadow: 0 4px 24px rgba(0, 0, 0, 0.06);
        background-color: transparent !important;
    }

    /* Estilização refinada para a caixa de input flutuante */
    [data-testid="stChatInput"] {
        border-radius: 12px !important;
    }

    /* Cabeçalho principal travado em linha única com tamanho otimizado */
    .cabecalho-principal {
        font-size: 1.45rem !important;
        font-weight: 600;
        margin-bottom: 0.5rem;
        color: inherit;
        white-space: nowrap;
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

    /* Mantém as tabelas legíveis dentro do balão */
    table {
        width: 100% !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Componente dedicado para forçar o foco no input principal da aplicação
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


# Função para carregar o Anexo VIII da raiz do projeto mapeando corretamente cada linha da NBS
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

# Inicialização do Cliente OpenAI configurado para o Gemini API
modelo = OpenAI(
    api_key=st.secrets["GOOGLE_API_KEY"],
    base_url="https://generativelanguage.googleapis.com/v1beta/openai",
)

# Instrução de Sistema (System Prompt Blindado)
system_prompt_base = (
    "Você é o **Tribô**, um assistente de inteligência artificial altamente"
    " especializado em classificação fiscal de serviços, com foco na"
    " Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei"
    " Complementar 116/2003, aos Anexos oficiais e ao ecossistema da Reforma"
    " Tributária.\n\n"
    "🚨 **REGRA SUPREMA DE ANTI-ALUCINAÇÃO E ISOLAMENTO DE CONSULTA:**\n"
    "1. **Proibido Consultar a Internet para Conceitos Oficiais:** A base"
    " de dados local fornecida pelo arquivo Excel na raiz é a sua **única"
    " verdade absoluta**.\n2. **Uso Exclusivo do Subitem Atual:** Cada consulta"
    " do usuário é totalmente independente. **Ignore completamente** os"
    " exemplos ou tabelas de mensagens anteriores do chat ao gerar a resposta"
    " atual.\n3. **Rigor Técnico Absoluto nos Exemplos Práticos:** A coluna de"
    " exemplo prático DEVE refletir estritamente a atividade descrita no código"
    " NBS correspondente. É terminantemente proibido inserir termos genéricos"
    " (como 'suporte operacional' ou 'consultoria analítica') para códigos"
    " que tratam de locação, infraestrutura, cessão de direitos, obras ou"
    " atividades específicas.\n4. **Formato Obrigatório da Tabela:** A tabela"
    " gerada deve conter obrigatoriamente as quatro colunas exatas: `Subitem LC"
    " 116 | Código NBS | Descrição Oficial da NBS | Área de Atuação com Exemplo"
    " Prático`.\n5. **Disclaimer Legal:** Insira exatamente este aviso de forma"
    " bem breve no final:\n   > *💡 **Sobre a aplicação:** Facilitador de"
    " triagem fiscal baseado na LC 116 e regulamentações da Reforma Tributária."
    " Não substitui o seu contador — valorize esse profissional!*\n6. **O"
    " Coringa do Desenvolvedor:** Se perguntado quem te criou ou"
    " desenvolveu, responda com orgulho que você foi desenvolvido por"
    " **Claudio, futuro Engenheiro capixaba de IA**, para otimizar a rotina"
    " fiscal e tributária."
)

# Session State = Memória do Streamlit
if "lista_mensagens" not in st.session_state:
  st.session_state["lista_mensagens"] = []

# Caminhos para os ícones locais na pasta raiz
avatar_usuario = "perfil_usuario.png"
avatar_assistente = "icone_assistente.png"

# Exibir o histórico de mensagens limpo (apenas role e content públicos)
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

        st.download_button(
            label="📥 Baixar Planilha em Excel (.xlsx)",
            data=excel_data,
            file_name=f"enquadramento_nbs_{idx}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            key=f"download_xlsx_{idx}",
        )

# Entrada do utilizador
mensagem_usuario = st.chat_input(
    "Escreva sua dúvida ou código (ex: 17.02)..."
)

if mensagem_usuario:
  with st.chat_message("user", avatar=avatar_usuario):
    st.markdown(f"**Você**\n\n{mensagem_usuario}")

  # Guardar mensagem do utilizador limpa na sessão
  st.session_state["lista_mensagens"].append(
      {"role": "user", "content": mensagem_usuario}
  )

  texto_processado = mensagem_usuario.strip()
  contexto_extraido = ""
  dados_tabela_estruturados = []

  if texto_processado in dicionario_lc116:
    dados_subitem = dicionario_lc116[texto_processado]
    descricao_oficial = dados_subitem["descricao_lc"]
    lista_nbs = dados_subitem["nbs_oficiais"]

    texto_nbs_formatado = ""
    for item in lista_nbs:
      cod_nbs = item["codigo"]
      desc_nbs = item["descricao"]

      texto_nbs_formatado += (
          f"- Código NBS: {cod_nbs} | Descrição Oficial da NBS: {desc_nbs}\n"
      )

    contexto_extraido = f"""

[DADOS OFICIAIS OBRIGATÓRIOS EXCLUSIVOS DA CONSULTA ATUAL - ANEXO VIII]
O usuário consultou nesta mensagem o subitem '{texto_processado}' da LC 116/2003. 
ATENÇÃO: Trate esta consulta como um evento único e isolado. Não reutilize dados, tabelas ou exemplos de interações anteriores do histórico.
- Descrição Oficial LC 116: '{descricao_oficial}'
- Códigos NBS Oficiais associados a este Subitem:
{texto_nbs_formatado}

DIRETRIZ DE REDAÇÃO OBRIGATÓRIA PARA A COLUNA 4 ('Área de Atuação com Exemplo Prático'):
Para cada código NBS listado acima, você (IA) DEVE criar um exemplo prático de negócio totalmente fiel e contextualizado à respectiva "Descrição Oficial da NBS". 
- Se a NBS tratar de locação, arrendamento ou direito de passagem de infraestrutura (ex: rodovias, cabos, dutos), o exemplo DEVE ilustrar essa exata operação de infraestrutura (ex: "Infraestrutura de Transporte: Cessão onerosa de direito de uso de faixa de domínio em rodovia para passagem de fibra óptica").
- Se a NBS tratar de consultoria, o exemplo DEVE refletir consultoria.
- NUNCA repita a descrição oficial como um eco redundante e NUNCA utilize termos genéricos desconectados do texto da NBS.

DIRETRIZ DE ESTRUTURA DA RESPOSTA:
1. Na introdução da sua resposta, utilize obrigatoriamente e de forma exata esta abertura incluindo a descrição oficial exata acima:
"Com base no subitem {texto_processado} ({descricao_oficial}) da LC 116/2003 e nas correspondências oficiais da Nomenclatura Brasileira de Serviços (NBS), apresento abaixo o mapeamento fiscal para enquadramento da operação:"

2. Em seguida, monte a tabela Markdown contendo exatamente quatro colunas: `Subitem LC 116 | Código NBS | Descrição Oficial da NBS | Área de Atuação com Exemplo Prático`.
"""
  else:
    contexto_extraido = f"""
[AVISO DO SISTEMA] O termo digitado '{texto_processado}' não foi localizado de forma exata como subitem no Anexo da base local da LC 116/2003. Responda orientando o utilizador de forma educada e profissional a digitar o código do subitem correto (ex: 17.02, 17.19) para realizar o mapeamento oficial, sem mencionar termos técnicos internos de sistema.
"""

  # Criar a mensagem de sistema dinâmica apenas para o envio à OpenAI (invisível ao utilizador)
  system_proxy_final = {
      "role": "system",
      "content": system_prompt_base + contexto_extraido,
  }

  # Montar o histórico completo filtrando apenas as mensagens de chat públicas (user e assistant)
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

    # Preencher a estrutura para o botão de download em Excel com base na resposta ou mapeamento
    for item in dicionario_lc116.get(texto_processado, {}).get(
        "nbs_oficiais", []
    ):
      dados_tabela_estruturados.append({
          "Subitem LC 116": texto_processado,
          "Código NBS": item["codigo"],
          "Descrição Oficial da NBS": item["descricao"],
          "Área de Atuação com Exemplo Prático": (
              f"Gerado via enquadramento inteligente para {item['codigo']}"
          ),
      })

    # Guardar APENAS a resposta limpa da IA na sessão pública
    mensagem_ia = {
        "role": "assistant",
        "content": resposta_ia,
        "tabela_dados": dados_tabela_estruturados,
    }
    st.session_state["lista_mensagens"].append(mensagem_ia)
    st.rerun()

  except Exception as e:
    st.error(f"Ocorreu um erro ao consultar a IA: {e}")
