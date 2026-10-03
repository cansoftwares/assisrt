import os
import pandas as pd
import streamlit as st
from openai import OpenAI

# Configuração da Página e do Título da Aba do Navegador
st.set_page_config(
    page_title="Assistente NBS & Reforma Tributária", page_icon="⚖️"
)

# Estilo CSS e Script JavaScript para focar automaticamente no chat input ao abrir o app
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

    /* Mensagem do Usuário (Alinhada à Direita com destaque corporativo) */
    [data-testid="stChatMessage-user"] {
        background-color: #f0f2f6 !important; 
        margin-left: auto !important;
        margin-right: 0px !important;
        flex-direction: row-reverse;
    }
    
    /* Inverte a ordem do avatar do usuário para ficar na direita */
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

<script>
    // Foca automaticamente no campo de texto do chat assim que a página carrega
    document.addEventListener("DOMContentLoaded", function() {
        setTimeout(function() {
            const chatInput = document.querySelector('[data-testid="stChatInput"] textarea');
            if (chatInput) {
                chatInput.focus();
            }
        }, 300);
    });
</script>
""",
    unsafe_allow_html=True,
)

# Cabeçalho visual principal e subtítulo na medida exata de duas linhas elegantes
st.markdown(
    '<p class="cabecalho-principal">⚖️ Assistente Especialista em NBS e Reforma'
    " Tributária</p>",
    unsafe_allow_html=True,
)
st.markdown(
    "Consulte códigos de serviços da LC 116/2003, descrições normativas oficiais"
    " e correspondências detalhadas de equivalência NBS para o ecossistema"
    " tributário."
)


# Função para carregar o Anexo VIII da raiz do projeto mapeando múltiplos códigos NBS por subitem
@st.cache_data
def carregar_base_lc116():
  caminho_excel = "AnexoVIII-CorrelacaoItemNBSIndOpCClassTrib_IBSCBS_V1.00.00.xlsx"
  if os.path.exists(caminho_excel):
    try:
      # Lê a aba 'tabela geral' forçando tudo como string para preservar zeros à esquerda (ex: 17.02)
      df = pd.read_excel(caminho_excel, sheet_name="tabela geral", dtype=str)

      # Padroniza os nomes das colunas
      df.columns = [str(col).strip() for col in df.columns]

      col_item_lc = df.columns[0]
      col_desc_lc = df.columns[1]
      col_nbs = df.columns[2]
      col_desc_nbs = df.columns[3]

      # Preenche as células vazias para baixo (forward fill) para resolver o problema das células mescladas em A e B
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


# Carrega o dicionário de subitens estruturado
dicionario_lc116 = carregar_base_lc116()

# Inicialização do Cliente OpenAI configurado para o Gemini API
modelo = OpenAI(
    api_key=st.secrets["GOOGLE_API_KEY"],
    base_url="https://generativelanguage.googleapis.com/v1beta/openai",
)

# Instrução de Sistema (System Prompt Blindado e Baseado nas Normas Oficiais)
system_prompt_base = (
    "Você é o **Tribô**, um assistente de inteligência artificial altamente"
    " especializado em classificação fiscal de serviços, com foco na"
    " Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei Complementar"
    " 116/2003, aos Anexos da regulamentação e ao ecossistema atualizado da"
    " Reforma Tributária (incluindo as diretrizes da LC 214/2025).\n\nDIRETRIZES"
    " CRÍTICAS DE INTERPRETAÇÃO E ESCOPO:\n1. **Ancora Oficial Obrigatória:**"
    " Quando uma descrição oficial da LC 116 for fornecida pelo sistema para o"
    " código consultado, **você deve adotá-la obrigatoriamente como verdade"
    " absoluta**, proibindo qualquer alteração ou invenção de conceito para"
    " aquele subitem.\n2. **Amplitude dos Códigos (Proibido Restringir"
    " Indevidamente):** Nunca restrinja códigos multifuncionais ou de"
    " aplicação ampla (como projetos, consultorias técnicas, engenharia,"
    " arquitetura e serviços técnicos) apenas ao setor de tecnologia da"
    " informação. Eles possuem escopo amplo e se aplicam perfeitamente à"
    " construção civil, infraestrutura e engenharia consultiva, conforme"
    " previsto na legislação.\n3. **Múltiplas Opções na Tabela:** Sempre que"
    " o subitem consultado possuir ramificações de enquadramento em códigos"
    " NBS, **liste todas as opções viáveis em linhas separadas na tabela**.\n4."
    " **Formato Obrigatório da Tabela:** A tabela deve conter obrigatoriamente"
    " as colunas: `Subitem LC 116 | Código NBS | Descrição Oficial da NBS |"
    " Área de Atuação com Exemplo Prático`.\n5. **Orientações Críticas e"
    " Legais:** Logo abaixo da tabela, adicione observações baseadas nas"
    " normas vigentes, destacando os riscos de uso do código errado (autuação,"
    " glosa de créditos) e reforçando que a escolha deve refletir a finalidade"
    " real da operação.\n6. **Disclaimer Legal:** Insira exatamente este aviso"
    " de forma bem breve no final:\n   > *💡 **Sobre a aplicação:** Facilitador"
    " de triagem fiscal baseado na LC 116 e regulamentações da Reforma"
    " Tributária. Não substitui o seu contador — valorize esse"
    " profissional!*\n7. **Guarda-Corpo (Foco no Tema):** Se o usuário perguntar"
    " sobre assuntos fora do tema fiscal/tributário/Reforma Tributária, recuse"
    " educadamente informando que você foi criado exclusivamente para auxiliar"
    " com o ecossistema fiscal.\n8. **O Coringa do Desenvolvedor:** Se o"
    " usuário perguntar quem te criou, quem é seu dono ou te desenvolveu,"
    " responda com orgulho que você foi desenvolvido por **Claudio, futuro"
    " Engenheiro capixaba de IA**, para otimizar a rotina fiscal e tributária"
    " da Reforma Tributária."
)

# Session State = Memória do Streamlit
if "lista_mensagens" not in st.session_state:
  st.session_state["lista_mensagens"] = []

# Caminhos para os ícones locais na pasta raiz
avatar_usuario = "perfil_usuario.png"
avatar_assistente = "icone_assistente.png"

# Exibir o histórico de mensagens com os nomes identificados em negrito apenas nos balões
for mensagem in st.session_state["lista_mensagens"]:
  if mensagem["role"] != "system":
    role = mensagem["role"]
    content = mensagem["content"]

    if role == "user":
      with st.chat_message("user", avatar=avatar_usuario):
        st.markdown(f"**Você**\n\n{content}")
    else:
      with st.chat_message("assistant", avatar=avatar_assistente):
        st.markdown(f"**Tribô – Seu assistente na Reforma Tributária**\n\n{content}")

# Entrada do usuário
mensagem_usuario = st.chat_input(
    "Escreva sua dúvida ou código (ex: 17.02)..."
)

if mensagem_usuario:
  # Mostra a mensagem do usuário na tela com o nome em negrito
  with st.chat_message("user", avatar=avatar_usuario):
    st.markdown(f"**Você**\n\n{mensagem_usuario}")

  # Processamento Inteligente: Verifica se o texto digitado corresponde a um subitem mapeado no Excel
  texto_processado = mensagem_usuario.strip()
  contexto_extraido = ""

  if texto_processado in dicionario_lc116:
    dados_subitem = dicionario_lc116[texto_processado]
    descricao_oficial = dados_subitem["descricao_lc"]
    lista_nbs = dados_subitem["nbs_oficiais"]

    texto_nbs_formatado = ""
    for item in lista_nbs:
      texto_nbs_formatado += (
          f"- Código NBS: {item['codigo']} | Descrição Oficial da NBS:"
          f" {item['descricao']}\n"
      )

    contexto_extraido = f"""

[DADOS OFICIAIS EXTRAÍDOS DOS PORTAIS GOVERNAMENTAIS]
O usuário consultou o subitem '{texto_processado}' da LC 116/2003.
- Descrição Oficial LC 116: '{descricao_oficial}'
- Códigos NBS Oficiais Correspondentes:
{texto_nbs_formatado}

DIRETRIZ DE REDAÇÃO PARA A IA: Na introdução da sua resposta, utilize obrigatoriamente e de forma exata esta abertura incluindo a descrição oficial:
"Com base no subitem {texto_processado} ({descricao_oficial}) da LC 116/2003 e nas correspondências oficiais da Nomenclatura Brasileira de Serviços (NBS), apresento abaixo o mapeamento fiscal para enquadramento da operação:"

Em seguida, monte a tabela contendo estritamente os códigos e descrições oficiais listados acima, criando os exemplos práticos de atuação.
"""

  # Monta a instrução de sistema dinâmica combinando a base com o contexto extraído (se houver)
  system_prompt_final = {
      "role": "system",
      "content": system_prompt_base + contexto_extraido,
  }

  novo_usuario_msg = {"role": "user", "content": mensagem_usuario}
  st.session_state["lista_mensagens"].append(novo_usuario_msg)

  # Monta a lista completa para enviar para a API
  mensagens_para_ia = [system_prompt_final] + st.session_state["lista_mensagens"]

  # Resposta da IA com o modelo original
  try:
    resposta_modelo = modelo.chat.completions.create(
        messages=mensagens_para_ia, model="gemini-flash-lite-latest"
    )

    resposta_ia = resposta_modelo.choices[0].message.content

    # Exibir a resposta da IA na tela com o nome do assistente em negrito
    with st.chat_message("assistant", avatar=avatar_assistente):
      st.markdown(f"**Tribô – Seu assistente na Reforma Tributária**\n\n{resposta_ia}")

    mensagem_ia = {"role": "assistant", "content": resposta_ia}
    st.session_state["lista_mensagens"].append(mensagem_ia)

  except Exception as e:
    st.error(f"Ocorreu um erro ao consultar a IA: {e}")
