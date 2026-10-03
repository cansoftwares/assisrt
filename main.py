import streamlit as st
import pandas as pd
from openai import OpenAI
import os

# Configuração da Página do Streamlit
st.set_page_config(page_title="Assistente NBS & Reforma Tributária", page_icon="⚖️")

# Estilo CSS avançado para largura equilibrada (70%) e layout estilo WhatsApp
st.markdown("""
<style>
    /* Define uma largura máxima de 70% para os balões de chat */
    .stChatMessage {
        max-width: 70% !important;
        width: 70% !important;
        border-radius: 15px;
        padding: 12px 18px;
        margin-bottom: 12px;
    }

    /* Mensagem do Usuário (Alinhada à Direita e com cor de destaque) */
    [data-testid="stChatMessage-user"] {
        background-color: #005c4b !important; /* Verde escuro tipo WhatsApp */
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
        background-color: #1f2c34 !important; /* Cinza/escuro tipo WhatsApp */
        margin-left: 0px !important;
        margin-right: auto !important;
    }

    /* Mantém as tabelas legíveis dentro do balão */
    table {
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

# Título do Chatbot
st.write("### ⚖️ Assistente Especialista em NBS e Reforma Tributária")
st.markdown("Consulte códigos de serviços (LC 116/2003), descrições oficiais, equivalências NBS e exemplos práticos.")

# Função para carregar o Anexo VIII da raiz do projeto e buscar a descrição da LC 116
@st.cache_data
def carregar_base_lc116():
    # Substitua pelo nome exato do arquivo que você subiu no GitHub se for diferente
    caminho_excel = "AnexoVIII-CorrelacaoItemNBSIndOpCClassTrib_IBSCBS_V1.00.00.xlsx"
    if os.path.exists(caminho_excel):
        try:
            # Lendo o Excel (ignorando cabeçalhos genéricos para pegar direto pelas colunas)
            df = pd.read_excel(caminho_excel, header=None)
            # Coluna 0 (A) = Subitem LC 116, Coluna 1 (B) = Descrição Oficial
            base_mapeada = {}
            for index, row in df.iterrows():
                subitem = str(row[0]).strip()
                descricao = str(row[1]).strip()
                if subitem and subitem != "nan":
                    base_mapeada[subitem] = descricao
            return base_mapeada
        except Exception as e:
            return {}
    return {}

# Carrega o dicionário de subitens
dicionario_lc116 = carregar_base_lc116()

# Inicialização do Cliente OpenAI configurado para o Gemini API
#modelo = OpenAI(
   # api_key="AQ.Ab8RN6IKsZFieIurPFiN1ywQ3MK-p8-viH_xxUTy_hGrkRAUZw",
   # base_url="https://generativelanguage.googleapis.com/v1beta/openai"
#)
modelo = OpenAI(
    api_key=st.secrets["GOOGLE_API_KEY"],
    base_url="https://generativelanguage.googleapis.com/v1beta/openai"
)

# Instrução de Sistema (System Prompt Blindado e Baseado no Anexo VIII)
system_prompt_base = (
    "Você é um assistente de inteligência artificial altamente especializado em classificação fiscal de serviços, "
    "com foco na Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei Complementar 116/2003, aos Anexos da regulamentação "
    "e ao ecossistema atualizado da Reforma Tributária (incluindo as diretrizes da LC 214/2025).\n\n"
    
    "DIRETRIZES CRÍTICAS DE INTERPRETAÇÃO E ESCOPO:\n"
    "1. **Ancora Oficial Obrigatória:** Quando uma descrição oficial da LC 116 for fornecida pelo sistema para o código consultado, **você deve adotá-la obrigatoriamente como verdade absoluta**, proibindo qualquer alteração ou invenção de conceito para aquele subitem.\n"
    "2. **Amplitude dos Códigos (Proibido Restringir Indevidamente):** Nunca restrinja códigos multifuncionais ou de aplicação ampla (como projetos, consultorias técnicas, engenharia, arquitetura e serviços técnicos) apenas ao setor de tecnologia da informação. Eles possuem escopo amplo e se aplicam perfeitamente à construção civil, infraestrutura e engenharia consultiva, conforme previsto na legislação.\n"
    "3. **Múltiplas Opções na Tabela:** Sempre que o subitem consultado possuir ramificações de enquadramento em códigos NBS, **liste todas as opções viáveis em linhas separadas na tabela**.\n"
    "4. **Formato Obrigatório da Tabela:** A tabela deve conter obrigatoriamente as colunas: "
    "`Subitem LC 116 | Código NBS | Descrição Oficial da NBS | Área de Atuação com Exemplo Prático`.\n"
    "5. **Orientações Críticas e Legais:** Logo abaixo da tabela, adicione observações baseadas nas normas vigentes, destacando os riscos de uso do código errado (autuação, glosa de créditos) e reforçando que a escolha deve refletir a finalidade real da operação.\n"
    "6. **Disclaimer Legal:** Insira exatamente este aviso de forma bem breve no final:\n"
    "   > *💡 **Sobre a aplicação:** Facilitador de triagem fiscal baseado na LC 116 e regulamentações da Reforma Tributária. Não substitui o seu contador — valorize esse profissional!*\n"
    "7. **Guarda-Corpo (Foco no Tema):** Se o usuário perguntar sobre assuntos fora do tema fiscal/tributário/Reforma Tributária, recuse educadamente informando que você foi criado exclusivamente para auxiliar com o ecossistema fiscal.\n"
    "8. **O Coringa do Desenvolvedor:** Se o usuário perguntar quem te criou, quem é seu dono ou te desenvolveu, responda com orgulho que você foi desenvolvido por **Claudio, futuro Engenheiro capixaba de IA**, para otimizar a rotina fiscal e tributária da Reforma Tributária."
)

# Session State = Memória do Streamlit
if "lista_mensagens" not in st.session_state:
    st.session_state["lista_mensagens"] = []

# Caminhos para os ícones locais na pasta raiz
avatar_usuario = "perfil_usuario.png"
avatar_assistente = "icone_assistente.png"

# Exibir o histórico de mensagens
for mensagem in st.session_state["lista_mensagens"]:
    if mensagem["role"] != "system":
        role = mensagem["role"]
        content = mensagem["content"]
        
        if role == "user":
            st.chat_message(role, avatar=avatar_usuario).write(content)
        else:
            st.chat_message(role, avatar=avatar_assistente).write(content)

# Entrada do usuário
mensagem_usuario = st.chat_input("Escreva sua dúvida ou código (ex: 17.19)...")

if mensagem_usuario:
    # Mostra a mensagem do usuário na tela
    st.chat_message("user", avatar=avatar_usuario).write(mensagem_usuario)
    
    # Processamento Inteligente: Verifica se o texto digitado corresponde a um subitem da LC 116 mapeado no Excel
    texto_processado = mensagem_usuario.strip()
    contexto_extraido = ""
    
    if texto_processado in dicionario_lc116:
        descricao_oficial = dicionario_lc116[texto_processado]
        contexto_extraido = f"\n\n[DADO OFICIAL EXTRAÍDO DO ANEXO VIII] O usuário consultou o subitem '{texto_processado}' da LC 116/2003, cuja descrição oficial exata e imutável é: '{descricao_oficial}'. Utilize estritamente esta definição."

    # Monta a instrução de sistema dinâmica combinando a base com o contexto extraído (se houver)
    system_prompt_final = {
        "role": "system",
        "content": system_prompt_base + contexto_extraido
    }

    novo_usuario_msg = {"role": "user", "content": mensagem_usuario}
    st.session_state["lista_mensagens"].append(novo_usuario_msg)

    # Monta a lista completa para enviar para a API
    mensagens_para_ia = [system_prompt_final] + st.session_state["lista_mensagens"]

    # Resposta da IA
    try:
        resposta_modelo = modelo.chat.completions.create(
            messages=mensagens_para_ia,
            model="gemini-flash-lite-latest"
        )
        
        resposta_ia = resposta_modelo.choices[0].message.content

        # Exibir a resposta da IA na tela
        st.chat_message("assistant", avatar=avatar_assistente).write(resposta_ia)
        mensagem_ia = {"role": "assistant", "content": resposta_ia}
        st.session_state["lista_mensagens"].append(mensagem_ia)
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao consultar a IA: {e}")
