import streamlit as st
from openai import OpenAI

# Configuração da Página do Streamlit
st.set_page_config(page_title="Assistente NBS & Reforma Tributária", page_icon="⚖️")

# Estilo CSS avançado para largura equilibrada (70%) e layout estilo WhatsApp
st.markdown("""
<style>
    /* Define uma largura máxima de 70% para os balões de chat (equilibrado: nem muito estreito, nem colado nas bordas) */
    .stChatMessage {
        max-width: 70% !important;
        width: 70% !important;
        border-radius: 15px;
        padding: 12px 18px;
        margin-bottom: 12px;
    }

    /* Mensagem do Usuário (Alinhada à Direita e com cor de destaque) */
    [data-testid="stChatMessage-user"] {
        background-color: #005c4b !important; /* Estilo verde escuro tipo WhatsApp */
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
        background-color: #1f2c34 !important; /* Estilo cinza/escuro tipo WhatsApp */
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
st.markdown("Consulte códigos de serviços (LC 116/2003), códigos NBS, descrições oficiais e exemplos práticos.")

# Chave e Inicialização do Cliente Gemini via OpenAI SDK
#modelo = OpenAI(
   # api_key="AQ.Ab8RN6IKsZFieIurPFiN1ywQ3MK-p8-viH_xxUTy_hGrkRAUZw",
   # base_url="https://generativelanguage.googleapis.com/v1beta/openai"
#)

modelo = OpenAI(
    api_key=st.secrets["GOOGLE_API_KEY"],
    base_url="https://generativelanguage.googleapis.com/v1beta/openai"
)

# Instrução de Sistema (System Prompt)
system_prompt = {
    "role": "system", 
    "content": (
        "Você é um assistente de inteligência artificial altamente especializado em classificação fiscal de serviços, "
        "com foco na Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei Complementar 116/2003, aos Anexos da regulamentação (como o Anexo VIII) "
        "e ao ecossistema atualizado da Reforma Tributária (incluindo as diretrizes da LC 214/2025 e normas correlatas).\n\n"
        
        "DIRETRIZES CRÍTICAS DE INTERPRETAÇÃO E ESCOPO:\n"
        "1. **Amplitude dos Códigos (Proibido Restringir Indevidamente):** Nunca restrinja códigos multifuncionais ou de aplicação ampla (como o código NBS `1.1403.10.00` e congêneres) apenas ao setor de tecnologia da informação ou software. Códigos de projetos, consultorias técnicas, engenharia, arquitetura e serviços técnicos especializados possuem escopo amplo e são perfeitamente válidos e aplicáveis à construção civil, infraestrutura e engenharia consultiva, conforme previsto na legislação de regência e nos anexos oficiais.\n"
        "2. **Múltiplas Opções na Tabela:** Sempre que um subitem da LC 116 ou código consultado possuir **mais de uma possibilidade ou ramificação** de enquadramento em códigos NBS, **você deve listar todas as opções viáveis em linhas separadas na tabela**, contemplando os desdobramentos previstos nos anexos oficiais.\n"
        "3. **Formato Obrigatório da Tabela:** A tabela deve conter obrigatoriamente as colunas: "
        "`Subitem LC 116 | Código NBS | Descrição Oficial da NBS | Área de Atuação com Exemplo Prático`.\n"
        "4. **Orientações Críticas e Legais:** Logo abaixo da tabela, adicione observações baseadas nas normas vigentes (citando o regramento dos anexos e da LC 214/2025 quando aplicável), destacando os riscos de uso do código errado (autuação, glosa de créditos) e reforçando que a escolha deve refletir a finalidade real da operação.\n"
        "5. **Disclaimer Legal:** Insira exatamente este aviso de forma bem breve no final:\n"
        "   > *💡 **Sobre a aplicação:** Facilitador de triagem fiscal baseado na LC 116 e regulamentações da Reforma Tributária. Não substitui o seu contador — valorize esse profissional!*\n"
        "6. **Guarda-Corpo (Foco no Tema):** Se o usuário perguntar sobre assuntos fora do tema fiscal/tributário/Reforma Tributária, recuse educadamente informando que você foi criado exclusivamente para auxiliar com o ecossistema fiscal.\n"
        "7. **O Coringa do Desenvolvedor:** Se o usuário perguntar quem te criou, quem é seu dono ou te desenvolveu, responda com orgulho que você foi desenvolvido por **Claudio, futuro Engenheiro capixaba de IA**, para otimizar a rotina fiscal e tributária da Reforma Tributária."
    )
}

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
mensagem_usuario = st.chat_input("Escreva sua dúvida ou código (ex: 17.01)...")

if mensagem_usuario:
    # Mostra a mensagem do usuário na tela usando a imagem local
    st.chat_message("user", avatar=avatar_usuario).write(mensagem_usuario)
    novo_usuario_msg = {"role": "user", "content": mensagem_usuario}
    st.session_state["lista_mensagens"].append(novo_usuario_msg)

    # Monta a lista completa para enviar para a API
    mensagens_para_ia = [system_prompt] + st.session_state["lista_mensagens"]

    # Resposta da IA
    try:
        resposta_modelo = modelo.chat.completions.create(
            messages=mensagens_para_ia,
            model="gemini-flash-lite-latest"
        )
        
        resposta_ia = resposta_modelo.choices[0].message.content

        # Exibir a resposta da IA na tela usando a imagem local
        st.chat_message("assistant", avatar=avatar_assistente).write(resposta_ia)
        mensagem_ia = {"role": "assistant", "content": resposta_ia}
        st.session_state["lista_mensagens"].append(mensagem_ia)
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao consultar a IA: {e}")
