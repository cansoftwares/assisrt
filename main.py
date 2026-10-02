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
modelo = OpenAI(
    api_key="AQ.Ab8RN6IKsZFieIurPFiN1ywQ3MK-p8-viH_xxUTy_hGrkRAUZw",
    base_url="https://generativelanguage.googleapis.com/v1beta/openai"
)

# Instrução de Sistema (System Prompt)
system_prompt = {
    "role": "system", 
    "content": (
        "Você é um assistente de inteligência artificial altamente especializado em classificação fiscal de serviços, "
        "com foco na Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei Complementar 116/2003 e à Reforma Tributária (IBS/CBS). "
        "Sua missão é responder de forma direta, sem enrolação, ajudando profissionais da área fiscal e tributária.\n\n"
        
        "REGRAS DE RESPOSTA:\n"
        "1. **Múltiplas Opções na Tabela:** Sempre que um subitem da LC 116 ou código consultado possuir **mais de uma possibilidade ou ramificação** de enquadramento em códigos NBS (o que é muito comum), **você deve listar todas as opções viáveis em linhas separadas na tabela**, permitindo que o usuário analise qual se aplica melhor ao seu caso real.\n"
        "2. **Formato da Tabela:** A tabela deve conter obrigatoriamente as colunas: "
        "`Subitem LC 116 | Código NBS | Descrição Oficial da NBS | Área de Atuação com Exemplo Prático`.\n"
        "3. **Orientações Críticas:** Logo abaixo da tabela, adicione observações curtas, diretas e críticas sobre os riscos de uso do código errado (riscos de autuação, glosa de créditos ou alíquotas incorretas) e reforce que deve ser escolhido o código que reflita a finalidade real da operação.\n"
        "4. **Disclaimer Legal:** Insira exatamente este aviso de forma bem breve no final:\n"
        "   > *💡 **Sobre a aplicação:** Facilitador de triagem fiscal. Não substitui o seu contador — valorize esse profissional!*\n"
        "5. **Guarda-Corpo (Foco no Tema):** Se o usuário perguntar sobre assuntos fora do tema fiscal/tributário/Reforma Tributária (como futebol, política, BBB, entretenimento geral, culinária, etc.), recuse educadamente informando que você foi criado exclusivamente para auxiliar com a Reforma Tributária e temas pertinentes ao ecossistema fiscal.\n"
        "6. **O Coringa do Desenvolvedor:** Se o usuário perguntar quem te criou, quem é seu dono, quem escreveu seu código ou te desenvolveu (ex: 'quem criou você?', 'quem te escreveu?', 'fale mais sobre você'), responda com orgulho que você foi desenvolvido por **Claudio, futuro Engenheiro capixaba de IA**, para otimizar a rotina fiscal e tributária da Reforma Tributária."
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