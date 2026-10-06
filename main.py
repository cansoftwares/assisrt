import io
import json
import os
import re
import pandas as pd
import streamlit as st
from openai import OpenAI
from openpyxl.styles import Alignment, Font

# Configuração da Página e do Título da Aba do Navegador
st.set_page_config(
    page_title="Assistente NBS & Reforma Tributária", page_icon="⚖️"
)

# Estilo CSS otimizado para o layout corporativo
st.markdown(
    """
<style>
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

    [data-testid="stChatInput"] {
        border-radius: 12px !important;
    }

    .cabecalho-principal {
        font-size: 1.45rem !important;
        font-weight: 600;
        margin-top: 0.5rem;
        margin-bottom: 0.75rem;
        color: inherit;
        word-break: break-word;
        line-height: 1.3 !important;
    }

    .stChatMessage {
        max-width: 100% !important;
        width: 100% !important;
        border-radius: 12px;
        padding: 12px 18px;
        margin-bottom: 12px;
        border: 1px solid rgba(49, 51, 63, 0.1);
    }

    [data-testid="stChatMessage-user"] {
        background-color: #f0f2f6 !important; 
        margin-left: auto !important;
        margin-right: 0px !important;
        flex-direction: row-reverse;
    }
    
    [data-testid="stChatMessage-user"] > div:first-child {
        flex-direction: row-reverse;
    }

    [data-testid="stChatMessage-assistant"] {
        background-color: #ffffff !important; 
        margin-left: 0px !important;
        margin-right: auto !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.03);
    }

    table {
        width: 100% !important;
    }
    
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

# ==========================================
# 1. INICIALIZAÇÃO DA MEMÓRIA DO CHAT
# ==========================================
if "lista_mensagens" not in st.session_state:
    st.session_state["lista_mensagens"] = []

if "pending_nbs_prompt" not in st.session_state:
    st.session_state["pending_nbs_prompt"] = None

# Componente para focar automaticamente no input
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

# Cabeçalho limpo e descritivo no topo da página
st.markdown(
    '<p class="cabecalho-principal">⚖️ Assistente NBS & Reforma Tributária</p>',
    unsafe_allow_html=True,
)
st.markdown(
    "Consulte códigos de serviços da LC 116/2003, descrições normativas oficiais e correspondências detalhadas de equivalência NBS para o ecossistema tributário."
)


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
        except Exception:
            return {}
    return {}


dicionario_lc116 = carregar_base_lc116()

resumo_base_texto = ""
for subitem_k, info_v in dicionario_lc116.items():
    resumo_base_texto += (
        f"Subitem LC 116: {subitem_k} - Descrição: {info_v['descricao_lc']}\n"
    )
    for nbs_item in info_v["nbs_oficiais"]:
        resumo_base_texto += (
            f"    -> NBS: {nbs_item['codigo']} | Descrição NBS:"
            f" {nbs_item['descricao']}\n"
        )

system_prompt_base = (
    "Você é o **Tribô**, um assistente de inteligência artificial altamente"
    " especializado em classificação fiscal de serviços, com foco na"
    " Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei"
    " Complementar 116/2003 e ao ecossistema da Reforma Tributária.\n\n"
    "🚨 **DIRETRIZES CRÍTICAS DE PREENCHIMENTO E ESCOPO:**\n"
    "1. **Restrição Absoluta de Tema:** Exclusivo para Reforma Tributária,"
    " LC 116/2003, NBS e classificação fiscal.\n"
    "2. **Tom em Primeira Pessoa do Singular:** Responda SEMPRE em **primeira"
    " pessoa do singular** (ex: 'identifiquei', 'apresento', 'consultei'). É"
    " proibido o uso do plural.\n"
    "3. **Estilo de Resposta Direto e Natural:** Diga qual subitem da LC"
    " 116/2003 você identificou e apresente o texto principal da análise.\n"
    "4. **Exaustividade Obrigatória:** Liste absolutamente todos os códigos"
    " NBS oficiais vinculados ao subitem na estrutura solicitada.\n"
    "5. **Exemplo Prático Real:** A última coluna deve descrever um caso real"
    " de mercado, sem copiar a descrição NBS.\n"
    "6. **Formato JSON Oculto:** Forneça no final o bloco JSON exato com a"
    " chave `dados_tabela` contendo: `subitem`, `codigo_nbs`, `descricao_nbs`, `exemplo_pratico`.\n"
    "7. **O Coringa do Desenvolvedor:** Desenvolvido por **Claudio, futuro"
    " Engenheiro capixaba de IA**.\n\n"
    "### TABELA DE REFERÊNCIA OFICIAL (LC 116 / NBS):\n"
    f"{resumo_base_texto}"
)

modelo = OpenAI(
    api_key=st.secrets["GOOGLE_API_KEY"],
    base_url="https://generativelanguage.googleapis.com/v1beta/openai",
)

avatar_usuario = "perfil_usuario.png"
avatar_assistente = "icone_assistente.png"
deve_focar_input = False

# Exibir histórico preservando conversas anteriores
for idx, mensagem in enumerate(st.session_state["lista_mensagens"]):
    role = mensagem["role"]
    content = mensagem["content"]

    if role == "user":
        with st.chat_message("user", avatar=avatar_usuario):
            st.markdown(f"**Você**\n\n{content}")
    elif role == "assistant":
        with st.chat_message("assistant", avatar=avatar_assistente):
            tabela_para_baixar = mensagem.get("tabela_dados", [])
            subitem_referencia = mensagem.get("subitem_ref", "Geral")

            if not tabela_para_baixar and subitem_referencia in dicionario_lc116:
                info_sub_rec = dicionario_lc116[subitem_referencia]
                tabela_para_baixar = []
                for nbs_obj in info_sub_rec["nbs_oficiais"]:
                    tabela_para_baixar.append({
                        "Subitem LC 116": subitem_referencia,
                        "Código NBS": nbs_obj["codigo"],
                        "Descrição Oficial da NBS": nbs_obj["descricao"],
                        "Área de Atuação com Exemplo Prático": (
                            f"Execução de serviços especializados para"
                            f" {info_sub_rec['descricao_lc'].lower()}."
                        ),
                    })

            # Nome do assistente ao lado do avatar do robô
            st.markdown("**Tribô – Seu assistente na Reforma Tributária**")

            # 1. Texto principal da resposta da IA
            st.markdown(content, unsafe_allow_html=True)

            # 2. Aviso importante resumido
            st.markdown(
                "*Importante: Escolha com precisão o NBS, a correta classificação garante a aplicação adequada das regras, mitigando riscos de bitributação ou autuações fiscais.*"
            )

            # 3. Botões de Ações Rápidas organizados em linhas estruturadas de 4 colunas (com texto visível e compacto)
            if tabela_para_baixar:
                st.markdown(
                    "<small><b>Ações rápidas:</b> <i>(Clique para aprofundar no código)</i></small>",
                    unsafe_allow_html=True,
                )
                
                itens_por_linha = 4
                for i in range(0, len(tabela_para_baixar), itens_por_linha):
                    lote_atual = tabela_para_baixar[i : i + itens_por_linha]
                    cols = st.columns(itens_por_linha)
                    
                    for j, row_data in enumerate(lote_atual):
                        cod_nbs_atual = row_data.get("Código NBS", "")
                        with cols[j]:
                            if st.button(
                                f"🔍 {cod_nbs_atual}",
                                key=f"btn_nbs_{idx}_{i+j}_{cod_nbs_atual}",
                                use_container_width=True,
                            ):
                                st.session_state["pending_nbs_prompt"] = (
                                    f"Por favor, traga mais detalhes estratégicos, regras de"
                                    f" tributação e enquadramento avançado para o código NBS"
                                    f" {cod_nbs_atual}."
                                )
