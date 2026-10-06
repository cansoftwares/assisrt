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
            
            col_ind_op = df.columns[6] if len(df.columns) > 6 else df.columns[4]
            col_c_clas = df.columns[8] if len(df.columns) > 8 else df.columns[5]

            df[col_item_lc] = df[col_item_lc].ffill()
            df[col_desc_lc] = df[col_desc_lc].ffill()

            base_mapeada = {}
            for _, row in df.iterrows():
                subitem = str(row[col_item_lc]).strip()
                desc_lc = str(row[col_desc_lc]).strip()
                cod_nbs = str(row[col_nbs]).strip()
                desc_nbs = str(row[col_desc_nbs]).strip()
                ind_op = str(row[col_ind_op]).strip() if col_ind_op in df.columns else ""
                c_clas = str(row[col_c_clas]).strip() if col_c_clas in df.columns else ""

                if subitem and subitem != "nan":
                    if subitem not in base_mapeada:
                        base_mapeada[subitem] = {
                            "descricao_lc": desc_lc,
                            "nbs_oficiais": [],
                        }

                    if cod_nbs and cod_nbs != "nan":
                        base_mapeada[subitem]["nbs_oficiais"].append(
                            {
                                "codigo": cod_nbs, 
                                "descricao": desc_nbs,
                                "ind_op": ind_op if ind_op != "nan" else "",
                                "c_clas": c_clas if c_clas != "nan" else ""
                            }
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
            f"    -> NBS: {nbs_item['codigo']} | Descrição NBS: {nbs_item['descricao']} | "
            f"IndOp: {nbs_item['ind_op']} | cClassTrib: {nbs_item['c_clas']}\n"
        )

system_prompt_base = (
    "Você é o **Tribô**, um assistente de inteligência artificial altamente"
    " especializado em classificação fiscal de serviços, com foco na"
    " Nomenclatura Brasileira de Serviços (NBS) vinculada à Lei"
    " Complementar 116/2003 e ao ecossistema da Reforma Tributária (NFSe Nacional, IBS e CBS).\n\n"
    "🚨 **DIRETRIZES CRÍTICAS DE PREENCHIMENTO E ESCOPO:**\n"
    "1. **Restrição Absoluta de Tema:** Exclusivo para Reforma Tributária,"
    " LC 116/2003, NBS e classificação fiscal.\n"
    "2. **Tom em Primeira Pessoa do Singular:** Responda SEMPRE em **primeira"
    " pessoa do singular** (ex: 'identifiquei', 'apresento', 'consultei'). É"
    " proibido o uso do plural.\n"
    "3. **Formato de Resposta para Subitens (Consulta Inicial):** Quando o usuário consultar um subitem da LC 116/2003 (ex: 17.19), "
    "apresente uma introdução **curta e direta** (apenas identificando o subitem e informando que apresenta a tabela abaixo, **sem repetições ou frases longas e genéricas**), e "
    "**obrigatoriamente inclua uma Tabela Markdown limpa com apenas 4 colunas**: "
    "Subitem LC 116, Código NBS, Descrição Oficial da NBS e Área de Atuação com Exemplo Prático. **NÃO inclua colunas IndOp ou cClassTrib nesta tabela inicial**.\n"
    "4. **Foco Prático na NFSe Nacional (Ao aprofundar em um NBS via clique):** Quando solicitado o detalhamento de um código NBS específico via clique no botão rápido, "
    "apresente uma análise completa estruturada com os parâmetros fiscais avançados: Item LC 116, CTN, NBS, IndOp, cClassTrib e CST IBS/CBS.\n"
    "5. **Formato JSON Oculto:** Forneça no final o bloco JSON exato com a"
    " chave `dados_tabela` contendo os dados estruturados correspondentes.\n"
    "6. **O Coringa do Desenvolvedor:** Desenvolvido por **Claudio, futuro"
    " Engenheiro capixaba de IA**.\n\n"
    "### TABELA DE REFERÊNCIA OFICIAL (LC 116 / NBS / IndOp / cClassTrib):\n"
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
            eh_aprofundamento = mensagem.get("eh_aprofundamento_nbs", False)

            # Nome do assistente ao lado do avatar do robô
            st.markdown("**Tribô – Seu assistente na Reforma Tributária**")

            # 1. Texto principal da resposta da IA
            st.markdown(content, unsafe_allow_html=True)

            # 2. Aviso importante resumido
            st.markdown(
                "*Importante: Escolha com precisão o NBS, a correta classificação garante a aplicação adequada das regras, mitigando riscos de bitributação ou autuações fiscais.*"
            )

            # 3. Botões de Ações Rápidas organizados em 6 colunas (exibidos apenas na listagem do subitem)
            if tabela_para_baixar and not eh_aprofundamento:
                st.markdown(
                    "<small><b>Ações rápidas:</b> <i>(Clique para aprofundar no código)</i></small>",
                    unsafe_allow_html=True,
                )
                
                itens_por_linha = 6
                for i in range(0, len(tabela_para_baixar), itens_por_linha):
                    lote_atual = tabela_para_baixar[i : i + itens_por_linha]
                    cols = st.columns(itens_por_linha)
                    
                    for j, row_data in enumerate(lote_atual):
                        cod_nbs_atual = row_data.get("Código NBS", row_data.get("NBS", ""))
                        with cols[j]:
                            if st.button(
                                f"🔍 {cod_nbs_atual}",
                                key=f"btn_nbs_{idx}_{i+j}_{cod_nbs_atual}",
                                use_container_width=True,
                            ):
                                st.session_state["pending_nbs_prompt"] = (
                                    f"Por favor, me detalhe o código NBS {cod_nbs_atual} e me informe quais códigos fiscais devem ser preenchidos nos documentos fiscais com a Reforma Tributária."
                                )
                                st.rerun()

            # 4. Rodapé de valorização do profissional contábil
            st.markdown(
                "<small><i>Esta ferramenta atua como um suporte estratégico e inteligente de alto nível, não tendo o objetivo de substituir seu contador — <b>Valorize sempre esse profissional!</b></i></small>",
                unsafe_allow_html=True,
            )

            # 5. Botão de Download do Excel adaptado ao contexto (4 colunas na listagem / 6 colunas no aprofundamento)
            if tabela_para_baixar:
                df_resposta = pd.DataFrame(tabela_para_baixar)
                
                if not eh_aprofundamento:
                    colunas_desejadas = [
                        "Subitem LC 116", 
                        "Código NBS", 
                        "Descrição Oficial da NBS", 
                        "Área de Atuação com Exemplo Prático"
                    ]
                    colunas_larguras = {"A": 14, "B": 14, "C": 35, "D": 50}
                else:
                    colunas_desejadas = [
                        "Item LC 116", 
                        "CTN", 
                        "NBS", 
                        "IndOp", 
                        "cClassTrib", 
                        "CST IBS/CBS"
                    ]
                    colunas_larguras = {"A": 14, "B": 10, "C": 14, "D": 12, "E": 14, "F": 25}

                for col in colunas_desejadas:
                    if col not in df_resposta.columns:
                        df_resposta[col] = ""
                
                df_resposta = df_resposta[colunas_desejadas]

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_resposta.to_excel(writer, index=False, sheet_name="Enquadramento")
                    workbook = writer.book
                    worksheet = writer.sheets["Enquadramento"]
                    
                    for coluna, largura in colunas_larguras.items():
                        worksheet.column_dimensions[coluna].width = largura

                    for row_idx, row in enumerate(worksheet.iter_rows(min_row=1), start=1):
                        for col_idx, cell in enumerate(row, start=1):
                            if row_idx == 1:
                                cell.font = Font(bold=True)
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center", wrap_text=True
                                )
                            else:
                                if col_idx in [1, 2, 4, 5]:
                                    cell.alignment = Alignment(
                                        horizontal="center", vertical="top", wrap_text=True
                                    )
                                else:
                                    cell.alignment = Alignment(
                                        horizontal="left", vertical="top", wrap_text=True
                                    )

                excel_data = output.getvalue()
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

# ==========================================
# 2. ENTRADA DE DADOS E PROCESSAMENTO DA IA
# ==========================================
mensagem_usuario = st.chat_input(
    "Escreva sua dúvida ou código (ex: 17.02, contabilidade...)"
)

texto_processado = None
if mensagem_usuario:
    texto_processado = mensagem_usuario.strip()
elif st.session_state["pending_nbs_prompt"]:
    texto_processado = st.session_state["pending_nbs_prompt"]
    st.session_state["pending_nbs_prompt"] = None

if texto_processado:
    with st.chat_message("user", avatar=avatar_usuario):
        st.markdown(f"**Você**\n\n{texto_processado}")
    st.session_state["lista_mensagens"].append(
        {"role": "user", "content": texto_processado}
    )

    dados_tabela_estruturados = []
    subitem_identificado_cache = "Geral"
    eh_aprofundamento_nbs = False

    match_nbs_clicado = re.search(r"código NBS\s*([\d\.]+)", texto_processado, re.IGNORECASE)
    if match_nbs_clicado:
        eh_aprofundamento_nbs = True
        nbs_alvo = match_nbs_clicado.group(1)
        
        for sub_k, info_v in dicionario_lc116.items():
            for nbs_item in info_v["nbs_oficiais"]:
                if nbs_item["codigo"] == nbs_alvo:
                    subitem_identificado_cache = sub_k
                    dados_tabela_estruturados.append({
                        "Item LC 116": sub_k,
                        "CTN": sub_k,
                        "NBS": nbs_item["codigo"],
                        "IndOp": nbs_item["ind_op"],
                        "cClassTrib": nbs_item["c_clas"],
                        "CST IBS/CBS": "Consultar Portal SVRS",
                    })

    subitem_encontrado_direto = None
    if not eh_aprofundamento_nbs:
        for sub in dicionario_lc116.keys():
            if sub.lower() in texto_processado.lower():
                subitem_encontrado_direto = sub
                break

    if subitem_encontrado_direto and not eh_aprofundamento_nbs:
        subitem_identificado_cache = subitem_encontrado_direto
        info_sub = dicionario_lc116[subitem_encontrado_direto]
        for nbs_obj in info_sub["nbs_oficiais"]:
            dados_tabela_estruturados.append({
                "Subitem LC 116": subitem_encontrado_direto,
                "Código NBS": nbs_obj["codigo"],
                "Descrição Oficial da NBS": nbs_obj["descricao"],
                "Área de Atuação com Exemplo Prático": (
                    f"Execução de serviços especializados para {info_sub['descricao_lc'].lower()}."
                ),
            })

        instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador mencionou diretamente o subitem '{subitem_encontrado_direto}' ({info_sub['descricao_lc']}).
Gere a resposta em PRIMEIRA PESSOA DO SINGULAR. Apresente apenas uma linha identificando o subitem de forma limpa e direta, indo direto para a tabela a seguir.
**ATENÇÃO AO FORMATO DA TABELA:** Gere obrigatoriamente uma Tabela Markdown limpa com **exatamente 4 colunas**: 
1. Subitem LC 116
2. Código NBS
3. Descrição Oficial da NBS
4. Área de Atuação com Exemplo Prático
**É terminantemente proibido incluir as colunas IndOp ou cClassTrib nesta listagem inicial.** No final, forneça o bloco JSON oculto correspondente.
"""
    elif eh_aprofundamento_nbs:
        instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador clicou no acesso rápido para aprofundar no código NBS {nbs_alvo}.
Apresente uma análise detalhada e estratégica em PRIMEIRA PESSOA DO SINGULAR sobre este NBS, explicando os parâmetros e a aplicação prática para a Reforma Tributária. 
Forneça a tabela Markdown detalhada contendo os campos técnicos para preenchimento da NFSe Nacional: Item LC 116, CTN, NBS, IndOp, cClassTrib e CST IBS/CBS.
No final, inclua o JSON oculto correspondente.
"""
    else:
        instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador fez a consulta: '{texto_processado}'.
Responda em PRIMEIRA PESSOA DO SINGULAR com foco estrito em LC 116 e Reforma Tributária.
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
            messages=mensagens_para_ia,
            model="gemini-flash-lite-latest",
            max_tokens=4000,
        )

        resposta_ia = resposta_modelo.choices[0].message.content

        try:
            if "```json" in resposta_ia:
                json_str = resposta_ia.split("```json")[1].split("```")[0].strip()
            elif "```" in resposta_ia:
                json_str = resposta_ia.split("```")[1].split("```")[0].strip()
            else:
                json_str = ""

            dados_json = json.loads(json_str)
            if "dados_tabela" in dados_json:
                if eh_aprofundamento_nbs:
                    dados_tabela_estruturados = []
                for item in dados_json["dados_tabela"]:
                    sub_val = item.get("subitem", item.get("subitem_lc_116", ""))
                    if sub_val and subitem_identificado_cache == "Geral":
                        subitem_identificado_cache = sub_val

                    if eh_aprofundamento_nbs:
                        dados_tabela_estruturados.append({
                            "Item LC 116": item.get("item_lc_116", item.get("subitem", "")),
                            "CTN": item.get("ctn", ""),
                            "NBS": item.get("nbs", item.get("codigo_nbs", "")),
                            "IndOp": item.get("ind_op", ""),
                            "cClassTrib": item.get("c_clas", item.get("cclas_trib", "")),
                            "CST IBS/CBS": item.get("cst_ibs_cbs", "Consultar Portal SVRS"),
                        })
                    else:
                        dados_tabela_estruturados.append({
                            "Subitem LC 116": sub_val,
                            "Código NBS": item.get("codigo_nbs", item.get("nbs", "")),
                            "Descrição Oficial da NBS": item.get("descricao_nbs", item.get("descricao", "")),
                            "Área de Atuação com Exemplo Prático": item.get("exemplo_pratico", item.get("area_de_atuacao_com_exemplo_pratico", "")),
                        })
        except Exception:
            pass

        if "```json" in resposta_ia:
            resposta_ia_exibicao = resposta_ia.split("```json")[0].strip()
        else:
            resposta_ia_exibicao = resposta_ia

        mensagem_ia = {
            "role": "assistant",
            "content": resposta_ia_exibicao,
            "tabela_dados": dados_tabela_estruturados,
            "subitem_ref": subitem_identificado_cache,
            "eh_aprofundamento_nbs": eh_aprofundamento_nbs,
        }
        st.session_state["lista_mensagens"].append(mensagem_ia)

        deve_focar_input = True
        st.rerun()

    except Exception as e:
        st.error(f"Ocorreu um erro ao consultar a IA: {e}")

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
