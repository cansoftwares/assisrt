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

            coluna_subitem_lc = df.columns[0]
            coluna_desc_lc = df.columns[1] if len(df.columns) > 1 else df.columns[0]
            col_nbs = next((c for c in df.columns if "mbs" in c.lower() or "nbs" in c.lower()), df.columns[2])
            col_desc_nbs = next((c for c in df.columns if "descrição" in c.lower() and ("mbs" in c.lower() or "hbs" in c.lower())), df.columns[3] if len(df.columns) > 3 else df.columns[2])

            col_ind_op = next((c for c in df.columns if "indop" in c.lower()), df.columns[6] if len(df.columns) > 6 else "")
            col_local_ibs = next((c for c in df.columns if "local" in c.lower() or "incidência" in c.lower()), df.columns[7] if len(df.columns) > 7 else "")
            col_c_clas = next((c for c in df.columns if "cclasstrib" in c.lower() or "cclas" in c.lower()), df.columns[8] if len(df.columns) > 8 else "")
            col_nome_c_clas = df.columns[9] if len(df.columns) > 9 else ""

            df[coluna_subitem_lc] = df[coluna_subitem_lc].ffill()
            df[coluna_desc_lc] = df[coluna_desc_lc].ffill()
            df[col_nbs] = df[col_nbs].ffill()
            if col_desc_nbs in df.columns:
                df[col_desc_nbs] = df[col_desc_nbs].ffill()
            if col_c_clas and col_c_clas in df.columns:
                df[col_c_clas] = df[col_c_clas].ffill()
            if col_nome_c_clas and col_nome_c_clas in df.columns:
                df[col_nome_c_clas] = df[col_nome_c_clas].ffill()
            if col_local_ibs and col_local_ibs in df.columns:
                df[col_local_ibs] = df[col_local_ibs].ffill()

            base_mapeada = {}
            for _, row in df.iterrows():
                subitem_bruto = str(row[coluna_subitem_lc]).strip()
                match_sub = re.search(r"\b(\d{2}\.\d{2})\b", subitem_bruto)
                subitem = match_sub.group(1) if match_sub else subitem_bruto

                desc_lc_oficial = str(row[coluna_desc_lc]).strip() if pd.notna(row[coluna_desc_lc]) else subitem_bruto
                cod_nbs = str(row[col_nbs]).strip()
                desc_nbs = str(row[col_desc_nbs]).strip() if col_desc_nbs in df.columns and pd.notna(row[col_desc_nbs]) else ""
                
                ind_op = str(row[col_ind_op]).strip() if col_ind_op and col_ind_op in df.columns and pd.notna(row[col_ind_op]) else ""
                local_desc = str(row[col_local_ibs]).strip() if col_local_ibs and col_local_ibs in df.columns and pd.notna(row[col_local_ibs]) else "Domicílio principal do adquirente"
                
                c_clas = str(row[col_c_clas]).strip() if col_c_clas and col_c_clas in df.columns and pd.notna(row[col_c_clas]) else ""
                nome_c_clas = str(row[col_nome_c_clas]).strip() if col_nome_c_clas and col_nome_c_clas in df.columns and pd.notna(row[col_nome_c_clas]) else "Situação tributada integralmente pelo IBS e CBS."

                if subitem and subitem != "nan" and cod_nbs and cod_nbs != "nan" and cod_nbs.startswith("1."):
                    if subitem not in base_mapeada:
                        base_mapeada[subitem] = {
                            "descricao_lc": desc_lc_oficial,
                            "nbs_oficiais": {},
                        }

                    if cod_nbs not in base_mapeada[subitem]["nbs_oficiais"]:
                        base_mapeada[subitem]["nbs_oficiais"][cod_nbs] = {
                            "descricao": desc_nbs if desc_nbs and desc_nbs != "nan" else "Serviço associado",
                            "ind_ops_detalhes": {},
                            "c_clas_detalhes": {}
                        }

                    if ind_op and ind_op != "nan":
                        base_mapeada[subitem]["nbs_oficiais"][cod_nbs]["ind_ops_detalhes"][ind_op] = local_desc

                    if c_clas and c_clas != "nan":
                        base_mapeada[subitem]["nbs_oficiais"][cod_nbs]["c_clas_detalhes"][c_clas] = nome_c_clas

            return base_mapeada
        except Exception as e:
            print(f"Erro ao carregar base: {e}")
            return {}
    return {}


dicionario_lc116 = carregar_base_lc116()

resumo_base_texto = ""
for subitem_k, info_v in dicionario_lc116.items():
    resumo_base_texto += (
        f"Subitem LC 116: {subitem_k} - Descrição do Item: {info_v['descricao_lc']}\n"
    )
    for nbs_k, nbs_v in info_v["nbs_oficiais"].items():
        ind_ops_resumo = ", ".join([f"{k} ({v})" for k, v in nbs_v["ind_ops_detalhes"].items()])
        c_clas_resumo = ", ".join([f"{k} ({v})" for k, v in nbs_v["c_clas_detalhes"].items()])
        resumo_base_texto += (
            f"    -> NBS: {nbs_k} | Descrição Oficial NBS: {nbs_v['descricao']} | "
            f"IndOps: [{ind_ops_resumo}] | cClassTribs: [{c_clas_resumo}]\n"
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
    " pessoa do singular** (ex: 'analisei', 'identifiquei', 'apresento', 'consultei'). É"
    " estritamente proibido o uso do plural.\n"
    "3. **Separação Rigorosa de Telas:**\n"
    "   - **Consulta Inicial (Subitem):** Apresente **apenas** a tabela com 4 colunas (`Subitem LC 116`, `Código NBS`, `Descrição Oficial da NBS`, `Área de Atuação com Exemplo Prático`) contendo **exclusivamente** os códigos NBS diretamente vinculados àquele subitem exato na base oficial.\n"
    "   - **Aprofundamento (Clique no NBS):** Apresente a tabela contendo as colunas exatas exigidas e logo abaixo inclua a **LEGENDA** detalhando individualmente o significado de cada código presente (Subitem LC, NBS, IndOp e cClassTrib) no formato `Código - Descrição`.\n"
    "4. **Formato JSON Obrigatório para Espelhamento Exato no Excel:** Inclua sempre um bloco de código JSON isolado contendo exatamente a chave `dados_tabela`.\n\n"
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

            st.markdown("**Tribô – Seu assistente na Reforma Tributária**")
            st.markdown(content, unsafe_allow_html=True)

            st.markdown(
                "*Importante: Escolha com precisão o NBS, a correta classificação garante a aplicação adequada das regras, mitigando riscos de bitributação ou autuações fiscais.*"
            )

            # Botões de Ações Rápidas
            if tabela_para_baixar and not eh_aprofundamento:
                st.markdown(
                    "<small><b>Ações rápidas:</b> <i>(Clique abaixo no NBS escolhido para se aprofundar sobre)</i></small>",
                    unsafe_allow_html=True,
                )
                
                itens_por_linha = 6
                for i in range(0, len(tabela_para_baixar), itens_por_linha):
                    lote_atual = tabela_para_baixar[i : i + itens_por_linha]
                    cols = st.columns(itens_por_linha)
                    
                    for j, row_data in enumerate(lote_atual):
                        cod_nbs_atual = row_data.get("Código NBS", row_data.get("NBS", ""))
                        subitem_atual_btn = row_data.get("Subitem LC 116", subitem_referencia)
                        with cols[j]:
                            if st.button(
                                f"🔍 {cod_nbs_atual}",
                                key=f"btn_nbs_{idx}_{i+j}_{cod_nbs_atual}",
                                use_container_width=True,
                            ):
                                st.session_state["pending_nbs_prompt"] = (
                                    f"Por favor, me detalhe o código NBS {cod_nbs_atual} do subitem {subitem_atual_btn} e me informe quais códigos fiscais corretos (IndOp, cClassTrib e CST) devem ser preenchidos conforme a tabela oficial da Reforma Tributária."
                                )
                                st.rerun()

            st.markdown(
                "<small><i>Esta ferramenta atua como um suporte estratégico e inteligente de alto nível, não tendo o objetivo de substituir seu contador — <b>Valorize sempre esse profissional!</b></i></small>",
                unsafe_allow_html=True,
            )

            # Botão de Download em Excel com nomes dinâmicos inteligentes
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
                    nome_arquivo_excel = f"Relatorio_NBS_Inteligente_-_Subitem_{subitem_referencia}.xlsx"
                else:
                    colunas_desejadas = [
                        "Subitem LC 116", 
                        "CTN", 
                        "Código NBS", 
                        "IndOp", 
                        "cClassTrib", 
                        "CST IBS/CBS", 
                        "Exemplos Práticos"
                    ]
                    colunas_larguras = {"A": 14, "B": 10, "C": 16, "D": 14, "E": 18, "F": 16, "G": 55}
                    
                    nbs_referencia_arquivo = ""
                    if tabela_para_baixar:
                        nbs_referencia_arquivo = tabela_para_baixar[0].get("Código NBS", "NBS")
                    
                    nome_arquivo_excel = f"Relatorio_NBS_Inteligente_-_Subitem_{subitem_referencia}_NBS_{nbs_referencia_arquivo}.xlsx"

                for col in colunas_desejadas:
                    if col not in df_resposta.columns:
                        df_resposta[col] = ""
                
                df_resposta = df_resposta[colunas_desejadas]

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_resposta.to_excel(writer, index=False, sheet_name="Enquadramento")
                    workbook = writer.book
                    worksheet = workbook.active
                    
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
                                if col_idx in [1, 2, 3, 4, 5, 6]:
                                    cell.alignment = Alignment(
                                        horizontal="center", vertical="top", wrap_text=True
                                    )
                                else:
                                    cell.alignment = Alignment(
                                        horizontal="left", vertical="top", wrap_text=True
                                    )

                excel_data = output.getvalue()
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

    match_nbs_clicado = re.search(r"código NBS\s*([\d\.]+)(?:\s+do\s+subitem\s+([\d\.]+))?", texto_processado, re.IGNORECASE)
    if match_nbs_clicado:
        eh_aprofundamento_nbs = True
        nbs_alvo = match_nbs_clicado.group(1)
        subitem_informado = match_nbs_clicado.group(2) if match_nbs_clicado.lastindex >= 2 else None
        
        subitem_encontrado_exato = None
        for sub_k, info_v in dicionario_lc116.items():
            if nbs_alvo in info_v["nbs_oficiais"]:
                if subitem_informado and subitem_informado == sub_k:
                    subitem_encontrado_exato = sub_k
                    break
                elif not subitem_encontrado_exato:
                    subitem_encontrado_exato = sub_k

        if subitem_encontrado_exato:
            subitem_identificado_cache = subitem_encontrado_exato
            nbs_obj = dicionario_lc116[subitem_encontrado_exato]["nbs_oficiais"][nbs_alvo]
            
            c_clas_dict = nbs_obj["c_clas_detalhes"]
            c_clas_str = ", ".join(c_clas_dict.keys())
            
            ind_ops_dict = nbs_obj["ind_ops_detalhes"]
            ind_ops_str = ", ".join(ind_ops_dict.keys())
            
            desc_nbs_oficial_base = nbs_obj["descricao"]
            
            if "demolição" in desc_nbs_oficial_base.lower():
                exemplos_consolidados = (
                    "1. Derrubada controlada de uma antiga edificação comercial para liberação do terreno no local do imóvel com vistas a um novo empreendimento.\n"
                    "2. Demolição parcial de paredes e estruturas internas em galpão industrial para readequação de layout operacional.\n"
                    "3. Remoção e desmonte de marquise em risco iminente de queda em fachada de edifício residencial."
                )
            else:
                exemplos_consolidados = (
                    f"1. Prestação Principal: Execução de serviços referentes a {desc_nbs_oficial_base.lower()} para atendimento corporativo.\n"
                    f"2. Operação Especializada: Atividades técnicas correlatas a {desc_nbs_oficial_base.lower()} com emissão de laudo técnico.\n"
                    f"3. Suporte Contínuo: Acompanhamento e suporte operacional especializado em {desc_nbs_oficial_base.lower()}."
                )

            dados_tabela_estruturados.append({
                "Subitem LC 116": subitem_encontrado_exato,
                "CTN": "",
                "Código NBS": nbs_alvo,
                "IndOp": ind_ops_str,
                "cClassTrib": c_clas_str,
                "CST IBS/CBS": "",
                "Exemplos Práticos": exemplos_consolidados,
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

        for nbs_k, nbs_obj in info_sub["nbs_oficiais"].items():
            desc_nbs_oficial = nbs_obj["descricao"]
            if "demolição" in desc_nbs_oficial.lower():
                exemplo_txt = "Construção Civil - Exemplo Prático: Demolição controlada de edifício comercial antigo para preparação do terreno para nova edificação."
            else:
                exemplo_txt = f"Prestação Principal: Execução de serviços referentes a {desc_nbs_oficial.lower()}."

            dados_tabela_estruturados.append({
                "Subitem LC 116": subitem_encontrado_direto,
                "Código NBS": nbs_k,
                "Descrição Oficial da NBS": desc_nbs_oficial,
                "Área de Atuação com Exemplo Prático": exemplo_txt,
            })

    if subitem_encontrado_direto and not eh_aprofundamento_nbs:
        instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador mencionou o subitem '{subitem_encontrado_direto}' ({info_sub['descricao_lc']}).
Inicie obrigatoriamente com a frase exata: "Analisei a solicitação referente ao subitem {subitem_encontrado_direto} da Lista de Serviços da Lei Complementar nº 116/2003, que trata de {info_sub['descricao_lc']}."
Vá direto ao ponto, **sem adicionar nenhuma frase intermediária ou explicativa**.
Apresente obrigatoriamente a Tabela Markdown limpa com **exatamente 4 colunas** contendo **apenas** os NBS mapeados na base oficial para este subitem: 
1. Subitem LC 116
2. Código NBS
3. Descrição Oficial da NBS
4. Área de Atuação com Exemplo Prático

**OBRIGATÓRIO - BLOCO JSON DE ESPELHAMENTO PARA O EXCEL:**
No final da resposta, inclua obrigatoriamente um bloco de código JSON isolado contendo exatamente a chave `dados_tabela` com a lista contendo os objetos exatos usando as chaves: "Subitem LC 116", "Código NBS", "Descrição Oficial da NBS", "Área de Atuação com Exemplo Prático".
"""
    elif eh_aprofundamento_nbs:
        c_clas_str = dados_tabela_estruturados[0]['cClassTrib'] if dados_tabela_estruturados else ''
        ind_ops_str = dados_tabela_estruturados[0]['IndOp'] if dados_tabela_estruturados else ''
        desc_lc_val = dicionario_lc116[subitem_identificado_cache]['descricao_lc'] if subitem_identificado_cache in dicionario_lc116 else 'Serviços'
        desc_nbs_oficial_val = desc_nbs_oficial_base if 'desc_nbs_oficial_base' in locals() else 'Serviço associado'
        c_clas_dict_val = c_clas_dict if 'c_clas_dict' in locals() else {}
        ind_ops_dict_val = ind_ops_dict if 'ind_ops_dict' in locals() else {}

        legenda_indop_linhas = "\n  - ".join([f"**{k}** - {v}" for k, v in ind_ops_dict_val.items()]) if ind_ops_dict_val else ""
        legenda_cclas_linhas = "\n  - ".join([f"**{k}** - {v}" for k, v in c_clas_dict_val.items()]) if c_clas_dict_val else ""

        instrucao_especifica = f"""
[ORIENTAÇÃO ESPECÍFICA PARA ESTA MENSAGEM]
O utilizador solicitou o aprofundamento no código NBS {nbs_alvo} vinculado ao subitem {subitem_identificado_cache}.
Inicie com uma abordagem simpática e acolhedora em primeira pessoa: "Claro! Analisei com atenção o código NBS {nbs_alvo}, cuja descrição oficial na base é "{desc_nbs_oficial_val}", vinculado ao subitem {subitem_identificado_cache} da Lei Complementar nº 116/2003 ({desc_lc_val})."

**REGRA DE OURO DE LAYOUT:** 
- Apresente **imediatamente** a tabela estruturada contendo **exatamente estas 7 colunas**: 
  `| Subitem LC 116 | CTN | Código NBS | IndOp | cClassTrib | CST IBS/CBS | Exemplos Práticos |`
- Preencha com **uma única linha** para este NBS, deixando as colunas **CTN** e **CST IBS/CBS** estritamente **em branco**, e contendo **de 2 a 5 exemplos práticos detalhados** na coluna `Exemplos Práticos`.
- Abaixo da tabela, inclua obrigatoriamente a **LEGENDA** contendo o significado específico de cada código no formato exacto solicitado:
  **LEGENDA:**
  - **Subitem LC:** {subitem_identificado_cache} - {desc_lc_val}
  - **NBS:** {nbs_alvo} - {desc_nbs_oficial_val}
  - **IndOp:**
    - {legenda_indop_linhas}
  - **cClassTrib:**
    - {legenda_cclas_linhas}

**OBRIGATÓRIO - BLOCO JSON DE ESPELHAMENTO PARA O EXCEL:**
No final da resposta, inclua obrigatoriamente um bloco de código JSON isolado contendo exatamente a chave `dados_tabela` com a lista contendo o objeto exato contendo: "Subitem LC 116", "CTN", "Código NBS", "IndOp", "cClassTrib", "CST IBS/CBS", "Exemplos Práticos".
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
                temp_estruturados = []
                for item in dados_json["dados_tabela"]:
                    if eh_aprofundamento_nbs:
                        temp_estruturados.append({
                            "Subitem LC 116": item.get("Subitem LC 116", subitem_identificado_cache),
                            "CTN": "",
                            "Código NBS": item.get("Código NBS", nbs_alvo),
                            "IndOp": item.get("IndOp", ind_ops_str),
                            "cClassTrib": item.get("cClassTrib", c_clas_str),
                            "CST IBS/CBS": "",
                            "Exemplos Práticos": item.get("Exemplos Práticos", item.get("Exemplo Prático", "")),
                        })
                    else:
                        sub_val = item.get("Subitem LC 116", subitem_identificado_cache)
                        cod_nbs_val = item.get("Código NBS", "")
                        desc_nbs_val = item.get("Descrição Oficial da NBS", "")
                        exemplo_pratico_val = item.get("Área de Atuação com Exemplo Prático", "")

                        temp_estruturados.append({
                            "Subitem LC 116": sub_val,
                            "Código NBS": cod_nbs_val,
                            "Descrição Oficial da NBS": desc_nbs_val,
                            "Área de Atuação com Exemplo Prático": exemplo_pratico_val,
                        })

                if temp_estruturados:
                    dados_tabela_estruturados = temp_estruturados
        except Exception:
            pass

        if not dados_tabela_estruturados and subitem_encontrado_direto:
            info_sub = dicionario_lc116[subitem_encontrado_direto]
            for nbs_k, nbs_obj in info_sub["nbs_oficiais"].items():
                desc_nbs_oficial = nbs_obj["descricao"]
                exemplo_txt = "Construção Civil - Exemplo Prático: Demolição controlada de edifício comercial antigo para preparação do terreno para nova edificação." if "demolição" in desc_nbs_oficial.lower() else f"Prestação Principal: Execução de serviços referentes a {desc_nbs_oficial.lower()}."
                dados_tabela_estruturados.append({
                    "Subitem LC 116": subitem_encontrado_direto,
                    "Código NBS": nbs_k,
                    "Descrição Oficial da NBS": desc_nbs_oficial,
                    "Área de Atuação com Exemplo Prático": exemplo_txt,
                })

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
