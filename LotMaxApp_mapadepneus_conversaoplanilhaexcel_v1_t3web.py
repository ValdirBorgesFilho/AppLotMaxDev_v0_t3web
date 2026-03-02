import sys
import os
import io
import warnings
import streamlit as st
import pandas as pd
import xlsxwriter

# --- 1. CONFIGURAÇÃO DA INTERFACE ---
st.set_page_config(
    page_title="AppLotMax | Mapeador Web", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. FUNÇÃO DE LEITURA (BLINDADA PARA PYTHON 3.14) ---
def ler_dados_excel(file, aba):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            # Lemos o Excel sem cache para evitar o erro de 'warnings' no Py 3.14
            df = pd.read_excel(file, sheet_name=aba, engine='openpyxl')
            return df.copy()
        except Exception as e:
            st.error(f"Erro ao processar a aba '{aba}': {e}")
            return None

# --- 3. CSS PARA OTIMIZAÇÃO (INTERFACE DESKTOP) ---
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding-top: 1rem !important; padding-bottom: 0rem !important; max-width: 98% !important; }
    div[data-baseweb="select"] > div { height: 32px !important; min-height: 32px !important; }
    .mapping-label { font-weight: 700; color: #2c3e50; margin-bottom: 2px; font-size: 0.85rem; display: block; }
    div[data-testid="stSelectbox"] { margin-bottom: -5px !important; }
    .val-error { color: #d63031; font-size: 0.65rem; font-weight: 700; margin-top: 2px; line-height: 1.1; }
    .val-warning { color: #f39c12; font-size: 0.65rem; font-weight: 700; margin-top: 2px; line-height: 1.1; }
    </style>
    """, unsafe_allow_html=True)

# --- 4. CABEÇALHO ---
c_logo, c_titulo = st.columns([1, 5])
with c_logo:
    logo_nome = "Lotmax_app_lotmax_2026.png"
    if os.path.exists(logo_nome):
        st.image(logo_nome, width=110)
    else:
        st.markdown("### 🚀 LotMax")
with c_titulo:
    st.markdown("<h3 style='margin-top: 10px;'>Mapeador de Planilhas de Pneus</h3>", unsafe_allow_html=True)

st.divider()

# --- 5. BARRA LATERAL ---
with st.sidebar:
    st.markdown("### 📂 Gestão de Arquivo")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx", "xls"], label_visibility="collapsed")
    if uploaded_file:
        if st.button("🗑️ Limpar Seleções"):
            st.session_state.map_state = {}
            st.rerun()

# --- 6. LÓGICA DE MAPEAMENTO E VALIDAÇÃO ---
if uploaded_file:
    # EXIBE O NOME DO ARQUIVO ACIMA DA SELEÇÃO DA ABA
    st.markdown(f"📄 **Arquivo:** `{uploaded_file.name}`")
    
    xls = pd.ExcelFile(uploaded_file)
    aba_sel = st.selectbox("Selecione a Aba desejada:", xls.sheet_names, key="aba_main")

    if aba_sel:
        df_origem = ler_dados_excel(uploaded_file, aba_sel)
        
        if df_origem is not None:
            colunas_planilha = df_origem.columns.tolist()
            lista_fixa_base = [
                "Placa ou Estoque","Marca","Recapadora","Tipo","Aplicacao",
                "Codigo aplicado","Condicao","Medida","Vida util atual",
                "Recapes possíveis","Vida util recapes","Codigo comercial",
                "DOT fabricado","Valor da compra"
            ]

            # Função auxiliar para formatar a lista de linhas com indicador de excesso "..."
            def format_rows(mask):
                lista_linhas = mask[mask].index.map(lambda x: int(x) + 2).tolist()
                total_erros = len(lista_linhas)
                if total_erros > 3:
                    return f"{lista_linhas[:3]}... (+{total_erros - 3})"
                return str(lista_linhas)

            if 'map_state' not in st.session_state:
                st.session_state.map_state = {item: "(Pular)" for item in lista_fixa_base}

            selecionados_atualmente = {v for k, v in st.session_state.map_state.items() if v != "(Pular)"}
            tem_erros_criticos = False

            grid = st.columns(4)
            for idx, item_fixo in enumerate(lista_fixa_base):
                with grid[idx % 4]:
                    st.markdown(f"<span class='mapping-label'>{item_fixo}</span>", unsafe_allow_html=True)
                    valor_salvo = st.session_state.map_state.get(item_fixo, "(Pular)")
                    
                    opcoes_disponiveis = ["(Pular)"] + [c for c in colunas_planilha if c not in selecionados_atualmente or c == valor_salvo]
                    
                    try:
                        idx_padrao = opcoes_disponiveis.index(valor_salvo)
                    except ValueError:
                        idx_padrao = 0

                    nova_escolha = st.selectbox(
                        f"sel_{item_fixo}", 
                        options=opcoes_disponiveis, 
                        index=idx_padrao, 
                        key=f"f_{item_fixo}", 
                        label_visibility="collapsed"
                    )
                    
                    if nova_escolha != valor_salvo:
                        st.session_state.map_state[item_fixo] = nova_escolha
                        st.rerun()

                    # --- REGRAS DE VALIDAÇÃO ---
                    # --- REGRAS DE VALIDAÇÃO COM TEXTOS ORIENTATIVOS ---
                    if nova_escolha != "(Pular)":
                        dados = df_origem[nova_escolha]
                        dados_limpos = dados.dropna()
                        
                        # Função auxiliar para formatar a lista de linhas
                        def format_rows(mask):
                            lista_linhas = mask[mask].index.map(lambda x: int(x) + 2).tolist()
                            total_erros = len(lista_linhas)
                            if total_erros > 3:
                                return f"{lista_linhas[:3]}... (+{total_erros - 3})"
                            return str(lista_linhas)

                        if item_fixo == "Placa ou Estoque":
                            mask = dados_limpos.apply(lambda x: len(str(x)) > 7 and str(x).strip().upper() != "ESTOQUE")
                            if mask.any():
                                st.markdown(f"<p class='val-error'>Linhas: {format_rows(mask)}<br>⚠️ Informe placa até 7 carac. ou 'Estoque'.</p>", unsafe_allow_html=True)

                        elif item_fixo == "Tipo":
                            validos = ["liso", "borrachudo", "borrachudo florestal off -road pesado", "borrachudo off-road leve", "single", "misto", "liso-reboque", "comercial leve", "comercial médio", "passeio"]
                            mask = ~dados_limpos.astype(str).str.lower().str.strip().isin(validos)
                            if mask.any():
                                st.markdown(f"<p class='val-warning'>Linhas: {format_rows(mask)}<br>⚠️ Use: liso, borrachudo, misto, etc.</p>", unsafe_allow_html=True)

                        elif item_fixo == "Aplicacao":
                            validos = ["pesado", "carreta", "leve ou medio", "passeio", "reboque"]
                            mask = ~dados_limpos.astype(str).str.lower().str.strip().isin(validos)
                            if mask.any():
                                st.markdown(f"<p class='val-warning'>Linhas: {format_rows(mask)}<br>⚠️ Use: pesado, carreta, leve/medio, passeio ou reboque.</p>", unsafe_allow_html=True)

                        elif item_fixo == "Codigo aplicado":
                            mask = dados.duplicated(keep=False) & dados.notna()
                            if mask.any():
                                st.markdown(f"<p class='val-error'>Linhas: {format_rows(mask)}<br>❌ Não permitir duplicações do código.</p>", unsafe_allow_html=True)
                                tem_erros_criticos = True

                        elif item_fixo == "Condicao":
                            validos = ["novo", "novo - em uso", "recapado", "recapado - em uso"]
                            mask = ~dados_limpos.astype(str).str.lower().str.strip().isin(validos)
                            if mask.any():
                                st.markdown(f"<p class='val-warning'>Linhas: {format_rows(mask)}<br>⚠️ Use: novo, recapado ou em uso.</p>", unsafe_allow_html=True)

                        elif item_fixo in ["Vida util atual", "Vida util recapes"]:
                            mask = pd.to_numeric(dados_limpos, errors='coerce').isna()
                            if mask.any():
                                st.markdown(f"<p class='val-error'>Linhas: {format_rows(mask)}<br>⚠️ Só valor numérico permitido.</p>", unsafe_allow_html=True)
                                tem_erros_criticos = True

                        elif item_fixo == "Recapes possíveis":
                            mask = ~dados_limpos.astype(str).str.strip().isin(["0", "1", "2", "3"])
                            if mask.any():
                                st.markdown(f"<p class='val-error'>Linhas: {format_rows(mask)}<br>⚠️ Valores permitidos: 0 a 3.</p>", unsafe_allow_html=True)
                                tem_erros_criticos = True

                        elif item_fixo == "DOT fabricado":
                            mask = dados_limpos.apply(lambda x: len(str(x).strip()) != 4)
                            if mask.any():
                                st.markdown(f"<p class='val-warning'>Linhas: {format_rows(mask)}<br>⚠️ Deve ter exatamente 4 dígitos.</p>", unsafe_allow_html=True)

                        elif item_fixo == "Valor da compra":
                            mask = pd.to_numeric(dados_limpos, errors='coerce').isna()
                            if mask.any():
                                st.markdown(f"<p class='val-error'>Linhas: {format_rows(mask)}<br>⚠️ Valor com duas casas decimais.</p>", unsafe_allow_html=True)

            # --- 7. EXPORTAÇÃO ---
            mapeamento_final = {v: k for k, v in st.session_state.map_state.items() if v != "(Pular)"}
            
            if mapeamento_final:
                st.divider()
                if tem_erros_criticos:
                    st.error("⚠️ Corrija os erros críticos (em vermelho) para habilitar o download.")
                else:
                    if st.button("🚀 GERAR PLANILHA CONVERTIDA"):
                        with st.spinner("Preparando download..."):
                            df_final = df_origem[list(mapeamento_final.keys())].copy()
                            df_final = df_final.rename(columns=mapeamento_final)
                            
                            if "Valor da compra" in df_final.columns:
                                df_final["Valor da compra"] = pd.to_numeric(df_final["Valor da compra"], errors='coerce').round(2)

                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                df_final.to_excel(writer, index=False)
                            
                            st.success("Planilha processada!")
                            st.download_button(
                                label="📥 BAIXAR AGORA (EXCEL)",
                                data=output.getvalue(),
                                file_name=f"Convertido_{uploaded_file.name}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
else:
    st.info("Aguardando upload do arquivo Excel pelo menu lateral...")
