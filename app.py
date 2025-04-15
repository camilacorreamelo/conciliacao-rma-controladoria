import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Validador de Notas", layout="wide")

st.title("🧾 Conciliação RMA - Setor de Contabilidade")
st.markdown("Faça upload dos dois arquivos na ordem correta: primeiro **Planilha RMA - Tesouro Nacional em .xlsx**, depois **Planilha de Entradas de Recebimento do metabase em .xlsx**.")

uploaded_files = st.file_uploader(
    "📤 Envie os dois arquivos Excel:",
    type="xlsx",
    accept_multiple_files=True
)

# Função para extrair DANFE, NFS-e, Processo e Processo Relacionado
def extract_danfe_processo(text):
    if pd.isna(text):
        return []
    danfe_matches = re.findall(r'(DANFES{0,2}|DANDES?|DAFES?|NFS[ES]?)\.?.*?\s*-?\s*((?:\d+\s*(?:/|,|E)?\s*)+)', text, re.IGNORECASE)
    processo_matches = re.findall(r'(?:PROCESSO|PROC)\.?\s*(\d{5,}\.?\d{6,}/\d{4}-\d{2})', text, re.IGNORECASE)
    processo_relacionado_matches = re.findall(r'(?:PROCESSO|PROC)\.?\s*RELACIONADO\s*(\d{5,}\.?\d{6,}/\d{4}-\d{2})', text, re.IGNORECASE)
    processo_relacionado = processo_relacionado_matches[0] if processo_relacionado_matches else None
    extracted_data = [(match[0], d.strip()) for match in danfe_matches for d in re.split(r'\s*(?:/|,|E)\s*', match[1])]
    return [(tipo_nota, d, p, processo_relacionado) for tipo_nota, d, p in [(m[0], m[1], proc) for m in extracted_data for proc in processo_matches]] if extracted_data else [(None, None, None, processo_relacionado)]

if uploaded_files and len(uploaded_files) == 2:
    rma_file = uploaded_files[0]
    query_file = uploaded_files[1]

    try:
        # Leitura do RMA com tratamento
        df_rma = pd.read_excel(rma_file, skiprows=2)
        df_rma.rename(columns={"Favorecido Doc.": "CNPJ"}, inplace=True)

        # Processamento de DANFE / Processo
        new_rows = []
        for _, row in df_rma.iterrows():
            matches = extract_danfe_processo(row.get('Doc - Observação', ''))
            for tipo, nota, proc, proc_rel in matches:
                new_rows.append({
                    'DH - Dia Emissão': row.get('DH - Dia Emissão', ''),
                    'Documento Origem': row.get('Documento Origem', ''),
                    'CNPJ Fornecedor': row.get('CNPJ', ''),
                    'Tipo de Nota': tipo,
                    'DANFE/NFS-e': nota,
                    'Processo': proc,
                    'Processo Relacionado': proc_rel
                })

        df_validado = pd.DataFrame(new_rows)

        # Leitura da query para comparação
        df_query = pd.read_excel(query_file, dtype=str)

        # Padronização
        df_validado['cnpj'] = df_validado['CNPJ Fornecedor'].apply(lambda x: re.sub(r'\D', '', str(x))).str.zfill(14)
        df_query['cnpj'] = df_query['cnpj'].apply(lambda x: re.sub(r'\D', '', str(x))).str.zfill(14)
        df_validado.rename(columns={'DANFE/NFS-e': 'nota_fiscal'}, inplace=True)

        df_validado['chave'] = df_validado['cnpj'] + "_" + df_validado['nota_fiscal']
        df_query['chave'] = df_query['cnpj'] + "_" + df_query['nota_fiscal']

        # Verificação cruzada
        df_validado['encontrado'] = df_validado['chave'].isin(df_query['chave']).map({True: "Foi encontrado", False: "Não foi encontrado"})
        df_query['encontrado'] = df_query['chave'].isin(df_validado['chave']).map({True: "Foi encontrado", False: "Não foi encontrado"})

        st.success("Análise concluída!")
        
        # Cálculo dos percentuais de notas encontradas
        percent_encontrado_validado = (
            df_validado['encontrado'].value_counts(normalize=True)
            .get('Foi encontrado', 0) * 100
        )

        percent_encontrado_query = (
            df_query['encontrado'].value_counts(normalize=True)
            .get('Foi encontrado', 0) * 100
        )

        # Exibição dos indicadores
        st.subheader("📊 Indicadores de Conciliação")

        col1, col2 = st.columns(2)

        with col1:
            st.metric(
                label="Notas Liquidadas Encontradas",
                value=f"{percent_encontrado_validado:.2f}%"
            )

        with col2:
            st.metric(
                label="Notas de Recebimento Encontradas",
                value=f"{percent_encontrado_query:.2f}%"
            )

        # Função para converter para download
        def to_excel_bytes(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.download_button("⬇️ Baixar Resultado - Conciliação em relação às Notas Liquidadas", to_excel_bytes(df_validado), file_name="resultado_rma.xlsx")
        st.download_button("⬇️ Baixar Resultado - Conciliação em relação às Notas de Recebimento", to_excel_bytes(df_query), file_name="resultado_query.xlsx")

    except Exception as e:
        st.error(f"Erro ao processar os arquivos: {e}")

else:
    st.info("Envie os dois arquivos XLSX na ordem correta para iniciar a análise.")


