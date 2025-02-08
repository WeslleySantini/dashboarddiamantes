import pandas as pd
import streamlit as st
import os
from datetime import datetime

# Definir nome da aba no navegador
st.set_page_config(page_title="Envio de Diamantes", page_icon="logo.png")

# Aplicar estilo nos botões
st.markdown("""
    <style>
        div.stButton > button:first-child, .stDownloadButton > button:first-child {
            background-color: #004A82;
            color: white;
            font-weight: bold;
            border-radius: 5px;
            padding: 8px 16px;
        }
        div.stButton > button:first-child:hover, .stDownloadButton > button:first-child:hover {
            background-color: #003366;
        }
    </style>
""", unsafe_allow_html=True)

# Função para gerar o nome do arquivo
def generate_filename():
    return f"diamantes_clubes_{datetime.today().strftime('%d-%m-%Y')}.xlsx"

# Carregar os dados salvos
file_path = "dados_diamantes.csv"
if "diamantes" not in st.session_state:
    if os.path.exists(file_path):
        st.session_state["diamantes"] = pd.read_csv(file_path)
    else:
        st.session_state["diamantes"] = pd.DataFrame(columns=["DATA", "HORÁRIO", "ID DO CLUBE", "NOME DO CLUBE", "QUANTIDADE", "VALOR", "RESPONSÁVEL"])

# Interface do Streamlit
st.image("logo.png", width=200)
st.title("📊 Dashboard de Envio de Diamantes")
st.markdown("Gerencie e registre o envio de diamantes para clubes de forma profissional!")

# Criar inputs para os dados
data = st.date_input("**Data**", value=datetime.today())
horario = st.time_input("**Horário**", value=datetime.now().time())
id_clube = st.text_input("**ID do Clube**", value="")
nome_clube = st.text_input("**Nome do Clube**", value="")
quantidade = st.number_input("**Quantidade**", min_value=0, step=1)
valor = st.text_input("**Valor (R$)**", value=f"R$ {quantidade * 0.10:.2f}", disabled=True)
responsavel = st.text_input("**Responsável**", value="")

# Botão para adicionar entrada de diamantes
if st.button("**Adicionar Envio**"):
    if id_clube and nome_clube and quantidade and valor and responsavel:
        try:
            valor_float = float(valor.replace("R$", "").replace(",", ".").strip())
            novo_dado = pd.DataFrame([[data, horario, id_clube, nome_clube, quantidade, valor_float, responsavel]],
                                     columns=["DATA", "HORÁRIO", "ID DO CLUBE", "NOME DO CLUBE", "QUANTIDADE", "VALOR", "RESPONSÁVEL"])
            st.session_state["diamantes"] = pd.concat([st.session_state["diamantes"], novo_dado], ignore_index=True)
            st.session_state["diamantes"].to_csv(file_path, index=False)
            st.success("Envio de diamantes registrado com sucesso!")
            st.rerun()
        except ValueError:
            st.error("Por favor, insira um valor válido para o campo Valor.")
    else:
        st.error("Todos os campos devem ser preenchidos.")

# Exibir os envios cadastrados
st.write("### 📋 Envios de Diamantes Registrados")
st.dataframe(st.session_state["diamantes"])

# Botão para excluir um envio específico
if not st.session_state["diamantes"].empty:
    excluir_index = st.number_input("Digite o índice do envio para excluir", min_value=0, max_value=len(st.session_state["diamantes"])-1, step=1)
    if st.button("**Excluir Envio**"):
        st.session_state["diamantes"] = st.session_state["diamantes"].drop(excluir_index).reset_index(drop=True)
        st.session_state["diamantes"].to_csv(file_path, index=False)
        st.success("Envio de diamantes excluído com sucesso!")
        st.rerun()

# Botão para limpar todos os envios sem confirmação
if st.button("**Limpar Todos os Envios**"):
    st.session_state["diamantes"] = pd.DataFrame(columns=["DATA", "HORÁRIO", "ID DO CLUBE", "NOME DO CLUBE", "QUANTIDADE", "VALOR", "RESPONSÁVEL"])
    if os.path.exists(file_path):
        os.remove(file_path)
    st.success("Todos os envios de diamantes foram removidos!")
    st.rerun()

# Botão para baixar a planilha com formatação
if not st.session_state["diamantes"].empty:
    filename = generate_filename()
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        st.session_state["diamantes"].to_excel(writer, index=False, sheet_name="Envios de Diamantes")
        worksheet = writer.sheets["Envios de Diamantes"]
        workbook = writer.book
        
        # Aplicando formatação ao cabeçalho
        header_format = workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#92D050",
            "border": 1
        })
        for col_num, value in enumerate(st.session_state["diamantes"].columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Ajustando colunas e centralizando texto
        center_format = workbook.add_format({"align": "center"})
        currency_format = workbook.add_format({"align": "center", "num_format": "R$ #,##0.00"})
        worksheet.set_column("A:A", 15, center_format)
        worksheet.set_column("B:B", 12, center_format)
        worksheet.set_column("C:C", 12, center_format)
        worksheet.set_column("D:D", 25, center_format)
        worksheet.set_column("E:E", 12, center_format)
        worksheet.set_column("F:F", 12, currency_format)
        worksheet.set_column("G:G", 15, center_format)
        
        writer.close()
    with open(filename, "rb") as file:
        st.download_button(
            label="**Baixar Planilha de Envios**",
            data=file,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
