import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
import base64
import os
from PIL import Image

# Título da aplicação
st.title("📦 Estoque - Filtro por Percentual da Coluna PT")

# Upload do arquivo
uploaded_file = st.file_uploader("Envie o arquivo .xlsx ou .csv", type=["xlsx", "csv"])

# Upload da logo
logo_file = "logo_houston.png"  # Nome fixo do arquivo da logo

# Campo para porcentagem
percentual = st.number_input(
    "Digite o percentual (%) para filtrar os itens que representam até esse valor acumulado da coluna PT",
    min_value=0.0, max_value=100.0, step=0.1
)

# Botão para calcular
if uploaded_file and st.button("Filtrar e Gerar PDF"):
    # Ler o arquivo
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

    # Verifica se as colunas necessárias existem
    required_cols = ["CODIGO", "DESCRICAO", "QT", "CM", "PT"]
    if not all(col in df.columns for col in required_cols):
        st.error("O arquivo deve conter as colunas: CODIGO, DESCRICAO, QT, CM, PT")
    else:
        # Calcular o total da coluna PT
        total_pt = df["PT"].sum()

        # Calcular a porcentagem que cada item representa
        df["%_PT"] = (df["PT"] / total_pt) * 100

        # Ordenar os dados por percentual decrescente
        df_sorted = df.sort_values("%_PT", ascending=False)

        # Calcular a porcentagem acumulada
        df_sorted["%_ACUMULADO"] = df_sorted["%_PT"].cumsum()

        # Filtrar os itens cuja soma acumulada seja menor ou igual ao limite
        limite_acumulado = 100.0 - percentual
        filtrados = df_sorted[df_sorted["%_ACUMULADO"] <= limite_acumulado]

        # Calcular totais das colunas
        total_qt = filtrados["QT"].sum()
        total_pt = filtrados["PT"].sum()
        total_percentual = filtrados["%_PT"].sum()
        total_acumulado = filtrados["%_ACUMULADO"].max()

        # Adicionar linha de total geral
        total_row = pd.DataFrame({
            "CODIGO": ["TOTAL GERAL"],
            "DESCRICAO": [""],
            "QT": [total_qt],
            "CM": [""],
            "PT": [total_pt],
            "%_PT": [total_percentual],
            "%_ACUMULADO": [total_acumulado]
        })

        df_final = pd.concat([filtrados, total_row], ignore_index=True)

        # Exibir a tabela
        st.dataframe(df_final)

        # Botão para baixar em Excel
        output_excel = BytesIO()
        df_final.to_excel(output_excel, index=False, engine="openpyxl")
        st.download_button("📥 Baixar Excel", data=output_excel.getvalue(), file_name="resultado_filtrado.xlsx")

        # Botão para baixar em CSV
        #output_csv = df_final.to_csv(index=False).encode("utf-8")
        #st.download_button("📥 Baixar CSV", data=output_csv, file_name="resultado_filtrado.csv")

        # Gerar PDF com logotipo
        pdf_buffer = BytesIO()
        doc = fitz.open()
        page = doc.new_page(width=842, height=595)  # A4 landscape

        # Inserir logo no canto superior direito
        try:
            logo_pixmap = fitz.Pixmap(logo_file)
            logo_width = 100
            logo_height = int(logo_width * logo_pixmap.height / logo_pixmap.width)
            x_pos = 842 - logo_width - 20
            y_pos = 20
            page.insert_image(
                fitz.Rect(x_pos, y_pos, x_pos + logo_width, y_pos + logo_height),
                pixmap=logo_pixmap
            )
        except Exception as e:
            st.warning(f"Não foi possível inserir a logo: {e}")

  
        # Cabeçalho
        header = f"Estoque GP54 - Itens que representam até {100 - percentual:.2f}% do total da coluna PT\n"
        page.insert_text((150, 50), header, fontsize=12, fontname="helv", fill=(0, 0, 0))

        # Definir colunas para exibir
        columns = ["CODIGO", "DESCRICAO", "QT", "CM", "PT", "%_PT", "%_ACUMULADO"]
        x_positions = [50, 120, 400, 460, 520, 580, 660]

        # Inserir cabeçalhos
        y = 90
        for i, col in enumerate(columns):
            page.insert_text((x_positions[i], y), col, fontsize=10, fontname="helv", fill=(0, 0, 0))
        y += 15

        # Inserir dados linha por linha
        for idx, row in df_final.iterrows():
            highlight = (idx == len(df_final) - 1)
            for i, col in enumerate(columns):
                value = row[col]
                if isinstance(value, float):
                    text = f"{value:,.2f}"
                else:
                    text = str(value)
                color = (0, 0, 0)
                if highlight:
                    color = (0.2, 0.2, 0.6)  # azul escuro para total
                page.insert_text((x_positions[i], y), text, fontsize=9, fontname="helv", fill=color)
            y += 15
            if y > 550:
                page = doc.new_page(width=842, height=595)
                y = 50

        doc.save(pdf_buffer)
        st.download_button("📥 Baixar PDF", data=pdf_buffer.getvalue(), file_name="resultado_filtrado.pdf")