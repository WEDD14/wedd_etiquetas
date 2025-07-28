
import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime

# --- Função para gerar etiquetas com layout final ---
def gerar_etiquetas_pdf(cliente, descricao, referencia, quantidade_texto, total_etiquetas):
    pdf = FPDF("P", "mm", "A4")
    rosa_claro = (255, 220, 230)
    margem_lateral = 25
    largura_caixa = 160

    for i in range(total_etiquetas):
        pdf.add_page()

        # Logótipo grande (imagem real a usar na prática)
        pdf.image("logo_wedd.jpg", x=margem_lateral, y=10, w=60)

        # Título mais abaixo
        pdf.set_font("Arial", "B", 18)
        pdf.set_xy(margem_lateral, 60)
        pdf.cell(w=largura_caixa, h=12, txt="ETIQUETA DE PALETE", ln=1, align="C")

        pdf.ln(8)
        pdf.set_x(margem_lateral)
        pdf.set_font("Arial", "", 12)
        pdf.cell(50, 10, "Cliente:", ln=0)
        pdf.cell(0, 10, cliente, ln=1)
        pdf.set_x(margem_lateral)
        pdf.cell(50, 10, "Código do Produto:", ln=0)
        pdf.cell(0, 10, str(referencia), ln=1)

        # Produto
        pdf.set_fill_color(*rosa_claro)
        pdf.ln(4)
        pdf.set_x(margem_lateral)
        pdf.set_font("Arial", "B", 14)
        pdf.cell(w=largura_caixa, h=10, txt="Produto:", ln=1)

        pdf.set_font("Arial", "", 12)
        pdf.set_x(margem_lateral)
        pdf.multi_cell(w=largura_caixa, h=10, txt=descricao, fill=True)

        # Quantidade
        pdf.ln(4)
        pdf.set_x(margem_lateral)
        pdf.set_font("Arial", "B", 14)
        pdf.cell(w=largura_caixa, h=10, txt="Quantidade:", ln=1)

        pdf.set_x(margem_lateral)
        pdf.set_font("Arial", "B", 28)
        pdf.set_fill_color(*rosa_claro)
        pdf.cell(w=largura_caixa, h=20, txt=quantidade_texto, ln=1, align="C", fill=True)

        # Lote
        pdf.ln(6)
        pdf.set_x(margem_lateral)
        lote = datetime.now().strftime("%Y-%m-%d %H:%M")
        pdf.set_font("Arial", "B", 14)
        pdf.set_fill_color(*rosa_claro)
        pdf.cell(w=largura_caixa, h=12, txt=f"Lote: {lote}", ln=1, fill=True)

        # Campos manuais
        pdf.ln(10)
        pdf.set_x(margem_lateral)
        pdf.set_font("Arial", "", 12)
        pdf.cell(w=largura_caixa, h=10, txt="Expedição palete nº: ____________________________", ln=1)
        pdf.set_x(margem_lateral)
        pdf.cell(w=largura_caixa, h=10, txt="Verificado por:     ____________________________", ln=1)

    return pdf.output(dest="S").encode("latin-1")

# --- Carregar dados do Excel ---
xls = pd.ExcelFile("Livro1.xlsm")
clientes = xls.parse(xls.sheet_names[1])
produtos = xls.parse(xls.sheet_names[2])

# Limpar dados
clientes = clientes.dropna(subset=["nome1"])
produtos = produtos.dropna(subset=["Descrição", "Referência"])

st.title("🎯 Gerador de Etiquetas WEDD — Layout Final")

# Dropdowns
cliente = st.selectbox("Seleciona o Cliente:", options=clientes["nome1"].tolist())
produto_desc = st.selectbox("Seleciona o Produto:", options=produtos["Descrição"].tolist())
referencia = produtos.loc[produtos["Descrição"] == produto_desc, "Referência"].values[0]

# Input manual
quantidade_texto = st.text_input("Indica a Quantidade a imprimir na etiqueta (ex: 12):", "1")

# Número de etiquetas
qtd_etiquetas = st.number_input("Quantas etiquetas deseja imprimir?", min_value=1, max_value=100, value=1, step=1)

# Pré-visualização
st.subheader("🖨 Pré-visualização da Etiqueta")
st.markdown(f"""
**Cliente:** {cliente}  
**Código do Produto:** {referencia}  
**Produto:** {produto_desc}  
**Quantidade:** {quantidade_texto or "_(a preencher)_"}  
**Lote:** _(gerado automaticamente)_  
**Expedição palete nº:** _(manual)_  
**Verificado por:** _(manual)_
""")

# Gerar PDF
if st.button("Gerar PDF com Etiquetas"):
    if not quantidade_texto.strip():
        st.warning("Por favor, introduz a quantidade antes de gerar o PDF.")
    else:
        pdf_bytes = gerar_etiquetas_pdf(cliente, produto_desc, referencia, quantidade_texto, qtd_etiquetas)
        st.download_button(
            label=f"📥 Descarregar {qtd_etiquetas} Etiqueta(s) PDF",
            data=pdf_bytes,
            file_name="etiquetas_paletes_wedd_final.pdf",
            mime="application/pdf"
        )
