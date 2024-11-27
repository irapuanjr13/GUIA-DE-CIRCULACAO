from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import gdown
from fpdf import FPDF
import os

app = Flask(__name__)

# ID do arquivo no Google Drive
GOOGLE_DRIVE_FILE_ID = "1mPYlc_uC3SfJnNQ_ToG6eVmn2ZYMhPCX"

def get_excel_from_google_drive():
    """Baixa a planilha do Google Drive e retorna o DataFrame."""
    url = f"https://drive.google.com/uc?id={GOOGLE_DRIVE_FILE_ID}"
    output_file = "patrimonio.xlsx"
    gdown.download(url, output_file, quiet=False)
    return pd.read_excel(output_file)

# Carregar a planilha no início do programa
df = get_excel_from_google_drive()

@app.route("/")
def menu_principal():
    return render_template("index.html")

@app.route("/consulta_bmp", methods=["GET", "POST"])
def consulta_bmp():
    results = pd.DataFrame()
    if request.method == "POST":
        search_query = request.form.get("bmp_query", "").strip().lower()
        if search_query:
            results = df[df['Nº BMP'].astype(str).str.lower().str.contains(search_query)]
    return render_template("consulta_bmp.html", results=results)

@app.route("/guia_bens", methods=["GET", "POST"])
def guia_bens():
    secoes_origem = df['Seção de Origem'].dropna().unique().tolist()
    secoes_destino = df['Seção de Destino'].dropna().unique().tolist()

    if request.method == "POST":
        bmp_numbers = request.form.get("bmp_numbers")
        secao_origem = request.form.get("secao_origem")
        secao_destino = request.form.get("secao_destino")
        chefia_origem = request.form.get("chefia_origem")
        chefia_destino = request.form.get("chefia_destino")

        # Validar campos obrigatórios
        if not (bmp_numbers and secao_origem and secao_destino and chefia_origem and chefia_destino):
            return render_template(
                "guia_bens.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Preencha todos os campos obrigatórios!"
            )

        bmp_list = [bmp.strip() for bmp in bmp_numbers.split(",")]
        dados_bmps = df[df["Nº BMP"].astype(str).str.strip().isin(bmp_list)]

        # Verificar se os BMPs existem
        if dados_bmps.empty:
            return render_template(
                "guia_bens.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Nenhum BMP encontrado para os números fornecidos."
            )

        # Verificar se há itens proibidos
        if dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").any():
            return render_template(
                "guia_bens.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Itens da conta '87 - MATERIAL DE CONSUMO DE USO DURADOURO' não podem ser processados."
            )

        # Gerar o PDF
        output_path = gerar_pdf(dados_bmps, secao_destino, chefia_origem, secao_origem, chefia_destino)
        return send_file(output_path, as_attachment=True)

    return render_template(
        "guia_bens.html",
        secoes_origem=secoes_origem,
        secoes_destino=secoes_destino
    )

def gerar_pdf(dados_bmps, secao_destino, chefia_origem, secao_origem, chefia_destino):
    """Gera o PDF da guia de movimentação."""
    pdf = PDF()
    pdf.add_page()
    pdf.add_table(dados_bmps)
    pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)

    output_path = "static/guia_circulacao_interna.pdf"
    pdf.output(output_path)
    return output_path

class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM MÓVEL PERMANENTE", ln=True, align="C")
        self.ln(10)

    def add_table(self, dados_bmps):
        col_widths = [25, 70, 55, 35]
        headers = ["Nº BMP", "Nomenclatura", "Nº Série", "Valor Atualizado"]

        self.set_font("Arial", "B", 10)
        for width, header in zip(col_widths, headers):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()

        self.set_font("Arial", size=10)
        for _, row in dados_bmps.iterrows():
            self.cell(col_widths[0], 10, str(row["Nº BMP"]), border=1, align="C")
            self.cell(col_widths[1], 10, str(row["NOMECLATURA/COMPONENTE"]), border=1, align="C")
            self.cell(col_widths[2], 10, str(row["Nº SERIE"]), border=1, align="C")
            self.cell(col_widths[3], 10, f"R$ {row['VL. ATUALIZ.']:.2f}".replace('.', ','), border=1, align="R")
            self.ln()

    def add_details(self, secao_destino, chefia_origem, secao_origem, chefia_destino):
        text = f"""
Solicitação de Transferência:
Solicito a transferência dos Bens Móveis Permanentes para a Seção {secao_destino}.

{chefia_origem}
{secao_origem}

Confirmação da Seção de Destino:
Confirmo o recebimento.

{chefia_destino}
{secao_destino}
"""
        self.multi_cell(0, 8, text)

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
