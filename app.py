from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import gdown
from fpdf import FPDF
import os
import io

app = Flask(__name__)

# ID do arquivo no Google Drive
GOOGLE_DRIVE_FILE_ID = "1mPYlc_uC3SfJnNQ_ToG6eVmn2ZYMhPCX"

def get_excel_from_google_drive():
    """Baixa a planilha do Google Drive e retorna o DataFrame."""
    url = f"https://drive.google.com/uc?id={GOOGLE_DRIVE_FILE_ID}"
    output_file = "patrimonio.xlsx"  # Nome temporário do arquivo baixado
    gdown.download(url, output_file, quiet=False)
    return pd.read_excel(output_file)

# Carregar a planilha no início do programa
df = get_excel_from_google_drive()

class PDF(FPDF):
    def __init__(self):
        super().__init__('P', 'mm', 'A4')  # Orientação retrato, milímetros, formato A4

    def fix_text(self, text):
        """Corrige caracteres incompatíveis com a codificação latin-1."""
        replacements = {
            "–": "-",  # Substituir travessão por hífen
            "“": '"',  # Substituir aspas abertas por aspas duplas
            "”": '"',  # Substituir aspas fechadas por aspas duplas
            "’": "'",  # Substituir apóstrofo por aspas simples
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text
    
    def header(self):
        # Adiciona o brasão no topo
        self.image('brasao.png', x=7, y=5, w=20)  # Ajuste x, y e w conforme necessário
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM MÓVEL PERMANENTE ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)

def add_table(self, dados):
    col_widths = [25, 70, 55, 35]
    headers = ["Nº BMP", "Nomenclatura", "Nº Série", "Valor Atualizado"]
    
    # Adicionar cabeçalho da tabela
    self.set_font("Arial", "B", 10)
    for width, header in zip(col_widths, headers):
        self.cell(width, 10, header, border=1, align="C")
    self.ln()

    # Adicionar as linhas da tabela
    self.set_font("Arial", size=10)
    for _, row in dados.iterrows():
        text = self.fix_text(row["NOMECLATURA/COMPONENTE"])
        max_chars = int(col_widths[1] / 3)  # Aprox. número de caracteres por linha
        line_count = len(text) // max_chars + 1
        row_height = 10 * line_count

        self.cell(col_widths[0], row_height, str(row["Nº BMP"]), border=1, align="C")
        
        # Multi-cell para a coluna "Nomenclatura"
        x, y = self.get_x(), self.get_y()
        self.multi_cell(col_widths[1], 10, text, border=1)
        self.set_xy(x + col_widths[1], y)  # Posiciona na próxima célula

        self.cell(col_widths[2], row_height, self.fix_text(str(row["Nº SERIE"])), border=1, align="C")
        self.cell(col_widths[3], row_height, f"R$ {row['VL. ATUALIZ.']:.2f}".replace('.', ','), border=1, align="R")
        self.ln()
        
    def add_details(self, secao_destino, chefia_origem, secao_origem, chefia_destino):
        self.set_font("Arial", size=12)
        self.ln(10)
        text = f"""
Solicitação de Transferência:
Informo à Senhora Chefe do GAP-LS que os bens especificados estão inservíveis para uso neste setor, classificados como ociosos, recuperáveis, reparados ou novos - aguardando distribuição. Diante disso, solicito autorização para transferir o(s) Bem(ns) Móvel(is) Permanente(s) acima discriminado(s), atualmente sob minha guarda, para a Seção {secao_destino}.

{chefia_origem}
{secao_origem}

Confirmação da Seção de Destino:
Estou ciente da movimentação informada acima e, devido à necessidade do setor, solicito à Senhora Dirigente Máximo autorização para manter sob minha guarda os Bens Móveis Permanentes especificados.

{chefia_destino}
{secao_destino}

DO AGENTE DE CONTROLE INTERNO AO DIRIGENTE MÁXIMO
Informo à Senhora que, após conferência, foi verificado que esta guia cumpre o disposto no Módulo D do RADA-e e, conforme a alínea "d" do item 5.3 da ICA 179-1, encaminho para apreciação e se for o caso, autorização.

KARINA RAQUEL VALIMAREANU  Maj Int
Chefe da ACI

DESPACHO DA AGENTE DIRETOR
Autorizo a movimentação solicitada e determino:
1. Que a Seção de Registro realize a movimentação no SILOMS.
2. Que a Seção de Registro publique a movimentação no próximo aditamento a ser confeccionado, conforme o item 2.14.2, Módulo do RADA-e.
3. Que os detentores realizem a movimentação física do(s) bem(ns).

LUCIANA DO AMARAL CORREA  Cel Int
Dirigente Máximo
"""
        self.multi_cell(0, 8, self.fix_text(text))

# Mock para as funções de validação
def validar_bmps(dados_bmps):
    """
    Valida se os BMPs fornecidos são válidos.
    Retorna True se todos os BMPs forem válidos e uma lista de erros se houver problemas.
    """
    erros = []
    for bmp in dados_bmps:
        if not bmp.isdigit():  # Exemplo simples: verifica se é numérico
            erros.append(f"BMP '{bmp}' não é um número válido.")
        elif int(bmp) <= 0:
            erros.append(f"BMP '{bmp}' deve ser maior que zero.")
    return True if not erros else erros

def validar_campos_obrigatorios(campos):
    """
    Verifica se os campos obrigatórios foram preenchidos.
    """
    erros = []
    for campo, valor in campos.items():
        if not valor:
            erros.append(f"O campo '{campo}' é obrigatório.")
    return True if not erros else erros

@app.route("/guia_bens", methods=["GET", "POST"])
def guia_bens():
    if request.method == "GET":
        secao_origem = df["Seção de Origem"].dropna().unique().tolist()
        secoes_destino = df["Seção de Destino"].dropna().unique().tolist()
        return render_template("guia_bens.html", secao_origem=secao_origem, secoes_destino=secoes_destino)

    elif request.method == "POST":
        # Receber dados enviados via JSON
        data = request.get_json(silent=True)
        if not data:
            return jsonify({"error": "Os dados enviados não estão no formato JSON!"}), 400

        # Validação de campos obrigatórios
        campos_obrigatorios = {
            "bmp_numbers": data.get("bmp_numbers"),
            "secao_origem": data.get("secao_origem"),
            "secao_destino": data.get("secao_destino"),
            "chefia_origem": data.get("chefia_origem"),
            "chefia_destino": data.get("chefia_destino")
        }
        erros = validar_campos_obrigatorios(campos_obrigatorios)
        if erros is not True:
            return jsonify({"error": erros}), 400

         # Simulação de retorno para validar o envio (ajuste conforme necessário)
        return jsonify({"success": "Dados recebidos com sucesso!"})

        # Validar BMPs
        bmp_list = [bmp.strip() for bmp in data["bmp_numbers"] if bmp.strip()]
        erros_bmps = validar_bmps(bmp_list)
        if erros_bmps is not True:
            return jsonify({"error": erros_bmps}), 400

        # Filtrar dados dos BMPs
        dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]
        if dados_bmps.empty:
            return jsonify({"error": "Nenhum BMP válido encontrado."}), 400
      
    # Extrair os dados do JSON
    bmp_numbers = data.get("bmp_numbers", [])
    secao_origem = data.get("secao_origem", "").strip()
    chefia_origem = data.get("chefia_origem", "").strip()
    secao_destino = data.get("secao_destino", "").strip()
    chefia_destino = data.get("chefia_destino", "").strip()
    
    # Validar a conta dos BMPs
    if not dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").any():
        return render_template(
            "guia_bens.html",
            secao_origem=secao_origem,
            secao_destino=secao_destino,
            error="Nenhum BMP válido encontrado."
        )

@app.route("/gerar_guia", methods=['POST'])
def gerar_guia():
    try:
        # Obtém os dados enviados na requisição
        data = request.get_json()

        if not data:
            return jsonify({"error": "Nenhum dado recebido"}), 400

        # Dados para gerar o PDF
        dados_bmps = data.get("bmp_numbers", [])
        secao_destino = data["secao_destino"]
        chefia_origem = data["chefia_origem"]
        secao_origem = data["secao_origem"]
        chefia_destino = data["chefia_destino"]

        # Geração do PDF em memória
        pdf_buffer = io.BytesIO()
        pdf = PDF()  # Sua classe de PDF
        pdf.add_page()
        pdf.add_table()
        pdf.add_details(
            secao_destino=secao_destino,
            chefia_origem=chefia_origem,
            secao_origem=secao_origem,
            chefia_destino=chefia_destino
        )
        pdf.output(pdf_buffer)
        pdf_buffer.seek(0)  # Retorna ao início para leitura

        # Enviar o PDF gerado como resposta
        return send_file(
            pdf_buffer,
            mimetype='application/pdf',
            as_attachment=True,
            download_name="guia_circulacao_interna.pdf"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
@app.route("/autocomplete", methods=["POST"])
def autocomplete():
    data = request.get_json()
    bmp_numbers = data.get("bmp_numbers", [])

    if not bmp_numbers:
        return jsonify({"error": "Nenhum BMP fornecido!"}), 400

    response = {}
    for bmp in bmp_numbers:
        filtro_bmp = df[df["Nº BMP"].astype(str) == bmp]
        if not filtro_bmp.empty:
            secao_origem = filtro_bmp["Seção de Origem"].values[0]
            chefia_origem = filtro_bmp["Chefia de Origem"].values[0]
            response[bmp] = {
                "secao_origem": secao_origem,
                "chefia_origem": chefia_origem
            }
        else:
            response[bmp] = {"secao_origem": "", "chefia_origem": ""}

    return jsonify(response)

@app.route("/get_chefia", methods=["POST"])
def get_chefia():
    data = request.json
    secao = data.get("secao")
    tipo = data.get("tipo")

    if tipo == "destino":
        chefia = df[df['Seção de Destino'] == secao]['Chefia de Destino'].dropna().unique()
    elif tipo == "origem":
        chefia = df[df['Seção de Origem'] == secao]['Chefia de Origem'].dropna().unique()
    else:
        return jsonify({"error": "Tipo inválido!"}), 400

    return jsonify({"chefia": chefia.tolist()})

@app.route("/consulta_bmp", methods=["GET", "POST"])
def consulta_bmp():
    results = pd.DataFrame()
    if request.method == "POST":
        search_query = request.form.get("bmp_query", "").strip().lower()
        if search_query:
            results = df[df['Nº BMP'].astype(str).str.lower().str.contains(search_query)]
    return render_template("consulta_bmp.html", results=results)

@app.route("/")
def menu_principal():
    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
