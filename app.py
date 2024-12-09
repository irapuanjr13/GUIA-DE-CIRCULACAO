from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import gdown
from fpdf import FPDF
import os
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

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

data_secoes_origem = df["Seção de Origem"].dropna().unique().tolist()
data_secoes_destino = df["Seção de Destino"].dropna().unique().tolist()
data_dados_bmps = df[df["Nº BMP"].astype(str).isin]

# Função para enviar e-mail
def enviar_email(pdf_bytes, destinatarios, assunto="Guia de Movimentação"):
    remetente = "jreletricidade@yahoo.com"  # Use seu email do Yahoo
    senha = "exlsxslvktxfphha"  # Use variáveis de ambiente para maior segurança

    # Configurar a mensagem do e-mail
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatarios)  # Múltiplos destinatários separados por vírgula
    msg['Subject'] = assunto

    # Corpo do e-mail
    corpo = "Segue em anexo a Guia de Movimentação de Bem Móvel Permanente."
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar o PDF
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(pdf_bytes.getvalue())
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename=guia_bens.pdf',
    )
    msg.attach(part)

    # Enviar o e-mail via servidor SMTP do Yahoo
    with smtplib.SMTP('smtp.mail.yahoo.com', 587) as server:
        server.starttls()  # Inicia a criptografia TLS
        server.login(remetente, senha)
        server.send_message(msg)
        print("E-mail enviado com sucesso!")

        
@app.route("/guia_bens", methods=["GET", "POST"])
def guia_bens():
    if request.method == "GET":
        secoes_origem = df["Seção de Origem"].dropna().unique().tolist()
        secoes_destino = df["Seção de Destino"].dropna().unique().tolist()
        return render_template("guia_bens.html", secoes_origem=secoes_origem, secoes_destino=secoes_destino)

    elif request.method == "POST":
        try:
            dados = request.json
            if not dados:
                return jsonify({"error": "Dados inválidos ou ausentes."}), 400

            campos_obrigatorios = ["bmp_numbers", "secao_origem", "secao_destino", "chefia_origem", "chefia_destino"]
            for campo in campos_obrigatorios:
                if not dados.get(campo):
                    return jsonify({"error": f"O campo '{campo}' é obrigatório."}), 400

            bmp_list = [bmp.strip() for bmp in dados["bmp_numbers"]]
            dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]
            if dados_bmps.empty:
                return jsonify({"error": "Nenhum BMP válido encontrado."}), 400

            
            if not dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").any():
                return render_template(
                    "guia_bens.html",
                    secoes_origem=secoes_origem,
                    secoes_destino=secoes_destino,
                    error="Estes itens pertencem à conta '87 - MATERIAL DE CONSUMO DE USO DURADOURO'."
                )

            # Processamento do PDF ou outros passos podem ir aqui
            return jsonify({"message": "Dados processados com sucesso."})

        except Exception as e:
            return jsonify({"error": str(e)}), 500
               
@app.route("/autocomplete", methods=["POST"])
def autocomplete():
    dados = request.get_json()
    bmp_numbers = dados.get("bmp_numbers", [])

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
    dados = request.json
    secao = dados.get("secao")
    tipo = dados.get("tipo")

    if tipo == "destino":
        chefia = df[df['Seção de Destino'] == secao]['Chefia de Destino'].dropna().unique()
    elif tipo == "origem":
        chefia = df[df['Seção de Origem'] == secao]['Chefia de Origem'].dropna().unique()
    else:
        return jsonify({"error": "Tipo inválido!"}), 400

    return jsonify({"chefia": chefia.tolist()})

@app.route('/validar_dados', methods=['POST'])
def validar_dados():
    dados = request.json
    # Valide os dados conforme necessário
    if not dados.get('secao_origem') or not dados.get('chefia_origem'):
        return jsonify({"error": "Dados de origem incompletos"}), 400
    if not dados.get('secao_destino') or not dados.get('chefia_destino'):
        return jsonify({"error": "Dados de destino incompletos"}), 400
    return jsonify({"message": "Dados válidos"})

# Rota para geração do PDF
@app.route('/gerar_guia', methods=['POST'])
def gerar_guia():
    dados = request.json

    # Imprime os dados recebidos para depuração
    print("Dados recebidos para gerar o PDF:", dados)

    if not dados:
        return jsonify({"error": "Nenhum dado enviado para gerar o PDF."}), 400

     # Validação dos BMPs
    try:
        bmp_list = [str(bmp).strip() for bmp in dados.get("bmp_numbers", [])]
    except Exception as e:
        return jsonify({"error": f"Erro ao processar BMPs: {str(e)}"}), 400

    dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]
    if dados_bmps.empty:
        return jsonify({"error": "Nenhum BMP válido encontrado."}), 400

    # Cria o PDF completo usando a classe PDF
    pdf = PDF()
    pdf.add_page()
    pdf.add_table(dados_bmps)
    pdf.add_details(
        secao_destino=dados.get('secao_destino'),
        chefia_origem=dados.get('chefia_origem'),
        secao_origem=dados.get('secao_origem'),
        chefia_destino=dados.get('chefia_destino')
    )

    # Salva o PDF localmente para depuração
    debug_pdf_path = "debug_guia_bens_completo.pdf"
    pdf.output(debug_pdf_path)
    print(f"PDF salvo para depuração: {debug_pdf_path}")

    # Gera o PDF em memória para envio
    pdf_output = BytesIO()
    pdf_output.write(pdf.output(dest='S').encode('latin1'))  # Corrige para o formato correto
    pdf_output.seek(0)

     # Enviar o e-mail após gerar o PDF
    destinatarios = dados.get("SREG.GAPLS@FAB.MIL.BR, TP.IRAPUANIMFJ@FAB.MIL.BR, TP.DORNELASKWDD@FAB.MIL.BR" , [])
    if destinatarios:
        enviar_email(pdf_output, destinatarios)

    # Retorna o PDF para o usuário
    return send_file(
        pdf_output,
        as_attachment=True,
        download_name='guia_circulacao_interna_completo.pdf',
        mimetype='application/pdf'
    )

class PDF(FPDF):
    def __init__(self):
        super().__init__('P', 'mm', 'A4')

    def fix_text(self, text):
        """Corrige caracteres incompatíveis com a codificação latin-1."""
        replacements = {
            "–": "-", "“": '"', "”": '"', "’": "'"
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text

    def header(self):
        # Adicionar o brasão
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM MÓVEL PERMANENTE ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)

    def add_table(self, dados_bmps):
        col_widths = [25, 70, 55, 35]
        headers = ["Nº BMP", "Nomenclatura", "Nº Série", "Valor Atualizado"]

        # Adicionar cabeçalho da tabela
        self.set_font("Arial", "B", 10)
        for width, header in zip(col_widths, headers):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()

        # Adicionar as linhas da tabela
        self.set_font("Arial", size=10)
        for _, row in dados_bmps.iterrows():
            # Calcular a altura necessária para a célula "Nomenclatura"
            text = self.fix_text(row["NOMECLATURA/COMPONENTE"])
            line_count = self.get_string_width(text) // col_widths[1] + 1
            row_height = 10 * line_count  # 10 é a altura padrão da célula
            
            self.cell(col_widths[0], row_height, str(row["Nº BMP"]), border=1, align="C")

            x, y = self.get_x(), self.get_y()
            self.multi_cell(col_widths[1], 10, text, border=1)
            self.set_xy(x + col_widths[1], y)  # Reposicionar para a próxima coluna

            self.cell(col_widths[2], row_height, self.fix_text(str(row["Nº SERIE"])if pd.notna(row["Nº SERIE"]) else ""), border=1, align="C")
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

def gerar_guia_pdf(dados_bmps):
    pdf = PDF()
    pdf.add_page()
    pdf.add_table(dados_bmps)

    # Salvar PDF localmente para debug
    pdf.output("debug_guia_bens.pdf")

    # Retornar o PDF
    pdf_output = BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)

    return send_file(
        pdf_output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="guia_bens.pdf"
    )

@app.route("/guia_duradouro", methods=["GET", "POST"])
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM DE USO DURADOURO ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)
    def add_table(self, dados_bmps):
        col_widths = [15, 55, 15, 30, 30, 30]
        headers = ["Nº BMP", "Nomenclatura", "Qtde", "Valor Atualizado", "Qtde a Movimentar", "Valor a Movimentar"]
        # Renderizar cabeçalhos
        self.set_font("Arial", "B", 9)
        for header, width in zip(headers, col_widths):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()
        # Adicionar dados
        self.set_font("Arial", size=9)
        for _, row in dados_bmps.iterrows():
            self.cell(col_widths[0], 10, str(row["Nº BMP"]), border=1, align="C")
            self.cell(col_widths[1], 10, str(row["NOMECLATURA/COMPONENTE"]), border=1, align="L")
            self.cell(col_widths[2], 10, str(row["QTD"]), border=1, align="C")
            self.cell(col_widths[3], 10, f"R$ {row['VL. ATUALIZ.']:.2f}".replace(".", ","), border=1, align="C")
            self.cell(col_widths[4], 10, str(row["Qtde a Movimentar"]), border=1, align="C")
            self.cell(col_widths[5], 10, f"R$ {row['Valor a Movimentar']:.2f}".replace(".", ","), border=1, align="C")
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
DO AGENTE DE CONTROLE INTERNO AO DIRIGENTE MÁXIMO:
Informo à Senhora que, após conferência, foi verificado que esta guia cumpre o disposto no Módulo D do RADA-e e, conforme a alínea "d" do item 5.3 da ICA 179-1, encaminho para apreciação e se for o caso, autorização.
KARINA RAQUEL VALIMAREANU  Maj Int
Chefe da ACI
DESPACHO DA AGENTE DIRETOR:
Autorizo a movimentação solicitada e determino:
1. Que a Seção de Registro realize a movimentação no SILOMS.
2. Que a Seção de Registro publique a movimentação no próximo aditamento a ser confeccionado, conforme o item 2.14.2, Módulo do RADA-e.
3. Que os detentores realizem a movimentação física do(s) bem(ns).
LUCIANA DO AMARAL CORREA  Cel Int
Dirigente Máximo
"""
        self.multi_cell(0, 8, text)
        
@app.route("/autocoomplete1", methods=["GET", "POST"])
def autocoomplete1():
    secoes_origem = df["Seção de Origem"].dropna().unique().tolist()
    secoes_destino = df["Seção de Destino"].dropna().unique().tolist()
    
    if request.method == "POST":
        bmp_numbers = request.form.get("bmp_numbers", "").strip()
        secao_origem = request.form.get("secao_origem")
        secao_destino = request.form.get("secao_destino")
        chefia_origem = request.form.get("chefia_origem")
        chefia_destino = request.form.get("chefia_destino")
        quantidades_movimentadas = {}
        
        for key, value in request.form.items():
            if key.startswith("quantidade_"):
                bmp_key = key.split("_")[1]
                quantidades_movimentadas[bmp_key] = float(value) if value.strip() else 0
                
        bmp_list = [bmp.strip() for bmp in bmp_numbers.split(",") if bmp.strip()]
        dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]
        if dados_bmps.empty:
            return render_template(
                "guia_duradouro.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Nenhum BMP encontrado ou inválido."
            )
        if not dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").all():
            return render_template(
                "guia_duradouro.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Os itens não pertencem à conta '87 - MATERIAL DE CONSUMO DE USO DURADOURO'."
            )
        
        # Cálculo de valores
        dados_bmps["Qtde a Movimentar"] = dados_bmps["Nº BMP"].astype(str).map(quantidades_movimentadas).fillna(0)
        dados_bmps["Valor a Movimentar"] = dados_bmps.apply(
            lambda row: (row["VL. ATUALIZ."] / row["QTD"] * row["Qtde a Movimentar"]) if row["QTD"] > 0 else 0, axis=1
        )
        pdf = PDF()
        pdf.add_page()
        pdf.add_table(dados_bmps)
        pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)  # Adicionando os detalhes ao PDF
        output_path = "static/guia_circulacao_interna.pdf"
        pdf.output(output_path)
        return send_file(output_path, as_attachment=True)
    return render_template(
        "guia_duradouro.html", secoes_origem=secoes_origem, secoes_destino=secoes_destino
    )

@app.route("/get_chefia1", methods=["POST"])
def get_chefia1():
    data = request.json
    secao = data.get("secao")
    tipo = data.get("tipo")
    if tipo == "origem":
        chefia = df[df['Seção de Origem'] == secao]['Chefia de Origem'].dropna().unique()
    elif tipo == "destino":
        chefia = df[df['Seção de Destino'] == secao]['Chefia de Destino'].dropna().unique()
    else:
        return jsonify({"error": "Tipo inválido"}), 400
    return jsonify({"chefia": chefia[0] if len(chefia) > 0 else ""})

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
   
