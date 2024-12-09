from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import gdown
from fpdf import FPDF
import io
from io import BytesIO
from datetime import datetime
from dotenv import load_dotenv
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

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

# Carregar variáveis de ambiente
load_dotenv()

# Configurações de e-mail
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")  # Nome da variável no arquivo .env
SENHA_EMAIL = os.getenv("SENHA_EMAIL")      # Nome da variável no arquivo .env

# Verificar se as variáveis foram carregadas corretamente
if not EMAIL_REMETENTE or not SENHA_EMAIL:
    raise ValueError("As variáveis de ambiente EMAIL_REMETENTE ou SENHA_EMAIL não foram configuradas corretamente.")

# Email fixo do destinatário
EMAIL_DESTINATARIO = "sreg.gapls@fab.mil.br"

def gerar_nome_anexo(prefixo="anexo"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{prefixo}_{timestamp}.pdf"

def testar_conexao_email():
    try:
        servidor = smtplib.SMTP("smtp.mail.yahoo.com", 587)
        servidor.starttls()
        servidor.login(EMAIL_REMETENTE, SENHA_EMAIL)
        print("Conexão bem-sucedida!")
        servidor.quit()
    except Exception as e:
        print(f"Erro na conexão SMTP: {e}")

# Chamar a função para testar
testar_conexao_email()
                
def enviar_email(destinatario, assunto, corpo, arquivo_anexo, nome_anexo="anexo.pdf"):
    try:
        # Configurar o servidor SMTP
        servidor = smtplib.SMTP("smtp.mail.yahoo.com", 587)
        servidor.set_debuglevel(1)  # Habilita detalhes de depuração
        servidor.starttls()
        servidor.login(EMAIL_REMETENTE, SENHA_EMAIL)

        # Configurar a mensagem
        mensagem = MIMEMultipart()
        mensagem["From"] = EMAIL_REMETENTE
        mensagem["To"] = destinatario
        mensagem["Subject"] = assunto
        mensagem.attach(MIMEText(corpo, "plain"))

        # Adicionar anexo
        if arquivo_anexo:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(arquivo_anexo.read())
            encoders.encode_base64(parte)
            parte.add_header(
                "Content-Disposition",
                f"attachment; filename={nome_anexo}",
            )
            mensagem.attach(parte)

        # Enviar o e-mail
        servidor.send_message(mensagem)
        servidor.quit()
        print("E-mail enviado com sucesso!")
        return True
    except Exception as e:
        print(f"Erro ao enviar o e-mail: {e}")
        return False

    destinatario = dados.get("email_destinatario", "")
    if not destinatario:
        print("Endereço de e-mail do destinatário está vazio!")
        
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

    # Envio de e-mail
    destinatario = dados.get("email_destinatario", "")
    if destinatario:
        sucesso = enviar_email(
            destinatario,
            "Guia de Movimentação de Bens",
            "Segue anexo o PDF gerado com as informações solicitadas.",
            pdf_output,
            nome_anexo=nome_anexo
        )
        if not sucesso:
            return jsonify({"error": "Erro ao enviar o e-mail."}), 500
       
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
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM MÓVEL PERMANENTE ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)

    def add_table(self, dados_bmps):
        col_widths = [20, 95, 50, 30]
        headers = ["Nº BMP", "Nomenclatura", "Nº Série", "Valor Atualizado"]

        # Adicionar cabeçalho da tabela
        self.set_font("Arial", "B", 10)
        for width, header in zip(col_widths, headers):
            self.cell(width, 10, header, border=1, align="C")
        self.ln()

        # Adicionar as linhas da tabela
        self.set_font("Arial", size=8)
        line_height = self.font_size + 2  # Define a altura da linha com base no tamanho da fonte

        for _, row in dados_bmps.iterrows():
            # Calcular a altura necessária para a célula "NOMECLATURA/COMPONENTE"
            text = self.fix_text(row["NOMECLATURA/COMPONENTE"])
            line_count = self.get_string_width(text) // col_widths[1] + 1
            row_height = line_height * line_count  # Altura ajustada ao tamanho do texto

            # Calcular a altura das outras células na mesma linha, baseando-se no valor máximo
            # Para a célula "Nº BMP"
            row_height_bmp = line_height  # Pode ser ajustado se necessário
            # Para a célula "Nº SERIE"
            row_height_serie = line_height  # Pode ser ajustado se necessário
            # Para a célula "VL. ATUALIZ."
            row_height_valor = line_height  # Pode ser ajustado se necessário

            # Definir a altura final da linha (a maior altura entre as células)
            row_height = max(row_height, row_height_bmp, row_height_serie, row_height_valor)
            
            # Adicionar célula "Nº BMP"
            self.cell(col_widths[0], row_height, str(row["Nº BMP"]), border=1, align="C")

            # Adicionar célula "NOMECLATURA/COMPONENTE" com quebra automática
            x, y = self.get_x(), self.get_y()
            self.multi_cell(col_widths[1], line_height, text, border=1)
            self.set_xy(x + col_widths[1], y)  # Reposicionar para a próxima coluna

            # Usar self.set_xy() para garantir que a próxima célula não sobreponha a anterior
            self.set_xy(x + col_widths[1], y)  # Reposicionar para a próxima coluna

            # Adicionar célula "Nº SERIE"
            self.cell(
                col_widths[2],
                row_height,
                self.fix_text(str(row["Nº SERIE"]) if pd.notna(row["Nº SERIE"]) else ""),
                border=1,
                align="C"
            )

            # Adicionar célula "VL. ATUALIZ."
            self.cell(
                col_widths[3],
                row_height,
                f"R$ {row['VL. ATUALIZ.']:.2f}".replace('.', ','),
                border=1,
                align="R"
            )

            self.ln()  # Próxima linha
                       
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

@app.route('/TTAC_apontamentos', methods=['GET'])
def TTAC_apontamentos_form():
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

@app.route('/validar_dados1', methods=['POST'])
def validar_dados1():
    dados = request.json
    # Valide os dados conforme necessário
    if not dados.get('secao_origem') or not dados.get('chefia_origem'):
        return jsonify({"error": "Dados de origem incompletos"}), 400
    if not dados.get('secao_destino') or not dados.get('chefia_destino'):
        return jsonify({"error": "Dados de destino incompletos"}), 400
    return jsonify({"message": "Dados válidos"})

# Rota para geração do PDF
@app.route('/gerar_ttac', methods=['POST'])
def gerar_ttac():
    dados = request.json

    # Imprime os dados recebidos para depuração
    print("Dados recebidos para gerar o PDF:", dados)

    if not dados:
        return jsonify({"error": "Nenhum dado enviado para gerar o PDF."}), 400

    # Cria o PDF completo usando a classe PDF
    pdf = PDF()
    pdf.add_page()
    pdf.add_details(
        secao_destino=dados.get('secao_destino'),
        chefia_origem=dados.get('chefia_origem'),
        secao_origem=dados.get('secao_origem'),
        chefia_destino=dados.get('chefia_destino')
    )

    # Salva o PDF localmente para depuração
    debug_pdf_path = "debug_TTAC_completo.pdf"
    pdf.output(debug_pdf_path)
    print(f"PDF salvo para depuração: {debug_pdf_path}")

    # Gera o PDF em memória para envio
    pdf_output = BytesIO()
    pdf_output.write(pdf.output(dest='S').encode('latin1'))  # Corrige para o formato correto
    pdf_output.seek(0)

    # Envio de e-mail
    destinatario = dados.get("email_destinatario", "")
    if destinatario:
        sucesso = enviar_email(
            destinatario,
            "Termo de Transmissão e Assunção ",
            "Segue anexo o PDF gerado com as informações solicitadas.",
            pdf_output,
            nome_anexo=nome_anexo
        )
        if not sucesso:
            return jsonify({"error": "Erro ao enviar o e-mail."}), 500
       
    # Retorna o PDF para o usuário
    return send_file(
        pdf_output,
        as_attachment=True,
        download_name='TTAC_completo.pdf',
        mimetype='application/pdf'
    )

class PDFTTAC(FPDF):
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
        self.set_font("Arial", "B", 12)
        self.cell(0, 6, "MINISTÉRIO DA DEFESA", ln=True, align="C")
        self.cell(0, 6, "COMANDO DA AERONÁUTICA", ln=True, align="C")
        self.cell(0, 6, "GRUPAMENTO DE APOIO DE LAGOA SANTA", ln=True, align="C")
        self.cell(0, 8, "GUIA DE MOVIMENTAÇÃO DE BEM MÓVEL PERMANENTE ENTRE AS SEÇÕES DO GAPLS", ln=True, align="C")
        self.ln(10)
                          
    def add_details(self, secao_destino, chefia_origem, secao_origem, chefia_destino):
        self.set_font("Arial", size=12)
        self.ln(10)
        text = f"""
Do(a) {chefia_origem}
À Dirigente Máximo

1. PASSAGEM

Comunico ao senhor que, em cumprimento à publicação constante do boletim interno {boletim_tipo} nº {numero_boletim}, de {data_boletim}; transmiti o cargo de Chefe da {secao_origem} ao meu substituto legal a quem fiz a entrega:

a) dos bens móveis permanentes, constantes da relação expedida pela Seção de Registro Patrimonial;
b) dos bens móveis de consumo de uso duradouro constantes da relação expedida pela Seção de Registro Patrimonial;
c) dos bens móveis de consumo em estoque;
d) do material em uso na Seção;
e) dos dois últimos Relatórios de Auditoria do CENCIAR;
f) dos balancetes referentes aos últimos dez anos; e
g) outros.

{data_confecção}

{chefia_origem}

2. RECEBIMENTO

Ao assumir o cargo de Chefe da {secao_origem}, declaro para todos os efeitos legais ter recebido tudo que consta do presente documento, inclusive os dois últimos Relatórios do CENCIAR, {restricoes}.

{data_confecção}

{chefia_origem}

3. ENCAMINHAMENTO

Encaminho ao senhor o presente Termo, informando que a escrituração se encontra em ordem e em dia, e que foram cumpridas as prescrições das normas vigentes.

{apontamentos}

{data_confecção}

KARINA RAQUEL VALIMAREANU  Maj Int
Chefe da ACI

4. SOLUÇÃO:

Por terem sido cumpridas todas as formalidades legais, determino que o presente termo seja transcrito, na íntegra, em boletim interno ({boletim_tipo}).

{data_confecção}

LUCIANA DO AMARAL CORREA  Cel Int
Dirigente Máximo
"""

@app.route('/PROCESS_TTAC_apontamentos', methods=['POST'])
def PROCESS_TTAC_apontamentos():
    # Receber e processar os dados do formulário
    form_data = request.form
    print("Dados recebidos:", form_data)  # Para depuração
    pdf = PDF()
    pdf.add_page()
    
    # Salvar PDF localmente para debug
    pdf.output("debug_TTAC.pdf")

    # Retornar o PDF
    pdf_output = BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)

    return send_file(
        pdf_output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="ttac.pdf"
    )    

@app.route("/consulta_bmp", methods=["GET", "POST"])
def consulta_bmp():
    results = pd.DataFrame()
    if request.method == "POST":
        search_query = request.form.get("bmp_query", "").strip().lower()
        if search_query:
            # Divide os BMPs por vírgula e remove espaços em excesso
            search_terms = [term.strip() for term in search_query.split(",") if term.strip()]
            # Filtra o dataframe para linhas que contenham qualquer um dos termos
            results = df[df['Nº BMP'].astype(str).str.lower().apply(
                lambda x: any(term in x for term in search_terms)
            )]
    return render_template("consulta_bmp.html", results=results)

@app.route("/")
def menu_principal():
    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
