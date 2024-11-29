from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import gdown
from fpdf import FPDF
import os

app = Flask(__name__)

# Classe PDF fictícia para o exemplo (substitua pela sua implementação)
class PDF:
    def add_page(self):
        print("Página adicionada ao PDF.")

    def add_table(self, data):
        print(f"Tabela adicionada com dados: {data}")

    def add_details(self, secao_destino, chefia_origem, secao_origem, chefia_destino):
        print(f"Detalhes adicionados: {secao_destino}, {chefia_origem}, {secao_origem}, {chefia_destino}")

    def output(self, path):
        print(f"PDF salvo em: {path}")


@app.route('/gerar_guia', methods=['GET', 'POST'])
def gerar_guia():
    try:
        # Para requisições GET
        if request.method == 'GET':
            return jsonify({
                "message": "Rota disponível para gerar guia via POST. Envie os dados no formato JSON.",
                "exemplo": {
                    "dados_bmps": [{"campo1": "valor1", "campo2": "valor2"}],
                    "secao_destino": "Seção Destino",
                    "chefia_origem": "Chefia Origem",
                    "secao_origem": "Seção Origem",
                    "chefia_destino": "Chefia Destino"
                }
            }), 200

        # Para requisições POST
        if request.method == 'POST':
            if not request.is_json:
                return jsonify({"error": "Conteúdo não é JSON. Verifique o cabeçalho 'Content-Type'."}), 415
            
            dados = request.json
            print(f"Dados recebidos: {dados}")

            # Extração dos dados do JSON
            dados_bmps = dados.get("dados_bmps")
            secao_destino = dados.get("secao_destino")
            chefia_origem = dados.get("chefia_origem")
            secao_origem = dados.get("secao_origem", "N/A")
            chefia_destino = dados.get("chefia_destino", "N/A")
            
            # Validações obrigatórias
            if not dados_bmps or not secao_destino or not chefia_origem:
                return jsonify({"error": "Dados obrigatórios ausentes no corpo da requisição!"}), 400

            # Geração do PDF
            print("Criando PDF...")
            pdf = PDF()
            pdf.add_page()
            pdf.add_table(dados_bmps)
            pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)

            # Gerenciamento de arquivos
            output_dir = "generated_pdfs"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            filename = f"guia_circulacao_{int(time.time())}.pdf"
            output_path = os.path.join(output_dir, filename)
            pdf.output(output_path)
            print(f"PDF gerado em {output_path}")
            
            # Retornar o arquivo gerado
            return send_file(output_path, as_attachment=True)

    except Exception as e:
        print(f"Erro ao gerar a guia: {e}")
        return jsonify({"error": f"Erro ao gerar a guia: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
