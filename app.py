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

        if not (bmp_numbers and secao_origem and secao_destino and chefia_origem and chefia_destino):
            return render_template(
                "guia_bens.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Preencha todos os campos!",
            )

        bmp_list = [bmp.strip() for bmp in bmp_numbers.split(",")]
        dados_bmps = df[df["Nº BMP"].astype(str).isin(bmp_list)]
        if dados_bmps.empty:
            return render_template(
                "guia_bens.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Nenhum BMP encontrado para os números fornecidos.",
            )

        if dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").any():
            return render_template(
                "guia_bens.html",
                secoes_origem=secoes_origem,
                secoes_destino=secoes_destino,
                error="Itens da conta '87 - MATERIAL DE CONSUMO DE USO DURADOURO' não podem ser processados."
            )

        pdf = PDF()
        pdf.add_page()
        pdf.add_table(dados_bmps)
        pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)

        output_path = "static/guia_circulacao_interna.pdf"
        pdf.output(output_path)
        return send_file(output_path, as_attachment=True)

    return render_template(
        "guia_bens.html", secoes_origem=secoes_origem, secoes_destino=secoes_destino
    )

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

@app.route('/get_chefia', methods=['POST'])
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

@app.route("/gerar_pdf", methods=["POST"])
def gerar_pdf_geral():
    # Obtendo dados do formulário
    secao_origem = request.form.get('secao_origem')
    secao_destino = request.form.get('secao_destino')
    chefia_origem = request.form.get('chefia_origem')
    chefia_destino = request.form.get('chefia_destino')

    bmps_input = request.form.get('bmps')  # Recebe BMPs como string (ex: "123,456,789")
    if not bmps_input:
        return jsonify({"error": "Nenhum BMP fornecido!"}), 400

    bmp_list = [bmp.strip() for bmp in bmps_input.split(",")]

    # Filtra os dados dos BMPs no DataFrame
    dados_bmps = df[df["Nº BMP"].astype(str).str.strip().isin(bmp_list)]
    if dados_bmps.empty:
        return jsonify({"error": "Nenhum BMP encontrado!"}), 404

    # Verifica se há itens proibidos
    if dados_bmps["CONTA"].eq("87 - MATERIAL DE CONSUMO DE USO DURADOURO").any():
        return jsonify({"error": "Itens proibidos encontrados!"}), 403

    # Gera o PDF em memória
    pdf = PDF()
    pdf.add_page()
    pdf.add_table(dados_bmps)
    pdf.add_details(secao_destino, chefia_origem, secao_origem, chefia_destino)

    # Salva o PDF temporariamente para envio
    pdf_output = BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)

    # Envia o arquivo para download
    return send_file(
        pdf_output,
        as_attachment=True,
        download_name="guia_circulacao_interna.pdf",
        mimetype="application/pdf"
    )

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

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
