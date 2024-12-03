document.addEventListener("DOMContentLoaded", () => {
    const secaoOrigemSelect = document.getElementById("secao_origem");
    const secaoDestinoSelect = document.getElementById("secao_destino");
    const errorMessage = document.getElementById("error-message");

    // Função para preencher as opções das seções
    fetch("/gerar_guia")
        .then(response => response.json())
        .then(data => {
            data.secoes_origem.forEach(secao => {
                const option = document.createElement("option");
                option.value = secao;
                option.textContent = secao;
                secaoOrigemSelect.appendChild(option);
            });

            data.secoes_destino.forEach(secao => {
                const option = document.createElement("option");
                option.value = secao;
                option.textContent = secao;
                secaoDestinoSelect.appendChild(option);
            });
        })
        .catch(error => {
            console.error("Erro ao carregar seções:", error);
            errorMessage.textContent = "Erro ao carregar as seções. Tente novamente mais tarde.";
        });

    // Submeter o formulário via POST
    const form = document.getElementById("guiaForm");
    form.addEventListener("submit", event => {
        event.preventDefault();

        const formData = new FormData(form);
        const data = Object.fromEntries(formData.entries());

        fetch("/gerar_guia", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(data)
        })
            .then(response => response.json())
            .then(result => {
                if (result.error) {
                    errorMessage.textContent = result.error;
                } else {
                    alert("Guia gerada com sucesso!");
                }
            })
            .catch(error => {
                console.error("Erro ao gerar guia:", error);
                errorMessage.textContent = "Erro ao gerar guia. Tente novamente.";
            });
    });
});
