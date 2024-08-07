function onSubmit(event) {
    event.preventDefault();
    const formData = new FormData(event.target);
    const data = {};

    formData.forEach((value, key) => {
        data[key] = value;
    });

    // Enviar dados para o Google Apps Script
    google.script.run.withSuccessHandler(function(response) {
        document.getElementById('responseMessage').innerText = response; // Mostra a URL do PDF
    }).withFailureHandler(function(error) {
        document.getElementById('responseMessage').innerText = 'Erro ao gerar PDF.';
        console.error('Error:', error);
    }).generatePdf(data);
}