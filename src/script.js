
document.getElementById('dataForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Impede o envio padrão do formulário

    const maquina = document.getElementById('maquina').value.trim();

    if (!maquina) {
        alert("Por favor, preencha o campo 'Máquina'.");
        return;
    }

    // Usa requestAnimationFrame para garantir que a tabela esteja completamente renderizada
    requestAnimationFrame(() => {
        // Coleta todos os dados da tabela
        const rows = Array.from(document.querySelectorAll('table tbody tr'));
    

        const ws_data = [
            ["Elemento (BASE)", "CM (g/m²)", "WW (%)", "Ret. (%)", "Peso Seco (g/m²)", "Consistência (%)", "Peso de Fibras Drenado (g/m²)", "Massa Total (g/m²)", "Quant. água retirada (g/m²)", "Quant. água retirada (l/min)", "Porcentual Drenado (%)", "Eficiência de Drenagem (%)"]
        ];

        console.log('Dados da tabela:', rows.map(row => {
            const base = row.querySelector('td').textContent.trim();
            const cells = Array.from(row.querySelectorAll('td input')).map(input => input.value.trim());
            return [base, ...cells];
        }));

        // Adiciona as linhas da tabela
    rows.forEach(row => {
        const base = row.querySelector('td').textContent.trim(); // Obtém o valor da primeira célula
        const cells = Array.from(row.querySelectorAll('td input')).map(input => input.value.trim());
        const rowData = [base, ...cells]; // Adiciona o valor da base e os dados das células
        ws_data.push(rowData);
    });

    // Cria uma nova planilha
    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    ws['!cols'] = [
        { width: 16 }, // "Elemento (BASE)"
        { width: 10 }, // "CM (g/m²)"
        { width: 10 }, // "WW (%)"
        { width: 10 }, // "Ret. (%)"
        { width: 18 }, // "Peso Seco (g/m²)"
        { width: 18 }, // "Consistência (%)"
        { width: 26 }, // "Peso de Fibras Drenado (g/m²)"
        { width: 20 }, // "Massa Total (g/m²)"
        { width: 26 }, // "Quant. água retirada (g/m²)"
        { width: 26 }, // "Quant. água retirada (l/min)"
        { width: 24 }, // "Porcentual Drenado (%)"
        { width: 26 }  // "Eficiência de Drenagem (%)"
    ];

     
    

    // Cria uma nova pasta de trabalho e adiciona a planilha a ela
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dados");

    // Usa o valor do campo "Máquina" como o nome do arquivo
    const fileName = `${maquina}_${new Date().toISOString().split('T')[0]}.xlsx`;

    // Gera o arquivo e inicia o download
    XLSX.writeFile(wb, fileName);
});
});

