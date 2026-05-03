function calculateTip(event) {
    event.preventDefault();
    
    let bill = document.getElementById('bill').value;
    let serviceQual = document.getElementById('serviceQual').value;
    let numOfPeople = document.getElementById('people').value;

    // Validação dos campos
    if (bill === '' || bill <= 0 || serviceQual === '0') {
        alert("Por favor, preencha todos os valores corretamente");
        return;
    }

    // Converte para números
    numOfPeople = parseInt(numOfPeople);
    
    if (numOfPeople <= 0) {
        numOfPeople = 1;
    }

    // Calcula a gorjeta total
    let totalTip = parseFloat(bill) * parseFloat(serviceQual);
    
    // Calcula a gorjeta por pessoa
    let tipPerPerson = totalTip / numOfPeople;
    
    // Formata para 2 casas decimais
    tipPerPerson = tipPerPerson.toFixed(2);
    totalTip = totalTip.toFixed(2);
    
    // Atualiza a exibição
    document.getElementById('tip').innerHTML = tipPerPerson;
    document.getElementById('totalTip').innerHTML = totalTip;
    
    // Mostra ou esconde a gorjeta total baseado no número de pessoas
    if (numOfPeople > 1) {
        document.getElementById('totalBox').style.display = "block";
    } else {
        document.getElementById('totalBox').style.display = "none";
    }
    
    document.getElementById('totaltip').style.display = "block";
}

function exportToExcel() {
    let bill = document.getElementById('bill').value;
    let serviceQual = document.getElementById('serviceQual').value;
    let numOfPeople = document.getElementById('people').value;
    let tipPerPerson = document.getElementById('tip').innerHTML;
    let totalTip = document.getElementById('totalTip').innerHTML;

    // Verifica se o resultado foi calculado
    if (document.getElementById('totaltip').style.display === 'none') {
        alert("Por favor, calcule a gorjeta primeiro!");
        return;
    }

    // Obtém a descrição da qualidade do serviço
    let serviceOptions = {
        '0.3': 'Incrível (30%)',
        '0.2': 'Bom (20%)',
        '0.15': 'Aceitável (15%)',
        '0.1': 'Ruim (10%)',
        '0.05': 'Péssimo (5%)'
    };
    let serviceDesc = serviceOptions[serviceQual];

    // Cria os dados para a planilha
    let data = [
        ['CALCULADORA DE GORJETAS', ''],
        ['', ''],
        ['DADOS DA CONTA', ''],
        ['Valor da Conta', 'R$ ' + parseFloat(bill).toFixed(2)],
        ['Qualidade do Serviço', serviceDesc],
        ['Número de Pessoas', numOfPeople],
        ['', ''],
        ['RESULTADOS', ''],
        ['Gorjeta por Pessoa', 'R$ ' + tipPerPerson],
        ['Gorjeta Total', 'R$ ' + totalTip],
        ['', ''],
        ['Data/Hora', new Date().toLocaleString('pt-BR')]
    ];

    // Cria a planilha
    let worksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Define as larguras das colunas
    worksheet['!cols'] = [
        { wch: 25 },
        { wch: 20 }
    ];

    // Estilo básico para as células (XLSX tem limitações em estilos)
    // Vamos melhorar a formatação manualmente
    let cellFormat = {
        font: { bold: true },
        fill: { fgColor: { rgb: "FFD966" } }
    };

    // Cria um novo workbook
    let workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Gorjeta");

    // Gera o nome do arquivo com data/hora
    let fileName = 'Gorjeta_' + new Date().toISOString().split('T')[0] + '_' + 
                   new Date().getHours() + '-' + 
                   String(new Date().getMinutes()).padStart(2, '0') + '.xlsx';

    // Faz o download
    XLSX.writeFile(workbook, fileName);
    
    alert("Arquivo exportado com sucesso!");
}

// Esconde o resultado inicialmente
document.getElementById("totaltip").style.display = "none";

// Adiciona o listener de submit ao formulário
document.getElementById('tipsForm').addEventListener('submit', calculateTip);

// Adiciona o listener do botão de exportação
document.getElementById('exportBtn').addEventListener('click', exportToExcel);