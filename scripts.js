function calculateTip(event) {
    event.preventDefault();
    
    let bill = document.getElementById('bill').value.trim();
    let serviceQual = document.getElementById('serviceQual').value;
    let numOfPeople = document.getElementById('people').value.trim();

    // Validação dos campos
    if (!bill || bill <= 0) {
        showError("Por favor, insira um valor válido para a conta");
        return;
    }

    if (serviceQual === '0') {
        showError("Por favor, selecione a qualidade do serviço");
        return;
    }

    // Converte para números
    numOfPeople = parseInt(numOfPeople);
    
    if (isNaN(numOfPeople) || numOfPeople <= 0) {
        showError("Por favor, insira um número válido de pessoas");
        return;
    }

    // Calcula a gorjeta total
    let totalTip = parseFloat(bill) * parseFloat(serviceQual);
    
    // Calcula a gorjeta por pessoa
    let tipPerPerson = totalTip / numOfPeople;
    
    // Formata para 2 casas decimais
    tipPerPerson = tipPerPerson.toFixed(2);
    totalTip = totalTip.toFixed(2);
    
    // Armazena os valores para exportação
    window.lastCalculation = {
        bill: parseFloat(bill).toFixed(2),
        serviceQual: serviceQual,
        numOfPeople: numOfPeople,
        tipPerPerson: tipPerPerson,
        totalTip: totalTip
    };
    
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
    showSuccess("Gorjeta calculada com sucesso!");
}

function showError(message) {
    alert("❌ " + message);
}

function showSuccess(message) {
    // Feedback visual (comentado para não interromper a experiência)
    console.log("✅ " + message);
}

function clearForm() {
    document.getElementById('tipsForm').reset();
    document.getElementById('totaltip').style.display = "none";
    document.getElementById('bill').focus();
}

function exportToExcel() {
    // Verifica se o resultado foi calculado
    if (!window.lastCalculation) {
        showError("Por favor, calcule a gorjeta primeiro!");
        return;
    }

    let { bill, serviceQual, numOfPeople, tipPerPerson, totalTip } = window.lastCalculation;

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
        ['Valor da Conta', 'R$ ' + bill],
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

    // Cria um novo workbook
    let workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Gorjeta");

    // Gera o nome do arquivo com data/hora
    let fileName = 'Gorjeta_' + new Date().toISOString().split('T')[0] + '_' + 
                   String(new Date().getHours()).padStart(2, '0') + '-' + 
                   String(new Date().getMinutes()).padStart(2, '0') + '.xlsx';

    // Faz o download
    XLSX.writeFile(workbook, fileName);
    
    showSuccess("Arquivo exportado com sucesso!");
}

// Esconde o resultado inicialmente
document.getElementById("totaltip").style.display = "none";

// Aguarda o carregamento completo do DOM
document.addEventListener('DOMContentLoaded', function() {
    // Adiciona o listener de submit ao formulário
    document.getElementById('tipsForm').addEventListener('submit', calculateTip);

    // Adiciona o listener do botão de exportação
    document.getElementById('exportBtn').addEventListener('click', exportToExcel);
    
    // Adiciona o listener do botão de limpar
    document.getElementById('clearBtn').addEventListener('click', clearForm);
    
    // Foca no primeiro campo ao carregar
    document.getElementById('bill').focus();
});

// Variável global para armazenar o último cálculo
window.lastCalculation = null;