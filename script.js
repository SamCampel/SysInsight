//processar dados da planilha
document.getElementById('process-data').addEventListener('click', function () {
    const fileInput = document.getElementById('file-input');
    const file = fileInput.files[0];

    if (!file) {
        alert('Faça o upload da planilha.');
        return;
    }

    const reader = new FileReader();

    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        //callback console
        console.log("Leu mano");

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        //converte a planilha pra json e faz debug
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        console.log("Conteúdo da planilha:", jsonData);

        processSheetData(jsonData);

        const metricsSection = document.getElementById('metrics-section');
        metricsSection.classList.remove('hidden');
    };

    reader.readAsArrayBuffer(file);
});

function processSheetData(data) {
    const atendentesData = {};//dados dos atendentes

    //lendo a planilha prr
    data.forEach(row => {
        console.log("Processando linha:", row);

        const atendente = row['Atendente'];
        
        if (!atendentesData[atendente]) {
            atendentesData[atendente] = {
                cont: 0, 
                tempo: 0, 
                demandasEmAberto: 0
            };
        }

        //excel serial pra Date
        const inicioAtendimento = row['Início do atendimento'] ? 
            new Date((row['Início do atendimento'] - 25569) * 86400 * 1000) : null; 

        const finalAtendimento = row['Final do atendimento'] ? 
            new Date((row['Final do atendimento'] - 25569) * 86400 * 1000) : null;

        if (!inicioAtendimento || isNaN(inicioAtendimento)) return;

        let tempoAtendimento = 0;
        if (finalAtendimento) {
            tempoAtendimento = (finalAtendimento - inicioAtendimento) / (1000 * 60); 
        } else {
            tempoAtendimento = 0; //demands em aberto
        }

        atendentesData[atendente].cont++;
        atendentesData[atendente].tempo += tempoAtendimento;
        if (!finalAtendimento) {
            atendentesData[atendente].demandasEmAberto++;
        }
    });

    const metricsSection = document.getElementById('metrics-section');
    metricsSection.classList.remove('hidden');

    const metricsContainer = document.getElementById('metrics');
    metricsContainer.innerHTML = '';

    //métrica por atendente
    for (const atendente in atendentesData) {
        const { cont, tempo, demandasEmAberto } = atendentesData[atendente];
        const tempoMedio = cont > 0 ? (tempo / cont).toFixed(2) : 0;

        const atendenteMetrics = document.createElement('p');
        atendenteMetrics.innerHTML = `
            <strong>Atendente ${atendente}:</strong><br>
            Demandas: ${cont} | Tempo Médio: ${tempoMedio}min | Demandas em Aberto: ${demandasEmAberto}
        `;
        metricsContainer.appendChild(atendenteMetrics);
    }
};