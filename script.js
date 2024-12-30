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

        // Exibir o botão "Enviar métricas aos gestores"
        const sendMetricsButton = document.getElementById('send-metrics');
        sendMetricsButton.classList.remove('hidden');
    };

    reader.readAsArrayBuffer(file);
});

function processSheetData(data) {
    const atendentesData = {};//dados dos atendentes

    //lendo a planilha
    data.forEach(row => {
        console.log("Processando linha:", row);

        const atendente = row['Atendente'];
        const emailGestor = row['e-mail gestor'];  // Leitura da coluna "e-mail gestor"
        
        if (!atendentesData[atendente]) {
            atendentesData[atendente] = {
                cont: 0, 
                tempo: 0, 
                demandasEmAberto: 0,
                emailGestor: emailGestor  // Armazenando o e-mail do gestor
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
            tempoAtendimento = 0; //demandas em aberto
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
        const { cont, tempo, demandasEmAberto, emailGestor } = atendentesData[atendente];
        const tempoMedio = cont > 0 ? (tempo / cont).toFixed(2) : 0;

        const atendenteMetrics = document.createElement('p');
        atendenteMetrics.innerHTML = `
            <strong>Atendente ${atendente}:</strong><br>
            Gestor: ${emailGestor}<br>
            Demandas: ${cont} | Tempo Médio: ${tempoMedio}min | Demandas em Aberto: ${demandasEmAberto}
        `;
        metricsContainer.appendChild(atendenteMetrics);
    }
};

document.getElementById('send-metrics').addEventListener('click', function() {
    alert('As métricas foram enviadas aos gestores pelos e-mails que estavam na planilha.');
});
