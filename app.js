document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileInfo = document.getElementById('file-info');
    const fileNameDisplay = document.getElementById('file-name');
    const exportBtn = document.getElementById('export-btn');
    const statusMessage = document.getElementById('status-message');

    let currentFile = null;

    // Prevenir comportamentos padrão para drag and drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    // Efeitos visuais para drag and drop
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('dragover');
        }, false);
    });

    // Lidar com arquivos soltos
    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    });

    // Lidar com arquivos selecionados via click
    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        if (files.length > 0) {
            const file = files[0];
            
            // Validar extensão do arquivo
            const validExtensions = ['.xlsx', '.xls', '.csv'];
            const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
            
            if (validExtensions.includes(fileExtension)) {
                currentFile = file;
                fileNameDisplay.textContent = file.name;
                dropZone.style.display = 'none';
                fileInfo.style.display = 'block';
                showMessage('', '');
            } else {
                showMessage('Formato de arquivo inválido. Por favor, envie uma planilha Excel (.xlsx, .xls) ou CSV.', 'error');
            }
        }
    }

    exportBtn.addEventListener('click', () => {
        if (!currentFile) return;
        
        exportBtn.disabled = true;
        exportBtn.innerHTML = 'Processando...';
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                // Ler o arquivo Excel (isso processará fórmulas se cellFormula for default, e pegará o v, que é o valor final)
                const workbook = XLSX.read(data, {type: 'array', cellDates: true});
                
                // Pegar a primeira aba
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Recalcular o range (!ref) para garantir que pegue todos os dados,
                // pois alguns sistemas exportam Excel com o !ref errado (só no cabeçalho)
                let calcMaxRow = 0;
                let calcMaxCol = 0;
                let hasKeys = false;
                for (let key in worksheet) {
                    if (key[0] === '!') continue; // Ignorar metadados
                    hasKeys = true;
                    const cellAddress = XLSX.utils.decode_cell(key);
                    if (cellAddress.r > calcMaxRow) calcMaxRow = cellAddress.r;
                    if (cellAddress.c > calcMaxCol) calcMaxCol = cellAddress.c;
                }
                
                if (hasKeys) {
                    worksheet['!ref'] = XLSX.utils.encode_range({s: {c: 0, r: 0}, e: {c: calcMaxCol, r: calcMaxRow}});
                }
                
                // Converter para array de arrays
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                // Transformar em TXT com |
                let txtOutput = "";
                let headerLength = 0;
                
                jsonData.forEach((row, rowIndex) => {
                    if (!row || row.length === 0) return;
                    
                    // Verificar se a linha tem algum dado real
                    const hasData = row.some(cell => cell !== "" && cell !== null && cell !== undefined);
                    if (!hasData) return;
                    
                    // Definir o tamanho do cabeçalho pela primeira linha válida
                    if (headerLength === 0) {
                        headerLength = row.length;
                    }
                    
                    // Preencher a linha com valores nulos até o tamanho do cabeçalho
                    // para garantir a mesma quantidade de pipes (||)
                    const paddedRow = [];
                    for (let i = 0; i < headerLength; i++) {
                        paddedRow.push(row[i]);
                    }
                    
                    const selectedFormat = document.getElementById('format-select').value;
                    const formattedRow = paddedRow.map(cell => formatCell(cell, selectedFormat));
                    txtOutput += formattedRow.join('|') + '\r\n';
                });
                
                // Fazer download do arquivo TXT
                const originalName = currentFile.name.substring(0, currentFile.name.lastIndexOf('.'));
                downloadFile(`${originalName}_Convertido.txt`, txtOutput);
                
                showMessage('Arquivo convertido com sucesso!', 'success');
                
                // Restaurar botão
                exportBtn.disabled = false;
                exportBtn.innerHTML = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg> Exportar para TXT';
                
            } catch (error) {
                console.error("Erro ao processar:", error);
                showMessage('Erro ao ler a planilha. Verifique se o arquivo não está corrompido.', 'error');
                exportBtn.disabled = false;
                exportBtn.innerHTML = 'Tentar Novamente';
            }
        };
        
        reader.readAsArrayBuffer(currentFile);
    });

    function formatCell(val, selectedFormat) {
        if (val === null || val === undefined || val === '') return '';
        
        if (typeof val === 'number') {
            if (selectedFormat === 'modelo1') {
                // Modelo 1: Duas casas decimais forçadas (ex: 1234,00)
                return val.toFixed(2).replace('.', ',');
            } else {
                // Modelo 2: "Geral" mantemos o número como está e trocamos ponto por vírgula nos decimais
                return String(val).replace('.', ',');
            }
        }
        
        if (val instanceof Date) {
            const pad = n => n.toString().padStart(2, '0');
            if (selectedFormat === 'modelo1') {
                // Modelo 1: Data com Horário
                return `${pad(val.getDate())}/${pad(val.getMonth()+1)}/${val.getFullYear()} ${pad(val.getHours())}:${pad(val.getMinutes())}:${pad(val.getSeconds())}`;
            } else {
                // Modelo 2: Data Abreviada
                return `${pad(val.getDate())}/${pad(val.getMonth()+1)}/${val.getFullYear()}`;
            }
        }
        
        // Se for string ou outro formato, tratar espaços extras
        let strVal = String(val).trim();
        
        // Se a string já tiver formato de data não vamos colocar aspas duplas
        if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(strVal)) {
             return strVal;
        }
        
        // Qualquer outro texto ganha aspas
        return `"${strVal}"`;
    }

    function downloadFile(filename, content) {
        const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function showMessage(msg, type) {
        statusMessage.textContent = msg;
        statusMessage.className = 'status-message';
        if (type) {
            statusMessage.classList.add(type);
        }
    }
});
