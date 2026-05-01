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
                let headers = [];
                let isFirstRow = true;
                
                jsonData.forEach((row, rowIndex) => {
                    if (!row || row.length === 0) return;
                    
                    // Verificar se a linha tem algum dado real
                    const hasData = row.some(cell => cell !== "" && cell !== null && cell !== undefined);
                    if (!hasData) return;
                    
                    // Definir o tamanho e guardar os cabeçalhos para mapeamento de tipos
                    if (headerLength === 0) {
                        headerLength = row.length;
                        headers = row.map(h => String(h || '').trim().toUpperCase());
                    }
                    
                    // Preencher a linha com valores nulos até o tamanho do cabeçalho
                    // para garantir a mesma quantidade de pipes (||)
                    const paddedRow = [];
                    for (let i = 0; i < headerLength; i++) {
                        paddedRow.push(row[i]);
                    }
                    
                    const selectedFormat = document.getElementById('format-select').value;
                    const formattedRow = paddedRow.map((cell, colIndex) => {
                        const headerName = headers[colIndex];
                        return formatCell(cell, selectedFormat, headerName, isFirstRow);
                    });
                    txtOutput += formattedRow.join('|') + '\r\n';
                    isFirstRow = false;
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

    function formatCell(val, selectedFormat, headerName, isHeader) {
        if (val === null || val === undefined || val === '') return '';
        
        // Formatar objeto Data corretamente para evitar "Horário Padrão de Brasília"
        let isDateObj = val instanceof Date;
        let formattedDateStr = "";
        if (isDateObj) {
            const pad = n => n.toString().padStart(2, '0');
            if (selectedFormat === 'modelo1') {
                // A pedido, fixar o horário em 00:00:00 independentemente do horário interno do JS/Excel
                formattedDateStr = `${pad(val.getDate())}/${pad(val.getMonth()+1)}/${val.getFullYear()} 00:00:00`;
            } else {
                formattedDateStr = `${pad(val.getDate())}/${pad(val.getMonth()+1)}/${val.getFullYear()}`;
            }
        }
        
        if (selectedFormat === 'modelo1') {
            if (isHeader) {
                // No modelo Access, os cabeçalhos sempre ganham aspas
                return `"${String(val)}"`;
            }
            
            // Mapeamento de tipos exato como a exportação original do Access
            const isNumericCol = ['CPF', 'MATRICULAMF', 'CPF OPERADOR'].includes(headerName);
            const isDateCol = ['ADMISSAO', 'ADMISSÃO'].includes(headerName);
            
            if (isNumericCol) {
                let num = Number(val);
                if (!isNaN(num) && String(val).trim() !== '') {
                    // Exporta como ,00 sem aspas
                    return num.toFixed(2).replace('.', ',');
                }
                return String(val); // Fallback
            }
            
            if (isDateCol || isDateObj) {
                if (isDateObj) return formattedDateStr;
                return String(val); // Se a data veio como texto no excel
            }
            
            // Para as demais colunas (incluindo FAIXA), tratamos como Texto (com aspas)
            // Mesmo se no Excel estiver um número como 31 ou 61
            return `"${String(val)}"`;
            
        } else {
            // Opção 2: Modelo de formatação flexível (Geral e Data Abreviada)
            if (typeof val === 'number') {
                return String(val).replace('.', ',');
            }
            
            if (isDateObj) {
                return formattedDateStr;
            }
            
            let strVal = String(val);
            if (/^\s*\d{1,2}\/\d{1,2}\/\d{4}/.test(strVal)) {
                 return strVal;
            }
            
            return `"${strVal}"`;
        }
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
