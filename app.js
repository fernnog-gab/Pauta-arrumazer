document.addEventListener("DOMContentLoaded", () => {
    // [Setup de UI omitido por brevidade - Mesma lógica de Drag/Drop e ocultar/mostrar seções]
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const loadingState = document.getElementById("loading-state");
    const actionsZone = document.getElementById("actions-zone");
    const btnDocx = document.getElementById("btn-download-docx");
    const btnHtml = document.getElementById("btn-download-html");
    const btnRestart = document.getElementById("btn-restart");

    let processedDocxBlob = null;
    let generatedHtmlString = "";

    // Regex Arquiteturalmente revisada:
    // ^\s* -> Garante que olha apenas o início da linha (tolerando espaços).
    // Mapeamento expandido com base no DOCX real.
    const regexExclusao = /^\s*(RECORRENTE|RECORRIDA|RECORRIDO|ADVOGADA|ADVOGADO|PERITO|PERITA|CUSTOS LEGIS|TERCEIRO|AGRAVANTE|AGRAVADA|AGRAVADO|TESTEMUNHA)[:\s]/i;

    // ... (Eventos de UI mantidos) ...
    dropZone.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });
    btnRestart.addEventListener("click", () => window.location.reload());

    async function handleFile(file) {
        if (!file.name.endsWith(".docx")) return alert("Apenas arquivos .docx");
        
        dropZone.classList.add("hidden");
        loadingState.classList.remove("hidden");
        
        // Uso de setTimeout para permitir que a UI renderize o estado de loading 
        // antes de bloquear a thread principal com o parser do XML.
        setTimeout(async () => {
            try {
                const arrayBuffer = await file.arrayBuffer();
                await processDocxCore(arrayBuffer);
                loadingState.classList.add("hidden");
                actionsZone.classList.remove("hidden");
            } catch (error) {
                console.error(error);
                alert("Falha ao processar arquivo. Verifique a integridade do DOCX.");
                window.location.reload();
            }
        }, 100);
    }

    async function processDocxCore(arrayBuffer) {
        const zip = new JSZip();
        const loadedZip = await zip.loadAsync(arrayBuffer);
        const documentXml = await loadedZip.file("word/document.xml").async("text");
        
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(documentXml, "application/xml");
        
        // Array.from congela a lista, permitindo deletar nós sem quebrar a iteração
        const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));
        
        let extrairParaHtml = [];
        let isCompetenciaRecursalIniciada = false;

        // Processamento O(n) estrito top-down
        for (const p of paragraphs) {
            const textContent = p.textContent.trim();
            if (!textContent) continue;

            // Gatilho da Máquina de Estado
            if (!isCompetenciaRecursalIniciada) {
                if (textContent.toUpperCase().includes("FEITOS DE COMPETÊNCIA RECURSAL")) {
                    isCompetenciaRecursalIniciada = true;
                }
                continue; // Pula o processamento até o gatilho
            }

            // Regra de Negócio: Deleta se bater com a Regex, senão preserva e joga pro HTML
            if (regexExclusao.test(textContent)) {
                p.parentNode.removeChild(p);
            } else {
                // Preserva no DOCX e salva para o HTML
                extrairParaHtml.push(textContent);
            }
        }

        // Serializa DOCX
        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);
        loadedZip.file("word/document.xml", newXmlString);
        processedDocxBlob = await loadedZip.generateAsync({ type: "blob" });

        // Gera HTML
        generateHtmlTemplate(extrairParaHtml);
    }

    function generateHtmlTemplate(linhas) {
        const tableRows = linhas.map(linha => {
            // Identifica se a linha é o início do processo (ex: "01 - PROCESSO...")
            const isProcesso = /^\d+\s*-/.test(linha); 
            return `
                <tr>
                    <td class="${isProcesso ? 'processo-header' : 'processo-relator'}">${linha}</td>
                </tr>
            `;
        }).join("");

        generatedHtmlString = `
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <style>
                    body { font-family: Arial, sans-serif; margin: 0; padding: 0; color: #000; }
                    table { width: 100%; border-collapse: collapse; }
                    td { padding: 4px 8px; font-size: 11pt; border-bottom: 1px solid #ddd; }
                    .processo-header { font-weight: bold; padding-top: 15px; border-bottom: none; font-size: 12pt; }
                    .processo-relator { padding-bottom: 15px; color: #333; }
                    @page { size: A4 portrait; margin: 2cm; }
                    @media print {
                        body { filter: grayscale(100%); }
                        td { page-break-inside: avoid; }
                    }
                </style>
            </head>
            <body>
                <table>
                    <tbody>${tableRows}</tbody>
                </table>
                <script>window.print();</script>
            </body>
            </html>
        `;
    }

    // Ações de download...
    btnDocx.addEventListener("click", () => triggerDownload(processedDocxBlob, "Pauta_Resumida.docx"));
    btnHtml.addEventListener("click", () => {
        const blob = new Blob([generatedHtmlString], { type: 'text/html' });
        triggerDownload(blob, "Pauta_Impressao.html");
    });

    function triggerDownload(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        a.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000); // Limpeza de memória segura
    }
});
