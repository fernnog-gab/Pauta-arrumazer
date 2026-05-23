document.addEventListener("DOMContentLoaded", () => {
    
    // Elementos da UI
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const loadingState = document.getElementById("loading-state");
    const actionsZone = document.getElementById("actions-zone");
    const btnDocx = document.getElementById("btn-download-docx");
    const btnHtml = document.getElementById("btn-download-html");
    const btnRestart = document.getElementById("btn-restart");

    let processedDocxBlob = null;
    let generatedHtmlString = "";

    // Eventos Drag & Drop
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });
    dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });

    dropZone.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });
    btnRestart.addEventListener("click", () => window.location.reload());

    async function handleFile(file) {
        if (!file.name.endsWith(".docx")) return alert("Apenas arquivos .docx são permitidos.");
        
        dropZone.classList.add("hidden");
        loadingState.classList.remove("hidden");
        
        setTimeout(async () => {
            try {
                const arrayBuffer = await file.arrayBuffer();
                await processDocxCore(arrayBuffer);
                loadingState.classList.add("hidden");
                actionsZone.classList.remove("hidden");
            } catch (error) {
                console.error("Erro no processamento:", error);
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
        
        const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));
        
        let extrairParaHtml = [];
        let isHeader = true;
        let isDeletandoLixo = false;

        for (const p of paragraphs) {
            const textContent = p.textContent.trim();

            // 1. Destruição de Cabeçalho Institucional
            if (isHeader) {
                if (textContent.toUpperCase().includes("FEITOS DE COMPETÊNCIA RECURSAL")) {
                    isHeader = false; 
                    extrairParaHtml.push(textContent);
                } else {
                    p.parentNode.removeChild(p); 
                }
                continue;
            }

            // 2. Destruição de linhas vazias extras
            if (!textContent) {
                p.parentNode.removeChild(p); 
                continue;
            }

            const isProcesso = /^\d+\s*-/.test(textContent);
            const isRelator = /^RELATOR[A]?:/i.test(textContent);

            // 3. Regra de Negócio: Guardar ou Deletar Blocos
            if (isProcesso) {
                isDeletandoLixo = false; 
                extrairParaHtml.push(textContent);
                
                // Injetar borda superior fina e cinza direto no XML do DOCX
                let pPr = p.getElementsByTagName("w:pPr")[0];
                if (!pPr) {
                    pPr = xmlDoc.createElement("w:pPr");
                    p.insertBefore(pPr, p.firstChild);
                }
                const pBdr = xmlDoc.createElement("w:pBdr");
                const topBdr = xmlDoc.createElement("w:top");
                topBdr.setAttribute("w:val", "single");
                topBdr.setAttribute("w:sz", "4"); // Espessura bem fina
                topBdr.setAttribute("w:space", "10");
                topBdr.setAttribute("w:color", "D3D3D3"); // Cinza Claro
                
                pBdr.appendChild(topBdr);
                pPr.appendChild(pBdr);

            } else if (isRelator) {
                extrairParaHtml.push(textContent);
                // Ativa a exclusão em massa para limpar plurais e quebras de linha
                isDeletandoLixo = true; 

            } else if (isDeletandoLixo) {
                // Apaga advogados, partes, testemunhas, textos longos, etc.
                p.parentNode.removeChild(p); 
            } else {
                // Margem de segurança: Remove sujeiras perdidas entre o Processo e o Relator
                p.parentNode.removeChild(p);
            }
        }

        // Empacota de volta para DOCX
        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);
        loadedZip.file("word/document.xml", newXmlString);
        processedDocxBlob = await loadedZip.generateAsync({ type: "blob" });

        // Gera a versão HTML da Pauta
        generateHtmlTemplate(extrairParaHtml);
    }

    function generateHtmlTemplate(linhas) {
        // Separa o título (FEITOS DE COMPETÊNCIA...)
        const titulo = linhas.shift(); 
        
        // Constrói os cards para HTML
        let htmlCards = "";
        for (let i = 0; i < linhas.length; i += 2) {
            const numeroProcesso = linhas[i] || "";
            const relatorProcesso = linhas[i + 1] || "";
            
            htmlCards += `
                <div class="processo-card">
                    <div class="processo-numero">${numeroProcesso}</div>
                    <div class="processo-relator">${relatorProcesso}</div>
                </div>
            `;
        }

        generatedHtmlString = `
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <title>Pauta de Julgamento - Resumida</title>
                <style>
                    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background: #fff; color: #333; }
                    .container { max-width: 800px; margin: 0 auto; padding: 2cm; }
                    .titulo { text-align: center; font-size: 14pt; font-weight: bold; margin-bottom: 25px; color: #1e293b; text-transform: uppercase; }
                    .processo-card { border-top: 1px solid #e2e8f0; padding: 12px 0; page-break-inside: avoid; }
                    .processo-numero { font-weight: 600; font-size: 11pt; color: #0f172a; margin-bottom: 4px; }
                    .processo-relator { font-size: 10pt; color: #475569; }
                    
                    @page { size: A4 portrait; margin: 1.5cm; }
                    @media print {
                        body { filter: grayscale(100%); }
                        .container { padding: 0; width: 100%; max-width: 100%; }
                        .titulo { color: #000; }
                        .processo-card { border-top: 1px solid #cbd5e1; }
                        .processo-numero { color: #000; }
                        .processo-relator { color: #333; }
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="titulo">${titulo}</div>
                    ${htmlCards}
                </div>
                <script>window.print();</script>
            </body>
            </html>
        `;
    }

    // Handlers de Download
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
        setTimeout(() => URL.revokeObjectURL(url), 1000); 
    }
});
