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

    // Função utilitária pura para sanitização de strings
    function normalizeTextForMatch(str) {
        if (!str) return "";
        return str.trim()
                  .replace(/\s+/g, " ")
                  .normalize("NFC")
                  .toUpperCase();
    }

    async function processDocxCore(arrayBuffer) {
        const zip = new JSZip();
        const loadedZip = await zip.loadAsync(arrayBuffer);
        const documentXml = await loadedZip.file("word/document.xml").async("text");
        
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(documentXml, "application/xml");
        const paragraphs = Array.from(xmlDoc.getElementsByTagName("w:p"));
        
        let isHeader = true;
        let isDeletandoLixo = false;
        let capturandoAviso = false;

        let dataSessao = "";
        let avisoEspecial = [];
        let processosExtraidos = [];
        let processoAtual = null;

        for (const p of paragraphs) {
            const textContent = p.textContent.trim();
            if (!textContent) {
                p.parentNode.removeChild(p); 
                continue; 
            }

            const textNorm = normalizeTextForMatch(textContent);

            // Fase 1: Análise de Cabeçalho e Gatilhos Prévios
            if (isHeader) {
                if (textNorm.includes("(HORÁRIO OFICIAL DE MATO GROSSO) DE")) {
                    const match = textContent.match(/de\s+([0-9]{1,2}\s+de\s+[A-Za-zÀ-ÿ]+\s+de\s+[0-9]{4})/i);
                    if (match) dataSessao = match[1];
                }

                if (textNorm.includes("§4º DO ART. 937")) {
                    capturandoAviso = true;
                    p.parentNode.removeChild(p);
                    continue; 
                } 
                else if (capturandoAviso && !textNorm.includes("FEITOS DE COMPETÊNCIA RECURSAL")) {
                    if (textContent) avisoEspecial.push(textContent); 
                    p.parentNode.removeChild(p);
                    continue;
                }

                if (textNorm.includes("FEITOS DE COMPETÊNCIA RECURSAL")) {
                    isHeader = false; 

                    // 1. Centralizar a expressão "FEITOS DE COMPETÊNCIA RECURSAL" no DOCX
                    let pPr = p.getElementsByTagName("w:pPr")[0];
                    if (!pPr) {
                        pPr = xmlDoc.createElement("w:pPr");
                        p.insertBefore(pPr, p.firstChild);
                    }
                    let jc = pPr.getElementsByTagName("w:jc")[0];
                    if (!jc) {
                        jc = xmlDoc.createElement("w:jc");
                        pPr.appendChild(jc);
                    }
                    jc.setAttribute("w:val", "center");

                    // Criamos um ponteiro de referência para inserir os próximos parágrafos em ordem
                    let referenceNode = p;

                    // 2. Injetar a Data extraída no DOCX
                    if (dataSessao) {
                        const dateXml = `
                            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                <w:pPr>
                                    <w:jc w:val="center"/>
                                    <w:spacing w:before="120" w:after="240"/>
                                </w:pPr>
                                <w:r>
                                    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
                                    <w:t>${dataSessao}</w:t>
                                </w:r>
                            </w:p>
                        `;
                        const dateNode = parser.parseFromString(dateXml, "application/xml").documentElement;
                        const importedDate = xmlDoc.importNode(dateNode, true);
                        referenceNode.parentNode.insertBefore(importedDate, referenceNode.nextSibling);
                        referenceNode = importedDate; // Atualiza o ponteiro
                    }

                    // 3. Injetar o Quadro de Aviso no DOCX (somente se não estiver vazio)
                    if (avisoEspecial.length > 0) {
                        const avisoTexto = avisoEspecial.join(" "); // Junta as linhas em um único texto contínuo
                        const avisoXml = `
                            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                                <w:pPr>
                                    <w:pBdr>
                                        <w:top w:val="single" w:sz="4" w:space="6" w:color="000000"/>
                                        <w:left w:val="single" w:sz="4" w:space="6" w:color="000000"/>
                                        <w:bottom w:val="single" w:sz="4" w:space="6" w:color="000000"/>
                                        <w:right w:val="single" w:sz="4" w:space="6" w:color="000000"/>
                                    </w:pBdr>
                                    <w:jc w:val="center"/>
                                    <w:spacing w:before="240" w:after="400"/>
                                </w:pPr>
                                <w:r>
                                    <w:rPr><w:b/><w:sz w:val="20"/></w:rPr>
                                    <w:t>${avisoTexto}</w:t>
                                </w:r>
                            </w:p>
                        `;
                        const avisoNode = parser.parseFromString(avisoXml, "application/xml").documentElement;
                        const importedAviso = xmlDoc.importNode(avisoNode, true);
                        referenceNode.parentNode.insertBefore(importedAviso, referenceNode.nextSibling);
                    }

                } else {
                    p.parentNode.removeChild(p); 
                }
                continue;
            }

            // Fase 2: Identificação de Processos e Relatores (Inalterado da versão anterior)
            const isProcesso = /^\d+\s*-/.test(textContent);
            const isRelator = textNorm.startsWith("RELATOR") || textNorm.startsWith("RELATORA");

            if (isProcesso) {
                isDeletandoLixo = false; 
                processoAtual = { numero: textContent, relator: "" };
                processosExtraidos.push(processoAtual);
                
                let pPr = p.getElementsByTagName("w:pPr")[0];
                if (!pPr) {
                    pPr = xmlDoc.createElement("w:pPr");
                    p.insertBefore(pPr, p.firstChild);
                }
                const pBdr = xmlDoc.createElement("w:pBdr");
                const topBdr = xmlDoc.createElement("w:top");
                topBdr.setAttribute("w:val", "single");
                topBdr.setAttribute("w:sz", "4");
                topBdr.setAttribute("w:space", "10");
                topBdr.setAttribute("w:color", "D3D3D3"); 
                pBdr.appendChild(topBdr);
                pPr.appendChild(pBdr);

            } else if (isRelator) {
                if (processoAtual) {
                    processoAtual.relator = textContent;
                }
                isDeletandoLixo = true; 

                // Mantém a injeção dos quadrados Acompanhar/Divergir nativos no DOCX
                const actionsXml = `
                    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                        <w:pPr>
                            <w:jc w:val="right"/>
                            <w:spacing w:before="60" w:after="240"/>
                        </w:pPr>
                        <w:r>
                            <w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="555555"/></w:rPr>
                            <w:t xml:space="preserve">DIVERGIR  </w:t>
                        </w:r>
                        <w:r>
                            <w:rPr><w:sz w:val="40"/><w:color w:val="333333"/></w:rPr>
                            <w:t>☐</w:t>
                        </w:r>
                        <w:r>
                            <w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="555555"/></w:rPr>
                            <w:t xml:space="preserve">      ACOMPANHAR  </w:t>
                        </w:r>
                        <w:r>
                            <w:rPr><w:sz w:val="40"/><w:color w:val="333333"/></w:rPr>
                            <w:t>☐</w:t>
                        </w:r>
                    </w:p>
                `;
                const actionNode = parser.parseFromString(actionsXml, "application/xml").documentElement;
                p.parentNode.insertBefore(xmlDoc.importNode(actionNode, true), p.nextSibling);

            } else if (isDeletandoLixo) {
                p.parentNode.removeChild(p); 
            } else {
                p.parentNode.removeChild(p);
            }
        }

        const serializer = new XMLSerializer();
        const newXmlString = serializer.serializeToString(xmlDoc);
        loadedZip.file("word/document.xml", newXmlString);
        processedDocxBlob = await loadedZip.generateAsync({ type: "blob" });

        generateHtmlTemplate(dataSessao, avisoEspecial, processosExtraidos);
    }

    function generateHtmlTemplate(dataSessao, avisoEspecial, processos) {
        let htmlCards = processos.map(proc => `
            <div class="processo-card">
                <div class="processo-info">
                    <div class="processo-numero">${proc.numero}</div>
                    <div class="processo-relator">${proc.relator}</div>
                </div>
                <div class="processo-actions">
                    <div class="action-box">
                        <span class="action-label">DIVERGIR</span>
                        <div class="square"></div>
                    </div>
                    <div class="action-box">
                        <span class="action-label">ACOMPANHAR</span>
                        <div class="square"></div>
                    </div>
                </div>
            </div>
        `).join("");

        // A caixa só será criada se o array tiver textos capturados APÓS o gatilho
        let htmlAviso = avisoEspecial.length > 0 
            ? `<div class="aviso-box">${avisoEspecial.join("<br>")}</div>` 
            : "";

        generatedHtmlString = `
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <title>Pauta de Julgamento - Resumida</title>
                <style>
                    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background: #fff; color: #333; }
                    .container { max-width: 900px; margin: 0 auto; padding: 2cm; }
                    
                    /* Cabeçalho */
                    .header-section { text-align: center; margin-bottom: 25px; }
                    .titulo { font-size: 14pt; font-weight: bold; color: #000; text-transform: uppercase; margin-bottom: 8px; }
                    .data-sessao { font-size: 12pt; color: #333; font-weight: 600; }
                    .aviso-box { border: 2px solid #000; padding: 15px; margin: 0 auto 30px auto; max-width: 80%; text-align: center; font-weight: bold; font-size: 10pt; line-height: 1.4; border-radius: 4px; }
                    
                    /* Layout Flexbox */
                    .processo-card { border-top: 1px solid #cbd5e1; padding: 15px 0; page-break-inside: avoid; display: flex; justify-content: space-between; align-items: flex-start; }
                    .processo-info { flex: 1; padding-right: 20px; }
                    .processo-numero { font-weight: 700; font-size: 11pt; color: #000; margin-bottom: 6px; }
                    .processo-relator { font-size: 10pt; color: #475569; line-height: 1.3; }
                    
                    /* Redução de 50% nos Quadrados (de 70px para 35px) */
                    .processo-actions { display: flex; gap: 20px; align-items: flex-start; margin-top: -2px; }
                    .action-box { display: flex; flex-direction: column; align-items: center; gap: 5px; }
                    .action-label { font-size: 8pt; font-weight: 700; color: #000; letter-spacing: 0.5px; }
                    .square { width: 35px; height: 35px; border: 2px solid #000; border-radius: 4px; }
                    
                    @page { size: A4 portrait; margin: 1.5cm; }
                    @media print {
                        body { filter: grayscale(100%); }
                        .container { padding: 0; width: 100%; max-width: 100%; }
                        .processo-card { border-top: 1px solid #000; }
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header-section">
                        <div class="titulo">FEITOS DE COMPETÊNCIA RECURSAL</div>
                        <div class="data-sessao">${dataSessao || ""}</div>
                    </div>
                    ${htmlAviso}
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
