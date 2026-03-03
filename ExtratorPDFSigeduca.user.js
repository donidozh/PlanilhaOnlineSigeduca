// ==UserScript==
// @name         Extrator de Codigos PDF Sigeduca
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Extrai codigos de alunos do PDF no Sigeduca
// @author       Elder Martins
// @match        *://sigeduca.seduc.mt.gov.br/ged/arralunossituacao.aspx*
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js
// @grant        GM_xmlhttpRequest
// @grant        GM_setClipboard
// ==/UserScript==

(function() {
    'use strict';

    // Configuração do PDF.js
    const pdfjsLib = window['pdfjs-dist/build/pdf'];
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    // Cria o botão na tela
    function criarBotao() {
        if (document.getElementById('btn-extrator-pdf')) return;

        const btn = document.createElement('button');
        btn.id = 'btn-extrator-pdf';
        btn.innerHTML = '🔍 Extrair Códigos do PDF';
        btn.style = "position:fixed; top:20px; right:20px; z-index:999999; padding:15px; background:#28a745; color:white; border:bold; border-radius:8px; cursor:pointer; font-weight:bold; box-shadow: 0px 4px 6px rgba(0,0,0,0.3);";
        document.body.appendChild(btn);

        btn.onclick = async function() {
            btn.innerHTML = '⏳ lendo PDF...';
            btn.style.background = '#ffc107';
            btn.disabled = true;

            try {
                const pdfUrl = window.location.href;

                GM_xmlhttpRequest({
                    method: "GET",
                    url: pdfUrl,
                    responseType: "arraybuffer",
                    onload: async function(response) {
                        const data = new Uint8Array(response.response);
                        const loadingTask = pdfjsLib.getDocument({data});
                        const pdf = await loadingTask.promise;
                        let fullText = "";

                        for (let i = 1; i <= pdf.numPages; i++) {
                            const page = await pdf.getPage(i);
                            const textContent = await page.getTextContent();
                            fullText += textContent.items.map(item => item.str).join(" ") + "\n";
                        }

                        // Regex para pegar números de 7 a 10 dígitos
                        const regex = /\b\d{7,10}\b/g;
                        const encontrados = fullText.match(regex);
                        const codigos = encontrados ? [...new Set(encontrados)] : [];

                        if (codigos.length > 0) {
                            GM_setClipboard(codigos.join('\n'));
                            alert('Sucesso! ' + codigos.length + ' códigos copiados para a área de transferência.');
                        } else {
                            alert('Não encontrei códigos com o padrão esperado no PDF.');
                        }

                        btn.innerHTML = '🔍 Extrair Códigos do PDF';
                        btn.style.background = '#28a745';
                        btn.disabled = false;
                    },
                    onerror: function() {
                        alert('Erro ao baixar o arquivo PDF.');
                        btn.disabled = false;
                    }
                });
            } catch (e) {
                console.error(e);
                alert('Erro durante o processamento.');
                btn.disabled = false;
            }
        };
    }

    // Tenta criar o botão após o carregamento
    setTimeout(criarBotao, 2000);
})();
