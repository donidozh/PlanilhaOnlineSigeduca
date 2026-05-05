// ==UserScript==
// @name         Envio para Planilha Online - Lote (Com Atestados)
// @namespace    http://tampermonkey.net/
// @version      3.1
// @description  Envio de turmas para planilha online, Geração de Impressão e Extração de Atestados em lote.
// @author       Elder Martins
// @match        *://sigeduca.seduc.mt.gov.br/ged/hwmgrhturma.aspx*
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
    'use strict';

    // Padrões de Link atualizados
    const urlWebAppPadrao = "";
    const urlPlanilhaPadrao = "";

    let urlWebapp = localStorage.getItem('sigeduca_url_webapp') || urlWebAppPadrao;
    let urlPlanilha = localStorage.getItem('sigeduca_url_planilha') || urlPlanilhaPadrao;

    const pdfjsLib = window['pdfjs-dist/build/pdf'];
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

    // Utilitários
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
    function isNotificationHidden(docObject) {
        var notification = docObject.getElementById('gx_ajax_notification');
        if (notification) return docObject.defaultView.getComputedStyle(notification).getPropertyValue('display') === 'none';
        return true;
    }

    let situacoesConhecidas = [
        "AFASTADO POR ABANDONO", "DEPENDENTE", "AFASTADO POR DESISTÊNCIA", "MATRICULADO",
        "MATRÍCULA EXTRAORDINÁRIA", "MATRÍCULA DE PROGRESSÃO PARCIAL", "RECLASSIFICADO",
        "TRANSFERIDO DA TURMA", "TRANSFERIDO DA ESCOLA", "TRANSF. ESCOLA - DEPENDENTE",
        "TRANSF. ESCOLA - MAT. PROGRE. PARC.", "TRANSF. ESCOLA - MAT. EXTRAORD.", "ÓBITO",
        "MATRICULA CANCELADA", "MATRICULA ESTORNADA", "TRANSFERÊNCIA CANCELADA",
        "TRANSFERÊNCIA ESTORNADA", "RECLASSIFICAÇÃO CANCELADA", "RECLASSIFICAÇÃO ESTORNADA",
        "MATRÍCULA PENDENTE", "SUPERADO", "SUPERAÇÃO ESTORNADA", "SUPERAÇÃO CANCELADA",
        "RESERVA DE MATRÍCULA", "AFASTADO C.H. COMPONENTE CURRICULAR"
    ];
    situacoesConhecidas.sort((a, b) => b.length - a.length);

    let turmasMapeadas = [];
    let isRodando = false;

    function criarPainelLote() {
        if (document.getElementById('painel-lote-sigeduca')) return;

        const painel = document.createElement('div');
        painel.id = 'painel-lote-sigeduca';
        painel.style = "position:fixed; top:20px; right:20px; z-index:999999; background:#f8f9fa; border:2px solid #343a40; border-radius:8px; box-shadow: 0px 4px 6px rgba(0,0,0,0.3); font-family: Arial, sans-serif; width: 340px; overflow: hidden; box-sizing: border-box;";

        const trilho = document.createElement('div');
        trilho.id = 'trilho-slider-lote';
        trilho.style = "display: flex; width: 680px; transition: transform 0.3s ease-in-out;";

        const telaPrincipal = document.createElement('div');
        telaPrincipal.style = "width: 340px; padding: 15px; box-sizing: border-box; position: relative;";
        telaPrincipal.innerHTML = `
            <div id="btn-engrenagem-lote" style="position: absolute; top: 12px; right: 15px; cursor: pointer; font-size: 18px;" title="Configurações">⚙️</div>
            <h4 style="margin: 0 0 10px 0; color: #17a2b8; font-size: 15px; text-align: center;">Ações em Lote (Turmas)</h4>
            <p style="font-size: 11px; color: #555; text-align: center; margin-bottom: 15px;">Filtre as turmas no Sigeduca e clique em Consultar. O script irá rastrear os links.</p>

            <button id="btn-mapear" style="width: 100%; padding: 8px; margin-bottom: 10px; background:#6c757d; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">1. Mapear Turmas na Tela</button>
            <div id="status-mapeamento" style="font-size: 12px; font-weight: bold; color: #333; text-align: center; margin-bottom: 10px; min-height: 15px;"></div>

            <select id="acao-lote" style="width: 100%; padding: 8px; margin-bottom: 10px; border-radius: 4px; border: 1px solid #ccc; box-sizing: border-box;">
                <option value="imprimir">🖨️ Gerar Impressão Manual</option>
                <option value="sheets">🚀 Enviar Turmas p/ Planilha</option>
                <option value="atestados">🩺 Extrair Atestados p/ Planilha</option>
            </select>

            <button id="btn-iniciar-lote" style="width: 100%; padding: 12px; background:#28a745; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold; opacity: 0.5;" disabled>2. Iniciar Processo 🚀</button>
            <button id="btn-abrir-planilha-lote" style="display: block; width: 100%; padding: 10px; margin-top: 10px; background:#007bff; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold; text-align: center; text-decoration: none; box-sizing: border-box; transition: background 0.3s;">📂 Abrir Planilha Online</button>

            <div id="log-lote" style="margin-top: 15px; font-size: 10px; color: #28a745; max-height: 150px; overflow-y: auto; border-top: 1px solid #ddd; padding-top: 5px; font-family: monospace;"></div>
        `;

        const telaConfig = document.createElement('div');
        telaConfig.style = "width: 340px; padding: 15px; box-sizing: border-box; position: relative; background: #ececec;";
        telaConfig.innerHTML = `
            <h4 style="margin: 0 0 10px 0; color: #d9534f; font-size: 13px; text-align: center; font-weight: bold;">⚠️ Configurações de Link</h4>

            <label style="font-size: 11px; font-weight: bold; color: #333;">Link da Implantação (Apps Script):</label>
            <input type="text" id="input-webapp-lote" value="${urlWebapp}" style="width: 100%; padding: 6px; margin-bottom: 10px; font-size: 11px; border: 1px solid #ccc; border-radius: 3px; box-sizing: border-box;">

            <label style="font-size: 11px; font-weight: bold; color: #333;">Link da Planilha (Para visualização):</label>
            <input type="text" id="input-planilha-lote" value="${urlPlanilha}" style="width: 100%; padding: 6px; margin-bottom: 15px; font-size: 11px; border: 1px solid #ccc; border-radius: 3px; box-sizing: border-box;">

            <div style="display: flex; gap: 10px;">
                <button id="btn-voltar-lote" style="flex: 1; padding: 8px; background:#6c757d; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">Voltar</button>
                <button id="btn-salvar-config-lote" style="flex: 1; padding: 8px; background:#d9534f; color:white; border:none; border-radius:4px; cursor:pointer; font-weight:bold;">Salvar</button>
            </div>
        `;

        trilho.appendChild(telaPrincipal);
        trilho.appendChild(telaConfig);
        painel.appendChild(trilho);
        document.body.appendChild(painel);

        document.getElementById('btn-engrenagem-lote').onclick = () => { trilho.style.transform = "translateX(-340px)"; };
        document.getElementById('btn-voltar-lote').onclick = () => { trilho.style.transform = "translateX(0)"; };

        document.getElementById('btn-abrir-planilha-lote').onclick = () => {
            if (!urlPlanilha) return alert("Nenhum link de planilha configurado.");
            window.open(urlPlanilha, '_blank');
        };

        document.getElementById('btn-salvar-config-lote').onclick = () => {
            const novoWebapp = document.getElementById('input-webapp-lote').value.trim();
            const novaPlanilha = document.getElementById('input-planilha-lote').value.trim();

            localStorage.setItem('sigeduca_url_webapp', novoWebapp);
            localStorage.setItem('sigeduca_url_planilha', novaPlanilha);
            urlWebapp = novoWebapp;
            urlPlanilha = novaPlanilha;

            alert("Configurações salvas com sucesso no seu navegador!");
            trilho.style.transform = "translateX(0)";
        };

        document.getElementById('btn-mapear').onclick = mapearTurmas;
        document.getElementById('btn-iniciar-lote').onclick = iniciarProcessoLote;
    }

    function addLog(msg, cor = "#333", id = null) {
        const logDiv = document.getElementById('log-lote');
        if (id) {
            let exist = document.getElementById(id);
            if (exist) {
                exist.innerHTML = `• ${msg}`;
                exist.style.color = cor;
                return;
            }
        }
        const newLine = document.createElement('div');
        newLine.style.color = cor;
        newLine.style.marginBottom = "5px";
        if (id) newLine.id = id;
        newLine.innerHTML = `• ${msg}`;
        logDiv.prepend(newLine);
    }

    function mapearTurmas() {
        turmasMapeadas = [];
        let trs = document.querySelectorAll('table tr');

        trs.forEach(tr => {
            let htmlLinha = tr.innerHTML;
            let match = htmlLinha.match(/(arralunossituacao\.aspx\?[^"'\s>]+)/i);

            if (match) {
                let urlLimpa = match[1].replace(/['");\\]+$/, "");
                let urlCompleta = "http://sigeduca.seduc.mt.gov.br/ged/" + urlLimpa;
                let spanNome = tr.querySelector('[id^="span_vGERTURSAL_"]');
                let nome = spanNome ? spanNome.innerText.trim() : "TURMA DESCONHECIDA";

                if (nome !== "" && nome.toUpperCase() !== "NOME DA TURMA") {
                    let turno = "DESCONHECIDO";
                    let textoTr = tr.innerText.toUpperCase();
                    if (textoTr.includes('MATUTINO')) turno = 'MATUTINO';
                    else if (textoTr.includes('VESPERTINO')) turno = 'VESPERTINO';
                    else if (textoTr.includes('NOTURNO')) turno = 'NOTURNO';
                    else if (textoTr.includes('INTEGRAL')) turno = 'INTEGRAL';

                    if (!turmasMapeadas.some(t => t.url === urlCompleta)) {
                        turmasMapeadas.push({ nome, turno, url: urlCompleta });
                    }
                }
            }
        });

        const status = document.getElementById('status-mapeamento');
        const btnIniciar = document.getElementById('btn-iniciar-lote');

        if (turmasMapeadas.length > 0) {
            status.innerHTML = `✅ ${turmasMapeadas.length} links capturados!`;
            status.style.color = "#28a745";
            btnIniciar.disabled = false;
            btnIniciar.style.opacity = "1";
        } else {
            status.innerHTML = `❌ Nenhuma URL de relatório encontrada na tela.`;
            status.style.color = "#d9534f";
        }
    }

    async function iniciarProcessoLote() {
        const acao = document.getElementById('acao-lote').value;

        if ((acao === 'sheets' || acao === 'atestados') && !urlWebapp) {
            return alert("Erro: Link do Apps Script não configurado! Clique na ⚙️ para adicionar.");
        }

        if (isRodando) return;
        isRodando = true;

        const btnIniciar = document.getElementById('btn-iniciar-lote');
        btnIniciar.innerHTML = "⏳ Processando...";
        btnIniciar.style.background = "#ffc107";
        btnIniciar.style.color = "#333";
        btnIniciar.disabled = true;
        document.getElementById('log-lote').innerHTML = "";

        let janelaImpressao = null;
        let htmlGeralImpressao = "";
        if (acao === 'imprimir') {
            janelaImpressao = window.open('', '_blank');
            if (!janelaImpressao) {
                alert("O navegador bloqueou a nova aba. Por favor, permita pop-ups para este site.");
                isRodando = false;
                resetarBotaoLote();
                return;
            }
            janelaImpressao.document.write("<h2>⏳ Processando turmas em lote... Por favor, aguarde.</h2>");
        }

        for (let i = 0; i < turmasMapeadas.length; i++) {
            let turma = turmasMapeadas[i];
            let urlPdf = turma.url;

            addLog(`[${i+1}/${turmasMapeadas.length}] Lendo: <a href="${urlPdf}" target="_blank" style="color:#007bff; text-decoration:underline;">${turma.nome}</a>`, "#333");

            try {
                let textoPdf = await baixarExtrairTextoPDF(urlPdf);
                let alunos = processarTextoDoPdf(textoPdf);

                if (alunos.length === 0) {
                    addLog(`⚠️ Lentidão detectada. Tentando novamente em 3s...`, "#ff9800");
                    await new Promise(r => setTimeout(r, 3000));
                    textoPdf = await baixarExtrairTextoPDF(urlPdf);
                    alunos = processarTextoDoPdf(textoPdf);
                }

                if (alunos.length > 0) {
                    if (acao === 'sheets') {
                        addLog(`>> ${alunos.length} alunos extraídos. Enviando relação...`, "#17a2b8");
                        await enviarParaSheets(turma.nome, turma.turno, alunos);
                        addLog(`✅ Relação salva na planilha!`, "#28a745");
                    }
                    else if (acao === 'imprimir') {
                        addLog(`>> ${alunos.length} alunos extraídos. Gerando página...`, "#17a2b8");
                        htmlGeralImpressao += gerarHtmlTurma(alunos, turma.nome, turma.turno);
                        addLog(`✅ Página gerada!`, "#28a745");
                    }
                    else if (acao === 'atestados') {
                        let logId = "log_atd_" + i;
                        addLog(`>> Preparando módulo de atestados...`, "#17a2b8", logId);
                        let atestados = await extrairAtestados(alunos, turma.nome, turma.turno, logId);

                        if (atestados.length > 0) {
                            addLog(`>> ${atestados.length} atestados coletados. Enviando...`, "#17a2b8", logId);
                            await enviarAtestadosParaSheets(atestados);
                            addLog(`✅ ${atestados.length} atestados salvos na planilha!`, "#28a745", logId);
                        } else {
                            addLog(`✅ Nenhum atestado encontrado nesta turma.`, "#28a745", logId);
                        }
                    }
                } else {
                    addLog(`❌ Arquivo vazio ou sem alunos processáveis.`, "#d9534f");
                }
            } catch (erro) {
                console.error(erro);
                addLog(`❌ Falha na leitura (Erro interno ou resposta HTML bloqueada).`, "#d9534f");
            }

            if (i < turmasMapeadas.length - 1) {
                await new Promise(r => setTimeout(r, 4000));
            }
        }

        if (acao === 'imprimir' && janelaImpressao) {
            finalizarJanelaImpressao(janelaImpressao, htmlGeralImpressao);
        }

        btnIniciar.innerHTML = "🎉 Lote Concluído!";
        btnIniciar.style.background = "#28a745";
        btnIniciar.style.color = "white";
        setTimeout(() => resetarBotaoLote(), 3000);
        isRodando = false;
    }

    function baixarExtrairTextoPDF(url) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: "GET", url: url, responseType: "arraybuffer",
                onload: async function(response) {
                    try {
                        const data = new Uint8Array(response.response);
                        const isHTML = String.fromCharCode.apply(null, data.subarray(0, 10)).includes('<!DOCTYPE') || String.fromCharCode.apply(null, data.subarray(0, 5)).includes('<html');
                        if (isHTML) return reject("O Genexus retornou HTML ao invés de PDF.");

                        const loadingTask = pdfjsLib.getDocument({data});
                        const pdf = await loadingTask.promise;
                        let fullText = "";
                        for (let i = 1; i <= pdf.numPages; i++) {
                            const page = await pdf.getPage(i);
                            const textContent = await page.getTextContent();
                            fullText += textContent.items.map(item => item.str).join(" ") + "\n";
                        }
                        resolve(fullText);
                    } catch (e) { reject(e); }
                },
                onerror: function(err) { reject(err); }
            });
        });
    }

    function processarTextoDoPdf(fullText) {
        const regexLinhaAluno = /\b(?:\d{1,3})\s+([A-ZÀ-Ÿ0-9\s\.\-]+?)\s+(\d{7,10})\s+((?:\d{2}\/\d{2}\/\d{2,4})|(?:\/\s*\/))\s+(\d{2}\/\d{2}\/\d{4})\s+(SIM|NÃO)\s+(SIM|NÃO)\s+(SIM|NÃO)/g;
        let match;
        const alunosExtraidos = [];

        while ((match = regexLinhaAluno.exec(fullText)) !== null) {
            let textoBloco = match[1].trim();
            let situacao2026 = "NÃO IDENTIFICADA";
            for (let sit of situacoesConhecidas) {
                if (textoBloco.startsWith(sit)) { situacao2026 = sit; break; }
            }
            let nome = textoBloco;
            let limpou = true;
            while(limpou) {
                limpou = false;
                for (let sit of situacoesConhecidas) {
                    if (nome.startsWith(sit)) {
                        nome = nome.substring(sit.length).trim();
                        limpou = true;
                    }
                }
            }
            nome = nome.replace(/^-+\s*/, '').trim();

            alunosExtraidos.push({
                codigo: match[2], nome: nome, situacao: situacao2026,
                dataMatricula: match[4],
                dataAjuste: (match[3].replace(/\s/g, '') === "//" || match[3].trim() === "") ? "" : match[3],
                alunoPaed: match[5], matPaed: match[6], transporte: match[7]
            });
        }
        alunosExtraidos.sort((a, b) => a.nome.localeCompare(b.nome));
        return alunosExtraidos;
    }

    function enviarParaSheets(turma, turno, alunos) {
        return new Promise((resolve, reject) => {
            const payload = { tipoIntegracao: "RELACAO", turma, turno, alunos };
            GM_xmlhttpRequest({
                method: "POST", url: urlWebapp, data: JSON.stringify(payload),
                headers: { "Content-Type": "application/json" },
                onload: function(response) { resolve(response); },
                onerror: function(err) { reject(err); }
            });
        });
    }

    // --- MÓDULO DE ATESTADOS LOTE CORRIGIDO ---
    function prepararIframeAtestados() {
        let container = document.getElementById('containerIframeAtestadosLote');
        // A principal correção: Destruir o Iframe antigo se existir,
        // garantindo que não aproveitaremos o cache do DOM da turma anterior.
        if (container) {
            container.remove();
        }

        container = document.createElement('div');
        container.id = 'containerIframeAtestadosLote';
        container.style.cssText = 'position: absolute; width: 0; height: 0; overflow: hidden; visibility: hidden; opacity: 0; border: none;';
        let iframe = document.createElement('iframe');
        iframe.id = 'iframeAtestadosLote'; iframe.name = 'iframeAtestadosLote';
        container.appendChild(iframe);
        document.body.appendChild(container);

        return document.getElementById('iframeAtestadosLote');
    }

    async function extrairAtestados(alunos, turma, turno, logId) {
        const iframe = prepararIframeAtestados();
        iframe.src = 'http://sigeduca.seduc.mt.gov.br/ged/hwmgedatestado.aspx';

        let iframeDoc = iframe.contentWindow.document;
        let tentativasLoad = 0;
        while (!iframeDoc.getElementById('vGEDALUCOD') && tentativasLoad < 30) {
            await delay(1000);
            iframeDoc = iframe.contentWindow.document;
            tentativasLoad++;
        }

        if (tentativasLoad >= 30) throw new Error("Falha ao carregar a tela de atestados do Sigeduca.");
        await delay(1500);

        let atestadosColetados = [];

        for (let i = 0; i < alunos.length; i++) {
            if (!isRodando) break;
            let aluno = alunos[i];

            addLog(`>> Lendo atestados: ${i + 1}/${alunos.length} (${aluno.nome.substring(0,18)}...)`, "#17a2b8", logId);

            let inputAluno = iframeDoc.getElementById('vGEDALUCOD');
            if (inputAluno) {
                inputAluno.value = aluno.codigo;
                if ("createEvent" in iframeDoc) {
                    var evt = iframeDoc.createEvent("HTMLEvents");
                    evt.initEvent("change", false, true);
                    inputAluno.dispatchEvent(evt);
                } else { inputAluno.fireEvent("onchange"); }
            } else { continue; }

            await delay(300);
            let btnConsultar = iframeDoc.getElementsByName('BCONSULTAR')[0] || iframeDoc.querySelector('.btnConsultar');
            if (btnConsultar) btnConsultar.click();
            await delay(300);

            while (!isNotificationHidden(iframeDoc)) { await delay(300); }

            let docTabela = iframeDoc;
            if (iframe.contentWindow.frames.length > 0 && iframe.contentWindow.frames[0].document.getElementById('GriddetalhesContainerTbl')) {
                docTabela = iframe.contentWindow.frames[0].document;
            }

            let selectPag = docTabela.getElementById('vPAG');
            let totalPaginas = selectPag ? selectPag.options.length : 1;

            for (let p = 1; p <= totalPaginas; p++) {
                if (p > 1 && selectPag) {
                    selectPag.value = p.toString();
                    if ("createEvent" in docTabela) {
                        var evtPag = docTabela.createEvent("HTMLEvents");
                        evtPag.initEvent("change", false, true);
                        selectPag.dispatchEvent(evtPag);
                    } else { selectPag.fireEvent("onchange"); }
                    try { docTabela.defaultView.gx.evt.execEvt('EVPAG.CLICK.', selectPag); } catch(e){}
                    await delay(300);
                    while (!isNotificationHidden(docTabela)) { await delay(300); }
                }

                let tabelaDetalhes = docTabela.getElementById('GriddetalhesContainerTbl');
                if (tabelaDetalhes && tabelaDetalhes.rows.length > 1) {
                    for (let n = 1; n < tabelaDetalhes.rows.length; n++) {
                        let numStr = ("0000" + n).slice(-4);
                        try {
                            let dataIni = docTabela.getElementById('span_vGEDATEPERINI_' + numStr)?.textContent.trim() || '';
                            let dataFim = docTabela.getElementById('span_vGEDATEPERFIN_' + numStr)?.textContent.trim() || '';
                            let tipoJust = docTabela.getElementById('span_vGEDATETIPO_' + numStr)?.textContent.trim() || '';

                            if (dataIni) {
                                let periodoStr = `de ${dataIni} a ${dataFim || dataIni}`;
                                atestadosColetados.push({
                                    codigo: aluno.codigo,
                                    nome: aluno.nome,
                                    turma: turma,
                                    turno: turno,
                                    periodo: periodoStr,
                                    tipoJustificativa: tipoJust
                                });
                            }
                        } catch (err) {}
                    }
                }
            }

            // Pausa de estabilidade final adicionada igual ao script individual
            await delay(200);
        }
        return atestadosColetados;
    }

    function enviarAtestadosParaSheets(atestados) {
        return new Promise((resolve, reject) => {
            const payload = {
                tipoIntegracao: "ATESTADOS",
                abaDestino: "ATESTADOS",
                dados: atestados
            };
            GM_xmlhttpRequest({
                method: "POST", url: urlWebapp, data: JSON.stringify(payload),
                headers: { "Content-Type": "application/json" },
                onload: function(response) { resolve(response); },
                onerror: function(err) { reject(err); }
            });
        });
    }

    // --- FUNÇÕES DE IMPRESSÃO ---
    function gerarHtmlTurma(alunos, turma, turno) {
        let linhasTabela = "";
        const alunosFiltrados = alunos.filter(a => !a.situacao.toUpperCase().includes("DEPENDENTE"));

        alunosFiltrados.forEach((aluno, index) => {
            let situacaoExibicao = aluno.situacao.toUpperCase();
            if (situacaoExibicao === "MATRICULADO") situacaoExibicao = "&nbsp;";
            else if (situacaoExibicao === "TRANSFERIDO DA ESCOLA") situacaoExibicao = "TRANSF.";
            else if (situacaoExibicao === "TRANSFERIDO DA TURMA") situacaoExibicao = "REMOV.";

            linhasTabela += `
                <tr>
                    <td class="centro">${index + 1}</td>
                    <td class="centro">${aluno.codigo}</td>
                    <td>${aluno.nome}</td>
                    <td class="centro">${situacaoExibicao}</td>
                    <td class="centro">${aluno.dataAjuste}</td>
                    <td class="centro">${aluno.dataMatricula}</td>
                </tr>
            `;
        });

        const linhasPorPagina = 55;
        let linhasRestantes = linhasPorPagina - (alunosFiltrados.length % linhasPorPagina);
        if (linhasRestantes < 5) linhasRestantes += linhasPorPagina;

        for (let i = 1; i <= linhasRestantes; i++) {
            linhasTabela += `
                <tr>
                    <td class="centro">${alunosFiltrados.length + i}</td>
                    <td class="centro">&nbsp;</td>
                    <td>&nbsp;</td>
                    <td class="centro">&nbsp;</td>
                    <td class="centro">&nbsp;</td>
                    <td class="centro">&nbsp;</td>
                </tr>
            `;
        }

        return `
            <div class="bloco-turma">
                <div class="cabecalho-impr">
                    <div class="titulo">RELAÇÃO DE ALUNOS POR SITUAÇÃO</div>
                    <div class="subtitulo">Turma: <b>${turma}</b> &nbsp;|&nbsp; Turno: <b>${turno}</b></div>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th style="width: 4%;">Nº</th>
                            <th style="width: 14%;">CÓDIGO</th>
                            <th style="width: 38%;">NOME DO ALUNO</th>
                            <th style="width: 20%;">SITUAÇÃO</th>
                            <th style="width: 12%;">DT AJUSTE</th>
                            <th style="width: 12%;">DT MATRÍCULA</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${linhasTabela}
                    </tbody>
                </table>
            </div>
        `;
    }

    function finalizarJanelaImpressao(novaJanela, htmlCorpo) {
        const htmlFinal = `
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <title>Lote de Impressão</title>
                <style>
                    body { font-family: Arial, sans-serif; padding: 15px; color: #333; margin: 0; }
                    .bloco-turma { page-break-after: always; padding-bottom: 20px; }
                    .bloco-turma:last-child { page-break-after: auto; }
                    .cabecalho-impr { text-align: center; margin-bottom: 5px; }
                    .cabecalho-impr .titulo { font-size: 10pt; text-transform: uppercase; font-weight: bold; color: #333; }
                    .cabecalho-impr .subtitulo { font-size: 10pt; text-transform: uppercase; color: #333; }
                    .cabecalho-impr .subtitulo b { color: #000; font-weight: bold; }
                    table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 7pt; border: 1px solid #000; }
                    th { border: none; border-top: 1px solid #000; border-bottom: 1px solid #000; padding: 3px 2px; text-transform: uppercase; background-color: #f2f2f2; font-weight: bold; text-align: center; }
                    th:first-child { border-left: 1px solid #000; }
                    th:last-child { border-right: 1px solid #000; }
                    td { border: none; border-top: 1px solid #000; border-bottom: 1px solid #000; padding: 3px 2px; text-transform: uppercase; }
                    tbody td:nth-child(1), tbody td:nth-child(2) { border-left: 1px solid #000; border-right: 1px solid #000; }
                    td.centro { text-align: center; }
                    @media print {
                        .no-print { display: none !important; }
                        body { padding: 0; }
                    }
                    .btn-imprimir {
                        display: block; width: 250px; margin: 15px auto; padding: 12px;
                        background: #007bff; color: white; text-align: center;
                        font-weight: bold; border-radius: 5px; cursor: pointer; border: none; font-size: 14px;
                    }
                    .btn-imprimir:hover { background: #0056b3; }
                </style>
            </head>
            <body>
                <button class="btn-imprimir no-print" onclick="window.print()">🖨️ Imprimir Todo o Lote</button>
                ${htmlCorpo}
            </body>
            </html>
        `;

        novaJanela.document.open();
        novaJanela.document.write(htmlFinal);
        novaJanela.document.close();
    }

    function resetarBotaoLote() {
        const btnIniciar = document.getElementById('btn-iniciar-lote');
        btnIniciar.innerHTML = "2. Iniciar Processo 🚀";
        btnIniciar.style.background = "#28a745";
        btnIniciar.disabled = false;
    }

    setTimeout(criarPainelLote, 1500);
})();
