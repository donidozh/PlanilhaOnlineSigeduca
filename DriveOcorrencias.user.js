// ==UserScript==
// @name         Drive de Ocorrências SIGEDUCA
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  Consulta e inserção de ocorrências com extração automática em background (Grid Oculto).
// @match        *://sigeduca.seduc.mt.gov.br/ged/*
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // =========================================================================
    // CONFIGURAÇÃO DA PLANILHA (Substitua pela URL do seu Web App)
    // =========================================================================
    const GAS_URL = 'https://script.google.com/macros/s/AKfycbyRK90uDaiIbI4j8tBEu8CwxsYKv9GSjsjWeQxcIDQ8rmWwf6FOPQ7PAPgJGbyDxa5C/exec';

    // =========================================================================
    // 1. INTERCEPTADOR DO MENU GENEXUS
    // =========================================================================
    const observer = new MutationObserver((mutations, obs) => {
        const gxState = document.querySelector('input[name="GXState"]');

        if (gxState && !gxState.dataset.modificadoOcorrencias) {
            try {
                let state = JSON.parse(gxState.value);

                let sigEscola = state.MPW0010vMENUDATACOLLECTION?.find(m => m.Title === "SIG Escola");
                if (sigEscola && sigEscola.Childs) {
                    let relatorios = sigEscola.Childs.find(c => c.Title === "Relatórios");

                    if (relatorios && relatorios.Childs) {
                        relatorios.Childs.push({
                            "Icon": "",
                            "Title": "Drive de Ocorrências",
                            "Url": "hwcmatriculasaluno.aspx?oco=1",
                            "Target": "_self",
                            "Description": "Central de Ocorrências Comportamentais",
                            "Childs": []
                        });
                    }
                }

                gxState.value = JSON.stringify(state);
                gxState.dataset.modificadoOcorrencias = "true";
                obs.disconnect();
            } catch(e) {
                console.error("Erro ao injetar menu Drive de Ocorrências:", e);
            }
        }
    });
    observer.observe(document.documentElement, { childList: true, subtree: true });


    // =========================================================================
    // 2. MUTANTE DE PÁGINA E CSS (Interface Principal)
    // =========================================================================
    window.addEventListener('DOMContentLoaded', () => {
        if (window.location.href.toLowerCase().includes('oco=1')) {
            injetarCSS();
            adaptarPaginaOcorrencias();
        }
    });

    // Força a ocultação da tabela nativa de matrículas para não poluir a tela
    function injetarCSS() {
        const style = document.createElement('style');
        style.innerHTML = `
            #GRIDS, #FreesgridContainerDiv {
                display: none !important;
            }
        `;
        document.head.appendChild(style);
    }

    function adaptarPaginaOcorrencias() {
        function setTitles() {
            document.title = "Drive de Ocorrências";
            const tituloTabela = document.getElementById('TTITULO');
            if (tituloTabela) tituloTabela.innerText = "Drive de Ocorrências";
        }
        setTitles();
        setTimeout(setTitles, 500);

        // Oculta os campos nativos desnecessários
        const camposParaOcultar = [
            'TBESCOLA', 'TBMUNICIPIO', 'vGERANOLETCOD', 'vGEDALUIDINEP', 'TBDATANASCIMENTO', 'TBNOMEMAE'
        ];

        camposParaOcultar.forEach(id => {
            const elemento = document.getElementById(id);
            if (elemento) {
                const linhaTabela = elemento.closest('tr');
                if (linhaTabela) linhaTabela.style.display = 'none';
            }
        });

        // Cria nossa área de resultados
        const containerPrincipal = document.querySelector('.Section');
        if (containerPrincipal) {
            const divResultados = document.createElement('div');
            divResultados.id = 'containerResultadosOcorrencias';
            divResultados.style.marginTop = '20px';
            divResultados.style.width = '100%';
            containerPrincipal.appendChild(divResultados);
        }

        const btnOriginal = document.querySelector('input[name="BCONSULTAR"]');
        if (btnOriginal) {
            // Oculta o botão de consulta original da Genexus, mas mantém ele no DOM para podermos clicar via script
            btnOriginal.style.display = 'none';

            const btnConsultar = document.createElement('input');
            btnConsultar.type = 'button';
            btnConsultar.value = 'Consultar Histórico';
            btnConsultar.name = 'BCONSULTAR_OCO';
            btnConsultar.className = btnOriginal.className;

            btnConsultar.addEventListener('click', (e) => {
                e.preventDefault();
                acionarConsulta();
            });

            const btnIncluir = document.createElement('input');
            btnIncluir.type = 'button';
            btnIncluir.value = 'Incluir Ocorrência';
            btnIncluir.name = 'BINCLUIR_OCO';
            btnIncluir.className = btnOriginal.className;
            btnIncluir.style.marginLeft = "15px";

            btnIncluir.addEventListener('click', (e) => {
                e.preventDefault();
                acionarInclusao();
            });

            const parent = btnOriginal.parentNode;
            parent.insertBefore(btnConsultar, btnOriginal);
            parent.insertBefore(btnIncluir, btnOriginal);
        }
    }


    // =========================================================================
    // 3. CAPTURA DE DADOS E LÓGICA DO MODAL (Com extração em Background)
    // =========================================================================

    function obterUsuarioLogado() {
        const loginSpan = document.getElementById('MPW0010TLOGIN');
        if (loginSpan) {
            let texto = loginSpan.innerText;
            if (texto.includes(':')) {
                return texto.split(':')[1].trim();
            }
            return texto.trim();
        }
        return "USUÁRIO DESCONHECIDO";
    }

    function extrairLotacaoAtual() {
        const lotacaoEl = document.getElementById("MPW0010TLOTACAO");
        if (lotacaoEl) {
            const textoLotacao = lotacaoEl.innerText || lotacaoEl.textContent;
            const partes = textoLotacao.split("-");
            if (partes.length > 0) {
                return partes[0].trim();
            }
        }
        return "";
    }

    function extrairTurmaTurno() {
        const schoolCode = extrairLotacaoAtual();
        let turma = "NÃO IDENTIFICADA";
        let turno = "NÃO IDENTIFICADO";

        const rows = document.querySelectorAll("tr[id^='FreesgridContainerRow_']");

        for (let row of rows) {
            let match = row.id.match(/FreesgridContainerRow_(\d+)/);
            if (match) {
                let suffix = match[1];
                let schoolSpan = document.getElementById(`span_vGERLOTNOMAUX_${suffix}`);

                // Usamos .textContent ao invés de .innerText porque os elementos estão ocultos via CSS
                if (schoolSpan && schoolSpan.textContent.includes(schoolCode)) {
                    let elTurma = document.getElementById(`span_vGERTURSAL_${suffix}`);
                    let elTurno = document.getElementById(`span_vGERTRNDSC_${suffix}`);

                    if (elTurma && elTurma.textContent.trim() !== "") turma = elTurma.textContent.trim();
                    if (elTurno && elTurno.textContent.trim() !== "") turno = elTurno.textContent.trim();

                    break;
                }
            }
        }
        return { turma, turno };
    }

    function capturarDadosTela() {
        const extracao = extrairTurmaTurno();

        return {
            anoLetivo: document.getElementById('vGERANOLETCOD')?.value || new Date().getFullYear().toString(),
            codAluno: document.getElementById('vGEDALUCOD')?.value || '',
            nomeAluno: document.getElementById('span_vGEDALUNOM')?.innerText || '',
            turma: extracao.turma,
            turno: extracao.turno,
            usuario: obterUsuarioLogado()
        };
    }

    function acionarInclusao() {
        const codAluno = document.getElementById('vGEDALUCOD')?.value || '';
        if(!codAluno || codAluno === "0") {
            alert("⚠️ Selecione um aluno clicando na lupa antes de incluir uma ocorrência.");
            return;
        }

        const btnIncluir = document.querySelector('input[name="BINCLUIR_OCO"]');
        const btnOriginal = document.querySelector('input[name="BCONSULTAR"]');

        // Altera o visual do botão enquanto carrega as matrículas em background
        btnIncluir.value = "Carregando turma...";
        btnIncluir.disabled = true;

        // Dispara o clique no botão oculto nativo do SigEduca
        if (btnOriginal) btnOriginal.click();

        // Loop de espera (Polling) para aguardar o AJAX da Genexus terminar de montar a tabela oculta
        let tentativas = 0;
        const maxTentativas = 30; // Limite de 15 segundos (30 * 500ms)

        const checkInterval = setInterval(() => {
            tentativas++;
            const rows = document.querySelectorAll("tr[id^='FreesgridContainerRow_']");
            const ajaxLoader = document.getElementById('gx_ajax_notification');
            const isAjaxLoading = ajaxLoader && ajaxLoader.style.display !== 'none';

            // Se as linhas foram criadas no DOM e o carregamento parou
            if (rows.length > 0 && !isAjaxLoading) {
                clearInterval(checkInterval);

                // Restaura o botão
                btnIncluir.value = "Incluir Ocorrência";
                btnIncluir.disabled = false;

                // Extrai os dados e abre o modal
                const dados = capturarDadosTela();
                if (dados.turma === "NÃO IDENTIFICADA") {
                    const confirmacao = confirm("Não identificamos uma matrícula recente para este aluno na nossa unidade. Deseja prosseguir com a ocorrência mesmo assim?");
                    if (!confirmacao) return;
                }
                abrirModal(dados);

            } else if (tentativas >= maxTentativas) {
                clearInterval(checkInterval);
                btnIncluir.value = "Incluir Ocorrência";
                btnIncluir.disabled = false;
                alert("⏳ Tempo limite excedido ao buscar a turma. Tente novamente.");
            }
        }, 500);
    }

    function abrirModal(dados) {
        let modalExistente = document.getElementById('modalOcorrenciaBg');
        if(modalExistente) modalExistente.remove();

        const dataHoje = new Date().toISOString().split('T')[0];

        const modalHtml = `
        <div id="modalOcorrenciaBg" style="position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:9999; display:flex; justify-content:center; align-items:center;">
            <div class="Form" style="background:#FFF; padding:0; border:2px solid #065195; border-radius:5px; width:780px; box-shadow: 0px 0px 15px rgba(0,0,0,0.5);">

                <table class="Table" cellpadding="1" cellspacing="2" style="width:100%;">
                    <tbody>
                        <tr>
                            <td bgcolor="#065195" style="padding: 5px;">
                                <span class="TextBlock" style="font-family:'Verdana'; font-size:8.0pt; font-weight:bold; color:#FFFFFF">Cadastro de Ocorrências</span>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 15px;">
                                <table class="Table" cellpadding="1" cellspacing="2" style="width: 100%;">
                                    <tbody>
                                        <tr>
                                            <td style="text-align:-khtml-right; width: 150px;">
                                                <p align="right"><span class="TituloCampo">Código do Aluno:</span></p>
                                            </td>
                                            <td>
                                                <span class="ReadonlyAttribute" style="font-weight:bold;">${dados.codAluno} - ${dados.nomeAluno}</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="text-align:-khtml-right">
                                                <p align="right"><span class="TituloCampo">Turma / Turno:</span></p>
                                            </td>
                                            <td>
                                                <span class="ReadonlyAttribute" style="color:#065195;"><b>${dados.turma}</b> (${dados.turno})</span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="text-align:-khtml-right">
                                                <p align="right"><span class="TituloCampo">Data da Ocorrência:</span></p>
                                            </td>
                                            <td>
                                                <input type="date" id="ocoData" class="Attribute" value="${dataHoje}" style="width: 130px;">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" style="text-align:-khtml-center; padding-top: 15px;">
                                                <fieldset style="-moz-border-radius:3pt;" class="GroupGrande">
                                                    <legend class="GroupGrandeTitle">Relato da Ocorrência</legend>
                                                    <textarea cols="80" rows="8" id="ocoRelato" class="Attribute" placeholder="Descreva detalhadamente o fato ocorrido..."></textarea>
                                                </fieldset>
                                            </td>
                                        </tr>
                                    </tbody>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td style="text-align:center; padding: 15px; background: #f0f0f0; border-top: 1px solid #ccc;">
                                <input type="button" id="btnSalvarOcorrencia" value="Confirmar" class="btnConfirmar" style="cursor:pointer;">
                                <input type="button" id="btnFecharModal" value="Voltar" class="btnVoltar" style="cursor:pointer; margin-left: 10px;">
                            </td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        `;

        document.body.insertAdjacentHTML('beforeend', modalHtml);

        document.getElementById('btnFecharModal').addEventListener('click', () => {
            document.getElementById('modalOcorrenciaBg').remove();
        });

        document.getElementById('btnSalvarOcorrencia').addEventListener('click', () => {
            salvarOcorrencia(dados);
        });
    }


    // =========================================================================
    // 4. INTEGRAÇÃO COM GOOGLE SHEETS
    // =========================================================================

    function salvarOcorrencia(dadosAluno) {
        const dataOcorrencia = document.getElementById('ocoData').value;
        const relato = document.getElementById('ocoRelato').value;
        const btnSalvar = document.getElementById('btnSalvarOcorrencia');

        if(!dataOcorrencia || relato.trim() === "") {
            alert("⚠️ Preencha a data e o relato da ocorrência.");
            return;
        }

        const [ano, mes, dia] = dataOcorrencia.split('-');
        const dataFormatada = `${dia}/${mes}/${ano}`;

        const payload = {
            acao: "inserir",
            anoLetivo: dadosAluno.anoLetivo,
            codAluno: dadosAluno.codAluno,
            nomeAluno: dadosAluno.nomeAluno,
            turma: dadosAluno.turma,
            turno: dadosAluno.turno,
            dataOcorrencia: dataFormatada,
            relato: relato,
            autor: dadosAluno.usuario
        };

        btnSalvar.value = "Salvando...";
        btnSalvar.disabled = true;

        fetch(GAS_URL, {
            method: 'POST',
            body: JSON.stringify(payload),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        })
        .then(response => response.json())
        .then(data => {
            if(data.status === 'sucesso') {
                alert("✅ Ocorrência salva no Drive com sucesso!");
                document.getElementById('modalOcorrenciaBg').remove();
            } else {
                alert("❌ Erro ao salvar: " + data.mensagem);
                btnSalvar.value = "Confirmar";
                btnSalvar.disabled = false;
            }
        })
        .catch(err => {
            console.error(err);
            alert("❌ Erro de conexão com o Google Sheets.");
            btnSalvar.value = "Confirmar";
            btnSalvar.disabled = false;
        });
    }

    function acionarConsulta() {
        const codAluno = document.getElementById('vGEDALUCOD')?.value || '';
        const nomeAluno = document.getElementById('span_vGEDALUNOM')?.innerText || '';
        if(!codAluno || codAluno === "0") {
            alert("⚠️ Selecione um aluno clicando na lupa para consultar o histórico.");
            return;
        }

        const btn = document.querySelector('input[name="BCONSULTAR_OCO"]');
        const container = document.getElementById('containerResultadosOcorrencias');

        btn.value = "Buscando...";
        btn.disabled = true;
        container.innerHTML = `<span style="font-family:Verdana; font-size:12px; color:#065195;">Buscando histórico na base de dados...</span>`;

        const payload = {
            acao: "consultar",
            codAluno: codAluno
        };

        fetch(GAS_URL, {
            method: 'POST',
            body: JSON.stringify(payload),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        })
        .then(response => response.json())
        .then(data => {
            if(data.status === 'sucesso') {
                renderizarTabelaOcorrencias(data.ocorrencias, nomeAluno);
            } else {
                container.innerHTML = `<span style="color:red; font-family:Verdana; font-size:12px;">Erro: ${data.mensagem}</span>`;
            }
        })
        .catch(err => {
            console.error(err);
            container.innerHTML = `<span style="color:red; font-family:Verdana; font-size:12px;">Erro de conexão.</span>`;
        })
        .finally(() => {
            btn.value = "Consultar Histórico";
            btn.disabled = false;
        });
    }

    function renderizarTabelaOcorrencias(ocorrencias, nomeAluno) {
        const container = document.getElementById('containerResultadosOcorrencias');

        if (!ocorrencias || ocorrencias.length === 0) {
            container.innerHTML = `
                <fieldset style="-moz-border-radius:3pt;" class="GroupGrande">
                    <legend class="GroupGrandeTitle">Histórico de Ocorrências</legend>
                    <div style="padding: 10px; font-family: Verdana; font-size: 11px;">Nenhuma ocorrência registrada para ${nomeAluno}.</div>
                </fieldset>`;
            return;
        }

        let linhas = ocorrencias.map(oco => `
            <tr>
                <td style="border-bottom: 1px solid #ccc; padding: 5px; font-family: Verdana; font-size: 10px; width: 80px; text-align: center;"><b>${oco.data}</b></td>
                <td style="border-bottom: 1px solid #ccc; padding: 5px; font-family: Verdana; font-size: 11px; width: 100px; text-align: center;">${oco.turma}<br><span style="font-size:9px; color:#666;">${oco.turno}</span></td>
                <td style="border-bottom: 1px solid #ccc; padding: 5px; font-family: Verdana; font-size: 11px;">${oco.relato}</td>
                <td style="border-bottom: 1px solid #ccc; padding: 5px; font-family: Verdana; font-size: 10px; width: 130px; text-align: center; color: #555;"><i>${oco.autor}</i></td>
            </tr>
        `).join('');

        container.innerHTML = `
            <fieldset style="-moz-border-radius:3pt;" class="GroupGrande">
                <legend class="GroupGrandeTitle">Histórico de Ocorrências - ${nomeAluno}</legend>
                <table width="100%" cellspacing="0" cellpadding="2" style="border-collapse: collapse;">
                    <thead>
                        <tr bgcolor="#f0f0f0">
                            <th style="padding: 5px; font-family: Verdana; font-size: 11px; text-align: center;">Data do Fato</th>
                            <th style="padding: 5px; font-family: Verdana; font-size: 11px; text-align: center;">Turma/Turno</th>
                            <th style="padding: 5px; font-family: Verdana; font-size: 11px; text-align: left;">Descrição da Ocorrência</th>
                            <th style="padding: 5px; font-family: Verdana; font-size: 11px; text-align: center;">Registrado por</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${linhas}
                    </tbody>
                </table>
            </fieldset>
        `;
    }

})();
