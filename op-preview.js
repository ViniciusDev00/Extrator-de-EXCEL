// op-preview.js - Carrega e exibe os dados da Ordem de Produção

document.addEventListener("DOMContentLoaded", () => {
    carregarDadosOP();
    atualizarDataGeracao();
});

function carregarDadosOP() {
    // Recuperar dados do localStorage
    const dadosGrupo = JSON.parse(localStorage.getItem("opGrupoDados"));
    const grupoSelecionado = localStorage.getItem("opGrupoSelecionado");

    if (!dadosGrupo || !grupoSelecionado) {
        alert("Nenhum dado encontrado. Retornando à página principal.");
        window.location.href = "index.html";
        return;
    }

    // Atualizar informações do cabeçalho
    atualizarCabecalho(grupoSelecionado, dadosGrupo);

    // Gerar tabela de produção (com agrupamento)
    gerarTabelaProducao(dadosGrupo);
}

function atualizarCabecalho(grupo, dados) {
    const nomeGrupo = obterNomeGrupo(grupo);
    const categoria = obterCategoriaGrupo(grupo);
    const semana = obterSemanaAtual();
    
    document.getElementById("grupoTexto").textContent = `GRUPO ${grupo} - ${nomeGrupo}`;
    document.getElementById("categoriaTexto").textContent = categoria;
    document.getElementById("semanaTexto").textContent = semana;
    
    // Gerar código da OP
    const codigoOp = `OP-${grupo}-${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, '0')}${String(new Date().getDate()).padStart(2, '0')}`;
    document.getElementById("codigoOp").textContent = codigoOp;
}

function gerarTabelaProducao(dados) {
    const tbody = document.getElementById("tableBody");
    tbody.innerHTML = "";

    // =================================================================================
    // LÓGICA DE AGRUPAMENTO: Juntar itens de mesma dimensão (Independente do Material)
    // =================================================================================
    
    const itensAgrupados = {};

    dados.forEach((item) => {
        // Normaliza a chave pela dimensão. Se não tiver dimensão, usa "N/A"
        // Remove espaços extras para garantir que "1000" seja igual a "1000 "
        const chaveDimensao = item.DIMENSÃO ? String(item.DIMENSÃO).trim() : "N/A";

        if (!itensAgrupados[chaveDimensao]) {
            itensAgrupados[chaveDimensao] = {
                dimensaoRaw: chaveDimensao,
                grupo: item.GRUPO, // Guarda o grupo para pegar o nome depois
                qtdTotal: 0,
                qtdPvc: 0,
                qtdInox: 0
            };
        }

        const quantidade = parseInt(item["QUANTIDADE TOTAL"]) || 0;
        const material = (item.BOJO || "").toUpperCase();
        
        // Lógica para detectar se é PVC/PP ou Inox
        const isPVC = material.includes("PVC") || material.includes("PP");

        // Soma ao total geral daquela dimensão
        itensAgrupados[chaveDimensao].qtdTotal += quantidade;

        // Distribui nas colunas específicas
        if (isPVC) {
            itensAgrupados[chaveDimensao].qtdPvc += quantidade;
        } else {
            itensAgrupados[chaveDimensao].qtdInox += quantidade;
        }
    });

    // Converter o objeto agrupado de volta para um array para podermos ordenar e exibir
    const listaFinal = Object.values(itensAgrupados);

    // Ordenar por dimensão numéricamente
    listaFinal.sort((a, b) => {
        const dimA = parseFloat(a.dimensaoRaw.replace(',', '.'));
        const dimB = parseFloat(b.dimensaoRaw.replace(',', '.'));
        
        if (!isNaN(dimA) && !isNaN(dimB)) {
            return dimA - dimB;
        }
        return a.dimensaoRaw.localeCompare(b.dimensaoRaw);
    });

    // Atualizar contador de itens no cabeçalho (agora reflete linhas únicas/agrupadas)
    document.getElementById("qtdItens").textContent = listaFinal.length;

    // =================================================================================
    // GERAÇÃO DO HTML
    // =================================================================================

    let totalQtdGeral = 0;
    let totalPvcGeral = 0;
    let totalInoxGeral = 0;

    // Obtém o nome do grupo para exibição
    const nomeGrupo = obterNomeGrupo(listaFinal.length > 0 ? listaFinal[0].grupo : "");

    listaFinal.forEach((item, index) => {
        const dimensaoFormatada = formatarDimensao(item.dimensaoRaw);
        
        // Nome exibido na coluna Pedido / Cliente (Agora é genérico do grupo + medida)
        const nomeParaExibir = `${nomeGrupo.toUpperCase()} - ${dimensaoFormatada}`;
        
        // Somar aos totais do rodapé
        totalQtdGeral += item.qtdTotal;
        totalPvcGeral += item.qtdPvc;
        totalInoxGeral += item.qtdInox;

        // Determinar status do estoque e badges
        const estoquePPClass = item.qtdPvc > 0 ? "estoque-fabricar" : "estoque-ok";
        const estoquePPText = item.qtdPvc > 0 ? "FABRICAR" : "-";
        
        const estoqueInoxClass = item.qtdInox > 0 ? "estoque-fabricar" : "estoque-ok";
        const estoqueInoxText = item.qtdInox > 0 ? "FABRICAR" : "-";

        // Estilização condicional para números
        const stylePvc = item.qtdPvc > 0 ? 'color: #0369a1; font-weight: 700;' : 'color: var(--neutral-300);';
        const styleInox = item.qtdInox > 0 ? 'color: var(--neutral-800); font-weight: 700;' : 'color: var(--neutral-300);';

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td class="col-item">${index + 1}</td>
            <td class="col-pedido">
                <div style="font-weight: 600; color: var(--neutral-800);">${nomeParaExibir}</div>
            </td>
            <td class="col-qtd" style="font-weight: 600; color: var(--neutral-900);">
                ${item.qtdTotal.toLocaleString()}
            </td>
            <td class="col-pvc" style="${stylePvc}">
                ${item.qtdPvc > 0 ? item.qtdPvc.toLocaleString() : '-'}
            </td>
            <td class="col-inox" style="${styleInox}">
                ${item.qtdInox > 0 ? item.qtdInox.toLocaleString() : '-'}
            </td>
            <td class="col-estoque">
                <span class="estoque-badge ${estoquePPClass}">${estoquePPText}</span>
            </td>
            <td class="col-estoque">
                <span class="estoque-badge ${estoqueInoxClass}">${estoqueInoxText}</span>
            </td>
            <td class="col-acao">
                <button class="acao-button" onclick="produzirItem(${index})">
                    <i class="fas fa-play"></i> PRODUZIR
                </button>
            </td>
        `;

        tbody.appendChild(tr);
    });

    // Atualizar totais da tabela no rodapé
    document.getElementById("totalQtd").textContent = totalQtdGeral.toLocaleString();
    document.getElementById("totalPvc").textContent = totalPvcGeral.toLocaleString();
    document.getElementById("totalInox").textContent = totalInoxGeral.toLocaleString();
}

function atualizarDataGeracao() {
    const agora = new Date();
    const dataFormatada = `${String(agora.getDate()).padStart(2, '0')}/${String(agora.getMonth() + 1).padStart(2, '0')}/${agora.getFullYear()}`;
    const hora = `${String(agora.getHours()).padStart(2, '0')}:${String(agora.getMinutes()).padStart(2, '0')}`;
    
    document.getElementById("dataGeracao").textContent = `${dataFormatada} às ${hora}`;
}

function fecharVisualizacao() {
    if (confirm("Deseja fechar a visualização da Ordem de Produção?")) {
        window.history.back();
    }
}

function produzirItem(index) {
    const botoes = document.querySelectorAll('.acao-button');
    if (botoes[index]) {
        const botaoOriginal = botoes[index].innerHTML;
        botoes[index].innerHTML = '<i class="fas fa-spinner fa-spin"></i> INICIADO';
        botoes[index].style.background = 'linear-gradient(135deg, #059669 0%, #10b981 100%)';
        botoes[index].disabled = true;
        
        mostrarNotificacao(`Produção do item ${index + 1} iniciada`, 'success');
        
        setTimeout(() => {
            botoes[index].innerHTML = botaoOriginal;
            botoes[index].style.background = 'linear-gradient(135deg, var(--primary-main) 0%, var(--primary-light) 100%)';
            botoes[index].disabled = false;
        }, 3000);
    }
}

function mostrarNotificacao(mensagem, tipo) {
    const notificacao = document.createElement('div');
    notificacao.className = `notificacao ${tipo}`;
    notificacao.innerHTML = `
        <i class="fas fa-${tipo === 'success' ? 'check-circle' : 'exclamation-circle'}"></i>
        <span>${mensagem}</span>
        <button onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    notificacao.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: ${tipo === 'success' ? '#10b981' : '#ef4444'};
        color: white;
        padding: 12px 16px;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        display: flex;
        align-items: center;
        gap: 10px;
        z-index: 1000;
        animation: slideIn 0.3s ease;
    `;
    
    document.body.appendChild(notificacao);
    
    setTimeout(() => {
        if (notificacao.parentElement) {
            notificacao.remove();
        }
    }, 5000);
}

// Adicionar estilos CSS para animação da notificação
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    .notificacao button { background: transparent; border: none; color: white; cursor: pointer; padding: 4px; margin-left: 10px; }
    .notificacao button:hover { opacity: 0.8; }
`;
document.head.appendChild(style);

// ====================================================================================
// FUNÇÕES AUXILIARES
// ====================================================================================

function obterNomeGrupo(grupo) {
    const nomes = {
        A: "VERTICAL ALTO 850",
        B: "AÇOUGUE CURVO",
        C: "ESPECIAIS",
        D: "ESPECIAIS 90",
        E: "ESPECIAIS COM PORTA",
        F: "PADARIA ESS",
        G: "AÇOUGUE RETO",
        H: "ESPECIAS 180",
        I: "ESPECIAIS 3P",
        J: "ILHAS CONGELADA",
        K: "PADARIA CURVA",
        L: "ACOPLADO",
        M: "AÇOUGUE RETO 1040",
        N: "VERTICAL ALTO 750",
        O: "ESPECIAL CENTRAL 1760"
    };
    return nomes[grupo] || `GRUPO ${grupo}`;
}

function obterCategoriaGrupo(grupo) {
    const categorias = {
        A: "VERTICAL ALTO 850",
        B: "AÇOUGUE CURVO",
        C: "ESPECIAIS",
        D: "ESPECIAIS 90",
        E: "ESPECIAIS COM PORTA",
        F: "PADARIA ESS",
        G: "AÇOUGUE RETO",
        H: "ESPECIAS 180",
        I: "ESPECIAIS 3P",
        J: "ILHAS CONGELADA",
        K: "PADARIA CURVA",
        L: "ACOPLADO",
        M: "AÇOUGUE RETO 1040",
        N: "VERTICAL ALTO 750",
        O: "ESPECIAL CENTRAL 1760"
    };
    return categorias[grupo] || "EXPOSITORES";
}

function obterSemanaAtual() {
    const semanaLida = localStorage.getItem("semanaOP");
    if (semanaLida) return semanaLida;
    
    const hoje = new Date();
    const inicioAno = new Date(hoje.getFullYear(), 0, 1);
    const dias = Math.floor((hoje - inicioAno) / (24 * 60 * 60 * 1000));
    const semana = Math.ceil((dias + 1) / 7);
    return `SEMANA ${String(semana).padStart(2, "0")}`;
}

function formatarDimensao(dimensao) {
    if (!dimensao || dimensao === "N/A") return "N/A";
    let dim = String(dimensao).replace(/\s/g, "").replace(".", ",");
    if (!dim.includes(",")) {
        dim = dim + ",00";
    } else {
        const partes = dim.split(",");
        if (partes[1].length === 1) {
            dim = dim + "0";
        } else if (partes[1].length > 2) {
            dim = partes[0] + "," + partes[1].substring(0, 2);
        }
    }
    return dim;
}