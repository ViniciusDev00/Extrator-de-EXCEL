// op-preview.js - Carrega e exibe os dados da Ordem de Produção

document.addEventListener("DOMContentLoaded", () => {
    carregarDadosOP();
    atualizarDataGeracao();
});

function carregarDadosOP() {
    // Recuperar dados do localStorage
    const dadosGrupo = JSON.parse(localStorage.getItem("opGrupoDados"));
    const grupoSelecionado = localStorage.getItem("opGrupoSelecionado");
    const todosDados = JSON.parse(localStorage.getItem("lotesDetalhes"));

    if (!dadosGrupo || !grupoSelecionado) {
        alert("Nenhum dado encontrado. Retornando à página principal.");
        window.location.href = "index.html";
        return;
    }

    // Atualizar informações do cabeçalho
    atualizarCabecalho(grupoSelecionado, dadosGrupo);

    // Gerar tabela de produção
    gerarTabelaProducao(dadosGrupo);

    // [Removida a chamada a atualizarResumoExecutivo para corrigir o erro 'null']
}

function atualizarCabecalho(grupo, dados) {
    const nomeGrupo = obterNomeGrupo(grupo);
    const categoria = obterCategoriaGrupo(grupo);
    const semana = obterSemanaAtual();
    
    // Calcular totais básicos
    const totalItens = dados.length;
    
    // Atualizar elementos
    document.getElementById("grupoTexto").textContent = `GRUPO ${grupo} - ${nomeGrupo}`;
    document.getElementById("categoriaTexto").textContent = categoria;
    document.getElementById("semanaTexto").textContent = semana;
    document.getElementById("qtdItens").textContent = totalItens;
    
    // Gerar código da OP
    const codigoOp = `OP-${grupo}-${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, '0')}${String(new Date().getDate()).padStart(2, '0')}`;
    document.getElementById("codigoOp").textContent = codigoOp;
}

function gerarTabelaProducao(dados) {
    const tbody = document.getElementById("tableBody");
    tbody.innerHTML = "";

    let totalQtd = 0;
    let totalPvc = 0;
    let totalInox = 0;

    // Obtém o grupo da primeira peça (assumindo que todos pertencem ao mesmo grupo)
    const grupoSelecionado = dados[0].GRUPO;
    const nomeGrupo = obterNomeGrupo(grupoSelecionado);
    
    // Ordenar dados
    dados.sort((a, b) => {
        // Agora ordenamos apenas pela dimensão e material, já que o nome do grupo é o mesmo
        if (a.DIMENSÃO !== b.DIMENSÃO) return a.DIMENSÃO.localeCompare(b.DIMENSÃO);
        if (a.BOJO !== b.BOJO) return a.BOJO.localeCompare(b.BOJO);
        return 0;
    });

    dados.forEach((item, index) => {
        const dimensao = formatarDimensao(item.DIMENSÃO);
        
        // NOVO: Usa a descrição do grupo em vez da sigla da peça
        const nomeParaExibir = `${nomeGrupo.toUpperCase()} - ${dimensao}`;
        
        const quantidade = parseInt(item["QUANTIDADE TOTAL"]) || 0;
        const material = (item.BOJO || "").toUpperCase();
        const isPVC = material.includes("PVC") || material.includes("PP");
        const qtdPvc = isPVC ? quantidade : 0;
        const qtdInox = !isPVC ? quantidade : 0;

        totalQtd += quantidade;
        totalPvc += qtdPvc;
        totalInox += qtdInox;

        // Determinar status do estoque
        const estoquePPClass = qtdPvc > 0 ? "estoque-fabricar" : "estoque-ok";
        const estoquePPText = qtdPvc > 0 ? "FABRICAR" : "-";
        
        const estoqueInoxClass = qtdInox > 0 ? "estoque-fabricar" : "estoque-ok";
        const estoqueInoxText = qtdInox > 0 ? "FABRICAR" : "-";

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td class="col-item">${index + 1}</td>
            <td class="col-pedido">
                <div style="font-weight: 600; color: var(--neutral-800);">${nomeParaExibir}</div>
                <div style="font-size: 0.75rem; color: var(--neutral-500); margin-top: 1px; display: none;">
                    Material: ${item.BOJO || 'N/A'}
                </div>
            </td>
            <td class="col-qtd" style="font-weight: 600; color: var(--neutral-900);">
                ${quantidade.toLocaleString()}
            </td>
            <td class="col-pvc" style="${qtdPvc > 0 ? 'color: #0369a1; font-weight: 700;' : 'color: var(--neutral-500);'}">
                ${qtdPvc > 0 ? qtdPvc.toLocaleString() : '-'}
            </td>
            <td class="col-inox" style="${qtdInox > 0 ? 'color: var(--neutral-800); font-weight: 700;' : 'color: var(--neutral-500);'}">
                ${qtdInox > 0 ? qtdInox.toLocaleString() : '-'}
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

    // Atualizar totais da tabela
    document.getElementById("totalQtd").textContent = totalQtd.toLocaleString();
    document.getElementById("totalPvc").textContent = totalPvc.toLocaleString();
    document.getElementById("totalInox").textContent = totalInox.toLocaleString();
}

function atualizarDataGeracao() {
    const agora = new Date();
    const dataFormatada = `${String(agora.getDate()).padStart(2, '0')}/${String(agora.getMonth() + 1).padStart(2, '0')}/${agora.getFullYear()}`;
    const hora = `${String(agora.getHours()).padStart(2, '0')}:${String(agora.getMinutes()).padStart(2, '0')}`;
    
    // Esta linha preenche o campo no cabeçalho.
    document.getElementById("dataGeracao").textContent = `${dataFormatada} às ${hora}`;
}

function fecharVisualizacao() {
    if (confirm("Deseja fechar a visualização da Ordem de Produção?")) {
        window.history.back();
    }
}

// [Removida a função exportarPDF]

function produzirItem(index) {
    // Simulação de início de produção
    const botoes = document.querySelectorAll('.acao-button');
    if (botoes[index]) {
        const botaoOriginal = botoes[index].innerHTML;
        botoes[index].innerHTML = '<i class="fas fa-spinner fa-spin"></i> INICIADO';
        botoes[index].style.background = 'linear-gradient(135deg, #059669 0%, #10b981 100%)';
        botoes[index].disabled = true;
        
        // Mostrar notificação
        mostrarNotificacao(`Produção do item ${index + 1} iniciada`, 'success');
        
        // Reverter após 3 segundos (simulação)
        setTimeout(() => {
            botoes[index].innerHTML = botaoOriginal;
            botoes[index].style.background = 'linear-gradient(135deg, var(--primary-main) 0%, var(--primary-light) 100%)';
            botoes[index].disabled = false;
        }, 3000);
    }
}

function mostrarNotificacao(mensagem, tipo) {
    // Criar elemento de notificação
    const notificacao = document.createElement('div');
    notificacao.className = `notificacao ${tipo}`;
    notificacao.innerHTML = `
        <i class="fas fa-${tipo === 'success' ? 'check-circle' : 'exclamation-circle'}"></i>
        <span>${mensagem}</span>
        <button onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    // Estilos inline para a notificação
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
    
    // Remover após 5 segundos
    setTimeout(() => {
        if (notificacao.parentElement) {
            notificacao.remove();
        }
    }, 5000);
}

// Adicionar estilos CSS para animação
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    .notificacao button {
        background: transparent;
        border: none;
        color: white;
        cursor: pointer;
        padding: 4px;
        margin-left: 10px;
    }
    
    .notificacao button:hover {
        opacity: 0.8;
    }
`;
document.head.appendChild(style);

// Funções auxiliares (mantidas do original)
function obterNomeGrupo(grupo) {
    const nomes = {
        A: "VERTICAL ALTO",
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
    };
    return nomes[grupo] || `GRUPO ${grupo}`;
}

function obterCategoriaGrupo(grupo) {
    const categorias = {
        A: "VERTICAL ALTO",
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
    };
    return categorias[grupo] || "EXPOSITORES";
}

function obterSemanaAtual() {
    // NOVO: Lê a semana da planilha salva no localStorage pelo script.js
    const semanaLida = localStorage.getItem("semanaOP");
    if (semanaLida) return semanaLida;
    
    // Fallback: Calcula a semana atual (como antes)
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