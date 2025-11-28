// ====================================================================================
// PARTE CRÍTICA: DEFINIÇÃO DE REGRAS DE AGRUPAMENTO
// ====================================================================================
const MAPA_MODELO_GRUPO = {
    // === BASEADO NOS 10 GRUPOS FORNECIDOS ANTERIORMENTE ===

    // Grupo A (Valp-850 e Variações)
    "VABP-PVT/850": "A", "VABP/850 MA": "A", "VACP-PVT/850": "A", "VACP-PVT/850 MA": "A",
    "VACP/850": "A", "VACP/750": "A", "VACP/850 MA": "A", "VAHP-PVT/850 MA": "A", "VAHP-PVT/850": "A", "VAHP/850": "A",
    "VALP- PVT/850": "A", "VALP-PVT/850": "A", "VALP-PVT/850 MA": "A", "VALP/850": "A", "VALP/750": "A", "VALP/750 MA": "A", "VABP/850": "A", "VABP/750": "A",
    "VALP/850 MA": "A", 

    // Grupo B (VAH/850)
    "VAH/850": "B", "VAL/850": "B",

    // Grupo C (VAP/850 e Variações)
    "VAP/850": "C", "VAP/850 MA": "C",

    // Grupo D (VCA2P/1040)
    "VCA2P/1040": "D",

    // Grupo E (VCAG Complexo)
    "VCAG (1,25) + VCA2P (2,50)/1040": "E", "VCAG/1040": "E", "VCAGR (1,25) + VCAG1P (1,25)/1040": "E",

    // Grupo F (VIL-2P/900)
    "VIL-2P/900": "F", "VIL-2P CENTRAL/1760": "F", "VIL-2P FRONTAL/900": "F",

    // Grupo G (VIL-2P/900 CANTO": "G",
    "VIL-2P/900 CANTO": "G",

    // Grupo H (VILP-2P/900 e Variações)
    "VILP-2P/900": "H", "VILP-2P/900 MA": "H",

    // Grupo I (VR-900 e Variações)
    "VR1P/900": "I", "VR2P/900": "I", "VR2P/900 MA": "I", "VR2PA/900": "I", "VRQU/900": "I", "VRQR/900": "I", "VRHB/900": "I",


    // Grupo J ()
    "VRA1P/1040": "J", "VRA2P/1040": "J", "VRAG (1,25) + VRA2P (2,50)/1040": "J",
    "VRAG/1040": "J", "VRAGR/1040": "J", "VRAG(3,75) + VRAG1P (1,25)/1040": "J", "VRAG2N/900": "J", "VRAG (2,50) +VRA1P (1,25)/1040": "J",

    // Grupo K ()
    "VIL-2P PONTA CT 180/900": "K",

    // Grupo L ()
    "VIL-3P/900": "L",

    // Grupo M ()
    "ICDT": "M", "ICFT": "M",

    // Grupo N ()
    "VC2P/900": "N", "VCQU/900": "N",
    
    // Grupo O ()
    "IRAS": "O",
    
};

// === MAPA DE CORES PARA AGRUPAMENTO (Para as linhas gerais) ===
const GRUPO_COLORS = {
    "A": "E0F7FA", // Azul Claro Suave
    "B": "F1F8E9", // Verde Menta Suave
    "C": "FFF3E0", // Âmbar Claro Suave
    "D": "FBE4E4", // Rosa Claro Suave
    "E": "E8EAF6", // Índigo Claro Suave
    "F": "F3E5F5", // Roxo Claro Suave
    "G": "E1F5FE", // Azul Bebê
    "H": "FFE0B2", // Laranja Claro
    "I": "DCF8C6", // Verde Limão
    "J": "FCE4EC", // Rosa Pastel
    "K": "FFFDE7", // Amarelo Gema Muito Claro
    "L": "F0F4C3", // Verde Oliva Muito Claro
    "M": "E1BEE7", // Malva/Lilás Pálido
    "N": "B2EBF2", // Ciano/Turquesa Claro
    "O": "FFCCBC", // Pêssego/Coral Claro
    "N/A": "DDDDDD" 
};

// === NOVAS CORES PARA O STATUS (Célula Específica) ===
const STATUS_COLORS = {
    "A PRODUZIR": "FFCCCC",   // Vermelho Suave
    "EM MONTAGEM": "FFF59D",  // Amarelo Suave
    "FINALIZADO": "C8E6C9"    // Verde Suave
};
// ====================================================================================

// Variáveis globais para armazenar os resultados
let lotesGerais = [];     
let lotesDetalhes = [];   
let todasCategorias = new Set(); 

// Inicialização do listener de habilitação do botão
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFileInput');
    const processButton = document.getElementById('processButton');
    
    fileInput.addEventListener('change', () => {
        processButton.disabled = fileInput.files.length === 0;
        document.getElementById('processButtonText').textContent = 'Processar Planilha';
    });
});

/**
 * Mapeia o nome do modelo para sua versão padrão, se uma regra existir.
 */
function mapearModeloEquivalente(modeloLimpo) {
    return MAPA_MODELO_GRUPO[modeloLimpo] || modeloLimpo;
}

/**
 * Funções de Limpeza de Dados
 */
function limparPrefixoModelo(modeloOriginal) {
    if (typeof modeloOriginal !== 'string') return String(modeloOriginal).trim().toUpperCase();
    let modelo = modeloOriginal.trim();
    const indexPrimeiroEspaco = modelo.indexOf(' ');
    if (indexPrimeiroEspaco !== -1) {
        modelo = modelo.substring(indexPrimeiroEspaco + 1).trim();
    }
    modelo = modelo.replace(/\s\s+/g, ' ');
    return modelo.toUpperCase();
}

function limparCategoria(nomeOriginal) {
    if (typeof nomeOriginal !== 'string') return String(nomeOriginal).trim().toUpperCase();
    let nome = nomeOriginal.trim().toUpperCase();
    nome = nome.replace(/ DE /g, ' ').trim(); 
    nome = nome.replace(/\s\s+/g, ' '); 
    return nome; 
}

/**
 * Funções de Interface e Exportação
 */
function gerarBotoesFiltro() {
    const containerXLSX = document.getElementById('filterColXLSX');
    const containerPDF = document.getElementById('filterColPDF');
    
    containerXLSX.innerHTML = '<h4>Download XLSX</h4>';
    containerPDF.innerHTML = '<h4>Download PDF</h4>';

    Array.from(todasCategorias).sort().forEach(categoria => {
        if (!categoria || categoria === 'N/A') return; 
        
        containerXLSX.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'xlsx', '${categoria}')" class="btn">${categoria}</button>`;
        containerPDF.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'pdf', '${categoria}')" class="btn">${categoria}</button>`;
    });
}

function selecionarColunas(data, isFiltered) {
    // Inclui a coluna STATUS
    const colunasPadrao = ['GRUPO', 'LINHA', 'BOJO', 'ALINHAMENTOS', 'DIMENSÃO', 'STATUS', 'QUANTIDADE TOTAL'];
    const colunasGerais = ['GRUPO', 'LINHA', 'BOJO', 'ALINHAMENTOS', 'STATUS', 'QUANTIDADE TOTAL'];

    let colunasFinais = data[0] && data[0].DIMENSÃO ? colunasPadrao : colunasGerais;

    if (isFiltered) {
        colunasFinais = colunasFinais.filter(col => col !== 'ALINHAMENTOS');
    }

    return data.map(item => {
        const novoItem = {};
        colunasFinais.forEach(col => {
            novoItem[col] = item[col];
        });
        return novoItem;
    });
}

function adicionarCabecalhoPersonalizado(ws, nomeRelatorio, dadosIniciais) {
    const colunasChave = Object.keys(dadosIniciais[0]);
    const numColunas = colunasChave.length; 
    const linhaInicial = 0; 
    
    const titulo = `RELATÓRIO OTIMIZADO DE LOTES - ${nomeRelatorio} | ${new Date().toLocaleDateString('pt-BR')}`;
    XLSX.utils.sheet_add_aoa(ws, [[titulo]], { origin: -1 });
    
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: linhaInicial, c: 0 }, e: { r: linhaInicial, c: numColunas - 1 } }); 

    const tituloCell = XLSX.utils.encode_cell({ r: linhaInicial, c: 0 });
    if (!ws[tituloCell]) ws[tituloCell] = { v: titulo, t: 's' };

    ws[tituloCell].s = {
        font: { name: "Arial", sz: 16, bold: true, color: { rgb: "FFFFFF" } }, 
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: "003366" } },
        border: { bottom: { style: "medium", color: { rgb: "000000" } } } 
    };
    
    return 1; 
}

/**
 * Aplica formatação e estilos no XLSX. 
 */
function aplicarFormatoBasico(dados, ws, nomeRelatorio, startRow) {
    if (dados.length === 0) return;

    const colunasChave = Object.keys(dados[0]);
    const numColunas = colunasChave.length;

    const range = { 
        s: { r: startRow, c: 0 }, 
        e: { r: startRow + dados.length, c: numColunas - 1 } 
    };
    ws['!ref'] = XLSX.utils.encode_range(range);
    
    // 1. Largura das Colunas
    ws['!cols'] = colunasChave.map(colName => {
        let wch = 25; 
        if (colName === 'GRUPO' || colName === 'BOJO') wch = 12; 
        if (colName === 'DIMENSÃO') wch = 15;
        if (colName === 'QUANTIDADE TOTAL') wch = 18;
        if (colName === 'STATUS') wch = 20; // Largura para status
        if (colName === 'LINHA') wch = 22; 
        return { wch: wch };
    });
    
    // 2. Filtro Automático
    ws['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
    
    // 3. Estilo do Cabeçalho
    const headerStyle = {
        fill: { fgColor: { rgb: "000000" } }, 
        font: { bold: true, color: { rgb: "FFFFFF" }, name: "Arial", sz: 10 }, 
        alignment: { horizontal: "center", vertical: "center" },
        border: { 
            top: { style: "medium", color: { rgb: "000000" } }, bottom: { style: "medium", color: { rgb: "FFFFFF" } }, 
            left: { style: "thin", color: { rgb: "FFFFFF" } }, right: { style: "thin", color: { rgb: "FFFFFF" } } 
        }
    };

    const centerStyle = { 
        alignment: { horizontal: "center", vertical: "center" }, 
        font: { name: "Arial", sz: 10 }
    };
    
    // Aplica estilo
    colunasChave.forEach((colName, index) => {
        // Cabeçalho
        const headerAddress = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c + index });
        const headerCell = ws[headerAddress];
        if (headerCell) {
            headerCell.s = headerStyle;
            if (typeof headerCell.v !== 'undefined') headerCell.t = 's';
        }

        // Corpo da Tabela
        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            const dataCellAddress = XLSX.utils.encode_cell({ r: r, c: range.s.c + index });
            const dataCell = ws[dataCellAddress];
            
            const grupoAddress = XLSX.utils.encode_cell({ r: r, c: 0 });
            const grupo = (ws[grupoAddress] && ws[grupoAddress].v) || 'N/A';
            let corFundo = GRUPO_COLORS[grupo] || GRUPO_COLORS['N/A'];

            // === LÓGICA DE COR PARA COLUNA STATUS ===
            if (colName === 'STATUS' && dataCell && dataCell.v) {
                const valorStatus = String(dataCell.v).toUpperCase();
                if (valorStatus.includes('A PRODUZIR')) {
                    corFundo = STATUS_COLORS['A PRODUZIR'];
                } else if (valorStatus.includes('MONTAGEM')) {
                    corFundo = STATUS_COLORS['EM MONTAGEM'];
                } else if (valorStatus.includes('FINALIZADO')) {
                    corFundo = STATUS_COLORS['FINALIZADO'];
                }
            }
            // ========================================

            if (dataCell) {
                const isNumeric = colName === 'QUANTIDADE TOTAL';
                
                const cellBorder = { 
                    border: { 
                        top: { style: "thin", color: { rgb: "DDDDDD" } }, bottom: { style: "thin", color: { rgb: "DDDDDD" } }, 
                        left: { style: "thin", color: { rgb: "DDDDDD" } }, right: { style: "thin", color: { rgb: "DDDDDD" } }
                    }
                };
                
                const fillStyle = { fill: { fgColor: { rgb: corFundo } } };
                
                dataCell.s = { 
                    ...centerStyle,
                    ...cellBorder,
                    ...fillStyle 
                };

                if (isNumeric) {
                    dataCell.t = 'n'; 
                } else if (typeof dataCell.v !== 'undefined') {
                    dataCell.t = 's'; 
                }
            }
        }
    });
}

function exportarXLSX(dadosParaExportar, nomeArquivo, isFiltered) {
    const dadosFinais = selecionarColunas(dadosParaExportar, isFiltered);
    
    const ws = {}; 
    const nomeRelatorioLimpo = nomeArquivo.replace(/_/g, ' '); 
    const totalTitleRows = adicionarCabecalhoPersonalizado(ws, nomeRelatorioLimpo, dadosFinais); 
    const START_ROW_DATA = totalTitleRows; 

    XLSX.utils.sheet_add_json(ws, dadosFinais, { origin: START_ROW_DATA, skipHeader: false });
    aplicarFormatoBasico(dadosFinais, ws, nomeRelatorioLimpo, START_ROW_DATA);
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lotes Produção");
    
    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`, { cellStyles: true }); 
}

function exportarOrdemProducaoXLSX(dados, nomeArquivo) {
    if (!dados || dados.length === 0) {
        alert("Não há dados para exportar.");
        return;
    }

    // 1. Agrupar dados
    const categoriaTitulo = dados[0].ALINHAMENTOS || "GERAL";
    const agrupado = {};

    dados.forEach(item => {
        const modelo = item.LINHA || "";
        const dimensao = item.DIMENSÃO || "";
        const bojo = (item.BOJO || "").toUpperCase();
        const qtd = item['QUANTIDADE TOTAL'] || 0;

        if (!modelo) return;

        if (!agrupado[modelo]) agrupado[modelo] = {};
        if (!agrupado[modelo][dimensao]) {
            agrupado[modelo][dimensao] = {
                modelo: modelo,
                dimensao: dimensao,
                total: 0,
                pvc: 0,
                inox: 0
            };
        }

        agrupado[modelo][dimensao].total += qtd;
        if (bojo.includes('INOX')) {
            agrupado[modelo][dimensao].inox += qtd;
        } else {
            agrupado[modelo][dimensao].pvc += qtd;
        }
    });

    // 2. Preparar dados para a planilha
    const ws_data = [];
    const merges = [];
    let currentRow = 0;

    // --- CABEÇALHO GERAL ---
    // LISTAGEM DE PRODUÇÃO
    ws_data.push(["LISTAGEM DE PRODUÇÃO"]);
    merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 7 } });
    currentRow++;

    // Linha azul (simulada com border bottom na célula anterior na etapa de estilo)

    // --- SUB-CABEÇALHO ---
    const dataAtual = new Date().toLocaleDateString('pt-BR');
    ws_data.push([`${categoriaTitulo} (SEM 048 A 050)`, "", "", "", "", "", "DATA:", dataAtual]);
    merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 5 } });
    currentRow++;

    // Espaço
    ws_data.push([""]);
    currentRow++;

    // --- ITERAR PELOS MODELOS ---
    let totalGeral = 0;
    let totalPVC = 0;
    let totalInox = 0;

    const modelos = Object.keys(agrupado).sort();

    modelos.forEach(modelo => {
        // Modelo Header
        ws_data.push([modelo.toUpperCase()]);
        merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 2 } });
        const rowModelo = currentRow;
        currentRow++;

        // Table Headers
        ws_data.push(["SEQ", "PEDIDO - CLIENTE", "QTD PRODUÇÃO", "PVC", "INOX", "ESTOQUE BOJO PP", "ESTOQUE BOJO INOX", "PRODUZIR"]);
        const rowHeader = currentRow;
        currentRow++;

        const dimensoes = Object.keys(agrupado[modelo]).sort((a,b) => {
            const valA = parseFloat(a.replace(',', '.'));
            const valB = parseFloat(b.replace(',', '.'));
            return (valA || 0) - (valB || 0);
        });

        dimensoes.forEach(dim => {
            const dadosDim = agrupado[modelo][dim];
            const nomePedido = `${modelo} ${dim}`;

            ws_data.push([
                "X",
                nomePedido,
                dadosDim.total,
                dadosDim.pvc,
                dadosDim.inox,
                "FABRICAR",
                "FABRICAR",
                "FABRICAR"
            ]);
            currentRow++;

            // Empty Row (Gap)
            ws_data.push(["", "", "", "", "", "", "", ""]);
            currentRow++;

            totalGeral += dadosDim.total;
            totalPVC += dadosDim.pvc;
            totalInox += dadosDim.inox;
        });
    });

    // --- TOTAIS ---
    const rowTotais = currentRow;
    ws_data.push(["", "", totalGeral, totalPVC, totalInox, "", "", ""]);
    currentRow++;

    // --- CRIAR PLANILHA ---
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    ws['!merges'] = merges;
    ws['!cols'] = [
        { wch: 5 },  // SEQ
        { wch: 40 }, // PEDIDO
        { wch: 15 }, // QTD
        { wch: 10 }, // PVC
        { wch: 10 }, // INOX
        { wch: 20 }, // ESTOQUE PP
        { wch: 20 }, // ESTOQUE INOX
        { wch: 20 }  // PRODUZIR
    ];

    // --- APLICAR ESTILOS ---
    const range = XLSX.utils.decode_range(ws['!ref']);

    // Estilo Base
    const styleCenter = { alignment: { horizontal: "center", vertical: "center" }, font: { name: "Arial", sz: 10 } };
    const styleBold = { font: { name: "Arial", sz: 10, bold: true } };
    const borderAll = {
        top: { style: "thin", color: { rgb: "000000" } },
        bottom: { style: "thin", color: { rgb: "000000" } },
        left: { style: "thin", color: { rgb: "000000" } },
        right: { style: "thin", color: { rgb: "000000" } }
    };
    const borderThickBox = {
        top: { style: "medium", color: { rgb: "000000" } },
        bottom: { style: "medium", color: { rgb: "000000" } },
        left: { style: "medium", color: { rgb: "000000" } },
        right: { style: "medium", color: { rgb: "000000" } }
    };

    // Estilo Título Principal (Row 0)
    const cellA1 = XLSX.utils.encode_cell({c:0, r:0});
    if(ws[cellA1]) {
        ws[cellA1].s = {
            font: { name: "Arial", sz: 14, bold: true, color: { rgb: "3366CC" } },
            alignment: { horizontal: "center" },
            border: { bottom: { style: "thick", color: { rgb: "3366CC" } } } // Blue bar equivalent
        };
    }

    // Sub-cabeçalho (Row 1)
    const cellA2 = XLSX.utils.encode_cell({c:0, r:1});
    if(ws[cellA2]) {
        ws[cellA2].s = {
            font: { name: "Arial", sz: 12, bold: true },
            alignment: { horizontal: "left" }
        };
    }
    const cellDataLabel = XLSX.utils.encode_cell({c:6, r:1}); // "DATA:"
    if(ws[cellDataLabel]) ws[cellDataLabel].s = { font: { bold: true, color: { rgb: "3366CC" } }, alignment: { horizontal: "right" } };
    const cellDataValue = XLSX.utils.encode_cell({c:7, r:1});
    if(ws[cellDataValue]) ws[cellDataValue].s = { font: { bold: true, italic: true, color: { rgb: "3366CC" } }, alignment: { horizontal: "center" } };

    // Iterar para aplicar estilos de linhas de dados
    let r = 3; // Começa após os cabeçalhos (Row 0, 1, 2 empty)

    // Note: r needs to track the logical rows we added.
    // However, I constructed ws_data sequentially, so I can just iterate through ws_data indices.

    // Helper to find row types is hard because I didn't store metadata per row.
    // I will iterate range.

    for (let R = 3; R <= range.e.r; ++R) {
        // Verificar conteúdo da primeira célula para identificar o tipo de linha
        const cellFirst = ws[XLSX.utils.encode_cell({c:0, r:R})];
        const valFirst = cellFirst ? cellFirst.v : "";

        // Header da Tabela (SEQ)
        if (valFirst === "SEQ") {
            for (let C = 0; C <= 7; ++C) {
                const cell = ws[XLSX.utils.encode_cell({c:C, r:R})];
                if (cell) {
                    cell.s = {
                        fill: { fgColor: { rgb: "FFF2CC" } }, // Beige
                        font: { bold: true, sz: 9 },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: borderAll
                    };
                }
            }
            continue;
        }

        // Linha de Dados (SEQ = X)
        if (valFirst === "X") {
            // Coluna SEQ (Boxed X)
            if (cellFirst) {
                cellFirst.s = {
                    alignment: { horizontal: "center", vertical: "center" },
                    font: { bold: true },
                    border: borderThickBox
                };
            }

            // Colunas de Dados (Preto com Texto Branco)
            // PEDIDO (1), QTD (2), PVC (3), INOX (4)
            for (let C = 1; C <= 4; ++C) {
                const cell = ws[XLSX.utils.encode_cell({c:C, r:R})];
                if (cell) {
                    cell.s = {
                        fill: { fgColor: { rgb: "000000" } },
                        font: { color: { rgb: "FFFFFF" }, bold: true },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "FFFFFF" } },
                            bottom: { style: "thin", color: { rgb: "FFFFFF" } },
                            left: { style: "thin", color: { rgb: "FFFFFF" } },
                            right: { style: "thin", color: { rgb: "FFFFFF" } }
                        }
                    };
                    if (C === 1) cell.s.alignment.horizontal = "left"; // Pedido align left
                }
            }

            // Colunas FABRICAR (5, 6, 7)
            for (let C = 5; C <= 7; ++C) {
                const cell = ws[XLSX.utils.encode_cell({c:C, r:R})];
                if (cell) {
                    cell.s = {
                        fill: { fgColor: { rgb: "FFFFFF" } },
                        font: { color: { rgb: "000000" }, bold: true },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: borderThickBox
                    };
                }
            }
            continue;
        }

        // Modelo Header (Assume it's the row BEFORE SEQ, but hard to track backwards)
        // Check if it matches a model name?
        // Let's assume if it is NOT SEQ, NOT X, NOT Empty, NOT Totals...
        // And has value in Col 0 but merged...

        // Actually, Totals row has empty first cell.
        // Empty row has empty first cell.

        // Model Header has value in Col 0 and is not SEQ or X.
        if (valFirst && valFirst !== "SEQ" && valFirst !== "X" && valFirst !== "LISTAGEM DE PRODUÇÃO" && !String(valFirst).includes("EXPOSITORES")) {
            // It is likely a Model Header (e.g. VERTICAL ALTOS)
            const cell = ws[XLSX.utils.encode_cell({c:0, r:R})];
            if (cell) {
                cell.s = {
                    font: { bold: true, sz: 11 },
                    border: {
                        top: { style: "double" }, bottom: { style: "double" },
                        left: { style: "double" }, right: { style: "double" }
                    },
                    alignment: { horizontal: "left" }
                };
            }
        }

        // Totals Row (empty first cell, numbers in 2,3,4)
        if (!valFirst) {
            const cellTotal = ws[XLSX.utils.encode_cell({c:2, r:R})];
            if (cellTotal && typeof cellTotal.v === 'number') {
                // This is the totals row
                for (let C = 0; C <= 7; ++C) {
                    const cell = ws[XLSX.utils.encode_cell({c:C, r:R})];
                    if (cell) {
                        cell.s = {
                            border: { top: { style: "double" }, bottom: { style: "double" } },
                            font: { bold: true },
                            alignment: { horizontal: "center" }
                        };
                    } else {
                        // Create empty cell for border
                         ws[XLSX.utils.encode_cell({c:C, r:R})] = { v: "", s: { border: { top: { style: "double" }, bottom: { style: "double" } } } };
                    }
                }
            }
        }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Ordem Produção");
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual.replace(/\//g, '-')}.xlsx`, { cellStyles: true });
}

function exportarPDF(dadosParaExportar, nomeArquivo, isFiltered) {
    if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') {
        alert("Erro: Biblioteca PDF não carregada.");
        return;
    }
    
    const dadosFinais = selecionarColunas(dadosParaExportar, isFiltered);
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    const headersMap = {
        'GRUPO': 'GRUPO', 
        'LINHA': 'Código do Modelo',
        'BOJO': 'Material (BOJO)',
        'ALINHAMENTOS': 'Categoria',
        'DIMENSÃO': 'DIMENSÃO', 
        'STATUS': 'Status', 
        'QUANTIDADE TOTAL': 'Quantidade Total'
    };
    
    const colunas = Object.keys(dadosFinais[0]).filter(key => headersMap[key]);
    const head = [colunas.map(key => headersMap[key])];
    const body = dadosFinais.map(item => colunas.map(key => item[key]));

    doc.autoTable({
        head: head,
        body: body,
        startY: 20,
        theme: 'grid', 
        styles: { 
            fontSize: 9, 
            font: 'helvetica', 
            textColor: [52, 58, 64] 
        },
        headStyles: { 
            fillColor: [0, 123, 255], 
            textColor: 255, 
            fontStyle: 'bold',
            fontSize: 10
        }, 
        alternateRowStyles: {
            fillColor: [248, 249, 250] 
        },
        bodyStyles: {
            lineColor: [222, 226, 230], 
            lineWidth: 0.1 
        },
        didDrawPage: function (data) {
            doc.setFontSize(14);
            doc.text("Lista Otimizada de Lotes - PCP", data.settings.margin.left, 10);
            doc.setFontSize(10);
            doc.text(`Relatório: ${nomeArquivo.replace(/_/g, ' ')} | Data: ${new Date().toLocaleDateString('pt-BR')}`, data.settings.margin.left, 15);
        }
    });

    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    doc.save(`${nomeArquivo}_${dataAtual}.pdf`);
}

function exportarRelatorio(tipo, formato, filtroCategoria = null) {
    let dadosParaExportar;
    let nomeArquivo;
    let isFiltered = filtroCategoria !== null; 

    if (tipo === 'geral') {
        dadosParaExportar = lotesGerais;
        nomeArquivo = 'Resumo_Geral_Lotes';
    } else { 
        if (isFiltered) {
            dadosParaExportar = lotesDetalhes.filter(lote => lote.ALINHAMENTOS === filtroCategoria);
            nomeArquivo = `Detalhe_Lotes_${filtroCategoria.replace(/\s/g, '_')}`;
        } else {
            dadosParaExportar = lotesDetalhes;
            nomeArquivo = 'Detalhe_Lotes_Completo';
        }
    }

    if (dadosParaExportar.length === 0) {
        alert("Nenhum dado encontrado para o filtro selecionado.");
        return;
    }

    if (formato === 'xlsx') {
        if (tipo === 'detalhe') {
             exportarOrdemProducaoXLSX(dadosParaExportar, nomeArquivo);
        } else {
             exportarXLSX(dadosParaExportar, nomeArquivo, isFiltered);
        }
    } else if (formato === 'pdf') {
        exportarPDF(dadosParaExportar, nomeArquivo, isFiltered); 
    }
}

function processarPlanilha() {
    const fileInput = document.getElementById('excelFileInput');
    const statusDiv = document.getElementById('statusMessage');
    
    statusDiv.textContent = 'Processando...';
    document.getElementById('processButton').disabled = true;
    document.getElementById('downloadSection').style.display = 'none'; 

    const file = fileInput.files[0];
    if (!file) {
        statusDiv.textContent = 'Erro: Nenhum arquivo selecionado.';
        document.getElementById('processButton').disabled = false;
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Define o intervalo de leitura na Linha 5 (Índice 4)
            const range = XLSX.utils.decode_range(worksheet['!ref']);
            range.s.r = 4; 
            const newRange = XLSX.utils.encode_range(range);
            
            const rawDataAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: newRange }); 

            if (!rawDataAOA || rawDataAOA.length === 0) {
                throw new Error("A planilha parece estar vazia ou o intervalo de leitura está incorreto.");
            }

            const headers = rawDataAOA[0].map(h => String(h || '').toUpperCase().trim());
            const INDEX_CATEGORIA = headers.indexOf('LINHA');
            const INDEX_MODELO = headers.indexOf('ALINHAMENTO');
            const INDEX_BOJO = headers.indexOf('BOJO');
            const INDEX_DIMENSAO = headers.indexOf('DIMENSÃO');
            const INDEX_STATUS = headers.indexOf('MONTAGEM'); 

            if (INDEX_CATEGORIA === -1 || INDEX_MODELO === -1) {
                const headersEncontrados = headers.join(", ");
                const msg = `Erro Crítico: Colunas não encontradas. Headers: [${headersEncontrados}]`;
                alert(msg);
                throw new Error(msg);
            }

            const lotesGeraisMap = {};   
            const lotesDetalhesMap = {}; 
            todasCategorias.clear();
            let linhasProcessadas = 0;

            for (let i = 1; i < rawDataAOA.length; i++) {
                const row = rawDataAOA[i];
                const maxIndex = Math.max(INDEX_CATEGORIA, INDEX_MODELO, INDEX_BOJO, INDEX_DIMENSAO);
                if (row.length <= maxIndex) continue;

                const modeloOriginal = String(row[INDEX_MODELO] || '').trim();
                const dimensaoOriginal = String(row[INDEX_DIMENSAO] || '').trim();
                const bojoOriginal = String(row[INDEX_BOJO] || '').trim();
                const categoriaOriginal = String(row[INDEX_CATEGORIA] || '').trim();
                
                let statusOriginal = "A PRODUZIR";
                if (INDEX_STATUS !== -1) {
                    const valorCelula = String(row[INDEX_STATUS] || '').trim().toUpperCase();
                    if (valorCelula !== "") {
                        statusOriginal = valorCelula;
                    }
                }

                const modeloLimpo = limparPrefixoModelo(modeloOriginal);
                const grupoLetra = MAPA_MODELO_GRUPO[modeloLimpo] || modeloLimpo;
                const categoriaLimpa = limparCategoria(categoriaOriginal); 
                const bojoNormalizado = (bojoOriginal || 'N/A').toUpperCase();
                const dimensaoNormalizada = (dimensaoOriginal || 'N/A').toUpperCase();

                if (!modeloLimpo || !categoriaLimpa) continue; 

                linhasProcessadas++;
                todasCategorias.add(categoriaLimpa);

                const chaveGeral = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${grupoLetra}|${statusOriginal}`; 
                const chaveDetalhe = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${dimensaoNormalizada}|${grupoLetra}|${statusOriginal}`; 

                if (lotesGeraisMap[chaveGeral]) {
                    lotesGeraisMap[chaveGeral]['QUANTIDADE TOTAL']++;
                } else {
                    lotesGeraisMap[chaveGeral] = {
                        'GRUPO': grupoLetra, 
                        'LINHA': modeloLimpo, 
                        'BOJO': bojoNormalizado,
                        'ALINHAMENTOS': categoriaLimpa,
                        'STATUS': statusOriginal,
                        'QUANTIDADE TOTAL': 1
                    };
                }
                
                if (lotesDetalhesMap[chaveDetalhe]) {
                    lotesDetalhesMap[chaveDetalhe]['QUANTIDADE TOTAL']++;
                } else {
                    lotesDetalhesMap[chaveDetalhe] = {
                        'GRUPO': grupoLetra, 
                        'LINHA': modeloLimpo, 
                        'BOJO': bojoNormalizado,
                        'ALINHAMENTOS': categoriaLimpa,
                        'DIMENSÃO': dimensaoNormalizada, 
                        'STATUS': statusOriginal, 
                        'QUANTIDADE TOTAL': 1
                    };
                }
            }

            lotesGerais = Object.values(lotesGeraisMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));
            lotesDetalhes = Object.values(lotesDetalhesMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));

            if (linhasProcessadas === 0) {
                 statusDiv.textContent = `Nenhuma linha de dados processada.`;
                 document.getElementById('processButton').disabled = false;
                 return;
            }

            statusDiv.textContent = `Processamento concluído! Itens: ${linhasProcessadas}.`;
            document.getElementById('downloadSection').style.display = 'block';
            gerarBotoesFiltro(); 

        } catch (error) {
            console.error("Erro:", error);
            statusDiv.textContent = `Erro: ${error.message}`;
        } finally {
            document.getElementById('processButton').disabled = false;
        }
    };

    reader.onerror = function(ex) {
        statusDiv.textContent = 'Erro ao ler o arquivo.';
        document.getElementById('processButton').disabled = false;
    };

    reader.readAsArrayBuffer(file);
}