// script.js (Após a correção)

// ====================================================================================
// SISTEMA NSF - OTIMIZADOR DE LOTES PCP
// FORMATO EXATO DA IMAGEM - ORDEM DE PRODUÇÃO PROFISSIONAL
// ====================================================================================

/* global XLSX */ // Apenas um aviso para o linter, não cria a variável

// ====================================================================================
// CONFIGURAÇÕES GLOBAIS E CONSTANTES
// ... (o restante do seu código)


const CONFIG_NSF = {
  // Cores corporativas NSF - ATUALIZADAS PARA MATCH COM A IMAGEM
  CORES: {
    PRETO: "000000",
    BRANCO: "FFFFFF",
    VERMELHO: "FF0000",
    CINZA_CLARO: "D9D9D9",
    CINZA_MEDIO: "BFBFBF",
    CINZA_ESCURO: "7F7F7F",
    AMARELO_CABECALHO: "D4A84B", // Cor dourada/amarela do cabeçalho da tabela
    AZUL_BOTAO: "4A90D9", // Cor azul dos botões FABRICAR
    VERDE_LOGO: "2E7D32", // Verde do logo NSF
    FUNDO_LINHA: "FFFFFF", // Fundo branco das linhas
    BORDA_FILTRO: "000000", // Borda preta do filtro
  },

  // Configurações de fonte
  FONTES: {
    PRINCIPAL: "Arial",
    TAMANHO_TITULO: 14,
    TAMANHO_SUBTITULO: 11,
    TAMANHO_CABECALHO: 9,
    TAMANHO_DADOS: 9,
    TAMANHO_ASSINATURA: 10,
    TAMANHO_DATA: 9,
  },

  // Configurações de layout - ATUALIZADAS
  LAYOUT: {
    LARGURA_COLUNA_SEQ: 5,
    LARGURA_COLUNA_PEDIDO: 30,
    LARGURA_COLUNA_QTD: 14,
    LARGURA_COLUNA_PVC: 8,
    LARGURA_COLUNA_INOX: 8,
    LARGURA_COLUNA_ESTOQUE_PP: 16,
    LARGURA_COLUNA_ESTOQUE_INOX: 18,
    LARGURA_COLUNA_PRODUZIR: 12,

    ALTURA_LINHA_TITULO: 22,
    ALTURA_LINHA_SUBTITULO: 18,
    ALTURA_LINHA_GRUPO: 18,
    ALTURA_LINHA_CABECALHO: 22,
    ALTURA_LINHA_DADOS: 20,
    ALTURA_LINHA_VAZIA: 12,
  },

  // Textos fixos
  TEXTOS: {
    TITULO_PRINCIPAL: "LISTAGEM DE PRODUÇÃO",
    FABRICAR: "FABRICAR",
    SUPERVISOR: "SUPERVISOR DE PRODUÇÃO",
    OPERADOR_01: "OPERADOR 01",
    OPERADOR_02: "OPERADOR 02",
    REP: "REP",
  },
}

// ====================================================================================
// MAPA COMPLETO DE AGRUPAMENTO DE MODELOS
// ====================================================================================

const MAPA_MODELO_GRUPO = {
  // Grupo A (Valp-850 e Variações)
  "VABP-PVT/850": "A",
  "VABP/850 MA": "A",
  "VACP-PVT/850": "A",
  "VACP-PVT/850 MA": "A",
  "VACP/850": "A",
  "VACP/750": "A",
  "VACP/850 MA": "A",
  "VAHP-PVT/850 MA": "A",
  "VAHP-PVT/850": "A",
  "VAHP/850": "A",
  "VALP- PVT/850": "A",
  "VALP-PVT/850": "A",
  "VALP-PVT/850 MA": "A",
  "VALP/850": "A",
  "VALP/750": "A",
  "VALP/750 MA": "A",
  "VABP/850": "A",
  "VABP/750": "A",
  "VALP/850 MA": "A",

  // Grupo B (VAH/850)
  "VAH/850": "B",
  "VAL/850": "B",

  // Grupo C (VAP/850 e Variações)
  "VAP/850": "C",
  "VAP/850 MA": "C",

  // Grupo D (VCA2P/1040)
  "VCA2P/1040": "D",

  // Grupo E (VCAG Complexo)
  "VCAG (1,25) + VCA2P (2,50)/1040": "E",
  "VCAG/1040": "E",
  "VCAGR (1,25) + VCAG1P (1,25)/1040": "E",

  // Grupo F (VIL-2P/900)
  "VIL-2P/900": "F",
  "VIL-2P CENTRAL/1760": "F",
  "VIL-2P FRONTAL/900": "F",

  // Grupo G (VIL-2P/900 CANTO)
  "VIL-2P/900 CANTO": "G",

  // Grupo H (VILP-2P/900 e Variações)
  "VILP-2P/900": "H",
  "VILP-2P/900 MA": "H",

  // Grupo I (VR-900 e Variações)
  "VR1P/900": "I",
  "VR2P/900": "I",
  "VR2P/900 MA": "I",
  "VR2PA/900": "I",
  "VRQU/900": "I",
  "VRQR/900": "I",
  "VRHB/900": "I",

  // Grupo J
  "VRA1P/1040": "J",
  "VRA2P/1040": "J",
  "VRAG (1,25) + VRA2P (2,50)/1040": "J",
  "VRAG/1040": "J",
  "VRAGR/1040": "J",
  "VRAG(3,75) + VRAG1P (1,25)/1040": "J",
  "VRAG2N/900": "J",
  "VRAG (2,50) +VRA1P (1,25)/1040": "J",

  // Grupo K
  "VIL-2P PONTA CT 180/900": "K",

  // Grupo L
  "VIL-3P/900": "L",

  // Grupo M
  ICDT: "M",
  ICFT: "M",

  // Grupo N
  "VC2P/900": "N",
  "VCQU/900": "N",

  // Grupo O
  IRAS: "O",
}

// ====================================================================================
// MAPA DE NOMES DE GRUPOS PARA EXIBIÇÃO
// ====================================================================================

const MAPA_NOME_GRUPO = {
  A: "VERTICAL ALTOS",
  B: "VERTICAL ALTOS",
  C: "VERTICAL ALTOS",
  D: "AÇOUGUE CURVO",
  E: "AÇOUGUE CURVO",
  F: "ILHAS",
  G: "ILHAS CANTO",
  H: "ILHAS",
  I: "REFRIGERAÇÃO",
  J: "REFRIGERAÇÃO ALTA",
  K: "ILHAS PONTA",
  L: "ILHAS 3P",
  M: "INTELIGENTE",
  N: "AÇOUGUE",
  O: "INTELIGENTE",
}

// ====================================================================================
// VARIÁVEIS GLOBAIS DO SISTEMA
// ====================================================================================

let lotesGerais = []
let lotesDetalhes = []
const todasCategorias = new Set()
const todosGrupos = new Set()

// ====================================================================================
// INICIALIZAÇÃO DO SISTEMA
// ====================================================================================

// script.js (Parte NOVA)
document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("excelFileInput")
  const processButton = document.getElementById("processButton")
  
  // O botão começa ATIVO no HTML (após a alteração acima).
  // Se você realmente quiser que ele comece desativado e só ative com o arquivo,
  // mantenha a lógica original abaixo.
  
  fileInput.addEventListener("change", () => {
    // 💡 Recomendação: Manter este bloco para garantir a funcionalidade
    processButton.disabled = fileInput.files.length === 0 
    document.getElementById("processButtonText").textContent = "Processar Planilha"
  })
})

// ====================================================================================
// FUNÇÕES AUXILIARES DE FORMATAÇÃO
// ====================================================================================

function getNomeCompletoGrupo(grupo) {
  return MAPA_NOME_GRUPO[grupo] || `GRUPO ${grupo}`
}

function getCategoriaGrupo(grupo) {
  const categorias = {
    A: "EXPOSITORES - REFRIGERAÇÃO",
    B: "EXPOSITORES - REFRIGERAÇÃO",
    C: "EXPOSITORES - REFRIGERAÇÃO",
    D: "EXPOSITORES - AÇOUGUE",
    E: "EXPOSITORES - AÇOUGUE",
    F: "ILHAS - CONGELADOS",
    G: "ILHAS - CONGELADOS",
    H: "ILHAS - CONGELADOS",
    I: "REFRIGERAÇÃO",
    J: "REFRIGERAÇÃO ALTA",
    K: "ILHAS - PONTA",
    L: "ILHAS - 3P",
    M: "INTELIGENTE",
    N: "AÇOUGUE",
    O: "INTELIGENTE",
  }
  return categorias[grupo] || "EXPOSITORES"
}

function formatarDimensaoCompleta(dimensao) {
  if (!dimensao || dimensao === "N/A") return "N/A"

  let dim = dimensao.toString().replace(/\s/g, "").replace(".", ",")

  // Formato da imagem: 1,250 / 1,500 / 2,500 etc
  if (!dim.includes(",")) {
    dim = dim + ",000"
  }

  return dim
}

function getSemanaAtualFormatada() {
  const hoje = new Date()
  const inicioAno = new Date(hoje.getFullYear(), 0, 1)
  const dias = Math.floor((hoje - inicioAno) / (24 * 60 * 60 * 1000))
  const semana = Math.ceil((dias + 1) / 7)

  return `SEM ${String(semana).padStart(3, "0")} A ${String(semana + 2).padStart(3, "0")}`
}

function getDataAtualFormatada() {
  const hoje = new Date()
  const dia = String(hoje.getDate()).padStart(2, "0")
  const mes = String(hoje.getMonth() + 1).padStart(2, "0")
  const ano = hoje.getFullYear()
  return `${dia}/${mes}/${ano}`
}

// ====================================================================================
// FUNÇÕES DE PROCESSAMENTO DE DADOS
// ====================================================================================

function limparPrefixoModelo(modeloOriginal) {
  if (typeof modeloOriginal !== "string") return String(modeloOriginal).trim().toUpperCase()
  let modelo = modeloOriginal.trim()
  const indexPrimeiroEspaco = modelo.indexOf(" ")
  if (indexPrimeiroEspaco !== -1) {
    modelo = modelo.substring(indexPrimeiroEspaco + 1).trim()
  }
  modelo = modelo.replace(/\s\s+/g, " ")
  return modelo.toUpperCase()
}

function limparCategoria(nomeOriginal) {
  if (typeof nomeOriginal !== "string") return String(nomeOriginal).trim().toUpperCase()
  let nome = nomeOriginal.trim().toUpperCase()
  nome = nome.replace(/ DE /g, " ").trim()
  nome = nome.replace(/\s\s+/g, " ")
  return nome
}

function normalizarDimensao(dimensao) {
  if (!dimensao || dimensao === "N/A") return "N/A"
  return dimensao.toString().replace(/\s/g, "").replace(",", ".")
}

// ====================================================================================
// FUNÇÕES DE INTERFACE
// ====================================================================================

function gerarBotoesFiltro() {
  const containerXLSX = document.getElementById("filterColXLSX")
  const containerPDF = document.getElementById("filterColPDF")

  containerXLSX.innerHTML = "<h4>Download XLSX</h4>"
  containerPDF.innerHTML = "<h4>Download PDF</h4>"

  Array.from(todasCategorias)
    .sort()
    .forEach((categoria) => {
      if (!categoria || categoria === "N/A") return

      containerXLSX.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'xlsx', '${categoria}')" class="btn">${categoria}</button>`
      containerPDF.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'pdf', '${categoria}')" class="btn">${categoria}</button>`
    })
}

function gerarBotoesOP() {
  const containerXLSX = document.getElementById("filterColOPXLSX")
  const containerPDF = document.getElementById("filterColOPPDF")

  containerXLSX.innerHTML = "<h4>Ordem de Produção (XLSX)</h4>"
  containerPDF.innerHTML = "<h4>Ordem de Produção (PDF)</h4>"

  Array.from(todosGrupos)
    .sort()
    .forEach((grupo) => {
      if (!grupo || grupo === "N/A") return

      containerXLSX.innerHTML += `<button onclick="exportarRelatorio('op', 'xlsx', '${grupo}')" class="btn">${grupo}</button>`
      containerPDF.innerHTML += `<button onclick="exportarRelatorio('op', 'pdf', '${grupo}')" class="btn">${grupo}</button>`
    })
}

// ====================================================================================
// FUNÇÃO PRINCIPAL DE EXPORTAÇÃO XLSX - LAYOUT IDÊNTICO À IMAGEM
// ====================================================================================

function exportarXLSX(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail) {
  if (tipo === "op") {
    // ========================================================================
    // LAYOUT EXATAMENTE IGUAL À IMAGEM - ORDEM DE PRODUÇÃO NSF
    // ========================================================================

    const dadosFinais = selecionarColunas(dadosParaExportar, tipo, isFilteredDetail)
    const nomeGrupo = nomeArquivo.replace("OP_GRUPO_", "").toUpperCase()

    const wb = XLSX.utils.book_new()
    const ws = {}

    // Definir range total da planilha
    const totalLinhas = 25 + dadosFinais.length * 2 // Estimativa
    const totalColunas = 18 // Colunas A até R

    // =============================================
    // LINHA 1-4: ÁREA DO LOGO (deixar vazia para o logo)
    // =============================================
    for (let row = 0; row < 4; row++) {
      for (let col = 0; col < 4; col++) {
        const cell = XLSX.utils.encode_cell({ r: row, c: col })
        ws[cell] = { v: "", t: "s", s: { fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } } } }
      }
    }

    // =============================================
    // LINHA 2: TÍTULO "LISTAGEM DE PRODUÇÃO" - Centralizado nas colunas E-N
    // =============================================
    const tituloCell = "E2"
    ws[tituloCell] = {
      v: CONFIG_NSF.TEXTOS.TITULO_PRINCIPAL,
      t: "s",
      s: {
        font: {
          name: CONFIG_NSF.FONTES.PRINCIPAL,
          sz: 16,
          bold: true,
          color: { rgb: CONFIG_NSF.CORES.PRETO },
        },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }
    if (!ws["!merges"]) ws["!merges"] = []
    ws["!merges"].push({ s: { r: 1, c: 4 }, e: { r: 1, c: 13 } }) // E2:N2

    // =============================================
    // LINHA 5: SUBTÍTULO COM CATEGORIA E SEMANA + DATA À DIREITA
    // =============================================
    const categoriaGrupo = getCategoriaGrupo(nomeGrupo)
    const semanaFormatada = getSemanaAtualFormatada()
    const subtitulo = `${categoriaGrupo} (${semanaFormatada})`

    // Subtítulo à esquerda (A5:L5)
    ws["A5"] = {
      v: subtitulo,
      t: "s",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 12, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "left", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }
    ws["!merges"].push({ s: { r: 4, c: 0 }, e: { r: 4, c: 11 } }) // A5:L5

    // DATA à direita (Q5:R5)
    ws["Q5"] = {
      v: "DATA:",
      t: "s",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 9, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "right", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }
    ws["R5"] = {
      v: getDataAtualFormatada(),
      t: "s",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 9, bold: false, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "left", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }

    // =============================================
    // LINHA 7-8: FILTRO DO GRUPO (caixa com borda)
    // =============================================
    const nomeCompletoGrupo = getNomeCompletoGrupo(nomeGrupo)
    ws["B7"] = {
      v: nomeCompletoGrupo,
      t: "s",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: false, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "left", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
        border: {
          top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
        },
      },
    }
    ws["!merges"].push({ s: { r: 6, c: 1 }, e: { r: 7, c: 3 } }) // B7:D8

    // =============================================
    // LINHA 10: CABEÇALHO DA TABELA - AMARELO/DOURADO COM FILTROS
    // =============================================
    const headers = [
      { col: 0, colEnd: 1, text: "SEQ", width: 2 },
      { col: 2, colEnd: 4, text: "PEDIDO - CLIENTE", width: 3 },
      { col: 5, colEnd: 6, text: "QTD PRODUÇÃO", width: 2 },
      { col: 7, colEnd: 7, text: "PVC", width: 1 },
      { col: 8, colEnd: 8, text: "INOX", width: 1 },
      { col: 9, colEnd: 11, text: "ESTOQUE BOJO PP", width: 3 },
      { col: 12, colEnd: 14, text: "ESTOQUE BOJO INOX", width: 3 },
      { col: 15, colEnd: 17, text: "PRODUZIR", width: 3 },
    ]

    const headerRow = 9 // Linha 10 (0-indexed = 9)

    headers.forEach((header) => {
      const startCell = XLSX.utils.encode_cell({ r: headerRow, c: header.col })
      ws[startCell] = {
        v: header.text + " ▼", // Adiciona o dropdown indicator
        t: "s",
        s: {
          font: {
            name: CONFIG_NSF.FONTES.PRINCIPAL,
            sz: 9,
            bold: true,
            color: { rgb: CONFIG_NSF.CORES.PRETO },
          },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AMARELO_CABECALHO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
          alignment: { horizontal: "center", vertical: "center" },
        },
      }

      // Merge das células do cabeçalho
      if (header.colEnd > header.col) {
        ws["!merges"].push({ s: { r: headerRow, c: header.col }, e: { r: headerRow, c: header.colEnd } })
      }

      // Preencher células mescladas com estilo
      for (let c = header.col + 1; c <= header.colEnd; c++) {
        const fillCell = XLSX.utils.encode_cell({ r: headerRow, c: c })
        ws[fillCell] = {
          v: "",
          t: "s",
          s: {
            fill: { fgColor: { rgb: CONFIG_NSF.CORES.AMARELO_CABECALHO } },
            border: {
              top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            },
          },
        }
      }
    })

    // =============================================
    // LINHAS DE DADOS - FORMATO EXATO DA IMAGEM
    // =============================================
    let currentDataRow = headerRow + 1 // Começa na linha 11
    let totalQtd = 0
    let totalPVC = 0
    let totalInox = 0

    dadosFinais.forEach((item, index) => {
      const dimensao = item["DIMENSÃO"] || "N/A"
      const dimensaoFormatada = formatarDimensaoCompleta(dimensao)
      const nomeGrupoDisplay = getNomeCompletoGrupo(item["GRUPO"] || nomeGrupo)
      const pedidoCliente = `${nomeGrupoDisplay} ${dimensaoFormatada}`
      const quantidade = Number.parseInt(item["QTD. FABRICAR"]) || 0

      // Calcular PVC e INOX baseado no material (bojo)
      const material = (item["MATERIAL (BOJO)"] || item["BOJO"] || "").toUpperCase()
      const isPVC = material.includes("PVC") || material.includes("PP")
      const qtdPVC = isPVC ? quantidade : 0
      const qtdInox = !isPVC ? quantidade : 0

      totalQtd += quantidade
      totalPVC += qtdPVC
      totalInox += qtdInox

      // Coluna SEQ (checkbox X) - A-B
      const seqCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 0 })
      ws[seqCell] = {
        v: "X",
        t: "s",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }
      ws["!merges"].push({ s: { r: currentDataRow, c: 0 }, e: { r: currentDataRow, c: 1 } })

      // Célula vazia para merge SEQ
      ws[XLSX.utils.encode_cell({ r: currentDataRow, c: 1 })] = {
        v: "",
        t: "s",
        s: {
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }

      // Coluna PEDIDO-CLIENTE (nome com dimensão) - C-E
      const pedidoCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 2 })
      ws[pedidoCell] = {
        v: pedidoCliente,
        t: "s",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
          alignment: { horizontal: "left", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.CINZA_CLARO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }
      ws["!merges"].push({ s: { r: currentDataRow, c: 2 }, e: { r: currentDataRow, c: 4 } })

      // Preencher células mescladas PEDIDO-CLIENTE
      for (let c = 3; c <= 4; c++) {
        ws[XLSX.utils.encode_cell({ r: currentDataRow, c: c })] = {
          v: "",
          t: "s",
          s: {
            fill: { fgColor: { rgb: CONFIG_NSF.CORES.CINZA_CLARO } },
            border: {
              top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            },
          },
        }
      }

      // Coluna QTD PRODUÇÃO - F-G (fundo amarelo)
      const qtdCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 5 })
      ws[qtdCell] = {
        v: quantidade,
        t: "n",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AMARELO_CABECALHO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }
      ws["!merges"].push({ s: { r: currentDataRow, c: 5 }, e: { r: currentDataRow, c: 6 } })
      ws[XLSX.utils.encode_cell({ r: currentDataRow, c: 6 })] = {
        v: "",
        t: "s",
        s: {
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AMARELO_CABECALHO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }

      // Coluna PVC - H (fundo amarelo)
      const pvcCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 7 })
      ws[pvcCell] = {
        v: qtdPVC,
        t: "n",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AMARELO_CABECALHO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }

      // Coluna INOX - I (fundo amarelo)
      const inoxCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 8 })
      ws[inoxCell] = {
        v: qtdInox,
        t: "n",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AMARELO_CABECALHO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }

      // Coluna ESTOQUE BOJO PP - J-L (botão FABRICAR azul)
      const estoquePPCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 9 })
      ws[estoquePPCell] = {
        v: CONFIG_NSF.TEXTOS.FABRICAR,
        t: "s",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 9, bold: true, color: { rgb: CONFIG_NSF.CORES.BRANCO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AZUL_BOTAO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }
      ws["!merges"].push({ s: { r: currentDataRow, c: 9 }, e: { r: currentDataRow, c: 11 } })
      for (let c = 10; c <= 11; c++) {
        ws[XLSX.utils.encode_cell({ r: currentDataRow, c: c })] = {
          v: "",
          t: "s",
          s: {
            fill: { fgColor: { rgb: CONFIG_NSF.CORES.AZUL_BOTAO } },
            border: {
              top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            },
          },
        }
      }

      // Coluna ESTOQUE BOJO INOX - M-O (botão FABRICAR azul)
      const estoqueInoxCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 12 })
      ws[estoqueInoxCell] = {
        v: CONFIG_NSF.TEXTOS.FABRICAR,
        t: "s",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 9, bold: true, color: { rgb: CONFIG_NSF.CORES.BRANCO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AZUL_BOTAO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }
      ws["!merges"].push({ s: { r: currentDataRow, c: 12 }, e: { r: currentDataRow, c: 14 } })
      for (let c = 13; c <= 14; c++) {
        ws[XLSX.utils.encode_cell({ r: currentDataRow, c: c })] = {
          v: "",
          t: "s",
          s: {
            fill: { fgColor: { rgb: CONFIG_NSF.CORES.AZUL_BOTAO } },
            border: {
              top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            },
          },
        }
      }

      // Coluna PRODUZIR - P-R (botão FABRICAR azul)
      const produzirCell = XLSX.utils.encode_cell({ r: currentDataRow, c: 15 })
      ws[produzirCell] = {
        v: CONFIG_NSF.TEXTOS.FABRICAR,
        t: "s",
        s: {
          font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 9, bold: true, color: { rgb: CONFIG_NSF.CORES.BRANCO } },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: CONFIG_NSF.CORES.AZUL_BOTAO } },
          border: {
            top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
          },
        },
      }
      ws["!merges"].push({ s: { r: currentDataRow, c: 15 }, e: { r: currentDataRow, c: 17 } })
      for (let c = 16; c <= 17; c++) {
        ws[XLSX.utils.encode_cell({ r: currentDataRow, c: c })] = {
          v: "",
          t: "s",
          s: {
            fill: { fgColor: { rgb: CONFIG_NSF.CORES.AZUL_BOTAO } },
            border: {
              top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
            },
          },
        }
      }

      currentDataRow++

      // Linha vazia entre itens (apenas se não for o último)
      if (index < dadosFinais.length - 1) {
        for (let c = 0; c < 18; c++) {
          const emptyCell = XLSX.utils.encode_cell({ r: currentDataRow, c: c })
          ws[emptyCell] = {
            v: "",
            t: "s",
            s: {
              fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
              border: {
                top: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
                bottom: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
                left: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
                right: { style: "thin", color: { rgb: CONFIG_NSF.CORES.PRETO } },
              },
            },
          }
        }
        currentDataRow++
      }
    })

    // =============================================
    // LINHA DE TOTAIS
    // =============================================
    currentDataRow += 1

    // Total QTD
    ws[XLSX.utils.encode_cell({ r: currentDataRow, c: 5 })] = {
      v: totalQtd,
      t: "n",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }
    ws["!merges"].push({ s: { r: currentDataRow, c: 5 }, e: { r: currentDataRow, c: 6 } })

    // Total PVC
    ws[XLSX.utils.encode_cell({ r: currentDataRow, c: 7 })] = {
      v: totalPVC,
      t: "n",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }

    // Total INOX
    ws[XLSX.utils.encode_cell({ r: currentDataRow, c: 8 })] = {
      v: totalInox,
      t: "n",
      s: {
        font: { name: CONFIG_NSF.FONTES.PRINCIPAL, sz: 10, bold: true, color: { rgb: CONFIG_NSF.CORES.PRETO } },
        alignment: { horizontal: "center", vertical: "center" },
        fill: { fgColor: { rgb: CONFIG_NSF.CORES.BRANCO } },
      },
    }

    // =============================================
    // CONFIGURAÇÕES FINAIS DA PLANILHA
    // =============================================

    // Definir larguras das colunas
    ws["!cols"] = [
      { wch: 4 }, // A
      { wch: 4 }, // B
      { wch: 12 }, // C
      { wch: 12 }, // D
      { wch: 12 }, // E
      { wch: 8 }, // F
      { wch: 8 }, // G
      { wch: 6 }, // H
      { wch: 6 }, // I
      { wch: 10 }, // J
      { wch: 10 }, // K
      { wch: 10 }, // L
      { wch: 10 }, // M
      { wch: 10 }, // N
      { wch: 10 }, // O
      { wch: 10 }, // P
      { wch: 8 }, // Q
      { wch: 10 }, // R
    ]

    // Definir alturas das linhas
    ws["!rows"] = []
    for (let i = 0; i <= currentDataRow + 5; i++) {
      ws["!rows"][i] = { hpt: 18 }
    }
    ws["!rows"][1] = { hpt: 25 } // Título
    ws["!rows"][headerRow] = { hpt: 22 } // Cabeçalho

    // Definir o range da planilha
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: currentDataRow + 5, c: 17 } })

    XLSX.utils.book_append_sheet(wb, ws, "OP Produção")

    const dataAtual = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-")
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`, { cellStyles: true })
  } else {
    // Layout padrão para outros relatórios
    const dadosFinais = selecionarColunas(dadosParaExportar, tipo, isFilteredDetail)
    const ws = XLSX.utils.json_to_sheet(dadosFinais)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, "Relatório")

    const dataAtual = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-")
    XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`)
  }
}

// ====================================================================================
// FUNÇÕES AUXILIARES RESTANTES
// ====================================================================================

function selecionarColunas(data, tipoRelatorio, isFilteredDetail) {
  if (data.length === 0) return []

  let colunasChave
  let headersMap

  if (tipoRelatorio === "op") {
    colunasChave = ["GRUPO", "LINHA", "BOJO", "DIMENSÃO", "ALINHAMENTOS", "QUANTIDADE TOTAL"]
    headersMap = {
      GRUPO: "GRUPO",
      LINHA: "MODELO",
      BOJO: "MATERIAL (BOJO)",
      DIMENSÃO: "DIMENSÃO",
      ALINHAMENTOS: "SETOR / ALINHAMENTO",
      "QUANTIDADE TOTAL": "QTD. FABRICAR",
    }
  } else if (tipoRelatorio === "detalhe") {
    colunasChave = ["GRUPO", "LINHA", "BOJO", "ALINHAMENTOS", "DIMENSÃO", "STATUS", "QUANTIDADE TOTAL"]
    headersMap = {
      GRUPO: "GRUPO",
      LINHA: "LINHA",
      BOJO: "BOJO",
      ALINHAMENTOS: "ALINHAMENTOS",
      DIMENSÃO: "DIMENSÃO",
      STATUS: "STATUS",
      "QUANTIDADE TOTAL": "QUANTIDADE TOTAL",
    }
    if (isFilteredDetail) {
      colunasChave = colunasChave.filter((col) => col !== "ALINHAMENTOS")
    }
  } else {
    colunasChave = ["GRUPO", "LINHA", "BOJO", "ALINHAMENTOS", "STATUS", "QUANTIDADE TOTAL"]
    headersMap = {
      GRUPO: "GRUPO",
      LINHA: "LINHA",
      BOJO: "BOJO",
      ALINHAMENTOS: "ALINHAMENTOS",
      STATUS: "STATUS",
      "QUANTIDADE TOTAL": "QUANTIDADE TOTAL",
    }
  }

  return data.map((item) => {
    const novoItem = {}
    colunasChave.forEach((col) => {
      novoItem[headersMap[col] || col] = item[col]
    })
    return novoItem
  })
}

function exportarPDF(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail) {
  if (typeof window.jspdf === "undefined" || typeof window.jspdf.jsPDF === "undefined") {
    alert("Erro: Biblioteca PDF não carregada.")
    return
  }

  const dadosFinais = selecionarColunas(dadosParaExportar, tipo, isFilteredDetail)
  const { jsPDF } = window.jspdf

  const doc = new jsPDF("landscape")

  const colunas = dadosFinais.length > 0 ? Object.keys(dadosFinais[0]) : []
  const head = [colunas]

  const body = dadosFinais.map((item) => colunas.map((key) => item[key]))

  doc.autoTable({
    head: head,
    body: body,
    startY: 20,
    theme: "grid",
    styles: {
      fontSize: 8,
      font: "helvetica",
      textColor: [52, 58, 64],
    },
    headStyles: {
      fillColor: [0, 123, 255],
      textColor: 255,
      fontStyle: "bold",
      fontSize: 9,
    },
    alternateRowStyles: {
      fillColor: [248, 249, 250],
    },
    bodyStyles: {
      lineColor: [222, 226, 230],
      lineWidth: 0.1,
    },
    didDrawPage: (data) => {
      doc.setFontSize(14)
      doc.text(`Lista Otimizada de Lotes - PCP (${tipo.toUpperCase()})`, data.settings.margin.left, 10)
      doc.setFontSize(10)
      doc.text(
        `Relatório: ${nomeArquivo.replace(/_/g, " ")} | Data: ${new Date().toLocaleDateString("pt-BR")}`,
        data.settings.margin.left,
        15,
      )
    },
  })

  const dataAtual = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-")
  doc.save(`${nomeArquivo}_${dataAtual}.pdf`)
}

function exportarRelatorio(tipo, formato, filtro = null) {
  let dadosParaExportar
  let nomeArquivo
  let isFilteredDetail = false

  if (tipo === "geral") {
    dadosParaExportar = lotesGerais
    nomeArquivo = "Resumo_Geral_Lotes"
  } else if (tipo === "detalhe") {
    if (filtro) {
      dadosParaExportar = lotesDetalhes.filter((lote) => lote.ALINHAMENTOS === filtro)
      nomeArquivo = `Detalhe_Lotes_${filtro.replace(/\s/g, "_")}`
      isFilteredDetail = true
    } else {
      dadosParaExportar = lotesDetalhes
      nomeArquivo = "Detalhe_Lotes_Completo"
    }
  } else if (tipo === "op") {
    if (!filtro) {
      alert("Erro: O relatório OP deve ter um Grupo de filtro selecionado.")
      return
    }

    dadosParaExportar = lotesDetalhes.filter((lote) => lote.GRUPO === filtro)

    dadosParaExportar.sort((a, b) => {
      if (a.BOJO.localeCompare(b.BOJO) !== 0) {
        return a.BOJO.localeCompare(b.BOJO)
      }
      if (a.LINHA.localeCompare(b.LINHA) !== 0) {
        return a.LINHA.localeCompare(b.LINHA)
      }
      return a.DIMENSÃO.localeCompare(b.DIMENSÃO)
    })

    nomeArquivo = `OP_GRUPO_${filtro.replace(/\s/g, "_")}`
  }

  if (dadosParaExportar.length === 0) {
    alert("Nenhum dado encontrado para o filtro selecionado.")
    return
  }

  if (formato === "xlsx") {
    exportarXLSX(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail)
  } else if (formato === "pdf") {
    exportarPDF(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail)
  }
}

function processarPlanilha() {
  const fileInput = document.getElementById("excelFileInput")
  const statusDiv = document.getElementById("statusMessage")

  statusDiv.textContent = "Processando..."
  document.getElementById("processButton").disabled = true
  document.getElementById("downloadSection").style.display = "none"

  const file = fileInput.files[0]
  if (!file) {
    statusDiv.textContent = "Erro: Nenhum arquivo selecionado."
    document.getElementById("processButton").disabled = false
    return
  }

  const reader = new FileReader()
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result)
      const workbook = XLSX.read(data, { type: "array" })

      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      const range = XLSX.utils.decode_range(worksheet["!ref"])
      range.s.r = 4
      const newRange = XLSX.utils.encode_range(range)

      const rawDataAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: newRange })

      if (!rawDataAOA || rawDataAOA.length === 0) {
        throw new Error("A planilha parece estar vazia ou o intervalo de leitura está incorreto.")
      }

      const headers = rawDataAOA[0].map((h) =>
        String(h || "")
          .toUpperCase()
          .trim(),
      )
      const INDEX_CATEGORIA = headers.indexOf("LINHA")
      const INDEX_MODELO = headers.indexOf("ALINHAMENTO")
      const INDEX_BOJO = headers.indexOf("BOJO")
      const INDEX_DIMENSAO = headers.indexOf("DIMENSÃO")
      const INDEX_STATUS = headers.indexOf("MONTAGEM")

      if (INDEX_CATEGORIA === -1 || INDEX_MODELO === -1) {
        const headersEncontrados = headers.join(", ")
        const msg = `Erro Crítico: Colunas não encontradas. Headers: [${headersEncontrados}]`
        alert(msg)
        throw new Error(msg)
      }

      const lotesGeraisMap = {}
      const lotesDetalhesMap = {}
      todasCategorias.clear()
      todosGrupos.clear()
      let linhasProcessadas = 0

      for (let i = 1; i < rawDataAOA.length; i++) {
        const row = rawDataAOA[i]
        const maxIndex = Math.max(INDEX_CATEGORIA, INDEX_MODELO, INDEX_BOJO, INDEX_DIMENSAO)
        if (row.length <= maxIndex) continue

        const modeloOriginal = String(row[INDEX_MODELO] || "").trim()
        const dimensaoOriginal = String(row[INDEX_DIMENSAO] || "").trim()
        const bojoOriginal = String(row[INDEX_BOJO] || "").trim()
        const categoriaOriginal = String(row[INDEX_CATEGORIA] || "").trim()

        let statusOriginal = "A PRODUZIR"
        if (INDEX_STATUS !== -1) {
          const valorCelula = String(row[INDEX_STATUS] || "")
            .trim()
            .toUpperCase()
          if (valorCelula !== "") {
            statusOriginal = valorCelula
          }
        }

        const modeloLimpo = limparPrefixoModelo(modeloOriginal)
        const grupoLetra = MAPA_MODELO_GRUPO[modeloLimpo] || "N/A"
        const categoriaLimpa = limparCategoria(categoriaOriginal)
        const bojoNormalizado = (bojoOriginal || "N/A").toUpperCase()
        const dimensaoNormalizada = normalizarDimensao(dimensaoOriginal)

        if (!modeloLimpo || !categoriaLimpa) continue

        linhasProcessadas++
        todasCategorias.add(categoriaLimpa)
        todosGrupos.add(grupoLetra)

        const chaveGeral = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${grupoLetra}|${statusOriginal}`
        const chaveDetalhe = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${dimensaoNormalizada}|${grupoLetra}|${statusOriginal}`

        if (lotesGeraisMap[chaveGeral]) {
          lotesGeraisMap[chaveGeral]["QUANTIDADE TOTAL"]++
        } else {
          lotesGeraisMap[chaveGeral] = {
            GRUPO: grupoLetra,
            LINHA: modeloLimpo,
            BOJO: bojoNormalizado,
            ALINHAMENTOS: categoriaLimpa,
            STATUS: statusOriginal,
            "QUANTIDADE TOTAL": 1,
          }
        }

        if (lotesDetalhesMap[chaveDetalhe]) {
          lotesDetalhesMap[chaveDetalhe]["QUANTIDADE TOTAL"]++
        } else {
          lotesDetalhesMap[chaveDetalhe] = {
            GRUPO: grupoLetra,
            LINHA: modeloLimpo,
            BOJO: bojoNormalizado,
            ALINHAMENTOS: categoriaLimpa,
            DIMENSÃO: dimensaoNormalizada,
            STATUS: statusOriginal,
            "QUANTIDADE TOTAL": 1,
          }
        }
      }

      lotesGerais = Object.values(lotesGeraisMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA))
      lotesDetalhes = Object.values(lotesDetalhesMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA))

      if (linhasProcessadas === 0) {
        statusDiv.textContent = `Nenhuma linha de dados processada.`
        document.getElementById("processButton").disabled = false
        return
      }

      statusDiv.textContent = `Processamento concluído! Itens: ${linhasProcessadas}.`
      document.getElementById("downloadSection").style.display = "block"
      gerarBotoesFiltro()
      gerarBotoesOP()
    } catch (error) {
      console.error("Erro:", error)
      statusDiv.textContent = `Erro: ${error.message}`
    } finally {
      document.getElementById("processButton").disabled = false
    }
  }

  reader.onerror = (ex) => {
    statusDiv.textContent = "Erro ao ler o arquivo."
    document.getElementById("processButton").disabled = false
  }

  reader.readAsArrayBuffer(file)
}
