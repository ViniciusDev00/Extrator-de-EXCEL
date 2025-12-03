// ====================================================================================
// SISTEMA NSF - OTIMIZADOR DE LOTES PCP
// FORMATO EXATO DO PDF - ORDEM DE PRODUÇÃO PROFISSIONAL
// ====================================================================================

/* global XLSX, jsPDF */ 

// ====================================================================================
// CONFIGURAÇÕES GLOBAIS E CONSTANTES
// ====================================================================================

const CONFIG_PDF = {
  // Configurações de página
  PAGINA: {
    ORIENTACAO: "portrait",
    UNIDADE: "mm",
    FORMATO: "a4",
    MARGEM_ESQUERDA: 10,
    MARGEM_DIREITA: 10,
    MARGEM_SUPERIOR: 10,
    MARGEM_INFERIOR: 10,
    LARGURA_UTIL: 190, // A4 width - margens
    ALTURA_UTIL: 277, // A4 height - margens
  },

  // Cores (em RGB para jsPDF)
  CORES: {
    PRETO: [0, 0, 0],
    BRANCO: [255, 255, 255],
    CINZA_TABELA: [240, 240, 240],
    CINZA_ESCURO: [100, 100, 100],
  },

  // Fontes e tamanhos
  FONTES: {
    TITULO: { tamanho: 16, estilo: "bold" },
    SUBTITULO: { tamanho: 12, estilo: "bold" },
    CABECALHO_TABELA: { tamanho: 10, estilo: "bold" },
    DADOS_TABELA: { tamanho: 10, estilo: "normal" },
    ASSINATURA: { tamanho: 9, estilo: "normal" },
    DATA: { tamanho: 10, estilo: "normal" },
    FILTRO: { tamanho: 11, estilo: "bold" },
  },

  // Layout da tabela
  TABELA: {
    ALTURA_CABECALHO: 8,
    ALTURA_LINHA: 7,
    PADDING_CELULA: 2,
    COLUNAS: [
      { nome: "Nº", largura: 10 },
      { nome: "PEDIDO: CLIENTE", largura: 60 },
      { nome: "QTD PRODUÇÃO", largura: 25 },
      { nome: "PVC", largura: 15 },
      { nome: "INOX", largura: 15 },
      { nome: "ESTOQUE PP", largura: 25 },
      { nome: "ESTOQUE INOX", largura: 25 },
      { nome: "PRODUZIR", largura: 25 },
    ],
  },
};

// ====================================================================================
// MAPA COMPLETO DE AGRUPAMENTO DE MODELOS (MESMO DO ANTERIOR)
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
};

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
};

// ====================================================================================
// VARIÁVEIS GLOBAIS DO SISTEMA
// ====================================================================================

let lotesGerais = [];
let lotesDetalhes = [];
const todasCategorias = new Set();
const todosGrupos = new Set();

// ====================================================================================
// INICIALIZAÇÃO DO SISTEMA
// ====================================================================================

document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("excelFileInput");
  const processButton = document.getElementById("processButton");
  
  fileInput.addEventListener("change", () => {
    processButton.disabled = fileInput.files.length === 0;
    document.getElementById("processButtonText").textContent = "Processar Planilha";
  });
});

// ====================================================================================
// FUNÇÕES AUXILIARES DE FORMATAÇÃO
// ====================================================================================

function getNomeCompletoGrupo(grupo) {
  return MAPA_NOME_GRUPO[grupo] || `GRUPO ${grupo}`;
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
  };
  return categorias[grupo] || "EXPOSITORES";
}

function getSemanaAtualFormatada() {
  const hoje = new Date();
  const inicioAno = new Date(hoje.getFullYear(), 0, 1);
  const dias = Math.floor((hoje - inicioAno) / (24 * 60 * 60 * 1000));
  const semana = Math.ceil((dias + 1) / 7);
  
  // Retorna no formato "SEMANA XX"
  return `SEMANA ${String(semana).padStart(2, "0")}`;
}

function getDataAtualFormatada() {
  const hoje = new Date();
  const dia = String(hoje.getDate()).padStart(2, "0");
  const mes = String(hoje.getMonth() + 1).padStart(2, "0");
  const ano = hoje.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

function formatarDimensaoParaPDF(dimensao) {
  if (!dimensao || dimensao === "N/A" || dimensao === "") return "N/A";
  
  // Remove espaços e pontos
  let dim = dimensao.toString().replace(/\s/g, "").replace(".", ",");
  
  // Garante que tenha formato decimal
  if (!dim.includes(",")) {
    dim = dim + ",000";
  }
  
  return dim;
}

// ====================================================================================
// FUNÇÃO AUXILIAR PARA CONTROLE DE PÁGINAS
// ====================================================================================

function verificarNovaPagina(doc, yAtual, alturaNecessaria) {
  if (yAtual + alturaNecessaria > CONFIG_PDF.PAGINA.ALTURA_UTIL) {
    doc.addPage();
    return CONFIG_PDF.PAGINA.MARGEM_SUPERIOR + 10;
  }
  return yAtual;
}

// ====================================================================================
// FUNÇÃO PRINCIPAL PARA GERAR PDF IDÊNTICO AO MODELO
// ====================================================================================

function gerarOrdemProducaoPDF(grupo) {
  // Verificar se a biblioteca jsPDF está carregada
  if (typeof window.jspdf === 'undefined') {
    alert("Erro: Biblioteca jsPDF não carregada. Aguarde o carregamento da página.");
    return;
  }
  
  // Filtrar dados do grupo selecionado
  const dadosGrupo = lotesDetalhes.filter((lote) => lote.GRUPO === grupo);
  
  if (dadosGrupo.length === 0) {
    alert(`Nenhum dado encontrado para o grupo ${grupo}`);
    return;
  }
  
  // Ordenar dados
  dadosGrupo.sort((a, b) => {
    if (a.BOJO !== b.BOJO) {
      return a.BOJO.localeCompare(b.BOJO);
    }
    if (a.LINHA !== b.LINHA) {
      return a.LINHA.localeCompare(b.LINHA);
    }
    if (a.DIMENSÃO !== b.DIMENSÃO) {
      return a.DIMENSÃO.localeCompare(b.DIMENSÃO);
    }
    return 0;
  });
  
  // Criar PDF
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({
    orientation: CONFIG_PDF.PAGINA.ORIENTACAO,
    unit: CONFIG_PDF.PAGINA.UNIDADE,
    format: CONFIG_PDF.PAGINA.FORMATO,
  });
  
  // =============================================
  // CABEÇALHO - TÍTULO PRINCIPAL
  // =============================================
  doc.setFontSize(CONFIG_PDF.FONTES.TITULO.tamanho);
  doc.setFont("helvetica", CONFIG_PDF.FONTES.TITULO.estilo);
  doc.text("LISTAGEM DE PRODUÇÃO", CONFIG_PDF.PAGINA.LARGURA_UTIL / 2, 20, { align: "center" });
  
  // =============================================
  // SUBTÍTULO - CATEGORIA E SEMANA
  // =============================================
  const categoria = getCategoriaGrupo(grupo);
  const semana = getSemanaAtualFormatada();
  const subtitulo = `${categoria} (${semana})`;
  
  doc.setFontSize(CONFIG_PDF.FONTES.SUBTITULO.tamanho);
  doc.setFont("helvetica", CONFIG_PDF.FONTES.SUBTITULO.estilo);
  doc.text(subtitulo, CONFIG_PDF.PAGINA.MARGEM_ESQUERDA, 35);
  
  // =============================================
  // DATA À DIREITA
  // =============================================
  const dataAtualTexto = getDataAtualFormatada();
  doc.setFontSize(CONFIG_PDF.FONTES.DATA.tamanho);
  doc.setFont("helvetica", "normal");
  doc.text(`DATA: ${dataAtualTexto}`, CONFIG_PDF.PAGINA.LARGURA_UTIL - 30, 35);
  
  // =============================================
  // FILTRO DO GRUPO (Caixa com borda)
  // =============================================
  const nomeGrupoDisplay = getNomeCompletoGrupo(grupo);
  const filtroX = CONFIG_PDF.PAGINA.MARGEM_ESQUERDA;
  const filtroY = 45;
  const filtroLargura = 50;
  const filtroAltura = 8;
  
  // Caixa do filtro
  doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
  doc.setFillColor(...CONFIG_PDF.CORES.BRANCO);
  doc.rect(filtroX, filtroY, filtroLargura, filtroAltura, "FD");
  
  // Texto do filtro
  doc.setFontSize(CONFIG_PDF.FONTES.FILTRO.tamanho);
  doc.setFont("helvetica", CONFIG_PDF.FONTES.FILTRO.estilo);
  doc.text(nomeGrupoDisplay, filtroX + 5, filtroY + 5);
  
  // =============================================
  // TABELA PRINCIPAL - CABEÇALHO
  // =============================================
  const tabelaX = CONFIG_PDF.PAGINA.MARGEM_ESQUERDA;
  const tabelaY = 60;
  
  // Desenhar cabeçalho da tabela
  let xAtual = tabelaX;
  CONFIG_PDF.TABELA.COLUNAS.forEach((coluna) => {
    // Retângulo do cabeçalho
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
    doc.rect(xAtual, tabelaY, coluna.largura, CONFIG_PDF.TABELA.ALTURA_CABECALHO, "FD");
    
    // Texto do cabeçalho
    doc.setFontSize(CONFIG_PDF.FONTES.CABECALHO_TABELA.tamanho);
    doc.setFont("helvetica", CONFIG_PDF.FONTES.CABECALHO_TABELA.estilo);
    doc.text(coluna.nome, xAtual + coluna.largura / 2, tabelaY + 5, { align: "center" });
    
    xAtual += coluna.largura;
  });
  
  // =============================================
  // DADOS DA TABELA
  // =============================================
  let yAtual = tabelaY + CONFIG_PDF.TABELA.ALTURA_CABECALHO;
  let totalQuantidade = 0;
  let totalPVC = 0;
  let totalINOX = 0;
  
  dadosGrupo.forEach((item, index) => {
    // Verificar se precisa de nova página
    yAtual = verificarNovaPagina(doc, yAtual, CONFIG_PDF.TABELA.ALTURA_LINHA);
    
    const dimensaoFormatada = formatarDimensaoParaPDF(item.DIMENSÃO);
    const pedidoCliente = `${item.LINHA} ${dimensaoFormatada}`;
    const quantidade = parseInt(item["QUANTIDADE TOTAL"]) || 0;
    const material = (item.BOJO || "").toUpperCase();
    const isPVC = material.includes("PVC") || material.includes("PP");
    const qtdPVC = isPVC ? quantidade : 0;
    const qtdINOX = !isPVC ? quantidade : 0;
    
    totalQuantidade += quantidade;
    totalPVC += qtdPVC;
    totalINOX += qtdINOX;
    
    // Desenhar linha da tabela
    xAtual = tabelaX;
    
    // Coluna 1: Nº (com "X")
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(...CONFIG_PDF.CORES.BRANCO);
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[0].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.text("X", xAtual + CONFIG_PDF.TABELA.COLUNAS[0].largura / 2, yAtual + 4.5, { align: "center" });
    xAtual += CONFIG_PDF.TABELA.COLUNAS[0].largura;
    
    // Coluna 2: PEDIDO: CLIENTE
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(...CONFIG_PDF.CORES.BRANCO);
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[1].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "normal");
    
    // Quebrar texto se necessário
    const textoPedido = pedidoCliente.length > 30 ? pedidoCliente.substring(0, 30) + "..." : pedidoCliente;
    doc.text(textoPedido, xAtual + 2, yAtual + 4.5);
    xAtual += CONFIG_PDF.TABELA.COLUNAS[1].largura;
    
    // Coluna 3: QTD PRODUÇÃO
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[2].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.text(quantidade.toString(), xAtual + CONFIG_PDF.TABELA.COLUNAS[2].largura / 2, yAtual + 4.5, { align: "center" });
    xAtual += CONFIG_PDF.TABELA.COLUNAS[2].largura;
    
    // Coluna 4: PVC
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[3].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.text(qtdPVC.toString(), xAtual + CONFIG_PDF.TABELA.COLUNAS[3].largura / 2, yAtual + 4.5, { align: "center" });
    xAtual += CONFIG_PDF.TABELA.COLUNAS[3].largura;
    
    // Coluna 5: INOX
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[4].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.text(qtdINOX.toString(), xAtual + CONFIG_PDF.TABELA.COLUNAS[4].largura / 2, yAtual + 4.5, { align: "center" });
    xAtual += CONFIG_PDF.TABELA.COLUNAS[4].largura;
    
    // Coluna 6: ESTOQUE PP (botão FABRICAR)
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(qtdPVC > 0 ? [52, 152, 219] : [200, 200, 200]); // Azul se tiver PVC, cinza se não
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[5].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...CONFIG_PDF.CORES.BRANCO);
    doc.text(qtdPVC > 0 ? "FABRICAR" : "-", xAtual + CONFIG_PDF.TABELA.COLUNAS[5].largura / 2, yAtual + 4.5, { align: "center" });
    doc.setTextColor(...CONFIG_PDF.CORES.PRETO);
    xAtual += CONFIG_PDF.TABELA.COLUNAS[5].largura;
    
    // Coluna 7: ESTOQUE INOX (botão FABRICAR)
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor(qtdINOX > 0 ? [52, 152, 219] : [200, 200, 200]); // Azul se tiver INOX, cinza se não
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[6].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...CONFIG_PDF.CORES.BRANCO);
    doc.text(qtdINOX > 0 ? "FABRICAR" : "-", xAtual + CONFIG_PDF.TABELA.COLUNAS[6].largura / 2, yAtual + 4.5, { align: "center" });
    doc.setTextColor(...CONFIG_PDF.CORES.PRETO);
    xAtual += CONFIG_PDF.TABELA.COLUNAS[6].largura;
    
    // Coluna 8: PRODUZIR (botão FABRICAR)
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.setFillColor([52, 152, 219]); // Azul sempre
    doc.rect(xAtual, yAtual, CONFIG_PDF.TABELA.COLUNAS[7].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
    doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...CONFIG_PDF.CORES.BRANCO);
    doc.text("FABRICAR", xAtual + CONFIG_PDF.TABELA.COLUNAS[7].largura / 2, yAtual + 4.5, { align: "center" });
    doc.setTextColor(...CONFIG_PDF.CORES.PRETO);
    
    yAtual += CONFIG_PDF.TABELA.ALTURA_LINHA;
  });
  
  // =============================================
  // LINHA DE TOTAIS
  // =============================================
  yAtual += 5;
  
  // Fundo para totais
  doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
  doc.setFillColor(...CONFIG_PDF.CORES.BRANCO);
  doc.rect(tabelaX, yAtual, 
    CONFIG_PDF.TABELA.COLUNAS[0].largura + CONFIG_PDF.TABELA.COLUNAS[1].largura, 
    CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
  
  // Texto "TOTAIS"
  doc.setFontSize(CONFIG_PDF.FONTES.DADOS_TABELA.tamanho);
  doc.setFont("helvetica", "bold");
  doc.text("TOTAIS", tabelaX + 5, yAtual + 4.5);
  
  // Total QTD PRODUÇÃO
  let xTotal = tabelaX + CONFIG_PDF.TABELA.COLUNAS[0].largura + CONFIG_PDF.TABELA.COLUNAS[1].largura;
  doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
  doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
  doc.rect(xTotal, yAtual, CONFIG_PDF.TABELA.COLUNAS[2].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
  doc.text(totalQuantidade.toString(), xTotal + CONFIG_PDF.TABELA.COLUNAS[2].largura / 2, yAtual + 4.5, { align: "center" });
  xTotal += CONFIG_PDF.TABELA.COLUNAS[2].largura;
  
  // Total PVC
  doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
  doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
  doc.rect(xTotal, yAtual, CONFIG_PDF.TABELA.COLUNAS[3].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
  doc.text(totalPVC.toString(), xTotal + CONFIG_PDF.TABELA.COLUNAS[3].largura / 2, yAtual + 4.5, { align: "center" });
  xTotal += CONFIG_PDF.TABELA.COLUNAS[3].largura;
  
  // Total INOX
  doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
  doc.setFillColor(...CONFIG_PDF.CORES.CINZA_TABELA);
  doc.rect(xTotal, yAtual, CONFIG_PDF.TABELA.COLUNAS[4].largura, CONFIG_PDF.TABELA.ALTURA_LINHA, "FD");
  doc.text(totalINOX.toString(), xTotal + CONFIG_PDF.TABELA.COLUNAS[4].largura / 2, yAtual + 4.5, { align: "center" });
  
  // =============================================
  // ÁREA DE ASSINATURAS
  // =============================================
  yAtual += 15;
  
  // Linha horizontal
  doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
  doc.line(tabelaX, yAtual, tabelaX + 180, yAtual);
  
  yAtual += 10;
  
  // Assinaturas
  const assinaturas = [
    { cargo: "PCP", x: tabelaX + 20 },
    { cargo: "SUPERVISOR DE PRODUÇÃO", x: tabelaX + 70 },
    { cargo: "OPERADOR 01", x: tabelaX + 120 },
    { cargo: "OPERADOR 02", x: tabelaX + 170 },
  ];
  
  assinaturas.forEach((assinatura) => {
    // Linha para assinatura
    doc.setDrawColor(...CONFIG_PDF.CORES.PRETO);
    doc.line(assinatura.x, yAtual, assinatura.x + 40, yAtual);
    
    // Cargo abaixo da linha
    doc.setFontSize(CONFIG_PDF.FONTES.ASSINATURA.tamanho);
    doc.setFont("helvetica", "normal");
    doc.text(assinatura.cargo, assinatura.x + 20, yAtual + 5, { align: "center" });
  });
  
  // =============================================
  // SALVAR PDF
  // =============================================
  const dataAtualParaNome = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-");
  const nomeArquivo = `OP_${grupo}_${nomeGrupoDisplay.replace(/\s/g, "_")}_${dataAtualParaNome}.pdf`;
  doc.save(nomeArquivo);
}

// ====================================================================================
// FUNÇÕES AUXILIARES RESTANTES
// ====================================================================================

function limparPrefixoModelo(modeloOriginal) {
  if (typeof modeloOriginal !== "string") return String(modeloOriginal).trim().toUpperCase();
  let modelo = modeloOriginal.trim();
  const indexPrimeiroEspaco = modelo.indexOf(" ");
  if (indexPrimeiroEspaco !== -1) {
    modelo = modelo.substring(indexPrimeiroEspaco + 1).trim();
  }
  modelo = modelo.replace(/\s\s+/g, " ");
  return modelo.toUpperCase();
}

function limparCategoria(nomeOriginal) {
  if (typeof nomeOriginal !== "string") return String(nomeOriginal).trim().toUpperCase();
  let nome = nomeOriginal.trim().toUpperCase();
  nome = nome.replace(/ DE /g, " ").trim();
  nome = nome.replace(/\s\s+/g, " ");
  return nome;
}

function normalizarDimensao(dimensao) {
  if (!dimensao || dimensao === "N/A") return "N/A";
  return dimensao.toString().replace(/\s/g, "").replace(",", ".");
}

function gerarBotoesFiltro() {
  const containerXLSX = document.getElementById("filterColXLSX");
  const containerPDF = document.getElementById("filterColPDF");

  containerXLSX.innerHTML = "<h4>Download XLSX</h4>";
  containerPDF.innerHTML = "<h4>Download PDF</h4>";

  Array.from(todasCategorias)
    .sort()
    .forEach((categoria) => {
      if (!categoria || categoria === "N/A") return;

      containerXLSX.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'xlsx', '${categoria}')" class="btn">${categoria}</button>`;
      containerPDF.innerHTML += `<button onclick="exportarRelatorio('detalhe', 'pdf', '${categoria}')" class="btn">${categoria}</button>`;
    });
}

// ====================================================================================
// FUNÇÃO PARA REDIRECIONAR PARA PRÉ-VISUALIZAÇÃO
// ====================================================================================

function visualizarOrdemProducao(grupo) {
  // Filtrar dados do grupo selecionado
  const dadosGrupo = lotesDetalhes.filter((lote) => lote.GRUPO === grupo);
  
  if (dadosGrupo.length === 0) {
    alert(`Nenhum dado encontrado para o grupo ${grupo}`);
    return;
  }
  
  // Salvar dados no localStorage para a próxima página
  localStorage.setItem('opGrupoDados', JSON.stringify(dadosGrupo));
  localStorage.setItem('opGrupoSelecionado', grupo);
  localStorage.setItem('lotesDetalhes', JSON.stringify(lotesDetalhes));
  
  // Redirecionar para a página de pré-visualização
  window.location.href = 'op-preview.html';
}

// ====================================================================================
// ATUALIZAR FUNÇÃO gerarBotoesOP()
// ====================================================================================

function gerarBotoesOP() {
  // MUDANÇA IMPORTANTE: O container agora é "filterColOP" (sem "PDF" no final)
  const containerOP = document.getElementById("filterColOP");

  // MUDANÇA NO HTML: Texto atualizado e estilo diferente
  containerOP.innerHTML = "<h4 style='color: #007bff;'>Pré-visualizar Ordem de Produção</h4>";

  Array.from(todosGrupos)
    .sort()
    .forEach((grupo) => {
      if (!grupo || grupo === "N/A") return;

      const nomeCompleto = getNomeCompletoGrupo(grupo);
      // MUDANÇA IMPORTANTE: Agora chama visualizarOrdemProducao() em vez de gerarOrdemProducaoPDF()
      containerOP.innerHTML += `
        <button onclick="visualizarOrdemProducao('${grupo}')" class="btn" style="background-color: #e3f2fd; color: #1976d2; border: 1px solid #bbdefb;">
          ${grupo} - ${nomeCompleto}
        </button>
      `;
    });
}

function exportarRelatorio(tipo, formato, filtro = null) {
  let dadosParaExportar;
  let nomeArquivo;
  let isFilteredDetail = false;

  if (tipo === "geral") {
    dadosParaExportar = lotesGerais;
    nomeArquivo = "Resumo_Geral_Lotes";
  } else if (tipo === "detalhe") {
    if (filtro) {
      dadosParaExportar = lotesDetalhes.filter((lote) => lote.ALINHAMENTOS === filtro);
      nomeArquivo = `Detalhe_Lotes_${filtro.replace(/\s/g, "_")}`;
      isFilteredDetail = true;
    } else {
      dadosParaExportar = lotesDetalhes;
      nomeArquivo = "Detalhe_Lotes_Completo";
    }
  }

  if (dadosParaExportar.length === 0) {
    alert("Nenhum dado encontrado para o filtro selecionado.");
    return;
  }

  if (formato === "xlsx") {
    exportarXLSX(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail);
  } else if (formato === "pdf") {
    exportarPDF(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail);
  }
}

function exportarXLSX(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail) {
  const dadosFinais = selecionarColunas(dadosParaExportar, tipo, isFilteredDetail);
  const ws = XLSX.utils.json_to_sheet(dadosFinais);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Relatório");

  const dataAtual = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-");
  XLSX.writeFile(wb, `${nomeArquivo}_${dataAtual}.xlsx`);
}

function exportarPDF(dadosParaExportar, nomeArquivo, tipo, isFilteredDetail) {
  if (typeof window.jspdf === "undefined" || typeof window.jspdf.jsPDF === "undefined") {
    alert("Erro: Biblioteca PDF não carregada.");
    return;
  }

  const dadosFinais = selecionarColunas(dadosParaExportar, tipo, isFilteredDetail);
  const { jsPDF } = window.jspdf;

  const doc = new jsPDF("landscape");

  const colunas = dadosFinais.length > 0 ? Object.keys(dadosFinais[0]) : [];
  const head = [colunas];

  const body = dadosFinais.map((item) => colunas.map((key) => item[key]));

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
      doc.setFontSize(14);
      doc.text(`Lista Otimizada de Lotes - PCP (${tipo.toUpperCase()})`, data.settings.margin.left, 10);
      doc.setFontSize(10);
      doc.text(
        `Relatório: ${nomeArquivo.replace(/_/g, " ")} | Data: ${new Date().toLocaleDateString("pt-BR")}`,
        data.settings.margin.left,
        15,
      );
    },
  });

  const dataAtual = new Date().toLocaleDateString("pt-BR").replace(/\//g, "-");
  doc.save(`${nomeArquivo}_${dataAtual}.pdf`);
}

function selecionarColunas(data, tipoRelatorio, isFilteredDetail) {
  if (data.length === 0) return [];

  let colunasChave;
  let headersMap;

  if (tipoRelatorio === "detalhe") {
    colunasChave = ["GRUPO", "LINHA", "BOJO", "ALINHAMENTOS", "DIMENSÃO", "STATUS", "QUANTIDADE TOTAL"];
    headersMap = {
      GRUPO: "GRUPO",
      LINHA: "LINHA",
      BOJO: "BOJO",
      ALINHAMENTOS: "ALINHAMENTOS",
      DIMENSÃO: "DIMENSÃO",
      STATUS: "STATUS",
      "QUANTIDADE TOTAL": "QUANTIDADE TOTAL",
    };
    if (isFilteredDetail) {
      colunasChave = colunasChave.filter((col) => col !== "ALINHAMENTOS");
    }
  } else {
    colunasChave = ["GRUPO", "LINHA", "BOJO", "ALINHAMENTOS", "STATUS", "QUANTIDADE TOTAL"];
    headersMap = {
      GRUPO: "GRUPO",
      LINHA: "LINHA",
      BOJO: "BOJO",
      ALINHAMENTOS: "ALINHAMENTOS",
      STATUS: "STATUS",
      "QUANTIDADE TOTAL": "QUANTIDADE TOTAL",
    };
  }

  return data.map((item) => {
    const novoItem = {};
    colunasChave.forEach((col) => {
      novoItem[headersMap[col] || col] = item[col];
    });
    return novoItem;
  });
}

function processarPlanilha() {
  const fileInput = document.getElementById("excelFileInput");
  const statusDiv = document.getElementById("statusMessage");
  const processButton = document.getElementById("processButton");

  statusDiv.textContent = "⏳ Processando...";
  statusDiv.style.color = "#007bff";
  processButton.disabled = true;
  document.getElementById("downloadSection").style.display = "none";

  const file = fileInput.files[0];
  if (!file) {
    statusDiv.textContent = "Erro: Nenhum arquivo selecionado.";
    statusDiv.style.color = "red";
    processButton.disabled = false;
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      range.s.r = 4;
      const newRange = XLSX.utils.encode_range(range);

      const rawDataAOA = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: newRange });

      if (!rawDataAOA || rawDataAOA.length === 0) {
        throw new Error("A planilha parece estar vazia ou o intervalo de leitura está incorreto.");
      }

      const headers = rawDataAOA[0].map((h) =>
        String(h || "")
          .toUpperCase()
          .trim(),
      );
      const INDEX_CATEGORIA = headers.indexOf("LINHA");
      const INDEX_MODELO = headers.indexOf("ALINHAMENTO");
      const INDEX_BOJO = headers.indexOf("BOJO");
      const INDEX_DIMENSAO = headers.indexOf("DIMENSÃO");
      const INDEX_STATUS = headers.indexOf("MONTAGEM");

      if (INDEX_CATEGORIA === -1 || INDEX_MODELO === -1) {
        const headersEncontrados = headers.join(", ");
        const msg = `Erro Crítico: Colunas não encontradas. Headers: [${headersEncontrados}]`;
        alert(msg);
        throw new Error(msg);
      }

      const lotesGeraisMap = {};
      const lotesDetalhesMap = {};
      todasCategorias.clear();
      todosGrupos.clear();
      let linhasProcessadas = 0;

      for (let i = 1; i < rawDataAOA.length; i++) {
        const row = rawDataAOA[i];
        const maxIndex = Math.max(INDEX_CATEGORIA, INDEX_MODELO, INDEX_BOJO, INDEX_DIMENSAO);
        if (row.length <= maxIndex) continue;

        const modeloOriginal = String(row[INDEX_MODELO] || "").trim();
        const dimensaoOriginal = String(row[INDEX_DIMENSAO] || "").trim();
        const bojoOriginal = String(row[INDEX_BOJO] || "").trim();
        const categoriaOriginal = String(row[INDEX_CATEGORIA] || "").trim();

        let statusOriginal = "A PRODUZIR";
        if (INDEX_STATUS !== -1) {
          const valorCelula = String(row[INDEX_STATUS] || "")
            .trim()
            .toUpperCase();
          if (valorCelula !== "") {
            statusOriginal = valorCelula;
          }
        }

        const modeloLimpo = limparPrefixoModelo(modeloOriginal);
        const grupoLetra = MAPA_MODELO_GRUPO[modeloLimpo] || "N/A";
        const categoriaLimpa = limparCategoria(categoriaOriginal);
        const bojoNormalizado = (bojoOriginal || "N/A").toUpperCase();
        const dimensaoNormalizada = normalizarDimensao(dimensaoOriginal);

        if (!modeloLimpo || !categoriaLimpa) continue;

        linhasProcessadas++;
        todasCategorias.add(categoriaLimpa);
        todosGrupos.add(grupoLetra);

        const chaveGeral = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${grupoLetra}|${statusOriginal}`;
        const chaveDetalhe = `${modeloLimpo}|${bojoNormalizado}|${categoriaLimpa}|${dimensaoNormalizada}|${grupoLetra}|${statusOriginal}`;

        if (lotesGeraisMap[chaveGeral]) {
          lotesGeraisMap[chaveGeral]["QUANTIDADE TOTAL"]++;
        } else {
          lotesGeraisMap[chaveGeral] = {
            GRUPO: grupoLetra,
            LINHA: modeloLimpo,
            BOJO: bojoNormalizado,
            ALINHAMENTOS: categoriaLimpa,
            STATUS: statusOriginal,
            "QUANTIDADE TOTAL": 1,
          };
        }

        if (lotesDetalhesMap[chaveDetalhe]) {
          lotesDetalhesMap[chaveDetalhe]["QUANTIDADE TOTAL"]++;
        } else {
          lotesDetalhesMap[chaveDetalhe] = {
            GRUPO: grupoLetra,
            LINHA: modeloLimpo,
            BOJO: bojoNormalizado,
            ALINHAMENTOS: categoriaLimpa,
            DIMENSÃO: dimensaoNormalizada,
            STATUS: statusOriginal,
            "QUANTIDADE TOTAL": 1,
          };
        }
      }

      lotesGerais = Object.values(lotesGeraisMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));
      lotesDetalhes = Object.values(lotesDetalhesMap).sort((a, b) => a.LINHA.localeCompare(b.LINHA));

      if (linhasProcessadas === 0) {
        statusDiv.textContent = `Nenhuma linha de dados processada.`;
        statusDiv.style.color = "orange";
        processButton.disabled = false;
        return;
      }

      statusDiv.textContent = `✅ Processamento concluído! Itens processados: ${linhasProcessadas}.`;
      statusDiv.style.color = "green";
      document.getElementById("downloadSection").style.display = "block";
      gerarBotoesFiltro();
      gerarBotoesOP();
    } catch (error) {
      console.error("Erro:", error);
      statusDiv.textContent = `❌ Erro: ${error.message}`;
      statusDiv.style.color = "red";
    } finally {
      processButton.disabled = false;
    }
  };

  reader.onerror = (ex) => {
    statusDiv.textContent = `❌ Erro ao ler o arquivo: ${ex.type}`;
    statusDiv.style.color = "red";
    processButton.disabled = false;
  };

  reader.readAsArrayBuffer(file);
}