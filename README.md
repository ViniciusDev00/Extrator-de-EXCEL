# NSF | Otimizador de Lotes PCP (Extrator de EXCEL)

![Logo NSF](Logo_NSF_C.png)

O **Otimizador de Lotes PCP** é uma ferramenta industrial avançada desenvolvida para transformar planilhas brutas de pedidos em listas de produção inteligentes e organizadas. Focado na eficiência do setor de **Planejamento e Controle de Produção (PCP)**, o sistema automatiza o agrupamento de lotes por modelo, material e dimensão.

## 🚀 Funcionalidades Principais

* **Processamento Inteligente de Planilhas**: Lê arquivos `.xlsx` e `.csv`, normaliza dimensões e mapeia automaticamente os modelos para seus respectivos grupos de produção.
* **Agrupamento Estratégico**:
    * **Resumo Geral**: Visão consolidada por Modelo, Material e Categoria.
    * **Detalhe de Produção**: Lista técnica com dimensões exatas para corte e fabricação (Lista de Corte).
* **Ordem de Produção (OP) Profissional**:
    * Interface de pré-visualização idêntica ao documento físico necessário no chão de fábrica.
    * Layout otimizado para impressão (A4) com campos para assinaturas de PCP, Supervisão e Operadores.
    * Controle de estoque visual com Badges dinâmicos ("FABRICAR" ou "OK").
* **Filtros Setoriais**: Exportação de arquivos PDF ou Excel filtrados por categorias específicas.

## 🛠️ Tecnologias Utilizadas

* **Frontend**: HTML5, CSS3 (com suporte avançado a `@media print` para relatórios) e JavaScript ES6+.
* **Processamento de Dados**: [SheetJS (XLSX)](https://sheetjs.com/) para leitura e manipulação de planilhas de alta performance.
* **Geração de Documentos**: [jsPDF](https://github.com/parallax/jsPDF) para criação de documentos técnicos em PDF.
* **UI/UX**: Design responsivo e interativo focado na usabilidade do ambiente industrial.

## 📋 Como Utilizar

1.  **Acesse a Ferramenta**: Abra o arquivo `index.html` em seu navegador ou acesse a [Demo Online](https://viniciusdev00.github.io/Extrator-de-EXCEL/).
2.  **Carregue os Dados**: Selecione a planilha de pedidos da semana no "Passo 1".
3.  **Processe**: O sistema fará o mapeamento automático baseado no `MAPA_MODELO_GRUPO` (ex: identifica que modelos `VABP` pertencem ao **Grupo A**).
4.  **Visualize e Baixe**:
    * Utilize a **Opção D** para visualizar a Ordem de Produção de um grupo específico.
    * Clique em **Imprimir Ordem** para gerar o documento físico para a produção.

## 🏗️ Estrutura do Repositório

* `index.html`: Dashboard principal com opções de carga e exportação.
* `script.js`: Core do sistema com as regras de negócio do PCP e mapeamento de grupos.
* `op-preview.html` / `.js` / `.css`: Módulo dedicado à visualização e formatação profissional da Ordem de Produção.
* `styles.css`: Estilização visual moderna com o tema da NSF.

---
**Desenvolvido por Vinicius Biancolini** © 2026 NSF INDUSTRIAL - Sistema de Produção PCP v1.0.0
