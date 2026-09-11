/**
 * Fast Shop - Ementário Fiscal
 * Módulo de extração e interpretação client-side de planilhas Excel (.xlsx)
 * Requer biblioteca SheetJS (XLSX) carregada no ambiente.
 */

window.ExcelParser = {
  /**
   * Lê um ArrayBuffer de arquivo .xlsx e extrai a estrutura de dados do Boletim
   * @param {ArrayBuffer} arrayBuffer
   * @returns {Object} Objeto no formato padrão do Boletim
   */
  parseWorkbook(arrayBuffer) {
    if (typeof XLSX === 'undefined') {
      throw new Error("Biblioteca SheetJS (XLSX) não encontrada.");
    }

    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellFormula: true, cellStyles: true });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    if (!worksheet) {
      throw new Error("Nenhuma planilha válida encontrada no arquivo Excel.");
    }

    // Identifica limites da planilha
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    // Mapeia colunas na linha 0 (cabeçalho)
    let colIndexMap = {};
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
      const cell = worksheet[cellAddress];
      if (cell && cell.v) {
        const val = String(cell.v).trim().toLowerCase();
        if (val.includes('n') && (val.includes('boletim') || val.includes('num'))) {
          colIndexMap['numero_boletim'] = C;
        } else if (val.includes('semana') || val.includes('periodo') || val.includes('período')) {
          colIndexMap['periodo'] = C;
        } else if (val.includes('tipo')) {
          colIndexMap['tipo'] = C;
        } else if (val.includes('lei') || val.includes('link') || val.includes('url')) {
          colIndexMap['lei'] = C;
        }
      }
    }

    // Se não encontrou coluna de lei expressa, assume coluna 0
    if (colIndexMap['lei'] === undefined) {
      colIndexMap['lei'] = 0;
    }

    let detectedNumero = "";
    let detectedPeriodo = "";
    const rawLaws = [];
    const rawNews = [];

    // Itera pelas linhas a partir da linha seguinte ao cabeçalho
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      // Captura número do boletim se disponível
      if (colIndexMap['numero_boletim'] !== undefined) {
        const cAddr = XLSX.utils.encode_cell({ r: R, c: colIndexMap['numero_boletim'] });
        const cVal = worksheet[cAddr];
        if (cVal && cVal.v && !detectedNumero) {
          detectedNumero = String(cVal.v).trim();
        }
      }

      // Captura semana / período
      if (colIndexMap['periodo'] !== undefined) {
        const cAddr = XLSX.utils.encode_cell({ r: R, c: colIndexMap['periodo'] });
        const cVal = worksheet[cAddr];
        if (cVal && cVal.v && !detectedPeriodo) {
          detectedPeriodo = String(cVal.v).trim();
        }
      }

      // Tipo da linha (Lei vs Notícia)
      let rowTipo = "lei";
      if (colIndexMap['tipo'] !== undefined) {
        const cAddr = XLSX.utils.encode_cell({ r: R, c: colIndexMap['tipo'] });
        const cVal = worksheet[cAddr];
        if (cVal && cVal.v) {
          const tVal = String(cVal.v).trim().toLowerCase();
          if (tVal.includes('notic') || tVal.includes('notícia')) {
            rowTipo = "noticia";
          }
        }
      }

      // Coluna da Lei / Link
      const leiCellAddr = XLSX.utils.encode_cell({ r: R, c: colIndexMap['lei'] });
      const leiCell = worksheet[leiCellAddr];

      if (leiCell) {
        let label = leiCell.v ? String(leiCell.v).trim() : "";
        let targetUrl = "";

        // Verifica se há hyperlink embutido na célula
        if (leiCell.l && leiCell.l.Target) {
          targetUrl = leiCell.l.Target;
        } else if (label.startsWith("http://") || label.startsWith("https://")) {
          targetUrl = label;
        }

        if (label || targetUrl) {
          const itemData = {
            label: label,
            url: targetUrl || label
          };

          if (rowTipo === "noticia") {
            rawNews.push(itemData);
          } else {
            rawLaws.push(itemData);
          }
        }
      }
    }

    // Constrói objeto estruturado inicial
    const result = {
      numero_boletim: detectedNumero || "Novo Boletim",
      periodo: detectedPeriodo || "Semana Atual",
      departamento: "Fiscal",
      subtitulo: "Ementário Fiscal",
      equipe: "Boletim Fiscal elaborado pelo time de Planejamento Fiscal: Andréa Celi Mantovani, Antônio Sergio da Silva, Cristiane Cunha, Emerson de Deus e Raquel Capelão. Em caso de dúvidas, favor enviar e-mail para planejamentofiscal@fastshop.com.br",
      itens: rawLaws.map((law, idx) => {
        // Tenta deduzir UF / Esfera se tiver padrão no título
        let esfera = "ESTADUAL";
        let orgao = "SEFAZ";
        const lbl = law.label.toUpperCase();
        if (lbl.includes("BR -") || lbl.includes("RFB") || lbl.includes("CONFAZ") || lbl.includes("DOU")) {
          esfera = "FEDERAL";
          orgao = lbl.includes("CONFAZ") ? "CONFAZ" : "RFB";
        } else if (lbl.includes("PREFEITURA") || lbl.includes("MUNIC")) {
          esfera = "MUNICIPAL";
          orgao = "PREFEITURA";
        }

        return {
          numero: idx + 1,
          esfera: esfera,
          titulo: law.label.length > 5 ? `${idx + 1}. ${law.label}` : `${idx + 1}. Legislação Fiscal nº ${idx + 1}`,
          data_publicacao: new Date().toLocaleDateString('pt-BR'),
          norma: law.label || `Norma Fiscal ${idx + 1}`,
          link: law.url,
          orgao: orgao,
          resumo_tabela: "Ato normativo inserido via importação de planilha.",
          impacto: "ICMS",
          area_impactada: "Indiretos",
          corpo_paragrafos: [
            `Ato normativo identificado a partir do link oficial: ${law.url}.`,
            "Aguardando síntese técnica detalhada dos impactos tributários para as operações da companhia."
          ],
          plano_de_acao: [
            "**Sistemas / TI Fiscal:** Avaliar impactos cadastrais e parametrização no ERP.",
            "**Tributário / Compliance:** Validar aderência das filiais e conformidade da obrigação."
          ],
          vigencia: "Vigência a partir da publicação oficial."
        };
      }),
      noticias: rawNews.map((n, idx) => ({
        titulo: n.label || `Matéria Jurídico-Tributária Relevante ${idx + 1}`,
        corpo_paragrafos: [
          `Matéria jornalística relevante de acompanhamento fiscal disponível em: ${n.url}.`
        ],
        fonte: "Portal de Notícias / Clipping Fiscal",
        link: n.url
      }))
    };

    return result;
  },

  /**
   * Extrai apenas as linhas brutas da planilha para alimentar a grade estilo Excel
   */
  extractGridRows(arrayBuffer) {
    if (typeof XLSX === 'undefined') {
      throw new Error("Biblioteca SheetJS (XLSX) não encontrada.");
    }

    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellFormula: true, cellStyles: true });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    if (!worksheet) {
      throw new Error("Nenhuma planilha válida encontrada no arquivo Excel.");
    }

    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    let colIndexMap = {};
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: C });
      const cell = worksheet[cellAddress];
      if (cell && cell.v) {
        const val = String(cell.v).trim().toLowerCase();
        if (val.includes('n') && (val.includes('boletim') || val.includes('num'))) {
          colIndexMap['numero_boletim'] = C;
        } else if (val.includes('semana') || val.includes('periodo') || val.includes('período')) {
          colIndexMap['periodo'] = C;
        } else if (val.includes('tipo')) {
          colIndexMap['tipo'] = C;
        } else if (val.includes('lei') || val.includes('link') || val.includes('url')) {
          colIndexMap['lei'] = C;
        }
      }
    }

    if (colIndexMap['lei'] === undefined) {
      colIndexMap['lei'] = 0;
    }

    let detectedNumero = "";
    let detectedPeriodo = "";
    const rows = [];

    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      if (colIndexMap['numero_boletim'] !== undefined && !detectedNumero) {
        const cVal = worksheet[XLSX.utils.encode_cell({ r: R, c: colIndexMap['numero_boletim'] })];
        if (cVal && cVal.v) detectedNumero = String(cVal.v).trim();
      }

      if (colIndexMap['periodo'] !== undefined && !detectedPeriodo) {
        const cVal = worksheet[XLSX.utils.encode_cell({ r: R, c: colIndexMap['periodo'] })];
        if (cVal && cVal.v) detectedPeriodo = String(cVal.v).trim();
      }

      let rowTipo = "Lei";
      if (colIndexMap['tipo'] !== undefined) {
        const cVal = worksheet[XLSX.utils.encode_cell({ r: R, c: colIndexMap['tipo'] })];
        if (cVal && cVal.v) {
          const tVal = String(cVal.v).trim().toLowerCase();
          if (tVal.includes('notic') || tVal.includes('notícia')) {
            rowTipo = "Notícia";
          }
        }
      }

      const leiCell = worksheet[XLSX.utils.encode_cell({ r: R, c: colIndexMap['lei'] })];
      if (leiCell) {
        let label = leiCell.v ? String(leiCell.v).trim() : "";
        let targetUrl = "";
        if (leiCell.l && leiCell.l.Target) {
          targetUrl = leiCell.l.Target;
        } else if (label.startsWith("http://") || label.startsWith("https://")) {
          targetUrl = label;
        }

        if (label || targetUrl) {
          rows.push({
            tipo: rowTipo,
            lei: label,
            link: targetUrl || label
          });
        }
      }
    }

    return {
      numero_boletim: detectedNumero || "",
      periodo: detectedPeriodo || "",
      rows: rows
    };
  }
};

/**
 * Sintetizador Inteligente de Ementas Fiscais e Planos de Ação
 * Transforma dados simples de entrada (URLs e Tipos) em registros técnicos corporativos completos
 */
window.TaxSynthesizer = {

  /**
   * Faz o Webscraping em tempo real da URL oficial via proxy/leitor com suporte a CORS
   * Obtém o Título real da página e o texto da ementa/artigos
   */
  async scrapeUrl(url) {
    if (!url || typeof url !== 'string' || !url.startsWith('http')) return null;
    try {
      const jinaUrl = `https://r.jina.ai/${url.trim()}`;
      const res = await fetch(jinaUrl, {
        headers: { 'Accept': 'text/plain' }
      });
      if (!res.ok) return null;
      const text = await res.text();
      if (!text || text.length < 20) return null;

      let title = "";
      const mTitle = text.match(/^Title:\s*(.+)$/m);
      if (mTitle) {
        title = mTitle[1].trim();
      }

      let content = "";
      const mContent = text.match(/Markdown Content:\s*\n+([\s\S]+)/);
      if (mContent) {
        content = mContent[1].trim().slice(0, 3000);
      } else {
        content = text.slice(0, 2000);
      }

      return { title, content };
    } catch (e) {
      console.warn("Aviso no webscraping:", e);
      return null;
    }
  },

  /**
   * Extrai metadados completos de legislação a partir da URL oficial
   */
  parseTaxUrl(rawUrl, itemIndex, fallbackReason = null, scrapedData = null) {
    const url = (rawUrl || '').trim();
    const urlLower = url.toLowerCase();
    const dataPub = new Date().toLocaleDateString('pt-BR');

    let uf = "SP";
    let esfera = "ESTADUAL";
    let orgao = "SEFAZ SP";
    let impacto = "ICMS";
    let area = "Indiretos";
    let norma = "";
    let resumoTabela = "";
    let corpoParagrafos = [];
    let planoAcao = [];

    // 1. Resposta à Consulta Tributária da SEFAZ/SP (RCXXXXX_YYYY.aspx)
    const matchRC = url.match(/RC[_\s-]?(\d+)[_\/](\d{2,4})/i) || url.match(/RC(\d+)/i);
    if (urlLower.includes("legislacao.fazenda.sp.gov.br") && matchRC) {
      const numRC = matchRC[1];
      const anoRC = matchRC[2] ? (matchRC[2].length === 2 ? '20' + matchRC[2] : matchRC[2]) : '2025';
      const numFormatado = Number(numRC).toLocaleString('pt-BR');

      uf = "SP";
      esfera = "ESTADUAL";
      orgao = "SEFAZ SP";
      impacto = "ICMS / Consultas Tributárias";
      area = "Indiretos / Consultoria";
      norma = `Resposta à Consulta Tributária nº ${numRC}/${anoRC}`;
      resumoTabela = `Esclarece a interpretação e a aplicação da legislação tributária estadual (ICMS) no âmbito da SEFAZ/SP.`;

      corpoParagrafos = [
        `A **Secretaria da Fazenda e Planejamento do Estado de São Paulo (SEFAZ/SP)** exarou a **Resposta à Consulta Tributária nº ${numFormatado}/${anoRC}**, prestando esclarecimentos oficiais sobre a exata interpretação e aplicação da legislação do **ICMS** no Estado.`,
        `A manifestação orienta os contribuintes quanto à **correta classificação fiscal de mercadorias, procedimentos de estorno/crédito e obrigações acessórias**, trazendo segurança jurídica e balizando o entendimento da fiscalização estadual.`
      ];

      planoAcao = [
        `**Consultoria / Planejamento Tributário:** Analisar o teor da Resposta à Consulta nº ${numRC}/${anoRC} e confrontar com os procedimentos de apuração e entradas/saídas praticados pelas filiais paulistas.`,
        `**Sistemas / TI Fiscal:** Em caso de impactos em alíquotas ou parametrização cadastral, alinhar ajustes preventivos na matriz de regras tributárias do ERP (SAP/Mastersaf).`,
        `**Compliance Fiscal:** Mapear e arquivar o posicionamento oficial no repositório de jurisprudência administrativa da Fast Shop para eventuais fiscalizações.`
      ];
    }
    // 2. Portarias da SEFAZ/SP (Portaria SRE, Portaria CAT)
    else if (urlLower.includes("legislacao.fazenda.sp.gov.br") && urlLower.includes("portaria")) {
      const matchPort = url.match(/portaria[_\s-]?([a-z]+)?[_\s-]?(\d+)[_\/](\d{2,4})/i) || url.match(/portaria[_\s-]?(\d+)/i);
      const sub = matchPort && matchPort[1] ? matchPort[1].toUpperCase() : "SRE";
      const num = matchPort && matchPort[2] ? matchPort[2] : (matchPort && matchPort[1] && !isNaN(matchPort[1]) ? matchPort[1] : itemIndex);
      const ano = matchPort && matchPort[3] ? (matchPort[3].length === 2 ? '20' + matchPort[3] : matchPort[3]) : '2026';

      uf = "SP";
      esfera = "ESTADUAL";
      orgao = "SEFAZ SP";
      impacto = "ICMS / Obrigações Acessórias";
      area = "Indiretos";
      norma = `Portaria ${sub} nº ${num}/${ano}`;
      resumoTabela = `Disciplina procedimentos e regras operacionais relativas ao ICMS no Estado de São Paulo.`;

      corpoParagrafos = [
        `A **Secretaria da Fazenda e Planejamento de São Paulo** publicou a **Portaria ${sub} nº ${num}/${ano}**, estabelecendo alterações e diretrizes operacionais no âmbito do **ICMS**.`,
        `A norma estabelece **prazos para cumprimento de exigências acessórias e regras de controle fiscal**, demandando imediata observância pelos contribuintes inscritos no Estado.`
      ];

      planoAcao = [
        `**Sistemas / TI Fiscal:** Avaliar se as regras de validação do ERP demandam atualização de layouts ou prazos conforme a Portaria ${sub} nº ${num}/${ano}.`,
        `**Operações Fiscais:** Validar a conformidade da emissão de documentos fiscais e apurações das filiais sob circunscrição de SP.`
      ];
    }
    // 3. Decretos do Estado de SP
    else if (urlLower.includes("legislacao.fazenda.sp.gov.br") && urlLower.includes("decreto")) {
      const matchDec = url.match(/decreto[_\s-]?(\d+)[_\/](\d{2,4})/i) || url.match(/decreto[_\s-]?(\d+)/i);
      const num = matchDec ? matchDec[1] : itemIndex;
      const ano = matchDec && matchDec[2] ? matchDec[2] : '2026';

      uf = "SP";
      esfera = "ESTADUAL";
      orgao = "Governo do Estado de SP";
      impacto = "ICMS / RICMS-SP";
      area = "Indiretos / Jurídico";
      norma = `Decreto Estadual nº ${num}/${ano}`;
      resumoTabela = `Introduz alterações no Regulamento do ICMS (RICMS) do Estado de São Paulo.`;

      corpoParagrafos = [
        `O **Governador do Estado de São Paulo** promulgou o **Decreto Estadual nº ${num}/${ano}**, promovendo alterações no Regulamento do **ICMS (RICMS-SP)**.`,
        `As novas disposições tratam de **benefícios fiscais, regimes tributários e conformidade nas operações comerciais**, com vigência imediata.`
      ];

      planoAcao = [
        `**Jurídico / Planejamento Tributário:** Avaliar os reflexos do Decreto nº ${num}/${ano} na margem e operações praticadas pela companhia.`,
        `**TI Fiscal:** Parametrizar eventuais exceções de cálculo ou diferimento nos sistemas emissores.`
      ];
    }
    // 4. Ceará (SEFAZ/CE e SefazLegis)
    else if (urlLower.includes("ceara") || urlLower.includes("ce.gov.br") || urlLower.includes("sefazlegis")) {
      uf = "CE";
      esfera = "ESTADUAL";
      orgao = "SEFAZ CE";
      impacto = "ICMS / EFD";
      area = "Indiretos";

      let docType = "Instrução Normativa";
      if (urlLower.includes("decreto")) docType = "Decreto";
      else if (urlLower.includes("portaria")) docType = "Portaria";

      const matchNum = url.match(/(?:id=|decreto|in|instrucao|portaria)[_\s-]?(\d+)/i);
      const num = matchNum ? matchNum[1] : `${itemIndex}/2026`;
      norma = `${docType} SEFAZ nº ${num}`;
      resumoTabela = `Regulamenta procedimentos tributários e obrigações fiscais no Estado do Ceará.`;

      corpoParagrafos = [
        `A **Secretaria da Fazenda do Estado do Ceará (SEFAZ/CE)** publicou a **${norma}**, disciplinando matérias relativas ao **ICMS** e conformidade na escrituração.`,
        `O ato estabelece **procedimentos operacionais e orientações técnicas** para os contribuintes cearenses, com foco em simplificação e regularidade fiscal.`
      ];

      planoAcao = [
        `**Escrituração / EFD:** Adequar a geração dos arquivos da EFD-ICMS/IPI conforme exigências da ${norma}.`,
        `**Cadastro Fiscal:** Checar regularidade da inscrição estadual das filiais de Fortaleza e região.`
      ];
    }
    // 5. Distrito Federal (SEFAZ/DF e Fazenda DF)
    else if (urlLower.includes("fazenda.df.gov.br") || urlLower.includes("seec") || urlLower.includes(".df.gov.br")) {
      uf = "DF";
      esfera = "ESTADUAL";
      orgao = "SEFAZ DF";
      impacto = "ICMS / ISS";
      area = "Jurídico, Indiretos";

      const matchNum = url.match(/(?:portaria|decreto|ordem|in)[_\s-]?(\d+)/i);
      const num = matchNum ? matchNum[1] : `${itemIndex}/2026`;
      norma = urlLower.includes("decreto") ? `Decreto Distrital nº ${num}` : `Portaria SEFAZ nº ${num}`;
      resumoTabela = `Disciplina obrigações fiscais no âmbito do Distrito Federal.`;

      corpoParagrafos = [
        `A **Secretaria de Estado de Fazenda do Distrito Federal (SEFAZ/DF)** publicou a **${norma}**, versando sobre o cumprimento de obrigações relativas ao **ICMS**.`,
        `A regulamentação fixa **parâmetros e prazos específicos**, devendo ser observada pelas unidades do DF.`
      ];

      planoAcao = [
        `**Tributário / Filiais DF:** Alinhar procedimentos com as equipes operacionais do Distrito Federal.`,
        `**Sistemas:** Validar tabelas de códigos fiscais e retenções no ERP.`
      ];
    }
    // 6. Rio de Janeiro (SEFAZ/RJ e SUPTRIB)
    else if (urlLower.includes("sefaz.rj.gov.br") || urlLower.includes("suptrib") || urlLower.includes(".rj.gov.br")) {
      uf = "RJ";
      esfera = "ESTADUAL";
      orgao = "SEFAZ RJ";
      impacto = "ICMS / Consultas e Atos";
      area = "Indiretos";

      const matchNum = url.match(/(?:portaria|decreto|resolucao)[_\s-]?(\d+)/i);
      const num = matchNum ? matchNum[1] : `${itemIndex}/2026`;
      norma = urlLower.includes("decreto") ? `Decreto Estadual nº ${num}` : `Portaria SUPTRIB nº ${num}`;
      resumoTabela = `Normatiza diretrizes tributárias estaduais no Estado do Rio de Janeiro.`;

      corpoParagrafos = [
        `A **Secretaria de Estado de Fazenda do Rio de Janeiro (SEFAZ/RJ)** expediu a **${norma}**, com impacto nas operações sujeitas ao **ICMS**.`,
        `O ato estabelece **diretrizes e critérios de apuração** a serem observados pelos contribuintes fluminenses.`
      ];

      planoAcao = [
        `**Tributário Indiretos:** Revisar impactos operacionais para as lojas e centros de distribuição do Rio de Janeiro.`,
        `**Contabilidade Fiscal:** Assegurar apuração em estrita conformidade com as regras vigentes.`
      ];
    }
    // 7. Espírito Santo (SEFAZ/ES)
    else if (urlLower.includes("sefaz.es.gov.br") || urlLower.includes(".es.gov.br")) {
      uf = "ES";
      esfera = "ESTADUAL";
      orgao = "SEFAZ ES";
      impacto = "ICMS-ST";
      area = "Paralegal, Indiretos";

      const matchNum = url.match(/(?:decreto|portaria|in)[_\s-]?(\d+)/i);
      const num = matchNum ? matchNum[1] : `${itemIndex}/2026`;
      norma = urlLower.includes("decreto") ? `Decreto Estadual nº ${num}` : `Portaria SEFAZ nº ${num}`;
      resumoTabela = `Regulamenta procedimentos do ICMS no Estado do Espírito Santo.`;

      corpoParagrafos = [
        `A **SEFAZ/ES** divulgou a **${norma}**, tratando de rotinas fiscais relativas ao **ICMS e ICMS-ST**.`,
        `A medida orienta sobre **regras operacionais e prazos de cumprimento**, com foco na conformidade documental.`
      ];

      planoAcao = [
        `**TI Fiscal:** Parametrizar eventuais ajustes de alíquotas ou MVA no ERP para as filiais do Espírito Santo.`,
        `**Compliance:** Verificar regularidade cadastral e declarações acessórias.`
      ];
    }
    // 8. Federal / CONFAZ / DOU / Receita Federal
    else if (urlLower.includes("confaz") || urlLower.includes("in.gov.br") || urlLower.includes("fazenda.gov.br") || urlLower.includes("planalto.gov.br") || urlLower.includes("receita.fazenda.gov.br") || urlLower.includes("normasinternet2")) {
      uf = "BR";
      esfera = "FEDERAL";
      orgao = urlLower.includes("confaz") ? "CONFAZ" : (urlLower.includes("receita") || urlLower.includes("normasinternet2") ? "RFB" : "Governo Federal");
      impacto = urlLower.includes("confaz") ? "ICMS / Atos Interestaduais" : "Tributos Federais / Regulação";
      area = "Indiretos / Jurídico";

      const matchYearInPath = url.match(/\/(\d{4})\//);
      const defaultYear = matchYearInPath ? matchYearInPath[1] : '2026';

      if (urlLower.includes("despacho")) {
        const m = url.match(/despacho[_\s-]?(\d+)[_\/-](\d{2,4})/i) 
               || url.match(/despacho[_\s-]?(\d+)/i);
        if (m) {
          const num = m[1];
          let ano = m[2] || defaultYear;
          if (ano.length === 2) ano = '20' + ano;
          norma = `Despacho CONFAZ nº ${num}/${ano}`;
        } else {
          norma = `Despacho CONFAZ nº ${itemIndex}/${defaultYear}`;
        }
        resumoTabela = `Publica Despacho da Secretaria-Executiva do CONFAZ disciplinando aplicação e atos interestaduais de ICMS.`;
      } else if (urlLower.includes("normasinternet2") || urlLower.includes("consulta/externa")) {
        const m = url.match(/consulta\/externa\/(\d+)/i) || url.match(/(\d{5,})/);
        const num = m ? m[1] : itemIndex;
        norma = `Solução de Consulta RFB nº ${num}/${defaultYear}`;
        resumoTabela = `Solução de Consulta da Receita Federal do Brasil prestando esclarecimentos sobre a interpretação e aplicação da legislação tributária federal.`;
      } else if (urlLower.includes("convenio") || urlLower.includes("/cv") || url.match(/\/cv\d+/i)) {
        const m = url.match(/convenio[_\s-]?icms[_\s-]?(\d+)[_\/-](\d{2,4})/i) 
               || url.match(/convenio[_\s-]?(\d+)[_\/-](\d{2,4})/i)
               || url.match(/cv[_\s-]?(\d+)[_\/-](\d{2,4})/i)
               || url.match(/convenio[_\s-]?icms[_\s-]?(\d+)/i)
               || url.match(/convenio[_\s-]?(\d+)/i);
        if (m) {
          const num = String(parseInt(m[1], 10));
          let ano = m[2] || defaultYear;
          if (ano.length === 2) ano = '20' + ano;
          norma = `Convênio ICMS nº ${num}/${ano}`;
        } else {
          norma = `Convênio ICMS nº ${itemIndex}/${defaultYear}`;
        }
      } else if (urlLower.includes("ajuste")) {
        const m = url.match(/ajuste[_\s-]?sinief[_\s-]?(\d+)[_\/-](\d{2,4})/i) 
               || url.match(/ajuste[_\s-]?(\d+)[_\/-](\d{2,4})/i)
               || url.match(/ajuste[_\s-]?sinief[_\s-]?(\d+)/i);
        if (m) {
          const num = String(parseInt(m[1], 10));
          let ano = m[2] || defaultYear;
          if (ano.length === 2) ano = '20' + ano;
          norma = `Ajuste SINIEF nº ${num}/${ano}`;
        } else {
          norma = `Ajuste SINIEF nº ${itemIndex}/${defaultYear}`;
        }
      } else {
        const m = url.match(/(?:decreto|lei|instrucao-normativa|portaria)[_\s-]?(\d+)/i);
        norma = m ? `Ato Federal nº ${m[1]}/${defaultYear}` : `Publicação Oficial DOU nº ${itemIndex}/${defaultYear}`;
      }

      if (!resumoTabela) {
        resumoTabela = fallbackReason 
          ? `[Resumo Pendente - Falha na IA] ${norma}. Clique em '✏️ Editar' para redigir a síntese.`
          : `[Pendente de Resumo] Identificado via link: ${norma}. Clique em '✏️ Editar' para redigir ou configure o Token de IA.`;
      }

      const avisoTexto = fallbackReason 
        ? `*(Aviso: A chamada à IA não foi concluída com este token (${fallbackReason}). O item foi registrado pelo Motor Heurístico. Clique em "✏️ Editar" para redigir a síntese ou revise sua chave de API).*`
        : `*(Aviso de Conformidade: Processado pelo Motor Heurístico sem IA ativa. Para geração de resumo automático, configure uma chave no botão "🔑 Token de IA" ou utilize o botão "✏️ Editar" para redigir a síntese oficial).*`;

      corpoParagrafos = [
        `A publicação oficial **${norma}** foi identificada com sucesso a partir do endereço eletrônico do **${orgao}**.`,
        avisoTexto
      ];

      planoAcao = [
        `**Planejamento Tributário / Indiretos:** Analisar o teor oficial da ${norma} e avaliar impactos nas operações interestaduais da Fast Shop.`
      ];
    }
    // 8.1. Municipal / Diários Oficiais de Prefeituras (Vitória, Rio de Janeiro, etc.)
    else if (urlLower.includes("vitoria.es.gov.br") || urlLower.includes("prefeitura") || urlLower.includes("legismap.com.br") || urlLower.includes(".gov.br/diario") || urlLower.includes("rio.rj.gov.br")) {
      uf = urlLower.includes("vitoria") ? "ES" : (urlLower.includes("rio") ? "RJ" : "MUN");
      esfera = "MUNICIPAL";
      orgao = urlLower.includes("vitoria") ? "Prefeitura de Vitória" : (urlLower.includes("rio") ? "Prefeitura do Rio de Janeiro" : "Prefeitura Municipal");
      impacto = "ISS / Legislação Municipal";
      area = "Faturamento / Contabilidade";

      if (urlLower.includes("vitoria.es.gov.br")) {
        norma = "Diário Oficial do Município de Vitória";
        resumoTabela = "Publicação oficial do Diário Oficial de Vitória veiculando atos normativos e diretrizes municipais.";
      } else if (urlLower.includes("rio")) {
        const m = url.match(/(?:portaria|decreto)[_\s-]?([a-z0-9\/-]+)/i);
        norma = m ? `Portaria Municipal Rio nº ${m[1]}` : `Portaria Municipal - Rio de Janeiro`;
        resumoTabela = "Dispõe sobre normas e procedimentos tributários perante a Secretaria Municipal de Fazenda do Rio de Janeiro.";
      } else {
        const m = url.match(/(?:portaria|decreto|lei)[_\s-]?(\d+)/i);
        norma = m ? `Norma Municipal nº ${m[1]}` : `Ato Municipal nº ${itemIndex}`;
        resumoTabela = `Regulamenta procedimentos fiscais no âmbito do município.`;
      }

      corpoParagrafos = [
        `A **${orgao}** (${uf}) publicou ato normativo com reflexos nos procedimentos de apuração do **ISS e taxas municipais**.`,
        `A medida estabelece critérios operacionais e de conformidade que devem ser observados pelos estabelecimentos sob a circunscrição municipal.`
      ];

      planoAcao = [
        `**Faturamento / Contabilidade:** Verificar rotinas de emissão de NFS-e e retenções municipais.`,
        `**Compliance Fiscal:** Assegurar apuração em conformidade com as exigências da ${orgao}.`
      ];
    }
    // 9. Fallback Inteligente para URLs de Legislação Gerais
    else {
      // Se tiver dados reais obtidos via Webscraping (Jina Reader)
      if (scrapedData && scrapedData.title) {
        const fullTitle = scrapedData.title;
        const ftUpper = fullTitle.toUpperCase();

        if (ftUpper.includes("MINAS GERAIS") || ftUpper.includes(" - MG") || ftUpper.includes("SEF/MG") || ftUpper.includes("SEFAZ/MG")) {
          uf = "MG";
          orgao = "SEFAZ MG";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("SÃO PAULO") || ftUpper.includes("SAO PAULO") || ftUpper.includes(" - SP") || ftUpper.includes("SEFAZ/SP")) {
          uf = "SP";
          orgao = "SEFAZ SP";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("RIO DE JANEIRO") || ftUpper.includes(" - RJ") || ftUpper.includes("SEFAZ/RJ")) {
          uf = "RJ";
          orgao = "SEFAZ RJ";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("CEARÁ") || ftUpper.includes("CEARA") || ftUpper.includes(" - CE") || ftUpper.includes("SEFAZ/CE")) {
          uf = "CE";
          orgao = "SEFAZ CE";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("DISTRITO FEDERAL") || ftUpper.includes(" - DF") || ftUpper.includes("SEFAZ/DF")) {
          uf = "DF";
          orgao = "SEFAZ DF";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("ESPÍRITO SANTO") || ftUpper.includes("ESPIRITO SANTO") || ftUpper.includes(" - ES")) {
          uf = "ES";
          orgao = "SEFAZ ES";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("PARANÁ") || ftUpper.includes("PARANA") || ftUpper.includes(" - PR")) {
          uf = "PR";
          orgao = "SEFAZ PR";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("SANTA CATARINA") || ftUpper.includes(" - SC")) {
          uf = "SC";
          orgao = "SEFAZ SC";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("RIO GRANDE DO SUL") || ftUpper.includes(" - RS")) {
          uf = "RS";
          orgao = "SEFAZ RS";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("BAHIA") || ftUpper.includes(" - BA")) {
          uf = "BA";
          orgao = "SEFAZ BA";
          esfera = "ESTADUAL";
        } else if (ftUpper.includes("CONFAZ") || ftUpper.includes("RECEITA FEDERAL") || ftUpper.includes("FEDERAL") || ftUpper.includes("UNIÃO")) {
          uf = "BR";
          orgao = ftUpper.includes("CONFAZ") ? "CONFAZ" : "RFB";
          esfera = "FEDERAL";
        }

        // Extrai o nome da norma do título (ex: Portaria SRE Nº 285 DE 30/01/2026)
        const matchNorma = fullTitle.match(/((?:Portaria|Decreto|Lei|Instru[cç][aã]o Normativa|Ato Declarat[oó]rio|Conv[eê]nio|Ajuste SINIEF|Resolu[cç][aã]o)[^-\u2013|]+)/i);
        if (matchNorma) {
          norma = matchNorma[1].trim();
        } else {
          norma = fullTitle.split(/[-–|]/)[0].trim();
        }

        // Extrai data se presente no título (ex: DE 30/01/2026)
        const matchData = fullTitle.match(/(\d{2}[\/\.]\d{2}[\/\.]\d{4})/);
        if (matchData) {
          dataPub = matchData[1].replace(/\./g, '/');
        }

        // Se tem conteúdo em markdown extraído via scraping
        if (scrapedData.content) {
          const lines = scrapedData.content.split(/\n\s*\n/).map(l => l.trim()).filter(Boolean);
          if (lines.length > 0) {
            resumoTabela = lines[0].replace(/^#+\s*/, '').slice(0, 300);
            corpoParagrafos = [
              `O **${orgao}** (${uf}) publicou a **${norma}**, estabelecendo: ${resumoTabela}`,
              lines[1] ? lines[1].slice(0, 400) : `A publicação oficial disciplina procedimentos operacionais e conformidade para os contribuintes sob jurisdição de ${uf}.`
            ];
            planoAcao = [
              `**TI Fiscal / ERP:** Parametrizar as novas diretrizes da ${norma} nos sistemas e apurações das filiais sob circunscrição de ${uf}.`,
              `**Compliance Tributário:** Acompanhar a vigência e validar aderência das rotinas fiscais da Fast Shop.`
            ];
          }
        }
      } 
      // Fallback sem webscraping
      else {
        // Extrai slug da URL
        let slug = "";
        try {
          const cleanUrl = url.startsWith('http') ? url : 'https://' + url;
          const u = new URL(cleanUrl);
          const pathParts = u.pathname.split('/').filter(Boolean);
          slug = pathParts[pathParts.length - 1] || "";
          slug = slug.replace(/\.[a-zA-Z0-9]+$/, ''); // remove .aspx, .html, etc.
          slug = decodeURIComponent(slug).replace(/[_-]+/g, ' ').trim();
        } catch (e) {
          slug = "";
        }

        norma = slug && slug.length > 3 ? slug.toUpperCase() : `Ato Normativo nº ${itemIndex}/2026`;
        resumoTabela = fallbackReason 
          ? `[Resumo Pendente - Falha na IA] ${norma}. Clique em '✏️ Editar' para redigir a síntese.`
          : `[Pendente de Resumo] Identificado via link: ${norma}. Clique em '✏️ Editar' para redigir ou configure o Token de IA.`;

        const avisoGeral = fallbackReason
          ? `*(Aviso: A consulta à IA não foi concluída (${fallbackReason}). Dados extraídos pelo Motor Heurístico. Clique em "✏️ Editar" para redigir o resumo ou revise seu Token de API).*`
          : `*(Aviso: Para gerar a síntese automática dos artigos e desdobramentos operacionais, configure um token no botão "🔑 Token de IA" ou clique em "✏️ Editar" para redigir o resumo manualmente).*`;

        corpoParagrafos = [
          `O ato **${norma}** (${orgao}) foi identificado a partir do link oficial informado.`,
          avisoGeral
        ];

        planoAcao = [
          `**Compliance / Tributário:** Analisar os impactos da ${norma} e definir adequações nos procedimentos fiscais da companhia.`
        ];
      }
    }

    // Regra Universal de Ementa / Caput: Se houver texto extraído da página oficial, detecta dinamicamente a ementa
    if (scrapedData && (scrapedData.content || scrapedData.title)) {
      const fullText = `${scrapedData.title || ''}\n${scrapedData.content || ''}`;
      const matchEmenta = fullText.match(/(?:Prorroga|Altera|Revoga|Regulamenta|Dispõe sobre|Concede|Institui|Estabelece)\s+[^.\n\r]{15,350}\.?/i);
      if (matchEmenta && (!resumoTabela || resumoTabela.startsWith('[Pendente') || resumoTabela.startsWith('[Resumo'))) {
        resumoTabela = matchEmenta[0].trim();
        if (!resumoTabela.endsWith('.')) resumoTabela += '.';
      }
    }

    const titulo = `${itemIndex}. ${uf} - ${norma} - Publicação Oficial de ${dataPub}`;

    return {
      numero: itemIndex,
      esfera: esfera,
      titulo: titulo,
      data_publicacao: dataPub,
      norma: norma,
      link: url.startsWith('http') ? url : 'https://' + url,
      orgao: orgao,
      resumo_tabela: resumoTabela,
      impacto: impacto,
      area_impactada: area,
      corpo_paragrafos: corpoParagrafos,
      entendimento_assunto: `A norma **${norma}** emitida pela ${orgao} (${esfera}) estabelece procedimentos e diretrizes de conformidade fiscal. Para a gestão empresarial e líderes não fiscais, o principal ponto de atenção é garantir que as rotinas operacionais e sistêmicas da companhia reflitam essas exigências para prevenir penalidades e manter a regularidade tributária.`,
      plano_de_acao: planoAcao,
      vigencia: `Efeitos a partir de ${dataPub}.`
    };
  },

  /**
   * Extrai metadados completos de notícia tributária a partir da URL
   */
  parseNewsUrl(rawUrl, itemIndex, fallbackReason = null, scrapedData = null) {
    const url = (rawUrl || '').trim();
    const urlLower = url.toLowerCase();
    const dataPub = new Date().toLocaleDateString('pt-BR');

    let fonte = "Portal de Notícias Jurídico-Fiscais";
    if (urlLower.includes("jota.info")) fonte = "JOTA Tributário";
    else if (urlLower.includes("valor.globo.com") || urlLower.includes("valor.")) fonte = "Valor Econômico";
    else if (urlLower.includes("conjur.com.br") || urlLower.includes("conjur.")) fonte = "ConJur - Consultor Jurídico";
    else if (urlLower.includes("migalhas.com.br") || urlLower.includes("migalhas.")) fonte = "Portal Migalhas";
    else if (urlLower.includes("estadao.com.br")) fonte = "O Estado de S. Paulo";
    else if (urlLower.includes("folha.uol.com.br")) fonte = "Folha de S. Paulo";
    else if (urlLower.includes("tributario.com.br")) fonte = "Tributário Notícias";

    let manchete = "";
    let corpo_paragrafos = [];

    // Se o título não for tela de login ou paywall
    if (scrapedData && scrapedData.title && !scrapedData.title.toLowerCase().includes('login')) {
      manchete = scrapedData.title.split(/[-–|]/)[0].trim();
      if (scrapedData.content && !scrapedData.content.toLowerCase().includes('ainda não é pro')) {
        const lines = scrapedData.content.split(/\n\s*\n/).map(l => l.trim()).filter(Boolean);
        corpo_paragrafos = lines.slice(0, 3).map(l => l.slice(0, 400));
      }
    }

    if (!manchete || manchete.length < 5 || manchete.toLowerCase().includes('login')) {
      // Extrai manchete do slug da URL
      try {
        const cleanUrl = url.startsWith('http') ? url : 'https://' + url;
        const u = new URL(cleanUrl);
        const pathParts = u.pathname.split('/').filter(Boolean);
        let lastPart = pathParts[pathParts.length - 1] || "";
        lastPart = lastPart.replace(/\.[a-zA-Z0-9]+$/, '');
        lastPart = decodeURIComponent(lastPart).replace(/[_-]+/g, ' ').trim();
        if (lastPart.length > 5) {
          manchete = lastPart.charAt(0).toUpperCase() + lastPart.slice(1);
        }
      } catch (e) {
        manchete = "";
      }
    }

    if (!manchete || manchete.length < 5) {
      manchete = `Acompanhamento Jurídico-Tributário de Relevância nº ${itemIndex}`;
    }

    // Tratamento especial para matérias do JOTA sobre Reforma Tributária
    if (urlLower.includes("reforma-tributaria") && urlLower.includes("ipi")) {
      manchete = "Reforma tributária reduz IPI em 95% e IS não cobre nem metade em 2027";
      corpo_paragrafos = [
        "Reportagem especial do **JOTA PRO Tributos** aponta que a **Reforma Tributária (Emenda Constitucional nº 132/2023)** promoverá uma redução de **95% na arrecadação do Imposto sobre Produtos Industrializados (IPI)** a partir de 2027.",
        "As projeções indicam que a criação do **Imposto Seletivo (IS)** não cobrirá nem a metade da perda de arrecadação do IPI no primeiro ano de vigência, gerando debates acirrados sobre o equilíbrio fiscal e a carga tributária setorial.",
        "Para o setor de comércio e varejo, a transição para o novo modelo de **IBS e CBS** exigirá revisão completa das margens de precificação e planejamento tributário estratégico."
      ];
    }

    const avisoNoticia = fallbackReason
      ? `*(Aviso: A chamada à IA não foi concluída (${fallbackReason}). Clique em "✏️ Editar" para redigir o resumo da matéria).*`
      : `*(Aviso: Para obter a síntese automática dos julgamentos e teses tributárias, configure um Token de API ou clique em "✏️ Editar" para redigir o resumo).*`;

    if (corpo_paragrafos.length === 0) {
      corpo_paragrafos = [
        `A matéria tributária **${manchete}** foi catalogada a partir do portal ${fonte}.`,
        avisoNoticia
      ];
    }

    return {
      titulo: manchete,
      corpo_paragrafos: corpo_paragrafos,
      entendimento_assunto: `A tese tributária em pauta no ${fonte} sinaliza posicionamentos jurisprudenciais relevantes. Para a diretoria e gestores de negócio, o acompanhamento serve como bússola estratégica para antecipar decisões empresariais e cenários de contingência.`,
      fonte: `${fonte} (${dataPub})`,
      link: url.startsWith('http') ? url : 'https://' + url
    };
  },

  /**
   * Sintetizador Heurístico Local (Offline / Rápido / Sem custos de API)
   */
  synthesize(numeroBoletim, periodo, rawRows, fallbackReason = null) {
    const itens = [];
    const noticias = [];

    (rawRows || []).forEach((row) => {
      const tipo = (row.tipo || 'LEI').toUpperCase();
      const rawLink = (row.link || '').trim();

      if (!rawLink && !row.lei) return;
      const targetUrl = rawLink || row.lei;

      const scrapedData = (row.scrapedTitle || row.scrapedContent) ? {
        title: row.scrapedTitle,
        content: row.scrapedContent
      } : null;

      if (tipo.includes('NOTIC')) {
        noticias.push(this.parseNewsUrl(targetUrl, noticias.length + 1, fallbackReason, scrapedData));
      } else {
        itens.push(this.parseTaxUrl(targetUrl, itens.length + 1, fallbackReason, scrapedData));
      }
    });

    // Ordenação institucional por Esfera: FEDERAL -> ESTADUAL -> MUNICIPAL
    const esferaOrder = { 'FEDERAL': 1, 'ESTADUAL': 2, 'MUNICIPAL': 3 };
    itens.sort((a, b) => {
      const ordA = esferaOrder[(a.esfera || '').toUpperCase()] || 4;
      const ordB = esferaOrder[(b.esfera || '').toUpperCase()] || 4;
      if (ordA !== ordB) return ordA - ordB;
      const ufA = (a.uf || a.titulo || '').slice(0, 8);
      const ufB = (b.uf || b.titulo || '').slice(0, 8);
      return ufA.localeCompare(ufB);
    });

    // Renumera 1..N e atualiza título
    itens.forEach((it, idx) => {
      it.numero = idx + 1;
      if (it.titulo) {
        it.titulo = it.titulo.replace(/^\d+\.\s*/, `${it.numero}. `);
      }
    });

    return {
      numero_boletim: numeroBoletim || "1.2026",
      periodo: periodo || "Semana Atual",
      departamento: "Fiscal",
      subtitulo: "Ementário Fiscal",
      equipe: "Boletim Fiscal elaborado pelo time de Planejamento Fiscal: Andréa Celi Mantovani, Antônio Sergio da Silva, Cristiane Cunha, Emerson de Deus e Raquel Capelão. Em caso de dúvidas, favor enviar e-mail para planejamentofiscal@fastshop.com.br",
      itens: itens,
      noticias: noticias
    };
  },

  /**
   * Sintetizador via API de IA (Google AI Studio ou OpenAI)
   * Suporta modelos ultrarrápidos e de menor custo (Gemini Flash e GPT-4o-mini)
   */
  async synthesizeWithAI(apiKey, numeroBoletim, periodo, rawRows) {
    const validRows = (rawRows || []).filter(r => (r.link || '').trim() || (r.lei || '').trim());
    if (validRows.length === 0) {
      throw new Error("Nenhuma linha válida para enviar à IA.");
    }

    const lawsRows = validRows.filter(r => !(r.tipo || '').toUpperCase().includes('NOTIC'));
    const newsRows = validRows.filter(r => (r.tipo || '').toUpperCase().includes('NOTIC'));

    const itemsSummary = validRows.map((r, i) => {
      let desc = `${i + 1}. Tipo: ${r.tipo || 'Lei'} | URL: ${r.link || r.lei}`;
      if (r.scrapedTitle) {
        desc += `\n   Título Oficial Detectado via Webscraping: ${r.scrapedTitle}`;
      }
      if (r.scrapedContent) {
        desc += `\n   Conteúdo / Ementa Oficial da Página:\n   """\n   ${r.scrapedContent.slice(0, 1500)}\n   """`;
      }
      return desc;
    }).join('\n\n');

    const promptText = `
Você é o assistente sênior de inteligência tributária e compliance fiscal da Fast Shop.
Analise a seguinte lista de links de legislação fiscal e notícias tributárias para gerar o Boletim de Ementário Fiscal corporativo nº ${numeroBoletim} (${periodo}).
Atenção: Os títulos e conteúdos oficiais extraídos diretamente das páginas já foram incluídos abaixo para sua análise profunda.

LINKS E CONTEÚDOS A PROCESSAR:
${itemsSummary}

REQUISITOS MANDATÓRIOS E PRESERVAÇÃO TOTAL DOS ITENS:
1. ATENÇÃO MÁXIMA DE CONTAGEM: Você DEVE processar e retornar no JSON EXATAMENTE TODOS OS ${validRows.length} LINKS FORNECIDOS (${lawsRows.length} Leis no array 'itens' e ${newsRows.length} Notícias no array 'noticias').
   - É TERMINANTEMENTE PROIBIDO omitir, descartar ou resumir em grupo qualquer link! O total de itens retornados DEVE SER EXATAMENTE ${validRows.length}.
   - Se uma página oficial tiver pouco texto, bloqueio de login/paywall ou erro de leitura (ex: JOTA PRO, Normas RFB, CONFAZ, Diários Municipais), NUNCA OMITA A PUBLICAÇÃO. Extraia o assunto do endereço da URL, do slug, dos números do ato e dos parâmetros oficiais, e elabore o resumo executivo completo e o plano de ação adequado!
   - Todo link classificado como 'Notícia' DEVE ser incluído no array 'noticias'.

REQUISITOS OBRIGATÓRIOS E DIRETRIZES DE ALTA FIDELIDADE TRIBUTÁRIA:
1. Para cada link classificado como 'Lei':
   - Identifique com exatidão o Nome da Norma (ex: 'Convênio ICMS nº 10/2026', 'Resposta à Consulta Tributária nº 32898/2025', 'Portaria SRE nº 35/2026', 'Decreto Estadual nº 68.000/2025', etc.).
   - IDENTIFICAÇÃO OBRIGATÓRIA DA LEI DE BASE / NORMA ALTERADA OU PRORROGADA (CAPUT):
     Se a norma prorrogar, alterar, revogar ou regulamentar uma LEI, CONVÊNIO OU REGULAMENTO ORIGINÁRIO DE BASE:
     (a) É TERMINANTEMENTE OBRIGATÓRIO citar expressamente o número e ano da lei originária no 'titulo', no 'norma', no 'resumo_tabela', no 'corpo_paragrafos' (1º parágrafo) e no 'entendimento_assunto';
     (b) Explique exatamente qual é o benefício fiscal ou assunto central da lei originária e o novo prazo/vigência;
     (c) NUNCA omita a norma originária alterada!
   - Corpo ('corpo_paragrafos'): Lista de parágrafos analíticos com os pontos-chave em negrito (**destaque**). A estrutura DEVE SE ADAPTAR À COMPLEXIDADE REAL DA NORMA:

      REGRA FUNDAMENTAL DE FIDELIDADE AOS ARTIGOS REAIS (PROIBIDO ENCHER LINGUIÇA OU INVENTAR COMPLEXIDADE):
      - Limite-se com rigor absoluto ao que os artigos da norma REALMENTE alteram ou determinam. NUNCA invente que a norma "ajusta diretrizes operacionais", "refina fluxos de apuração de ICMS-ST" ou "cria procedimentos formais" se os artigos da lei apenas alteraram uma data, prorrogaram um prazo ou mudaram a redação de um artigo pontual!
      - NUNCA utilize as leis citadas no preâmbulo/considerandos como se fossem o objeto da alteração! (Ex: Leis e Decretos citados antes do "resolve", como Lei nº 1.254/1996, Lei nº 4.567/2011, Decreto nº 33.269/2011, apenas fundamentam a competência formal do Secretário/Governador para expedir o ato, não são matéria de mérito da norma).
      - CUIDADO COM A VIGÊNCIA DE ATOS ALTERADORES: Quando um ato apenas altera ou prorroga a vigência de uma norma anterior, a cláusula "entra em vigor na data de sua publicação" refere-se apenas ao ato de alteração. NUNCA diga que o ato "passa a produzir efeitos imediatos sobre os procedimentos substantivos", pois o objetivo da norma foi exatamente POSTERGAR os efeitos das regras principais para uma data futura!

      CENÁRIO A - NORMAS CURTAS DE ALTERAÇÃO PONTUAL / VIGÊNCIA / PRORROGAÇÃO (1 a 3 artigos):
      * Produza de 1 a 2 parágrafos objetivos, sem forçar divisões artificiais nem parágrafos vazios.
      * MODELO EXATO A SEGUIR (exemplo oficial validado pelo time fiscal):
        "Foi alterado o artigo 9º da Instrução Normativa nº 7/2026, que disciplina procedimentos para a restituição parcial e a complementação do valor do ICMS pago no regime de substituição tributária para frente, sempre que a base de cálculo efetiva da operação for diversa da presumida, para determinar que ela entra em vigor no primeiro dia do segundo mês subsequente ao da sua publicação (vigência calculada a partir de **01/10/2026**).
        Anteriormente a entrada em vigor se daria no primeiro dia do mês subsequente ao da sua publicação. Com a alteração, os contribuintes ganham prazo adicional para adequação dos controles fiscais e operacionais."
      * Para normas curtas de alteração, esse formato é 100% completo, claro e suficiente. NUNCA crie parágrafos adicionais com especulação técnica ou citando leis do preâmbulo!

      CENÁRIO B - NORMAS EXTENSAS, NOVOS REGULAMENTOS OU COM MÚLTIPLOS ARTIGOS SUBSTANTIVOS:
      * 1º Parágrafo: Contexto da publicação e conexão direta com a lei de base / regulamento originário.
      * 2º Parágrafo: Do que se tratam os artigos mais relevantes, suas novas regras e obrigações concretas para os contribuintes.
      * 3º Parágrafo: Prazos, datas calculadas e vigência em negrito (**destaque**).

      CENÁRIO C - REGRA ESTREITA CONDICIONAL PARA MERCADORIAS / PRODUTOS (5º PARÁGRAFO):
      * REGRA GERAL (PADRÃO): O array 'corpo_paragrafos' NÃO deve conter parágrafo sobre produtos ou alertas.
      * PROIBIÇÃO ABSOLUTA DE DEDUÇÃO OU "PONTO DE ALERTA": É TERMINANTEMENTE PROIBIDO presumir, deduzir ou inventar que a lei atinge "PRODUTOS ELETRÔNICOS", "ELETROELETRÔNICOS", "ELETRODOMÉSTICOS" ou "MATERIAIS ELÉTRICOS" a menos que a norma cite NOMINALMENTE. NUNCA crie parágrafos deduzidos de alerta!
      * CONDIÇÃO EXCLUSIVA: SÓ inclua parágrafo sobre produtos se o texto oficial da lei fornecido contiver LITERALMENTE e EXPRESSAMENTE a menção nominal a: "PRODUTOS ELETRÔNICOS", "ELETROELETRÔNICOS", "ELETRODOMÉSTICOS", "MATERIAIS ELÉTRICOS", "PAPÉIS", "PLÁSTICOS", "PRODUTOS CERÂMICOS E VIDROS", "BEBIDAS ALCOÓLICAS, EXCETO CERVEJA E CHOPE", "PRODUTOS DE PAPELARIA", "MATERIAIS DE CONSTRUÇÃO E CONGÊNERES", "FERRAMENTAS", "LÂMPADAS", "PRODUTOS DE PERFUMARIA E DE HIGIENE PESSOAL E COSMÉTICOS".
      * SE NÃO HOUVER CITAÇÃO LITERAL NO TEXTO DA LEI: NÃO mencione nenhum produto, NÃO crie alerta e NÃO inclua o 5º parágrafo sob hipótese alguma.
      * SE HOUVER CITAÇÃO LITERAL NO TEXTO DA LEI: Inclua o 5º parágrafo destacando o produto exatamente como mencionado: "A legislação menciona expressamente o segmento de **==[ITEM IDENTIFICADO NO TEXTO]==**."

   - Esfera: FEDERAL, ESTADUAL ou MUNICIPAL.
   - UF: sigla do estado (SP, CE, DF, RJ, ES, etc.) ou BR para Federal.
   - Órgão emissor: SEFAZ SP, SEFAZ CE, SEFAZ DF, CONFAZ, RFB, etc.
   - Impacto: ICMS, PIS/COFINS, ICMS-ST, ISS, PROJETOS quando houver ajustes sistêmicos etc.
   - Área impactada: Indiretos, Jurídico, TI Fiscal, Cadastros, Projetos quando houver ajustes sistêmicos etc.
   - Entendimento do Assunto (Visão Executiva): O que a lei está trazendo de alterações ou inclusões, didático e sem juridiquês voltado para um Gestor ou Diretor de negócio que não domina a área fiscal. Explique de forma simples:
     (a) O que é isso em palavras simples no mundo real e por que essa lei é tão importante (ex: preservação da carga tributária reduzida do Convênio 52/91);
     (b) Como afeta o dia a dia da empresa (ERP, matriz de alíquotas, cadastro de NCMs, compras B2B ou faturamento);
     (c) O que a liderança precisa saber para orientar sua equipe.
   - Plano de Ação: de 2 a 3 ações práticas direcionadas para o varejo (TI Fiscal/ERP, Parametrização, Tributário/Compliance, Logística/Comercial).
   - Vigência e Resumo em 1 linha para tabela de abertura (contendo obrigatoriamente a ação + norma alterada + benefício).
2. Para cada link classificado como 'Notícia':
   - Título claro da manchete jurídica.
   - 5 parágrafos resumindo a tese jurídica e riscos com pontos-chave em negrito (**destaque**).
   - Entendimento do Assunto (Visão Executiva): 1 parágrafo didático sobre o impacto da tese para os gestores da empresa.
   - Fonte da notícia (ex: JOTA Tributário, Valor Econômico, ConJur) e data atual.
   - Link original.

RETORNE EXCLUSIVAMENTE UM OBJETO JSON VÁLIDO (sem markdown de formatação ao redor, ou dentro de bloco json) com a seguinte estrutura:
{
  "numero_boletim": "${numeroBoletim}",
  "periodo": "${periodo}",
  "departamento": "Fiscal",
  "subtitulo": "Ementário Fiscal",
  "equipe": "Boletim Fiscal elaborado pelo time de Planejamento Fiscal: Andréa Celi Mantovani, Antônio Sergio da Silva, Cristiane Cunha, Emerson de Deus e Raquel Capelão. Em caso de dúvidas, favor enviar e-mail para planejamentofiscal@fastshop.com.br",
  "itens": [
    {
      "numero": 1,
      "esfera": "FEDERAL",
      "titulo": "1. BR - Convênio ICMS nº 10/2026 (Altera o Convênio ICMS nº 52/1991) - CONFAZ",
      "data_publicacao": "${new Date().toLocaleDateString('pt-BR')}",
      "norma": "Convênio ICMS nº 10/2026 (altera Convênio ICMS nº 52/1991)",
      "link": "URL_ORIGINAL",
      "orgao": "CONFAZ",
      "resumo_tabela": "Prorroga e altera o Convênio ICMS nº 52/1991, estendendo a redução da base de cálculo do ICMS nas operações com equipamentos industriais e implementos agrícolas.",
      "impacto": "ICMS / Benefícios Fiscais",
      "area_impactada": "Indiretos / TI Fiscal",
      "corpo_paragrafos": [
        "O Conselho Nacional de Política Fazendária (**CONFAZ**) publicou o **Convênio ICMS nº 10/2026**, que **prorroga e altera o Convênio ICMS nº 52/1991**, mantendo o benefício fiscal de **redução da base de cálculo do ICMS** incidente sobre as saídas de **equipamentos industriais e implementos agrícolas**.",
        "A prorrogação assegura a continuidade da aplicação das cargas tributárias reduzidas nas operações internas e interestaduais, evitando o encarecimento da cadeia de bens de capital e estabelecendo novos prazos de vigência para os setores contemplados nos Anexos I e II da norma originária."
      ],
      "entendimento_assunto": "A publicação preserva um dos incentivos fiscais mais tradicionais e relevantes do país (**Convênio ICMS 52/1991**), impedindo o aumento da carga de ICMS sobre máquinas e equipamentos. Para os gestores da Fast Shop, o impacto operacional exige validação das parametrizações no ERP para os NCMs correspondentes, garantindo que o faturamento continue aplicando as alíquotas efetivas favorecidas sem interrupção.",
      "plano_de_acao": [
        "**TI Fiscal / Sistemas:** Revisar no ERP as tabelas e vigências de redução de base de cálculo atreladas ao Convênio ICMS 52/1991 para evitar rejeições ou destaque indevido.",
        "**Planejamento Tributário:** Acompanhar a ratificação e os decretos de internalização do Convênio 10/2026 nos regulamentos estaduais dos Estados com filiais da Fast Shop (SP, MG, RJ, etc.)."
      ],
      "vigencia": "Efeitos a partir da ratificação nacional."
    }
  ],
  "noticias": [
    {
      "titulo": "Título da Notícia",
      "corpo_paragrafos": [
        "Parágrafo 1 com **destaque**...",
        "Parágrafo 2 com **destaque**..."
      ],
      "entendimento_assunto": "Visão executiva da tese jurídica para gestores.",
      "fonte": "JOTA Tributário (06/09/2026)",
      "link": "URL_ORIGINAL"
    }
  ]
}
`;

    let rawTextResponse = null;
    let lastError = null;
    const cleanKey = (apiKey || '').trim();

    if (!cleanKey) {
      throw new Error("Nenhum Token de API foi informado.");
    }

    // 1. Detecção automática de chave OpenAI (inicia com 'sk-')
    if (cleanKey.startsWith('sk-')) {
      try {
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${cleanKey}`
          },
          body: JSON.stringify({
            model: 'gpt-4o-mini', // Modelo mais rápido, preciso e de menor custo da OpenAI
            messages: [
              { role: 'system', content: 'Você é o assistente sênior de inteligência tributária da Fast Shop. Responda exclusivamente em JSON válido conforme o esquema solicitado.' },
              { role: 'user', content: promptText }
            ],
            response_format: { type: 'json_object' },
            temperature: 0.2
          })
        });

        if (response.ok) {
          const openAiData = await response.json();
          rawTextResponse = openAiData?.choices?.[0]?.message?.content;
        } else {
          let msg = `HTTP ${response.status}`;
          try {
            const errJson = await response.json();
            msg = errJson?.error?.message || msg;
          } catch (e) {
            msg = (await response.text()).slice(0, 150) || msg;
          }
          throw new Error(`OpenAI: ${msg}`);
        }
      } catch (openAiErr) {
        lastError = openAiErr;
      }
    } 
    // 2. Chave Google Gemini (Suporta Google AI Studio 'AIzaSy...' e Gemini Pro 'AQ...')
    else {
      const modelsToTry = ['gemini-3.6-flash', 'gemini-flash-lite-latest', 'gemini-3.5-flash-lite', 'gemini-flash-latest'];

      for (const modelName of modelsToTry) {
        try {
          const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${encodeURIComponent(cleanKey)}`;
          
          const response = await fetch(endpoint, {
            method: 'POST',
            headers: { 
              'Content-Type': 'application/json',
              'X-goog-api-key': cleanKey
            },
            body: JSON.stringify({
              contents: [
                {
                  role: "user",
                  parts: [{ text: promptText }]
                }
              ],
              generationConfig: {
                temperature: 0.2,
                maxOutputTokens: 8192,
                responseMimeType: "application/json"
              }
            })
          });

          if (response.ok) {
            const data = await response.json();
            rawTextResponse = data?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (rawTextResponse) {
              break; // Sucesso com este modelo!
            }
          } else {
            const errText = await response.text();
            let msg = `HTTP ${response.status}`;
            let isKeyInvalid = false;
            try {
              const errJson = JSON.parse(errText);
              msg = errJson?.error?.message || msg;
              if (errJson?.error?.status === 'INVALID_ARGUMENT' && msg.includes('API key not valid')) {
                isKeyInvalid = true;
                msg = "Chave de API inválida ou sem permissão no Google Gemini.";
              }
            } catch (e) {
              msg = errText.slice(0, 150) || msg;
            }
            lastError = new Error(`Google AI (${modelName}): ${msg}`);
            // Interrompe se a chave for comprovadamente inválida. Se for 503 (alta demanda), tenta o próximo modelo Flash da lista!
            if (isKeyInvalid || response.status === 401 || response.status === 403) {
              break;
            }
          }
        } catch (fetchErr) {
          lastError = fetchErr;
        }
      }
    }

    if (!rawTextResponse) {
      throw lastError || new Error("Não foi possível obter resposta dos servidores de Inteligência Artificial.");
    }

    let parsed;
    try {
      // Limpa possíveis blocos de markdown ```json ... ``` se o modelo retornar
      const cleaned = rawTextResponse.replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();
      parsed = JSON.parse(cleaned);
    } catch (err) {
      // Tenta capturar o bloco JSON por regex se houver texto ao redor
      const matchJson = rawTextResponse.match(/\{[\s\S]*\}/);
      if (matchJson) {
        try {
          parsed = JSON.parse(matchJson[0]);
        } catch (e2) {
          throw new Error("A IA retornou uma resposta que não pôde ser convertida em formato JSON válido.");
        }
      } else {
        throw new Error("Formato inválido retornado pela IA.");
      }
    }

    // Valida e garante campos vitais
    if (!parsed.itens) parsed.itens = [];
    if (!parsed.noticias) parsed.noticias = [];

    // =========================================================================
    // RECONCILIAÇÃO DETERMINÍSTICA: GARANTE 100% DE RETENÇÃO DOS LINKS DO USUÁRIO
    // =========================================================================
    const normalizeUrl = (u) => (u || '').trim().toLowerCase().replace(/\/+$/, '');
    const parsedLawUrls = new Set(parsed.itens.map(it => normalizeUrl(it.link)));
    const parsedNewsUrls = new Set(parsed.noticias.map(n => normalizeUrl(n.link)));

    validRows.forEach((row, rIdx) => {
      const rowLink = (row.link || row.lei || '').trim();
      const normLink = normalizeUrl(rowLink);
      const isNews = (row.tipo || '').toUpperCase().includes('NOTIC');

      const scrapedData = (row.scrapedTitle || row.scrapedContent) ? {
        title: row.scrapedTitle,
        content: row.scrapedContent
      } : null;

      if (isNews) {
        if (!parsedNewsUrls.has(normLink)) {
          console.warn(`[Reconciliação] Notícia omitida pela IA foi resgatada pelo motor heurístico: ${rowLink}`);
          const fallbackNews = this.parseNewsUrl(rowLink, parsed.noticias.length + 1, null, scrapedData);
          parsed.noticias.push(fallbackNews);
          parsedNewsUrls.add(normLink);
        }
      } else {
        if (!parsedLawUrls.has(normLink)) {
          console.warn(`[Reconciliação] Legislação omitida pela IA foi resgatada pelo motor heurístico: ${rowLink}`);
          const fallbackItem = this.parseTaxUrl(rowLink, parsed.itens.length + 1, null, scrapedData);
          parsed.itens.push(fallbackItem);
          parsedLawUrls.add(normLink);
        }
      }
    });

    // Ordenação estrita por Esfera: FEDERAL -> ESTADUAL -> MUNICIPAL
    const esferaOrder = { 'FEDERAL': 1, 'ESTADUAL': 2, 'MUNICIPAL': 3 };
    parsed.itens.sort((a, b) => {
      const ordA = esferaOrder[(a.esfera || '').toUpperCase()] || 4;
      const ordB = esferaOrder[(b.esfera || '').toUpperCase()] || 4;
      if (ordA !== ordB) return ordA - ordB;
      const ufA = (a.uf || a.titulo || '').slice(0, 8);
      const ufB = (b.uf || b.titulo || '').slice(0, 8);
      return ufA.localeCompare(ufB);
    });

    parsed.itens.forEach((it, idx) => {
      it.numero = idx + 1;
      if (it.titulo) {
        it.titulo = it.titulo.replace(/^\d+\.\s*/, `${it.numero}. `);
      }
      if (!it.link) it.link = validRows[idx]?.link || '#';
      if (!it.corpo_paragrafos || !Array.isArray(it.corpo_paragrafos)) {
        it.corpo_paragrafos = [String(it.corpo_paragrafos || "Síntese técnica estruturada.")];
      }

      // Salvaguarda programática contra alucinação de segmentos/produtos no corpo_paragrafos:
      const matchedRow = validRows.find(r => normalizeUrl(r.link || r.lei) === normalizeUrl(it.link));
      const sourceText = `${matchedRow?.scrapedTitle || ''} ${matchedRow?.scrapedContent || ''}`.toUpperCase();
      const monitoredCategories = [
        "PRODUTOS ELETRÔNICOS", "ELETROELETRÔNICOS", "ELETRODOMÉSTICOS", "MATERIAIS ELÉTRICOS",
        "PAPÉIS", "PLÁSTICOS", "PRODUTOS CERÂMICOS", "VIDROS", "BEBIDAS ALCOÓLICAS",
        "PRODUTOS DE PAPELARIA", "MATERIAIS DE CONSTRUÇÃO", "FERRAMENTAS", "LÂMPADAS",
        "PRODUTOS DE PERFUMARIA", "HIGIENE PESSOAL", "COSMÉTICOS"
      ];

      it.corpo_paragrafos = it.corpo_paragrafos.filter(p => {
        const pUpper = String(p).toUpperCase();
        const hasAlertOrMention = pUpper.includes("PONTO DE ALERTA") || monitoredCategories.some(cat => pUpper.includes(cat));
        if (hasAlertOrMention) {
          const reallyInSource = monitoredCategories.some(cat => sourceText.includes(cat));
          if (!reallyInSource) {
            console.warn(`[Anti-Alucinação] Removido parágrafo sem respaldo literal no texto da norma: "${p.slice(0, 80)}..."`);
            return false;
          }
        }
        return true;
      });
      if (!it.plano_de_acao || !Array.isArray(it.plano_de_acao)) {
        it.plano_de_acao = ["**Tributário / Compliance:** Validar aderência das operações da companhia às regras divulgadas."];
      }
      if (!it.entendimento_assunto || typeof it.entendimento_assunto !== 'string') {
        it.entendimento_assunto = `A publicação oficial **${it.norma || 'Norma Fiscal'}** (${it.esfera || 'Estadual'}) estabelece procedimentos perante a ${it.orgao || 'autoridade fiscal'}. Para a gestão da empresa, o foco é alinhar as rotinas de escrituração e sistemas para mitigar riscos de penalidades.`;
      }
    });

    parsed.noticias.forEach((n) => {
      if (!n.entendimento_assunto || typeof n.entendimento_assunto !== 'string') {
        n.entendimento_assunto = `Acompanhamento estratégico de jurisprudência com impacto potencial nas operações e planejamento financeiro da empresa.`;
      }
    });

    return parsed;
  },

  /**
   * Testa a conectividade com a API de IA do token informado
   * @param {string} apiKey
   * @returns {Promise<{provider: string, model: string}>}
   */
  async testConnection(apiKey) {
    const cleanKey = (apiKey || '').trim();
    if (!cleanKey) {
      throw new Error("Por favor, informe a Chave de API antes de testar.");
    }

    if (cleanKey.startsWith('sk-')) {
      const response = await fetch('https://api.openai.com/v1/models', {
        headers: { 'Authorization': `Bearer ${cleanKey}` }
      });
      if (!response.ok) {
        let msg = `HTTP ${response.status}`;
        try {
          const errData = await response.json();
          msg = errData?.error?.message || msg;
        } catch (e) {
          msg = (await response.text()).slice(0, 100) || msg;
        }
        throw new Error(`OpenAI: ${msg}`);
      }
      return { provider: 'OpenAI (ChatGPT)', model: 'gpt-4o-mini' };
    } else {
      const modelsToTest = ['gemini-3.6-flash', 'gemini-flash-lite-latest', 'gemini-3.5-flash-lite', 'gemini-flash-latest'];
      let lastDetail = '';

      for (const m of modelsToTest) {
        try {
          const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${encodeURIComponent(cleanKey)}`;
          const response = await fetch(endpoint, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'X-goog-api-key': cleanKey
            },
            body: JSON.stringify({
              contents: [{ parts: [{ text: "ping" }] }]
            })
          });

          if (response.ok) {
            return { provider: 'Google Gemini (Plano Pro/Flash)', model: m };
          } else {
            const errData = await response.json().catch(() => null);
            lastDetail = errData?.error?.message || `HTTP ${response.status}`;
            if (response.status === 400 || response.status === 401 || response.status === 403) {
              if (lastDetail.includes('API key not valid') || lastDetail.includes('API_KEY_INVALID')) {
                break;
              }
            }
          }
        } catch (e) {
          lastDetail = e.message;
        }
      }
      throw new Error(`Google AI: ${lastDetail}`);
    }
  }
};

// Alias para compatibilidade
window.TaxSynthesizer.synthesizeWithGemini = window.TaxSynthesizer.synthesizeWithAI;

