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
   * Extrai metadados completos de legislação a partir da URL oficial
   */
  parseTaxUrl(rawUrl, itemIndex) {
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
    // 8. Federal / CONFAZ / DOU
    else if (urlLower.includes("confaz") || urlLower.includes("in.gov.br") || urlLower.includes("fazenda.gov.br") || urlLower.includes("planalto.gov.br") || urlLower.includes("receita.fazenda.gov.br")) {
      uf = "BR";
      esfera = "FEDERAL";
      orgao = urlLower.includes("confaz") ? "CONFAZ" : (urlLower.includes("receita") ? "RFB" : "Governo Federal");
      impacto = urlLower.includes("confaz") ? "ICMS / Atos Interestaduais" : "Tributos Federais / Regulação";
      area = "Indiretos / Jurídico";

      if (urlLower.includes("convenio")) {
        const m = url.match(/convenio[_\s-]?icms[_\s-]?(\d+)[_\/](\d{2,4})/i) || url.match(/convenio[_\s-]?(\d+)/i);
        norma = m ? `Convênio ICMS nº ${m[1]}/${m[2] || '2026'}` : `Convênio ICMS nº ${itemIndex}/2026`;
      } else if (urlLower.includes("ajuste")) {
        const m = url.match(/ajuste[_\s-]?sinief[_\s-]?(\d+)[_\/](\d{2,4})/i) || url.match(/ajuste[_\s-]?(\d+)/i);
        norma = m ? `Ajuste SINIEF nº ${m[1]}/${m[2] || '2026'}` : `Ajuste SINIEF nº ${itemIndex}/2026`;
      } else {
        const m = url.match(/(?:decreto|lei|instrucao-normativa|portaria)[_\s-]?(\d+)/i);
        norma = m ? `Ato Federal nº ${m[1]}` : `Publicação Oficial DOU nº ${itemIndex}/2026`;
      }

      resumoTabela = `Ato normativo de alcance nacional disciplinado pelo ${orgao}.`;

      corpoParagrafos = [
        `O órgão **${orgao}** publicou a **${norma}**, versando sobre normas gerais e harmonização tributária em nível nacional.`,
        `A medida estabelece **diretrizes que impactam o comércio interestadual**, com desdobramentos diretos nas operações de faturamento e logística.`
      ];

      planoAcao = [
        `**Planejamento Tributário:** Avaliar impactos nas operações interestaduais e aderência aos acordos federativos.`,
        `**TI Fiscal:** Parametrizar regras de emissão e contingência no sistema emissor corporativo.`
      ];
    }
    // 9. Fallback Inteligente para URLs de Legislação Gerais
    else {
      // Extrai slug da URL
      let slug = "";
      try {
        const u = new URL(url);
        const pathParts = u.pathname.split('/').filter(Boolean);
        slug = pathParts[pathParts.length - 1] || "";
        slug = slug.replace(/\.[a-zA-Z0-9]+$/, ''); // remove .aspx, .html, etc.
        slug = decodeURIComponent(slug).replace(/[_-]+/g, ' ').trim();
      } catch (e) {
        slug = "";
      }

      norma = slug && slug.length > 3 ? slug.toUpperCase() : `Ato Normativo nº ${itemIndex}/2026`;
      resumoTabela = `Publicação oficial de regulamentação fiscal e procedimentos tributários.`;

      corpoParagrafos = [
        `O ato normativo publicado pelo órgão **${orgao}** estabelece alterações e regras operacionais relativas ao **${impacto}**, demandando estrita observância pelos contribuintes com operações na circunscrição.`,
        `A norma define **novo cronograma de cumprimento das obrigações**, impactando diretamente os procedimentos de emissão documental e a **conformidade na Escrituração Fiscal Digital (EFD)**.`
      ];

      planoAcao = [
        `**Sistemas / TI Fiscal:** Verificar se a tabela de códigos e parâmetros no ERP está atualizada com as diretrizes da **${orgao}** para evitar rejeições cadastrais.`,
        `**Escrituração / Tributário:** Validar a correta apuração e emissão dos documentos fiscais pertinentes às filiais da região.`,
        `**Alinhamento de Processos:** Divulgar as novas regras para as áreas operacionais correlatas (Logística, Comercial e Tesouraria).`
      ];
    }

    const titulo = `${itemIndex}. ${uf} - ${norma} - Publicação Oficial de ${dataPub}`;

    return {
      numero: itemIndex,
      esfera: esfera,
      titulo: titulo,
      data_publicacao: dataPub,
      norma: norma,
      link: url,
      orgao: orgao,
      resumo_tabela: resumoTabela,
      impacto: impacto,
      area_impactada: area,
      corpo_paragrafos: corpoParagrafos,
      plano_de_acao: planoAcao,
      vigencia: `Efeitos a partir de ${dataPub}.`
    };
  },

  /**
   * Extrai metadados completos de notícia tributária a partir da URL
   */
  parseNewsUrl(rawUrl, itemIndex) {
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

    // Extrai manchete do slug da URL
    let manchete = "";
    try {
      const u = new URL(url);
      const pathParts = u.pathname.split('/').filter(Boolean);
      let lastPart = pathParts[pathParts.length - 1] || "";
      lastPart = lastPart.replace(/\.[a-zA-Z0-9]+$/, '');
      lastPart = decodeURIComponent(lastPart).replace(/[_-]+/g, ' ').trim();
      if (lastPart.length > 5) {
        // Capitaliza as palavras
        manchete = lastPart.charAt(0).toUpperCase() + lastPart.slice(1);
      }
    } catch (e) {
      manchete = "";
    }

    if (!manchete || manchete.length < 5) {
      manchete = `Acompanhamento Jurídico-Tributário de Relevância nº ${itemIndex}`;
    }

    return {
      titulo: manchete,
      corpo_paragrafos: [
        `A matéria em destaque analisa **julgamentos, discussões regulatórias e tendências jurisprudenciais** de elevado impacto econômico para o setor varejista e operações interestaduais.`,
        `O monitoramento contínuo do tema é estratégico para mapear **riscos contingenciais**, orientar o posicionamento institucional da companhia e subsidiar decisões tributárias corporativas.`
      ],
      fonte: `${fonte} (${dataPub})`,
      link: url
    };
  },

  /**
   * Sintetizador Heurístico Local (Offline / Rápido / Sem custos de API)
   */
  synthesize(numeroBoletim, periodo, rawRows) {
    const itens = [];
    const noticias = [];

    (rawRows || []).forEach((row) => {
      const tipo = (row.tipo || 'LEI').toUpperCase();
      const rawLink = (row.link || '').trim();

      if (!rawLink && !row.lei) return;
      const targetUrl = rawLink || row.lei;

      if (tipo.includes('NOTIC')) {
        noticias.push(this.parseNewsUrl(targetUrl, noticias.length + 1));
      } else {
        itens.push(this.parseTaxUrl(targetUrl, itens.length + 1));
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
   * Sintetizador via API Google Gemini (Google AI Studio)
   * Realiza consulta à IA em tempo real quando o usuário configurou sua chave no site
   */
  async synthesizeWithGemini(apiKey, numeroBoletim, periodo, rawRows) {
    const validRows = (rawRows || []).filter(r => (r.link || '').trim() || (r.lei || '').trim());
    if (validRows.length === 0) {
      throw new Error("Nenhuma linha válida para enviar à IA.");
    }

    const itemsSummary = validRows.map((r, i) => 
      `${i + 1}. Tipo: ${r.tipo || 'Lei'} | URL: ${r.link || r.lei}`
    ).join('\n');

    const promptText = `
Você é o assistente sênior de inteligência tributária e compliance fiscal da Fast Shop.
Analise a seguinte lista de links de legislação fiscal e notícias tributárias para gerar o Boletim de Ementário Fiscal corporativo nº ${numeroBoletim} (${periodo}).

LINKS A PROCESSAR:
${itemsSummary}

REQUISITOS OBRIGATÓRIOS:
1. Para cada link classificado como 'Lei':
   - Identifique com exatidão o Nome da Norma (ex: 'Resposta à Consulta Tributária nº 32898/2025', 'Portaria SRE nº 35/2026', 'Decreto Estadual nº 68.000/2025', etc.).
   - Esfera: FEDERAL, ESTADUAL ou MUNICIPAL.
   - UF: sigla do estado (SP, CE, DF, RJ, ES, etc.) ou BR para Federal.
   - Órgão emissor: SEFAZ SP, SEFAZ CE, SEFAZ DF, CONFAZ, RFB, etc.
   - Impacto: ICMS, PIS/COFINS, ICMS-ST, ISS, etc.
   - Área impactada: Indiretos, Jurídico, TI Fiscal, Cadastros, etc.
   - Corpo: 2 parágrafos concisos com os pontos-chave em negrito (**destaque**).
   - Plano de Ação: de 2 a 3 ações práticas direcionadas para o varejo (TI Fiscal/ERP, Parametrização, Tributário/Compliance, Logística/Comercial).
   - Vigência e Resumo em 1 linha para tabela de abertura.
2. Para cada link classificado como 'Notícia':
   - Título claro da manchete jurídica.
   - 2 parágrafos resumindo a tese jurídica e riscos com pontos-chave em negrito (**destaque**).
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
      "esfera": "ESTADUAL",
      "titulo": "1. SP - Resposta à Consulta Tributária nº 32898, de 2025 - SEFAZ SP",
      "data_publicacao": "${new Date().toLocaleDateString('pt-BR')}",
      "norma": "Resposta à Consulta Tributária nº 32898/2025",
      "link": "URL_ORIGINAL",
      "orgao": "SEFAZ SP",
      "resumo_tabela": "Resumo em 1 frase para a tabela inicial.",
      "impacto": "ICMS / Consultas",
      "area_impactada": "Indiretos",
      "corpo_paragrafos": [
        "Parágrafo 1 com **destaque em negrito**...",
        "Parágrafo 2 com **destaque em negrito**..."
      ],
      "plano_de_acao": [
        "**TI Fiscal:** ...",
        "**Tributário:** ..."
      ],
      "vigencia": "Efeitos a partir da publicação."
    }
  ],
  "noticias": [
    {
      "titulo": "Título da Notícia",
      "corpo_paragrafos": [
        "Parágrafo 1 com **destaque**...",
        "Parágrafo 2 com **destaque**..."
      ],
      "fonte": "JOTA Tributário (06/09/2026)",
      "link": "URL_ORIGINAL"
    }
  ]
}
`;

    // Modelos oficiais do Google AI Studio suportados (em ordem de preferência)
    const modelsToTry = ['gemini-1.5-flash', 'gemini-2.0-flash', 'gemini-1.5-pro'];
    let lastError = null;
    let rawTextResponse = null;

    for (const modelName of modelsToTry) {
      try {
        const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${encodeURIComponent(apiKey.trim())}`;
        
        const response = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [
              {
                role: "user",
                parts: [{ text: promptText }]
              }
            ],
            generationConfig: {
              temperature: 0.2,
              responseMimeType: "application/json"
            }
          })
        });

        if (response.ok) {
          const data = await response.json();
          rawTextResponse = data?.candidates?.[0]?.content?.parts?.[0]?.text;
          if (rawTextResponse) {
            break; // Sucesso!
          }
        } else {
          const errText = await response.text();
          let msg = `HTTP ${response.status}`;
          try {
            const errJson = JSON.parse(errText);
            if (errJson.error && errJson.error.message) {
              msg = errJson.error.message;
            }
          } catch (e) {}
          lastError = new Error(msg);
          // Se for erro de autenticação (chave inválida), não adianta tentar outros modelos
          if (response.status === 400 && msg.includes('API_KEY_INVALID')) {
            throw new Error(`Chave do Gemini inválida (${msg}). Verifique sua chave no Google AI Studio.`);
          }
        }
      } catch (err) {
        lastError = err;
        if (err.message && err.message.includes('API_KEY_INVALID')) {
          throw err;
        }
      }
    }

    if (!rawTextResponse) {
      throw lastError || new Error("Não foi possível obter resposta dos servidores do Google Gemini.");
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

    parsed.itens.forEach((it, idx) => {
      it.numero = idx + 1;
      if (!it.link) it.link = validRows[idx]?.link || '#';
      if (!it.corpo_paragrafos || !Array.isArray(it.corpo_paragrafos)) {
        it.corpo_paragrafos = [String(it.corpo_paragrafos || "Síntese técnica estruturada.")];
      }
      if (!it.plano_de_acao || !Array.isArray(it.plano_de_acao)) {
        it.plano_de_acao = ["**Tributário / Compliance:** Validar aderência das operações da companhia às regras divulgadas."];
      }
    });

    return parsed;
  }
};

