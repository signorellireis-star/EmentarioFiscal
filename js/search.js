/**
 * EmpresaX - Ementário Fiscal
 * Motor de Busca Global e Filtragem Dinâmica
 */

window.SearchEngine = {
  /**
   * Constrói o índice linear de pesquisa a partir de uma lista de boletins
   */
  flattenDatabase(boletins) {
    const flatItems = [];

    (boletins || []).forEach(bol => {
      // Itens de Legislação
      (bol.itens || []).forEach((item, idx) => {
        flatItems.push({
          tipo_registro: 'LEI',
          item_index: idx,
          boletim_numero: bol.numero_boletim,
          boletim_periodo: bol.periodo,
          numero: item.numero,
          esfera: (item.esfera || 'ESTADUAL').toUpperCase(),
          titulo: item.titulo || '',
          data_publicacao: item.data_publicacao || '',
          norma: item.norma || '',
          link: item.link || '#',
          orgao: item.orgao || '',
          resumo_tabela: item.resumo_tabela || '',
          impacto: item.impacto || '',
          area_impactada: item.area_impactada || '',
          corpo_paragrafos: item.corpo_paragrafos || [],
          entendimento_assunto: item.entendimento_assunto || '',
          plano_de_acao: item.plano_de_acao || [],
          vigencia: item.vigencia || '',
          // Campo textual consolidado para busca rápida
          search_blob: [
            bol.numero_boletim,
            bol.periodo,
            item.titulo,
            item.norma,
            item.orgao,
            item.resumo_tabela,
            item.impacto,
            item.area_impactada,
            (item.corpo_paragrafos || []).join(' '),
            item.entendimento_assunto || '',
            (item.plano_de_acao || []).join(' '),
            item.vigencia
          ].join(' ').toLowerCase()
        });
      });

      // Itens de Notícia
      (bol.noticias || []).forEach((noticia, nIdx) => {
        flatItems.push({
          tipo_registro: 'NOTICIA',
          item_index: nIdx,
          boletim_numero: bol.numero_boletim,
          boletim_periodo: bol.periodo,
          numero: nIdx + 1,
          esfera: 'NOTÍCIA',
          titulo: noticia.titulo || '',
          data_publicacao: bol.periodo || '',
          norma: 'Notícia / Clipping',
          link: noticia.link || '#',
          orgao: noticia.fonte || 'Mídia Especializada',
          resumo_tabela: noticia.titulo || '',
          impacto: 'Conjuntura / Jurisprudência',
          area_impactada: 'Jurídico, Indiretos',
          corpo_paragrafos: noticia.corpo_paragrafos || [],
          entendimento_assunto: noticia.entendimento_assunto || '',
          plano_de_acao: noticia.plano_de_acao || [],
          vigencia: '',
          search_blob: [
            bol.numero_boletim,
            noticia.titulo,
            noticia.fonte,
            (noticia.corpo_paragrafos || []).join(' '),
            noticia.entendimento_assunto || ''
          ].join(' ').toLowerCase()
        });
      });
    });

    return flatItems;
  },

  /**
   * Converte marcação markdown simples (**negrito**) para HTML seguro
   */
  formatMarkdown(text) {
    if (!text) return '';
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/==(.+?)==/g, '<mark class="highlight-yellow" style="background-color: #ffeb3b; color: #1a1a1a; font-weight: bold; padding: 1px 4px; border-radius: 2px;">$1</mark>')
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>');
  },

  /**
   * Filtra os itens com base nos critérios da interface
   */
  filterItems(items, query, filters) {
    const q = (query || '').trim().toLowerCase();
    const tokens = q ? q.split(/\s+/).filter(Boolean) : [];

    return items.filter(item => {
      // 1. Busca textual global por todos os tokens digitados
      if (tokens.length > 0) {
        const matchesAll = tokens.every(tok => item.search_blob.includes(tok));
        if (!matchesAll) return false;
      }

      // 2. Filtro por Esfera
      if (filters.esfera && filters.esfera !== 'TODAS') {
        if (item.esfera !== filters.esfera) return false;
      }

      // 3. Filtro por Boletim
      if (filters.boletim && filters.boletim !== 'TODOS') {
        if (item.boletim_numero !== filters.boletim) return false;
      }

      // 4. Filtro por Órgão
      if (filters.orgao && filters.orgao !== 'TODOS') {
        if (item.orgao !== filters.orgao) return false;
      }

      // 5. Filtro por Tributo / Impacto
      if (filters.impacto && filters.impacto !== 'TODOS') {
        if (!item.impacto.toLowerCase().includes(filters.impacto.toLowerCase())) return false;
      }

      return true;
    });
  },

  /**
   * Renderiza a lista de cards com base nos itens filtrados
   */
  renderResults(containerEl, items, onSelectBoletimCallback) {
    if (!containerEl) return;

    if (!items || items.length === 0) {
      containerEl.innerHTML = `
        <div style="text-align: center; padding: 4rem 1rem; color: var(--text-muted); background: var(--bg-surface); border-radius: var(--radius-lg); border: 1px dashed var(--border-subtle);">
          <div style="font-size: 2.5rem; margin-bottom: 0.5rem; opacity: 0.6;">🔍</div>
          <h3 style="font-size: 1.1rem; color: var(--text-primary); margin-bottom: 0.25rem;">Nenhuma legislação encontrada</h3>
          <p style="font-size: 0.85rem;">Tente ajustar os termos da pesquisa ou redefinir os filtros aplicados.</p>
        </div>
      `;
      return;
    }

    const cardsHtml = items.map(item => {
      // Classes de Esfera
      let esfClass = 'esfera-estadual';
      let tagEsfClass = 'tag-est';
      if (item.esfera === 'FEDERAL') {
        esfClass = 'esfera-federal';
        tagEsfClass = 'tag-fed';
      } else if (item.esfera === 'MUNICIPAL') {
        esfClass = 'esfera-municipal';
        tagEsfClass = 'tag-mun';
      } else if (item.esfera === 'NOTÍCIA') {
        esfClass = 'esfera-noticia';
        tagEsfClass = 'tag-not';
      }

      // Parágrafos
      const paras = (item.corpo_paragrafos || [])
        .map(p => `<p>${this.formatMarkdown(p)}</p>`)
        .join('');

      // Entendimento do Assunto (Visão Executiva)
      let understandingHtml = '';
      if (item.entendimento_assunto) {
        understandingHtml = `
          <div class="card-understanding">
            <div class="card-understanding-title">
              <span>💡</span> Entendimento do Assunto (Visão Executiva):
            </div>
            <div class="card-understanding-text">
              ${this.formatMarkdown(item.entendimento_assunto)}
            </div>
          </div>
        `;
      }

      // Plano de Ação
      let actionPlanHtml = '';
      if (item.plano_de_acao && item.plano_de_acao.length > 0) {
        const lis = item.plano_de_acao
          .map(act => `<li>${this.formatMarkdown(act)}</li>`)
          .join('');
        actionPlanHtml = `
          <div class="card-action-plan">
            <div class="card-action-plan-title">
              <span>⚡</span> Plano de Ação Sugerido:
            </div>
            <ul class="card-action-plan-list">
              ${lis}
            </ul>
          </div>
        `;
      }

      // Vigência
      let vigHtml = '';
      if (item.vigencia) {
        vigHtml = `<span style="font-style: italic; color: var(--text-secondary);">${this.formatMarkdown(item.vigencia)}</span>`;
      }

      return `
        <article class="law-card ${esfClass}">
          <div class="card-top">
            <div class="card-tags">
              <span class="tag ${tagEsfClass}">${item.esfera}</span>
              <span class="tag tag-boletim">Boletim ${item.boletim_numero}</span>
              <span class="tag" style="background: rgba(255,255,255,0.05); color: var(--text-primary);">${item.impacto}</span>
            </div>
            <div class="card-date">${item.data_publicacao}</div>
          </div>

          <h3 class="card-title">${item.titulo}</h3>

          <table class="card-mini-table">
            <thead>
              <tr>
                <th style="width: 20%;">Norma</th>
                <th style="width: 15%;">Órgão</th>
                <th style="width: 45%;">Resumo Síntese</th>
                <th style="width: 20%;">Área Impactada</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><strong><a href="${item.link}" target="_blank" rel="noopener noreferrer">${item.norma} ↗</a></strong></td>
                <td>${item.orgao}</td>
                <td>${this.formatMarkdown(item.resumo_tabela)}</td>
                <td>${item.area_impactada}</td>
              </tr>
            </tbody>
          </table>

          <div class="card-body-text">
            ${paras}
          </div>

          ${understandingHtml}

          ${actionPlanHtml}

          <div class="card-footer">
            <div>
              ${vigHtml}
            </div>
            <div class="card-actions">
              <button class="btn-card-action" onclick="App.openEditItemModal('${item.tipo_registro}', '${item.boletim_numero}', ${item.item_index})" style="color: #818CF8; border-color: rgba(99, 102, 241, 0.4); font-weight: 600;">
                ✏️ Editar ${item.tipo_registro === 'NOTICIA' ? 'Notícia' : 'Item'}
              </button>
              <button class="btn-card-action" onclick="App.copyCardText('${encodeURIComponent(item.titulo)}', '${encodeURIComponent((item.corpo_paragrafos || []).join('\n\n'))}')">
                📋 Copiar Síntese
              </button>
              <button class="btn-card-action" onclick="App.viewInViewer('${item.boletim_numero}')">
                📄 Ver no Boletim Completo
              </button>
              <a href="${item.link}" target="_blank" rel="noopener noreferrer" class="btn-card-action" style="color: #38BDF8;">
                🔗 Acessar na Íntegra ↗
              </a>
            </div>
          </div>
        </article>
      `;
    }).join('');

    containerEl.innerHTML = cardsHtml;
  }
};
