/**
 * EmpresaX - Portal do Ementário Fiscal
 * Aplicação Principal (Estado, Navegação, Persistência e Renderização A4)
 */

window.App = {
  // Estado Global da Aplicação
  state: {
    boletins: [],
    flatItems: [],
    activeTab: 'search',
    activeBoletimNum: null,
    searchQuery: '',
    filters: {
      esfera: 'TODAS',
      boletim: 'TODOS',
      orgao: 'TODOS',
      impacto: 'TODOS'
    },
    // Estado do Gerador / Novo Boletim (Estilo Excel)
    gridRows: [
      { tipo: 'Lei', link: '' },
      { tipo: 'Lei', link: '' },
      { tipo: 'Lei', link: '' }
    ],
    generatedBoletim: null
  },

  // =========================================================================
  // CLIENTE DE BANCO EM NUVEM SUPABASE (SINCRONIZAÇÃO EM TEMPO REAL)
  // =========================================================================
  supabase: {
    url: 'https://pkognzyrwisqwoqzpeus.supabase.co',
    anonKey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBrb2duenlyd2lzcXdvcXpwZXVzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODg3MjA4MzYsImV4cCI6MjEwNDI5NjgzNn0.PlOltgppl3dJ4hddk8Rmjy7WiJgil4EsupKNwOgvras',

    async fetchBoletins() {
      try {
        const res = await fetch(`${this.url}/rest/v1/boletins?select=*&order=numero_boletim.desc`, {
          headers: {
            'apikey': this.anonKey,
            'Authorization': `Bearer ${this.anonKey}`
          }
        });
        if (res.ok) {
          const rows = await res.json();
          if (Array.isArray(rows) && rows.length > 0) {
            return rows.map(r => ({
              numero_boletim: r.numero_boletim,
              periodo: r.periodo || '',
              departamento: r.departamento || 'Fiscal',
              subtitulo: r.subtitulo || 'Ementário Fiscal',
              equipe: r.equipe || '',
              itens: r.itens || [],
              noticias: r.noticias || []
            }));
          }
        }
      } catch (err) {
        console.warn('[Supabase] Erro ao buscar dados em nuvem:', err);
      }
      return null;
    },

    async upsertBoletim(boletim) {
      if (!boletim || !boletim.numero_boletim) return;
      try {
        const payload = {
          numero_boletim: String(boletim.numero_boletim),
          periodo: boletim.periodo || '',
          departamento: boletim.departamento || 'Fiscal',
          subtitulo: boletim.subtitulo || 'Ementário Fiscal',
          equipe: boletim.equipe || '',
          itens: boletim.itens || [],
          noticias: boletim.noticias || [],
          updated_at: new Date().toISOString()
        };

        await fetch(`${this.url}/rest/v1/boletins`, {
          method: 'POST',
          headers: {
            'apikey': this.anonKey,
            'Authorization': `Bearer ${this.anonKey}`,
            'Content-Type': 'application/json',
            'Prefer': 'resolution=merge-duplicates'
          },
          body: JSON.stringify(payload)
        });
        console.log(`[Supabase] Boletim ${boletim.numero_boletim} sincronizado com a nuvem.`);
      } catch (err) {
        console.error('[Supabase] Erro ao salvar na nuvem:', err);
      }
    },

    async deleteBoletim(num) {
      if (!num) return;
      try {
        await fetch(`${this.url}/rest/v1/boletins?numero_boletim=eq.${encodeURIComponent(String(num))}`, {
          method: 'DELETE',
          headers: {
            'apikey': this.anonKey,
            'Authorization': `Bearer ${this.anonKey}`
          }
        });
        console.log(`[Supabase] Boletim ${num} excluído da nuvem.`);
      } catch (err) {
        console.error('[Supabase] Erro ao excluir na nuvem:', err);
      }
    }
  },

  updateCloudBadge(status, label) {
    const badge = document.getElementById('cloud-status-badge');
    const labelEl = document.getElementById('cloud-status-label');
    if (!badge) return;
    badge.className = `cloud-status-badge ${status}`;
    if (labelEl && label) labelEl.textContent = label;
  },

  async syncFromSupabase() {
    this.updateCloudBadge('syncing', 'Sincronizando...');
    const cloudBoletins = await this.supabase.fetchBoletins();
    if (cloudBoletins && cloudBoletins.length > 0) {
      this.state.boletins = cloudBoletins;
      localStorage.setItem('empresax_boletins_db', JSON.stringify(this.state.boletins));
      this.rebuildSearchIndex();
      this.updateMetrics();
      this.populateFilterDropdowns();
      this.runSearch();

      if (!this.state.activeBoletimNum || !this.state.boletins.some(b => b.numero_boletim === this.state.activeBoletimNum)) {
        this.state.activeBoletimNum = this.state.boletins[0].numero_boletim;
        this.renderViewer();
      }
      this.updateCloudBadge('', 'Nuvem Ativa');
    } else {
      this.updateCloudBadge('', 'Nuvem Ativa');
    }
  },

  /**
   * Inicialização da aplicação
   */
  async init() {
    this.loadDatabase();
    this.setupNavigation();
    this.setupSearchEvents();
    this.setupGeneratorEvents();
    this.setupDeleteEvents();
    this.setupThemeToggle();
    this.updateGeminiStatusIndicator();
    this.updateMetrics();
    this.populateFilterDropdowns();
    this.runSearch();

    // Seleciona o primeiro boletim para o visualizador por padrão
    if (this.state.boletins.length > 0) {
      this.state.activeBoletimNum = this.state.boletins[0].numero_boletim;
      this.renderViewer();
    }

    // Sincronização automática transparente com o Supabase em Nuvem
    this.syncFromSupabase();
  },

  /**
   * Carrega a base de dados do LocalStorage ou do arquivo inicial
   */
  loadDatabase() {
    const localData = localStorage.getItem('empresax_boletins_db') || localStorage.getItem('fastshop_boletins_db');
    if (localData) {
      try {
        this.state.boletins = JSON.parse(localData);
      } catch (e) {
        console.error("Erro ao carregar dados locais:", e);
      }
    }

    // Se não houver dados locais, utiliza os dados pré-carregados
    if (!this.state.boletins || this.state.boletins.length === 0) {
      if (window.INITIAL_BOLETINS_DB && window.INITIAL_BOLETINS_DB.length > 0) {
        this.state.boletins = JSON.parse(JSON.stringify(window.INITIAL_BOLETINS_DB));
        this.saveDatabase();
      }
    }

    // Normalização de conformidade: atualiza qualquer menção antiga a 'SEEC DF' para 'SEFAZ DF'
    let migrated = false;
    (this.state.boletins || []).forEach(b => {
      (b.itens || []).forEach(it => {
        if (it.orgao === 'SEEC DF') {
          it.orgao = 'SEFAZ DF';
          migrated = true;
        }
        if (!it.entendimento_assunto) {
          it.entendimento_assunto = `A norma **${it.norma || 'Fiscal'}** (${it.esfera || 'Estadual'}) estabelece diretrizes operacionais de conformidade perante a ${it.orgao || 'autoridade fazendária'}. Para a liderança e gestores, o foco principal é alinhar processos internos e rotinas de sistemas para mitigar riscos de penalidades.`;
          migrated = true;
        }
      });
      (b.noticias || []).forEach(n => {
        if (!n.entendimento_assunto) {
          n.entendimento_assunto = `Acompanhamento estratégico de jurisprudência com impacto potencial para tomada de decisão e planejamento corporativo.`;
          migrated = true;
        }
      });
    });
    if (migrated) {
      localStorage.setItem('empresax_boletins_db', JSON.stringify(this.state.boletins));
    }

    this.rebuildSearchIndex();
  },

  /**
   * Salva o estado atual no LocalStorage e no Supabase (Nuvem)
   */
  saveDatabase(targetBoletim = null) {
    localStorage.setItem('empresax_boletins_db', JSON.stringify(this.state.boletins));
    this.rebuildSearchIndex();
    this.updateMetrics();
    this.populateFilterDropdowns();

    // Sincronização em nuvem
    if (targetBoletim) {
      this.supabase.upsertBoletim(targetBoletim);
    } else {
      (this.state.boletins || []).forEach(b => this.supabase.upsertBoletim(b));
    }
  },

  /**
   * Reconstrói o índice linear para busca rápida
   */
  rebuildSearchIndex() {
    this.state.flatItems = SearchEngine.flattenDatabase(this.state.boletins);
  },

  /**
   * Controle de abas e navegação
   */
  setupNavigation() {
    const tabs = document.querySelectorAll('.nav-btn');
    tabs.forEach(btn => {
      btn.addEventListener('click', () => {
        const targetTab = btn.dataset.tab;
        this.switchTab(targetTab);
      });
    });
  },

  switchTab(tabId) {
    this.state.activeTab = tabId;
    document.querySelectorAll('.nav-btn').forEach(b => {
      b.classList.toggle('active', b.dataset.tab === tabId);
    });

    document.querySelectorAll('.tab-pane').forEach(p => {
      p.style.display = (p.id === `tab-${tabId}`) ? 'block' : 'none';
    });

    if (tabId === 'viewer') {
      this.renderViewer();
    } else if (tabId === 'metrics') {
      this.renderDetailedMetrics();
    }
  },

  /**
   * Configuração de Eventos de Busca
   */
  setupSearchEvents() {
    const searchInput = document.getElementById('search-input');
    const filterEsfera = document.getElementById('filter-esfera');
    const filterBoletim = document.getElementById('filter-boletim');
    const filterOrgao = document.getElementById('filter-orgao');
    const filterImpacto = document.getElementById('filter-impacto');

    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        this.state.searchQuery = e.target.value;
        this.runSearch();
      });
    }

    const onFilterChange = () => {
      this.state.filters.esfera = filterEsfera ? filterEsfera.value : 'TODAS';
      this.state.filters.boletim = filterBoletim ? filterBoletim.value : 'TODOS';
      this.state.filters.orgao = filterOrgao ? filterOrgao.value : 'TODOS';
      this.state.filters.impacto = filterImpacto ? filterImpacto.value : 'TODOS';
      this.runSearch();
    };

    if (filterEsfera) filterEsfera.addEventListener('change', onFilterChange);
    if (filterBoletim) filterBoletim.addEventListener('change', onFilterChange);
    if (filterOrgao) filterOrgao.addEventListener('change', onFilterChange);
    if (filterImpacto) filterImpacto.addEventListener('change', onFilterChange);

    // Chips de esfera
    document.querySelectorAll('.chip-btn').forEach(chip => {
      chip.addEventListener('click', () => {
        document.querySelectorAll('.chip-btn').forEach(c => c.classList.remove('active'));
        chip.classList.add('active');
        const esf = chip.dataset.esfera;
        if (filterEsfera) {
          filterEsfera.value = esf;
          this.state.filters.esfera = esf;
          this.runSearch();
        }
      });
    });
  },

  /**
   * Executa a busca e atualiza a listagem de cards
   */
  runSearch() {
    const filtered = SearchEngine.filterItems(
      this.state.flatItems,
      this.state.searchQuery,
      this.state.filters
    );

    const container = document.getElementById('results-grid');
    const countEl = document.getElementById('results-count');

    if (countEl) {
      countEl.textContent = `${filtered.length} norma(s) encontrada(s)`;
    }

    SearchEngine.renderResults(container, filtered);
  },

  /**
   * Preenche os menus suspensos de filtro com base nos dados reais
   */
  populateFilterDropdowns() {
    const selBol = document.getElementById('filter-boletim');
    const selViewerBol = document.getElementById('viewer-boletim-select');
    const selOrgao = document.getElementById('filter-orgao');
    const selImpacto = document.getElementById('filter-impacto');

    if (selBol) {
      const nums = this.state.boletins.map(b => b.numero_boletim);
      selBol.innerHTML = `<option value="TODOS">Todos os Boletins</option>` +
        nums.map(n => `<option value="${n}">Boletim ${n}</option>`).join('');
    }

    if (selViewerBol) {
      selViewerBol.innerHTML = this.state.boletins.map(b => 
        `<option value="${b.numero_boletim}" ${b.numero_boletim === this.state.activeBoletimNum ? 'selected' : ''}>Boletim ${b.numero_boletim} (${b.periodo})</option>`
      ).join('');
    }

    if (selOrgao) {
      const orgaos = [...new Set(this.state.flatItems.map(i => i.orgao).filter(Boolean))].sort();
      selOrgao.innerHTML = `<option value="TODOS">Todos os Órgãos</option>` +
        orgaos.map(o => `<option value="${o}">${o}</option>`).join('');
    }

    if (selImpacto) {
      const impactos = [...new Set(this.state.flatItems.map(i => i.impacto).filter(Boolean))].sort();
      selImpacto.innerHTML = `<option value="TODOS">Todos os Impactos</option>` +
        impactos.map(im => `<option value="${im}">${im}</option>`).join('');
    }
  },

  /**
   * Atualiza os cartões de métricas rápidas no topo
   */
  updateMetrics() {
    const totalBoletins = this.state.boletins.length;
    const totalItens = this.state.flatItems.filter(i => i.tipo_registro === 'LEI').length;
    const totalNoticias = this.state.flatItems.filter(i => i.tipo_registro === 'NOTICIA').length;
    const estCount = this.state.flatItems.filter(i => i.esfera === 'ESTADUAL').length;

    const elBol = document.getElementById('metric-total-boletins');
    const elItens = document.getElementById('metric-total-itens');
    const elEst = document.getElementById('metric-total-estaduais');
    const elNot = document.getElementById('metric-total-noticias');

    if (elBol) elBol.textContent = totalBoletins;
    if (elItens) elItens.textContent = totalItens;
    if (elEst) elEst.textContent = estCount;
    if (elNot) elNot.textContent = totalNoticias;
  },

  /**
   * Renderiza a página detalhada de métricas
   */
  renderDetailedMetrics() {
    const container = document.getElementById('metrics-detailed-content');
    if (!container) return;

    // Contagem por Esfera
    const esferaCounts = { FEDERAL: 0, ESTADUAL: 0, MUNICIPAL: 0, NOTICIA: 0 };
    this.state.flatItems.forEach(it => {
      if (it.tipo_registro === 'NOTICIA') esferaCounts.NOTICIA++;
      else esferaCounts[it.esfera] = (esferaCounts[it.esfera] || 0) + 1;
    });

    // Contagem por Órgão
    const orgaoCounts = {};
    this.state.flatItems.forEach(it => {
      if (it.orgao) {
        orgaoCounts[it.orgao] = (orgaoCounts[it.orgao] || 0) + 1;
      }
    });

    const topOrgaos = Object.entries(orgaoCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);

    container.innerHTML = `
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; margin-top: 1.5rem;">
        <div class="gen-box">
          <h3 style="margin-bottom: 1rem; font-family: var(--font-display);">Distribuição por Esfera Jurídica</h3>
          <table class="card-mini-table">
            <thead>
              <tr>
                <th>Esfera / Categoria</th>
                <th style="text-align: right;">Quantidade de Atos</th>
                <th style="text-align: right;">Participação (%)</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td><span class="tag tag-fed">FEDERAL</span></td>
                <td style="text-align: right;"><strong>${esferaCounts.FEDERAL}</strong></td>
                <td style="text-align: right;">${((esferaCounts.FEDERAL / (this.state.flatItems.length || 1)) * 100).toFixed(1)}%</td>
              </tr>
              <tr>
                <td><span class="tag tag-est">ESTADUAL</span></td>
                <td style="text-align: right;"><strong>${esferaCounts.ESTADUAL}</strong></td>
                <td style="text-align: right;">${((esferaCounts.ESTADUAL / (this.state.flatItems.length || 1)) * 100).toFixed(1)}%</td>
              </tr>
              <tr>
                <td><span class="tag tag-mun">MUNICIPAL</span></td>
                <td style="text-align: right;"><strong>${esferaCounts.MUNICIPAL}</strong></td>
                <td style="text-align: right;">${((esferaCounts.MUNICIPAL / (this.state.flatItems.length || 1)) * 100).toFixed(1)}%</td>
              </tr>
              <tr>
                <td><span class="tag tag-not">NOTÍCIAS</span></td>
                <td style="text-align: right;"><strong>${esferaCounts.NOTICIA}</strong></td>
                <td style="text-align: right;">${((esferaCounts.NOTICIA / (this.state.flatItems.length || 1)) * 100).toFixed(1)}%</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div class="gen-box">
          <h3 style="margin-bottom: 1rem; font-family: var(--font-display);">Top Órgãos Emissores</h3>
          <table class="card-mini-table">
            <thead>
              <tr>
                <th>Órgão / Ente Fiscal</th>
                <th style="text-align: right;">Total de Publicações</th>
              </tr>
            </thead>
            <tbody>
              ${topOrgaos.map(([org, count]) => `
                <tr>
                  <td><strong>${org}</strong></td>
                  <td style="text-align: right;">${count}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  },

  /**
   * Configuração de Eventos do Gerador de Boletins (Entrada estilo Excel)
   */
  setupGeneratorEvents() {
    const fileInput = document.getElementById('excel-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
          this.processExcelFile(e.target.files[0]);
        }
      });
    }

    this.renderGridRows();
  },

  /**
   * Processa arquivo Excel e carrega as linhas na grade
   */
  async processExcelFile(file) {
    try {
      const buffer = await file.arrayBuffer();
      const extracted = ExcelParser.extractGridRows(buffer);

      if (extracted.numero_boletim) {
        const numEl = document.getElementById('draft-numero');
        if (numEl) numEl.value = extracted.numero_boletim;
      }
      if (extracted.periodo) {
        const perEl = document.getElementById('draft-periodo');
        if (perEl) perEl.value = extracted.periodo;
      }

      if (extracted.rows && extracted.rows.length > 0) {
        this.state.gridRows = extracted.rows.map(r => ({
          tipo: r.tipo || 'Lei',
          link: r.link || r.lei || ''
        }));
        this.renderGridRows();
        alert(`Planilha carregada com sucesso! ${extracted.rows.length} publicações inseridas na grade.`);
      } else {
        alert("Nenhuma linha de publicação identificada na planilha.");
      }
    } catch (err) {
      console.error(err);
      alert("Erro ao ler planilha: " + err.message);
    }
  },

  /**
   * Renderiza a grade de linhas da planilha (Links e Classificação)
   */
  renderGridRows() {
    const tbody = document.getElementById('excel-grid-tbody');
    const badge = document.getElementById('grid-row-count-badge');
    if (!tbody) return;

    if (!this.state.gridRows || this.state.gridRows.length === 0) {
      this.state.gridRows = [{ tipo: 'Lei', link: '' }];
    }

    if (badge) {
      badge.textContent = `${this.state.gridRows.length} linha(s) configurada(s)`;
    }

    tbody.innerHTML = this.state.gridRows.map((row, idx) => `
      <tr data-index="${idx}">
        <td style="text-align: center; color: var(--text-muted); font-size: 0.8rem; font-weight: 600;">${idx + 1}</td>
        <td>
          <select class="grid-select" onchange="App.updateGridCell(${idx}, 'tipo', this.value)">
            <option value="Lei" ${row.tipo === 'Lei' ? 'selected' : ''}>⚖️ Legislação / Ato</option>
            <option value="Notícia" ${row.tipo === 'Notícia' ? 'selected' : ''}>📰 Notícia</option>
          </select>
        </td>
        <td>
          <input 
            type="text" 
            class="grid-input" 
            placeholder="Cole a URL ou Link oficial (ex: https://legislacao.fazenda.sp.gov.br/...)"
            value="${row.link ? row.link.replace(/"/g, '&quot;') : (row.lei ? row.lei.replace(/"/g, '&quot;') : '')}" 
            oninput="App.updateGridCell(${idx}, 'link', this.value)"
          >
        </td>
        <td style="text-align: center;">
          <button type="button" class="btn-remove-row" onclick="App.removeGridRow(${idx})" title="Remover linha">✕</button>
        </td>
      </tr>
    `).join('');
  },

  updateGridCell(index, field, value) {
    if (this.state.gridRows[index]) {
      this.state.gridRows[index][field] = value;
    }
  },

  addGridRow(data) {
    this.state.gridRows.push({
      tipo: data?.tipo || 'Lei',
      link: data?.link || data?.lei || ''
    });
    this.renderGridRows();
  },

  removeGridRow(index) {
    if (this.state.gridRows.length > 1) {
      this.state.gridRows.splice(index, 1);
    } else {
      this.state.gridRows = [{ tipo: 'Lei', link: '' }];
    }
    this.renderGridRows();
  },

  clearGridRows() {
    this.state.gridRows = [
      { tipo: 'Lei', link: '' },
      { tipo: 'Lei', link: '' },
      { tipo: 'Lei', link: '' }
    ];
    this.renderGridRows();

    // Limpa o seletor de arquivo de planilha e seu badge de arquivo, se houver
    const fileInput = document.getElementById('excel-file-input');
    if (fileInput) fileInput.value = '';
    const fileBadge = document.getElementById('upload-file-name-badge');
    if (fileBadge) fileBadge.textContent = '';

    // Oculta painel de pré-visualização gerada
    const panel = document.getElementById('generated-preview-panel');
    if (panel) panel.style.display = 'none';

    this.showToast("🗑️ Grade de publicações limpa com sucesso!", "info");
  },

  openBatchPasteModal() {
    const modal = document.getElementById('batch-paste-modal');
    if (modal) {
      modal.classList.add('active');
      const ta = document.getElementById('batch-paste-textarea');
      if (ta) {
        ta.value = '';
        setTimeout(() => ta.focus(), 150);
      }
    }
  },

  closeBatchPasteModal() {
    const modal = document.getElementById('batch-paste-modal');
    if (modal) modal.classList.remove('active');
  },

  applyBatchPaste() {
    const ta = document.getElementById('batch-paste-textarea');
    if (!ta || !ta.value.trim()) {
      alert("Nenhum conteúdo colado.");
      return;
    }

    const lines = ta.value.split('\n').map(l => l.trim()).filter(Boolean);
    const newRows = [];

    lines.forEach(line => {
      let parts = line.includes('\t') ? line.split('\t') : line.split(';');
      parts = parts.map(p => p.trim()).filter(Boolean);

      const linkCandidate = parts[0].startsWith('http') ? parts[0] : (parts[1] && parts[1].startsWith('http') ? parts[1] : parts[0]);
      const linkLower = linkCandidate.toLowerCase();
      const isNewsDomain = [
        'jota.info', 'conjur.com.br', 'valor.globo.com', 'valor.com.br',
        'migalhas.com.br', 'tributario.com.br', 'estadao.com.br', 'folha.uol.com.br'
      ].some(d => linkLower.includes(d));

      if (parts.length === 1) {
        newRows.push({
          tipo: isNewsDomain ? 'Notícia' : 'Lei',
          link: linkCandidate
        });
      } else if (parts.length >= 2) {
        const isNoticia = parts[1].toLowerCase().includes('notic') || parts[0].toLowerCase().includes('notic') || isNewsDomain;
        newRows.push({
          tipo: isNoticia ? 'Notícia' : 'Lei',
          link: linkCandidate
        });
      }
    });

    if (newRows.length > 0) {
      this.state.gridRows = this.state.gridRows.filter(r => (r.link || '').trim());
      this.state.gridRows.push(...newRows);
      this.renderGridRows();
      this.closeBatchPasteModal();
      alert(`${newRows.length} linha(s) adicionada(s) à grade!`);
    }
  },

  syncGridFromDOM() {
    const tbody = document.getElementById('excel-grid-tbody');
    if (!tbody) return;
    const trList = tbody.querySelectorAll('tr');
    if (trList.length === 0) return;

    const rows = [];
    trList.forEach((tr, idx) => {
      const select = tr.querySelector('.grid-select');
      const input = tr.querySelector('.grid-input');
      const tipo = select ? select.value : (this.state.gridRows[idx]?.tipo || 'Lei');
      const link = input ? input.value.trim() : (this.state.gridRows[idx]?.link || '');
      rows.push({ tipo, link });
    });
    if (rows.length > 0) {
      this.state.gridRows = rows;
    }
  },

  showProcessingOverlay(options) {
    const modal = document.getElementById('processing-overlay');
    if (!modal) return;

    const titleEl = document.getElementById('processing-title');
    const subEl = document.getElementById('processing-subtitle');
    const s2Text = document.getElementById('proc-step-2-text');

    if (titleEl && options?.title) titleEl.textContent = options.title;
    if (subEl && options?.subtitle) subEl.textContent = options.subtitle;
    if (s2Text && options?.step2) s2Text.textContent = options.step2;

    const icon1 = document.getElementById('proc-icon-1');
    const icon2 = document.getElementById('proc-icon-2');
    const icon3 = document.getElementById('proc-icon-3');

    if (icon1) icon1.textContent = '✓';
    if (icon2) icon2.textContent = '⏳';
    if (icon3) icon3.textContent = '○';

    modal.classList.add('active');
  },

  updateProcessingStep(stepNum, text, isDone) {
    const icon = document.getElementById(`proc-icon-${stepNum}`);
    const row = document.getElementById(`proc-step-${stepNum}`);
    if (icon) {
      icon.textContent = isDone ? '✓' : '⏳';
      icon.style.color = isDone ? '#10B981' : '#818CF8';
    }
    if (row && text) {
      const span = row.querySelector('span:last-child');
      if (span) span.textContent = text;
    }
  },

  hideProcessingOverlay() {
    const modal = document.getElementById('processing-overlay');
    if (modal) modal.classList.remove('active');
  },

  /**
   * Mescla e agrupa um boletim existente com novas publicações processadas,
   * garantindo a ordenação institucional por Esfera (FEDERAL, ESTADUAL, MUNICIPAL)
   * e a renumeração sequencial contínua sem duplicação de links.
   */
  mergeAndSortBoletim(existingBol, newGenerated) {
    if (!existingBol) return newGenerated;
    if (!newGenerated) return existingBol;

    const normalizeUrl = (u) => (u || '').trim().toLowerCase().replace(/\/+$/, '');

    // 1. Legislações
    const mergedItens = [];
    const seenLinks = new Set();

    // Itens preexistentes do boletim
    (existingBol.itens || []).forEach(it => {
      const norm = normalizeUrl(it.link);
      if (norm && !seenLinks.has(norm)) {
        seenLinks.add(norm);
        mergedItens.push(JSON.parse(JSON.stringify(it)));
      }
    });

    // Novos itens processados
    (newGenerated.itens || []).forEach(it => {
      const norm = normalizeUrl(it.link);
      if (norm && !seenLinks.has(norm)) {
        seenLinks.add(norm);
        mergedItens.push(JSON.parse(JSON.stringify(it)));
      }
    });

    // Ordenação estrita por Esfera: FEDERAL -> ESTADUAL -> MUNICIPAL
    const esferaOrder = { 'FEDERAL': 1, 'ESTADUAL': 2, 'MUNICIPAL': 3 };
    mergedItens.sort((a, b) => {
      const ordA = esferaOrder[(a.esfera || '').toUpperCase()] || 4;
      const ordB = esferaOrder[(b.esfera || '').toUpperCase()] || 4;
      if (ordA !== ordB) return ordA - ordB;
      const ufA = (a.uf || a.titulo || '').slice(0, 8);
      const ufB = (b.uf || b.titulo || '').slice(0, 8);
      return ufA.localeCompare(ufB);
    });

    // Renumera sequencialmente 1..N e ajusta o título
    mergedItens.forEach((it, idx) => {
      it.numero = idx + 1;
      if (it.titulo) {
        it.titulo = it.titulo.replace(/^\d+\.\s*/, `${it.numero}. `);
      }
    });

    // 2. Notícias
    const mergedNoticias = [];
    const seenNews = new Set();

    (existingBol.noticias || []).forEach(n => {
      const norm = normalizeUrl(n.link) || (n.titulo || '').toLowerCase();
      if (norm && !seenNews.has(norm)) {
        seenNews.add(norm);
        mergedNoticias.push(JSON.parse(JSON.stringify(n)));
      }
    });

    (newGenerated.noticias || []).forEach(n => {
      const norm = normalizeUrl(n.link) || (n.titulo || '').toLowerCase();
      if (norm && !seenNews.has(norm)) {
        seenNews.add(norm);
        mergedNoticias.push(JSON.parse(JSON.stringify(n)));
      }
    });

    return {
      numero_boletim: newGenerated.numero_boletim || existingBol.numero_boletim,
      periodo: newGenerated.periodo || existingBol.periodo,
      departamento: newGenerated.departamento || existingBol.departamento || 'Fiscal',
      subtitulo: newGenerated.subtitulo || existingBol.subtitulo || 'Ementário Fiscal',
      equipe: newGenerated.equipe || existingBol.equipe || 'Boletim Fiscal elaborado pelo time de Planejamento Fiscal: Andréa Celi Mantovani, Antônio Sergio da Silva, Cristiane Cunha, Emerson de Deus e Raquel Capelão. Em caso de dúvidas, favor enviar e-mail para planejamentofiscal@empresax.com.br',
      itens: mergedItens,
      noticias: mergedNoticias
    };
  },

  /**
   * Processamento automatizado da grade:
   * Determina automaticamente a norma, órgão, esfera e elabora ementa com destaques e planos de ação
   * Exibe tela animada de 'Pesquisando...' e suporta IA Gemini em tempo real com fallback automático
   */
  async processGridData(forceWithoutGemini = false) {
    // 1. Sincroniza valores digitados diretamente da tabela DOM
    this.syncGridFromDOM();

    let num = (document.getElementById('draft-numero')?.value || '').trim();
    let per = (document.getElementById('draft-periodo')?.value || '').trim();

    // Valores padrão inteligentes caso o usuário não tenha preenchido
    if (!num) {
      num = "1.2026";
      const numInput = document.getElementById('draft-numero');
      if (numInput) numInput.value = num;
    }
    if (!per) {
      per = "Semana Atual";
      const perInput = document.getElementById('draft-periodo');
      if (perInput) perInput.value = per;
    }

    const validRows = this.state.gridRows.filter(r => (r.link || '').trim() || (r.lei || '').trim());
    if (validRows.length === 0) {
      alert("Por favor, cole pelo menos um link oficial na tabela antes de clicar em Processar.");
      return;
    }

    const geminiKey = (localStorage.getItem('empresax_gemini_api_key') || localStorage.getItem('fastshop_gemini_api_key') || '').trim();
    const hasGemini = geminiKey.length > 10;

    // Se o usuário não possui a chave configurada e não optou explicitamente por prosseguir manual, alerta com o modal explicativo
    if (!hasGemini && !forceWithoutGemini) {
      this.openGeminiMissingModal();
      return;
    }

    // 2. Abre a tela de 'Pesquisando & Analisando...'
    this.showProcessingOverlay({
      title: "Pesquisando & Analisando Legislação...",
      subtitle: "Acessando páginas oficiais e lendo o conteúdo dos atos normativos em tempo real...",
      step2: hasGemini 
        ? "Consultando IA Gemini com o conteúdo real da página..." 
        : "Decodificando órgãos oficiais e normas tributárias..."
    });

    // Etapa 1: Webscraping em tempo real de cada link oficial
    const enrichedRows = [];
    for (let i = 0; i < validRows.length; i++) {
      const r = validRows[i];
      const link = (r.link || r.lei || '').trim();
      this.updateProcessingStep(1, `Acessando e lendo link ${i + 1} de ${validRows.length}...`, false);
      const scraped = await TaxSynthesizer.scrapeUrl(link);
      enrichedRows.push({
        ...r,
        scrapedTitle: scraped?.title || "",
        scrapedContent: scraped?.content || ""
      });
    }
    this.updateProcessingStep(1, "Publicações oficiais lidas e conteúdos extraídos com sucesso!", true);

    // Pausa técnica para permitir animação fluida da interface
    await new Promise(r => setTimeout(r, 400));

    let generated = null;

    if (hasGemini) {
      try {
        this.updateProcessingStep(2, "Consultando Inteligência Artificial com os dados oficiais da página...", false);
        generated = await TaxSynthesizer.synthesizeWithAI(geminiKey, num, per, enrichedRows);
        this.updateProcessingStep(2, "IA concluiu a síntese com sucesso!", true);
      } catch (geminiErr) {
        console.warn("Aviso na chamada à API de IA:", geminiErr);
        this.updateProcessingStep(2, `Aviso: Falha na IA. Ativando Motor Heurístico...`, true);
        await new Promise(r => setTimeout(r, 700));

        alert(`⚠️ Aviso sobre a Inteligência Artificial:\n\nA consulta à IA retornou uma falha:\n${geminiErr.message}\n\nO boletim foi estruturado com sucesso a partir do conteúdo real extraído da página. Você pode complementar ou ajustar os dados clicando em "✏️ Editar".`);

        generated = TaxSynthesizer.synthesize(num, per, enrichedRows, geminiErr.message);
      }
    } else {
      this.updateProcessingStep(2, "Motor Heurístico estruturou a publicação com base na leitura real da página!", true);
      generated = TaxSynthesizer.synthesize(num, per, enrichedRows);
    }

    this.updateProcessingStep(3, "Ementário estruturado e diagramado com sucesso!", true);
    await new Promise(r => setTimeout(r, 500));

    this.hideProcessingOverlay();

    if (!generated) {
      alert("Não foi possível gerar o boletim. Verifique os links informados.");
      return;
    }

    // =========================================================================
    // AGRUPAMENTO AUTOMÁTICO POR NÚMERO DE BOLETIM E PERÍODO
    // =========================================================================
    const existingIndex = this.state.boletins.findIndex(b => 
      (b.numero_boletim || '').trim().toLowerCase() === String(num).trim().toLowerCase()
    );

    if (existingIndex >= 0) {
      const existingBol = this.state.boletins[existingIndex];
      const prevCount = (existingBol.itens || []).length;
      const prevNewsCount = (existingBol.noticias || []).length;

      generated = this.mergeAndSortBoletim(existingBol, generated);

      alert(`🔄 Agrupamento Realizado com Sucesso!\n\nFoi identificado que o Boletim nº ${num} já existe no repositório.\nOs novos links foram agrupados ao boletim existente mantendo a ordenação institucional por Esfera:\n1º FEDERAL\n2º ESTADUAL\n3º MUNICIPAL\n\nTotal consolidado: ${generated.itens.length} Legislações e ${generated.noticias.length} Notícia(s).`);
    }

    this.state.generatedBoletim = generated;

    // Renderiza a pré-visualização estruturada dos cards
    const panel = document.getElementById('generated-preview-panel');
    const container = document.getElementById('generated-preview-cards');

    if (panel && container) {
      panel.style.display = 'block';
      const flat = SearchEngine.flattenDatabase([generated]);
      SearchEngine.renderResults(container, flat);
      panel.scrollIntoView({ behavior: 'smooth' });
    }
  },

  /**
   * Salva o boletim processado no repositório geral e abre o visualizador A4
   */
  saveGeneratedToRepository() {
    if (!this.state.generatedBoletim) {
      alert("Nenhum boletim processado para salvar.");
      return;
    }

    const num = this.state.generatedBoletim.numero_boletim;
    const existingIndex = this.state.boletins.findIndex(b => 
      (b.numero_boletim || '').trim().toLowerCase() === String(num).trim().toLowerCase()
    );

    let finalBoletim = JSON.parse(JSON.stringify(this.state.generatedBoletim));

    if (existingIndex >= 0) {
      finalBoletim = this.mergeAndSortBoletim(this.state.boletins[existingIndex], finalBoletim);
      this.state.boletins[existingIndex] = finalBoletim;
    } else {
      this.state.boletins.unshift(finalBoletim);
    }

    this.saveDatabase(finalBoletim);
    this.state.activeBoletimNum = finalBoletim.numero_boletim;
    this.switchTab('viewer');
    alert(`Boletim ${num} consolidado e salvo com sucesso no Repositório! Todas as páginas diagramadas com hiperlinks estão disponíveis.`);
  },

  /**
   * Baixa o JSON do boletim atual ou gerado
   */
  downloadDraftJson() {
    const target = this.state.generatedBoletim || this.state.boletins.find(b => b.numero_boletim === this.state.activeBoletimNum) || this.state.boletins[0];
    const num = target ? String(target.numero_boletim || '').trim() : 'boletim';
    const per = target && target.periodo ? String(target.periodo || '').trim().replace(/[\/\\:*?"<>|]/g, '.') : '';
    const filename = per ? `Ementário - Boletim - ${num} - ${per}.json` : `Ementário - Boletim - ${num}.json`;
    const jsonStr = JSON.stringify(target, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  },

  /**
   * Exporta todo o banco de dados em formato JSON (Backup)
   */
  exportFullDatabase() {
    const jsonStr = JSON.stringify(this.state.boletins, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `empresax_ementario_db_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  },

  /**
   * Importa um arquivo JSON de backup do banco de dados
   */
  importFullDatabase() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = async (e) => {
      if (e.target.files.length > 0) {
        try {
          const text = await e.target.files[0].text();
          const parsed = JSON.parse(text);
          if (Array.isArray(parsed)) {
            this.state.boletins = parsed;
            this.saveDatabase();
            this.runSearch();
            alert(`Base importada com sucesso! ${parsed.length} boletins carregados.`);
          } else {
            alert("O arquivo não possui o formato de lista de boletins esperado.");
          }
        } catch (err) {
          alert("Erro ao importar base: " + err.message);
        }
      }
    };
    input.click();
  },

  /**
   * Redireciona a visualização para o boletim selecionado no repositório
   */
  viewInViewer(numeroBoletim) {
    this.state.activeBoletimNum = numeroBoletim;
    const sel = document.getElementById('viewer-boletim-select');
    if (sel) sel.value = numeroBoletim;
    this.switchTab('viewer');
  },

  /**
   * Copia o texto da ementa para a área de transferência
   */
  copyCardText(encodedTitle, encodedBody) {
    const title = decodeURIComponent(encodedTitle);
    const body = decodeURIComponent(encodedBody);
    const fullText = `*${title}*\n\n${body}`;
    navigator.clipboard.writeText(fullText).then(() => {
      alert("Síntese fiscal copiada com sucesso para a área de transferência!");
    });
  },

  /**
   * Alterna entre modo escuro e claro
   */
  setupThemeToggle() {
    const btn = document.getElementById('btn-theme-toggle');
    if (btn) {
      btn.addEventListener('click', () => {
        document.body.classList.toggle('light-theme');
        const isLight = document.body.classList.contains('light-theme');
        btn.innerHTML = isLight ? '🌙' : '☀️';
        localStorage.setItem('empresax_theme', isLight ? 'light' : 'dark');
      });
    }

    if ((localStorage.getItem('empresax_theme') || localStorage.getItem('fastshop_theme')) === 'light') {
      document.body.classList.add('light-theme');
      if (btn) btn.innerHTML = '🌙';
    }
  },

  /**
   * Renderização do Visualizador Institucional A4 (Exatamente como o Word / PDF Oficial)
   */
  renderViewer() {
    const container = document.getElementById('pdf-preview-container');
    const select = document.getElementById('viewer-boletim-select');
    if (!container) return;

    const num = select ? select.value : this.state.activeBoletimNum;
    const bol = this.state.boletins.find(b => b.numero_boletim === num) || this.state.boletins[0];

    if (!bol) {
      container.innerHTML = `<p style="padding: 3rem; text-align: center;">Nenhum boletim selecionado.</p>`;
      return;
    }

    // Configura o título do documento para o padrão oficial ao salvar em PDF
    const numStr = (bol.numero_boletim || '').trim();
    const perStr = (bol.periodo || '').trim().replace(/[\/\\:*?"<>|]/g, '.');
    document.title = perStr ? `Ementário - Boletim - ${numStr} - ${perStr}` : `Ementário - Boletim - ${numStr}`;

    const logoUri = window.EMPRESAX_LOGO_URI || '';
    const logoHtml = logoUri 
      ? `<img src="${logoUri}" alt="EmpresaX">` 
      : `<span class="header-brand-name" style="font-size: 15pt; font-weight: 800; color: #111827; letter-spacing: -0.5px; font-family: sans-serif;">EmpresaX</span>`;

    const renderHeader = () => `
      <table class="header-table">
        <tr>
          <td class="header-logo-cell">
            ${logoHtml}
          </td>
          <td class="header-title-cell">Boletim Fiscal</td>
          <td class="header-info-cell">
            Número Boletim<br>
            <strong>${bol.numero_boletim}</strong>
          </td>
        </tr>
        <tr>
          <td class="header-dept-cell">${bol.departamento || 'Fiscal'}</td>
          <td class="header-subtitle-cell">${bol.subtitulo || 'Ementário Fiscal'}</td>
          <td class="header-date-cell">${bol.periodo}</td>
        </tr>
      </table>
    `;

    const sheetsHtml = [];

    // --- PÁGINA 1: SUMÁRIO ---
    const tocRows = [];
    let currentEsfera = null;

    (bol.itens || []).forEach((item, idx) => {
      const pNum = idx + 2;
      const esf = (item.esfera || 'ESTADUAL').toUpperCase();
      if (esf !== currentEsfera) {
        currentEsfera = esf;
        tocRows.push(`
          <div class="toc-category">
            <a href="#sheet-${pNum}" class="toc-link">
              <span>${esf}</span>
              <span class="toc-dots"></span>
              <span class="toc-page">${pNum}</span>
            </a>
          </div>
        `);
      }

      // Limpa título para o sumário
      let cleanTitle = item.titulo;
      if (cleanTitle.startsWith(`${item.numero}. `)) {
        cleanTitle = cleanTitle.substring(`${item.numero}. `.length);
      }

      tocRows.push(`
        <div class="toc-item">
          <a href="#sheet-${pNum}" class="toc-link">
            <span class="toc-num">${item.numero}.</span>
            <span class="toc-text">${cleanTitle}</span>
            <span class="toc-dots"></span>
            <span class="toc-page">${pNum}</span>
          </a>
        </div>
      `);
    });

    if (bol.noticias && bol.noticias.length > 0) {
      const notiPageStart = (bol.itens || []).length + 2;
      tocRows.push(`
        <div class="toc-category">
          <a href="#sheet-${notiPageStart}" class="toc-link">
            <span>NOTÍCIAS</span>
            <span class="toc-dots"></span>
            <span class="toc-page">${notiPageStart}</span>
          </a>
        </div>
      `);

      bol.noticias.forEach((noticia, nIdx) => {
        const notiPNum = notiPageStart + nIdx;
        let notiTitle = noticia.titulo || `Matéria Jurídico-Tributária ${nIdx + 1}`;
        tocRows.push(`
          <div class="toc-item">
            <a href="#sheet-${notiPNum}" class="toc-link">
              <span class="toc-num">${nIdx + 1}.</span>
              <span class="toc-text">${notiTitle}</span>
              <span class="toc-dots"></span>
              <span class="toc-page">${notiPNum}</span>
            </a>
          </div>
        `);
      });
    }

    const p1 = `
      <section class="a4-sheet" id="sheet-1">
        <div>
          ${renderHeader()}
          <div class="sumario-title">Sumário</div>
          <div>${tocRows.join('')}</div>
        </div>
        <div class="footer">
          <div class="footer-disclaimer">${bol.equipe}</div>
        </div>
      </section>
    `;
    sheetsHtml.push(p1);

    // --- PÁGINAS 2..N: LEGISLAÇÕES ---
    let lastRenderedCategory = null;
    (bol.itens || []).forEach((item, idx) => {
      const pNum = idx + 2;
      const esf = (item.esfera || 'ESTADUAL').toUpperCase();
      let catHeader = '';
      if (esf !== lastRenderedCategory) {
        lastRenderedCategory = esf;
        catHeader = `<div class="section-title">${esf}</div>`;
      }

      const paras = (item.corpo_paragrafos || [])
        .map(p => `<p>${SearchEngine.formatMarkdown(p)}</p>`)
        .join('');

      let understandingBox = '';
      if (item.entendimento_assunto) {
        understandingBox = `
          <div class="pdf-understanding-box">
            <div class="pdf-understanding-title">💡 Entendimento do Assunto (Visão Executiva):</div>
            <div class="pdf-understanding-text">${SearchEngine.formatMarkdown(item.entendimento_assunto)}</div>
          </div>
        `;
      }

      let actionPlan = '';
      if (item.plano_de_acao && item.plano_de_acao.length > 0) {
        const lis = item.plano_de_acao
          .map(a => `<li>${SearchEngine.formatMarkdown(a)}</li>`)
          .join('');
        actionPlan = `
          <div class="action-plan">
            <div class="action-plan-title">Plano de Ação Sugerido:</div>
            <ul class="action-plan-list">
              ${lis}
            </ul>
          </div>
        `;
      }

      const vig = item.vigencia ? `<p class="effects-note">${SearchEngine.formatMarkdown(item.vigencia)}</p>` : '';

      const pItem = `
        <section class="a4-sheet" id="sheet-${pNum}">
          <div>
            ${renderHeader()}
            ${catHeader}
            <div class="item-title">${item.titulo}</div>

            <table class="info-table">
              <tr>
                <th class="col-data">DATA PUBLICAÇÃO</th>
                <th class="col-norma">NORMA</th>
                <th class="col-orgao">ÓRGÃO</th>
                <th class="col-resumo">RESUMO</th>
                <th class="col-impacto">IMPACTO</th>
                <th class="col-area">ÁREA IMPACTADA</th>
              </tr>
              <tr>
                <td class="col-data">${item.data_publicacao}</td>
                <td class="col-norma"><a href="${item.link}" target="_blank">${item.norma}</a></td>
                <td class="col-orgao">${item.orgao}</td>
                <td class="col-resumo">${SearchEngine.formatMarkdown(item.resumo_tabela)}</td>
                <td class="col-impacto">${item.impacto}</td>
                <td class="col-area">${item.area_impactada}</td>
              </tr>
            </table>

            <div class="content-body">
              ${paras}
              ${understandingBox}
              ${actionPlan}
              ${vig}
            </div>
          </div>

          <div class="footer">
            <div class="footer-page-num">${pNum}</div>
            <div class="footer-instruction">
              Para acesso a legislação na íntegra, basta clicar com o botão esquerdo do mouse sobre o campo “Norma” na tabela de informações.
            </div>
            <div class="footer-disclaimer">${bol.equipe}</div>
          </div>
        </section>
      `;
      sheetsHtml.push(pItem);
    });

    // --- PÁGINAS DE NOTÍCIAS ---
    let notiPageNum = (bol.itens || []).length + 2;
    (bol.noticias || []).forEach(noticia => {
      const paras = (noticia.corpo_paragrafos || [])
        .map(p => `<p>${SearchEngine.formatMarkdown(p)}</p>`)
        .join('');

      let notiUnderstanding = '';
      if (noticia.entendimento_assunto) {
        notiUnderstanding = `
          <div class="pdf-understanding-box">
            <div class="pdf-understanding-title">💡 Entendimento do Assunto (Visão Executiva):</div>
            <div class="pdf-understanding-text">${SearchEngine.formatMarkdown(noticia.entendimento_assunto)}</div>
          </div>
        `;
      }

      const pNoti = `
        <section class="a4-sheet" id="sheet-${notiPageNum}">
          <div>
            ${renderHeader()}
            <div class="section-title">NOTÍCIAS</div>
            <div class="item-title" style="font-size: 13pt; text-align: left; margin: 20pt 0 15pt 0;">${noticia.titulo}</div>

            <div class="content-body">
              ${paras}
              ${notiUnderstanding}
              <div style="margin-top: 25pt; font-size: 11pt;">
                <strong>Fonte:</strong> ${noticia.fonte || 'Clipping Tributário'}<br>
                <strong>Link:</strong> <a href="${noticia.link}" target="_blank" style="color: #0563c1;">${noticia.link}</a>
              </div>
            </div>
          </div>

          <div class="footer">
            <div class="footer-page-num">${notiPageNum}</div>
            <div class="footer-instruction">
              Para acesso a legislação na íntegra, basta clicar com o botão esquerdo do mouse sobre o campo “Norma” na tabela de informações.
            </div>
            <div class="footer-disclaimer">${bol.equipe}</div>
          </div>
        </section>
      `;
      sheetsHtml.push(pNoti);
      notiPageNum++;
    });

    container.innerHTML = sheetsHtml.join('');
  },

  /**
   * Modal de Adição Manual de Legislações / Notícias


  /**
   * Configuração de Eventos de Exclusão de Boletim
   */
  setupDeleteEvents() {
    const input = document.getElementById('delete-confirm-input');
    const btnConfirm = document.getElementById('btn-confirm-delete');
    if (input && btnConfirm) {
      input.addEventListener('input', (e) => {
        const val = (e.target.value || '').trim().toUpperCase();
        btnConfirm.disabled = (val !== 'EXCLUIR');
      });

      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !btnConfirm.disabled) {
          this.confirmDeleteBoletim();
        }
      });
    }
  },

  /**
   * Abre o modal de confirmação de exclusão
   */
  openDeleteModal(numeroBoletim) {
    const target = numeroBoletim || this.state.activeBoletimNum;
    if (!target) {
      alert("Nenhum boletim selecionado para exclusão.");
      return;
    }

    this.state.targetDeleteBoletim = target;
    const titleEl = document.getElementById('delete-modal-boletim-title');
    if (titleEl) titleEl.textContent = `Boletim ${target}`;

    const input = document.getElementById('delete-confirm-input');
    const btn = document.getElementById('btn-confirm-delete');
    if (input) input.value = '';
    if (btn) btn.disabled = true;

    const modal = document.getElementById('delete-boletim-modal');
    if (modal) {
      modal.classList.add('active');
      setTimeout(() => input && input.focus(), 150);
    }
  },

  closeDeleteModal() {
    const modal = document.getElementById('delete-boletim-modal');
    if (modal) modal.classList.remove('active');
  },

  /**
   * Confirma e executa a exclusão definitiva do boletim
   */
  confirmDeleteBoletim() {
    const target = this.state.targetDeleteBoletim;
    if (!target) return;

    this.state.boletins = this.state.boletins.filter(b => b.numero_boletim !== target);
    this.saveDatabase();
    this.supabase.deleteBoletim(target);

    this.state.activeBoletimNum = this.state.boletins.length > 0 ? this.state.boletins[0].numero_boletim : null;
    this.populateFilterDropdowns();
    this.runSearch();
    this.renderViewer();
    this.closeDeleteModal();

    alert(`Boletim ${target} foi excluído com sucesso do repositório.`);
  },

  /**
   * Dispara a impressão do navegador ou salvar em PDF
   */
  printBulletin() {
    this.switchTab('viewer');
    const select = document.getElementById('viewer-boletim-select');
    const num = select ? select.value : this.state.activeBoletimNum;
    const bol = this.state.boletins.find(b => b.numero_boletim === num) || this.state.boletins[0];

    if (bol) {
      const numStr = (bol.numero_boletim || '').trim();
      const perStr = (bol.periodo || '').trim().replace(/[\/\\:*?"<>|]/g, '.');
      document.title = perStr ? `Ementário - Boletim - ${numStr} - ${perStr}` : `Ementário - Boletim - ${numStr}`;
    }

    setTimeout(() => {
      window.print();
    }, 250);
  },

  /**
   * Sistema de Notificação Flutuante (Toast)
   */
  showToast(message, type = 'info') {
    let container = document.getElementById('toast-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'toast-container';
      document.body.appendChild(container);
    }

    const toast = document.createElement('div');
    toast.className = `app-toast toast-${type}`;
    
    let borderColor = 'var(--border-subtle)';
    if (type === 'success') borderColor = 'rgba(16, 185, 129, 0.6)';
    if (type === 'warning') borderColor = 'rgba(245, 158, 11, 0.6)';
    if (type === 'error') borderColor = 'rgba(239, 68, 68, 0.6)';
    toast.style.borderColor = borderColor;

    toast.innerHTML = `<span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, 3200);
  },

  // =========================================================================
  // GESTÃO DA CHAVE DE API DE IA (GOOGLE AI STUDIO OU OPENAI)
  // =========================================================================

  openGeminiModal() {
    const modal = document.getElementById('gemini-key-modal');
    const input = document.getElementById('gemini-api-key-input');
    const statusEl = document.getElementById('gemini-modal-status');
    const resultBox = document.getElementById('gemini-test-result');
    const currentKey = (localStorage.getItem('empresax_gemini_api_key') || localStorage.getItem('fastshop_gemini_api_key') || '').trim();

    if (resultBox) resultBox.style.display = 'none';
    if (input) input.value = currentKey;

    if (statusEl) {
      if (!currentKey) {
        statusEl.textContent = "Status: Não configurado (Motor Heurístico ativo)";
        statusEl.style.color = "var(--text-muted)";
      } else if (currentKey.startsWith('sk-')) {
        statusEl.textContent = "Status: Token OpenAI (ChatGPT) configurado";
        statusEl.style.color = "#10B981";
      } else if (currentKey.startsWith('AIzaSy') || currentKey.startsWith('AQ.')) {
        statusEl.textContent = "Status: Token Google Gemini (Plano Pro/Flash) configurado";
        statusEl.style.color = "#10B981";
      } else {
        statusEl.textContent = `Status: Token configurado (${currentKey.slice(0, 5)}...)`;
        statusEl.style.color = "#10B981";
      }
    }
    if (modal) modal.classList.add('active');
  },

  closeGeminiModal() {
    const modal = document.getElementById('gemini-key-modal');
    if (modal) modal.classList.remove('active');
  },

  toggleKeyVisibility() {
    const input = document.getElementById('gemini-api-key-input');
    if (input) {
      input.type = input.type === 'password' ? 'text' : 'password';
    }
  },

  async testApiKey() {
    const input = document.getElementById('gemini-api-key-input');
    const resultBox = document.getElementById('gemini-test-result');
    const testBtn = document.getElementById('btn-test-gemini-key');
    const key = (input ? input.value : '').trim();

    if (!key) {
      alert("Por favor, cole seu Token de API antes de clicar em Testar.");
      return;
    }

    if (resultBox) {
      resultBox.style.display = 'block';
      resultBox.style.background = 'rgba(99, 102, 241, 0.12)';
      resultBox.style.border = '1px solid rgba(99, 102, 241, 0.35)';
      resultBox.style.color = '#818CF8';
      resultBox.innerHTML = '⏳ Conectando aos servidores da IA em tempo real para validar...';
    }
    if (testBtn) testBtn.disabled = true;

    try {
      const res = await TaxSynthesizer.testConnection(key);
      if (resultBox) {
        resultBox.style.background = 'rgba(16, 185, 129, 0.12)';
        resultBox.style.border = '1px solid rgba(16, 185, 129, 0.4)';
        resultBox.style.color = '#34D399';
        resultBox.innerHTML = `✅ <strong>Conexão bem-sucedida!</strong> Provedor: <strong>${res.provider}</strong> | Modelo: <code>${res.model}</code>`;
      }
    } catch (err) {
      if (resultBox) {
        resultBox.style.background = 'rgba(239, 68, 68, 0.12)';
        resultBox.style.border = '1px solid rgba(239, 68, 68, 0.4)';
        resultBox.style.color = '#F87171';
        resultBox.innerHTML = `❌ <strong>Falha na validação do token:</strong><br>${err.message}`;
      }
    } finally {
      if (testBtn) testBtn.disabled = false;
    }
  },

  saveGeminiKey() {
    const input = document.getElementById('gemini-api-key-input');
    const key = (input ? input.value : '').trim();

    if (!key) {
      this.showToast("Por favor, cole seu token de API antes de salvar.", "warning");
      return;
    }

    localStorage.setItem('empresax_gemini_api_key', key);
    this.updateGeminiStatusIndicator();

    const resultBox = document.getElementById('gemini-test-result');
    if (resultBox) {
      resultBox.style.display = 'block';
      resultBox.style.background = 'rgba(16, 185, 129, 0.1)';
      resultBox.style.border = '1px solid rgba(16, 185, 129, 0.3)';
      resultBox.style.color = '#34D399';
      resultBox.innerHTML = '✅ <strong>Token salvo com sucesso!</strong> As próximas análises de leis utilizarão a IA.';
    }

    this.showToast("💾 Token de API salvo com sucesso no navegador!", "success");

    setTimeout(() => {
      this.closeGeminiModal();
      if (resultBox) resultBox.style.display = 'none';
    }, 900);
  },

  clearGeminiKey() {
    localStorage.removeItem('empresax_gemini_api_key');
    localStorage.removeItem('fastshop_gemini_api_key');
    const input = document.getElementById('gemini-api-key-input');
    if (input) input.value = '';

    const resultBox = document.getElementById('gemini-test-result');
    if (resultBox) {
      resultBox.style.display = 'block';
      resultBox.style.background = 'rgba(16, 185, 129, 0.1)';
      resultBox.style.border = '1px solid rgba(16, 185, 129, 0.3)';
      resultBox.style.color = '#34D399';
      resultBox.innerHTML = '✅ <strong>Token removido com sucesso.</strong> O portal voltou a operar com o Motor Heurístico Integrado (100% gratuito).';
    }

    this.updateGeminiStatusIndicator();
    this.showToast("🔑 Token de API removido com sucesso.", "info");

    setTimeout(() => {
      this.closeGeminiModal();
      if (resultBox) resultBox.style.display = 'none';
    }, 1200);
  },

  updateGeminiStatusIndicator() {
    const key = (localStorage.getItem('empresax_gemini_api_key') || localStorage.getItem('fastshop_gemini_api_key') || '').trim();
    const dot = document.getElementById('gemini-status-indicator');
    const btnLabel = document.getElementById('gemini-btn-label');

    if (dot) {
      if (key && key.length > 10) {
        dot.style.background = '#10B981';
        dot.title = 'Token de IA Conectado (Tempo Real)';
        if (btnLabel) btnLabel.textContent = 'Token Ativo';
      } else {
        dot.style.background = '#9CA3AF';
        dot.title = 'Motor Heurístico Local Ativo';
        if (btnLabel) btnLabel.textContent = 'Token de IA';
      }
    }
  },

  openGeminiMissingModal() {
    const m = document.getElementById('gemini-missing-modal');
    if (m) m.classList.add('active');
  },

  closeGeminiMissingModal() {
    const m = document.getElementById('gemini-missing-modal');
    if (m) m.classList.remove('active');
  },

  openGeminiModalFromWarning() {
    this.closeGeminiMissingModal();
    this.openGeminiModal();
  },

  proceedWithoutGemini() {
    this.closeGeminiMissingModal();
    this.processGridData(true);
  },

  // =========================================================================
  // EDIÇÃO DE ITENS FINALIZADOS (LEGISLAÇÃO E NOTÍCIAS)
  // =========================================================================

  openEditItemModal(tipoRegistro, boletimNum, itemIndex) {
    const boletim = this.state.boletins.find(b => b.numero_boletim === boletimNum) 
      || (this.state.generatedBoletim && this.state.generatedBoletim.numero_boletim === boletimNum ? this.state.generatedBoletim : null);

    if (!boletim) {
      alert(`Boletim ${boletimNum} não encontrado.`);
      return;
    }

    const badge = document.getElementById('edit-modal-boletim-badge');
    if (badge) badge.textContent = `Boletim ${boletimNum}`;

    document.getElementById('edit-tipo-registro').value = tipoRegistro;
    document.getElementById('edit-boletim-num').value = boletimNum;
    document.getElementById('edit-item-index').value = itemIndex;

    const lawContainer = document.getElementById('edit-law-container');
    const newsContainer = document.getElementById('edit-news-container');

    if (tipoRegistro === 'LEI') {
      if (lawContainer) lawContainer.style.display = 'block';
      if (newsContainer) newsContainer.style.display = 'none';

      const item = boletim.itens ? boletim.itens[itemIndex] : null;
      if (!item) {
        alert("Publicação não encontrada neste boletim.");
        return;
      }

      document.getElementById('edit-titulo').value = item.titulo || '';
      document.getElementById('edit-norma').value = item.norma || '';
      document.getElementById('edit-orgao').value = item.orgao || '';
      document.getElementById('edit-esfera').value = (item.esfera || 'ESTADUAL').toUpperCase();
      document.getElementById('edit-impacto').value = item.impacto || '';
      document.getElementById('edit-area').value = item.area_impactada || '';
      document.getElementById('edit-vigencia').value = item.vigencia || '';
      document.getElementById('edit-resumo-tabela').value = item.resumo_tabela || '';
      document.getElementById('edit-link').value = item.link || '';
      document.getElementById('edit-corpo').value = (item.corpo_paragrafos || []).join('\n\n');
      const entInput = document.getElementById('edit-entendimento');
      if (entInput) entInput.value = item.entendimento_assunto || '';
      document.getElementById('edit-plano-acao').value = (item.plano_de_acao || []).join('\n');
    } else {
      // Notícia
      if (lawContainer) lawContainer.style.display = 'none';
      if (newsContainer) newsContainer.style.display = 'block';

      const noticia = boletim.noticias ? boletim.noticias[itemIndex] : null;
      if (!noticia) {
        alert("Notícia não encontrada neste boletim.");
        return;
      }

      document.getElementById('edit-news-titulo').value = noticia.titulo || '';
      document.getElementById('edit-news-fonte').value = noticia.fonte || '';
      document.getElementById('edit-news-link').value = noticia.link || '';
      document.getElementById('edit-news-corpo').value = (noticia.corpo_paragrafos || []).join('\n\n');
      const newsEntInput = document.getElementById('edit-news-entendimento');
      if (newsEntInput) newsEntInput.value = noticia.entendimento_assunto || '';
    }

    const modal = document.getElementById('edit-item-modal');
    if (modal) modal.classList.add('active');
  },

  closeEditItemModal() {
    const modal = document.getElementById('edit-item-modal');
    if (modal) modal.classList.remove('active');
  },

  saveEditedItem() {
    const tipoRegistro = document.getElementById('edit-tipo-registro').value;
    const boletimNum = document.getElementById('edit-boletim-num').value;
    const itemIndex = parseInt(document.getElementById('edit-item-index').value, 10);

    const targetBoletins = [];
    const mainBol = this.state.boletins.find(b => b.numero_boletim === boletimNum);
    if (mainBol) targetBoletins.push(mainBol);
    if (this.state.generatedBoletim && this.state.generatedBoletim.numero_boletim === boletimNum) {
      targetBoletins.push(this.state.generatedBoletim);
    }

    if (targetBoletins.length === 0) {
      alert(`Boletim ${boletimNum} não encontrado.`);
      return;
    }

    targetBoletins.forEach(boletim => {
      if (tipoRegistro === 'LEI') {
        const item = boletim.itens ? boletim.itens[itemIndex] : null;
        if (!item) return;

        item.titulo = (document.getElementById('edit-titulo').value || '').trim();
        item.norma = (document.getElementById('edit-norma').value || '').trim();
        item.orgao = (document.getElementById('edit-orgao').value || '').trim();
        item.esfera = (document.getElementById('edit-esfera').value || 'ESTADUAL').trim();
        item.impacto = (document.getElementById('edit-impacto').value || '').trim();
        item.area_impactada = (document.getElementById('edit-area').value || '').trim();
        item.vigencia = (document.getElementById('edit-vigencia').value || '').trim();
        item.resumo_tabela = (document.getElementById('edit-resumo-tabela').value || '').trim();
        item.link = (document.getElementById('edit-link').value || '').trim();

        const rawCorpo = document.getElementById('edit-corpo').value || '';
        item.corpo_paragrafos = rawCorpo.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);

        const entInput = document.getElementById('edit-entendimento');
        if (entInput) item.entendimento_assunto = entInput.value.trim();

        const rawPlan = document.getElementById('edit-plano-acao').value || '';
        item.plano_de_acao = rawPlan.split('\n').map(l => l.trim()).filter(Boolean);
      } else {
        const noticia = boletim.noticias ? boletim.noticias[itemIndex] : null;
        if (!noticia) return;

        noticia.titulo = (document.getElementById('edit-news-titulo').value || '').trim();
        noticia.fonte = (document.getElementById('edit-news-fonte').value || '').trim();
        noticia.link = (document.getElementById('edit-news-link').value || '').trim();

        const rawCorpo = document.getElementById('edit-news-corpo').value || '';
        noticia.corpo_paragrafos = rawCorpo.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);

        const newsEntInput = document.getElementById('edit-news-entendimento');
        if (newsEntInput) noticia.entendimento_assunto = newsEntInput.value.trim();
      }
    });

    // Persiste no LocalStorage e atualiza todos os componentes
    this.saveDatabase(mainBol);
    this.populateFilterDropdowns();
    this.runSearch();

    // Se o boletim editado for o atualmente exibido no visualizador A4/PDF, atualiza em tempo real
    if (this.state.activeBoletimNum === boletimNum) {
      this.renderViewer();
    }

    // Se a pré-visualização de novo boletim estiver visível, atualiza os cards
    const draftPreviewCards = document.getElementById('generated-preview-cards');
    if (draftPreviewCards && this.state.generatedBoletim && this.state.generatedBoletim.numero_boletim === boletimNum) {
      const flat = SearchEngine.flattenDatabase([this.state.generatedBoletim]);
      SearchEngine.renderResults(draftPreviewCards, flat);
    }

    this.closeEditItemModal();
    alert("Publicação atualizada com sucesso! As alterações estão salvas no repositório e refletidas no visualizador de PDF.");
  }
};

// Inicialização automática ao carregar o DOM
document.addEventListener('DOMContentLoaded', () => {
  App.init();
});

