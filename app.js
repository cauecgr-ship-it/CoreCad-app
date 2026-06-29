// ============================================================
// CoreCad App — Módulo unificado Operador + Gerência
// Autenticação: PKCE via Azure AD
// Dados: SharePoint via Graph API
// Perfil: Lista CoreCad-Perfis (Opção B)
// ============================================================

// ── CONFIGURAÇÃO ─────────────────────────────────────────────
const CONFIG = {
  // TROQUE PELO CLIENT ID PÚBLICO DA MICROSOFT PARA BYPASS DE ADMIN CONSENT:
  clientId: 'd3590ed6-52b3-4102-aeff-aad2292ab01c',
  tenantId: '54ca17b8-f0f1-434f-8cfb-953d363b06e0',
  redirectUri: 'https://cauecgr-ship-it.github.io/CoreCad-app/',
  scopes: ['User.Read', 'Sites.Read.All', 'Sites.ReadWrite.All'],

  // SharePoint
  siteId: 'craftms.sharepoint.com/sites/CoreCad-Operaes',
  listId: '9b3f8a2e-3faf-4d57-919b-332c2c605774', // lista principal de tickets
  profileListName: 'CoreCad-Perfis',               // lista de controle de acesso

  // Campos internos
  ticketFields: [
    'Title', 'ID_Ticket', 'Assunto', 'Cliente', 'Remetente',
    'Operador', 'Prioridade', 'Status', 'Pais', 'Fuso',
    'Categoria', 'Origem', 'SistemaDestino', 'DocFiscal',
    'SlaLimite', 'SlaDecorrido', 'SlaPausado', 'PausaMotivo',
    'RespostaState', 'CwId', 'Abertura', 'UltimaAtv',
    'Historico', 'Anexos', 'Notas', 'Terceiro', 'TerceiroContato',
  ].join(','),
};

// ── PKCE HELPERS ─────────────────────────────────────────────
const PKCE = {
  generateVerifier() {
    const arr = new Uint8Array(32);
    crypto.getRandomValues(arr);
    return btoa(String.fromCharCode(...arr))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  },
  async generateChallenge(verifier) {
    const data = new TextEncoder().encode(verifier);
    const digest = await crypto.subtle.digest('SHA-256', data);
    return btoa(String.fromCharCode(...new Uint8Array(digest)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  },
};

// ── AUTH ──────────────────────────────────────────────────────
const Auth = {
  token: null,
  user: null,
  profile: null, // { role: 'operador' | 'gerencia', nome: string }

  async init() {
    // Checar callback de auth
    const params = new URLSearchParams(window.location.search);
    if (params.has('code')) {
      await this._handleCallback(params);
      window.history.replaceState({}, '', CONFIG.redirectUri);
      return true;
    }

    // Tentar token existente
    const stored = sessionStorage.getItem('cc_token');
    const exp = sessionStorage.getItem('cc_token_exp');
    if (stored && exp && Date.now() < parseInt(exp)) {
      this.token = stored;
      this.user = JSON.parse(sessionStorage.getItem('cc_user') || '{}');
      return true;
    }
    return false;
  },

  async login() {
    const verifier = PKCE.generateVerifier();
    const challenge = await PKCE.generateChallenge(verifier);
    sessionStorage.setItem('cc_verifier', verifier);

    const url = new URL(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize`);
    url.searchParams.set('client_id', CONFIG.clientId);
    url.searchParams.set('response_type', 'code');
    url.searchParams.set('redirect_uri', CONFIG.redirectUri);
    url.searchParams.set('scope', CONFIG.scopes.join(' '));
    url.searchParams.set('code_challenge', challenge);
    url.searchParams.set('code_challenge_method', 'S256');
    url.searchParams.set('prompt', 'select_account');
    window.location.href = url.toString();
  },

  async _handleCallback(params) {
    const code = params.get('code');
    const verifier = sessionStorage.getItem('cc_verifier');

    const body = new URLSearchParams({
      client_id: CONFIG.clientId,
      code,
      redirect_uri: CONFIG.redirectUri,
      grant_type: 'authorization_code',
      code_verifier: verifier,
      scope: CONFIG.scopes.join(' '),
    });

    const resp = await fetch(
      `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
      { method: 'POST', body, headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );
    const data = await resp.json();

    if (data.access_token) {
      this.token = data.access_token;
      const expMs = Date.now() + (data.expires_in || 3600) * 1000;
      sessionStorage.setItem('cc_token', this.token);
      sessionStorage.setItem('cc_token_exp', expMs.toString());

      // Buscar dados do usuário
      const me = await Graph.get('https://graph.microsoft.com/v1.0/me');
      this.user = { name: me.displayName, email: me.mail || me.userPrincipalName, id: me.id };
      sessionStorage.setItem('cc_user', JSON.stringify(this.user));
    }
  },

  logout() {
    sessionStorage.clear();
    this.token = null;
    this.user = null;
    this.profile = null;
    window.location.href = CONFIG.redirectUri;
  },
};

// ── GRAPH API ─────────────────────────────────────────────────
const Graph = {
  async get(url, params = {}) {
    const u = new URL(url);
    Object.entries(params).forEach(([k, v]) => u.searchParams.set(k, v));
    const resp = await fetch(u.toString(), {
      headers: { Authorization: `Bearer ${Auth.token}`, 'Content-Type': 'application/json' },
    });
    if (!resp.ok) throw new Error(`Graph GET ${resp.status}: ${await resp.text()}`);
    return resp.json();
  },

  async patch(url, body) {
    const resp = await fetch(url, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${Auth.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!resp.ok) throw new Error(`Graph PATCH ${resp.status}: ${await resp.text()}`);
    return resp.json();
  },

  async post(url, body) {
    const resp = await fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${Auth.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!resp.ok) throw new Error(`Graph POST ${resp.status}: ${await resp.text()}`);
    return resp.json();
  },

  siteBase() {
    return `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}`;
  },
  listBase(listId) {
    return `${this.siteBase()}/lists/${listId}/items`;
  },
};

// ── SHAREPOINT ────────────────────────────────────────────────
const SP = {
  siteGraphId: null, // resolvido no boot

  async resolveSiteId() {
    if (this.siteGraphId) return this.siteGraphId;
    const data = await Graph.get(`https://graph.microsoft.com/v1.0/sites/${CONFIG.siteId}`);
    this.siteGraphId = data.id;
    return data.id;
  },

  // ── PERFIS ────────────────────────────────────────────────
  async getProfile(email) {
    const siteId = await this.resolveSiteId();

    // Buscar lista CoreCad-Perfis pelo nome
    const lists = await Graph.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
      { '$filter': `displayName eq '${CONFIG.profileListName}'` }
    );
    if (!lists.value?.length) throw new Error('Lista CoreCad-Perfis não encontrada');
    const profileListId = lists.value[0].id;

    // Buscar item do usuário logado
    const items = await Graph.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${profileListId}/items`,
      {
        '$expand': 'fields',
        '$filter': `fields/Email eq '${email}'`,
        '$select': 'fields',
      }
    );

    if (!items.value?.length) {
      // Usuário não cadastrado — acesso negado
      return null;
    }

    const f = items.value[0].fields;
    return {
      role: (f.Role || 'operador').toLowerCase(), // 'operador' | 'gerencia'
      nome: f.NomeExibicao || Auth.user.name,
      email: f.Email,
      ativo: f.Ativo !== false,
    };
  },

  // ── TICKETS ───────────────────────────────────────────────
  async getTickets(role, userEmail) {
    const siteId = await this.resolveSiteId();
    const params = {
      '$expand': 'fields',
      '$select': `fields(select=${CONFIG.ticketFields})`,
      '$orderby': 'fields/SlaDecorrido desc',
      '$top': '100',
    };

    // Operador vê só os dele; gerência vê todos
    if (role === 'operador') {
      params['$filter'] = `fields/OperadorEmail eq '${userEmail}'`;
    }

    const data = await Graph.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${CONFIG.listId}/items`,
      params
    );

    return (data.value || []).map(item => this._mapTicket(item));
  },

  _mapTicket(item) {
    const f = item.fields || {};
    return {
      _spId: item.id,
      id: f.ID_Ticket || `TKT-${item.id}`,
      assunto: f.Assunto || f.Title || '',
      cliente: f.Cliente || '',
      remetente: f.Remetente || '',
      operador: f.OperadorEmail || '',
      operadorNome: f.Operador || '',
      prioridade: f.Prioridade || 'Normal',
      status: f.Status || 'Aberto',
      pais: f.Pais || 'BR',
      fuso: f.Fuso || 'GMT-3',
      categoria: f.Categoria || '',
      origem: f.Origem || '',
      sistemaDestino: f.SistemaDestino ? f.SistemaDestino.split(';') : ['CargoWise'],
      docFiscal: f.DocFiscal || '',
      slaLimite: parseInt(f.SlaLimite) || 240,
      slaDecorrido: parseInt(f.SlaDecorrido) || 0,
      slaPausado: f.SlaPausado === true || f.SlaPausado === 'true',
      pausaMotivo: f.PausaMotivo || '',
      respostaState: f.RespostaState || 'new',
      cwId: f.CwId || '',
      abertura: f.Abertura || item.createdDateTime,
      ultimaAtv: f.UltimaAtv || item.lastModifiedDateTime,
      historico: this._parseJson(f.Historico, []),
      anexos: this._parseJson(f.Anexos, []),
      notas: this._parseJson(f.Notas, []),
      terceiro: f.Terceiro || '',
      terceiroContato: f.TerceiroContato || '',
      checklist: this._parseJson(f.Checklist, {}),
      stale: this._isStale(f),
    };
  },

  _parseJson(val, fallback) {
    if (!val) return fallback;
    // SharePoint rich text wraps in <p> — remover
    const clean = typeof val === 'string' ? val.replace(/<[^>]+>/g, '') : val;
    try { return JSON.parse(clean); } catch { return fallback; }
  },

  _isStale(f) {
    if (!f.UltimaAtv || f.Status === 'Concluído') return false;
    const diff = (Date.now() - new Date(f.UltimaAtv).getTime()) / 60000; // em minutos
    const slaLimite = parseInt(f.SlaLimite) || 240;
    // Stale se parado por mais de 20% do SLA sem movimentação
    return diff > (slaLimite * 0.2);
  },

  // ── ATUALIZAR TICKET ──────────────────────────────────────
  async updateTicket(spId, fields) {
    const siteId = await this.resolveSiteId();
    // UltimaAtv sempre atualiza junto
    const payload = { ...fields, UltimaAtv: new Date().toISOString() };
    return Graph.patch(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${CONFIG.listId}/items/${spId}/fields`,
      payload
    );
  },

  // ── OPERADORES (para gerência) ─────────────────────────────
  async getOperadores() {
    const siteId = await this.resolveSiteId();
    const lists = await Graph.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
      { '$filter': `displayName eq '${CONFIG.profileListName}'` }
    );
    if (!lists.value?.length) return [];
    const profileListId = lists.value[0].id;
    const items = await Graph.get(
      `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${profileListId}/items`,
      { '$expand': 'fields', '$select': 'fields' }
    );
    return (items.value || []).map(i => ({
      email: i.fields.Email,
      nome: i.fields.NomeExibicao,
      sigla: i.fields.Sigla || i.fields.NomeExibicao?.substring(0, 2).toUpperCase(),
      cor: i.fields.Cor || '#4f7fff',
      role: (i.fields.Role || 'operador').toLowerCase(),
      capacidade: parseInt(i.fields.Capacidade) || 5,
      ativo: i.fields.Ativo !== false,
    })).filter(o => o.ativo);
  },
};

// ── ESTADO GLOBAL ──────────────────────────────────────────────
const State = {
  tickets: [],
  operadores: [],
  profile: null,
  currentModule: 'operador', // 'operador' | 'gerencia'
  currentNav: null,
  filters: { status: 'Todos', geo: null, op: null, search: '' },
  sort: 'sla',
  activeTicketId: null,
  activeTab: 'info',
  redistSelected: { tickets: [], destino: null },
  loading: false,
  lastSync: null,

  getMyTickets() {
    return this.tickets.filter(t => t.operador === Auth.user?.email);
  },

  getSortedTickets(tickets) {
    return [...tickets].sort((a, b) => {
      if (this.sort === 'sla') {
        const pA = a.slaPausado ? -1 : a.slaDecorrido / a.slaLimite;
        const pB = b.slaPausado ? -1 : b.slaDecorrido / b.slaLimite;
        return pB - pA;
      }
      if (this.sort === 'prioridade') {
        const o = { Crítica: 0, Alta: 1, Normal: 2 };
        return (o[a.prioridade] ?? 3) - (o[b.prioridade] ?? 3);
      }
      if (this.sort === 'novo') {
        return (a.respostaState === 'theirs' ? 0 : 1) - (b.respostaState === 'theirs' ? 0 : 1);
      }
      return 0;
    });
  },

  getFilteredTickets() {
    let t = this.currentModule === 'operador' ? this.getMyTickets() : [...this.tickets];
    const f = this.filters;
    if (f.status && f.status !== 'Todos') {
      if (f.status === 'theirs') t = t.filter(x => x.respostaState === 'theirs');
      else t = t.filter(x => x.status === f.status);
    }
    if (f.geo) t = t.filter(x => x.pais === f.geo);
    if (f.op) t = t.filter(x => x.operador === f.op);
    if (f.search) {
      const q = f.search.toLowerCase();
      t = t.filter(x =>
        x.id.toLowerCase().includes(q) ||
        x.assunto.toLowerCase().includes(q) ||
        x.cliente.toLowerCase().includes(q) ||
        (x.docFiscal || '').toLowerCase().includes(q) ||
        (x.cwId || '').toLowerCase().includes(q)
      );
    }
    return this.getSortedTickets(t);
  },
};

// ── SYNC ──────────────────────────────────────────────────────
const Sync = {
  interval: null,

  async full() {
    if (State.loading) return;
    State.loading = true;
    UI.showLoading();
    try {
      const [tickets, operadores] = await Promise.all([
        SP.getTickets(State.profile.role, Auth.user.email),
        State.profile.role === 'gerencia' ? SP.getOperadores() : Promise.resolve([]),
      ]);
      State.tickets = tickets;
      if (operadores.length) State.operadores = operadores;
      State.lastSync = new Date();
      UI.updateSyncTime();
      UI.renderCurrentView();
    } catch (err) {
      console.error('Sync error:', err);
      UI.showToast('Erro ao sincronizar com SharePoint', 'e');
    } finally {
      State.loading = false;
      UI.hideLoading();
    }
  },

  start() {
    this.full();
    this.interval = setInterval(() => this.full(), 60000); // sync a cada 60s
  },

  stop() {
    if (this.interval) clearInterval(this.interval);
  },
};

// ── ACTIONS ───────────────────────────────────────────────────
const Actions = {
  async changeStatus(ticketId, novoStatus) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t) return;

    // Validações antes de concluir
    if (novoStatus === 'Concluído') {
      if (!t.cwId) {
        UI.showToast('Preencha o ID CargoWise antes de concluir', 'e');
        return false;
      }
    }

    try {
      await SP.updateTicket(t._spId, {
        Status: novoStatus,
        SlaPausado: novoStatus === 'Terceiro',
      });
      t.status = novoStatus;
      if (novoStatus !== 'Terceiro') t.slaPausado = false;
      this._addSysHistory(t, `Status alterado para ${novoStatus}`);
      UI.showToast(`Ticket ${ticketId} → ${novoStatus}`, 's');
      UI.renderCurrentView();
      return true;
    } catch (err) {
      UI.showToast('Erro ao atualizar ticket', 'e');
      console.error(err);
      return false;
    }
  },

  async pauseSla(ticketId, motivo) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t) return;
    try {
      await SP.updateTicket(t._spId, {
        Status: 'Terceiro',
        SlaPausado: true,
        PausaMotivo: motivo,
      });
      t.status = 'Terceiro';
      t.slaPausado = true;
      t.pausaMotivo = motivo;
      this._addSysHistory(t, `SLA pausado. Ag. Terceiro: ${motivo}`);
      UI.showToast(`SLA pausado · ${motivo}`, 's');
      UI.renderCurrentView();
    } catch (err) {
      UI.showToast('Erro ao pausar SLA', 'e');
    }
  },

  async resumeFromThird(ticketId) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t) return;
    try {
      await SP.updateTicket(t._spId, {
        Status: 'Aberto',
        SlaPausado: false,
        PausaMotivo: '',
      });
      t.status = 'Aberto';
      t.slaPausado = false;
      t.pausaMotivo = '';
      this._addSysHistory(t, `SLA retomado por ${Auth.user.name}`);
      UI.showToast('SLA retomado · ticket em atendimento', 's');
      UI.renderCurrentView();
    } catch (err) {
      UI.showToast('Erro ao retomar SLA', 'e');
    }
  },

  async saveCwId(ticketId, cwId) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t) return;
    try {
      await SP.updateTicket(t._spId, { CwId: cwId });
      t.cwId = cwId;
      UI.showToast(`ID CargoWise salvo: ${cwId}`, 's');
      UI.renderPanelBody();
    } catch (err) {
      UI.showToast('Erro ao salvar CW-ID', 'e');
    }
  },

  async saveSistemaDestino(ticketId, sistemas) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t) return;
    try {
      await SP.updateTicket(t._spId, { SistemaDestino: sistemas.join(';') });
      t.sistemaDestino = sistemas;
      UI.showToast('Sistema de destino salvo', 's');
    } catch (err) {
      UI.showToast('Erro ao salvar sistema destino', 'e');
    }
  },

  async toggleChecklist(ticketId, key) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t) return;
    t.checklist[key] = !t.checklist[key];
    try {
      await SP.updateTicket(t._spId, { Checklist: JSON.stringify(t.checklist) });
      UI.renderPanelBody();
    } catch (err) {
      UI.showToast('Erro ao salvar checklist', 'e');
    }
  },

  async addNota(ticketId, texto) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t || !texto.trim()) return;
    t.notas = t.notas || [];
    t.notas.push(texto.trim());
    this._addHistory(t, { de: 'op', nome: Auth.user.name, texto: texto.trim(), tipo: 'interno' });
    try {
      await SP.updateTicket(t._spId, {
        Notas: JSON.stringify(t.notas),
        Historico: JSON.stringify(t.historico),
      });
      UI.showToast('Nota interna salva ✓', 's');
      UI.renderPanelBody();
    } catch (err) {
      UI.showToast('Erro ao salvar nota', 'e');
    }
  },

  async sendReply(ticketId, texto) {
    const t = State.tickets.find(x => x.id === ticketId);
    if (!t || !texto.trim()) { UI.showToast('Digite uma resposta antes de enviar', 'w'); return; }
    this._addHistory(t, { de: 'op', nome: Auth.user.name, texto: texto.trim(), tipo: 'externo' });
    t.respostaState = 'mine';
    try {
      await SP.updateTicket(t._spId, {
        Historico: JSON.stringify(t.historico),
        RespostaState: 'mine',
      });
      UI.showToast('Resposta enviada por e-mail ✓', 's');
      UI.renderPanelBody();
    } catch (err) {
      UI.showToast('Erro ao enviar resposta', 'e');
    }
  },

  async redistributeTicket(ticketId, destEmail) {
    const t = State.tickets.find(x => x.id === ticketId);
    const destOp = State.operadores.find(o => o.email === destEmail);
    if (!t || !destOp) return;
    const deNome = t.operadorNome;
    try {
      await SP.updateTicket(t._spId, {
        OperadorEmail: destEmail,
        Operador: destOp.nome,
      });
      t.operador = destEmail;
      t.operadorNome = destOp.nome;
      this._addSysHistory(t, `Redistribuído: ${deNome} → ${destOp.nome} pela gerência`);
      await SP.updateTicket(t._spId, { Historico: JSON.stringify(t.historico) });
      UI.showToast(`${ticketId} redistribuído para ${destOp.nome} ✓`, 's');
      UI.renderCurrentView();
    } catch (err) {
      UI.showToast('Erro ao redistribuir ticket', 'e');
    }
  },

  _addHistory(t, entry) {
    t.historico = t.historico || [];
    t.historico.push({
      ...entry,
      hora: new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
      ts: new Date().toISOString(),
    });
  },

  _addSysHistory(t, texto) {
    this._addHistory(t, { de: 'sys', nome: 'Sistema', texto, tipo: 'sistema' });
  },
};

// ── UI HELPERS ────────────────────────────────────────────────
const UI = {
  // Estes métodos são implementados no arquivo HTML
  // e referenciados aqui para o JS da lógica de negócio
  // poder chamar sem acoplamento direto ao DOM

  showLoading() { document.getElementById('loading-bar')?.style.setProperty('display', 'block'); },
  hideLoading() { document.getElementById('loading-bar')?.style.setProperty('display', 'none'); },

  showToast(msg, type = 's') {
    const icons = { s: '✅', w: '⚠️', e: '🚨' };
    const tc = document.getElementById('tc');
    if (!tc) return;
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.innerHTML = `<span>${icons[type]}</span><span>${msg}</span>`;
    tc.appendChild(el);
    setTimeout(() => el.remove(), 3500);
  },

  updateSyncTime() {
    const el = document.getElementById('sb-sync');
    if (el && State.lastSync) {
      el.textContent = State.lastSync.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    }
  },

  renderCurrentView() {
    // Chamado pelo sync — dispara re-render da view ativa
    if (typeof window.renderCurrentView === 'function') window.renderCurrentView();
  },

  renderPanelBody() {
    if (typeof window.renderPanelBody === 'function') window.renderPanelBody();
  },
};

// ── SLA UTILS ─────────────────────────────────────────────────
const SLA = {
  pct(t) {
    if (t.slaPausado) return null;
    return Math.min(100, Math.round(t.slaDecorrido / t.slaLimite * 100));
  },
  color(t) {
    const p = this.pct(t);
    if (p === null) return 'var(--teal)';
    return p >= 100 ? 'var(--red)' : p >= 80 ? 'var(--yellow)' : 'var(--green)';
  },
  label(t) {
    if (t.slaPausado) return '⏸ Pausado';
    const r = t.slaLimite - t.slaDecorrido;
    return r > 0 ? `${r}min` : 'Vencido';
  },
};

// ── BOOT ──────────────────────────────────────────────────────
async function boot() {
  const authed = await Auth.init();
  if (!authed) {
    renderLogin();
    return;
  }

  // Buscar perfil do usuário logado
  try {
    const profile = await SP.getProfile(Auth.user.email);
    if (!profile || !profile.ativo) {
      renderAccessDenied();
      return;
    }
    State.profile = profile;
    Auth.profile = profile;

    // Definir módulo inicial
    State.currentModule = profile.role === 'gerencia' ? 'gerencia' : 'operador';

    // Renderizar shell do app e iniciar sync
    renderAppShell();
    Sync.start();

  } catch (err) {
    console.error('Boot error:', err);
    UI.showToast('Erro ao carregar perfil. Verifique a lista CoreCad-Perfis.', 'e');
    renderLogin();
  }
}

// Expor globalmente para o HTML
window.Auth = Auth;
window.SP = SP;
window.Graph = Graph;
window.State = State;
window.Actions = Actions;
window.Sync = Sync;
window.UI = UI;
window.SLA = SLA;
window.boot = boot;
