// ============================================================
// CoreCad — Módulo Operador
// Renderização completa da interface do operador
// Depende de: app.js (State, Actions, SLA, UI, Auth)
// ============================================================

const PAISES = { BR:'🇧🇷', MX:'🇲🇽', AR:'🇦🇷', CL:'🇨🇱', CO:'🇨🇴', US:'🇺🇸' };

const CHECKLISTS = {
  'CNEE': [
    { layer: '📄 Documental', items: [
      { id: 'd1', label: 'Documento fiscal recebido', sub: 'RUT/CUIT/NIT/EIN conforme país' },
      { id: 'd2', label: 'Documento validado (dígito verificador)', sub: 'Verificar algoritmo por país' },
      { id: 'd3', label: 'Documento não vencido', sub: 'Validade do documento fiscal' },
    ]},
    { layer: '⚙️ Sistêmica', items: [
      { id: 's1', label: 'Verificar duplicidade no CargoWise', sub: 'Buscar por doc. fiscal e razão social', required: true },
      { id: 's2', label: 'Registrar CNEE no CargoWise', sub: 'Módulo: Organization > CNEE' },
      { id: 's3', label: 'Configurar contatos e endereços', sub: 'Mínimo: endereço fiscal e contato principal' },
      { id: 's4', label: 'Confirmar ID CargoWise gerado', sub: 'Preencher campo CW-ID antes de concluir', required: true },
    ]},
    { layer: '✅ Aprovação', items: [
      { id: 'a1', label: 'Confirmar cadastro ao cliente', sub: 'E-mail com ID CargoWise gerado' },
    ]},
  ],
  'Fornecedor': [
    { layer: '📄 Documental', items: [
      { id: 'd1', label: 'CNPJ/doc. fiscal do fornecedor recebido', sub: '' },
      { id: 'd2', label: 'Dados bancários recebidos (se aplicável)', sub: '' },
    ]},
    { layer: '⚙️ Sistêmica', items: [
      { id: 's1', label: 'Verificar duplicidade', sub: '', required: true },
      { id: 's2', label: 'Registrar fornecedor no CargoWise', sub: 'Módulo: Creditor' },
      { id: 's3', label: 'Registrar no CraftHub (se aplicável)', sub: '' },
      { id: 's4', label: 'Confirmar ID CargoWise', sub: '', required: true },
    ]},
    { layer: '✅ Aprovação', items: [
      { id: 'a1', label: 'Notificar financeiro sobre novo fornecedor', sub: '' },
      { id: 'a2', label: 'Confirmar ao solicitante', sub: '' },
    ]},
  ],
  'Agente': [
    { layer: '📄 Documental', items: [
      { id: 'd1', label: 'Contrato de representação recebido', sub: '' },
      { id: 'd2', label: 'Documentação fiscal do agente', sub: '' },
    ]},
    { layer: '⚙️ Sistêmica', items: [
      { id: 's1', label: 'Verificar duplicidade', sub: '', required: true },
      { id: 's2', label: 'Cadastrar agente no CargoWise', sub: 'Módulo: Agent' },
      { id: 's3', label: 'Configurar rotas e modais', sub: '' },
      { id: 's4', label: 'Confirmar ID CargoWise', sub: '', required: true },
    ]},
    { layer: '✅ Aprovação', items: [
      { id: 'a1', label: 'Aprovação do gestor de parceiros', sub: '' },
      { id: 'a2', label: 'Confirmar ativação ao agente', sub: '' },
    ]},
  ],
  'default': [
    { layer: '📄 Documental', items: [
      { id: 'd1', label: 'Documentação recebida', sub: '' },
      { id: 'd2', label: 'Documentação validada', sub: '' },
    ]},
    { layer: '⚙️ Sistêmica', items: [
      { id: 's1', label: 'Verificar duplicidade', sub: '', required: true },
      { id: 's2', label: 'Executar cadastro no sistema de destino', sub: '' },
      { id: 's3', label: 'Confirmar ID gerado', sub: '', required: true },
    ]},
    { layer: '✅ Aprovação', items: [
      { id: 'a1', label: 'Confirmar ao solicitante', sub: '' },
    ]},
  ],
};

const TEMPLATES = {
  '': '',
  'recebido': 'Olá,\n\nConfirmamos o recebimento da sua solicitação ({ID}).\nEstamos processando as informações e retornaremos em breve.\n\nAtenciosamente,\nEquipe CoreCad — Craft Multimodal',
  'doc_pendente': 'Olá,\n\nAgradecemos o contato. Para dar continuidade ao cadastro ({ID}), precisamos dos seguintes documentos:\n\n• [LISTAR DOCUMENTOS NECESSÁRIOS]\n\nPor favor, encaminhe assim que possível.\n\nAtenciosamente,\nEquipe CoreCad — Craft Multimodal',
  'concluido': 'Olá,\n\nInformamos que o cadastro referente à solicitação {ID} foi concluído com sucesso.\n\nID CargoWise: {CW_ID}\n\nEm caso de dúvidas, estamos à disposição.\n\nAtenciosamente,\nEquipe CoreCad — Craft Multimodal',
  'espera_gov': 'Olá,\n\nSua solicitação {ID} está em análise pelo setor de Governança.\nAssim que tivermos retorno, entraremos em contato.\n\nAtenciosamente,\nEquipe CoreCad — Craft Multimodal',
};

// ── RENDER PAINEL OPERADOR ────────────────────────────────────
function renderOperadorPainel() {
  const myT = State.getMyTickets();
  const abertos = myT.filter(t => t.status === 'Aberto').length;
  const aguardaResp = myT.filter(t => t.respostaState === 'theirs').length;
  const espera = myT.filter(t => t.status === 'Espera').length;
  const terceiro = myT.filter(t => t.status === 'Terceiro').length;
  const concluidos = myT.filter(t => t.status === 'Concluído').length;
  const criticos = myT.filter(t => t.prioridade === 'Crítica' && t.status === 'Aberto').length;

  const h = new Date().getHours();
  const greeting = h < 12 ? 'Bom dia' : h < 18 ? 'Boa tarde' : 'Boa noite';
  const nome = State.profile?.nome?.split(' ')[0] || Auth.user?.name?.split(' ')[0] || '';

  const main = document.getElementById('main');
  main.innerHTML = `
    <div class="ph">
      <div>
        <div class="pt">${greeting}, ${nome} 👋</div>
        <div class="ps">Fila ordenada por urgência · ${new Date().toLocaleDateString('pt-BR', { weekday:'long', day:'2-digit', month:'long' })}</div>
      </div>
      <div class="pa">
        <div class="vsw">
          <button class="vbtn on" id="vbtn-cards" onclick="switchViewOp('cards')">Cards</button>
          <button class="vbtn" id="vbtn-tabela" onclick="switchViewOp('tabela')">Tabela</button>
        </div>
        <button class="btn bg sm" onclick="Sync.full()">↺ Sincronizar</button>
      </div>
    </div>

    ${criticos > 0 ? `<div class="alert-banner r">🚨 <span>${criticos} ticket crítico com SLA em risco — ação imediata.</span></div>` : ''}
    ${aguardaResp > 0 ? `<div class="alert-banner y">💬 <span>${aguardaResp} ticket com resposta nova do cliente aguardando sua ação.</span></div>` : ''}

    <div class="stats">
      <div class="sc ${State.filters.status === 'Aberto' ? 'af' : ''}" onclick="setFilterOp('Aberto')">
        <div class="scl">Abertos</div><div class="scv cb">${abertos}</div><div class="scs">em atendimento</div>
      </div>
      <div class="sc ${State.filters.status === 'theirs' ? 'af' : ''}" onclick="setFilterOp('theirs')">
        <div class="scl">Resp. Pendente</div><div class="scv" style="color:var(--orange)">${aguardaResp}</div><div class="scs">cliente respondeu</div>
      </div>
      <div class="sc ${State.filters.status === 'Espera' ? 'af' : ''}" onclick="setFilterOp('Espera')">
        <div class="scl">Em Espera</div><div class="scv cy">${espera}</div><div class="scs">aguardando retorno</div>
      </div>
      <div class="sc ${State.filters.status === 'Terceiro' ? 'af' : ''}" onclick="setFilterOp('Terceiro')">
        <div class="scl">Ag. Terceiro</div><div class="scv cp">${terceiro}</div><div class="scs">SLA pausado</div>
      </div>
      <div class="sc ${State.filters.status === 'Concluído' ? 'af' : ''}" onclick="setFilterOp('Concluído')">
        <div class="scl">Concluídos</div><div class="scv cg">${concluidos}</div><div class="scs">hoje</div>
      </div>
    </div>

    <div class="fbar">
      <span class="fbar-label">Filtro:</span>
      ${['Todos','Aberto','theirs','Espera','Terceiro','Concluído'].map(f =>
        `<div class="chip ${State.filters.status === f ? 'on' : ''}" onclick="setFilterOp('${f}')">${f === 'theirs' ? '💬 Resp. Pendente' : f}</div>`
      ).join('')}
      <div class="vsw" style="margin-left:auto;">
        <button class="sort-chip ${State.sort === 'sla' ? 'on' : ''}" onclick="setSortOp('sla')">⏱ SLA</button>
        <button class="sort-chip ${State.sort === 'prioridade' ? 'on' : ''}" onclick="setSortOp('prioridade')">🔺 Prioridade</button>
        <button class="sort-chip ${State.sort === 'novo' ? 'on' : ''}" onclick="setSortOp('novo')">💬 Novos</button>
      </div>
      <input class="srch" value="${State.filters.search}" placeholder="🔍 ID, cliente, doc. fiscal..." oninput="setSearchOp(this.value)">
    </div>

    <div id="ticket-list">${renderTicketCards(State.getSortedTickets(applyOpFilters(myT)))}</div>
  `;
}

function applyOpFilters(tickets) {
  let t = [...tickets];
  const f = State.filters;
  if (f.status && f.status !== 'Todos') {
    if (f.status === 'theirs') t = t.filter(x => x.respostaState === 'theirs');
    else t = t.filter(x => x.status === f.status);
  }
  if (f.search) {
    const q = f.search.toLowerCase();
    t = t.filter(x =>
      x.id.toLowerCase().includes(q) ||
      x.assunto.toLowerCase().includes(q) ||
      x.cliente.toLowerCase().includes(q) ||
      (x.docFiscal || '').toLowerCase().includes(q)
    );
  }
  return t;
}

function setFilterOp(f) {
  State.filters.status = f;
  renderOperadorPainel();
}

function setSortOp(s) {
  State.sort = s;
  renderOperadorPainel();
}

function setSearchOp(q) {
  State.filters.search = q;
  const myT = State.getMyTickets();
  document.getElementById('ticket-list').innerHTML = renderTicketCards(
    State.getSortedTickets(applyOpFilters(myT))
  );
}

function switchViewOp(v) {
  document.getElementById('vbtn-cards').classList.toggle('on', v === 'cards');
  document.getElementById('vbtn-tabela').classList.toggle('on', v === 'tabela');
  const myT = State.getSortedTickets(State.getMyTickets());
  document.getElementById('ticket-list').innerHTML =
    v === 'cards' ? renderTicketCards(myT) : renderTicketTable(myT, 'operator');
}

// ── CARDS ─────────────────────────────────────────────────────
function renderTicketCards(tickets) {
  if (!tickets.length) return `<div class="empty"><div class="ei">📭</div><p>Nenhum ticket neste filtro.</p></div>`;
  return tickets.map(t => renderTicketCard(t)).join('');
}

function renderTicketCard(t) {
  const pct = SLA.pct(t);
  const slaColor = SLA.color(t);
  const slaLabel = SLA.label(t);
  const flag = PAISES[t.pais] || '🌐';
  const cl = CHECKLISTS[t.categoria] || CHECKLISTS['default'];
  const allKeys = cl.flatMap(l => l.items.map(i => i.id));
  const clChecked = allKeys.filter(k => t.checklist[k]).length;

  const rstate = t.respostaState === 'theirs'
    ? `<span class="rstate theirs">💬 Cliente respondeu</span>`
    : t.respostaState === 'new'
    ? `<span class="rstate new">🆕 Novo</span>`
    : `<span class="rstate mine">↩ Aguardando cliente</span>`;

  return `
  <div class="tcard ${t.stale ? 'stale' : ''}" onclick="openTicketPanel('${t.id}')">
    ${t.stale ? `<div class="stale-warn">⏰ Parado +${Math.round((Date.now() - new Date(t.ultimaAtv)) / 60000)}min</div>` : ''}
    <div class="tcard-inner">
      <div class="flag" title="${t.pais} · ${t.fuso}">${flag}</div>
      <div class="tcard-main">
        <div class="tcard-top">
          <span class="tcard-id">${t.id}</span>
          <span class="badge b-${t.prioridade.toLowerCase()}">${t.prioridade}</span>
          <span class="badge b-${t.status === 'Terceiro' ? 'terceiro' : t.status.toLowerCase().replace('í','i').replace('é','e')}">${t.status === 'Terceiro' ? '⏳ Ag. Terceiro' : t.status}</span>
          ${t.slaPausado ? `<span class="badge b-pausado">⏸ SLA pausado</span>` : ''}
          <span class="tag">${t.categoria}</span>
          ${rstate}
        </div>
        <div class="tcard-subject">${t.assunto}</div>
        <div class="tcard-meta">
          <span>${t.cliente}</span>
          ${t.docFiscal ? `<span>· <span style="font-family:monospace;color:var(--accent-op);">${t.docFiscal}</span></span>` : ''}
          <span>· ${t.origem}</span>
          <span>· CL: ${clChecked}/${allKeys.length}</span>
          ${t.cwId ? `<span>· <strong style="color:var(--green);">CW: ${t.cwId}</strong></span>` : ''}
        </div>
      </div>
      <div class="tcard-side">
        ${t.slaPausado ? `
          <div class="tcard-sla-time" style="color:var(--teal);">⏸ Pausado</div>
          <div class="sla-pct" style="color:var(--teal);font-size:10px;">${t.pausaMotivo || 'Ag. terceiro'}</div>
        ` : `
          <div class="tcard-sla-time" style="color:${slaColor};">${slaLabel}</div>
          <div class="sla-track"><div class="sla-fill" style="width:${pct}%;background:${slaColor};"></div></div>
          <div class="sla-pct">${pct}% do SLA</div>
        `}
        <button class="btn bp btn-sm" onclick="event.stopPropagation();openTicketPanel('${t.id}')">Abrir →</button>
      </div>
    </div>
  </div>`;
}

// ── TABLE ─────────────────────────────────────────────────────
function renderTicketTable(tickets, mode) {
  const isMgr = mode === 'manager';
  const cols = isMgr
    ? '85px 1fr 100px 70px 85px 90px 100px 70px'
    : '85px 1fr 70px 85px 90px 110px 70px';

  return `
  <div class="tbl-wrap">
    <div class="tbl-head" style="grid-template-columns:${cols};">
      <span>ID</span><span>Assunto / Cliente</span>
      ${isMgr ? '<span>Operador</span>' : ''}
      <span>País</span><span>Prioridade</span><span>Status</span><span>SLA</span><span>Ação</span>
    </div>
    ${tickets.map(t => {
      const pct = SLA.pct(t);
      const sc = SLA.color(t);
      const label = SLA.label(t);
      const flag = PAISES[t.pais] || '🌐';
      const op = State.operadores.find(o => o.email === t.operador);
      return `<div class="tbl-row" style="grid-template-columns:${cols};" onclick="openTicketPanel('${t.id}')">
        <div class="tcard-id">${t.id}</div>
        <div>
          <div style="font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px;">${t.assunto}</div>
          <div style="font-size:11px;color:var(--text-muted);">${t.cliente}</div>
        </div>
        ${isMgr ? `<div><span style="font-size:11px;">${op ? `<span style="background:${op.cor};border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:8px;font-weight:800;color:#fff;margin-right:4px;">${op.sigla}</span>${op.nome.split(' ')[0]}` : t.operadorNome}</span></div>` : ''}
        <div>${flag} <span style="font-size:10px;color:var(--text-muted);">${t.pais}</span></div>
        <div><span class="badge b-${t.prioridade.toLowerCase()}">${t.prioridade}</span></div>
        <div><span class="badge b-${t.status === 'Terceiro' ? 'terceiro' : t.status.toLowerCase().replace('í','i').replace('é','e')}">${t.status === 'Terceiro' ? 'Ag.Terc' : t.status}</span></div>
        <div>${t.slaPausado
          ? `<span style="color:var(--teal);font-size:10px;font-weight:700;">⏸ Pausado</span>`
          : `<div style="color:${sc};font-size:11px;font-weight:700;">${label}</div>
             <div class="sla-track" style="width:90px;"><div class="sla-fill" style="width:${pct}%;background:${sc};"></div></div>`
        }</div>
        <div><button class="btn bg btn-xs" onclick="event.stopPropagation();openTicketPanel('${t.id}')">Ver</button></div>
      </div>`;
    }).join('')}
  </div>
  <div style="display:flex;justify-content:flex-end;margin-top:10px;">
    <button class="btn bg sm" onclick="exportCSV()">⬇ Exportar CSV</button>
  </div>`;
}

// ── BUSCA GLOBAL ──────────────────────────────────────────────
function renderBuscaGlobal() {
  document.getElementById('main').innerHTML = `
    <div class="ph"><div><div class="pt">🔍 Busca Global</div><div class="ps">Busque por cliente, documento fiscal, ID CargoWise ou ID de ticket</div></div></div>
    <input class="srch" style="width:100%;font-size:14px;padding:10px 14px;margin-bottom:16px;" id="gs" placeholder="CNPJ, RUT, CUIT, NIT, EIN, razão social, TKT-XXXX, CW-ID..." oninput="doGlobalSearch(this.value)">
    <div id="gs-results"><div class="empty"><div class="ei">🔍</div><p>Digite para buscar.</p></div></div>
  `;
}

function doGlobalSearch(q) {
  if (!q || q.length < 2) { document.getElementById('gs-results').innerHTML = ''; return; }
  const results = State.tickets.filter(t =>
    t.id.toLowerCase().includes(q.toLowerCase()) ||
    t.assunto.toLowerCase().includes(q.toLowerCase()) ||
    t.cliente.toLowerCase().includes(q.toLowerCase()) ||
    (t.docFiscal || '').toLowerCase().includes(q.toLowerCase()) ||
    (t.cwId || '').toLowerCase().includes(q.toLowerCase())
  );
  document.getElementById('gs-results').innerHTML = results.length
    ? renderTicketCards(State.getSortedTickets(results))
    : `<div class="empty"><div class="ei">🚫</div><p>Nenhum resultado para "<strong>${q}</strong>".</p></div>`;
}

// ── DETAIL PANEL ──────────────────────────────────────────────
function openTicketPanel(id) {
  State.activeTicketId = id;
  State.activeTab = 'info';
  renderPanel();
  document.getElementById('overlay').classList.add('open');
}

function renderPanel() {
  const t = State.tickets.find(x => x.id === State.activeTicketId);
  if (!t) return;
  const flag = PAISES[t.pais] || '🌐';
  const sc = SLA.color(t);

  document.getElementById('phdr').innerHTML = `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
      <span class="tcard-id" style="font-size:13px;">${t.id}</span>
      <span class="badge b-${t.prioridade.toLowerCase()}">${t.prioridade}</span>
      <span class="badge b-${t.status === 'Terceiro' ? 'terceiro' : t.status.toLowerCase().replace('í','i').replace('é','e')}">${t.status === 'Terceiro' ? '⏳ Ag. Terceiro' : t.status}</span>
      ${t.slaPausado ? `<span class="badge b-pausado">⏸ SLA pausado</span>` : ''}
      <button class="panel-close" onclick="closePanelUI()" style="margin-left:auto;">✕</button>
    </div>
    <div style="font-size:14px;font-weight:800;margin-bottom:3px;">${flag} ${t.assunto}</div>
    <div style="font-size:11px;color:var(--text-muted);">${t.cliente} · ${t.remetente} · ${t.fuso}</div>
  `;

  document.getElementById('ptabs').innerHTML = ['info','checklist','historico','responder','auditoria'].map(tab =>
    `<div class="ptab ${State.activeTab === tab ? 'on' : ''}" onclick="switchTabOp('${tab}')">
      ${{ info:'ℹ️ Info', checklist:'✅ Checklist', historico:'💬 Histórico', responder:'📧 Responder', auditoria:'🔍 Auditoria' }[tab]}
    </div>`
  ).join('');

  renderPanelBodyOp(t);
  renderPanelFooterOp(t);
}

function switchTabOp(tab) {
  State.activeTab = tab;
  renderPanel();
}

// Expor para o UI.renderPanelBody
window.renderPanelBody = () => { renderPanel(); };

function renderPanelBodyOp(t) {
  const pct = SLA.pct(t);
  const sc = SLA.color(t);
  const restante = t.slaLimite - t.slaDecorrido;
  let html = '';

  if (State.activeTab === 'info') {
    html = `
      <div class="ds-title">SLA</div>
      <div class="sla-box">
        <div class="sla-box-top">
          <span class="sla-box-label">${t.slaPausado ? 'SLA pausado' : 'Tempo utilizado'}</span>
          <span class="sla-box-time" style="color:${sc};">${t.slaPausado ? '⏸ Pausado' : restante > 0 ? `${restante}min restantes` : 'VENCIDO'}</span>
        </div>
        ${!t.slaPausado ? `<div class="sla-big"><div class="sla-big-fill" style="width:${pct}%;background:${sc};"></div></div>` : ''}
        <div class="sla-box-sub">
          <span>${!t.slaPausado ? `${t.slaDecorrido}min / ${t.slaLimite}min (${pct}%)` : 'SLA não corre durante dependência externa'}</span>
          <span>Prioridade ${t.prioridade}</span>
        </div>
        ${t.slaPausado ? `<div class="pause-note">⏸ ${t.pausaMotivo || 'Ag. dependência externa'}</div>` : ''}
      </div>

      <div class="ds-title">Informações</div>
      <div class="info-grid">
        <div class="info-item"><label>Cliente</label><span>${t.cliente}</span></div>
        <div class="info-item"><label>Remetente</label><span>${t.remetente}</span></div>
        <div class="info-item"><label>Doc. Fiscal</label><span style="font-family:monospace;color:var(--accent-op);">${t.docFiscal || '—'}</span></div>
        <div class="info-item"><label>País · Fuso</label><span>${PAISES[t.pais] || ''} ${t.pais} · ${t.fuso}</span></div>
        <div class="info-item"><label>Categoria</label><span>${t.categoria}</span></div>
        <div class="info-item"><label>Origem</label><span>${t.origem}</span></div>
        <div class="info-item"><label>Abertura</label><span>${new Date(t.abertura).toLocaleString('pt-BR')}</span></div>
        <div class="info-item"><label>Última atividade</label><span>${new Date(t.ultimaAtv).toLocaleString('pt-BR')}</span></div>
        ${t.terceiro ? `<div class="info-item" style="grid-column:1/-1;"><label>Aguardando</label><span style="color:var(--purple);">${t.terceiro}${t.terceiroContato ? ' — ' + t.terceiroContato : ''}</span></div>` : ''}
      </div>

      <div class="ds-title">Sistema de Destino</div>
      <div style="display:flex;gap:7px;margin-bottom:16px;">
        ${['CargoWise','CraftHub'].map(s =>
          `<div class="sys-chip ${(t.sistemaDestino || []).includes(s) ? 'on' : ''}"
           onclick="toggleSistemaOp('${t.id}','${s}')">${s}</div>`
        ).join('')}
      </div>

      <div class="cw-field">
        <label>ID CargoWise <span style="color:var(--red)">*</span> — obrigatório para concluir</label>
        <div class="cw-input-wrap">
          <input class="cw-input" id="cw-${t.id}" value="${t.cwId || ''}" placeholder="Ex: CW-MX-00471">
          <button class="btn bg sm" onclick="Actions.saveCwId('${t.id}', document.getElementById('cw-${t.id}').value)">Salvar</button>
        </div>
        ${t.cwId ? `<div style="margin-top:5px;font-size:11px;color:var(--green);">✓ ID registrado: <strong style="font-family:monospace;">${t.cwId}</strong></div>` : ''}
      </div>

      <div class="ds-title">Anexos (${(t.anexos || []).length})</div>
      <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:16px;">
        ${(t.anexos || []).length
          ? t.anexos.map(a => `
            <div style="display:flex;align-items:center;gap:8px;background:var(--surface2);border-radius:var(--r-sm);padding:8px 10px;">
              <span style="font-size:15px;">📎</span>
              <span style="font-size:12px;font-weight:600;flex:1;">${a}</span>
              <button class="btn bg btn-xs" onclick="UI.showToast('Abrindo ${a}','s')">Ver</button>
            </div>`).join('')
          : '<div style="font-size:12px;color:var(--text-muted);">Nenhum anexo recebido.</div>'}
      </div>

      <div class="ds-title">Nota Interna</div>
      <div style="margin-bottom:6px;font-size:10px;color:var(--text-muted);font-weight:700;">⚠️ Visível apenas para a equipe CoreCad</div>
      ${(t.notas || []).map(n => `
        <div style="background:var(--yellow-soft);border:1px solid rgba(245,158,11,.15);border-radius:var(--r-sm);padding:9px 11px;margin-bottom:7px;">
          <div style="font-size:10px;font-weight:800;color:var(--yellow);margin-bottom:4px;">🔒 Nota Interna</div>
          <div style="font-size:12px;">${n}</div>
        </div>`).join('')}
      <textarea class="reply-ta internal-ta" id="nota-ta-${t.id}" rows="2" placeholder="Adicionar nota interna..."></textarea>
      <button class="btn by sm" style="margin-top:7px;"
        onclick="Actions.addNota('${t.id}', document.getElementById('nota-ta-${t.id}').value)">+ Salvar nota interna</button>
    `;
  }

  else if (State.activeTab === 'checklist') {
    const cl = CHECKLISTS[t.categoria] || CHECKLISTS['default'];
    const allKeys = cl.flatMap(l => l.items.map(i => i.id));
    const checked = allKeys.filter(k => t.checklist[k]).length;
    const pctCl = allKeys.length ? Math.round(checked / allKeys.length * 100) : 0;

    html = `
      <div class="cl-progress">
        <span style="font-size:12px;font-weight:700;">Progresso</span>
        <div class="cl-progress-bar"><div class="cl-progress-fill" style="width:${pctCl}%;"></div></div>
        <span class="cl-progress-text">${checked}/${allKeys.length} · ${pctCl}%</span>
      </div>
      ${cl.map(layer => `
        <div class="cl-layer">
          <div class="cl-layer-title">${layer.layer}</div>
          ${layer.items.map(item => `
            <div class="cl-item ${t.checklist[item.id] ? 'done' : ''}" onclick="Actions.toggleChecklist('${t.id}','${item.id}')">
              <div class="cl-check"></div>
              <div style="flex:1;">
                <div class="cl-label">${item.label}</div>
                ${item.sub ? `<div class="cl-sub">${item.sub}</div>` : ''}
              </div>
              ${item.required ? `<div class="cl-req">Obrigatório</div>` : ''}
            </div>`).join('')}
        </div>`).join('')}
      ${pctCl < 100 ? `<div style="background:var(--yellow-soft);border:1px solid rgba(245,158,11,.2);border-radius:var(--r-sm);padding:10px 12px;font-size:11px;color:var(--yellow);font-weight:600;">⚠️ Checklist incompleto. Ao concluir, você assumirá responsabilidade documentada pelos itens pendentes.</div>` : ''}
    `;
  }

  else if (State.activeTab === 'historico') {
    html = `
      <div class="hist-list">
        ${(t.historico || []).map(h => {
          const isSys = h.tipo === 'sistema';
          const isInt = h.tipo === 'interno';
          const sigla = h.de === 'op'
            ? (State.operadores.find(o => o.nome === h.nome)?.sigla || h.nome?.substring(0,2).toUpperCase() || 'OP')
            : h.nome?.split(' ').map(w => w[0]).join('').substring(0,2).toUpperCase() || 'CL';
          return `<div class="hist-item">
            <div class="hist-av ${isSys ? 'sys' : h.de}">${isSys ? '⚙️' : sigla}</div>
            <div class="hist-content">
              <div class="hist-meta"><strong>${h.nome}</strong> · ${h.hora || ''}</div>
              <div class="hist-bubble ${isSys ? 'sys' : h.de === 'op' ? 'op' : ''} ${isInt ? 'internal' : ''}">
                ${isInt ? `<div class="internal-label">🔒 Nota interna</div>` : ''}
                ${h.texto}
              </div>
            </div>
          </div>`;
        }).join('') || '<div class="empty"><div class="ei">💬</div><p>Sem histórico ainda.</p></div>'}
      </div>
    `;
  }

  else if (State.activeTab === 'responder') {
    html = `
      <div style="margin-bottom:12px;font-size:12px;color:var(--text-muted);">Resposta enviada por e-mail para <strong style="color:var(--text);">${t.remetente}</strong></div>
      <div class="ds-title">Template</div>
      <select class="template-select" id="tmpl-${t.id}" onchange="applyTemplateOp('${t.id}')">
        <option value="">— Selecionar template —</option>
        <option value="recebido">✅ Confirmação de recebimento</option>
        <option value="doc_pendente">📄 Documentação pendente</option>
        <option value="espera_gov">⏳ Aguardando Governance</option>
        <option value="concluido">🎉 Cadastro concluído</option>
      </select>
      <div class="ds-title" style="margin-top:14px;">Resposta</div>
      <textarea class="reply-ta" id="reply-ta-${t.id}" rows="7" placeholder="Digite a resposta para o cliente..."></textarea>
      <div class="reply-footer">
        <button class="btn bg sm" onclick="UI.showToast('Rascunho salvo','s')">Salvar rascunho</button>
        <button class="btn bp" onclick="Actions.sendReply('${t.id}', document.getElementById('reply-ta-${t.id}').value)">📧 Enviar</button>
      </div>
    `;
  }

  else if (State.activeTab === 'auditoria') {
    html = `
      <div class="ds-title">Log de Auditoria</div>
      <div class="audit-list">
        ${[
          { cor: 'var(--accent-op)', texto: `Ticket criado · SLA iniciado`, time: new Date(t.abertura).toLocaleString('pt-BR') },
          ...(t.historico || []).filter(h => h.de === 'op' || h.tipo === 'sistema').map(h => ({
            cor: h.tipo === 'sistema' ? 'var(--text-muted)' : 'var(--accent-op)',
            texto: h.tipo === 'sistema' ? h.texto : `Mensagem enviada por <strong>${h.nome}</strong>`,
            time: h.hora || '',
          })),
          ...(t.slaPausado ? [{ cor: 'var(--teal)', texto: `SLA pausado · ${t.pausaMotivo}`, time: '' }] : []),
          ...(t.cwId ? [{ cor: 'var(--green)', texto: `ID CargoWise registrado: <strong style="font-family:monospace">${t.cwId}</strong>`, time: '' }] : []),
          ...(t.status === 'Concluído' ? [{ cor: 'var(--green)', texto: `Ticket concluído`, time: new Date(t.ultimaAtv).toLocaleString('pt-BR') }] : []),
        ].map(a => `
          <div class="audit-item">
            <div class="audit-dot" style="background:${a.cor};"></div>
            <div><div class="audit-text">${a.texto}</div>${a.time ? `<div class="audit-time">${a.time}</div>` : ''}</div>
          </div>`).join('')}
      </div>
    `;
  }

  document.getElementById('pbody').innerHTML = html;
}

function renderPanelFooterOp(t) {
  const cl = CHECKLISTS[t.categoria] || CHECKLISTS['default'];
  const requiredKeys = cl.flatMap(l => l.items.filter(i => i.required).map(i => i.id));
  const requiredOk = requiredKeys.every(k => t.checklist[k]);

  let btns = '';
  if (t.status === 'Aberto') {
    btns = `
      <button class="btn bgrn" onclick="concludeTicket('${t.id}')">✅ Concluir</button>
      <button class="btn by" onclick="Actions.changeStatus('${t.id}','Espera')">⏸️ Em Espera</button>
      <button class="btn bg" onclick="pauseSlaPrompt('${t.id}')">🔗 Ag. Terceiro</button>
    `;
  } else if (t.status === 'Espera') {
    btns = `<button class="btn bp" onclick="Actions.changeStatus('${t.id}','Aberto')">▶️ Retomar</button>`;
  } else if (t.status === 'Terceiro') {
    btns = `<button class="btn bp" onclick="Actions.resumeFromThird('${t.id}')">▶️ Retomar (SLA)</button>`;
  } else if (t.status === 'Concluído') {
    btns = `<button class="btn bg" onclick="Actions.changeStatus('${t.id}','Aberto')">↺ Reabrir</button>`;
  }

  document.getElementById('pfooter').innerHTML = `
    ${btns}
    <button class="btn bg" onclick="openRedistRequest('${t.id}')">↗️ Solicitar redistribuição</button>
    <button class="btn bg" style="margin-left:auto;color:var(--text-dim);" onclick="closePanelUI()">Fechar</button>
  `;
}

// ── AÇÕES DO PAINEL ───────────────────────────────────────────
function toggleSistemaOp(id, sys) {
  const t = State.tickets.find(x => x.id === id);
  if (!t) return;
  const arr = [...(t.sistemaDestino || [])];
  const idx = arr.indexOf(sys);
  if (idx > -1) arr.splice(idx, 1); else arr.push(sys);
  t.sistemaDestino = arr;
  Actions.saveSistemaDestino(id, arr);
  renderPanel();
}

function applyTemplateOp(id) {
  const t = State.tickets.find(x => x.id === id);
  const sel = document.getElementById('tmpl-' + id);
  const ta = document.getElementById('reply-ta-' + id);
  if (!sel || !ta || !t) return;
  ta.value = (TEMPLATES[sel.value] || '')
    .replace('{ID}', id)
    .replace('{CW_ID}', t.cwId || '[ID CargoWise]');
}

async function concludeTicket(id) {
  const t = State.tickets.find(x => x.id === id);
  if (!t) return;
  if (!t.cwId) {
    UI.showToast('⚠️ Preencha o ID CargoWise antes de concluir', 'e');
    State.activeTab = 'info';
    renderPanel();
    return;
  }
  const cl = CHECKLISTS[t.categoria] || CHECKLISTS['default'];
  const requiredKeys = cl.flatMap(l => l.items.filter(i => i.required).map(i => i.id));
  const missingRequired = requiredKeys.filter(k => !t.checklist[k]);
  if (missingRequired.length) {
    UI.showToast('⚠️ Itens obrigatórios do checklist incompletos', 'e');
    State.activeTab = 'checklist';
    renderPanel();
    return;
  }
  const allKeys = cl.flatMap(l => l.items.map(i => i.id));
  const allDone = allKeys.every(k => t.checklist[k]);
  if (!allDone && !confirm('Checklist incompleto. Confirmar conclusão e assumir responsabilidade pelos itens pendentes?')) return;
  await Actions.changeStatus(id, 'Concluído');
  closePanelUI();
}

function pauseSlaPrompt(id) {
  const motivo = prompt('Motivo para Aguardando Terceiro\n(ex: Governance — Daniela Milan):');
  if (!motivo) return;
  Actions.pauseSla(id, motivo);
  closePanelUI();
}

function openRedistRequest(id) {
  const ops = State.operadores.filter(o => o.email !== Auth.user?.email && o.role === 'operador');
  if (!ops.length) { UI.showToast('Nenhum operador disponível para redistribuição', 'w'); return; }
  const opcoes = ops.map(o => `${o.sigla} — ${o.nome} (${State.tickets.filter(t=>t.operador===o.email&&t.status!=='Concluído').length} ativos)`).join('\n');
  const escolha = prompt(`Solicitação de redistribuição para ${id}.\n\nSelecione o operador (número):\n${ops.map((o,i)=>`${i+1}. ${o.nome}`).join('\n')}`);
  if (!escolha) return;
  const idx = parseInt(escolha) - 1;
  if (idx < 0 || idx >= ops.length) { UI.showToast('Opção inválida', 'w'); return; }
  const motivo = prompt('Motivo da solicitação:');
  if (!motivo) return;
  UI.showToast(`Solicitação enviada para aprovação da gerência · Destino: ${ops[idx].nome}`, 's');
}

function closePanelUI() {
  document.getElementById('overlay').classList.remove('open');
}

// ── EXPORT CSV ────────────────────────────────────────────────
function exportCSV() {
  const rows = [['ID','Assunto','Cliente','País','Prioridade','Status','SLA%','CW-ID']];
  State.getMyTickets().forEach(t => {
    const pct = t.slaPausado ? 'Pausado' : SLA.pct(t) + '%';
    rows.push([t.id, t.assunto, t.cliente, t.pais, t.prioridade, t.status, pct, t.cwId || '']);
  });
  const csv = rows.map(r => r.map(v => `"${v}"`).join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = `corecad-operador-${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
  UI.showToast('CSV exportado ✓', 's');
}

// Expor globalmente
window.renderOperadorPainel = renderOperadorPainel;
window.renderTicketCards = renderTicketCards;
window.renderTicketTable = renderTicketTable;
window.renderBuscaGlobal = renderBuscaGlobal;
window.doGlobalSearch = doGlobalSearch;
window.openTicketPanel = openTicketPanel;
window.renderPanel = renderPanel;
window.switchTabOp = switchTabOp;
window.setFilterOp = setFilterOp;
window.setSortOp = setSortOp;
window.setSearchOp = setSearchOp;
window.switchViewOp = switchViewOp;
window.toggleSistemaOp = toggleSistemaOp;
window.applyTemplateOp = applyTemplateOp;
window.concludeTicket = concludeTicket;
window.pauseSlaPrompt = pauseSlaPrompt;
window.openRedistRequest = openRedistRequest;
window.closePanelUI = closePanelUI;
window.exportCSV = exportCSV;
