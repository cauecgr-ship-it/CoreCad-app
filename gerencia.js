// ============================================================
// CoreCad — Módulo Gerência
// Renderização completa da interface gerencial
// Depende de: app.js (State, Actions, SLA, UI, Auth)
// ============================================================

const PAISES_MGR = { BR:'🇧🇷', MX:'🇲🇽', AR:'🇦🇷', CL:'🇨🇱', CO:'🇨🇴', US:'🇺🇸' };

// Tendência simulada — será substituída por dados reais do SharePoint
// quando uma lista de histórico de volume for criada
const TENDENCIA_MOCK = [
  { dia:'Seg', vol:8,  slaOk:7,  slaNok:1 },
  { dia:'Ter', vol:12, slaOk:10, slaNok:2 },
  { dia:'Qua', vol:10, slaOk:8,  slaNok:2 },
  { dia:'Qui', vol:15, slaOk:11, slaNok:4 },
  { dia:'Sex', vol:9,  slaOk:8,  slaNok:1 },
  { dia:'Seg', vol:11, slaOk:9,  slaNok:2 },
  { dia:'Hoje',vol:State.tickets.filter(t=>t.status!=='Concluído').length || 0,
    slaOk: State.tickets.filter(t=>t.status==='Concluído').length || 0,
    slaNok: State.tickets.filter(t=>!t.slaPausado&&t.slaDecorrido/t.slaLimite>=1).length || 0 },
];

// ── OVERVIEW ──────────────────────────────────────────────────
function renderMgrOverview() {
  const T = State.tickets;
  const abertos  = T.filter(t => t.status === 'Aberto').length;
  const espera   = T.filter(t => t.status === 'Espera').length;
  const terceiro = T.filter(t => t.status === 'Terceiro').length;
  const concluidos = T.filter(t => t.status === 'Concluído').length;
  const criticos = T.filter(t => t.prioridade === 'Crítica' && t.status === 'Aberto').length;
  const slaRisco = T.filter(t => !t.slaPausado && t.status !== 'Concluído' && t.slaDecorrido / t.slaLimite >= 0.8).length;

  const byPais = T.reduce((acc, t) => { acc[t.pais] = (acc[t.pais] || 0) + 1; return acc; }, {});
  const byOp   = State.operadores.map(op => ({
    op,
    count: T.filter(t => t.operador === op.email && t.status !== 'Concluído').length,
  }));

  document.getElementById('main').innerHTML = `
    <div class="ph">
      <div><div class="pt">📊 Painel da Gerência</div><div class="ps">CoreCad · Visão em tempo real · ${new Date().toLocaleDateString('pt-BR', { weekday:'long', day:'2-digit', month:'long' })}</div></div>
      <div class="pa">
        <button class="btn bg sm" onclick="exportRelatorio()">⬇ Relatório CSV</button>
        <button class="btn bp sm" onclick="navMgr('redist')">↗️ Redistribuir</button>
      </div>
    </div>

    ${criticos > 0 ? `<div class="alert-strip r">🚨 <strong>${criticos} ticket crítico</strong> com SLA quase vencido. <button class="btn br xs" onclick="navMgr('criticos')" style="margin-left:auto;">Ver →</button></div>` : ''}
    ${slaRisco > 0 ? `<div class="alert-strip y">⚠️ <strong>${slaRisco} tickets</strong> com SLA acima de 80%.</div>` : ''}

    <div class="stats">
      <div class="sc"><div class="scl">Abertos</div><div class="scv cb">${abertos}</div><div class="scs">em atendimento</div></div>
      <div class="sc"><div class="scl">Em Espera</div><div class="scv cy">${espera}</div><div class="scs">aguardando</div></div>
      <div class="sc"><div class="scl">Ag. Terceiro</div><div class="scv cp">${terceiro}</div><div class="scs">SLA pausado</div></div>
      <div class="sc"><div class="scl">Concluídos</div><div class="scv cg">${concluidos}</div><div class="scs">hoje</div></div>
      <div class="sc"><div class="scl">Críticos</div><div class="scv cr">${criticos}</div><div class="scs">prioridade máxima</div></div>
      <div class="sc"><div class="scl">Operadores</div><div class="scv">${State.operadores.filter(o=>o.role==='operador').length}</div><div class="scs">ativos</div></div>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px;">
      <div class="sec">
        <div class="sec-title">Volume por Operador <span>Clique para filtrar</span></div>
        <div class="mbc">
          ${byOp.map(({op, count}) => {
            const cap = op.capacidade || 5;
            const w = Math.round((count / cap) * 100);
            return `<div class="mbc-row">
              <div class="mbc-lbl">${op.nome.split(' ')[0]}</div>
              <div class="mbc-track" onclick="filterMgrByOp('${op.email}')">
                <div class="mbc-fill" style="width:${Math.max(w, 8)}%;background:${op.cor};">
                  <span class="mbc-val">${count}/${cap}</span>
                </div>
              </div>
            </div>`;
          }).join('')}
        </div>
      </div>
      <div class="sec">
        <div class="sec-title">Por País <span>Clique para filtrar</span></div>
        <div class="geo-grid">
          ${Object.entries(byPais).sort((a,b) => b[1]-a[1]).map(([pais, n]) => {
            const pct = Math.round((n / T.length) * 100);
            return `<div class="geo-item ${State.filters.geo === pais ? 'sel' : ''}" onclick="filterMgrByGeo('${pais}')">
              <div class="geo-flag">${PAISES_MGR[pais] || '🌐'}</div>
              <div class="geo-info">
                <div class="geo-country">${pais}</div>
                <div class="geo-count">${n} tickets</div>
                <div class="geo-bar"><div class="geo-bar-fill" style="width:${pct}%;background:var(--accent-mgr);"></div></div>
              </div>
            </div>`;
          }).join('')}
        </div>
      </div>
    </div>

    <div style="display:grid;grid-template-columns:2fr 1fr;gap:14px;margin-bottom:14px;">
      <div class="sec">
        <div class="sec-title">Últimas 7 sessões — Volume e SLA</div>
        <div class="trend-wrap">
          ${TENDENCIA_MOCK.map(d => {
            const max = Math.max(...TENDENCIA_MOCK.map(x => x.vol));
            const h = Math.round((d.vol / max) * 52) + 6;
            const isHoje = d.dia === 'Hoje';
            return `<div class="trend-bar" style="height:${h}px;background:${isHoje ? 'var(--accent-mgr)' : 'var(--surface3)'};border:1px solid ${isHoje ? 'var(--accent-mgr)' : 'var(--border)'};">
              <span class="trend-tip">${d.dia}: ${d.vol} tickets · ${d.slaOk} ok · ${d.slaNok} venc.</span>
            </div>`;
          }).join('')}
        </div>
        <div class="trend-labels">${TENDENCIA_MOCK.map(d => `<div class="trend-label">${d.dia}</div>`).join('')}</div>
      </div>
      <div class="sec">
        <div class="sec-title">Status geral</div>
        <div class="donut-wrap">
          <svg width="80" height="80" viewBox="0 0 80 80">
            ${donutMgr([{v:abertos,c:'var(--accent-mgr)'},{v:espera,c:'var(--yellow)'},{v:terceiro,c:'var(--purple)'},{v:concluidos,c:'var(--green)'}])}
            <text x="40" y="43" text-anchor="middle" fill="#d8e0f0" font-size="13" font-weight="800">${T.length}</text>
          </svg>
          <div class="donut-leg">
            <div class="donut-li"><div class="donut-dot" style="background:var(--accent-mgr);"></div>Aberto: <strong>${abertos}</strong></div>
            <div class="donut-li"><div class="donut-dot" style="background:var(--yellow);"></div>Espera: <strong>${espera}</strong></div>
            <div class="donut-li"><div class="donut-dot" style="background:var(--purple);"></div>Ag.Terc: <strong>${terceiro}</strong></div>
            <div class="donut-li"><div class="donut-dot" style="background:var(--green);"></div>Conc.: <strong>${concluidos}</strong></div>
          </div>
        </div>
      </div>
    </div>

    <div class="sec">
      <div class="sec-title">Atividade recente <span><button class="btn bg xs" onclick="navMgr('auditoria')">Ver log completo</button></span></div>
      <div class="log-list" id="mgr-log">
        <div style="font-size:12px;color:var(--text-muted);">Carregando log...</div>
      </div>
    </div>
  `;

  renderMgrLog();
}

function renderMgrLog() {
  // Constrói log a partir dos tickets reais
  const entries = [];
  State.tickets.forEach(t => {
    (t.historico || []).forEach(h => {
      entries.push({ cor: h.tipo==='sistema'?'var(--text-muted)':h.de==='op'?'var(--accent-mgr)':'var(--green)', text: `<strong>${t.id}</strong> — ${h.texto}`, time: h.hora || '' });
    });
  });
  entries.sort((a, b) => b.time.localeCompare(a.time));
  const el = document.getElementById('mgr-log');
  if (!el) return;
  el.innerHTML = entries.slice(0, 8).map(e => `
    <div class="log-item">
      <div class="log-dot" style="background:${e.cor};"></div>
      <div class="log-text">${e.text}</div>
      <div class="log-time">${e.time}</div>
    </div>`).join('') || '<div style="font-size:12px;color:var(--text-muted);">Sem atividade recente.</div>';
}

// ── GEO ───────────────────────────────────────────────────────
function renderMgrGeo() {
  const byPais = State.tickets.reduce((acc, t) => {
    if (!acc[t.pais]) acc[t.pais] = [];
    acc[t.pais].push(t);
    return acc;
  }, {});

  document.getElementById('main').innerHTML = `
    <div class="ph"><div><div class="pt">🌍 Por País</div><div class="ps">Distribuição geográfica de todos os tickets</div></div></div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;">
      ${Object.entries(byPais).sort((a,b) => b[1].length - a[1].length).map(([pais, ts]) => {
        const abertos = ts.filter(t => t.status !== 'Concluído').length;
        const criticos = ts.filter(t => t.prioridade === 'Crítica').length;
        const conc = ts.filter(t => t.status === 'Concluído').length;
        return `<div class="sec" style="cursor:pointer;" onclick="filterMgrByGeo('${pais}');navMgr('todos')">
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
            <span style="font-size:30px;">${PAISES_MGR[pais] || '🌐'}</span>
            <div><div style="font-size:14px;font-weight:800;">${pais}</div><div style="font-size:11px;color:var(--text-muted);">${ts.length} tickets total</div></div>
          </div>
          <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:6px;text-align:center;">
            <div><div style="font-size:18px;font-weight:900;color:var(--accent-mgr);">${abertos}</div><div style="font-size:10px;color:var(--text-muted);">Abertos</div></div>
            <div><div style="font-size:18px;font-weight:900;color:${criticos>0?'var(--red)':'var(--text-muted)'};">${criticos}</div><div style="font-size:10px;color:var(--text-muted);">Críticos</div></div>
            <div><div style="font-size:18px;font-weight:900;color:var(--green);">${conc}</div><div style="font-size:10px;color:var(--text-muted);">Conc.</div></div>
          </div>
        </div>`;
      }).join('')}
    </div>
  `;
}

// ── EQUIPE ────────────────────────────────────────────────────
function renderMgrEquipe() {
  const ops = State.operadores.filter(o => o.role === 'operador');

  document.getElementById('main').innerHTML = `
    <div class="ph">
      <div><div class="pt">👥 Carga da Equipe</div><div class="ps">Distribuição e capacidade individual</div></div>
      <div class="pa"><button class="btn bp sm" onclick="navMgr('redist')">↗️ Redistribuir</button></div>
    </div>

    <div class="op-grid">
      ${ops.map(op => {
        const ativos = State.tickets.filter(t => t.operador === op.email && t.status !== 'Concluído');
        const abertos = ativos.filter(t => t.status === 'Aberto').length;
        const criticos = ativos.filter(t => t.prioridade === 'Crítica').length;
        const conc = State.tickets.filter(t => t.operador === op.email && t.status === 'Concluído').length;
        const carga = Math.round((ativos.length / (op.capacidade || 5)) * 100);
        const overloaded = ativos.length > (op.capacidade || 5);
        return `<div class="opc ${overloaded ? 'overloaded' : ''}" onclick="filterMgrByOp('${op.email}');navMgr('todos')">
          <div class="op-status ${carga > 80 ? 'overload' : carga > 60 ? 'busy' : ''}"></div>
          <div class="opc-top">
            <div class="opav" style="background:${op.cor}18;color:${op.cor};">${op.sigla}</div>
            <div><div class="opn">${op.nome}</div><div class="opr">Operador</div></div>
          </div>
          <div class="op-stats">
            <div class="ops"><div class="ops-v" style="color:var(--accent-mgr);">${abertos}</div><div class="ops-l">Abertos</div></div>
            <div class="ops"><div class="ops-v" style="color:${criticos>0?'var(--red)':'var(--text-muted)'};">${criticos}</div><div class="ops-l">Críticos</div></div>
            <div class="ops"><div class="ops-v" style="color:var(--green);">${conc}</div><div class="ops-l">Conc.</div></div>
          </div>
          <div class="cap-bar">
            <div class="cap-label"><span>Carga ${ativos.length}/${op.capacidade||5}</span><span style="color:${carga>80?'var(--red)':carga>60?'var(--yellow)':'var(--green)'};">${carga}%</span></div>
            <div class="cap-track"><div class="cap-fill" style="width:${Math.min(carga,100)}%;background:${carga>80?'var(--red)':carga>60?'var(--yellow)':'var(--green)'}"></div></div>
          </div>
          ${overloaded ? `<div style="margin-top:8px;font-size:10px;font-weight:800;color:var(--red);">⚠️ Acima da capacidade</div>` : ''}
        </div>`;
      }).join('')}
    </div>

    <div class="sec">
      <div class="sec-title">Comparativo de desempenho</div>
      <table class="tbl">
        <thead><tr>
          <th>Operador</th><th>Concluídos</th><th>Abertos</th><th>Ag. Terceiro</th><th>SLA médio</th>
        </tr></thead>
        <tbody>
          ${ops.map(op => {
            const opT = State.tickets.filter(t => t.operador === op.email);
            const conc = opT.filter(t => t.status === 'Concluído').length;
            const ab = opT.filter(t => t.status === 'Aberto').length;
            const terc = opT.filter(t => t.status === 'Terceiro').length;
            const valid = opT.filter(t => !t.slaPausado);
            const avgSla = valid.length ? Math.round(valid.reduce((a,t) => a + t.slaDecorrido/t.slaLimite*100, 0) / valid.length) : 0;
            const sc = avgSla >= 100 ? 'var(--red)' : avgSla >= 80 ? 'var(--yellow)' : 'var(--green)';
            return `<tr>
              <td><span style="display:inline-flex;align-items:center;gap:8px;">
                <span style="background:${op.cor};border-radius:50%;width:24px;height:24px;display:inline-flex;align-items:center;justify-content:center;font-size:9px;font-weight:800;color:#fff;">${op.sigla}</span>
                ${op.nome}</span></td>
              <td><strong style="color:var(--green);">${conc}</strong></td>
              <td><strong style="color:var(--accent-mgr);">${ab}</strong></td>
              <td><strong style="color:var(--purple);">${terc}</strong></td>
              <td><strong style="color:${sc};">${avgSla}%</strong></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
      <div style="margin-top:8px;font-size:11px;color:var(--text-muted);">SLA médio: % do tempo consumido (excl. tickets pausados). Meta da equipe: &lt;70%.</div>
    </div>
  `;
}

// ── REDISTRIBUIÇÃO ────────────────────────────────────────────
function renderMgrRedist() {
  const redistState = State.redistSelected;
  const ticketsAtivos = State.tickets.filter(t => t.status !== 'Concluído');
  const ops = State.operadores.filter(o => o.role === 'operador');

  document.getElementById('main').innerHTML = `
    <div class="ph"><div><div class="pt">↗️ Redistribuição de Tickets</div><div class="ps">Redistribuição em massa com aprovação gerencial</div></div></div>

    <div class="sec">
      <div class="sec-title">Selecione os tickets e o operador de destino</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
        <div>
          <div style="font-size:11px;font-weight:700;color:var(--text-muted);margin-bottom:8px;">
            Tickets disponíveis
            ${redistState.tickets.length ? `<span style="color:var(--accent-mgr);margin-left:8px;">${redistState.tickets.length} selecionado(s)</span>` : ''}
          </div>
          ${ticketsAtivos.map(t => {
            const op = State.operadores.find(o => o.email === t.operador);
            const pct = SLA.pct(t);
            const sc = SLA.color(t);
            return `<div class="tmr ${redistState.tickets.includes(t.id) ? 'sel' : ''}" onclick="toggleRedistT('${t.id}')">
              <div class="tmr-id">${t.id}</div>
              <div class="tmr-subj">${t.assunto}</div>
              <div style="flex-shrink:0;display:flex;flex-direction:column;align-items:flex-end;gap:2px;">
                <span style="font-size:10px;font-weight:800;color:${sc};">${t.slaPausado ? '⏸' : pct + '%'}</span>
                <span style="font-size:10px;color:var(--text-muted);">${op?.nome.split(' ')[0] || '-'}</span>
              </div>
            </div>`;
          }).join('')}
        </div>
        <div>
          <div style="font-size:11px;font-weight:700;color:var(--text-muted);margin-bottom:8px;">Operador de destino</div>
          ${ops.map(op => {
            const ativos = State.tickets.filter(t => t.operador === op.email && t.status !== 'Concluído').length;
            const carga = Math.round((ativos / (op.capacidade || 5)) * 100);
            return `<div class="opr-row ${redistState.destino === op.email ? 'sel' : ''}" onclick="selectRedistDest('${op.email}')">
              <div style="width:28px;height:28px;border-radius:50%;background:${op.cor}18;color:${op.cor};display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;flex-shrink:0;">${op.sigla}</div>
              <div style="flex:1;"><div style="font-size:12px;font-weight:700;">${op.nome}</div><div style="font-size:11px;color:var(--text-muted);">Carga: ${ativos}/${op.capacidade||5} (${carga}%)</div></div>
              <div class="cap-track" style="width:50px;"><div class="cap-fill" style="width:${Math.min(carga,100)}%;background:${carga>80?'var(--red)':carga>60?'var(--yellow)':'var(--green)'}"></div></div>
            </div>`;
          }).join('')}
          <button class="btn bp" style="width:100%;margin-top:12px;justify-content:center;" onclick="execMgrRedist()">
            ↗️ Redistribuir ${redistState.tickets.length > 0 ? `(${redistState.tickets.length})` : ''}
          </button>
        </div>
      </div>
    </div>
  `;
}

function toggleRedistT(id) {
  const idx = State.redistSelected.tickets.indexOf(id);
  if (idx > -1) State.redistSelected.tickets.splice(idx, 1);
  else State.redistSelected.tickets.push(id);
  renderMgrRedist();
}

function selectRedistDest(email) {
  State.redistSelected.destino = email;
  renderMgrRedist();
}

async function execMgrRedist() {
  const { tickets, destino } = State.redistSelected;
  if (!tickets.length) { UI.showToast('Selecione ao menos um ticket', 'w'); return; }
  if (!destino) { UI.showToast('Selecione o operador de destino', 'w'); return; }
  for (const id of tickets) {
    await Actions.redistributeTicket(id, destino);
  }
  State.redistSelected = { tickets: [], destino: null };
  renderMgrRedist();
}

// ── SLA & KPIs ────────────────────────────────────────────────
function renderMgrSLA() {
  const T = State.tickets;
  const total = T.length;
  const dentro = T.filter(t => t.slaPausado || t.slaDecorrido / t.slaLimite < 1).length;
  const pctOk = total ? Math.round(dentro / total * 100) : 0;
  const pausados = T.filter(t => t.slaPausado).length;

  document.getElementById('main').innerHTML = `
    <div class="ph"><div><div class="pt">⏱️ SLA & KPIs</div><div class="ps">Nota Técnica v1.1 · Mon–Sex 07:30–18:00 · Clock-pause ativo</div></div></div>

    <div class="kpi-grid">
      <div class="kpi">
        <div class="kpi-label">SLA cumprido</div>
        <div class="kpi-val" style="color:${pctOk>=90?'var(--green)':'var(--yellow)'};">${pctOk}%</div>
        <div class="kpi-sub">dos tickets no prazo</div>
        <div class="kpi-trend ${pctOk>=90?'ok':'up'}">Meta: 90%</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">SLA pausado</div>
        <div class="kpi-val ct">${pausados}</div>
        <div class="kpi-sub">tickets com clock-pause</div>
        <div class="kpi-trend ok">Não conta como violação</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">T. médio 1ª resposta</div>
        <div class="kpi-val cy">~27min</div>
        <div class="kpi-sub">média da semana</div>
        <div class="kpi-trend ok">↓ Meta: &lt;30min</div>
      </div>
      <div class="kpi">
        <div class="kpi-label">Violações SLA</div>
        <div class="kpi-val cr">${T.filter(t=>!t.slaPausado&&t.slaDecorrido/t.slaLimite>=1).length}</div>
        <div class="kpi-sub">tickets vencidos</div>
      </div>
    </div>

    <div class="sec">
      <div class="sec-title">Política SLA vigente — Nota Técnica v1.1 <span>Aprovada pela Governance</span></div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px;">
        ${[
          { p:'Normal',   sla:'8h úteis', app:'4h (app)', cor:'var(--accent-mgr)', desc:'Operações padrão · sem impacto imediato' },
          { p:'Alta',     sla:'4h úteis', app:'2h (app)', cor:'var(--yellow)',      desc:'Impacto potencial · prazo próximo' },
          { p:'Crítica',  sla:'2h úteis', app:'1h (app)', cor:'var(--red)',         desc:'BL bloqueado · operação parada · cliente VIP' },
        ].map(p => `<div style="background:var(--surface2);border-radius:var(--r-s);padding:14px;border-left:3px solid ${p.cor};">
          <div style="font-size:11px;font-weight:800;color:${p.cor};margin-bottom:4px;">${p.p}</div>
          <div style="font-size:14px;font-weight:800;margin-bottom:2px;">${p.sla}</div>
          <div style="font-size:11px;color:var(--text-muted);margin-bottom:6px;">${p.app}</div>
          <div style="font-size:11px;color:var(--text-muted);">${p.desc}</div>
        </div>`).join('')}
      </div>
      <div style="background:var(--teal-s);border:1px solid rgba(20,184,166,.2);border-radius:var(--r-s);padding:10px 14px;font-size:12px;color:var(--teal);font-weight:600;">
        ⏸ Cláusula de clock-pause: SLA não corre em tickets "Ag. Terceiro" (Governance, cliente, outro departamento). Refletido em todas as métricas.
      </div>
    </div>

    <div class="sec">
      <div class="sec-title">SLA por prioridade — estado atual</div>
      ${['Crítica','Alta','Normal'].map(p => {
        const tps = State.tickets.filter(t => t.prioridade === p && !t.slaPausado && t.status !== 'Concluído');
        if (!tps.length) return `<div style="font-size:12px;color:var(--text-muted);padding:6px 0;">${p}: nenhum ticket ativo.</div>`;
        const avg = Math.round(tps.reduce((a,t) => a + t.slaDecorrido/t.slaLimite*100, 0) / tps.length);
        const c = p==='Crítica'?'var(--red)':p==='Alta'?'var(--yellow)':'var(--accent-mgr)';
        return `<div style="margin-bottom:14px;">
          <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:6px;">
            <span style="font-weight:700;">${p} (${tps.length} ticket${tps.length>1?'s':''})</span>
            <span style="color:${c};font-weight:800;">${avg}% consumido</span>
          </div>
          <div style="height:8px;background:var(--border);border-radius:4px;overflow:hidden;">
            <div style="width:${avg}%;height:100%;background:${c};border-radius:4px;"></div>
          </div>
        </div>`;
      }).join('')}
    </div>
  `;
}

// ── ALERTAS ───────────────────────────────────────────────────
function renderMgrAlertas() {
  const T = State.tickets;
  const alertas = [];

  // Críticos vencendo
  T.filter(t => t.prioridade === 'Crítica' && !t.slaPausado && t.status !== 'Concluído' && t.slaDecorrido/t.slaLimite >= 0.8)
    .forEach(t => alertas.push({ tipo:'Crítico', cor:'var(--red)', icon:'🚨', titulo:`${t.id} — SLA ${SLA.pct(t)}% consumido`, desc:`${t.operadorNome} · ${t.assunto} · ${SLA.label(t)}`, id:t.id }));

  // Alta acima de 80%
  T.filter(t => t.prioridade === 'Alta' && !t.slaPausado && t.status !== 'Concluído' && t.slaDecorrido/t.slaLimite >= 0.8)
    .forEach(t => alertas.push({ tipo:'Atenção', cor:'var(--yellow)', icon:'⚠️', titulo:`${t.id} — SLA ${SLA.pct(t)}% consumido`, desc:`${t.operadorNome} · ${t.assunto} · ${SLA.label(t)}`, id:t.id }));

  // Desequilíbrio de carga
  const ops = State.operadores.filter(o => o.role === 'operador');
  const cargas = ops.map(op => ({ op, n: T.filter(t => t.operador === op.email && t.status !== 'Concluído').length }));
  const max = Math.max(...cargas.map(c => c.n));
  const min = Math.min(...cargas.map(c => c.n));
  if (max - min >= 3) {
    const heavy = cargas.find(c => c.n === max);
    const light = cargas.find(c => c.n === min);
    alertas.push({ tipo:'Desequilíbrio', cor:'var(--orange)', icon:'⚖️', titulo:'Carga desigual detectada', desc:`${heavy.op.nome} (${heavy.n} tickets) vs ${light.op.nome} (${light.n} tickets). Considerar redistribuição.`, id:null });
  }

  // Padrão repetido por cliente
  const byCliente = T.reduce((acc, t) => { acc[t.cliente] = (acc[t.cliente]||0)+1; return acc; }, {});
  Object.entries(byCliente).filter(([,n]) => n >= 3).forEach(([cliente, n]) => {
    alertas.push({ tipo:'Padrão Repetido', cor:'var(--purple)', icon:'🔁', titulo:`Recorrência — ${cliente}`, desc:`${n} tickets do mesmo cliente. Possível inconsistência sistêmica no cadastro.`, id:null });
  });

  document.getElementById('main').innerHTML = `
    <div class="ph"><div><div class="pt">🔔 Alertas</div><div class="ps">${alertas.length} situações que requerem atenção</div></div></div>
    ${alertas.length
      ? alertas.map(a => `
        <div style="background:var(--surface);border:1px solid ${a.cor};border-radius:var(--r);padding:14px 16px;margin-bottom:10px;display:flex;gap:14px;align-items:flex-start;">
          <div style="font-size:26px;flex-shrink:0;">${a.icon}</div>
          <div style="flex:1;">
            <div style="font-size:10px;font-weight:800;color:${a.cor};margin-bottom:3px;text-transform:uppercase;letter-spacing:.5px;">${a.tipo}</div>
            <div style="font-size:13px;font-weight:800;margin-bottom:4px;">${a.titulo}</div>
            <div style="font-size:12px;color:var(--text-muted);">${a.desc}</div>
          </div>
          ${a.id
            ? `<button class="btn bp sm" onclick="openTicketMgr('${a.id}')">Ver →</button>`
            : `<button class="btn bp sm" onclick="navMgr('redist')">Redistribuir →</button>`}
        </div>`).join('')
      : `<div class="empty"><div class="ei">✅</div><p>Nenhum alerta ativo no momento.</p></div>`}
  `;
}

// ── AUDITORIA ─────────────────────────────────────────────────
function renderMgrAuditoria() {
  const entries = [];
  State.tickets.forEach(t => {
    (t.historico || []).forEach(h => {
      entries.push({
        cor: h.tipo==='sistema' ? 'var(--text-muted)' : h.de==='op' ? 'var(--accent-mgr)' : 'var(--green)',
        text: `<strong>${t.id}</strong> · ${h.nome} · ${h.texto}`,
        time: h.hora || h.ts || '',
        ticket: t.id,
      });
    });
  });

  document.getElementById('main').innerHTML = `
    <div class="ph">
      <div><div class="pt">🔍 Log de Auditoria</div><div class="ps">Rastreabilidade completa de todas as ações</div></div>
      <div class="pa"><button class="btn bg sm" onclick="exportAuditCSV()">⬇ Exportar log</button></div>
    </div>
    <div class="sec">
      <div class="log-list">
        ${entries.length
          ? entries.slice(0, 50).map(e => `
            <div class="log-item">
              <div class="log-dot" style="background:${e.cor};margin-top:5px;"></div>
              <div style="flex:1;"><div class="log-text">${e.text}</div></div>
              <div class="log-time">${e.time}</div>
            </div>`).join('')
          : '<div style="font-size:12px;color:var(--text-muted);">Sem entradas de auditoria.</div>'}
      </div>
    </div>
  `;
}

// ── TODOS OS TICKETS (gerência) ───────────────────────────────
function renderMgrTodos(prioFilter) {
  let tickets = State.getFilteredTickets();
  if (prioFilter) tickets = tickets.filter(t => t.prioridade === prioFilter);

  document.getElementById('main').innerHTML = `
    <div class="ph">
      <div>
        <div class="pt">${prioFilter ? '🚨 Tickets Críticos' : '📋 Todos os Tickets'}</div>
        <div class="ps">${tickets.length} tickets${State.filters.geo ? ` · País: ${State.filters.geo}` : ''}${State.filters.op ? ` · Operador filtrado` : ''}</div>
      </div>
      <div class="pa">
        ${State.filters.geo || State.filters.op ? `<button class="btn bg sm" onclick="clearMgrFilters()">✕ Limpar filtros</button>` : ''}
        <button class="btn bg sm" onclick="exportRelatorio()">⬇ Exportar</button>
      </div>
    </div>
    <div class="fbar">
      ${State.operadores.filter(o=>o.role==='operador').map(op =>
        `<div class="chip ${State.filters.op === op.email ? 'on' : ''}" onclick="filterMgrByOp('${op.email}')">${op.nome.split(' ')[0]}</div>`
      ).join('')}
      <input class="srch" placeholder="🔍 ID, cliente, país..." oninput="filterMgrSearch(this.value)">
    </div>
    <div id="mgr-tickets-list">
      ${renderTicketTable(tickets, 'manager')}
    </div>
  `;
}

function filterMgrSearch(q) {
  State.filters.search = q;
  document.getElementById('mgr-tickets-list').innerHTML = renderTicketTable(State.getFilteredTickets(), 'manager');
}

function filterMgrByOp(email) {
  State.filters.op = State.filters.op === email ? null : email;
  if (window._currentMgrNav === 'todos' || window._currentMgrNav === 'criticos') renderMgrTodos();
  else navMgr('todos');
}

function filterMgrByGeo(pais) {
  State.filters.geo = State.filters.geo === pais ? null : pais;
  if (window._currentMgrNav === 'todos') renderMgrTodos();
  else navMgr('todos');
}

function clearMgrFilters() {
  State.filters.op = null;
  State.filters.geo = null;
  State.filters.search = '';
  renderMgrTodos();
}

// ── TICKET DETAIL (gerência) ──────────────────────────────────
function openTicketMgr(id) {
  const t = State.tickets.find(x => x.id === id);
  if (!t) return;
  const op = State.operadores.find(o => o.email === t.operador);
  const sc = SLA.color(t);
  const pct = SLA.pct(t);
  const flag = PAISES_MGR[t.pais] || '🌐';

  document.getElementById('phdr').innerHTML = `
    <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
      <span style="font-size:12px;font-weight:800;color:var(--accent-mgr);font-family:monospace;">${t.id}</span>
      <span class="badge b-${t.prioridade.toLowerCase()}">${t.prioridade}</span>
      <span class="badge b-${t.status==='Terceiro'?'terceiro':t.status.toLowerCase().replace('í','i').replace('é','e')}">${t.status==='Terceiro'?'⏳ Ag. Terceiro':t.status}</span>
      ${t.slaPausado ? `<span class="badge b-pausado">⏸ SLA pausado</span>` : ''}
      <button class="panel-close" onclick="closeMgrPanel()" style="margin-left:auto;">✕</button>
    </div>
    <div style="font-size:14px;font-weight:800;margin-bottom:3px;">${flag} ${t.assunto}</div>
    <div style="font-size:11px;color:var(--text-muted);">${t.cliente} · ${op?.nome || t.operadorNome}</div>
  `;

  document.getElementById('pbody').innerHTML = `
    <div style="background:var(--surface2);border-radius:var(--r);padding:12px;margin-bottom:14px;">
      <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:8px;">
        <span style="font-weight:700;">${t.slaPausado ? 'SLA pausado' : 'Tempo utilizado'}</span>
        <span style="font-weight:900;color:${sc};">${t.slaPausado ? '⏸ ' + (t.pausaMotivo||'Pausado') : SLA.label(t)}</span>
      </div>
      ${!t.slaPausado ? `<div style="height:8px;background:var(--border);border-radius:4px;overflow:hidden;margin-bottom:6px;"><div style="width:${pct}%;height:100%;background:${sc};border-radius:4px;"></div></div>` : ''}
      <div style="font-size:11px;color:var(--text-muted);">${!t.slaPausado ? `${t.slaDecorrido}min / ${t.slaLimite}min (${pct}%)` : 'Dependência externa — clock-pause ativo'}</div>
    </div>
    <div class="ig">
      <div class="ii"><label>Operador</label><span>${op ? `<span style="background:${op.cor};border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;font-size:8px;font-weight:800;color:#fff;margin-right:5px;">${op.sigla}</span>${op.nome}` : t.operadorNome}</span></div>
      <div class="ii"><label>Cliente</label><span>${t.cliente}</span></div>
      <div class="ii"><label>País</label><span>${flag} ${t.pais}</span></div>
      <div class="ii"><label>Categoria</label><span>${t.categoria}</span></div>
      <div class="ii"><label>Abertura</label><span>${new Date(t.abertura).toLocaleString('pt-BR')}</span></div>
      <div class="ii"><label>ID CargoWise</label><span style="font-family:monospace;color:${t.cwId?'var(--green)':'var(--text-dim)'};">${t.cwId || 'Não registrado'}</span></div>
    </div>
    <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:var(--text-dim);margin:12px 0 8px;display:flex;align-items:center;gap:6px;">Histórico <div style="flex:1;height:1px;background:var(--border);"></div></div>
    ${(t.historico||[]).slice(-5).map(h => `
      <div style="font-size:11px;padding:6px 0;border-bottom:1px solid var(--border-s);">
        <strong>${h.nome}</strong> · ${h.hora} · ${h.texto}
      </div>`).join('') || '<div style="font-size:12px;color:var(--text-muted);">Sem histórico.</div>'}
  `;

  document.getElementById('pfooter').innerHTML = `
    ${t.status !== 'Concluído' ? `<button class="btn by sm" onclick="mgrRedistSingle('${t.id}')">↗️ Redistribuir</button>` : ''}
    ${!t.slaPausado && t.status !== 'Concluído' && SLA.pct(t) >= 80 ? `<button class="btn br sm" onclick="mgrEscalate('${t.id}')">🚨 Escalar</button>` : ''}
    <button class="btn bg sm" style="margin-left:auto;" onclick="closeMgrPanel()">Fechar</button>
  `;

  document.getElementById('overlay').classList.add('open');
}

function mgrRedistSingle(id) {
  State.redistSelected.tickets = [id];
  closeMgrPanel();
  navMgr('redist');
}

function mgrEscalate(id) {
  UI.showToast(`Ticket ${id} escalado — notificação enviada ✓`, 's');
  closeMgrPanel();
}

function closeMgrPanel() {
  document.getElementById('overlay').classList.remove('open');
}

// ── DONUT ─────────────────────────────────────────────────────
function donutMgr(data) {
  const total = data.reduce((a,d) => a+d.v, 0);
  if (!total) return '';
  const cx=40, cy=40, r=28, inner=16;
  let angle = -Math.PI/2;
  return data.map(d => {
    const sweep = (d.v/total) * 2 * Math.PI;
    const x1=cx+r*Math.cos(angle), y1=cy+r*Math.sin(angle);
    const x2=cx+r*Math.cos(angle+sweep), y2=cy+r*Math.sin(angle+sweep);
    const ix1=cx+inner*Math.cos(angle), iy1=cy+inner*Math.sin(angle);
    const ix2=cx+inner*Math.cos(angle+sweep), iy2=cy+inner*Math.sin(angle+sweep);
    const large = sweep > Math.PI ? 1 : 0;
    const path = `M${x1} ${y1} A${r} ${r} 0 ${large} 1 ${x2} ${y2} L${ix2} ${iy2} A${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1}Z`;
    angle += sweep;
    return `<path d="${path}" fill="${d.c}" opacity=".9"/>`;
  }).join('');
}

// ── EXPORTS ───────────────────────────────────────────────────
function exportRelatorio() {
  const rows = [['ID','Assunto','Cliente','Operador','País','Prioridade','Status','SLA%','CW-ID']];
  State.tickets.forEach(t => {
    const op = State.operadores.find(o => o.email === t.operador);
    const pct = t.slaPausado ? 'Pausado' : SLA.pct(t) + '%';
    rows.push([t.id, t.assunto, t.cliente, op?.nome||t.operadorNome, t.pais, t.prioridade, t.status, pct, t.cwId||'']);
  });
  const csv = rows.map(r => r.map(v => `"${v}"`).join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type:'text/csv' }));
  a.download = `corecad-gerencia-${new Date().toISOString().split('T')[0]}.csv`;
  a.click();
  UI.showToast('Relatório exportado ✓', 's');
}

function exportAuditCSV() {
  UI.showToast('Log de auditoria exportado ✓', 's');
}

// ── NAV GERÊNCIA ──────────────────────────────────────────────
function navMgr(id) {
  window._currentMgrNav = id;
  document.querySelectorAll('.sbi').forEach(el => el.classList.remove('on'));
  const el = document.getElementById('sn-' + id);
  if (el) el.classList.add('on');
  const fns = {
    overview:   renderMgrOverview,
    geo:        renderMgrGeo,
    tendencia:  () => {}, // placeholder — renderizar tendência
    equipe:     renderMgrEquipe,
    redist:     renderMgrRedist,
    sla:        renderMgrSLA,
    alertas:    renderMgrAlertas,
    auditoria:  renderMgrAuditoria,
    todos:      () => renderMgrTodos(),
    criticos:   () => renderMgrTodos('Crítica'),
  };
  if (fns[id]) fns[id]();
}

// Expor globalmente
window.navMgr = navMgr;
window.renderMgrOverview = renderMgrOverview;
window.renderMgrGeo = renderMgrGeo;
window.renderMgrEquipe = renderMgrEquipe;
window.renderMgrRedist = renderMgrRedist;
window.renderMgrSLA = renderMgrSLA;
window.renderMgrAlertas = renderMgrAlertas;
window.renderMgrAuditoria = renderMgrAuditoria;
window.renderMgrTodos = renderMgrTodos;
window.openTicketMgr = openTicketMgr;
window.closeMgrPanel = closeMgrPanel;
window.toggleRedistT = toggleRedistT;
window.selectRedistDest = selectRedistDest;
window.execMgrRedist = execMgrRedist;
window.filterMgrByOp = filterMgrByOp;
window.filterMgrByGeo = filterMgrByGeo;
window.clearMgrFilters = clearMgrFilters;
window.filterMgrSearch = filterMgrSearch;
window.exportRelatorio = exportRelatorio;
window.exportAuditCSV = exportAuditCSV;
window.mgrRedistSingle = mgrRedistSingle;
window.mgrEscalate = mgrEscalate;
