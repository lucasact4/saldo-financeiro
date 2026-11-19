(function () {
  'use strict';
  const M = window.FinanceModel, $ = selector => document.querySelector(selector);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = value => value === null || value === undefined || !Number.isFinite(Number(value)) ? 'Não informado' : Number(value).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  const compactMoney = value => Number(value).toLocaleString('pt-BR', { notation: Math.abs(value) >= 10000 ? 'compact' : 'standard', maximumFractionDigits: 0 });
  const displayDate = value => { if (!(value instanceof Date)) return typeof value === 'string' && /^\d{4}-\d\d-\d\d$/.test(value) ? new Date(value + 'T12:00:00').toLocaleDateString('pt-BR') : 'Não informada'; return Number.isNaN(value.getTime()) ? 'Não informada' : value.toLocaleDateString('pt-BR'); };
  // Excel dates have no timezone. SheetJS exposes them at UTC midnight; preserve that calendar day.
  const displayExcelDate = value => value instanceof Date && !Number.isNaN(value.getTime()) ? value.toLocaleDateString('pt-BR', { timeZone: 'UTC' }) : displayDate(value);
  const paths = {
    overview: '<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',
    cycles: '<path d="M4 7h16M4 12h12M4 17h9"/><path d="m17 15 3 3-3 3"/>',
    goals: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/>',
    sheets: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M4 9h16M4 15h16M10 9v12"/>',
    upload: '<path d="M12 16V3m-4 4 4-4 4 4M4 15v5a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-5"/>',
    download: '<path d="M12 3v13m-4-4 4 4 4-4M4 15v5a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-5"/>',
    restore: '<path d="M3 10a9 9 0 1 1 2 8M3 4v6h6M12 7v5l3 2"/>',
    shield: '<path d="M12 3 4 6v6c0 4 4 7 8 9 4-2 8-5 8-9V6z"/><path d="m8 12 3 3 5-6"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7h.01"/>',
    file: '<path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6M8 13h8M8 17h5"/>',
    chart: '<path d="M4 3v17h17M8 15V9m5 6V5m5 10v-4"/>',
    check: '<path d="m5 12 4 4L19 6"/>',
    close: '<path d="m6 6 12 12M6 18 18 6"/>',
    wallet: '<path d="M4 7V5a2 2 0 0 1 2-2h12v4M4 7h15a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8z"/><path d="M21 12h-5v5h5"/>',
    trend: '<path d="m3 17 6-6 4 4 8-10M15 5h6v6"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    eye: '<path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12"/><circle cx="12" cy="12" r="3"/>',
    calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M7 3v4m10-4v4M3 11h18M7 15h3m4 0h3"/>',
    card: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 10h18M7 15h3"/>',
    home: '<path d="m3 10 9-7 9 7M5 9v12h14V9M10 21v-7h4v7"/>',
    monitor: '<rect x="3" y="4" width="18" height="13" rx="2"/><path d="M8 21h8m-4-4v4"/>',
    heart: '<path d="M20.8 4.6a5.5 5.5 0 0 0-7.8 0L12 5.7l-1.1-1.1a5.5 5.5 0 0 0-7.8 7.8L12 21l8.8-8.6a5.5 5.5 0 0 0 0-7.8z"/>',
    bike: '<circle cx="5" cy="17" r="4"/><circle cx="19" cy="17" r="4"/><path d="m5 17 5-9 6 9M8 8h5m2-5h3l2 5M10 8l5 9"/>',
    search: '<circle cx="10.5" cy="10.5" r="6.5"/><path d="m16 16 5 5"/>',
    spark: '<path d="m12 3 2.5 6.5L21 12l-6.5 2.5L12 21l-2.5-6.5L3 12l6.5-2.5z"/>'
  };
  const icon = (name, cls = '') => `<svg class="${cls}" aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.65" stroke-linecap="round" stroke-linejoin="round">${paths[name] || paths.goals}</svg>`;
  function icons(scope = document) { scope.querySelectorAll('[data-icon]').forEach(el => { el.innerHTML = icon(el.dataset.icon); }); }
  const categoryIcon = category => ({ 'Moto': 'bike', 'TI / Setup': 'monitor', 'Moradia': 'home', 'Saúde': 'heart', 'Reserva': 'wallet' }[category] || 'goals');
  const STORE_KEY = 'saldo-local-v1', views = ['overview', 'cycles', 'goals', 'sheets'];
  let book = null, year = 2026, view = 'overview', selectedCycle = 0, demo = false, charts = [], goalTab = 'active', goalFilter = '', goalCategory = '', sheetName = '', formulaMode = false, sheetSearch = '', editingGoal = null, toastTimer = 0;
  let db = { version: 1, books: {} }, local = { closed: {}, goals: {}, customGoals: [], budgets: {} };
  try { const raw = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); if (raw?.version === 1 && raw.books && typeof raw.books === 'object' && !Array.isArray(raw.books)) db = raw; } catch (_) { /* Storage may be unavailable in private/file mode. */ }
  function toast(message, error = false) { clearTimeout(toastTimer); const t = $('#toast'); t.textContent = message; t.className = 'toast' + (error ? ' error' : ''); t.hidden = false; toastTimer = setTimeout(() => { t.hidden = true; }, error ? 9000 : 4500); }
  function save() {
    if (!book || demo) return;
    try { db.books[book.key] = M.validState(local); localStorage.setItem(STORE_KEY, JSON.stringify(db)); }
    catch (_) { toast('Não foi possível salvar neste navegador. Exporte um backup das metas antes de fechar.', true); }
  }
  const current = () => book.years.find(y => y.year === year);
  const closedCount = () => local.closed[year] ?? 0;
  const isClosed = p => p.index < closedCount();
  const statusTag = p => isClosed(p) ? '<span class="tag success">Fechado</span>' : '<span class="tag forecast">Previsão</span>';
  const allGoals = () => [...book.goals, ...local.customGoals];
  const activeGoals = () => allGoals().filter(g => !local.goals[g.id]?.hidden);
  const goalProgress = g => local.goals[g.id] || {};
  const gs = g => M.goalStats(g, goalProgress(g));
  function destroyCharts() { charts.forEach(c => c.destroy()); charts = []; }
  function showInfo(title, content, actions = null) {
    $('#info-title').textContent = title; $('#info-content').innerHTML = content;
    $('#info-footer').innerHTML = actions || '<button class="button button-primary" data-close-dialog="info-dialog">Entendi</button>';
    $('#info-dialog').showModal();
  }
  function privacy() { showInfo('Sua planilha fica com você', '<p>O arquivo é aberto diretamente neste navegador. Não há conta, servidor de dados, rastreamento nem conexão com seu banco ou OneDrive.</p><ul><li>A planilha fica apenas na memória enquanto a página está aberta.</li><li>Valores guardados, prazos, orçamentos anotados e preferências ficam no armazenamento local deste navegador.</li><li>O backup contém suas anotações e metas locais. Guarde-o como um arquivo pessoal.</li><li>As bibliotecas já estão incluídas no projeto. A página não usa CDN.</li></ul><p>Ao trocar de navegador, computador, endereço do site ou limpar os dados de navegação, restaure o backup das metas e abra novamente sua planilha.</p>'); }
  async function loadFile(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) { toast('Escolha um arquivo .xlsx, como a sua planilha original.', true); return; }
    if (file.size > 20 * 1024 * 1024) { toast('O arquivo excede 20 MB. Use uma cópia menor da planilha.', true); return; }
    if (!window.XLSX || !M) { toast('Não foi possível carregar as bibliotecas locais. Extraia todo o ZIP antes de abrir o index.html.', true); return; }
    $('#loading').hidden = false;
    try {
      await new Promise(resolve => setTimeout(resolve, 30));
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: true, cellFormula: true, sheetRows: 500 });
      const parsed = M.parse(workbook, file.name);
      activate(parsed, false); toast('Planilha aberta. O arquivo original não foi alterado.');
    } catch (error) { toast(error.message?.includes('password') ? 'Este arquivo está protegido por senha. Abra uma cópia sem senha para leitura.' : error.message || 'Não foi possível ler o arquivo. Baixe outra cópia no OneDrive e tente novamente.', true); }
    finally { $('#loading').hidden = true; $('#file-input').value = ''; }
  }
  function activate(parsed, isDemo) {
    destroyCharts(); book = parsed; demo = isDemo;
    local = M.validState(isDemo ? {} : db.books[book.key] || {});
    if (book.years.some(y => y.year === 2026) && local.closed[2026] === undefined) local.closed[2026] = 10;
    if (isDemo) {
      book.goals.forEach((g, i) => { local.goals[g.id] = { saved: [2400, 650, 1200][i] ?? 0, monthly: [350, 200, 400][i] ?? 100, priority: g.priority, deadline: '2027-09-30', hidden: false }; });
    }
    year = book.years.some(y => y.year === new Date().getFullYear()) ? new Date().getFullYear() : book.years.at(-1).year;
    selectedCycle = Math.min(local.closed[year] || 0, 11); sheetName = current().name;
    $('#welcome').hidden = true; $('#loaded-app').hidden = false; $('#demo-tag').hidden = !isDemo;
    $('#open-file-label').textContent = 'Trocar planilha';
    $('#source-title').textContent = isDemo ? 'Arquivo de exemplo' : 'Arquivo aberto';
    $('#source-caption').textContent = `${book.sheetNames.length} abas · leitura local`;
    $('#filename-label').textContent = book.filename;
    $('#goal-nav-count').textContent = activeGoals().length; $('#goal-nav-count').hidden = false;
    save(); render();
  }
  function render() {
    if (!book) return;
    destroyCharts();
    document.querySelectorAll('.nav-item').forEach(el => { el.classList.toggle('active', el.dataset.view === view); el.setAttribute('aria-current', el.dataset.view === view ? 'page' : 'false'); });
    const titles = { overview: ['Visão geral', 'O que entrou, o que saiu e o que vem pela frente.'], cycles: ['Ciclos e gastos', 'Um ciclo por vez. Entenda os detalhes sem se perder nas colunas.'], goals: ['Metas e objetivos', 'Dê um destino para a sua próxima conquista.'], sheets: ['Suas planilhas', 'Consulte o conteúdo do arquivo, com suas abas originais.'] };
    $('#page-title').textContent = titles[view][0]; $('#page-description').textContent = titles[view][1];
    $('#year-switch').innerHTML = book.years.map(y => `<button class="${y.year === year ? 'active' : ''}" data-year="${y.year}" aria-pressed="${y.year === year}">${y.year}</button>`).join('');
    $('#year-switch').hidden = ['goals', 'sheets'].includes(view);
    $('#context-bar').hidden = ['goals', 'sheets'].includes(view);
    $('#source-year-label').textContent = current().name;
    $('#closed-select').innerHTML = '<option value="0">Nenhum ciclo</option>' + current().periods.map(p => `<option value="${p.index + 1}">${esc(p.label)} ${p.year}</option>`).join('');
    $('#closed-select').value = closedCount();
    for (const name of views) $('#view-' + name).hidden = view !== name;
    if (view === 'overview') renderOverview();
    if (view === 'cycles') renderCycles();
    if (view === 'goals') renderGoals();
    if (view === 'sheets') renderSheets();
    $('#goal-nav-count').textContent = activeGoals().length; icons($('#main'));
  }
  function kpi(label, value, caption, symbol, featured = false, negative = false, empty = false) {
    return `<article class="kpi ${featured ? 'featured' : ''}"><div class="kpi-top"><span class="kpi-label">${esc(label)}</span><span class="kpi-icon">${icon(symbol)}</span></div><strong class="kpi-value ${negative ? 'negative' : ''} ${empty ? 'empty' : ''}">${esc(value)}</strong><span class="kpi-foot">${esc(caption)}</span></article>`;
  }
  function goalPreview(g) { const s = gs(g); return `<div class="goal-preview-row"><span class="goal-preview-icon">${icon(categoryIcon(g.category))}</span><div><strong>${esc(g.name)}</strong><small>${s.saved === null ? 'Valor guardado ainda não informado' : `${money(s.saved)} guardados · ${Math.round(s.progress)}%`}</small><div class="mini-progress"><i style="width:${s.progress ?? 0}%"></i></div></div><span class="target">${money(g.target)}</span></div>`; }
  function renderOverview() {
    const y = current(), closed = closedCount(), p = y.periods[selectedCycle], last = closed ? y.periods[closed - 1] : null, goals = activeGoals();
    const goalCost = M.sum(goals.map(g => g.target)), known = goals.filter(g => gs(g).saved !== null).length;
    const ranks = {Alta:0,Média:1,Baixa:2,'A definir':3};
    const previews = [...goals].sort((a, b) => (ranks[goalProgress(a).priority || a.priority] ?? 3) - (ranks[goalProgress(b).priority || b.priority] ?? 3)).slice(0, 3);
    const worst = y.periods.slice(0, closed).filter(i => i.net < 0).sort((a, b) => a.net - b.net)[0];
    const firstCard = y.periods[0].items.filter(i => i.group === 'Cartão').reduce((a, i) => a + i.amount, 0), lastCard = M.sum(y.periods[11].items.filter(i => i.group === 'Cartão').map(i => i.amount));
    const insight1 = worst ? { title: `${worst.label} teve sobra negativa`, text: `${money(worst.net)} no ciclo. O saldo acumulado é mostrado separadamente para preservar sua leitura.`, ico: 'chart' } : { title: closed ? 'Seus ciclos fechados em perspectiva' : 'Planejamento ainda em aberto', text: closed ? `${closed} ciclos definidos como fechados. Você pode ajustar esse marco no seletor acima.` : 'Defina o último ciclo fechado para separar histórico e previsão nos gráficos.', ico: 'calendar' };
    const insight2 = lastCard < firstCard ? { title: 'O cartão perde peso na previsão', text: `O valor passa de ${money(firstCard)} para ${money(lastCard)} entre o primeiro e o último ciclo do ano.`, ico: 'card' } : { title: 'Olhe além do saldo final', text: `A menor sobra do ano é ${money(Math.min(...y.periods.map(i => i.net)))}. Os ciclos ajudam a enxergar os meses mais apertados.`, ico: 'trend' };
    const insight3 = { title: known < goals.length ? 'Falta dar contexto às suas metas' : 'Seus objetivos estão acompanhados', text: known < goals.length ? `${goals.length - known} metas ainda sem valor guardado. Informe também o prazo para calcular um aporte de referência.` : 'Aportes de metas são planejamento. Não são descontados novamente dos saldos da planilha.', ico: 'goals' };
    $('#view-overview').innerHTML = `<div class="kpi-grid">
      ${kpi('Saldo nos ciclos fechados', last ? money(last.balance) : 'Sem fechamento', last ? `Até ${last.label} · ${closed} ciclos fechados` : 'Defina o último ciclo no seletor acima', 'wallet', true, false, !last)}
      ${kpi('Saldo previsto no fim do ano', money(y.periods[11].balance), `${12 - closed} ciclos em aberto · saldo inicial ${money(y.opening)}`, 'trend', false, y.periods[11].balance < 0)}
      ${kpi(`Sobra em ${p.label}`, money(p.net), `${isClosed(p) ? 'Fechado' : 'Previsão'} · receitas menos despesas`, 'cycles', false, p.net < 0)}
      ${kpi('Objetivos em foco', String(goals.length).padStart(2, '0'), `${money(goalCost)} em objetivos · ${known} com progresso`, 'goals')}
    </div><div class="chart-grid"><article class="panel"><div class="panel-heading"><div><h2>O caminho do seu saldo</h2><p>Saldo acumulado ao fim de cada ciclo de ${year}.</p></div><span class="tag">${year}</span></div><div class="chart-wrap tall"><canvas id="balance-chart" role="img" aria-label="Evolução do saldo acumulado; os valores estão disponíveis na aba Ciclos e gastos."></canvas></div><div class="chart-legend"><span><i class="legend-line"></i>Ciclos fechados</span><span><i class="legend-line forecast"></i>Previsão da planilha</span></div></article><article class="panel"><div class="panel-heading"><div><h2>Entradas e saídas</h2><p>Receitas e despesas registradas por ciclo.</p></div></div><div class="chart-wrap tall"><canvas id="flow-chart" role="img" aria-label="Comparação das receitas e despesas por ciclo, detalhadas em Ciclos e gastos."></canvas></div><div class="chart-legend"><span><i class="legend-square"></i>Receitas</span><span><i class="legend-square expenses"></i>Despesas</span></div></article></div>
    <div class="lower-grid"><article class="panel"><div class="panel-heading"><div><h2>O dinheiro neste ciclo</h2><p>Grupos do seu quadro mensal, sem detalhar a fatura duas vezes.</p></div><div class="cycle-selector"><select data-cycle-select aria-label="Ciclo dos indicadores">${y.periods.map(i => `<option value="${i.index}" ${i.index === selectedCycle ? 'selected' : ''}>${esc(i.label)}</option>`).join('')}</select></div></div>${breakdown(p)}<button class="text-button" data-open-cycle="${p.index}">Ver todos os lançamentos</button></article><article class="panel"><div class="panel-heading"><div><h2>Suas próximas conquistas</h2><p>Prioridades e objetivos que vieram da planilha.</p></div><button class="text-button" data-goto="goals">Ver metas</button></div>${previews.length ? previews.map(goalPreview).join('') : '<p class="small-note">Adicione uma meta para começar.</p>'}</article></div>
    <div class="insights">${[insight1, insight2, insight3].map(i => `<article class="insight"><span>${icon(i.ico)}</span><div><strong>${esc(i.title)}</strong><p>${esc(i.text)}</p></div></article>`).join('')}</div>`;
    makeCharts(y);
  }
  function breakdown(p) {
    const colors = { 'Cartão': '#315d75', 'Gastos fixos': '#479579', 'Gastos previstos': '#a6c4b6', 'Outros lançamentos': '#a2b3ca', 'Créditos e reembolsos': '#c09870' };
    const grouped = new Map(); for (const item of p.items) grouped.set(item.group, (grouped.get(item.group) || 0) + item.amount);
    const positive = M.sum([...grouped.values()].filter(n => n > 0));
    return [...grouped].sort((a, b) => b[1] - a[1]).map(([name, value]) => `<div class="breakdown-row"><span class="breakdown-label"><i class="category-dot" style="background:${colors[name]}"></i>${esc(name)}</span><strong class="breakdown-value">${money(value)}</strong><div class="breakdown-track"><i style="background:${colors[name]};width:${positive > 0 ? Math.min(100, Math.abs(value) / positive * 100) : 0}%"></i></div></div>`).join('') + `<div class="breakdown-total"><span>${isClosed(p) ? 'Despesa registrada' : 'Despesa prevista'}</span><strong>${money(p.expense)}</strong></div><p class="small-note">${esc(p.sheet)} · ${esc(p.expenseCell)}. Valores negativos são créditos ou ajustes que reduzem a despesa.</p>`;
  }
  function makeCharts(y) {
    if (!window.Chart) { toast('A biblioteca local de gráficos não foi carregada. Confira a pasta vendor.', true); return; }
    Chart.defaults.font.family = 'Inter, Segoe UI, Arial, sans-serif'; Chart.defaults.font.size = 12; Chart.defaults.color = '#728294';
    const reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    const shared = { responsive: true, maintainAspectRatio: false, animation: reduced ? false : { duration: 320 }, plugins: { legend: { display: false }, tooltip: { backgroundColor: '#172c3e', padding: 12, titleFont: { size: 12 }, bodyFont: { size: 12 }, displayColors: false, callbacks: { label: c => `${c.dataset.label}: ${money(c.parsed.y)}` } } }, interaction: { mode: 'index', intersect: false }, scales: { x: { grid: { display: false }, border: { display: false }, ticks: { color: '#8b97a6', font: { size: 10 }, maxRotation: 0, autoSkip: true, maxTicksLimit: 6 } }, y: { beginAtZero: true, border: { display: false }, grid: { color: '#edf1f5', drawTicks: false }, ticks: { padding: 10, maxTicksLimit: 5, callback: value => 'R$ ' + compactMoney(value) } } } };
    const n = closedCount(), labels = y.periods.map(p => p.label);
    charts.push(new Chart($('#balance-chart'), { type: 'line', data: { labels, datasets: [ { label: 'Saldo fechado', data: y.periods.map((p, i) => i < n ? p.balance : null), borderColor: '#238361', backgroundColor: '#2383610b', borderWidth: 2.5, pointRadius: n === 1 ? 4 : 2.5, pointHoverRadius: 5, pointBackgroundColor: '#fff', pointBorderWidth: 2, fill: n > 1, tension: .22 }, { label: 'Saldo previsto', data: y.periods.map((p, i) => i >= Math.max(0, n - 1) && n < 12 ? p.balance : null), borderColor: '#91a7bd', borderWidth: 2, pointRadius: 2.5, pointBackgroundColor: '#fff', pointHoverRadius: 5, borderDash: [5, 5], tension: .22 } ] }, options: shared }));
    charts.push(new Chart($('#flow-chart'), { type: 'bar', data: { labels, datasets: [ { label: 'Receitas', data: y.periods.map(p => p.income), backgroundColor: '#3b856a', borderRadius: 3, maxBarThickness: 11 }, { label: 'Despesas', data: y.periods.map(p => p.expense), backgroundColor: '#b8c8db', borderRadius: 3, maxBarThickness: 11 } ] }, options: shared }));
  }
  function renderCycles() {
    const y = current(), n = closedCount(), totalIncome = M.sum(y.periods.slice(0, n).map(p => p.income)), totalExpense = M.sum(y.periods.slice(0, n).map(p => p.expense));
    const schedule = book.schedules.find(s => s.source === y.name) || book.schedules.at(-1);
    const moto = book.workbook.Sheets.Moto;
    $('#view-cycles').innerHTML = `<div class="section-toolbar"><p>Os valores vêm dos totais da sua planilha. “Fechado” é o marco que você definiu acima. Clique em um ciclo para ver os lançamentos e anotar um orçamento.</p><span class="tag">${n} de 12 ciclos fechados</span></div>
    <div class="panel table-panel"><div class="table-top panel-heading"><div><h2>Seu ano, ciclo a ciclo</h2><p>${esc(y.name)} · saldo inicial ${money(y.opening)}</p></div></div><div class="table-scroll"><table class="data-table"><thead><tr><th>Ciclo</th><th>Situação</th><th class="num">Receitas</th><th class="num">Despesas</th><th class="num">Sobra do ciclo</th><th class="num">Saldo acumulado</th><th class="num">Real − orçamento</th><th><span class="subtle">Detalhes</span></th></tr></thead><tbody>${y.periods.map(p => {
      const budget = local.budgets[`${year}-${p.index + 1}`], delta = isClosed(p) && budget !== undefined ? p.expense - budget : null;
      return `<tr class="clickable" data-open-cycle="${p.index}"><td class="cycle-name">${esc(p.label)}</td><td>${statusTag(p)}</td><td class="num">${money(p.income)}</td><td class="num">${money(p.expense)}</td><td class="num ${p.net < 0 ? 'negative' : 'positive'}">${money(p.net)}</td><td class="num ${p.balance < 0 ? 'negative' : ''}">${money(p.balance)}</td><td class="num ${delta > 0 ? 'negative' : delta < 0 ? 'positive' : ''}">${delta === null ? '<span class="subtle">—</span>' : money(delta)}</td><td><button class="icon-button" data-open-cycle="${p.index}" aria-label="Ver lançamentos de ${esc(p.label)}">${icon('eye')}</button></td></tr>`;
    }).join('')}${n ? `<tr class="totals-row"><td colspan="2">Totais dos ${n} ciclos fechados</td><td class="num">${money(totalIncome)}</td><td class="num">${money(totalExpense)}</td><td class="num">${money(totalIncome - totalExpense)}</td><td class="num">${money(y.periods[n - 1].balance)}</td><td colspan="2"></td></tr>` : ''}</tbody></table></div></div>
    <p class="small-note">“Real − orçamento” só é calculado quando o ciclo está fechado e você informou um orçamento local. Não inventamos o orçamento dos meses anteriores.</p>
    <div class="split-panels"><article class="panel"><div class="panel-heading"><div><h2>Parcelas identificadas</h2><p>${schedule ? `${esc(schedule.source)} · referência do quadro: ${displayExcelDate(schedule.date)}` : 'Nenhum quadro de parcelas encontrado.'}</p></div><span class="tag">Quadro do arquivo</span></div>${schedule?.items.length ? schedule.items.map(i => `<div class="installment-row"><div><strong>${esc(i.name)}</strong><small>Parcela ${i.current} de ${i.count} · ${i.remaining} após a indicada</small></div><div>${money(i.monthly)}<small>${money(i.remainingValue)} a vencer</small></div></div>`).join('') + `<p class="small-note">Após pagar as parcelas indicadas: ${money(schedule.remaining)} restantes. Isso não é uma nova cobrança no painel. Os valores já integram sua organização na planilha.</p>` : '<p class="small-note">As previsões de cartão no quadro mensal continuam disponíveis na tabela acima.</p>'}</article>
    <article class="panel"><div class="panel-heading"><div><h2>${moto ? 'Seu planejamento da moto' : 'Como ler este painel'}</h2><p>${moto ? 'Resumo do quadro de pagamento da aba Moto.' : 'A planilha continua sendo a fonte dos dados.'}</p></div><span>${icon(moto ? 'bike' : 'info')}</span></div>${moto ? [['Total da moto', 'I46'], ['Pago, incluindo sua entrada', 'I52'], ['Restante a pagar', 'I49']].map(([label, a]) => `<div class="simple-metric"><span>${label}</span><strong>${money(M.number(moto[a]?.v))}</strong></div>`).join('') : ''}<p class="small-note">O leitor usa os resultados salvos no arquivo, sem recalcular suas fórmulas. Para atualizar o painel, atualize sua planilha no OneDrive, baixe o .xlsx e abra a nova cópia.</p>${y.sharedHistoricalInputs ? '<p class="small-note">Na planilha original, alguns meses compartilham o total atual de gastos fixos. Alterar esse total no Excel também pode alterar o histórico. Este leitor preserva os resultados do arquivo importado.</p>' : ''}${y.issues.length ? `<p class="sheet-warning">${y.issues.map(i => esc(i.message)).join('<br>')}</p>` : ''}</article></div>`;
  }
  function openCycle(index) {
    const p = current().periods[index]; if (!p) return;
    const key = `${year}-${index + 1}`, budget = local.budgets[key];
    $('#cycle-modal-source').textContent = `${p.sheet} · ${p.expenseCell}, ${p.netCell}, ${p.balanceCell}`;
    $('#cycle-modal-title').textContent = `${p.label} ${year}`;
    $('#cycle-modal-content').innerHTML = `<div class="cycle-stats"><div><span>Receitas</span><strong>${money(p.income)}</strong></div><div><span>Despesas</span><strong>${money(p.expense)}</strong></div><div><span>Sobra do ciclo</span><strong class="${p.net < 0 ? 'negative' : ''}">${money(p.net)}</strong></div></div>${statusTag(p)}<p class="small-note">${esc(p.fullLabel)}. Receitas = despesas + sobra registrada. Créditos e reembolsos aparecem com sinal negativo.</p><div class="cycle-items">${p.items.map(i => `<div class="cycle-item"><div>${esc(i.name)}<small>${esc(i.cell)} · ${esc(i.group)}</small></div><strong>${money(i.amount)}</strong></div>`).join('')}</div><form id="budget-form" class="budget-form" data-cycle="${index}"><label for="cycle-budget">Orçamento de gastos deste ciclo (opcional)</label><div><input class="input" id="cycle-budget" type="number" min="0" max="1000000000000" step="0.01" placeholder="Informe um valor em R$" value="${budget ?? ''}"><button class="button button-primary" type="submit">Salvar</button></div><p class="small-note">Fica salvo neste navegador. Deixe em branco para remover. Não modifica o .xlsx.</p><p class="budget-result">${budget === undefined ? 'Nenhum orçamento informado.' : isClosed(p) ? `${p.expense <= budget ? 'Dentro do orçamento' : 'Acima do orçamento'} · diferença ${money(p.expense - budget)}` : `Previsão versus orçamento: ${money(p.expense - budget)}. O realizado só é comparado ao fechar o ciclo.`}</p></form>`;
    $('#cycle-dialog').showModal();
  }
  function renderGoals() {
    const all = activeGoals(), known = all.filter(g => gs(g).saved !== null), saved = M.sum(known.map(g => gs(g).saved)), monthly = M.sum(all.map(g => gs(g).monthly));
    const filters = all.filter(g => (!goalFilter || M.norm(g.name).includes(M.norm(goalFilter))) && (!goalCategory || g.category === goalCategory));
    const categories = [...new Set(all.map(g => g.category))].sort();
    $('#view-goals').innerHTML = `<div class="section-toolbar"><div class="segmented"><button data-goal-tab="active" class="${goalTab === 'active' ? 'active' : ''}">Em foco (${all.length})</button><button data-goal-tab="library" class="${goalTab === 'library' ? 'active' : ''}">Listas da planilha</button></div><button class="button button-primary" data-add-goal>${icon('plus')}Nova meta</button></div>
    <div class="goal-summary"><div><span>Valor-alvo dos objetivos</span><strong>${money(M.sum(all.map(g => g.target)))}</strong></div><div><span>Guardado informado · ${known.length} de ${all.length} metas</span><strong>${known.length ? money(saved) : 'Não informado'}</strong></div><div><span>Aportes mensais planejados</span><strong>${all.some(g => gs(g).monthly !== null) ? money(monthly) : 'Não informado'}</strong></div></div>
    ${goalTab === 'active' ? `<div class="section-toolbar"><p>${book.goalSourceYear ? `Metas importadas de Financeiro ${book.goalSourceYear}, sem repetir os itens de Prioridade.` : 'Nenhuma meta encontrada nas abas Financeiro.'} Seu progresso fica neste navegador.</p><div class="filter-row"><input class="input" id="goal-search" type="search" placeholder="Buscar uma meta…" aria-label="Buscar meta" value="${esc(goalFilter)}"><select id="goal-category" aria-label="Categoria das metas"><option value="">Todas as categorias</option>${categories.map(c => `<option ${c === goalCategory ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></div></div><div id="goal-cards" class="goals-grid">${filters.map(goalCard).join('')}</div>${!filters.length ? '<div class="empty-state"><h2>Nenhuma meta por aqui</h2><p>Adicione um objetivo ou ajuste os filtros.</p><button class="button button-primary" data-add-goal>Adicionar meta</button></div>' : ''}<p class="small-note">${allGoals().some(g => goalProgress(g).hidden) ? '<button class="text-button" data-show-hidden>Mostrar metas ocultas</button> · ' : ''}As estimativas usam o aporte mensal informado, sem rendimentos. Valores guardados já fazem parte do seu patrimônio; não são descontados novamente do saldo.</p>` : `<p class="sheet-warning">Estas são listas de referência das outras abas. Confira itens já contemplados nas metas atuais antes de adicionar outra meta com o mesmo custo.</p><div class="library-grid">${book.library.map(plan => `<article class="plan-card"><span class="tag">${esc(plan.source)}</span><h3 style="margin-top:15px">${esc(plan.name)}</h3><p>${plan.items.length} itens encontrados no arquivo. Valores conforme a planilha.</p><div class="plan-total">${money(plan.target)}</div><div class="plan-items">${plan.items.map(i => `<div><span>${esc(i.name)}</span><strong>${money(i.target)}</strong></div>`).join('')}</div><button class="button" data-add-plan="${esc(plan.id)}">${icon('plus')}Acompanhar este plano</button></article>`).join('')}</div>`}`;
  }
  function goalCard(g) {
    const s = gs(g), progress = goalProgress(g), priority = progress.priority || g.priority;
    return `<article class="goal-card"><div class="goal-card-top"><span class="goal-preview-icon">${icon(categoryIcon(g.category))}</span><span class="tag ${priority === 'Alta' ? 'warn' : ''}">${priority === 'A definir' ? 'Prioridade a definir' : 'Prioridade ' + esc(priority.toLowerCase())}</span></div><h3>${esc(g.name)}</h3><span class="goal-category">${esc(g.category)} · ${g.imported ? 'Da planilha' : 'Meta local'}</span><div class="goal-values"><div><strong>${s.saved === null ? '—' : money(s.saved)}</strong></div><div class="muted">de ${money(g.target)}<br>${s.remaining === null ? 'guardado não informado' : `${money(s.remaining)} para chegar lá`}</div></div><div class="goal-progress ${s.progress === null ? 'unknown' : ''}"><i style="width:${s.progress ?? 0}%"></i></div><div class="goal-status"><span>${esc(s.status)}</span><span>${s.progress === null ? 'Sem dados' : Math.round(s.progress) + '%'}</span></div><div class="goal-data-row"><span>Aporte mensal</span><strong>${s.monthly === null ? 'Não informado' : money(s.monthly)}</strong></div><div class="goal-data-row"><span>Prazo desejado</span><strong>${s.deadline ? displayDate(s.deadline) : 'A definir'}</strong></div>${s.suggested !== null && s.remaining > 0 ? `<div class="goal-data-row"><span>Aporte para o prazo</span><strong>${money(s.suggested)}/mês</strong></div>` : ''}<button class="button" data-edit-goal="${esc(g.id)}">${s.saved === null ? 'Começar a acompanhar' : 'Atualizar progresso'}</button></article>`;
  }
  function openGoal(id = null, plan = null) {
    editingGoal = id ? allGoals().find(g => g.id === id) : null;
    if (!editingGoal && plan) editingGoal = { id: '', name: plan.name, target: plan.target, category: plan.category, priority: 'A definir', imported: false };
    const g = editingGoal, p = g?.id ? goalProgress(g) : {}, custom = !g?.imported;
    $('#goal-dialog-title').textContent = g?.id ? 'Acompanhar meta' : 'Uma nova conquista';
    const categories = [...new Set(['Moto', 'TI / Setup', 'Saúde', 'Pessoal', 'Trabalho', 'Moradia', 'Reserva', 'Outros', g?.category].filter(Boolean))];
    $('#goal-form-fields').innerHTML = `<div class="form-grid"><div class="form-field full"><label for="goal-name">Nome da meta</label><input class="input" id="goal-name" required maxlength="140" value="${esc(g?.name || '')}" ${custom ? '' : 'readonly'}></div><div class="form-field"><label for="goal-target">Valor-alvo (R$)</label><input class="input" id="goal-target" type="number" required min="0.01" max="1000000000000" step="0.01" value="${g?.target ?? ''}" ${custom ? '' : 'readonly'}></div><div class="form-field"><label for="goal-cat">Categoria</label><select id="goal-cat" ${custom ? '' : 'disabled'}>${categories.map(c => `<option ${c === (g?.category || 'Outros') ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></div><div class="form-field"><label for="goal-saved">Quanto já guardou (R$)</label><input class="input" id="goal-saved" type="number" min="0" max="1000000000000" step="0.01" value="${p.saved ?? ''}" placeholder="Não informado"><small>Informe 0 se ainda não reservou dinheiro.</small></div><div class="form-field"><label for="goal-monthly">Aporte mensal (R$)</label><input class="input" id="goal-monthly" type="number" min="0" max="1000000000000" step="0.01" value="${p.monthly ?? ''}" placeholder="Não informado"></div><div class="form-field"><label for="goal-deadline">Prazo desejado</label><input class="input" id="goal-deadline" type="date" min="2000-01-01" max="2100-12-31" value="${esc(p.deadline || '')}"></div><div class="form-field"><label for="goal-priority">Prioridade</label><select id="goal-priority">${['Alta', 'Média', 'Baixa', 'A definir'].map(v => `<option ${(p.priority || g?.priority || 'A definir') === v ? 'selected' : ''}>${v}</option>`).join('')}</select></div></div><div id="goal-live-preview" class="goal-preview-math"></div><p class="form-help">${g?.imported ? `Nome e valor-alvo vêm de ${esc(g.source)}. Para alterá-los, atualize a planilha e abra a nova cópia. Progresso, aporte e prazo ficam apenas no navegador.` : 'Esta meta será salva neste navegador e incluída no backup. A sua planilha não será alterada.'}</p>${g?.id ? '<button type="button" class="link-danger" id="hide-goal">Ocultar esta meta do painel</button>' : ''}`;
    $('#goal-form-error').hidden = true; previewGoal(); $('#goal-dialog').showModal();
  }
  function previewGoal() {
    const name = $('#goal-name')?.value || '', target = M.number($('#goal-target')?.value), saved = M.number($('#goal-saved')?.value), monthly = M.number($('#goal-monthly')?.value), deadline = $('#goal-deadline')?.value || '';
    const s = M.goalStats({ name, target }, { saved, monthly, deadline });
    $('#goal-live-preview').innerHTML = `<div><small>Falta guardar</small><strong>${s.remaining === null ? '—' : money(s.remaining)}</strong></div><div><small>${s.suggested !== null ? 'Aporte para o prazo' : 'Meses com esse aporte'}</small><strong>${s.suggested !== null ? money(s.suggested) + '/mês' : s.monthsToGoal !== null ? s.monthsToGoal + ' meses' : '—'}</strong></div>`;
  }
  function submitGoal(event) {
    event.preventDefault(); const form = $('#goal-form'); if (!form.reportValidity()) return;
    const target = M.number($('#goal-target').value), saved = M.number($('#goal-saved').value), monthly = M.number($('#goal-monthly').value), name = $('#goal-name').value.trim();
    if (!name || target === null || target <= 0 || (saved !== null && saved < 0) || (monthly !== null && monthly < 0)) { $('#goal-form-error').textContent = 'Confira o nome e os valores da meta.'; $('#goal-form-error').hidden = false; return; }
    const id = editingGoal?.id || 'custom-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
    const priority = $('#goal-priority').value;
    if (!editingGoal?.imported) {
      const g = { id, name, target, category: $('#goal-cat').value, priority, imported: false };
      const old = local.customGoals.findIndex(i => i.id === id); if (old >= 0) local.customGoals[old] = g; else local.customGoals.push(g);
    }
    local.goals[id] = { saved, monthly, deadline: $('#goal-deadline').value, priority, hidden: false };
    save(); $('#goal-dialog').close(); render(); toast(demo ? 'Meta atualizada na demonstração.' : 'Progresso salvo neste navegador.');
  }
  function renderSheets() {
    if (!book.sheetNames.includes(sheetName)) sheetName = current().name;
    $('#view-sheets').innerHTML = `<div class="sheet-viewer-controls"><label>Aba do arquivo<select id="sheet-select">${book.sheetNames.map(n => `<option ${n === sheetName ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select></label><div class="filter-row"><input class="input" id="sheet-search" type="search" placeholder="Buscar nesta aba…" aria-label="Buscar na aba" value="${esc(sheetSearch)}"><label style="display:flex;align-items:center;gap:7px;white-space:nowrap"><input id="formula-mode" type="checkbox" ${formulaMode ? 'checked' : ''}>Mostrar fórmulas</label></div></div><div id="sheet-content"></div><p class="small-note">Consulta somente de leitura. Até 500 linhas e 26 colunas por aba. Fórmulas são exibidas como texto e não são executadas pelo leitor.</p>`;
    renderSheetTable();
  }
  function renderSheetTable() {
    const sheet = book.workbook.Sheets[sheetName], range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    const endRow = Math.min(range.e.r, 499), endCol = Math.min(range.e.c, 25), rows = [];
    for (let r = range.s.r; r <= endRow; r++) {
      const cells = [];
      for (let c = range.s.c; c <= endCol; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })]; let v = '';
        if (cell) v = formulaMode && cell.f ? '=' + cell.f : cell.v instanceof Date ? displayExcelDate(cell.v) : typeof cell.v === 'number' && /R\$|\$/.test(cell.z || '') ? money(cell.v) : cell.w ?? cell.v ?? '';
        cells.push({ v: String(v), numeric: typeof cell?.v === 'number' && !(formulaMode && cell?.f) });
      }
      if (sheetSearch && !M.norm(cells.map(i => i.v).join(' ')).includes(M.norm(sheetSearch))) continue;
      if (!cells.some(c => c.v)) continue;
      const line = M.norm(cells.map(c => c.v).join(' ')), cls = line.includes('total dividas') || line.includes('sobra total') ? 'total-like' : cells.some(c => ['nº', 'nome', 'vacina'].includes(c.v.toLowerCase())) ? 'header-like' : '';
      rows.push(`<tr class="${cls}"><td>${r + 1}</td>${cells.map(c => `<td class="${c.numeric ? 'num' : ''}" title="${esc(c.v)}">${esc(c.v)}</td>`).join('')}</tr>`);
    }
    $('#sheet-content').innerHTML = rows.length ? `<div class="sheet-grid"><table aria-label="Conteúdo de ${esc(sheetName)}"><thead><tr><th>#</th>${Array.from({ length: endCol - range.s.c + 1 }, (_, i) => `<th>${M.col(i + range.s.c)}</th>`).join('')}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '<div class="empty-state"><p>Nenhuma linha encontrada.</p></div>';
  }
  function backup() {
    if (!book) { toast('Abra sua planilha primeiro para salvar um backup das metas.'); return; }
    const body = { application: 'saldo-local', version: 1, exportedAt: new Date().toISOString(), workbookName: book.filename, state: M.validState(local) };
    const blob = new Blob([JSON.stringify(body, null, 2)], { type: 'application/json' }), url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = 'saldo-backup-' + new Date().toISOString().slice(0, 10) + '.json'; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast('Backup exportado. Ele contém metas e anotações, sem a planilha.');
  }
  async function restore(file) {
    if (!file || !book) return;
    try {
      if (file.size > 2 * 1024 * 1024) throw new Error('O backup é grande demais. Escolha um backup exportado por este painel.');
      const data = JSON.parse(await file.text());
      if (data.application !== 'saldo-local' || data.version !== 1 || !data.state) throw new Error('Este arquivo não é um backup compatível com o painel.');
      const restored = M.validState(data.state);
      showInfo('Restaurar suas anotações?', `<p>Backup de <strong>${esc(data.workbookName || 'uma planilha')}</strong>, exportado em ${displayDate(new Date(data.exportedAt))}.</p><p>Há ${Object.keys(restored.goals).length} acompanhamentos e ${restored.customGoals.length} metas locais. A restauração substitui as anotações deste painel, sem modificar sua planilha.</p>`, '<button class="button" data-close-dialog="info-dialog">Cancelar</button><button class="button button-primary" id="confirm-restore">Restaurar backup</button>');
      $('#confirm-restore').addEventListener('click', () => { local = restored; save(); $('#info-dialog').close(); render(); toast('Backup restaurado neste navegador.'); }, { once: true });
    } catch (error) { toast(error.message || 'Não foi possível restaurar este backup.', true); }
    finally { $('#backup-input').value = ''; }
  }
  function demoWorkbook() {
    const wb = XLSX.utils.book_new();
    for (const yy of [2026, 2027]) {
      const rows = Array.from({ length: 88 }, () => []); let balance = yy === 2026 ? 1700 : 0;
      rows[45][5] = balance; rows[45][7] = 'Receita principal'; rows[45][8] = 4200; rows[45][10] = 'Receita complementar'; rows[45][11] = 800;
      rows[45][17] = new Date('2026-11-01T12:00:00');
      for (let i = 0; i < 12; i++) {
        const c = 2 + i % 6 * 3, start = i < 6 ? 2 : 24, totalRow = start + 14, netRow = start + 15, balanceRow = start + 16;
        const card = yy === 2026 ? [1500, 1760, 1400, 2450, 1380, 1120, 1600, 1870, 1300, 1210, 1100, 950][i] : Math.max(0, 850 - i * 100);
        const fixed = 1480, variable = yy === 2026 ? [880, 700, 920, 1300, 820, 700, 940, 650, 860, 900, 900, 900][i] : 950;
        rows[start][c - 1] = 'Cartão NuBank'; rows[start][c] = card; rows[start + 1][c - 1] = 'Gastos Fixos'; rows[start + 1][c] = fixed; rows[start + 2][c - 1] = 'Gastos Previstos'; rows[start + 2][c] = variable;
        const expense = card + fixed + variable, net = 5000 - expense; balance += net;
        rows[totalRow][c - 1] = 'Total Dívidas'; rows[totalRow][c] = expense; rows[netRow][c - 1] = 'Sobra'; rows[netRow][c] = net; rows[balanceRow][c - 1] = 'Sobra Total'; rows[balanceRow][c] = balance; rows[balanceRow + 2][c - 1] = M.shortCycle(i) + ' ' + yy;
      }
      rows[68][0] = 'Metas Atuais'; rows[69][1] = 'Metas Prioridade'; rows[69][4] = 'Metas Pessoal'; rows[69][7] = 'Metas TI/Setup';
      if (yy === 2026) { rows[70][1] = 'Reserva de emergência'; rows[70][2] = 10000; rows[70][4] = 'Viagem de férias'; rows[70][5] = 5000; rows[70][7] = 'Notebook novo'; rows[70][8] = 4500; }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Financeiro ' + yy);
    }
    return wb;
  }
  document.addEventListener('click', event => {
    const el = event.target.closest('button, [data-open-cycle], a[data-goto]'); if (!el) return;
    if (el.dataset.closeDialog) { $('#' + el.dataset.closeDialog).close(); return; }
    if (el.dataset.localAction) {
      $('#info-dialog').close();
      if (el.dataset.localAction === 'backup') backup();
      if (el.dataset.localAction === 'restore') { if (!book) toast('Abra sua planilha antes de restaurar as metas.'); else $('#backup-input').click(); }
      if (el.dataset.localAction === 'privacy') privacy();
      return;
    }
    if (el.dataset.view || el.dataset.goto) {
      const next = el.dataset.view || el.dataset.goto; if (!book) { toast('Abra sua planilha ou explore o exemplo para começar.'); return; }
      if (views.includes(next)) { view = next; location.hash = view; render(); window.scrollTo({ top: 0, behavior: 'auto' }); } return;
    }
    if (el.dataset.year) { year = Number(el.dataset.year); selectedCycle = Math.min(closedCount(), 11); render(); return; }
    if (el.dataset.openCycle !== undefined) { openCycle(Number(el.dataset.openCycle)); return; }
    if (el.dataset.goalTab) { goalTab = el.dataset.goalTab; render(); return; }
    if (el.hasAttribute('data-add-goal')) { openGoal(); return; }
    if (el.dataset.editGoal) { openGoal(el.dataset.editGoal); return; }
    if (el.dataset.addPlan) { openGoal(null, book.library.find(p => p.id === el.dataset.addPlan)); return; }
    if (el.hasAttribute('data-show-hidden')) { for (const p of Object.values(local.goals)) p.hidden = false; save(); render(); return; }
    if (el.id === 'hide-goal' && editingGoal?.id) { local.goals[editingGoal.id] = { ...goalProgress(editingGoal), hidden: true }; save(); $('#goal-dialog').close(); render(); toast('Meta oculta. Você pode mostrá-la novamente no fim da página de metas.'); }
  });
  document.addEventListener('change', event => {
    const el = event.target;
    if (el.id === 'closed-select') { local.closed[year] = Math.max(0, Math.min(12, Number(el.value))); selectedCycle = Math.min(closedCount(), 11); save(); render(); }
    if (el.hasAttribute('data-cycle-select')) { selectedCycle = Number(el.value); render(); }
    if (el.id === 'goal-category') { goalCategory = el.value; renderGoals(); }
    if (el.id === 'sheet-select') { sheetName = el.value; renderSheetTable(); }
    if (el.id === 'formula-mode') { formulaMode = el.checked; renderSheetTable(); }
  });
  document.addEventListener('input', event => {
    if (event.target.id === 'goal-search') { goalFilter = event.target.value; const filtered = activeGoals().filter(g => (!goalFilter || M.norm(g.name).includes(M.norm(goalFilter))) && (!goalCategory || g.category === goalCategory)); $('#goal-cards').innerHTML = filtered.map(goalCard).join('') || '<p class="small-note">Nenhuma meta encontrada.</p>'; }
    if (event.target.id === 'sheet-search') { sheetSearch = event.target.value; renderSheetTable(); }
    if (event.target.closest('#goal-form')) previewGoal();
  });
  document.addEventListener('submit', event => {
    if (event.target.id !== 'budget-form') return; event.preventDefault();
    if (!event.target.reportValidity()) return;
    const index = Number(event.target.dataset.cycle), cost = M.number($('#cycle-budget').value), key = `${year}-${index + 1}`;
    if (cost === null) delete local.budgets[key]; else local.budgets[key] = cost;
    save(); $('#cycle-dialog').close(); render(); toast('Orçamento salvo neste navegador.');
  });
  $('#goal-form').addEventListener('submit', submitGoal);
  $('#open-file').addEventListener('click', () => $('#file-input').click());
  $('#welcome-open').addEventListener('click', () => $('#file-input').click());
  $('#file-input').addEventListener('change', event => loadFile(event.target.files[0]));
  $('#demo-button').addEventListener('click', () => { if (!window.XLSX || !M) { toast('Extraia todos os arquivos do ZIP antes de abrir o painel.', true); return; } activate(M.parse(demoWorkbook(), 'Exemplo demonstrativo.xlsx'), true); });
  $('#backup-button').addEventListener('click', backup);
  $('#restore-button').addEventListener('click', () => { if (!book) { toast('Abra sua planilha antes de restaurar as metas.'); return; } $('#backup-input').click(); });
  $('#backup-input').addEventListener('change', event => restore(event.target.files[0]));
  $('#privacy-button').addEventListener('click', privacy);
  $('#mobile-tools').addEventListener('click', () => showInfo('Suas metas, no seu navegador', '<p>Exporte o progresso das suas metas para guardar uma cópia ou levar suas anotações para outro dispositivo.</p><div class="mobile-tool-list"><button class="button" data-local-action="backup">Exportar backup</button><button class="button" data-local-action="restore">Restaurar backup</button><button class="button" data-local-action="privacy">Sobre a privacidade</button></div>'));
  $('#source-info').addEventListener('click', () => book ? showInfo('Sobre os dados exibidos', `<p><strong>${esc(book.filename)}</strong> · ${book.sheetNames.length} abas.</p><p>O painel reconheceu ${book.years.length} anos financeiros. A data registrada em ${esc(current().name)} é <strong>${displayExcelDate(current().updatedAt)}</strong>. Essa data é um campo da planilha, não uma confirmação de atualização.</p><p>As fórmulas não são recalculadas aqui. O leitor mostra os resultados salvos no .xlsx. Para refletir mudanças no OneDrive, baixe e abra novamente o arquivo.</p><p>O marco inicial de 2026 foi definido até Out–Nov, conforme seu fechamento informado. Ajuste o seletor quando necessário. A seleção é local e não altera a planilha.</p>`) : privacy());
  const drop = $('#dropzone');
  for (const name of ['dragenter', 'dragover']) drop.addEventListener(name, event => { event.preventDefault(); drop.classList.add('dragover'); });
  for (const name of ['dragleave', 'drop']) drop.addEventListener(name, event => { event.preventDefault(); drop.classList.remove('dragover'); });
  drop.addEventListener('drop', event => loadFile(event.dataTransfer.files[0]));
  window.addEventListener('dragover', e => e.preventDefault()); window.addEventListener('drop', e => e.preventDefault());
  window.addEventListener('hashchange', () => { const next = location.hash.slice(1); if (book && views.includes(next) && next !== view) { view = next; render(); } });
  document.querySelector('.brand').addEventListener('click', () => { if (book) { view = 'overview'; render(); } });
  document.querySelectorAll('dialog').forEach(dialog => dialog.addEventListener('click', event => { if (event.target === dialog) { const r = dialog.getBoundingClientRect(); if (event.clientX < r.left || event.clientX > r.right || event.clientY < r.top || event.clientY > r.bottom) dialog.close(); } }));
  icons();
  if (!M || !window.XLSX) toast('Algum arquivo do projeto não carregou. Extraia o ZIP inteiro e abra o index.html dentro da pasta.', true);
})();
