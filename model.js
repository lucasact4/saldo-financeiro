/* Leitura do formato da planilha original. Nenhuma célula é alterada. */
(function (root) {
  'use strict';
  const norm = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase().replace(/\s+/g, ' ');
  const hash = value => { let n = 2166136261; for (const c of String(value)) n = Math.imul(n ^ c.charCodeAt(0), 16777619); return (n >>> 0).toString(36); };
  const number = value => {
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;
    if (typeof value !== 'string' || !value.trim()) return null;
    let s = value.trim().replace(/R\$|\s/g, '');
    const negative = /^\(.*\)$/.test(s); s = s.replace(/[()]/g, '');
    if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    const n = Number(s); return Number.isFinite(n) ? (negative ? -n : n) : null;
  };
  const sum = values => values.reduce((a, b) => a + (typeof b === 'number' ? b : 0), 0);
  const validDate = value => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return false;
    const [y,m,d] = value.split('-').map(Number), date = new Date(y,m-1,d);
    return y >= 2000 && y <= 2100 && date.getFullYear() === y && date.getMonth() === m-1 && date.getDate() === d;
  };
  const col = n => { let s = ''; for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const at = (sheet, row, column) => sheet[col(column) + row];
  const value = (sheet, row, column) => at(sheet, row, column)?.v ?? null;
  const amount = (sheet, row, column) => number(value(sheet, row, column));
  const text = (sheet, row, column) => String(value(sheet, row, column) ?? '').trim();
  const months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  const shortCycle = i => `${months[i]}–${months[(i + 1) % 12]}`;
  const classify = name => {
    const s = norm(name);
    if (s.includes('cartao') || s === 'nubank') return 'Cartão';
    if (s.includes('gastos fixos')) return 'Gastos fixos';
    if (s.includes('previsto')) return 'Gastos previstos';
    return 'Outros lançamentos';
  };
  function financial(sheet, name, year) {
    const anchors = [];
    for (const address of Object.keys(sheet)) {
      if (address[0] === '!' || norm(sheet[address]?.v) !== 'sobra total') continue;
      const match = address.match(/^([A-Z]+)(\d+)$/); if (!match) continue;
      let column = 0; for (const c of match[1]) column = column * 26 + c.charCodeAt(0) - 64;
      const row = Number(match[2]), labelCol = column - 1;
      if (norm(value(sheet, row - 1, labelCol)) !== 'sobra') continue;
      anchors.push({ row, labelCol });
    }
    anchors.sort((a, b) => a.row - b.row || a.labelCol - b.labelCol);
    if (anchors.length !== 12) throw new Error(`A aba “${name}” não tem os 12 blocos de “Sobra Total” esperados. Use a planilha original.`);
    const issues = [];
    const periods = anchors.map((anchor, index) => {
      const { row, labelCol } = anchor, c = labelCol + 1;
      const expense = amount(sheet, row - 2, c), net = amount(sheet, row - 1, c), balance = amount(sheet, row, c);
      if ([expense, net, balance].some(v => v === null)) throw new Error(`Faltam resultados numéricos em ${name}, ciclo ${shortCycle(index)}. Salve o arquivo no Excel Web e baixe outra cópia com os resultados das fórmulas.`);
      const items = [];
      for (let r = row - 16; r <= row - 3; r++) {
        const label = text(sheet, r, labelCol), cost = amount(sheet, r, c);
        if (label && cost !== null) items.push({ name: label, amount: cost, group: cost < 0 ? 'Créditos e reembolsos' : classify(label), cell: col(c) + r });
      }
      if (Math.abs(sum(items.map(i => i.amount)) - expense) > .02) issues.push({ type: 'total', index, message: `${shortCycle(index)}: o total registrado difere da soma dos itens visíveis.` });
      return { index, year, label: shortCycle(index), fullLabel: text(sheet, row + 2, labelCol) || `${shortCycle(index)} ${year}`, expense, net, balance, income: expense + net, items, sheet: name, expenseCell: col(c) + (row - 2), netCell: col(c) + (row - 1), balanceCell: col(c) + row };
    });
    const opening = amount(sheet, 46, 5) ?? periods[0].balance - periods[0].net;
    for (const [i, p] of periods.entries()) if (Math.abs((i ? periods[i - 1].balance : opening) + p.net - p.balance) > .02) issues.push({ type: 'balance', index: i, message: `${p.label}: confira a continuidade do saldo acumulado no arquivo.` });
    const sharedHistoricalInputs = anchors.some(a => /^.*R67/i.test(String(at(sheet, a.row - 15, a.labelCol + 1)?.f || '')));
    return { year, name, opening, periods, issues, sharedHistoricalInputs, updatedAt: value(sheet, 46, 17) };
  }
  function readGoals(sheet, sheetName, year) {
    let headerRow = 0;
    for (let r = 1; r < 300; r++) if (norm(value(sheet, r, 0)).startsWith('metas atuais')) { headerRow = r + 1; break; }
    if (!headerRow) return [];
    const found = new Map();
    for (let c = 0; c < 24; c++) {
      const heading = text(sheet, headerRow, c); if (!norm(heading).startsWith('metas ')) continue;
      const priorityList = norm(heading).includes('prioridade');
      const category = heading.replace(/^Metas\s+/i, '').replace('TI/Setup', 'TI / Setup').replace('Prioridade', 'Outros');
      for (let r = headerRow + 1; r < headerRow + 30; r++) {
        const name = text(sheet, r, c), target = amount(sheet, r, c + 1);
        if (norm(name).startsWith('total ')) break;
        if (!name || target === null || target <= 0) continue;
        const key = norm(name);
        if (!found.has(key)) found.set(key, { id: 'sheet-' + hash(key), name, target, category, priority: priorityList ? 'Alta' : 'A definir', source: `${sheetName}!${col(c)}${r}`, sourceYear: year, imported: true });
        else { const g = found.get(key); if (priorityList) g.priority = 'Alta'; else g.category = category; }
      }
    }
    return [...found.values()];
  }
  function readLibrary(workbook) {
    const plans = [];
    const livingName = workbook.SheetNames.find(n => norm(n) === 'morar so');
    if (livingName) {
      const s = workbook.Sheets[livingName];
      for (const [labelCol, first, last, title] of [[1, 3, 20, 'Morar só · Essencial'], [4, 3, 20, 'Morar só · Depois da mudança'], [7, 3, 20, 'Morar só · Conforto'], [1, 25, 38, 'Morar só · Comprar antes']]) {
        const items = [];
        for (let r = first; r <= last; r++) { const name = text(s, r, labelCol), cost = amount(s, r, labelCol + 1); if (name && cost !== null && cost > 0) items.push({ name, target: cost }); }
        if (items.length) plans.push({ id: 'plan-' + hash(title), name: title, category: 'Moradia', target: sum(items.map(i => i.target)), items, source: livingName });
      }
    }
    for (const name of ['Setup', 'Quarto', 'Outros', 'Treino']) {
      const sheet = workbook.Sheets[name]; if (!sheet) continue;
      const items = [], priceCol = name === 'Treino' ? 2 : 3;
      for (let r = 2; r <= 40; r++) { const title = text(sheet, r, 1), price = amount(sheet, r, priceCol); if (title && !['nome', 'valor', 'preco'].includes(norm(title)) && price !== null && price > 0) items.push({ name: title, target: price }); }
      if (items.length) plans.push({ id: 'plan-' + hash(name), name, category: name === 'Setup' ? 'TI / Setup' : name === 'Treino' ? 'Saúde' : 'Pessoal', items, target: sum(items.map(i => i.target)), source: name });
    }
    return plans;
  }
  function readInstallments(workbook, sheetName) {
    const s = workbook.Sheets[sheetName], items = [];
    if (!s) return { items, total: 0, source: sheetName, date: null };
    for (let r = 50; r <= 66; r++) {
      const label = text(s, r, 10), monthly = amount(s, r, 11), match = label.match(/\b(\d+)\s*\/\s*(\d+)\b/);
      if (!match || monthly === null || monthly <= 0) continue;
      const current = Number(match[1]), count = Number(match[2]); if (current < 1 || count < current || count > 600) continue;
      items.push({ name: label.replace(/\s*[-–]?\s*\d+\s*\/\s*\d+\s*$/, '').trim(), current, count, monthly, remaining: count - current, remainingValue: (count - current) * monthly });
    }
    return { items, total: sum(items.map(i => i.monthly)), remaining: sum(items.map(i => i.remainingValue)), source: sheetName, date: value(s, 49, 11) };
  }
  function parse(workbook, filename) {
    const years = workbook.SheetNames.map(name => ({ name, match: name.match(/^Financeiro\s+(20\d{2})$/i) })).filter(i => i.match).map(i => financial(workbook.Sheets[i.name], i.name, Number(i.match[1]))).sort((a, b) => a.year - b.year);
    if (!years.length) throw new Error('Não encontrei abas como “Financeiro 2026”. Escolha sua planilha financeira original.');
    const goalsByYear = years.map(y => ({ year: y.year, goals: readGoals(workbook.Sheets[y.name], y.name, y.year) }));
    const goalSource = [...goalsByYear].reverse().find(y => y.goals.length);
    const schedules = years.map(y => readInstallments(workbook, y.name)).filter(s => s.items.length);
    return { filename, key: 'book-' + hash(norm(filename.replace(/\.xlsx$/i, '').replace(/\s*\(\d+\)$/, ''))), years, goals: goalSource?.goals || [], goalSourceYear: goalSource?.year, library: readLibrary(workbook), schedules, sheetNames: workbook.SheetNames, workbook };
  }
  function goalStats(goal, progress = {}, now = new Date()) {
    const saved = number(progress.saved), monthly = number(progress.monthly), target = number(goal.target);
    const remaining = saved !== null && target !== null ? Math.max(0, target - saved) : null;
    const deadline = validDate(progress.deadline) ? new Date(progress.deadline + 'T23:59:59') : null;
    const monthsToGoal = remaining === 0 ? 0 : remaining !== null && monthly > 0 ? Math.ceil(remaining / monthly) : null;
    const monthDiff = deadline ? Math.max(1, (deadline.getFullYear() - now.getFullYear()) * 12 + deadline.getMonth() - now.getMonth() + (deadline.getDate() > now.getDate() ? 1 : 0)) : null;
    const suggested = remaining !== null && deadline && deadline >= now ? remaining / monthDiff : null;
    let status = saved === null ? 'Informe o valor guardado' : remaining === 0 ? 'Concluída' : deadline && deadline < now ? 'Prazo vencido' : monthly === null ? 'Defina um aporte' : monthly === 0 ? 'Sem aporte mensal' : !deadline ? 'Defina um prazo' : monthly + .005 < suggested ? 'Revise prazo ou aporte' : 'Dentro do plano';
    return { saved, monthly, remaining, monthsToGoal, suggested, deadline, progress: saved !== null && target > 0 ? Math.min(100, saved / target * 100) : null, status };
  }
  function validProgress(input) {
    const result = {};
    if (!input || typeof input !== 'object' || Array.isArray(input)) return result;
    for (const [id, v] of Object.entries(input).slice(0, 500)) {
      if (!/^(sheet|custom|plan)-[a-z0-9-]+$/.test(id) || !v || typeof v !== 'object') continue;
      const saved = number(v.saved), monthly = number(v.monthly);
      result[id] = { saved: saved !== null && saved >= 0 && saved <= 1e12 ? saved : null, monthly: monthly !== null && monthly >= 0 && monthly <= 1e12 ? monthly : null, deadline: validDate(v.deadline) ? v.deadline : '', priority: ['Alta', 'Média', 'Baixa', 'A definir'].includes(v.priority) ? v.priority : '', hidden: v.hidden === true };
    }
    return result;
  }
  function validState(input) {
    const result = { closed: {}, goals: validProgress(input?.goals), customGoals: [], budgets: {} };
    for (const [key, n] of Object.entries(input?.closed || {})) if (/^20\d{2}$/.test(key) && Number.isInteger(n) && n >= 0 && n <= 12) result.closed[key] = n;
    for (const [key, raw] of Object.entries(input?.budgets || {})) { const n = number(raw); if (/^20\d{2}-\d{1,2}$/.test(key) && n !== null && n >= 0 && n <= 1e12) result.budgets[key] = n; }
    for (const g of (Array.isArray(input?.customGoals) ? input.customGoals.slice(0, 200) : [])) {
      const target = number(g.target); if (!/^custom-[a-z0-9-]+$/.test(g.id || '') || typeof g.name !== 'string' || !g.name.trim() || target === null || target <= 0 || target > 1e12) continue;
      result.customGoals.push({ id: g.id, name: g.name.slice(0, 140), target, category: String(g.category || 'Outros').slice(0, 50), priority: ['Alta', 'Média', 'Baixa'].includes(g.priority) ? g.priority : 'A definir', imported: false });
    }
    return result;
  }
  root.FinanceModel = { norm, hash, number, sum, col, months, shortCycle, parse, goalStats, validState };
})(typeof globalThis !== 'undefined' ? globalThis : window);
