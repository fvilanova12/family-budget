// Our Budget — shared household budget for two.
// Plain JavaScript, no build step. Data lives in Supabase (protected by Row
// Level Security); receipts are read on the device (see ocr.js).
(function () {
  'use strict';

  // =====================================================================
  // Data layer — Supabase. demo.js provides the same interface in memory.
  // =====================================================================
  function SupaAPI(cfg) {
    const sb = window.supabase.createClient(cfg.supabaseUrl, cfg.supabaseKey, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
    });
    const run = async (p) => {
      const { data, error } = await p;
      if (error) throw error;
      return data;
    };
    const filt = (q, o) => {
      o = o || {};
      for (const k in o.eq || {}) q = q.eq(k, o.eq[k]);
      for (const k in o.gte || {}) q = q.gte(k, o.gte[k]);
      for (const k in o.gt || {}) q = q.gt(k, o.gt[k]);
      for (const k in o.lt || {}) q = q.lt(k, o.lt[k]);
      for (const k in o.lte || {}) q = q.lte(k, o.lte[k]);
      for (const k of o.notnull || []) q = q.not(k, 'is', null);
      for (const [col, asc] of o.order || []) q = q.order(col, { ascending: asc });
      if (o.limit) q = q.limit(o.limit);
      return q;
    };
    return {
      async session() { return (await sb.auth.getSession()).data.session; },
      onAuth(cb) { sb.auth.onAuthStateChange((event, session) => cb(event, session)); },
      signIn: (email, password) => run(sb.auth.signInWithPassword({ email, password })),
      signOut: () => sb.auth.signOut(),
      resetPassword: (email) => run(sb.auth.resetPasswordForEmail(email, { redirectTo: location.origin + location.pathname })),
      updatePassword: (password) => run(sb.auth.updateUser({ password })),
      select: (t, o) => run(filt(sb.from(t).select('*'), o)),
      insert: (t, row) => run(sb.from(t).insert(row).select().single()),
      update: (t, id, patch) => run(sb.from(t).update(patch).eq(t === 'members' ? 'user_id' : 'id', id).select().single()),
      updateWhere: (t, o, patch) => run(filt(sb.from(t).update(patch), o)),
      upsertBudget: (row) => run(sb.from('category_budgets').upsert(row, { onConflict: 'category_id,month' })),
      remove: (t, id) => run(sb.from(t).delete().eq('id', id)),
      payBill: (id, amount, date, method) => run(sb.rpc('pay_bill', { p_bill_id: id, p_amount: amount, p_paid_on: date, p_payment_method: method || null })),
      upload: (path, blob) => run(sb.storage.from('receipts').upload(path, blob, { contentType: 'image/jpeg', upsert: false })),
      signedUrl: async (path) => (await run(sb.storage.from('receipts').createSignedUrl(path, 600))).signedUrl,
      removeFiles: (paths) => run(sb.storage.from('receipts').remove(paths)),
      subscribe(cb) { sb.channel('household').on('postgres_changes', { event: '*', schema: 'public' }, cb).subscribe(); },
    };
  }

  const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';

  // =====================================================================
  // Helpers
  // =====================================================================
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const pad = (n) => String(n).padStart(2, '0');
  const isoOf = (d) => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const todayISO = () => isoOf(new Date());
  const monthOf = (iso) => iso.slice(0, 7) + '-01';
  const addMonths = (m, n) => { const [y, mo] = m.split('-').map(Number); return isoOf(new Date(y, mo - 1 + n, 1)); };
  const addDays = (iso, n) => { const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() + n); return isoOf(d); };
  const daysBetween = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 86400000);
  const fmtDate = (iso) => (iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '');
  const fmtDM = (iso) => (iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) : '');
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const MON = MONTHS.map((m) => m.slice(0, 3));
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const monthLabel = (m, forceYear) => {
    const y = +m.slice(0, 4);
    return MONTHS[+m.slice(5, 7) - 1] + (forceYear || y !== new Date().getFullYear() ? ' ' + y : '');
  };
  const dayLabel = (iso) => {
    const diff = daysBetween(iso, todayISO());
    if (diff === 0) return 'Today';
    if (diff === 1) return 'Yesterday';
    const d = new Date(iso + 'T12:00:00');
    return DAYS[d.getDay()] + ' ' + d.getDate() + ' ' + MON[d.getMonth()] + (d.getFullYear() !== new Date().getFullYear() ? ' ' + d.getFullYear() : '');
  };
  const round2 = (n) => Math.round(n * 100) / 100;
  const NF2 = new Intl.NumberFormat('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const NF0 = new Intl.NumberFormat('de-DE', { maximumFractionDigits: 0 });
  // European format: 1.234,56 €
  const eur = (n, opts) => {
    opts = opts || {};
    const a = Math.abs(n || 0);
    const whole = opts.smart && Math.abs(a - Math.round(a)) < 0.005;
    let s = (opts.dec === false || whole ? NF0 : NF2).format(a) + ' €';
    if (n < 0) s = '−' + s;
    else if (opts.sign && n > 0) s = '+' + s;
    return s;
  };
  const numText = (n) => (n == null || n === '' ? '' : Math.abs(n - Math.round(n)) < 0.005 ? NF0.format(n) : NF2.format(n));
  // Accepts 1.234,56 · 1234,56 · 1234.56 · 1.700
  function parseAmount(s) {
    s = String(s ?? '').replace(/[€\s]/g, '');
    if (!s) return NaN;
    if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
    const v = parseFloat(s);
    return isNaN(v) ? NaN : round2(v);
  }
  const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now() + '-' + Math.random().toString(16).slice(2));

  // ---------- Icons (stroke SVG, as in the design) ----------
  const P = {
    home: '<path d="M3 12l9-8 9 8"/><path d="M5 10v10h14V10"/>',
    cart: '<circle cx="9" cy="20" r="1.5"/><circle cx="18" cy="20" r="1.5"/><path d="M2 3h3l2.5 12h11.5l2-8H6.2"/>',
    fork: '<path d="M7 3v8a2 2 0 0 0 4 0V3"/><path d="M9 11v10"/><path d="M17 3c-2 0-3 2-3 6s1 4 3 4v8"/>',
    star: '<path d="M12 3l2.6 5.6 6 .7-4.5 4.1 1.2 6L12 16.5 6.7 19.4l1.2-6L3.4 9.3l6-.7z"/>',
    car: '<path d="M5 17h14v-5l-2-5H7l-2 5z"/><circle cx="7.5" cy="17.5" r="1.5"/><circle cx="16.5" cy="17.5" r="1.5"/><path d="M5 12h14"/>',
    bolt: '<path d="M13 2L4 14h7l-1 8 9-12h-7z"/>',
    heart: '<path d="M12 20s-7-4.4-7-10a4 4 0 0 1 7-2.6A4 4 0 0 1 19 10c0 5.6-7 10-7 10z"/>',
    school: '<path d="M2 9l10-5 10 5-10 5z"/><path d="M6 11v5c3 2 9 2 12 0v-5"/>',
    tennis: '<circle cx="12" cy="12" r="9"/><path d="M5.6 5.6c3 3 3 9.8 0 12.8M18.4 5.6c-3 3-3 9.8 0 12.8"/>',
    shirt: '<path d="M8 3l-5 3 2 4 3-1v12h8V9l3 1 2-4-5-3c0 2-1.5 3-4 3S8 5 8 3z"/>',
    phone: '<rect x="7" y="2" width="10" height="20" rx="2"/><path d="M11 18h2"/>',
    sparkle: '<path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2.5 2.5M15.5 15.5L18 18M6 18l2.5-2.5M15.5 8.5L18 6"/>',
    box: '<path d="M3 7l9-4 9 4v10l-9 4-9-4z"/><path d="M3 7l9 4 9-4M12 11v10"/>',
    briefcase: '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M9 7V5h6v2M3 12h18"/>',
    coin: '<circle cx="12" cy="12" r="9"/><path d="M15 9.5c-.5-1-1.6-1.5-3-1.5-2 0-3 1-3 2s1 1.6 3 2 3 1 3 2-1 2-3 2c-1.4 0-2.5-.5-3-1.5M12 6v2M12 16v2"/>',
    user: '<circle cx="12" cy="8" r="4"/><path d="M4 21c1-4 4-6 8-6s7 2 8 6"/>',
    plus: '<path d="M12 5v14M5 12h14"/>',
    camera: '<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
    close: '<path d="M6 6l12 12M18 6L6 18"/>',
    left: '<path d="M15 18l-6-6 6-6"/>',
    right: '<path d="M9 18l6-6-6-6"/>',
    cal: '<rect x="3" y="5" width="18" height="16" rx="3"/><path d="M3 10h18M8 3v4M16 3v4"/>',
    search: '<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>',
    check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    list: '<path d="M8 6h13M8 12h13M8 18h13"/><circle cx="3.5" cy="6" r="1"/><circle cx="3.5" cy="12" r="1"/><circle cx="3.5" cy="18" r="1"/>',
    coins: '<ellipse cx="12" cy="6" rx="7" ry="3"/><path d="M5 6v6c0 1.7 3.1 3 7 3s7-1.3 7-3V6"/><path d="M5 12v6c0 1.7 3.1 3 7 3s7-1.3 7-3v-6"/>',
    pie: '<path d="M21 12a9 9 0 1 1-9-9v9z"/><path d="M15 3.5A9 9 0 0 1 20.5 9H15z"/>',
    photos: '<rect x="3" y="4" width="18" height="16" rx="3"/><circle cx="9" cy="10" r="1.8"/><path d="M21 16l-5-5-9 9"/>',
    keyboard: '<rect x="3" y="6" width="18" height="12" rx="2"/><path d="M7 10h.01M11 10h.01M15 10h.01M8 14h8"/>',
    trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
    download: '<path d="M12 3v12M7 10l5 5 5-5M5 21h14"/>',
    repeat: '<path d="M17 2l4 4-4 4"/><path d="M3 11V9a3 3 0 0 1 3-3h15M7 22l-4-4 4-4"/><path d="M21 13v2a3 3 0 0 1-3 3H3"/>',
    logout: '<path d="M15 4h4a1 1 0 0 1 1 1v14a1 1 0 0 1-1 1h-4M10 17l5-5-5-5M15 12H3"/>',
    key: '<circle cx="8" cy="15" r="4"/><path d="M11 12l9-9M17 6l3 3"/>',
    card: '<rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20"/>',
    cash: '<rect x="2" y="6" width="20" height="12" rx="2"/><circle cx="12" cy="12" r="2.5"/>',
    bank: '<path d="M3 10l9-6 9 6M5 10v8M9 10v8M15 10v8M19 10v8M3 20h18"/>',
    clip: '<path d="M21 11l-8.5 8.5a5 5 0 0 1-7-7L14 4a3.5 3.5 0 0 1 5 5l-8.5 8.5a2 2 0 0 1-3-3L15 7"/>',
    goal: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/>',
    note: '<path d="M4 4h16v12l-4 4H4z"/><path d="M16 20v-4h4M8 9h8M8 13h5"/>',
  };
  const svg = (name, size, sw) => '<svg aria-hidden="true" width="' + (size || 20) + '" height="' + (size || 20) + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="' + (sw || 2) + '" stroke-linecap="round" stroke-linejoin="round">' + (P[name] || P.box) + '</svg>';

  function catIcon(c) {
    if (!c) return 'box';
    const n = Receipt.norm(c.name);
    if (c.kind === 'income') return /work|salary|nomina|project/.test(n) ? 'briefcase' : S.members.some((m) => Receipt.norm(m.display_name) === n) ? 'user' : 'coin';
    const map = [[/hous|home|alquiler|rent|mortgage|hipoteca/, 'home'], [/grocer|super|food/, 'cart'], [/tennis|tenis|\bsports?\b/, 'tennis'],
      [/child|educ|school|colegio|kid/, 'school'], [/health|insur|medic|seguro/, 'heart'], [/transport|\bcar\b|fuel|gasol|parking/, 'car'],
      [/utilit|electric|water|luz|agua/, 'bolt'], [/dining|eat|restaur|cafe/, 'fork'], [/cloth|ropa/, 'shirt'],
      [/entertain|subs|fun|leisure|ocio/, 'star'], [/phone|internet|movil/, 'phone'], [/personal|care|beauty/, 'sparkle']];
    for (const [re, icon] of map) if (re.test(n)) return icon;
    return 'box';
  }

  const METHODS = [['card', 'Card', 'card'], ['cash', 'Cash', 'cash'], ['transfer', 'Transfer', 'bank'], ['direct_debit', 'Direct debit', 'repeat']];
  const methodLabel = (m) => (METHODS.find((x) => x[0] === m) || [null, ''])[1];
  const FREQS = [['weekly', 'Weekly'], ['monthly', 'Monthly'], ['bimonthly', 'Every 2 months'], ['quarterly', 'Every 3 months'], ['yearly', 'Yearly']];
  const freqLabel = (f) => (FREQS.find((x) => x[0] === f) || [null, f])[1];

  function blend(hex, base, k) {
    const h = (x) => [1, 3, 5].map((i) => parseInt(x.slice(i, i + 2), 16));
    const a = h(hex), b = h(base);
    return '#' + a.map((v, i) => pad(Math.round(v * k + b[i] * (1 - k)).toString(16))).join('');
  }

  // =====================================================================
  // State
  // =====================================================================
  const S = {
    api: null, user: null, me: null, ready: false,
    members: [], cats: [], bills: [], goals: [], contribs: [], settings: { monthly_savings_target: 0 }, rules: [], budgetRows: [],
    month: monthOf(todayISO()), txs: [], totals: [],
    route: 'home', params: {}, sheet: null, form: null, scan: null, toast: null,
    hist: { who: 'all', q: '', search: false }, budgetScope: 'forward',
    exp: { from: new Date().getFullYear() + '-01-01', to: todayISO() },
    login: { email: '', password: '', err: '', busy: false }, busy: false,
  };
  const api = () => S.api;

  const member = (id) => S.members.find((m) => m.user_id === id);
  const membersOrdered = () => (S.me ? [S.me].concat(S.members.filter((m) => m.user_id !== S.me.user_id)) : S.members);
  const cat = (id) => S.cats.find((c) => c.id === id);
  const expenseCats = () => S.cats.filter((c) => c.kind === 'expense' && !c.archived);
  const incomeCats = () => S.cats.filter((c) => c.kind === 'income' && !c.archived);
  function avatar(m, cls) {
    if (!m) return '';
    const fg = m.color && /^#[0-9a-f]{6}$/i.test(m.color) ? m.color : '#C8F03C';
    return '<span class="avatar ' + (cls || '') + '" style="background:' + blend(fg, '#0E0F12', 0.16) + ';color:' + fg + '" title="' + esc(m.display_name) + '">' + esc((m.display_name || '?').trim()[0].toUpperCase()) + '</span>';
  }

  // Budget in force for a month: the latest one set on or before it. Months
  // before the first budget use the earliest one, so history stays comparable.
  function budgetFor(catId, month) {
    let best = null, first = null;
    for (const r of S.budgetRows) {
      if (r.category_id !== catId) continue;
      if (r.month <= month && (!best || r.month > best.month)) best = r;
      if (!first || r.month < first.month) first = r;
    }
    return best ? +best.amount : first ? +first.amount : 0;
  }
  const totalBudget = (month) => expenseCats().reduce((s, c) => s + budgetFor(c.id, month), 0);

  function monthStats() {
    const spentBy = {};
    let spent = 0, income = 0;
    for (const t of S.txs) {
      if (t.kind === 'expense') { spent += +t.amount; spentBy[t.category_id] = (spentBy[t.category_id] || 0) + +t.amount; }
      else income += +t.amount;
    }
    return { spent: round2(spent), income: round2(income), spentBy, budget: round2(totalBudget(S.month)) };
  }

  // Totals per month from the monthly_category_totals view.
  function totalsByMonth() {
    const out = {};
    for (const r of S.totals) {
      const o = out[r.month] || (out[r.month] = { income: 0, expense: 0, byCat: {} });
      o[r.kind] += +r.total;
      if (r.kind === 'expense') o.byCat[r.category_id] = (o.byCat[r.category_id] || 0) + +r.total;
    }
    return out;
  }

  // =====================================================================
  // Loading
  // =====================================================================
  async function loadBase() {
    const [members, cats, bills, goals, contribs, settings, rules, budgetRows] = await Promise.all([
      api().select('members'),
      api().select('categories', { order: [['sort_order', true]] }),
      api().select('recurring_bills', { order: [['next_due_on', true]] }),
      api().select('savings_goals', { order: [['created_at', true]] }),
      api().select('savings_contributions', { order: [['occurred_on', false]] }),
      api().select('household_settings', { eq: { id: 1 } }),
      api().select('merchant_rules'),
      api().select('category_budgets', { order: [['month', false]] }),
    ]);
    Object.assign(S, { members, cats, bills, goals, contribs, rules, budgetRows, settings: settings[0] || { monthly_savings_target: 0 } });
    S.me = member(S.user.id) || null;
  }

  async function loadMonth() {
    const cur = monthOf(todayISO());
    const from = addMonths(S.month < cur ? S.month : cur, -12);
    const [txs, totals] = await Promise.all([
      api().select('transactions', { gte: { occurred_on: S.month }, lt: { occurred_on: addMonths(S.month, 1) }, order: [['occurred_on', false], ['created_at', false]] }),
      api().select('monthly_category_totals', { gte: { month: from } }),
    ]);
    S.txs = txs;
    S.totals = totals;
  }

  let reloadTimer = null;
  function onRemoteChange() {
    clearTimeout(reloadTimer);
    reloadTimer = setTimeout(async () => {
      try { await loadBase(); await loadMonth(); render(); } catch (e) { /* offline: keep what we have */ }
    }, 400);
  }

  async function refresh() {
    await loadBase();
    await loadMonth();
  }

  // =====================================================================
  // Routing
  // =====================================================================
  function parseHash() {
    const h = location.hash.replace(/^#\/?/, '');
    if (!h || /access_token|error_description|type=recovery/.test(h)) return { route: 'home', params: {} };
    const [path, query] = h.split('?');
    const parts = path.split('/');
    const params = Object.fromEntries(new URLSearchParams(query || ''));
    if (parts[0] === 'edit') params.id = parts[1];
    return { route: parts[0] || 'home', params };
  }

  function newForm(kind) {
    return {
      id: null, kind: kind || 'expense', amountStr: '', amountText: null, occurred_on: todayISO(), paid_by: S.me ? S.me.user_id : null,
      category_id: null, merchant: '', note: '', payment_method: kind === 'income' ? 'transfer' : 'card', catTouched: false,
      showNote: false, receipt_path: null, recurring_bill_id: null, refund: false,
    };
  }

  function reloadIfIdle() {
    if (S.updateReady && !S.sheet && !S.busy && ['home', 'history', 'savings', 'budgets'].includes(S.route)) location.reload();
  }

  function onRoute() {
    const { route, params } = parseHash();
    if (['home', 'history', 'savings', 'budgets'].includes(route)) S.lastMain = route;
    S.route = route;
    S.params = params;
    S.sheet = S.recovery ? { type: 'password', recovery: true } : null;
    if (route === 'add') S.form = newForm(params.kind === 'income' ? 'income' : 'expense');
    if (route === 'edit') {
      const t = S.txs.find((x) => x.id === params.id);
      if (!t) { location.hash = '#/history'; return; }
      S.form = Object.assign(newForm(t.kind), {
        id: t.id, kind: t.kind, refund: +t.amount < 0, amountStr: Number.isInteger(round2(Math.abs(t.amount))) ? String(round2(Math.abs(t.amount))) : round2(Math.abs(t.amount)).toFixed(2), occurred_on: t.occurred_on, paid_by: t.paid_by,
        category_id: t.category_id, merchant: t.merchant || '', note: t.note || '', payment_method: t.payment_method,
        catTouched: true, showNote: !!t.note, receipt_path: t.receipt_path, recurring_bill_id: t.recurring_bill_id,
      });
    }
    if (route === 'scan') {
      if (S.scan && S.scan.previewUrl) URL.revokeObjectURL(S.scan.previewUrl);
      S.scan = { step: 'pick' };
    }
    render();
    window.scrollTo(0, 0);
    reloadIfIdle();
  }
  const go = (hash) => { if (location.hash === hash) onRoute(); else location.hash = hash; };

  // =====================================================================
  // Rendering
  // =====================================================================
  function toast(msg, error) {
    S.toast = { msg, error: !!error };
    renderToast();
    clearTimeout(toast.t);
    toast.t = setTimeout(() => { S.toast = null; renderToast(); }, error ? 4500 : 2600);
  }
  function renderToast() {
    let el = document.getElementById('toast');
    if (!S.toast) { if (el) el.remove(); return; }
    if (!el) { el = document.createElement('div'); el.id = 'toast'; el.setAttribute('role', 'status'); document.body.appendChild(el); }
    el.className = 'toast' + (S.toast.error ? ' error' : '');
    el.textContent = S.toast.msg;
  }
  function errMsg(e) {
    const m = (e && (e.message || e.error_description)) || String(e);
    if (/Invalid login credentials/i.test(m)) return 'Wrong email or password.';
    if (/Failed to fetch|NetworkError|Load failed/i.test(m)) return 'No connection. Check your internet and try again.';
    if (/duplicate key/i.test(m)) return 'That name already exists.';
    return m;
  }

  function render() {
    const app = document.getElementById('app');
    const f = document.activeElement;
    const focus = f && f.id ? { id: f.id, s: f.selectionStart, e: f.selectionEnd } : null;
    let html;
    if (!S.user) html = viewLogin();
    else if (!S.ready) html = '<div class="boot">Loading…</div>';
    else if (!S.me) html = viewNotMember();
    else {
      const views = { home: viewHome, history: viewHistory, savings: viewSavings, budgets: viewBudgets, add: viewAdd, edit: viewAdd, scan: viewScan };
      html = (views[S.route] || viewHome)();
    }
    app.innerHTML = html + viewSheet();
    if (focus) {
      const el = document.getElementById(focus.id);
      if (el) {
        el.focus({ preventScroll: true });
        try { if (focus.s != null) el.setSelectionRange(focus.s, focus.e); } catch (e) { /* not a text input */ }
      }
    }
  }

  const navItem = (route, icon, label) => '<a href="#/' + route + '"' + (S.route === route ? ' aria-current="page"' : '') + '>' + svg(icon, 24) + label + '</a>';
  const nav = () => '<nav class="nav" aria-label="Main">' + navItem('home', 'home', 'Home') + navItem('history', 'list', 'History') + navItem('savings', 'coins', 'Savings') + navItem('budgets', 'pie', 'Budgets') + '</nav>';
  const fabs = () => '<div class="fab-row"><a class="fab" href="#/add">' + svg('plus', 22, 2.5) + 'Add</a><a class="fab scan" href="#/scan">' + svg('camera', 22) + 'Scan receipt</a></div>';
  const monthNav = (titleTag) => {
    const cur = monthOf(todayISO());
    return '<div class="month-nav"><button class="icon-btn" data-act="month" data-d="-1" aria-label="Previous month">' + svg('left') + '</button>' +
      '<' + titleTag + ' class="title">' + esc(monthLabel(S.month)) + '</' + titleTag + '>' +
      '<button class="icon-btn" data-act="month" data-d="1" aria-label="Next month"' + (S.month >= addMonths(cur, 3) ? ' disabled' : '') + '>' + svg('right') + '</button></div>';
  };
  const avatars = () => '<div class="avatars" aria-label="Shared by ' + esc(S.members.map((m) => m.display_name).join(' and ')) + '">' + membersOrdered().map((m) => avatar(m)).join('') + '</div>';

  // ---------- Login ----------
  function viewLogin() {
    const L = S.login;
    return '<main class="login">' +
      '<div class="brand"><img src="icons/icon-192.png" alt=""><div><h1>Our Budget</h1><p class="muted">Shared household budget</p></div></div>' +
      '<form data-submit="signIn" novalidate>' +
      '<label class="field" for="l-email">Email<input class="input" id="l-email" type="email" autocomplete="username" inputmode="email" data-model="login.email" value="' + esc(L.email) + '" required></label>' +
      '<label class="field" for="l-pw">Password<input class="input" id="l-pw" type="password" autocomplete="current-password" data-model="login.password" value="' + esc(L.password) + '" required></label>' +
      (L.err ? '<p class="err" role="alert">' + esc(L.err) + '</p>' : '') +
      '<button class="btn" type="submit"' + (L.busy ? ' disabled' : '') + '>' + (L.busy ? 'Signing in…' : 'Sign in') + '</button>' +
      '<button class="btn ghost" type="button" data-act="forgot">Forgot password?</button>' +
      '</form></main>';
  }
  function viewNotMember() {
    return '<main class="login"><h1>Not part of this household</h1><p class="muted">This account can sign in, but it is not one of the two household members, so it can’t see any data.</p><button class="btn secondary" data-act="signOut">Sign out</button></main>';
  }

  // ---------- Home ----------
  function dueBills() {
    const today = todayISO();
    return S.bills.filter((b) => b.active && b.next_due_on <= addDays(today, Math.max(0, b.remind_days_before ?? 3)))
      .sort((a, b) => (a.next_due_on < b.next_due_on ? -1 : 1));
  }
  function dueText(b) {
    const d = daysBetween(todayISO(), b.next_due_on);
    if (d < 0) return 'Overdue · ' + (-d === 1 ? '1 day' : -d + ' days');
    if (d === 0) return 'Due today';
    if (d === 1) return 'Due tomorrow';
    return 'Due in ' + d + ' days · ' + fmtDM(b.next_due_on);
  }

  // Monthly cost of the active bills in a category (e.g. rent, school).
  function fixedMonthly(catId) {
    const per = { weekly: 52 / 12, monthly: 1, bimonthly: 0.5, quarterly: 1 / 3, yearly: 1 / 12 };
    return S.bills.filter((b) => b.active && b.category_id === catId).reduce((s, b) => s + +b.amount * (per[b.frequency] || 1), 0);
  }
  const isFixedCat = (catId, budget) => budget > 0 && fixedMonthly(catId) >= budget * 0.8;

  function catCard(c, spent, budget) {
    const pct = budget > 0 ? (spent / budget) * 100 : spent > 0 ? 100 : 0;
    // Mostly-fixed categories (rent, school…) only warn when actually over.
    const warn = budget > 0 ? (isFixedCat(c.id, budget) ? spent > budget + 0.5 : pct >= 90) : spent > 0;
    const left = budget - spent;
    let val;
    if (!warn) val = '<strong>' + eur(spent, { smart: true }) + '</strong> of ' + eur(budget, { smart: true });
    else if (left >= 0) val = '<strong>' + eur(left, { smart: true }) + '</strong> left of ' + eur(budget, { smart: true });
    else val = '<strong>' + eur(-left, { smart: true }) + '</strong> over ' + (budget > 0 ? eur(budget, { smart: true }) : 'budget');
    return '<button class="cat-card' + (warn ? ' warn' : '') + '" data-act="catHistory" data-id="' + c.id + '" aria-label="' + esc(c.name + ': ' + eur(spent) + ' of ' + eur(budget) + ' — show entries') + '">' +
      '<div class="head"><span class="ico sm' + (warn ? ' warn' : '') + '">' + svg(catIcon(c), 17) + '</span><span class="name">' + esc(c.name) + '</span></div>' +
      '<div class="val">' + val + '</div>' +
      '<div class="track"><div style="width:' + Math.min(100, pct).toFixed(1) + '%"></div></div></button>';
  }

  function txRow(t) {
    const c = cat(t.category_id);
    const who = member(t.paid_by);
    const refund = t.kind === 'expense' && +t.amount < 0;
    const sub = [refund ? 'Refund' : '', c ? c.name : '', who ? who.display_name : '', methodLabel(t.payment_method)].filter(Boolean).join(' · ');
    const income = t.kind === 'income' || refund;
    return '<button class="row" data-act="openTx" data-id="' + t.id + '">' +
      '<span class="ico' + (income ? ' in' : '') + '">' + svg(catIcon(c), 18) + '</span>' +
      '<span class="grow"><span class="t">' + esc(t.merchant || (c ? c.name : 'Entry')) + (t.receipt_path ? ' <span class="muted">' + svg('clip', 13) + '</span>' : '') + (t.note ? ' <span class="muted">' + svg('note', 13) + '</span>' : '') + '</span><span class="s">' + esc(sub) + '</span></span>' +
      '<span class="amt' + (income ? ' in' : '') + '">' + (income ? eur(Math.abs(t.amount), { sign: true }) : eur(+t.amount)) + '</span></button>';
  }

  function viewHome() {
    const st = monthStats();
    const cur = monthOf(todayISO());
    const isCur = S.month === cur;
    const left = round2(st.budget - st.spent);
    const pct = st.budget > 0 ? Math.min(100, (st.spent / st.budget) * 100) : 0;
    let paceText = '';
    if (isCur) {
      const lastDay = addDays(addMonths(S.month, 1), -1);
      const daysLeft = daysBetween(todayISO(), lastDay) + 1;
      paceText = left > 0 ? '~' + eur(left / daysLeft, { dec: false }) + '/day · ' + daysLeft + (daysLeft === 1 ? ' day' : ' days') + ' left' : daysLeft + (daysLeft === 1 ? ' day' : ' days') + ' left';
    } else if (S.month < cur) paceText = 'Month closed';
    else paceText = 'Upcoming month';
    const saved = round2(st.income - st.spent);
    const target = +S.settings.monthly_savings_target || 0;
    const savePct = target > 0 ? Math.max(0, Math.min(100, (saved / target) * 100)) : 0;

    const hero = '<section class="hero' + (left < 0 ? ' over' : '') + '" aria-label="Month summary">' +
      '<div style="display:flex;flex-direction:column;gap:2px"><span class="label">' + (left < 0 ? 'Over budget by' : 'Left to spend') + '</span><span class="big">' + eur(Math.abs(left)) + '</span></div>' +
      '<div class="bar"><div style="width:' + pct.toFixed(1) + '%"></div></div>' +
      '<div class="meta"><span>' + eur(st.spent, { dec: false }) + ' of ' + eur(st.budget, { dec: false }) + ' budget</span><span>' + paceText + '</span></div>' +
      '<div class="boxes"><div class="box"><span>Money in</span><span>' + eur(st.income, { sign: true, dec: false }) + '</span></div><div class="box"><span>Money out</span><span>' + (st.spent ? '−' : '') + eur(st.spent, { dec: false }) + '</span></div></div>' +
      (target > 0 ? '<a class="saving" href="#/savings" style="text-decoration:none"><div class="line"><span>Savings this month</span><span><strong>' + eur(saved, { dec: false }) + '</strong> of ' + eur(target, { dec: false }) + '</span></div><div class="mini"><div style="width:' + savePct.toFixed(1) + '%"></div></div></a>' : '') +
      '</section>';

    const due = dueBills();
    const dueHtml = due.length ? '<section class="section" aria-labelledby="due-h"><div class="section-head"><h2 id="due-h">Bills due</h2><a class="link-btn" href="#/budgets?bills=1" style="text-decoration:none">All bills</a></div><div class="due">' +
      due.map((b) => {
        const over = b.next_due_on < todayISO();
        return '<div class="due-item' + (over ? ' overdue' : '') + '"><span class="ico sm' + (over ? ' warn' : '') + '">' + svg(catIcon(cat(b.category_id)), 17) + '</span><div class="grow"><span class="t">' + esc(b.name) + '</span><span class="s">' + (b.amount_is_estimate ? '~' : '') + eur(+b.amount) + ' · ' + dueText(b) + '</span></div><button class="pill-btn" data-act="payBill" data-id="' + b.id + '">' + svg('check', 16, 2.5) + 'Paid</button></div>';
      }).join('') + '</div></section>' : '';

    const cats = expenseCats().map((c) => {
      const budget = budgetFor(c.id, S.month);
      return { c, spent: st.spentBy[c.id] || 0, budget, fixed: isFixedCat(c.id, budget) };
    })
      // Stable order so cards don't move: day-to-day categories first, fixed costs after.
      .sort((a, b) => (a.fixed - b.fixed) || (a.c.sort_order - b.c.sort_order));
    const catHtml = '<section class="section" aria-labelledby="cat-h"><div class="section-head"><h2 id="cat-h">By category</h2>' +
      '<a class="link-btn" href="#/budgets" style="text-decoration:none">Edit budgets</a></div>' +
      '<div class="cat-grid">' + cats.map((x) => catCard(x.c, x.spent, x.budget)).join('') + '</div></section>';

    const recent = S.txs.slice(0, 6);
    const recentHtml = '<section class="section" aria-labelledby="recent-h"><div class="section-head"><h2 id="recent-h">Recent</h2><a class="link-btn" href="#/history" style="text-decoration:none">History</a></div>' +
      (recent.length ? '<div class="list">' + recent.map(txRow).join('') + '</div>' : '<div class="card empty">Nothing yet this month. Tap <strong>Add</strong> or <strong>Scan receipt</strong>.</div>') + '</section>';

    return '<main class="screen"><header class="topbar">' + monthNav('h1') + avatars() + '</header>' +
      '<div class="home-grid"><div>' + hero + dueHtml + catHtml + '</div><div>' + recentHtml + '</div></div></main>' + fabs() + nav();
  }

  // ---------- Add / edit ----------
  function viewAdd() {
    const f = S.form;
    const income = f.kind === 'income';
    const plus = income || f.refund;
    const amtDisplay = f.amountStr ? f.amountStr.replace('.', ',') : '0';
    const dateLabel = f.occurred_on === todayISO() ? 'Today' : f.occurred_on === addDays(todayISO(), -1) ? 'Yesterday' : fmtDate(f.occurred_on);
    const people = membersOrdered();
    // Day-to-day categories first; ones mostly covered by fixed bills last.
    const cats = income ? incomeCats() : expenseCats().slice().sort((x, y) => isFixedCat(x.id, budgetFor(x.id, S.month)) - isFixedCat(y.id, budgetFor(y.id, S.month)));
    if (f.category_id && !cats.find((c) => c.id === f.category_id) && cat(f.category_id)) cats.push(cat(f.category_id));
    const keys = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '.', '0', 'del'];
    return '<main class="add">' +
      '<header class="add-head">' +
      '<button class="icon-btn filled" data-act="closeForm" aria-label="Close">' + svg('close', 20, 2.2) + '</button>' +
      '<div class="seg round" role="group" aria-label="Type"><button data-act="kind" data-k="expense" aria-pressed="' + !income + '">Expense</button><button data-act="kind" data-k="income" aria-pressed="' + income + '">Income</button></div>' +
      '<label class="date-pill">' + svg('cal', 16) + '<span>' + esc(dateLabel) + '</span><input id="f-date" type="date" aria-label="Date" value="' + f.occurred_on + '" max="' + addDays(todayISO(), 400) + '" data-model="form.occurred_on" data-rerender></label>' +
      '</header>' +
      '<div class="amount' + (!f.amountStr ? ' empty-amt' : plus ? ' in' : '') + '" aria-live="polite">' + (plus && f.amountStr ? '+' : '') + esc(amtDisplay) + ' €</div>' +
      (S.route === 'edit'
        ? (f.receipt_path ? '<button class="scan-link" data-act="viewReceipt" data-path="' + esc(f.receipt_path) + '">' + svg('clip', 18) + 'View receipt photo</button>' : '')
        : !income ? '<a class="scan-link" href="#/scan">' + svg('camera', 20) + 'Take a photo of the receipt</a>' : '<div class="hint">Money coming in: salary, refunds, gifts…</div>') +
      '<div class="seg" role="group" aria-label="' + (income ? 'Who received it' : 'Who paid') + '">' +
      people.map((m) => '<button data-act="who" data-id="' + m.user_id + '" aria-pressed="' + (f.paid_by === m.user_id) + '">' + esc(m.display_name) + (income ? ' received' : ' paid') + '</button>').join('') + '</div>' +
      '<input class="input" id="f-merchant" type="text" autocomplete="off" maxlength="80" placeholder="' + (income ? 'e.g. October salary' : 'Shop or description, e.g. Mercadona') + '" aria-label="' + (income ? 'Description' : 'Shop or description') + '" value="' + esc(f.merchant) + '" data-model="form.merchant" data-after="merchantTyped">' +
      '<div class="cat-chips" role="group" aria-label="Category">' + cats.map((c) => '<button class="chip' + (income ? ' income' : '') + '" data-act="pickCat" data-id="' + c.id + '" aria-pressed="' + (f.category_id === c.id) + '">' + svg(catIcon(c), 16) + esc(c.name) + '</button>').join('') + '</div>' +
      '<div class="chips" role="group" aria-label="Payment method">' + METHODS.map((m) => '<button class="chip sm" data-act="method" data-m="' + m[0] + '" aria-pressed="' + (f.payment_method === m[0]) + '">' + svg(m[2], 15) + m[1] + '</button>').join('') +
      (income ? '' : '<button class="chip sm" data-act="toggleRefund" aria-pressed="' + !!f.refund + '">' + svg('repeat', 15) + 'Refund</button>') +
      (f.showNote ? '' : '<button class="chip sm" data-act="showNote">' + svg('note', 15) + 'Add note</button>') + '</div>' +
      (f.showNote ? '<input class="input" id="f-note" type="text" maxlength="500" placeholder="Note (optional)" aria-label="Note" value="' + esc(f.note) + '" data-model="form.note">' : '') +
      '<div class="keypad">' + keys.map((k) => '<button data-act="key" data-k="' + k + '" aria-label="' + (k === 'del' ? 'Delete' : k === '.' ? 'Decimal comma' : k) + '">' + (k === 'del' ? '⌫' : k === '.' ? ',' : k) + '</button>').join('') + '</div>' +
      '<button class="btn' + (income ? ' income' : '') + '" data-act="saveTx"' + (S.busy ? ' disabled' : '') + '>' + svg('check', 20, 2.5) + (S.busy ? 'Saving…' : f.id ? 'Save changes' : income ? 'Save income' : f.refund ? 'Save refund' : 'Save expense') + '</button>' +
      (f.id ? '<button class="btn danger" data-act="deleteTx">' + svg('trash', 18) + 'Delete</button>' : '') +
      '</main>';
  }

  // ---------- Scan ----------
  function viewScan() {
    const sc = S.scan;
    const inputs = '<input id="file-cam" class="sr-only" type="file" accept="image/*" capture="environment" data-change="fileChosen" tabindex="-1">' +
      '<input id="file-any" class="sr-only" type="file" accept="image/*,application/pdf" data-change="fileChosen" tabindex="-1">';
    if (sc.step === 'pick') {
      return '<main class="scan-screen"><header class="topbar"><a class="icon-btn" href="#/home" aria-label="Close">' + svg('close', 20, 2.2) + '</a><h1>Scan receipt</h1><span style="width:44px"></span></header>' +
        '<div class="scan-frame"><div class="frame"><i></i><i></i><i></i><i></i><div class="paper" aria-hidden="true"><b class="h" style="width:60%;align-self:center"></b><b style="width:40%;align-self:center"></b><hr><b style="width:80%"></b><b style="width:70%"></b><b style="width:85%"></b><b style="width:60%"></b><b style="width:75%"></b><hr><b class="h" style="width:50%;align-self:flex-end"></b></div></div>' +
        '<p style="font-size:15px;font-weight:500;color:#E4E9E6;text-align:center">Lay the receipt flat in good light,<br>then take the photo</p></div>' +
        '<div class="shutter-row">' +
        '<label class="side-btn" for="file-any" style="justify-self:start"><span>' + svg('photos') + '</span>Photo / PDF</label>' +
        '<label class="shutter" for="file-cam" aria-label="Take photo"><span></span></label>' +
        '<a class="side-btn" href="#/add" style="justify-self:end"><span>' + svg('keyboard') + '</span>Type it</a>' +
        '</div>' + inputs + '</main>';
    }
    if (sc.step === 'reading') {
      return '<main class="scan-screen"><header class="topbar"><a class="icon-btn" href="#/home" aria-label="Cancel">' + svg('close', 20, 2.2) + '</a><h1>Reading receipt…</h1><span style="width:44px"></span></header>' +
        '<div class="reading">' + (sc.previewUrl ? '<img src="' + sc.previewUrl + '" alt="">' : '') +
        '<div class="progress" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="' + Math.round((sc.progress || 0) * 100) + '"><div style="width:' + Math.round((sc.progress || 0) * 100) + '%"></div></div>' +
        '<p style="color:#E4E9E6;font-size:14px;max-width:280px">' + esc(sc.status || 'Reading on your phone. Nothing is sent anywhere until you save.') + '</p></div></main>';
    }
    // review
    const f = S.form;
    const p = sc.parsed || {};
    const ok = p.amount && p.amountSure && p.date;
    const people = membersOrdered();
    const missing = [];
    if (!p.amount) missing.push('total');
    if (!p.date) missing.push('date');
    if (!p.merchant) missing.push('shop');
    let title = ok ? 'Receipt read' : 'Please check the details';
    let sub = ok ? 'Check the details, then save. The photo is kept for 12 months.' : missing.length ? 'Couldn’t find the ' + missing.join(', ') + '. Fill in what’s missing.' : 'Check the total especially, then save.';
    if (sc.failed) { title = 'Couldn’t read this receipt'; sub = 'Type the details in. The photo will still be saved.'; }
    return '<main class="review"><header class="topbar"><button class="icon-btn filled" data-act="retake" aria-label="Retake photo">' + svg('left') + '</button><h1>Check the receipt</h1><span style="width:44px"></span></header>' +
      '<button class="read-note' + (ok ? '' : ' warn') + '" data-act="viewPreview">' + (sc.previewUrl ? '<img src="' + sc.previewUrl + '" alt="Receipt photo">' : '') + '<span style="display:flex;flex-direction:column;gap:2px"><span class="t">' + title + '</span><span class="s">' + sub + '</span></span></button>' +
      '<label class="field" for="s-merchant">Store<input class="input" id="s-merchant" type="text" maxlength="80" value="' + esc(f.merchant) + '" data-model="form.merchant" data-after="merchantTyped"></label>' +
      '<div class="grid-2"><label class="field" for="s-total">Total (€)<input class="input big" id="s-total" type="text" inputmode="decimal" value="' + esc(f.amountText || '') + '" placeholder="0,00" data-model="form.amountText"></label>' +
      '<label class="field" for="s-date">Date' + dateIn('s-date', 'form.occurred_on', f.occurred_on) + '</label></div>' +
      '<label class="field" for="s-cat">Category<select class="select" id="s-cat" data-change="selectCat">' + (f.category_id ? '' : '<option value="">Choose a category…</option>') +
      expenseCats().map((c) => '<option value="' + c.id + '"' + (f.category_id === c.id ? ' selected' : '') + '>' + esc(c.name) + '</option>').join('') + '</select></label>' +
      '<div class="seg" role="group" aria-label="Who paid">' + people.map((m) => '<button data-act="who" data-id="' + m.user_id + '" aria-pressed="' + (f.paid_by === m.user_id) + '">' + esc(m.display_name) + ' paid</button>').join('') + '</div>' +
      '<div class="chips" role="group" aria-label="Payment method">' + METHODS.map((m) => '<button class="chip sm" data-act="method" data-m="' + m[0] + '" aria-pressed="' + (f.payment_method === m[0]) + '">' + svg(m[2], 15) + m[1] + '</button>').join('') + '</div>' +
      '<input class="input" id="s-note" type="text" maxlength="500" placeholder="Note (optional)" aria-label="Note" value="' + esc(f.note) + '" data-model="form.note">' +
      '<div style="margin-top:auto;display:flex;flex-direction:column;gap:6px"><button class="btn" data-act="saveTx"' + (S.busy ? ' disabled' : '') + '>' + svg('check', 20, 2.5) + (S.busy ? 'Saving…' : 'Save expense') + '</button>' +
      '<button class="btn ghost" data-act="retake">Retake photo</button></div></main>';
  }

  // ---------- History ----------
  function viewHistory() {
    const H = S.hist;
    const people = membersOrdered();
    const byMonth = totalsByMonth();
    // 12-month trend ending at the selected month
    const months = [];
    for (let i = 11; i >= 0; i--) months.push(addMonths(S.month, -i));
    const max = Math.max(1, ...months.map((m) => Math.max((byMonth[m] || {}).income || 0, (byMonth[m] || {}).expense || 0)));
    const trend = '<section class="card pad section" aria-labelledby="trend-h"><div class="section-head"><h2 id="trend-h">Last 12 months</h2><div class="legend"><span>Spent</span><span class="inc">Income</span></div></div>' +
      '<div class="trend">' + months.map((m) => {
        const o = byMonth[m] || { income: 0, expense: 0 };
        return '<button data-act="setMonth" data-m="' + m + '" aria-pressed="' + (m === S.month) + '" aria-label="' + monthLabel(m, true) + ': spent ' + eur(o.expense) + ', income ' + eur(o.income) + '"><span class="bars"><i class="out" style="height:' + ((o.expense / max) * 100).toFixed(1) + '%"></i><i class="inc" style="height:' + ((o.income / max) * 100).toFixed(1) + '%"></i></span>' + MON[+m.slice(5, 7) - 1].slice(0, 1) + '</button>';
      }).join('') + '</div></section>';

    const cur = byMonth[S.month] || { byCat: {}, expense: 0 };
    const prev = byMonth[addMonths(S.month, -1)] || { byCat: {} };
    const rows = Object.keys(cur.byCat).map((id) => ({ c: cat(id), v: cur.byCat[id], p: prev.byCat[id] || 0 })).filter((x) => x.c).sort((a, b) => b.v - a.v);
    const maxC = Math.max(1, ...rows.map((r) => r.v));
    const breakdown = rows.length ? '<section class="card pad section breakdown" aria-labelledby="bd-h"><div class="section-head"><h2 id="bd-h">Spending by category</h2><span class="muted num" style="font-weight:700">' + eur(cur.expense) + '</span></div><div>' +
      rows.map((r) => {
        const delta = r.p > 0 ? Math.round(((r.v - r.p) / r.p) * 100) : null;
        return '<button class="row" data-act="catHistory" data-id="' + r.c.id + '" style="padding:8px 0"><span class="dot" style="background:' + esc(r.c.color) + '"></span><span class="grow"><span class="t" style="font-size:14px">' + esc(r.c.name) + '</span><span class="hbar"><span style="display:block;height:100%;width:' + ((r.v / maxC) * 100).toFixed(1) + '%;background:' + esc(r.c.color) + '"></span></span></span>' +
          '<span style="display:flex;flex-direction:column;align-items:flex-end"><span class="amt" style="font-size:14px">' + eur(r.v) + '</span>' + (delta !== null && delta !== 0 ? '<span class="delta ' + (delta > 0 ? 'up' : 'down') + '">' + (delta > 0 ? '▲ ' : '▼ ') + Math.abs(delta) + '% vs ' + MON[+addMonths(S.month, -1).slice(5, 7) - 1] + '</span>' : '') + '</span></button>';
      }).join('') + '</div></section>' : '';

    const q = Receipt.norm(H.q).trim();
    const list = S.txs.filter((t) => (H.who === 'all' || t.paid_by === H.who) && (!H.cat || t.category_id === H.cat) &&
      (!q || Receipt.norm([t.merchant, t.note, (cat(t.category_id) || {}).name].join(' ')).includes(q)));
    const days = [];
    for (const t of list) {
      let d = days[days.length - 1];
      if (!d || d.date !== t.occurred_on) { d = { date: t.occurred_on, items: [], total: 0 }; days.push(d); }
      d.items.push(t);
      d.total += t.kind === 'income' ? +t.amount : -t.amount;
    }
    const catFilter = H.cat && cat(H.cat) ? '<div class="chips"><button class="chip" aria-pressed="true" data-act="clearCatFilter">' + esc(cat(H.cat).name) + ' ' + svg('close', 14, 2.5) + '</button></div>' : '';
    const listHtml = days.length ? days.map((d) => '<section class="section" style="gap:6px"><div class="day-head"><span>' + esc(dayLabel(d.date)) + '</span><span class="num">' + eur(round2(d.total), { sign: true }) + '</span></div><div class="list">' + d.items.map(txRow).join('') + '</div></section>').join('')
      : '<div class="card empty">' + (S.txs.length ? 'Nothing matches.' : 'No entries in ' + esc(monthLabel(S.month)) + '.') + '</div>';

    return '<main class="screen"><header class="topbar"><h1 class="title">History</h1><button class="icon-btn filled" data-act="toggleSearch" aria-label="Search entries" aria-expanded="' + H.search + '">' + svg('search') + '</button></header>' +
      '<div style="display:flex;justify-content:center;margin-top:-10px">' + monthNav('h2') + '</div>' +
      (H.search ? '<input class="input" id="h-q" type="search" placeholder="Search this month" value="' + esc(H.q) + '" data-model="hist.q" data-rerender aria-label="Search this month">' : '') +
      '<div class="filter-chips" role="group" aria-label="Filter by person">' +
      [['all', 'Both of us']].concat(people.map((m) => [m.user_id, m.display_name])).map((x) => '<button class="chip" data-act="histWho" data-id="' + x[0] + '" aria-pressed="' + (H.who === x[0]) + '">' + esc(x[1]) + '</button>').join('') + '</div>' +
      catFilter +
      '<div class="wide-2"><div>' + trend + breakdown + '</div><div>' + listHtml + '</div></div></main>' + fabs() + nav();
  }

  // ---------- Savings ----------
  function viewSavings() {
    const cur = monthOf(todayISO());
    const byMonth = totalsByMonth();
    const target = +S.settings.monthly_savings_target || 0;
    const thisM = byMonth[cur] || { income: 0, expense: 0 };
    const savedNow = round2(thisM.income - thisM.expense);
    const goals = S.goals.filter((g) => !g.archived);
    const savedIn = (g) => round2(S.contribs.filter((c) => c.goal_id === g.id).reduce((s, c) => s + +c.amount, 0));
    const totalSaved = round2(goals.reduce((s, g) => s + savedIn(g), 0));

    const monthCard = '<section class="total-card" aria-label="This month"><div class="section-head"><span class="muted" style="font-weight:600">Saved this month (income − spending)</span><button class="link-btn" data-act="editTarget">Target</button></div>' +
      '<span class="big">' + eur(savedNow, { dec: false }) + '</span>' +
      (target > 0 ? '<div class="track" style="height:8px"><div style="width:' + Math.max(0, Math.min(100, (savedNow / target) * 100)).toFixed(1) + '%"></div></div><span class="muted" style="font-size:13px">Target ' + eur(target, { dec: false }) + ' a month · ' + (savedNow >= target ? 'reached' : eur(target - savedNow, { dec: false }) + ' to go') + '</span>' : '<span class="muted" style="font-size:13px">No monthly target set.</span>') + '</section>';

    const totalCard = '<section class="total-card" aria-label="Total saved"><span class="muted" style="font-weight:600">Saved together in goals</span><span class="big">' + eur(totalSaved, { dec: false }) + '</span><span class="muted" style="font-size:13px">' + (goals.length ? 'In ' + goals.length + (goals.length === 1 ? ' goal' : ' goals') : 'No goals yet') + '</span></section>';

    const past = [];
    for (let i = 1; i <= 3; i++) {
      const m = addMonths(cur, -i);
      const o = byMonth[m];
      if (!o || !o.expense) continue;
      const b = round2(expenseCats().reduce((s, c) => s + budgetFor(c.id, m), 0));
      past.push({ m, spent: round2(o.expense), budget: b, diff: round2(b - o.expense) });
    }
    const pot = past.find((p) => p.diff > 0);
    const leftover = past.length ? '<section class="section" aria-labelledby="left-h"><h2 id="left-h" style="font-size:17px">Left over from past months</h2><div class="list">' +
      past.map((p) => '<div class="row"><span class="grow"><span class="t">' + monthLabel(p.m) + '</span><span class="s">Spent ' + eur(p.spent, { dec: false }) + ' of ' + eur(p.budget, { dec: false }) + '</span></span><span class="' + (p.diff >= 0 ? 'under' : 'over') + ' num">' + (p.diff >= 0 ? '+' + eur(p.diff, { dec: false }) + ' under' : eur(-p.diff, { dec: false }) + ' over') + '</span></div>').join('') + '</div>' +
      (pot ? '<button class="btn secondary" data-act="putLeftover" data-v="' + pot.diff + '">Put ' + eur(pot.diff, { dec: false }) + ' from ' + MONTHS[+pot.m.slice(5, 7) - 1] + ' into a goal ' + svg('right', 18) + '</button>' : '') + '</section>' : '';

    const months = (n) => monthLabel(addMonths(cur, n));
    const goalsHtml = '<section class="section" aria-labelledby="goals-h"><div class="section-head"><h2 id="goals-h">Goals</h2></div>' +
      (goals.length ? goals.map((g) => {
        const saved = savedIn(g);
        const target2 = +g.target_amount;
        const left = Math.max(0, target2 - saved);
        let plan = '';
        if (left <= 0) plan = 'Goal reached';
        else if (+g.monthly_plan > 0) plan = 'Plan: ' + eur(+g.monthly_plan, { dec: false }) + ' a month · ready by ' + months(Math.ceil(left / +g.monthly_plan));
        else if (g.target_date) plan = 'Target date ' + fmtDate(g.target_date);
        return '<button class="goal" data-act="openGoal" data-id="' + g.id + '"><span class="line"><span class="name">' + esc(g.name) + '</span><span class="muted" style="font-size:13px"><strong style="color:var(--ink);font-size:15px">' + eur(saved, { dec: false }) + '</strong> of ' + eur(target2, { dec: false }) + '</span></span>' +
          '<span class="track" style="display:block"><span style="display:block;height:100%;border-radius:3px;background:var(--lime);width:' + Math.min(100, (saved / target2) * 100).toFixed(1) + '%"></span></span>' + (plan ? '<span class="plan">' + plan + '</span>' : '') + '</button>';
      }).join('') : '<div class="card empty">No goals yet. Create one when you’re ready, e.g. a summer holiday or an emergency fund.</div>') + '</section>';

    return '<main class="screen"><header class="topbar"><h1 class="title">Savings</h1><button class="pill-btn" data-act="newGoal">' + svg('plus', 16, 2.5) + 'New goal</button></header>' +
      '<div class="wide-2"><div>' + monthCard + totalCard + leftover + '</div><div>' + goalsHtml + '</div></div></main>' + nav();
  }

  // ---------- Budgets & settings ----------
  function viewBudgets() {
    const cats = expenseCats();
    const total = round2(cats.reduce((s, c) => s + budgetFor(c.id, S.month), 0));
    const limits = '<section class="section" aria-labelledby="lim-h"><div class="section-head"><h2 id="lim-h">Category limits</h2></div>' +
      '<div class="seg round" role="group" aria-label="Budget changes apply to"><button data-act="scope" data-s="forward" aria-pressed="' + (S.budgetScope === 'forward') + '">From ' + MON[+S.month.slice(5, 7) - 1] + ' on</button><button data-act="scope" data-s="once" aria-pressed="' + (S.budgetScope === 'once') + '">Only ' + MON[+S.month.slice(5, 7) - 1] + '</button></div>' +
      '<div class="list">' + cats.map((c) => '<div class="limit-row"><button class="name" data-act="openCat" data-id="' + c.id + '"><span class="ico sm">' + svg(catIcon(c), 16) + '</span><span>' + esc(c.name) + '</span></button>' +
        '<label class="money-in"><input id="b-' + c.id + '" type="text" inputmode="decimal" value="' + numText(budgetFor(c.id, S.month)) + '" data-change="setBudget" data-id="' + c.id + '" aria-label="' + esc(c.name) + ' budget"><span aria-hidden="true">€</span></label></div>').join('') + '</div>' +
      '<button class="dashed" data-act="newCat" data-kind="expense">' + svg('plus', 18, 2.2) + 'Add category</button></section>';

    const bills = S.bills.slice().sort((a, b) => (b.active - a.active) || (a.next_due_on < b.next_due_on ? -1 : 1));
    const billsHtml = '<section class="section" id="bills" aria-labelledby="bills-h"><div class="section-head"><h2 id="bills-h">Recurring bills</h2><span class="muted num" style="font-size:13px">' + eur(round2(bills.filter((b) => b.active && b.frequency === 'monthly').reduce((s, b) => s + +b.amount, 0)), { dec: false }) + ' / month</span></div>' +
      '<div class="list">' + (bills.length ? bills.map((b) => '<button class="row" data-act="openBill" data-id="' + b.id + '"><span class="ico sm">' + svg(catIcon(cat(b.category_id)), 16) + '</span><span class="grow"><span class="t">' + esc(b.name) + (!b.active ? '<span class="tag paused">' + (b.ends_on && b.next_due_on > b.ends_on ? 'Ended' : 'Paused') + '</span>' : '') + '</span><span class="s">' + freqLabel(b.frequency) + (b.active ? ' · next ' + fmtDate(b.next_due_on) : '') + (b.ends_on ? ' · until ' + fmtDate(b.ends_on) : '') + '</span></span><span class="amt">' + (b.amount_is_estimate ? '~' : '') + eur(+b.amount, { smart: true }) + '</span></button>').join('') : '<div class="empty">No bills yet.</div>') + '</div>' +
      '<button class="dashed" data-act="newBill">' + svg('plus', 18, 2.2) + 'Add bill</button></section>';

    const incomeHtml = '<section class="section" aria-labelledby="inc-h"><h2 id="inc-h" style="font-size:17px">Income categories</h2><div class="chips wrap">' +
      incomeCats().map((c) => '<button class="chip" data-act="openCat" data-id="' + c.id + '">' + svg(catIcon(c), 16) + esc(c.name) + '</button>').join('') +
      '<button class="chip" data-act="newCat" data-kind="income">' + svg('plus', 16) + 'Add</button></div></section>';

    const exportHtml = '<section class="section" aria-labelledby="exp-h"><h2 id="exp-h" style="font-size:17px">Export</h2><div class="card pad" style="display:flex;flex-direction:column;gap:12px">' +
      '<div class="grid-2"><label class="field" for="e-from">From' + dateIn('e-from', 'exp.from', S.exp.from) + '</label><label class="field" for="e-to">To' + dateIn('e-to', 'exp.to', S.exp.to) + '</label></div>' +
      '<div class="grid-2"><button class="btn secondary small" data-act="export" data-f="csv">' + svg('download', 18) + 'CSV</button><button class="btn secondary small" data-act="export" data-f="xlsx">' + svg('download', 18) + 'Excel</button></div></div></section>';

    const others = membersOrdered().slice(1).map((m) => m.display_name).join(', ');
    const account = '<section class="section" aria-label="Household"><div class="card pad" style="display:flex;align-items:center;gap:12px">' + avatars() +
      '<div style="flex:1;display:flex;flex-direction:column"><span style="font-weight:600">' + esc(S.members.map((m) => m.display_name).join(' & ')) + '</span><span class="muted" style="font-size:13px">Signed in as ' + esc(S.me.display_name) + ' · syncs on every device</span></div></div>' +
      '<div class="list">' +
      '<button class="row" data-act="changePassword"><span class="ico sm">' + svg('key', 16) + '</span><span class="grow"><span class="t">Change password</span></span>' + svg('right', 18) + '</button>' +
      (api().demo ? '<button class="row" data-act="demoSwitch"><span class="ico sm">' + svg('user', 16) + '</span><span class="grow"><span class="t">Demo: act as ' + esc(others) + '</span></span></button>' : '') +
      '<button class="row" data-act="signOut"><span class="ico sm warn">' + svg('logout', 16) + '</span><span class="grow"><span class="t">Sign out</span></span></button>' +
      '</div></section>';

    return '<main class="screen"><header class="topbar"><h1 class="title">Budgets</h1>' + monthNav('h2').replace('class="title"', 'class="title" style="font-size:18px"') + '</header>' +
      '<div class="wide-2"><div><section class="total-row" aria-label="Monthly total"><div style="display:flex;flex-direction:column;gap:2px"><span class="muted" style="font-weight:600;font-size:14px">Monthly budget</span><span class="big">' + eur(total, { dec: false }) + '</span></div><span class="muted" style="font-size:13px;text-align:right;max-width:130px">Sum of the categories below</span></section>' + limits + '</div>' +
      '<div>' + billsHtml + incomeHtml + exportHtml + account + '</div></div></main>' + nav();
  }

  // ---------- Sheets ----------
  function sheetWrap(title, body) {
    return '<div class="overlay"><div class="sheet" role="dialog" aria-modal="true" aria-label="' + esc(title) + '"><div class="grab" aria-hidden="true"></div><div class="sheet-head"><h2>' + esc(title) + '</h2><button class="icon-btn filled" data-act="closeSheet" aria-label="Close" style="background:var(--card-2)">' + svg('close', 18, 2.2) + '</button></div>' + body + '</div></div>';
  }
  const field = (label, id, inner) => '<label class="field" for="' + id + '">' + label + inner + '</label>';
  const textIn = (id, model, value, attrs) => '<input class="input" id="' + id + '" type="text" value="' + esc(value ?? '') + '" data-model="' + model + '" ' + (attrs || '') + '>';
  const moneyIn = (id, model, value) => '<input class="input" id="' + id + '" type="text" inputmode="decimal" value="' + esc(value ?? '') + '" data-model="' + model + '" placeholder="0,00">';
  // Shows the date as dd/mm/yyyy whatever the phone's region; tapping opens the native picker.
  const dateIn = (id, model, value, empty) => '<span class="input date-field"><span>' + (value ? fmtDate(value) : '<span class="muted">' + (empty || 'Not set') + '</span>') + '</span>' + svg('cal', 18) +
    '<input id="' + id + '" type="date" value="' + esc(value || '') + '" data-model="' + model + '" data-rerender></span>';
  const selectIn = (id, model, value, opts) => '<select class="select" id="' + id + '" data-model="' + model + '">' + opts.map((o) => '<option value="' + esc(o[0]) + '"' + (String(value) === String(o[0]) ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') + '</select>';
  const methodChips = (model, value) => '<div class="chips wrap" role="group" aria-label="Payment method">' + METHODS.map((m) => '<button class="chip sm" data-act="setSheet" data-k="' + model + '" data-v="' + m[0] + '" aria-pressed="' + (value === m[0]) + '">' + svg(m[2], 15) + m[1] + '</button>').join('') + '</div>';

  function viewSheet() {
    const sh = S.sheet;
    if (!sh) return '';
    const busy = S.busy ? ' disabled' : '';
    switch (sh.type) {
      case 'pay': {
        const b = S.bills.find((x) => x.id === sh.billId);
        if (!b) return '';
        return sheetWrap('Mark as paid', '<p class="muted">' + esc(b.name) + ' · due ' + fmtDate(b.next_due_on) + (b.amount_is_estimate ? ' · the amount varies, enter the real one' : '') + '</p>' +
          '<div class="grid-2">' + field('Amount (€)', 'p-amt', moneyIn('p-amt', 'sheet.amount', sh.amount)) + field('Paid on', 'p-date', dateIn('p-date', 'sheet.date', sh.date)) + '</div>' +
          methodChips('method', sh.method) +
          '<button class="btn" data-act="confirmPay"' + busy + '>' + svg('check', 20, 2.5) + 'Record payment</button>' +
          '<p class="muted" style="font-size:13px">Adds the expense and moves the next due date to ' + fmtDate(nextDue(b)) + '.</p>');
      }
      case 'bill': {
        const f = sh.f;
        return sheetWrap(f.id ? 'Edit bill' : 'New bill',
          field('Name', 'bl-name', textIn('bl-name', 'sheet.f.name', f.name, 'maxlength="80" placeholder="e.g. Alquiler"')) +
          '<div class="grid-2">' + field('Amount (€)', 'bl-amt', moneyIn('bl-amt', 'sheet.f.amount', f.amount)) + field('How often', 'bl-freq', selectIn('bl-freq', 'sheet.f.frequency', f.frequency, FREQS)) + '</div>' +
          '<label class="check"><input type="checkbox" id="bl-est" data-model="sheet.f.amount_is_estimate"' + (f.amount_is_estimate ? ' checked' : '') + '>Amount varies (ask for the real amount when paying)</label>' +
          field('Category', 'bl-cat', selectIn('bl-cat', 'sheet.f.category_id', f.category_id, expenseCats().map((c) => [c.id, c.name]))) +
          '<div class="grid-2">' + field('Next due', 'bl-due', dateIn('bl-due', 'sheet.f.next_due_on', f.next_due_on)) + field('Ends (optional)', 'bl-end', dateIn('bl-end', 'sheet.f.ends_on', f.ends_on, 'No end')) + '</div>' +
          (f.ends_on ? '<button class="link-btn" style="align-self:flex-end;margin-top:-10px" data-act="setSheet" data-k="f.ends_on" data-v="">Remove end date</button>' : '') +
          field('Remind me (days before)', 'bl-rem', '<input class="input" id="bl-rem" type="number" min="0" max="60" inputmode="numeric" value="' + esc(f.remind_days_before) + '" data-model="sheet.f.remind_days_before">') +
          methodChips('f.payment_method', f.payment_method) +
          (f.id ? '<label class="check"><input type="checkbox" id="bl-act" data-model="sheet.f.active"' + (f.active ? ' checked' : '') + '>Active (untick to pause, e.g. school in summer)</label>' : '') +
          '<button class="btn" data-act="saveBill"' + busy + '>' + svg('check', 20, 2.5) + 'Save bill</button>' +
          (f.id ? '<button class="btn danger" data-act="deleteBill">' + svg('trash', 18) + 'Delete bill</button>' : ''));
      }
      case 'cat': {
        const f = sh.f;
        return sheetWrap(f.id ? 'Edit category' : f.kind === 'income' ? 'New income category' : 'New category',
          field('Name', 'c-name', textIn('c-name', 'sheet.f.name', f.name, 'maxlength="40"')) +
          (!f.id && f.kind === 'expense' ? field('Monthly budget (€)', 'c-bud', moneyIn('c-bud', 'sheet.f.budget', f.budget)) : '') +
          field('Chart colour', 'c-col', '<input class="input" id="c-col" type="color" value="' + esc(f.color) + '" data-model="sheet.f.color" style="padding:4px">') +
          '<button class="btn" data-act="saveCat"' + busy + '>' + svg('check', 20, 2.5) + 'Save</button>' +
          (f.id ? '<button class="btn danger" data-act="archiveCat">' + svg('trash', 18) + 'Hide category</button><p class="muted" style="font-size:13px">Hiding keeps its past entries in History and exports.</p>' : ''));
      }
      case 'goal': {
        const g = S.goals.find((x) => x.id === sh.goalId);
        if (!g) return '';
        const list = S.contribs.filter((c) => c.goal_id === g.id);
        const saved = round2(list.reduce((s, c) => s + +c.amount, 0));
        return sheetWrap(g.name, '<div class="total-card" style="background:var(--card-2)"><span class="big">' + eur(saved, { dec: false }) + '</span><span class="muted">of ' + eur(+g.target_amount, { dec: false }) + (g.target_date ? ' · by ' + fmtDate(g.target_date) : '') + '</span></div>' +
          '<div class="grid-2"><button class="btn small" data-act="contribute" data-id="' + g.id + '" data-dir="in">' + svg('plus', 18, 2.5) + 'Add money</button><button class="btn secondary small" data-act="contribute" data-id="' + g.id + '" data-dir="out">Take out</button></div>' +
          (list.length ? '<div class="list" style="background:var(--card-2)">' + list.slice(0, 12).map((c) => '<div class="row"><span class="grow"><span class="t">' + fmtDate(c.occurred_on) + '</span><span class="s">' + esc((member(c.created_by) || {}).display_name || '') + '</span></span><span class="amt' + (c.amount > 0 ? ' in' : '') + '">' + eur(+c.amount, { sign: true }) + '</span><button class="icon-btn" data-act="deleteContribution" data-id="' + c.id + '" aria-label="Delete this entry">' + svg('trash', 16) + '</button></div>').join('') + '</div>' : '<p class="muted">Nothing saved in this goal yet.</p>') +
          '<button class="btn ghost" data-act="editGoal" data-id="' + g.id + '">Edit goal</button>');
      }
      case 'goalEdit': {
        const f = sh.f;
        return sheetWrap(f.id ? 'Edit goal' : 'New goal',
          field('Name', 'g-name', textIn('g-name', 'sheet.f.name', f.name, 'maxlength="60" placeholder="e.g. Vacaciones verano"')) +
          '<div class="grid-2">' + field('Target (€)', 'g-target', moneyIn('g-target', 'sheet.f.target_amount', f.target_amount)) + field('Plan per month (€)', 'g-plan', moneyIn('g-plan', 'sheet.f.monthly_plan', f.monthly_plan)) + '</div>' +
          field('Target date (optional)', 'g-date', dateIn('g-date', 'sheet.f.target_date', f.target_date)) +
          '<button class="btn" data-act="saveGoal"' + busy + '>' + svg('check', 20, 2.5) + 'Save goal</button>' +
          (f.id ? '<button class="btn danger" data-act="archiveGoal">' + svg('trash', 18) + 'Archive goal</button>' : ''));
      }
      case 'contribute': {
        const goals = S.goals.filter((g) => !g.archived);
        return sheetWrap(sh.dir === 'out' ? 'Take money out' : 'Add to a goal',
          field('Goal', 'k-goal', selectIn('k-goal', 'sheet.goalId', sh.goalId, goals.map((g) => [g.id, g.name]))) +
          '<div class="grid-2">' + field('Amount (€)', 'k-amt', moneyIn('k-amt', 'sheet.amount', sh.amount)) + field('Date', 'k-date', dateIn('k-date', 'sheet.date', sh.date)) + '</div>' +
          '<button class="btn" data-act="saveContribution"' + busy + '>' + svg('check', 20, 2.5) + (sh.dir === 'out' ? 'Take out' : 'Add money') + '</button>');
      }
      case 'target':
        return sheetWrap('Monthly savings target', field('Amount per month (€)', 't-amt', moneyIn('t-amt', 'sheet.amount', sh.amount)) + '<p class="muted" style="font-size:13px">Compared with income minus spending each month.</p><button class="btn" data-act="saveTarget"' + busy + '>Save</button>');
      case 'password':
        return sheetWrap(sh.recovery ? 'Choose a new password' : 'Change password',
          field('New password', 'pw1', '<input class="input" id="pw1" type="password" autocomplete="new-password" data-model="sheet.pw1" value="' + esc(sh.pw1 || '') + '">') +
          field('Repeat it', 'pw2', '<input class="input" id="pw2" type="password" autocomplete="new-password" data-model="sheet.pw2" value="' + esc(sh.pw2 || '') + '">') +
          '<p class="muted" style="font-size:13px">At least 10 characters.</p><button class="btn" data-act="savePassword"' + busy + '>Save password</button>');
      case 'image':
        return sheetWrap('Receipt', sh.url ? '<img class="receipt-view" src="' + esc(sh.url) + '" alt="Receipt photo">' : '<p class="muted">Loading…</p>');
      case 'confirm':
        return sheetWrap(sh.title, '<p class="muted">' + esc(sh.text) + '</p><button class="btn" style="background:var(--danger)" data-act="confirmYes">' + esc(sh.yes) + '</button><button class="btn ghost" data-act="closeSheet">Cancel</button>');
      default:
        return '';
    }
  }

  function nextDue(b) {
    const d = new Date(b.next_due_on + 'T12:00:00');
    const step = { weekly: [0, 7], monthly: [1, 0], bimonthly: [2, 0], quarterly: [3, 0], yearly: [12, 0] }[b.frequency] || [1, 0];
    d.setMonth(d.getMonth() + step[0]);
    d.setDate(d.getDate() + step[1]);
    return isoOf(d);
  }

  // =====================================================================
  // Actions
  // =====================================================================
  function setModel(path, value) {
    const keys = path.split('.');
    let o = S;
    for (let i = 0; i < keys.length - 1; i++) o = o[keys[i]];
    o[keys[keys.length - 1]] = value;
  }

  async function withBusy(fn) {
    if (S.busy) return;
    S.busy = true;
    render();
    try { await fn(); } catch (e) { toast(errMsg(e), true); } finally { S.busy = false; render(); }
  }

  async function learnRule(merchant, categoryId) {
    if (!merchant || !categoryId) return;
    const r = Receipt.matchRule(merchant, S.rules);
    try {
      if (r) {
        if (r.category_id === null || r.category_id === categoryId) return;
        await api().update('merchant_rules', r.id, { category_id: categoryId, updated_at: new Date().toISOString() });
      } else {
        const pattern = Receipt.norm(merchant).replace(/[^a-z0-9 &.\-]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 40);
        if (pattern.length < 3) return;
        await api().insert('merchant_rules', { pattern, display_name: merchant.trim().slice(0, 60), category_id: categoryId });
      }
    } catch (e) { /* learning is best effort */ }
  }

  async function exportData(format) {
    const { from, to } = S.exp;
    if (!from || !to || from > to) { toast('Check the dates.', true); return; }
    const rows = await api().select('transactions', { gte: { occurred_on: from }, lte: { occurred_on: to }, order: [['occurred_on', true], ['created_at', true]] });
    if (!rows.length) { toast('No entries in those dates.'); return; }
    const head = ['Date', 'Type', 'Category', 'Shop / description', 'Amount (€)', 'Payment method', 'Paid by', 'Added by', 'Note', 'Receipt photo'];
    const data = rows.map((t) => [t.occurred_on, t.kind === 'income' ? 'Income' : +t.amount < 0 ? 'Refund' : 'Expense', (cat(t.category_id) || {}).name || '', t.merchant || '',
      t.kind === 'income' ? +t.amount : -t.amount, methodLabel(t.payment_method), (member(t.paid_by) || {}).display_name || '', (member(t.created_by) || {}).display_name || '', t.note || '', t.receipt_path ? 'Yes' : 'No']);
    const name = 'budget_' + from + '_' + to;
    if (format === 'csv') {
      const cell = (v) => { const s = String(v); return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
      const lines = [head].concat(data.map((r) => r.map((v, i) => (i === 0 ? fmtDate(v) : i === 4 ? NF2.format(v).replace(/\./g, '') : v)))).map((r) => r.map(cell).join(';'));
      await deliver(new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), name + '.csv');
      return;
    }
    await Receipt.loadScript(XLSX_URL);
    const X = window.XLSX;
    const wb = X.utils.book_new();
    const sheetRows = [head].concat(data.map((r) => [new Date(r[0] + 'T12:00:00')].concat(r.slice(1))));
    const ws = X.utils.aoa_to_sheet(sheetRows, { cellDates: true });
    for (let i = 1; i < sheetRows.length; i++) {
      const d = ws[X.utils.encode_cell({ r: i, c: 0 })]; if (d) d.z = 'dd/mm/yyyy';
      const a = ws[X.utils.encode_cell({ r: i, c: 4 })]; if (a) a.z = '#,##0.00 "€";-#,##0.00 "€"';
    }
    ws['!cols'] = [12, 10, 22, 30, 12, 14, 10, 10, 30, 8].map((w) => ({ wch: w }));
    X.utils.book_append_sheet(wb, ws, 'Entries');
    // Monthly summary by category
    const months = [...new Set(rows.map((t) => t.occurred_on.slice(0, 7)))].sort();
    const ecats = S.cats.filter((c) => c.kind === 'expense');
    const sum = [['Month', 'Income', 'Expenses', 'Balance'].concat(ecats.map((c) => c.name))];
    for (const m of months) {
      const mt = rows.filter((t) => t.occurred_on.startsWith(m));
      const inc = round2(mt.filter((t) => t.kind === 'income').reduce((s, t) => s + +t.amount, 0));
      const exp = round2(mt.filter((t) => t.kind === 'expense').reduce((s, t) => s + +t.amount, 0));
      sum.push([MONTHS[+m.slice(5, 7) - 1] + ' ' + m.slice(0, 4), inc, exp, round2(inc - exp)].concat(ecats.map((c) => round2(mt.filter((t) => t.kind === 'expense' && t.category_id === c.id).reduce((s, t) => s + +t.amount, 0)))));
    }
    const ws2 = X.utils.aoa_to_sheet(sum);
    for (let r = 1; r < sum.length; r++) for (let c = 1; c < sum[0].length; c++) { const x = ws2[X.utils.encode_cell({ r, c })]; if (x) x.z = '#,##0.00 "€"'; }
    ws2['!cols'] = sum[0].map((h, i) => ({ wch: i === 0 ? 16 : Math.max(11, h.length + 2) }));
    X.utils.book_append_sheet(wb, ws2, 'By month');
    const out = X.write(wb, { bookType: 'xlsx', type: 'array' });
    await deliver(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name + '.xlsx');
  }

  async function deliver(blob, filename) {
    const file = new File([blob], filename, { type: blob.type });
    if (/iPhone|iPad|Android/i.test(navigator.userAgent) && navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: filename }); return; } catch (e) { if (e.name === 'AbortError') return; }
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 3000);
    toast('Downloaded ' + filename);
  }

  async function handleFile(file) {
    if (!file) return;
    if (file.size > 25 * 1024 * 1024) { toast('That file is too big (max 25 MB).', true); return; }
    S.scan = { step: 'reading', progress: 0.02, previewUrl: /^image\//.test(file.type) ? URL.createObjectURL(file) : null };
    S.form = newForm('expense');
    render();
    let result = null;
    try {
      result = await Receipt.read(file, (p) => {
        S.scan.progress = Math.max(S.scan.progress, p);
        const bar = document.querySelector('.progress > div');
        if (bar) bar.style.width = Math.round(S.scan.progress * 100) + '%';
      });
    } catch (e) {
      console.error(e);
    }
    if (S.route !== 'scan') return;
    if (S.scan.previewUrl) URL.revokeObjectURL(S.scan.previewUrl);
    const parsed = result ? Receipt.parse(result.text, S.rules, todayISO()) : {};
    S.scan = { step: 'review', parsed, photo: result ? result.photo : null, previewUrl: result ? result.previewUrl : null, failed: !result, text: result ? result.text : '' };
    Object.assign(S.form, {
      merchant: parsed.merchant || '', amountText: parsed.amount ? NF2.format(parsed.amount) : '',
      occurred_on: parsed.date || todayISO(), category_id: parsed.categoryId || null, catTouched: !!parsed.categoryId,
      payment_method: parsed.payment || 'card',
    });
    render();
  }

  const A = {
    // navigation
    async month(d) {
      S.month = addMonths(S.month, +d.d);
      await loadMonth().catch((e) => toast(errMsg(e), true));
      render();
    },
    async setMonth(d) { S.month = d.m; await loadMonth().catch((e) => toast(errMsg(e), true)); render(); },
    catHistory(d) { S.hist.cat = d.id; go('#/history'); },
    clearCatFilter() { S.hist.cat = null; render(); },
    openTx(d) { go('#/edit/' + d.id); },
    histWho(d) { S.hist.who = d.id; render(); },
    toggleSearch() { S.hist.search = !S.hist.search; if (!S.hist.search) S.hist.q = ''; render(); if (S.hist.search) { const el = document.getElementById('h-q'); if (el) el.focus(); } },
    closeSheet() { S.sheet = null; S.recovery = false; render(); },
    setSheet(d) { setModel('sheet.' + d.k, d.v); render(); },
    closeForm() { go('#/' + (S.lastMain || 'home')); },

    // login
    async signIn() {
      const L = S.login;
      if (!L.email || !L.password) { L.err = 'Enter your email and password.'; render(); return; }
      L.busy = true; L.err = ''; render();
      try {
        const res = await api().signIn(L.email.trim(), L.password);
        L.password = '';
        await start(res.user);
      } catch (e) { L.err = errMsg(e); }
      L.busy = false;
      render();
    },
    async forgot() {
      const email = S.login.email.trim();
      if (!email) { S.login.err = 'Type your email above first, then tap “Forgot password?”.'; render(); return; }
      try { await api().resetPassword(email); toast('Check your email for a reset link.'); } catch (e) { toast(errMsg(e), true); }
    },
    async signOut() { await api().signOut(); S.user = null; S.ready = false; S.me = null; location.hash = ''; render(); },
    changePassword() { S.sheet = { type: 'password' }; render(); },
    async savePassword() {
      const sh = S.sheet;
      if (!sh.pw1 || sh.pw1.length < 10) { toast('Use at least 10 characters.', true); return; }
      if (sh.pw1 !== sh.pw2) { toast('The two passwords don’t match.', true); return; }
      await withBusy(async () => { await api().updatePassword(sh.pw1); S.recovery = false; S.sheet = null; toast('Password changed.'); });
    },
    demoSwitch() { const id = api().switchUser(); S.user = { id }; S.me = member(id); render(); toast('Now acting as ' + S.me.display_name); },

    // add / edit
    kind(d) {
      const f = S.form;
      if (f.kind === d.k) return;
      f.kind = d.k;
      f.refund = false;
      f.category_id = null;
      f.catTouched = false;
      f.payment_method = d.k === 'income' ? 'transfer' : 'card';
      render();
    },
    key(d) {
      const f = S.form;
      let a = f.amountStr;
      if (d.k === 'del') a = a.slice(0, -1);
      else if (d.k === '.') { if (!a.includes('.')) a = (a || '0') + '.'; }
      else {
        const dot = a.indexOf('.');
        if (dot !== -1 && a.length - dot > 2) return;
        if (a === '0') a = '';
        if (a.replace('.', '').length >= 7) return;
        a += d.k;
      }
      f.amountStr = a;
      render();
    },
    who(d) { S.form.paid_by = d.id; render(); },
    pickCat(d) { S.form.category_id = d.id; S.form.catTouched = true; render(); },
    selectCat(d, el) { S.form.category_id = el.value || null; S.form.catTouched = true; },
    method(d) { S.form.payment_method = d.m; render(); },
    showNote() { S.form.showNote = true; render(); const el = document.getElementById('f-note'); if (el) el.focus(); },
    toggleRefund() { S.form.refund = !S.form.refund; render(); },
    merchantTyped() {
      const f = S.form;
      if (f.catTouched || f.kind !== 'expense') return;
      const r = Receipt.matchRule(f.merchant, S.rules);
      const next = r && r.category_id ? r.category_id : null;
      if (next !== f.category_id) { f.category_id = next; render(); }
    },
    async saveTx() {
      const f = S.form;
      const amount = f.amountText != null ? parseAmount(f.amountText) : round2(parseFloat(f.amountStr || '0'));
      if (!(amount > 0)) { toast('Enter the amount.', true); return; }
      if (!f.category_id) { toast('Pick a category.', true); return; }
      if (!f.occurred_on) { toast('Pick a date.', true); return; }
      await withBusy(async () => {
        let receiptPath = f.receipt_path;
        if (S.route === 'scan' && S.scan && S.scan.photo) {
          receiptPath = f.occurred_on.slice(0, 4) + '/' + f.occurred_on.slice(5, 7) + '/' + uuid() + '.jpg';
          await api().upload(receiptPath, S.scan.photo);
        }
        const signed = f.kind === 'expense' && f.refund ? -amount : amount;
        const row = {
          kind: f.kind, amount: signed, occurred_on: f.occurred_on, merchant: f.merchant.trim() || null, category_id: f.category_id,
          payment_method: f.payment_method || null, paid_by: f.paid_by, note: (f.note || '').trim() || null, receipt_path: receiptPath || null,
        };
        try {
          if (f.id) await api().update('transactions', f.id, row);
          else await api().insert('transactions', row);
        } catch (e) {
          if (receiptPath && receiptPath !== f.receipt_path) api().removeFiles([receiptPath]).catch(() => {});
          throw e;
        }
        if (f.kind === 'expense' && !f.refund) await learnRule(row.merchant, row.category_id);
        if (monthOf(f.occurred_on) !== S.month) S.month = monthOf(f.occurred_on);
        await refresh();
        toast((f.id ? 'Saved · ' : f.refund ? 'Refund added · ' : 'Added · ') + eur(amount) + (row.merchant ? ' · ' + row.merchant : ''));
        S.busy = false;
        go('#/' + (S.lastMain || 'home'));
      });
    },
    deleteTx() {
      const f = S.form;
      S.sheet = { type: 'confirm', title: 'Delete this entry?', text: 'This removes it for both of you' + (f.receipt_path ? ', including the receipt photo' : '') + '.', yes: 'Delete', onYes: async () => {
        await api().remove('transactions', f.id);
        if (f.receipt_path) api().removeFiles([f.receipt_path]).catch(() => {});
        await refresh();
        toast('Deleted.');
        go('#/' + (S.lastMain || 'home'));
      } };
      render();
    },
    async confirmYes() { const fn = S.sheet && S.sheet.onYes; S.sheet = null; if (fn) await withBusy(fn); },
    async viewReceipt(d) {
      S.sheet = { type: 'image', url: null };
      render();
      try { S.sheet.url = await api().signedUrl(d.path); render(); } catch (e) { S.sheet = null; toast(errMsg(e), true); }
    },

    // scan
    async fileChosen(d, el) { const file = el.files && el.files[0]; el.value = ''; await handleFile(file); },
    retake() { if (S.scan && S.scan.previewUrl) URL.revokeObjectURL(S.scan.previewUrl); S.scan = { step: 'pick' }; render(); },
    viewPreview() { if (S.scan.previewUrl) { S.sheet = { type: 'image', url: S.scan.previewUrl }; render(); } },

    // bills
    payBill(d) {
      const b = S.bills.find((x) => x.id === d.id);
      const today = todayISO();
      S.sheet = { type: 'pay', billId: b.id, amount: NF2.format(+b.amount), date: b.next_due_on < today ? b.next_due_on : today, method: b.payment_method || 'direct_debit' };
      render();
    },
    async confirmPay() {
      const sh = S.sheet;
      const amount = parseAmount(sh.amount);
      if (!(amount > 0)) { toast('Enter the amount.', true); return; }
      await withBusy(async () => {
        const b = S.bills.find((x) => x.id === sh.billId);
        await api().payBill(sh.billId, amount, sh.date, sh.method);
        S.sheet = null;
        await refresh();
        toast('Paid · ' + b.name + ' ' + eur(amount));
      });
    },
    newBill() {
      S.sheet = { type: 'bill', f: { id: null, name: '', amount: '', amount_is_estimate: false, category_id: (expenseCats()[0] || {}).id, payment_method: 'direct_debit', frequency: 'monthly', next_due_on: addMonths(monthOf(todayISO()), 1), ends_on: '', remind_days_before: 3, active: true } };
      render();
    },
    openBill(d) {
      const b = S.bills.find((x) => x.id === d.id);
      S.sheet = { type: 'bill', f: Object.assign({}, b, { amount: NF2.format(+b.amount), ends_on: b.ends_on || '' }) };
      render();
    },
    async saveBill() {
      const f = S.sheet.f;
      const amount = parseAmount(f.amount);
      if (!f.name.trim()) { toast('Give the bill a name.', true); return; }
      if (!(amount > 0)) { toast('Enter the amount.', true); return; }
      if (!f.next_due_on) { toast('Pick the next due date.', true); return; }
      const row = { name: f.name.trim(), amount, amount_is_estimate: !!f.amount_is_estimate, category_id: f.category_id, payment_method: f.payment_method,
        frequency: f.frequency, next_due_on: f.next_due_on, ends_on: f.ends_on || null, remind_days_before: Math.max(0, Math.min(60, parseInt(f.remind_days_before, 10) || 0)), active: f.id ? !!f.active : true };
      await withBusy(async () => {
        if (f.id) await api().update('recurring_bills', f.id, row); else await api().insert('recurring_bills', row);
        S.sheet = null;
        await refresh();
        toast('Bill saved.');
      });
    },
    deleteBill() {
      const f = S.sheet.f;
      S.sheet = { type: 'confirm', title: 'Delete “' + f.name + '”?', text: 'Past payments stay in History. Tip: to stop it for a while, untick “Active” instead.', yes: 'Delete bill', onYes: async () => {
        await api().remove('recurring_bills', f.id); await refresh(); toast('Bill deleted.');
      } };
      render();
    },

    // budgets & categories
    scope(d) { S.budgetScope = d.s; render(); },
    async setBudget(d, el) {
      const v = parseAmount(el.value || '0');
      const old = budgetFor(d.id, S.month);
      if (isNaN(v) || v < 0) { el.value = numText(old); toast('Enter a number, e.g. 1.200', true); return; }
      if (v === old) { el.value = numText(v); return; }
      await withBusy(async () => {
        if (S.budgetScope === 'once') {
          const next = addMonths(S.month, 1);
          if (!S.budgetRows.find((r) => r.category_id === d.id && r.month === next)) await api().upsertBudget({ category_id: d.id, month: next, amount: old });
          await api().upsertBudget({ category_id: d.id, month: S.month, amount: v });
        } else {
          await api().upsertBudget({ category_id: d.id, month: S.month, amount: v });
          await api().updateWhere('category_budgets', { eq: { category_id: d.id }, gt: { month: S.month } }, { amount: v });
        }
        await refresh();
        toast((cat(d.id) || {}).name + ': ' + eur(v, { smart: true }) + (S.budgetScope === 'once' ? ' in ' + monthLabel(S.month) : ' from ' + monthLabel(S.month)));
      });
    },
    newCat(d) { S.sheet = { type: 'cat', f: { id: null, name: '', kind: d.kind, budget: '', color: '#64748B' } }; render(); },
    openCat(d) { const c = cat(d.id); S.sheet = { type: 'cat', f: { id: c.id, name: c.name, kind: c.kind, color: c.color || '#64748B' } }; render(); },
    async saveCat() {
      const f = S.sheet.f;
      if (!f.name.trim()) { toast('Give it a name.', true); return; }
      await withBusy(async () => {
        if (f.id) await api().update('categories', f.id, { name: f.name.trim(), color: f.color });
        else {
          const sort = Math.max(0, ...S.cats.filter((c) => c.kind === f.kind).map((c) => c.sort_order || 0)) + 1;
          const c = await api().insert('categories', { name: f.name.trim(), kind: f.kind, color: f.color, sort_order: sort });
          const b = parseAmount(f.budget || '0');
          if (f.kind === 'expense' && b > 0) await api().upsertBudget({ category_id: c.id, month: S.month, amount: b });
        }
        S.sheet = null;
        await refresh();
        toast('Category saved.');
      });
    },
    archiveCat() {
      const f = S.sheet.f;
      S.sheet = { type: 'confirm', title: 'Hide “' + f.name + '”?', text: 'It disappears from budgets and pickers. Past entries keep it.', yes: 'Hide', onYes: async () => {
        await api().update('categories', f.id, { archived: true }); await refresh(); toast('Category hidden.');
      } };
      render();
    },

    // savings
    newGoal() { S.sheet = { type: 'goalEdit', f: { id: null, name: '', target_amount: '', monthly_plan: '', target_date: '' } }; render(); },
    openGoal(d) { S.sheet = { type: 'goal', goalId: d.id }; render(); },
    editGoal(d) {
      const g = S.goals.find((x) => x.id === d.id);
      S.sheet = { type: 'goalEdit', f: { id: g.id, name: g.name, target_amount: numText(+g.target_amount), monthly_plan: g.monthly_plan ? numText(+g.monthly_plan) : '', target_date: g.target_date || '' } };
      render();
    },
    async saveGoal() {
      const f = S.sheet.f;
      const target = parseAmount(f.target_amount);
      const plan = f.monthly_plan ? parseAmount(f.monthly_plan) : null;
      if (!f.name.trim()) { toast('Give the goal a name.', true); return; }
      if (!(target > 0)) { toast('Enter the target amount.', true); return; }
      const row = { name: f.name.trim(), target_amount: target, monthly_plan: plan && plan > 0 ? plan : null, target_date: f.target_date || null };
      await withBusy(async () => {
        if (f.id) await api().update('savings_goals', f.id, row); else await api().insert('savings_goals', row);
        S.sheet = null;
        await refresh();
        toast('Goal saved.');
      });
    },
    archiveGoal() {
      const f = S.sheet.f;
      S.sheet = { type: 'confirm', title: 'Archive “' + f.name + '”?', text: 'It disappears from Savings. Nothing else changes.', yes: 'Archive', onYes: async () => {
        await api().update('savings_goals', f.id, { archived: true }); await refresh(); toast('Goal archived.');
      } };
      render();
    },
    contribute(d) { S.sheet = { type: 'contribute', goalId: d.id, dir: d.dir, amount: '', date: todayISO() }; render(); },
    putLeftover(d) {
      const goals = S.goals.filter((g) => !g.archived);
      if (!goals.length) { toast('Create a goal first.'); A.newGoal(); return; }
      S.sheet = { type: 'contribute', goalId: goals[0].id, dir: 'in', amount: numText(+d.v), date: todayISO() };
      render();
    },
    async saveContribution() {
      const sh = S.sheet;
      const v = parseAmount(sh.amount);
      if (!(v > 0)) { toast('Enter the amount.', true); return; }
      await withBusy(async () => {
        await api().insert('savings_contributions', { goal_id: sh.goalId, amount: sh.dir === 'out' ? -v : v, occurred_on: sh.date || todayISO() });
        const goalId = sh.goalId;
        await refresh();
        S.sheet = { type: 'goal', goalId };
        toast(sh.dir === 'out' ? 'Taken out.' : 'Added to the goal.');
      });
    },
    async deleteContribution(d) {
      await withBusy(async () => { await api().remove('savings_contributions', d.id); await refresh(); });
    },
    editTarget() { S.sheet = { type: 'target', amount: numText(+S.settings.monthly_savings_target || 0) }; render(); },
    async saveTarget() {
      const v = parseAmount(S.sheet.amount || '0');
      if (isNaN(v) || v < 0) { toast('Enter a number.', true); return; }
      await withBusy(async () => { await api().update('household_settings', 1, { monthly_savings_target: v, updated_at: new Date().toISOString() }); S.sheet = null; await refresh(); toast('Target saved.'); });
    },

    // export
    async export(d) { await withBusy(() => exportData(d.f)); },
  };

  // =====================================================================
  // Events
  // =====================================================================
  document.addEventListener('click', (e) => {
    if (e.target.classList && e.target.classList.contains('overlay')) { S.sheet = null; render(); return; }
    if (e.target.type === 'date' && e.target.showPicker) { try { e.target.showPicker(); } catch (err) { /* iOS opens its own picker */ } }
    const el = e.target.closest('[data-act]');
    if (!el || el.disabled) return;
    const fn = A[el.dataset.act];
    if (fn) { e.preventDefault(); Promise.resolve(fn(el.dataset, el, e)).catch((err) => toast(errMsg(err), true)); }
  });
  document.addEventListener('input', (e) => {
    const el = e.target;
    if (!el.dataset || !el.dataset.model) return;
    setModel(el.dataset.model, el.type === 'checkbox' ? el.checked : el.value);
    if (el.dataset.after) A[el.dataset.after](el.dataset, el);
    if ('rerender' in el.dataset) render();
  });
  document.addEventListener('change', (e) => {
    const el = e.target;
    if (el.type === 'checkbox' && el.dataset.model) setModel(el.dataset.model, el.checked);
    if (el.dataset && el.dataset.change) Promise.resolve(A[el.dataset.change](el.dataset, el, e)).catch((err) => toast(errMsg(err), true));
  });
  document.addEventListener('submit', (e) => {
    const f = e.target;
    if (f.dataset.submit) { e.preventDefault(); A[f.dataset.submit](f.dataset, f); }
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && S.sheet) { S.sheet = null; render(); return; }
    if ((S.route !== 'add' && S.route !== 'edit') || S.sheet || !S.form) return;
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
    if (/^[0-9]$/.test(e.key)) A.key({ k: e.key });
    else if (e.key === ',' || e.key === '.') A.key({ k: '.' });
    else if (e.key === 'Backspace') A.key({ k: 'del' });
    else if (e.key === 'Enter') A.saveTx();
    else return;
    e.preventDefault();
  });

  // =====================================================================
  // Start
  // =====================================================================
  async function cleanupReceipts() {
    try {
      const key = 'ob-receipt-cleanup';
      if (localStorage.getItem(key) === todayISO()) return;
      const d = new Date();
      d.setFullYear(d.getFullYear() - 1);
      const old = await api().select('transactions', { notnull: ['receipt_path'], lt: { created_at: d.toISOString() }, limit: 50 });
      if (old.length) {
        await api().removeFiles(old.map((t) => t.receipt_path));
        for (const t of old) await api().update('transactions', t.id, { receipt_path: null });
      }
      localStorage.setItem(key, todayISO());
    } catch (e) { /* try again next time */ }
  }

  async function start(user) {
    S.user = user;
    render();
    try {
      await refresh();
    } catch (e) {
      toast(errMsg(e), true);
    }
    S.ready = true;
    onRoute();
    if (S.me) {
      api().subscribe(onRemoteChange);
      cleanupReceipts();
    }
  }

  async function boot() {
    const local = /^(localhost|127\.0\.0\.1)$/.test(location.hostname);
    // Demo mode (local testing only): demo.js is not deployed.
    const demo = local && new URLSearchParams(location.search).has('demo');
    if (demo) await Receipt.loadScript('demo.js');
    S.api = demo ? window.DemoAPI() : SupaAPI(window.APP_CONFIG);
    window.addEventListener('hashchange', () => { if (S.ready) onRoute(); });
    document.addEventListener('visibilitychange', () => { if (!document.hidden && S.ready && S.me) onRemoteChange(); });
    S.api.onAuth((event) => {
      if (event === 'PASSWORD_RECOVERY') { S.recovery = true; S.sheet = { type: 'password', recovery: true }; render(); }
      if (event === 'SIGNED_OUT' && S.user) { S.user = null; S.ready = false; S.me = null; render(); }
    });
    try {
      const session = await S.api.session();
      if (session) await start(session.user);
      else render();
    } catch (e) {
      render();
      toast(errMsg(e), true);
    }
    if ('serviceWorker' in navigator && (location.protocol === 'https:' || local) && !demo) {
      // When a new version is published, reload once it's safe (not mid-entry).
      const hadController = !!navigator.serviceWorker.controller;
      navigator.serviceWorker.addEventListener('controllerchange', () => {
        if (!hadController) return;
        S.updateReady = true;
        reloadIfIdle();
      });
      navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).then((reg) => {
        document.addEventListener('visibilitychange', () => { if (!document.hidden) reg.update().catch(() => {}); });
      }).catch(() => {});
    }
  }

  boot();
})();
