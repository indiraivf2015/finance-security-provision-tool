// Shared screens for every provision page (Upload → Download). Copied from the approved Security page's script;
// the provision-specific texts and figures come from the provision's engine (window.ProvisionEngine.ui).
// DOM is built with textContent only (no innerHTML).
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const E = window.ProvisionEngine;
  const U = E.ui;
  $('sopver').textContent = E.SOP.version;
  $('foot').textContent = E.toolName + ' ' + E.TOOL_VERSION + ' · ' + E.SOP.version + ' · works offline';
  for (const k of ['extract', 'opening', 'rajeev']) { $('t-' + k).textContent = U.files[k].title; $('s-' + k).textContent = U.files[k].desc; }
  $('gen-note').textContent = U.note;
  $('points-note').textContent = U.pointsNote;

  const files = { extract: null, opening: null, rajeev: null };
  const names = { extract: '', opening: '', rajeev: '' };
  let wbs = null, state = null, opts = {}, output = null;
  const SVGNS = 'http://www.w3.org/2000/svg';

  function el(tag, props, kids) {
    const n = document.createElement(tag);
    if (props) for (const [k, v] of Object.entries(props)) { if (k === 'class') n.className = v; else if (k === 'text') n.textContent = v; else n.setAttribute(k, v); }
    (kids || []).forEach((c) => n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c));
    return n;
  }
  function icon(id) { const s = document.createElementNS(SVGNS, 'svg'); s.setAttribute('class', 'i'); const u = document.createElementNS(SVGNS, 'use'); u.setAttribute('href', '#' + id); s.appendChild(u); return s; }
  const clear = (n) => { while (n.firstChild) n.removeChild(n.firstChild); };
  const inr = (x) => E.fmtINR(x, U.decimals); // U.decimals: 0 = whole rupees on screen (exact values are kept)

  function showPage(n) {
    ['sec-files', 'sec-gen'].forEach((id, i) => $(id).classList.toggle('hidden', i + 1 !== n));
    setStep(n);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
  $('st1').addEventListener('click', () => { if ($('st1').classList.contains('done')) showPage(1); });
  function setStep(n) {
    for (let i = 1; i <= 2; i++) {
      const s = $('st' + i); s.className = 'st' + (i < n ? ' done' : i === n ? ' cur' : '');
      const badge = s.querySelector('.n'); clear(badge);
      if (i < n) badge.appendChild(icon('i-check')); else badge.textContent = String(i);
    }
  }

  function resetDownstream() {
    wbs = null; state = null; opts = {}; output = null;
    ['sec-gen', 'result', 'errors', 'choice', 'month-chip'].forEach((id) => $(id).classList.add('hidden'));
    $('sec-files').classList.remove('hidden');
    setStep(1);
  }

  async function setFile(k, f) {
    resetDownstream();
    const box = $('d-' + k), nm = $('n-' + k); clear(nm);
    if (!f) { files[k] = null; names[k] = ''; nm.appendChild(el('span', { text: 'No file chosen' })); box.classList.remove('set'); }
    else {
      files[k] = await f.arrayBuffer(); names[k] = f.name;
      nm.appendChild(icon('i-check')); nm.appendChild(el('span', { text: f.name + ' · ' + (f.size / 1024).toFixed(0) + ' KB', title: f.name }));
      box.classList.add('set');
    }
    box.querySelector('.ic').replaceChildren(icon(f ? 'i-sheet' : 'i-upload'));
    $('b-read').disabled = !(files.extract && files.opening && files.rajeev);
  }
  for (const k of Object.keys(files)) {
    $('f-' + k).addEventListener('change', (ev) => setFile(k, ev.target.files && ev.target.files[0]));
    const box = $('d-' + k);
    box.addEventListener('dragover', (ev) => { ev.preventDefault(); box.classList.add('over'); });
    box.addEventListener('dragleave', () => box.classList.remove('over'));
    box.addEventListener('drop', (ev) => { ev.preventDefault(); box.classList.remove('over'); const f = ev.dataTransfer && ev.dataTransfer.files[0]; if (f) setFile(k, f); });
  }

  function showErrors(list) {
    const ul = $('error-list'); clear(ul);
    list.forEach((m) => ul.appendChild(el('li', { text: m })));
    $('errors').classList.remove('hidden');
  }

  $('b-read').addEventListener('click', async () => {
    if (state && output) { showPage(2); return; } // same files already processed
    resetDownstream();
    $('b-read').disabled = true; $('read-status').textContent = 'Reading files…';
    try {
      wbs = await E.load(files);
      await runAnalyse();
    } catch (e) {
      showErrors(e.list || ['Unexpected error: ' + e.message]);
    } finally { $('b-read').disabled = false; $('read-status').textContent = ''; }
  });

  async function runAnalyse() {
    $('errors').classList.add('hidden'); $('choice').classList.add('hidden');
    let st;
    try { st = E.analyse(wbs, opts); } catch (e) { showErrors(e.list || ['Unexpected error: ' + e.message]); return; }
    if (st.needChoice) { askChoice(st.needChoice); return; }
    state = st; output = null;
    $('month-txt').textContent = st.period.display; $('month-chip').classList.remove('hidden');
    $('read-status').textContent = 'Building the statement…';
    try {
      const now = new Date();
      // No year choices: lines whose description has no year are counted as the current FY (finance rule)
      const res = await E.build(state, new Map(), { now, nowText: now.toLocaleString('en-IN'), names });
      output = { blob: new Blob([res.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name: U.outputName(state.period) };
      showResult(res.result);
    } catch (e) {
      state = null; showErrors(e.list || ['Unexpected error: ' + e.message]);
    } finally { $('read-status').textContent = ''; }
  }

  function askChoice(c) {
    const body = $('choice-body'); clear(body);
    let get;
    if (c.kind === 'openingYtdCol') {
      $('choice-title').textContent = `Last month's statement has more than one "${c.header}" column. Which one is the closing balance?`;
      body.appendChild(el('p', { class: 'hint', text: 'The SOP carries last month\'s closing balance into this month\'s opening. Choose the column that holds the final closing figure.' }));
      const sel = el('select'); c.options.forEach((o) => sel.appendChild(el('option', { value: o.col, text: `Column ${o.col} — "${o.header}" — total ${inr(o.total)}` })));
      body.appendChild(sel); get = () => { opts.openingYtdCol = sel.value; };
    } else if (c.kind === 'openingSheet' || c.kind === 'rajeevSheet') {
      $('choice-title').textContent = c.kind === 'openingSheet' ? 'More than one sheet in last month\'s file looks like a statement. Which one should be used?' : `Which sheet of the provision file holds the ${c.noun} provision?`;
      const sel = el('select'); c.options.forEach((o) => sel.appendChild(el('option', { value: o, text: o })));
      body.appendChild(sel); get = () => { opts[c.kind] = sel.value; };
    } else if (c.kind === 'confirmOpeningMonth') {
      $('choice-title').textContent = 'Please confirm the month of last month\'s statement';
      const cb = el('input', { type: 'checkbox', id: 'cb-month' });
      body.appendChild(el('label', { style: 'display:flex;gap:8px;margin-top:8px' }, [cb, `Sheet "${c.sheet}" has no "for the month of …" title. I confirm it is the statement for ${c.month}.`]));
      get = () => { if (!cb.checked) return false; opts.confirmOpeningMonth = true; };
    }
    $('b-choice').onclick = async () => { if (get() === false) return; $('b-choice').disabled = true; try { await runAnalyse(); } finally { $('b-choice').disabled = false; } };
    $('choice').classList.remove('hidden');
  }

  function showResult(R) {
    const failed = R.checks.filter((c) => c.status === 'FAIL').length;
    $('result').className = 'result ' + (failed ? 'bad' : 'ok');
    $('r-ico').replaceChildren(icon(failed ? 'i-x-circle' : 'i-check-circle'));
    $('r-status').textContent = failed ? `NOT READY — ${failed} check(s) failed. Do not use this statement.` : `READY — all ${R.checks.length} checks passed`;
    $('r-status').className = 'status';
    $('r-file').textContent = output.name + ' · ' + U.sheets;
    const keys = $('r-keys'); clear(keys);
    for (const [label, v, note] of U.keys(R)) keys.appendChild(el('div', { class: 'key' }, [el('span', { text: label }), el('strong', { text: inr(v) }), el('small', { text: note })]));
    $('r-details').open = !!failed;
    const np = R.points.length;
    $('r-points-sum').replaceChildren(icon('i-alert'), document.createTextNode(np ? `${np} point${np > 1 ? 's' : ''} to check` : 'No points to check this month'));
    const pb = $('r-points-body'); clear(pb);
    for (const p of R.points) pb.appendChild(el('tr', {}, [el('td', { text: p.code != null ? `${p.code} ${p.name}` : '—' }), el('td', { text: p.point }),
      el('td', { class: 'num', text: p.amount == null ? '' : inr(p.amount) }), el('td', { text: p.check })]));
    $('r-points').classList.toggle('hidden', !np);
    const t = $('r-totals'); clear(t);
    for (const [k, v] of U.totals(R)) t.appendChild(el('div', { class: 'kv' }, [el('span', { text: k }), el('strong', { text: inr(v) })]));
    const cb = $('r-checks'); clear(cb);
    for (const c of R.checks) {
      const isCount = c.kind === 'count0' || c.kind === 'count0fail';
      cb.appendChild(el('tr', {}, [el('td', { text: String(c.id) }), el('td', { text: c.sop }), el('td', { text: c.text }),
        el('td', { class: 'num', text: isCount ? String(c.expected) : inr(c.expected) }), el('td', { class: 'num', text: isCount ? String(c.actual) : inr(c.actual) }),
        el('td', {}, [el('span', { class: 'badge ' + c.status, text: c.status })])]));
    }
    $('result').classList.remove('hidden');
    showPage(2);
  }

  $('b-back2').addEventListener('click', () => showPage(1));
  $('b-new').addEventListener('click', () => { if (confirm('Start again with new files? The current statement will be cleared from this page (download it first if you need it).')) location.reload(); });
  $('b-dl').addEventListener('click', () => {
    if (!output) return;
    const url = URL.createObjectURL(output.blob);
    const a = el('a', { href: url, download: output.name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  });
  setStep(1);
})();
