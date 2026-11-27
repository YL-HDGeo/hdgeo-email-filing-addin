// HD Geo email filing panel.
// Runs in two modes:
//  - Preview: a normal browser, with sample emails and projects (nothing is saved).
//  - Outlook: reads the open email, files it through the "Email filing - File email"
//    Power Automate flow and searches through "Email filing - Search".
//    Both flow links live in the user's Outlook roaming settings, never in this code.
//
// Filing rules:
//  - Meeting invites are never filed.
//  - Subject with exactly one HD number: filed automatically (if auto-file is on).
//  - No HD number: look for a clue in filed emails (same thread, same client contact,
//    recent ones count most). One clear project -> "File to HDxxxx?" Yes/No.
//    Otherwise the likely projects are listed, or the person types an HD number or
//    searches project sites by job name, address or client.
//  - Several HD numbers: the person picks.
//  - Search results open in SharePoint's file viewer, not as a download.
//  - A filed email is marked on the email itself, so opening it again doesn't re-file.

(function () {
  const data = window.SampleData;
  const M = window.Matching;
  const state = {
    mode: 'preview',
    email: null,
    result: null,
    selectedHd: null,
    projects: data.projects,
    filedLog: data.filedLog.slice(),
    previewFiled: {},       // preview only: messageId -> filed record
    card: { kind: 'idle' },
    chooserOpen: false,
    token: 0,               // bumps on every new email, so late replies are ignored
    askToken: 0,            // bumps on every search, so late replies are ignored
    props: null,            // Outlook custom properties of the open email
  };

  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nzDate = iso => {
    if (!iso) return '';
    const [y, m, d] = String(iso).slice(0, 10).split('-');
    return d ? `${d}/${m}/${y}` : '';
  };
  const projectByHd = hd => state.projects.find(p => p.hdNumber === hd);
  // Outlook has no project list yet: an HD number stands in for the project and
  // the flow checks it against the Project Register.
  const stubProject = (hd, name) => ({ hdNumber: hd, jobName: name || '', client: '', status: '', siteUrl: '', folders: [] });
  const addStub = (hd, name) => {
    const p = projectByHd(hd);
    if (!p) state.projects.push(stubProject(hd, name));
    else if (name && !p.jobName) p.jobName = name;
  };
  // "HD4275 - Classic Developments - 300 Peacocke" -> "Classic Developments - 300 Peacocke"
  const jobFromSiteTitle = t => String(t || '').replace(/^\s*HD\s?\d{3,5}(-\d+)?\s*[-–:]?\s*/i, '').trim();

  // ---------- Settings ----------

  const KEYS = { file: 'fileFlowUrl', search: 'searchFlowUrl', auto: 'autoFile' };
  function getSetting(key) {
    try {
      if (state.mode === 'outlook') return Office.context.roamingSettings.get(key);
      return localStorage.getItem(key);
    } catch (e) { return null; }
  }
  function saveSettings(values, done) {
    if (state.mode === 'outlook') {
      const rs = Office.context.roamingSettings;
      Object.entries(values).forEach(([k, v]) => (v === '' || v == null ? rs.remove(k) : rs.set(k, v)));
      rs.saveAsync(r => done(r.status === Office.AsyncResultStatus.Succeeded));
    } else {
      try {
        Object.entries(values).forEach(([k, v]) => (v === '' || v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v)));
        done(true);
      } catch (e) { done(false); }
    }
  }
  const fileUrl = () => getSetting(KEYS.file) || '';
  const searchUrl = () => getSetting(KEYS.search) || '';
  const autoOn = () => getSetting(KEYS.auto) !== '0';
  const looksLikeFlowUrl = u => /^https:\/\/[^/]+\.(powerplatform\.com|logic\.azure\.com)(:\d+)?\//i.test(u);

  function openSettings(open) {
    $('view-main').hidden = open;
    $('view-settings').hidden = !open;
    if (open) {
      $('flow-url').value = fileUrl();
      $('search-url').value = searchUrl();
      $('auto-file').checked = autoOn();
      $('settings-help').textContent = 'Saved in your own Outlook settings only. Treat these links like passwords: anyone who has them can file or search project emails.';
    }
  }
  function wireSettings() {
    $('open-settings').addEventListener('click', () => openSettings(true));
    $('close-settings').addEventListener('click', () => { openSettings(false); renderCard(); });
    $('save-settings').addEventListener('click', () => {
      const f = $('flow-url').value.trim();
      const s = $('search-url').value.trim();
      if ((f && !looksLikeFlowUrl(f)) || (s && !looksLikeFlowUrl(s))) {
        $('settings-help').textContent = 'One of those doesn\'t look like a Power Automate flow link. Copy the HTTP URL from the flow\'s trigger.';
        return;
      }
      saveSettings({ [KEYS.file]: f, [KEYS.search]: s, [KEYS.auto]: $('auto-file').checked ? '1' : '0' }, ok => {
        $('settings-help').textContent = ok ? 'Saved.' : 'Could not save. Try again.';
        if (ok) setTimeout(() => { openSettings(false); if (state.email) showEmail(state.email); }, 500);
      });
    });
  }

  // ---------- Start-up ----------

  function start(inOutlook) {
    state.mode = inOutlook ? 'outlook' : 'preview';
    wireSettings();
    wireAsk();

    if (inOutlook) {
      state.projects = [];
      state.filedLog = [];
      $('ask-input').placeholder = 'e.g. tom waikato bilders retaning wall last month';
      loadOutlookEmail();
      // Pinned pane: follow the person from email to email.
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadOutlookEmail);
    } else {
      startPreview();
    }
  }

  function startPreview() {
    document.body.classList.add('preview-mode');
    ['ribbon', 'mail-list', 'reading', 'host-head'].forEach(id => { $(id).hidden = false; });
    // The preview pretends both flow links are set so every state can be seen.
    state.previewLinks = true;

    const list = $('mail-list');
    list.innerHTML = data.emails.map((e, i) => `
      <button type="button" class="mail-item" data-i="${i}">
        <span class="who">${esc(e.fromName)}</span>
        <span class="subj">${esc(e.subject)}</span>
        <span class="scenario">${esc(e.label)}</span>
      </button>`).join('');
    list.querySelectorAll('.mail-item').forEach(btn => btn.addEventListener('click', () => {
      list.querySelectorAll('.mail-item').forEach(b => b.classList.toggle('active', b === btn));
      showEmail(data.emails[btn.dataset.i]);
    }));
    const setPane = open => {
      document.body.classList.toggle('pane-closed', !open);
      $('ribbon-file').setAttribute('aria-pressed', open);
    };
    $('ribbon-file').addEventListener('click', () => setPane(document.body.classList.contains('pane-closed')));
    $('host-close').addEventListener('click', () => setPane(false));
    list.querySelector('.mail-item').click();
    runAsk('');
  }

  if (window.Office && Office.onReady) {
    Office.onReady(info => start(info.host === Office.HostType.Outlook));
  } else {
    document.addEventListener('DOMContentLoaded', () => start(false));
  }

  // ---------- Reading the email ----------

  function loadOutlookEmail() {
    const item = Office.context.mailbox.item;
    if (!item) return;
    const addr = a => (a && a.emailAddress) || '';
    showEmail({
      messageId: item.itemId,
      conversationId: item.conversationId,
      itemType: item.itemType,
      itemClass: item.itemClass,
      subject: item.subject || '',
      from: addr(item.from),
      fromName: (item.from && item.from.displayName) || '',
      to: (item.to || []).map(addr),
      cc: (item.cc || []).map(addr),
      date: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : '',
      body: '',
    });
  }

  function showEmail(email) {
    const token = ++state.token;
    state.email = email;
    state.chooserOpen = false;
    state.props = null;
    if (state.mode === 'preview') renderSampleEmail(email);
    else M.extractHdNumbers(email.subject).forEach(addStub);

    state.result = M.suggestProjects(email, state.projects, state.filedLog);
    state.selectedHd = state.result.autoSelect;

    if (state.result.blocked === 'calendar') { setCard({ kind: 'blocked' }); return; }

    readFiledMarker(record => {
      if (token !== state.token) return;
      if (record) { setCard({ kind: 'filed', ...record, earlier: true }); return; }
      const hdsInSubject = M.extractHdNumbers(email.subject);
      if (autoOn() && hasFileLink() && hdsInSubject.length === 1) {
        fileNow(hdsInSubject[0], true);
      } else if (hdsInSubject.length === 0 && hasSearchLink()) {
        // No HD number: look for a clear clue (same thread or same client filed before).
        setCard({ kind: 'looking' });
        findClue(email, (clue, candidates) => {
          if (token !== state.token) return;
          (candidates || []).forEach(c => addStub(c.hd, c.job));
          if (clue) {
            state.selectedHd = clue.hd;
            setCard({ kind: 'suggest', ...clue });
          } else if (candidates && candidates.length && state.mode === 'outlook') {
            // No clear winner: show the likely projects to pick from.
            state.result = { ...state.result, autoSelect: null,
              suggestions: candidates.map(c => ({ project: projectByHd(c.hd), reason: c.why })) };
            setCard({ kind: 'choose' });
          } else {
            setCard({ kind: 'choose', noClue: true });
          }
        });
      } else {
        setCard({ kind: 'choose' });
      }
    });
  }

  // ---------- Clues for emails without an HD number ----------

  // Calls the search flow with an exact SharePoint (KQL) search. Resolves to search rows.
  function searchExact(kql, params) {
    return fetch(searchUrl(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: 'exact', kql, params: params || '' }),
    })
      .then(r => (r.ok ? r.json() : Promise.reject(new Error('search ' + r.status))))
      .then(res => rawRows(res.search));
  }

  // HD number of a search row: dash-job folder in the path first, then the site.
  function hdFromRow(row) {
    const path = String(row.Path || '');
    const dash = path.match(/\/(HD\d{3,5}-\d+)\s*-[^/]*\//i);
    if (dash) return dash[1].toUpperCase();
    const t = String(row.SiteTitle || '').match(/^\s*HD\s?(\d{3,5})\b/i)
      || path.match(/\/sites\/HD(\d{3,5})(?![\d-])/i);
    return t ? 'HD' + t[1] : '';
  }

  const ownDomain = /@hdgeo\.co\.nz$/i;
  const topicOf = subject => String(subject || '')
    .replace(/^\s*((re|fw|fwd|aw)\s*:\s*)+/i, '')
    .replace(/[^\w\s&]/g, ' ').replace(/\s+/g, ' ').trim();

  // cb(clue or null). A clue is only given when one project clearly wins.
  function findClue(email, cb) {
    if (state.mode === 'preview') {
      const r = state.result || { suggestions: [] };
      const s = r.suggestions[0];
      if (s && (r.autoSelect || r.suggestions.length === 1)) {
        setTimeout(() => cb({ hd: s.project.hdNumber, job: s.project.jobName, why: s.reason }), 500);
      } else setTimeout(() => cb(null), 500);
      return;
    }
    const addresses = [email.from].concat(email.to || [], email.cc || [])
      .filter(a => a && !ownDomain.test(a))
      .filter((a, i, all) => all.indexOf(a) === i).slice(0, 4);
    const topic = topicOf(email.subject);
    const useTopic = topic.split(' ').length >= 3 || topic.length >= 15;
    const parts = addresses.map(a => `"${a}"`);
    if (useTopic) parts.unshift(`Title:"${topic}"`);
    if (!parts.length) { cb(null); return; }

    searchExact(`(${parts.join(' OR ')}) (FileExtension:eml OR FileExtension:msg)`)
      .then(rows => {
        const norm = s => topicOf(String(s || '').replace(/^\d{14}\s*-\s*/, '').replace(/_/g, ' ')).toLowerCase();
        const now = Date.now();
        const groups = {};
        rows.forEach(row => {
          const hd = hdFromRow(row);
          if (!hd) return;
          // Filed date from the file name (yyyyMMddHHmmss - …), else the modified date.
          const m = String(row.Title || '').match(/^(\d{4})(\d{2})(\d{2})\d{6}/);
          const iso = m ? `${m[1]}-${m[2]}-${m[3]}` : String(row.Write || '').slice(0, 10);
          const days = iso ? (now - Date.parse(iso)) / 864e5 : 9999;
          // Recent emails count most: a contact's old projects shouldn't win over the current one.
          const weight = days <= 60 ? 5 : days <= 180 ? 2 : days <= 365 ? 0.7 : 0.3;
          const g = groups[hd] || (groups[hd] = { hd, job: jobFromSiteTitle(row.SiteTitle), score: 0, thread: 0, client: 0, last: '' });
          if (useTopic && norm(row.Title).includes(topic.toLowerCase())) { g.score += 3 * weight; g.thread++; } else { g.score += weight; g.client++; }
          if (iso > g.last) g.last = iso;
        });
        const ranked = Object.values(groups).sort((a, b) => b.score - a.score).slice(0, 4);
        ranked.forEach(g => {
          g.why = [
            g.thread ? `${g.thread} earlier email${g.thread === 1 ? '' : 's'} in this thread` : '',
            g.client ? `${g.client} filed email${g.client === 1 ? '' : 's'} with this contact` : '',
            g.last ? `last ${nzDate(g.last)}` : '',
          ].filter(Boolean).join(' · ');
        });
        const top = ranked[0];
        const second = ranked[1];
        const clear = top && top.score >= 2 && (!second || top.score >= 2 * second.score);
        cb(clear ? top : null, ranked);
      })
      .catch(() => cb(null, []));
  }

  // Project search by job name, address or client (site titles), for the "choose" card.
  function searchProjects(text) {
    const words = text.split(/\s+/).map(w => w.replace(/[^\w&-]/g, '')).filter(w => w.length >= 2).slice(0, 5);
    if (!words.length) return Promise.resolve([]);
    const kql = words.map(w => `SiteTitle:${w}*`).join(' AND ');
    return searchExact(kql, '&collapsespecification=%27SPSiteURL:1%27').then(rows => {
      const seen = {};
      return rows.map(r => {
        const m = String(r.SiteTitle || '').match(/^\s*HD\s?(\d{3,5})\b/i);
        return { hd: m ? 'HD' + m[1] : '', job: jobFromSiteTitle(r.SiteTitle) };
      })
        .filter(p => p.hd && !seen[p.hd] && (seen[p.hd] = true))
        .slice(0, 8);
    });
  }

  function renderSampleEmail(e) {
    const isMeeting = M.isCalendarItem(e);
    $('sample-email').innerHTML = `
      <h2>${esc(e.subject)}</h2>
      <p class="meta">${esc(e.fromName)} &lt;${esc(e.from)}&gt; · to ${esc((e.to || []).join(', '))} · ${nzDate(e.date)}
        ${isMeeting ? '<span class="badge">Meeting request</span>' : ''}</p>
      <p>${esc(e.body)}</p>`;
  }

  const hasFileLink = () => (state.mode === 'preview' ? !!state.previewLinks : !!fileUrl());
  const hasSearchLink = () => (state.mode === 'preview' ? !!state.previewLinks : !!searchUrl());

  // ---------- "Already filed" marker, kept on the email itself ----------

  function readFiledMarker(cb) {
    if (state.mode === 'preview') { cb(state.previewFiled[state.email.messageId] || null); return; }
    const item = Office.context.mailbox.item;
    if (!item || !item.loadCustomPropertiesAsync) { cb(null); return; }
    item.loadCustomPropertiesAsync(r => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) { cb(null); return; }
      state.props = r.value;
      let rec = null;
      try { rec = JSON.parse(r.value.get('hdgeoFiled') || 'null'); } catch (e) { rec = null; }
      cb(rec);
    });
  }
  function writeFiledMarker(rec) {
    if (state.mode === 'preview') { state.previewFiled[state.email.messageId] = rec; return; }
    if (!state.props) return;
    try { state.props.set('hdgeoFiled', JSON.stringify(rec)); state.props.saveAsync(() => {}); } catch (e) { /* best effort */ }
  }

  // ---------- Filing ----------

  function fileNow(hd, automatic) {
    const token = state.token;
    const email = state.email;
    setCard({ kind: 'filing', hd, automatic });
    sendToFlow(hd, email, (err, where, name) => {
      if (token !== state.token) return;
      if (err) { setCard({ kind: 'error', hd, message: err }); return; }
      const rec = { hd, where, name, at: new Date().toISOString() };
      writeFiledMarker(rec);
      state.filedLog.push({
        messageId: email.messageId, conversationId: email.conversationId, hdNumber: hd,
        folder: where, subject: email.subject, from: email.from, fromName: email.fromName,
        client: (projectByHd(hd) || {}).client || '', date: (email.date || '').slice(0, 10), preview: email.body || '',
      });
      setCard({ kind: 'filed', ...rec, automatic });
    });
  }

  // cb(errorMessage, where, fileName)
  function sendToFlow(hd, email, cb) {
    if (state.mode === 'preview') {
      const p = projectByHd(hd);
      const folder = p && p.folders ? p.folders[0] : 'Emails';
      setTimeout(() => cb(null, folder.replace(/\//g, ' › '), `${email.subject}.eml (preview, not saved)`), 700);
      return;
    }
    const item = Office.context.mailbox.item;
    if (!Office.context.requirements.isSetSupported('Mailbox', '1.14') || !item.getAsFileAsync) {
      cb('This Outlook can\'t hand the email over as a file. Use new Outlook or Outlook on the web.');
      return;
    }
    item.getAsFileAsync(res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        cb('Could not read the email from Outlook. ' + ((res.error && res.error.message) || ''));
        return;
      }
      fetch(fileUrl(), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          hdNumber: hd,
          subject: email.subject,
          receivedDateTime: email.date,
          from: email.from,
          conversationId: email.conversationId,
          filedBy: (Office.context.mailbox.userProfile || {}).emailAddress || '',
          emlBase64: res.value,
        }),
      })
        .then(r => r.text().then(text => ({ ok: r.ok, status: r.status, text })))
        .then(({ ok, status, text }) => {
          if (!ok) { cb(text || `Not filed (error ${status}).`); return; }
          let info = {};
          try { info = JSON.parse(text); } catch (e) { /* plain text */ }
          const where = (info.Path || '').replace(/^\/Shared Documents\//, '').replace(/\/[^/]*$/, '').replace(/\//g, ' › ') || 'Emails';
          cb(null, where, info.Name || '');
        })
        .catch(() => cb('Could not reach the filing flow. Check your connection and the link in Settings.'));
    });
  }

  // ---------- The file card ----------

  function setCard(card) { state.card = card; renderCard(); }

  function optionHtml(project, why) {
    const closed = /completed|cancelled|closed/i.test(project.status || '')
      ? ` · <span class="closed">${esc(project.status)}</span>` : '';
    const sel = project.hdNumber === state.selectedHd ? ' selected' : '';
    const name = project.jobName ? `${project.hdNumber} – ${project.jobName}` : project.hdNumber;
    const sub = [project.client, why].filter(Boolean).join(' · ');
    return `<button type="button" class="option${sel}" data-hd="${esc(project.hdNumber)}">
      <span class="name">${esc(name)}</span>
      <span class="why">${esc(sub)}${closed}</span>
    </button>`;
  }

  function statusHtml(cls, icon, head, sub) {
    return `<div class="status ${cls}"><span class="status-icon">${icon}</span>
      <div class="status-text"><div class="status-head">${head}</div>${sub ? `<div class="status-sub">${sub}</div>` : ''}</div></div>`;
  }

  function chooserHtml(title) {
    const r = state.result || { suggestions: [] };
    const listed = r.suggestions.map(s => optionHtml(s.project, s.reason));
    const picked = state.selectedHd && !r.suggestions.some(s => s.project.hdNumber === state.selectedHd)
      ? projectByHd(state.selectedHd) : null;
    if (picked) listed.unshift(optionHtml(picked, 'Chosen'));
    const sugg = listed.length ? listed.join('') : '<p class="empty">No project found in this email.</p>';
    const several = !r.autoSelect && r.suggestions.length > 1 && !state.selectedHd
      ? '<p class="empty">More than one project matches. Pick the right one.</p>' : '';
    const sel = state.selectedHd;
    const noLink = !hasFileLink();
    return `
      ${title ? `<p class="card-title">${title}</p>` : ''}
      ${several}
      <div class="choices" id="choices">${sugg}</div>
      <label class="label" for="pick-input">Another project</label>
      <input id="pick-input" type="search" placeholder="HD number, job name, address or client" autocomplete="off" />
      <div class="choices" id="pick-results"></div>
      <button id="file-btn" class="btn-primary" type="button" ${!sel || noLink ? 'disabled' : ''}>${sel ? `File to ${esc(sel)}` : 'Choose a project'}</button>
      ${noLink ? '<p class="help">Add the filing flow link in Settings (⚙ at the top) first.</p>' : ''}`;
  }

  function renderCard() {
    const c = state.card;
    const box = $('file-card');
    let html = '';
    switch (c.kind) {
      case 'blocked':
        html = statusHtml('muted', '–', 'Meeting invite', 'Calendar items aren\'t filed to project sites.');
        break;
      case 'looking':
        html = statusHtml('', '<span class="spinner"></span>', 'Looking for the project…',
          'No HD number in the subject. Checking this thread and contact in filed emails.');
        break;
      case 'suggest':
        html = `<div class="suggest">${statusHtml('warn', '?', `File to ${esc(c.hd)}?`,
          `${c.job ? `<strong>${esc(c.job)}</strong><br>` : ''}${esc(c.why || '')}`)}
          <button type="button" id="yes-btn" class="btn-primary" ${hasFileLink() ? '' : 'disabled'}>Yes, file to ${esc(c.hd)}</button>
          <button type="button" id="no-btn" class="btn-secondary">No, choose another project</button>
          ${hasFileLink() ? '' : '<p class="help">Add the filing flow link in Settings (⚙ at the top) first.</p>'}</div>`;
        break;
      case 'filing':
        html = statusHtml('', '<span class="spinner"></span>', `Filing to ${esc(c.hd)}…`,
          c.automatic ? 'HD number found in the subject. Saving the email and attachments.' : 'Saving the email and attachments.');
        break;
      case 'filed':
        html = statusHtml('ok', '✓', `${c.earlier ? 'Filed' : 'Saved'} to ${esc(c.hd)}`,
          `${esc(c.where || 'Emails')}${c.name ? ' · ' + esc(c.name) : ''}${c.earlier && c.at ? ' · ' + nzDate(c.at) : ''}`)
          + (state.chooserOpen ? chooserHtml('File a copy to another project')
            : '<button type="button" id="other-btn" class="link-btn">File a copy to another project</button>');
        break;
      case 'error':
        html = statusHtml('err', '!', `Not filed to ${esc(c.hd)}`, esc(c.message))
          + `<button type="button" id="retry-btn" class="btn-primary">Try again</button>`
          + chooserHtml('Or file to another project');
        break;
      default: {
        const hdCount = state.email ? M.extractHdNumbers(state.email.subject).length : 0;
        const sub = !state.email ? 'Open an email to file it.'
          : hdCount > 1 ? 'This subject has more than one HD number.'
            : (state.result && state.result.suggestions.length) ? 'Check the project and click File.'
              : c.noClue ? 'No HD number and no clear match. Type the HD number, or search by job name, address or client.'
                : 'No HD number in this email. Pick the project.';
        html = statusHtml('warn', '?', 'File this email?', sub) + chooserHtml('');
      }
    }
    box.innerHTML = html;
    wireCard();
  }

  function wireCard() {
    const box = $('file-card');
    box.querySelectorAll('#choices .option').forEach(btn => btn.addEventListener('click', () => {
      state.selectedHd = btn.dataset.hd;
      renderCard();
    }));
    const pick = $('pick-input');
    if (pick) pick.addEventListener('input', renderPickResults);
    const fileBtn = $('file-btn');
    if (fileBtn) fileBtn.addEventListener('click', () => { if (state.selectedHd) { state.chooserOpen = false; fileNow(state.selectedHd, false); } });
    const other = $('other-btn');
    if (other) other.addEventListener('click', () => { state.chooserOpen = true; state.selectedHd = null; renderCard(); });
    const retry = $('retry-btn');
    if (retry) retry.addEventListener('click', () => fileNow(state.card.hd, false));
    const yes = $('yes-btn');
    if (yes) yes.addEventListener('click', () => fileNow(state.card.hd, false));
    const no = $('no-btn');
    if (no) no.addEventListener('click', () => { state.selectedHd = null; setCard({ kind: 'choose' }); const p = $('pick-input'); if (p) p.focus(); });
  }

  function renderPickResults() {
    const q = $('pick-input').value.trim().toLowerCase();
    const out = $('pick-results');
    if (q.length < 2) { out.innerHTML = ''; return; }
    let options;
    if (state.mode === 'outlook') {
      const hds = M.extractHdNumbers(q.replace(/^(\d)/, 'HD$1'));
      if (hds.length) {
        options = hds.map(hd => optionHtml(stubProject(hd), 'Typed HD number'));
      } else if (q.length >= 3 && hasSearchLink()) {
        // Not an HD number: search project sites by job name, address or client.
        out.innerHTML = '<p class="empty">Searching projects…</p>';
        clearTimeout(state.pickTimer);
        const ask = q;
        state.pickTimer = setTimeout(() => {
          searchProjects(ask).then(found => {
            const now = $('pick-input');
            if (!now || now.value.trim().toLowerCase() !== ask) return;
            out.innerHTML = found.length
              ? found.map(p => optionHtml(stubProject(p.hd, p.job), '')).join('')
              : '<p class="empty">No project sites match. Try fewer words or the HD number.</p>';
            wirePickOptions(out);
          }).catch(() => { out.innerHTML = '<p class="empty">Project search failed. Type the HD number instead.</p>'; });
        }, 500);
        return;
      } else {
        options = ['<p class="empty">Type the HD number (e.g. HD4400 or HD4156-1), or a job name, address or client.</p>'];
      }
    } else {
      const matches = state.projects.filter(p =>
        [p.hdNumber, p.jobName, p.client].join(' ').toLowerCase().includes(q)).slice(0, 8);
      options = matches.length ? matches.map(p => optionHtml(p, '')) : ['<p class="empty">No projects match.</p>'];
    }
    out.innerHTML = options.join('');
    wirePickOptions(out);
  }

  function wirePickOptions(out) {
    out.querySelectorAll('.option').forEach(btn => btn.addEventListener('click', () => {
      const name = (btn.querySelector('.name').textContent.split(' – ')[1] || '').trim();
      addStub(btn.dataset.hd, name);
      state.selectedHd = btn.dataset.hd;
      renderCard();
    }));
  }

  // ---------- Ask about filed emails ----------

  function wireAsk() {
    $('ask-form').addEventListener('submit', ev => { ev.preventDefault(); runAsk($('ask-input').value); });
    if (state.mode === 'preview') $('ask-input').addEventListener('input', () => runAsk($('ask-input').value));
  }

  function openLink(url) {
    if (state.mode === 'outlook' && Office.context.ui && Office.context.ui.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url);
    } else {
      window.open(url, '_blank', 'noopener');
    }
  }

  // Turns the SharePoint search reply (from the search flow) into simple rows.
  // File names look like "20260528084726 - Re_ HD3884 - subject": the date comes from the front.
  function rawRows(search) {
    const table = search && search.PrimaryQueryResult && search.PrimaryQueryResult.RelevantResults
      && search.PrimaryQueryResult.RelevantResults.Table;
    return ((table && table.Rows) || []).map(r => {
      const o = {};
      (r.Cells || []).forEach(c => { o[c.Key] = c.Value; });
      return o;
    });
  }

  // Link that opens the file in SharePoint's viewer instead of downloading it:
  // <site>/<library>/Forms/AllItems.aspx?id=<file>&parent=<folder>
  function viewerUrl(path) {
    const m = String(path || '').match(/^(https:\/\/[^/]+)((?:\/(?:sites|teams)\/[^/]+)?)\/([^/]+)\/(.+)$/i);
    if (!m) return path;
    const [, origin, site, lib, rest] = m;
    const file = `${site}/${lib}/${rest}`;
    const parent = file.replace(/\/[^/]*$/, '');
    return `${origin}${site}/${encodeURI(lib)}/Forms/AllItems.aspx?id=${encodeURIComponent(file)}&parent=${encodeURIComponent(parent)}`;
  }

  function searchRows(search) {
    return rawRows(search).map(o => {
      const title = String(o.Title || '');
      const m = title.match(/^(\d{4})(\d{2})(\d{2})\d{6} - (.*)$/);
      const clean = s => String(s || '').replace(/<\/?c0>/g, '').replace(/<ddd\/>/g, '…').replace(/\s+/g, ' ').trim();
      return {
        title: (m ? m[4] : title).replace(/_ /g, ': ').replace(/^RE: /i, 'RE: '),
        path: o.Path || '',
        url: o.Path ? viewerUrl(o.Path) : '',
        site: o.SiteTitle || '',
        date: m ? `${m[1]}-${m[2]}-${m[3]}` : (o.Write || ''),
        summary: clean(o.HitHighlightedSummary).slice(0, 220),
      };
    }).filter(x => x.url);
  }

  function runAsk(text) {
    text = (text || '').trim();
    const status = $('ask-status');
    const list = $('ask-results');
    const chips = $('ask-chips');

    if (state.mode === 'preview') {
      // Preview: rule-based search over the sample filed emails.
      const query = M.parseSearch(text, state.projects, state.filedLog, new Date());
      chips.innerHTML = M.describeSearch(query)
        .map(([k, v]) => `<span class="chip"><span class="k">${esc(k)}</span> ${esc(v)}</span>`).join('');
      const results = M.searchFiled(state.filedLog, state.projects, query).slice(0, text ? 20 : 5);
      status.textContent = text ? `${results.length} matching filed email${results.length === 1 ? '' : 's'} (sample data)` : 'Recently filed (sample data)';
      list.innerHTML = results.map(f => {
        const p = projectByHd(f.hdNumber);
        return `<li><span class="subject">${esc(f.subject)}</span>
          <span class="meta">${esc(f.fromName || f.from)} · ${esc(f.hdNumber)}${p ? ' ' + esc(p.jobName) : ''} · ${nzDate(f.date)}</span>
          <span class="snippet">${esc(f.preview)}</span></li>`;
      }).join('');
      return;
    }

    if (!text) { status.textContent = ''; list.innerHTML = ''; chips.innerHTML = ''; return; }
    if (!hasSearchLink()) { status.textContent = 'Add the search flow link in Settings (⚙ at the top) first.'; return; }

    const btn = document.querySelector('.btn-ask');
    const answerBox = $('ask-answer');
    btn.disabled = true;
    status.innerHTML = '<span class="spinner dark"></span> Searching filed emails and reading the best matches…';
    list.innerHTML = '';
    chips.innerHTML = '';
    answerBox.hidden = true;
    const token = ++state.askToken;

    fetch(searchUrl(), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: text }),
    })
      .then(r => r.text().then(t => ({ ok: r.ok, t })))
      .then(({ ok, t }) => {
        if (!ok) throw new Error(t || 'Search failed.');
        let res = {};
        try { res = JSON.parse(t); } catch (e) { throw new Error('Search reply could not be read.'); }
        const rows = res.results || searchRows(res.search);
        if (rows.length) return { rows, ai: parseAnswer(res.answer) };
        // The AI's search found nothing (or was malformed): retry with a plain loose search.
        return searchExact(looseKql(text)).then(raw => ({
          rows: searchRows({ PrimaryQueryResult: { RelevantResults: { Table: { Rows: raw.map(o => ({ Cells: Object.entries(o).map(([Key, Value]) => ({ Key, Value })) })) } } } }),
          ai: null, loose: true,
        }));
      })
      .then(({ rows, ai, loose }) => {
        if (token !== state.askToken) return;
        // Put the AI's picks first, in its order.
        const picks = ((ai && ai.best) || []).map(p => rows.find(r => r.path === p)).filter(Boolean);
        const rest = rows.filter(r => !picks.includes(r)).slice(0, picks.length ? 8 : 15);
        if (ai && ai.answer) {
          answerBox.innerHTML = `<span class="answer-label">Answer</span>${esc(ai.answer)}`;
          answerBox.hidden = false;
        }
        status.textContent = !rows.length ? 'No filed emails found. Try fewer or different words, a project name or HD number.'
          : loose ? 'Showing looser matches.' : '';
        const item = (x, best) => `<li${best ? ' class="best"' : ''}>
          ${best ? '<span class="best-tag">Best match</span>' : ''}
          <a href="${esc(x.url)}">${esc(x.title)}</a>
          <span class="meta">${esc(x.site || '')}${x.date ? ' · ' + nzDate(x.date) : ''}</span>
          ${x.summary ? `<span class="snippet">${esc(x.summary)}</span>` : ''}</li>`;
        list.innerHTML = picks.map(x => item(x, true)).join('')
          + (picks.length && rest.length ? '<li class="more-head">Other matches</li>' : '')
          + rest.map(x => item(x, false)).join('');
        list.querySelectorAll('a').forEach(a => a.addEventListener('click', ev => { ev.preventDefault(); openLink(a.getAttribute('href')); }));
      })
      .catch(err => {
        if (token !== state.askToken) return;
        status.textContent = (err && err.message && !/fetch/i.test(err.message)) ? err.message
          : 'Could not reach the search flow. Check your connection and the link in Settings.';
      })
      .finally(() => { btn.disabled = false; });
  }

  // The answer step replies with JSON text: {"answer": "...", "best": [paths]}.
  function parseAnswer(text) {
    if (!text) return null;
    const s = String(text).replace(/^```(json)?/i, '').replace(/```\s*$/, '').trim();
    try {
      const o = JSON.parse(s.slice(s.indexOf('{'), s.lastIndexOf('}') + 1));
      return { answer: String(o.answer || ''), best: Array.isArray(o.best) ? o.best.map(String) : [] };
    } catch (e) { return { answer: s, best: [] }; }
  }

  // Loose fallback: any of the words, start-of-word matching, filed emails only.
  function looseKql(text) {
    const stop = /^(the|and|for|from|with|about|email|emails|find|show|me|in|on|of|to|a|an|most|recent|latest|last|newest)$/i;
    const words = text.split(/\s+/).map(w => w.replace(/[^\w-]/g, '')).filter(w => w.length > 2 && !stop.test(w)).slice(0, 6);
    const any = words.map(w => `(${w}* OR SiteTitle:${w}*)`).join(' OR ');
    return `(${any || '*'}) (FileExtension:eml OR FileExtension:msg)`;
  }
})();
