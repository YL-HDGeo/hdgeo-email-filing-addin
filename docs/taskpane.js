// HD Geo email filing panel.
// Runs in two modes:
//  - Preview: a normal browser, using sample emails and projects.
//  - Outlook: reads the open email from Outlook. Projects and saving still use
//    sample data until the app registration (Microsoft Graph access) is set up.

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
  };

  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nzDate = iso => {
    if (!iso) return '';
    const [y, m, d] = iso.slice(0, 10).split('-');
    return `${d}/${m}/${y}`;
  };
  const projectByHd = hd => state.projects.find(p => p.hdNumber === hd);

  // ---------- Start-up ----------

  function start(inOutlook) {
    state.mode = inOutlook ? 'outlook' : 'preview';
    wireTabs();
    wireFileView();
    wireSearchView();

    if (inOutlook) {
      loadOutlookEmail();
      // Keep working when the pane is pinned and the person opens another email.
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadOutlookEmail);
    } else {
      startPreview();
    }
  }

  // Preview only: mock Outlook window with an inbox, reading pane and ribbon button.
  function startPreview() {
    document.body.classList.add('preview-mode');
    ['ribbon', 'mail-list', 'reading', 'host-head'].forEach(id => { $(id).hidden = false; });

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
    $('ribbon-file').addEventListener('click', () =>
      setPane(document.body.classList.contains('pane-closed')));
    $('host-close').addEventListener('click', () => setPane(false));

    list.querySelector('.mail-item').click();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(info => start(info.host === Office.HostType.Outlook));
  } else {
    // office.js could not load (offline preview): still run the preview.
    document.addEventListener('DOMContentLoaded', () => start(false));
  }

  // ---------- Reading the email ----------

  function loadOutlookEmail() {
    const item = Office.context.mailbox.item;
    if (!item) return;
    const addr = a => (a && a.emailAddress) || '';
    const email = {
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
    };
    showEmail(email);
  }

  function showEmail(email) {
    state.email = email;
    if (state.mode === 'preview') renderSampleEmail(email);
    state.result = M.suggestProjects(email, state.projects, state.filedLog);
    state.selectedHd = state.result.autoSelect;
    $('project-search').value = '';
    $('project-results').innerHTML = '';
    renderFileView();
  }

  function renderSampleEmail(e) {
    const isMeeting = M.isCalendarItem(e);
    $('sample-email').innerHTML = `
      <h2>${esc(e.subject)}</h2>
      <p class="meta">${esc(e.fromName)} &lt;${esc(e.from)}&gt; · to ${esc((e.to || []).join(', '))} · ${nzDate(e.date)}
        ${isMeeting ? '<span class="badge">Meeting request</span>' : ''}</p>
      <p>${esc(e.body)}</p>`;
  }

  // ---------- File view ----------

  function wireFileView() {
    $('project-search').addEventListener('input', renderProjectSearch);
    $('file-btn').addEventListener('click', fileEmail);
  }

  function optionHtml(project, why) {
    const closed = /completed|cancelled|closed/i.test(project.status)
      ? ` · <span class="closed">${esc(project.status)}</span>` : '';
    const sel = project.hdNumber === state.selectedHd ? ' selected' : '';
    return `<button type="button" class="option${sel}" data-hd="${esc(project.hdNumber)}">
      <span class="name">${esc(project.hdNumber)} – ${esc(project.jobName)}</span>
      <span class="why">${esc(project.client)}${why ? ' · ' + esc(why) : ''}${closed}</span>
    </button>`;
  }

  function wireOptions(container) {
    container.querySelectorAll('.option').forEach(btn => btn.addEventListener('click', () => {
      state.selectedHd = btn.dataset.hd;
      renderFileView();
      renderProjectSearch();
    }));
  }

  function renderFileView() {
    const r = state.result;
    const notice = $('file-notice');
    const form = $('file-form');

    if (r.blocked === 'calendar') {
      notice.innerHTML = '<div class="notice warn">This is a meeting invite or calendar reply. These aren\'t filed to project sites.</div>';
      form.hidden = true;
      return;
    }
    form.hidden = false;
    notice.innerHTML = r.alreadyFiled
      ? `<div class="notice ok">Already filed to ${esc(r.alreadyFiled.hdNumber)} › ${esc(r.alreadyFiled.folder)}. Filing again saves another copy.</div>`
      : '';

    const box = $('suggestions');
    if (r.suggestions.length) {
      box.innerHTML = r.suggestions.map(s => optionHtml(s.project, s.reason)).join('');
      if (!r.autoSelect && r.suggestions.length > 1 && !state.selectedHd) {
        box.insertAdjacentHTML('afterbegin', '<p class="empty">This client has more than one active project. Pick the right one.</p>');
      }
    } else {
      box.innerHTML = '<p class="empty">No match found. Search for the project below.</p>';
    }
    wireOptions(box);

    // Folder list for the chosen project. Dash jobs default to their variation folder.
    const folderSelect = $('folder-select');
    folderSelect.innerHTML = '';
    const project = projectByHd(state.selectedHd);
    if (project) {
      project.folders.forEach(f => folderSelect.add(new Option(f, f)));
      folderSelect.disabled = false;
    } else {
      folderSelect.add(new Option('Choose a project first', ''));
      folderSelect.disabled = true;
    }

    const btn = $('file-btn');
    btn.disabled = !project;
    btn.textContent = project ? `File to ${project.hdNumber}` : 'File email';
    $('file-help').textContent = project
      ? `Saves to ${project.siteUrl.replace('https://', '')} › ${folderSelect.value}`
      : '';
  }

  function renderProjectSearch() {
    const q = $('project-search').value.trim().toLowerCase();
    const out = $('project-results');
    if (q.length < 2) { out.innerHTML = ''; return; }
    const matches = state.projects.filter(p =>
      [p.hdNumber, p.jobName, p.client].join(' ').toLowerCase().includes(q)).slice(0, 20);
    out.innerHTML = matches.length
      ? matches.map(p => optionHtml(p, '')).join('')
      : '<p class="empty">No projects match.</p>';
    wireOptions(out);
  }

  function fileEmail() {
    const project = projectByHd(state.selectedHd);
    if (!project) return;
    const folder = $('folder-select').value;
    const e = state.email;

    if (state.mode === 'outlook') {
      // Saving needs Microsoft Graph access, which comes with the app registration.
      $('file-help').textContent = 'Saving to SharePoint switches on once the app registration is set up. Nothing was saved.';
      return;
    }

    // Preview: record it in the in-memory filed list so thread matching and search pick it up.
    state.filedLog.push({
      messageId: e.messageId, conversationId: e.conversationId, hdNumber: project.hdNumber, folder,
      subject: e.subject, from: e.from, fromName: e.fromName, client: project.client,
      date: (e.date || '').slice(0, 10), preview: e.body,
    });
    state.result = M.suggestProjects(e, state.projects, state.filedLog);
    renderFileView();
    $('file-notice').innerHTML = `<div class="notice ok">Filed to ${esc(project.hdNumber)} › ${esc(folder)} (preview only, nothing saved to SharePoint).</div>`;
    runSearch();
  }

  // ---------- Search view ----------

  function wireTabs() {
    const show = which => {
      ['file', 'search'].forEach(v => {
        $('view-' + v).hidden = v !== which;
        $('tab-' + v).classList.toggle('active', v === which);
        $('tab-' + v).setAttribute('aria-selected', v === which);
      });
      if (which === 'search') runSearch();
    };
    $('tab-file').addEventListener('click', () => show('file'));
    $('tab-search').addEventListener('click', () => show('search'));
  }

  function wireSearchView() {
    $('keyword-input').addEventListener('input', runSearch);
  }

  function runSearch() {
    const text = $('keyword-input').value;
    const query = M.parseSearch(text, state.projects, state.filedLog, new Date());
    $('search-chips').innerHTML = M.describeSearch(query)
      .map(([k, v]) => `<span class="chip"><span class="k">${esc(k)}</span> ${esc(v)}</span>`).join('');
    const results = text.trim()
      ? M.searchFiled(state.filedLog, state.projects, query)
      : M.searchFiled(state.filedLog, state.projects, query).slice(0, 10);
    $('search-count').textContent = text.trim()
      ? `${results.length} matching filed email${results.length === 1 ? '' : 's'}`
      : 'Most recent filed emails';
    $('search-results').innerHTML = results.map(f => {
      const p = projectByHd(f.hdNumber);
      return `<li>
        <span class="subject">${esc(f.subject)}</span>
        <span class="meta">${esc(f.fromName || f.from)} · ${esc(f.hdNumber)}${p ? ' ' + esc(p.jobName) : ''} · ${nzDate(f.date)}</span>
        <span class="snippet">${esc(f.preview)}</span>
      </li>`;
    }).join('');
  }
})();
