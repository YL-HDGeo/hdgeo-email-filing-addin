// HD Geo email filing panel.
// Runs in two modes:
//  - Preview: a normal browser, using sample emails and projects.
//  - Outlook: reads the open email from Outlook and files it through the
//    "Email filing - File email" Power Automate flow. The flow link is saved
//    in the user's Outlook roaming settings (never in this code).

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
  // Outlook mode has no project list yet: an HD number stands in for the project,
  // and the flow checks it against the Project Register.
  const stubProject = hd => ({ hdNumber: hd, jobName: '', client: '', status: '', siteUrl: '', folders: [] });
  const addStub = hd => { if (!projectByHd(hd)) state.projects.push(stubProject(hd)); };

  // ---------- Settings (flow link) ----------

  const SETTING_KEY = 'fileFlowUrl';
  function getFlowUrl() {
    try {
      if (state.mode === 'outlook') return Office.context.roamingSettings.get(SETTING_KEY) || '';
      return localStorage.getItem(SETTING_KEY) || '';
    } catch (e) { return ''; }
  }
  function saveFlowUrl(url, done) {
    if (state.mode === 'outlook') {
      if (url) Office.context.roamingSettings.set(SETTING_KEY, url);
      else Office.context.roamingSettings.remove(SETTING_KEY);
      Office.context.roamingSettings.saveAsync(r => done(r.status === Office.AsyncResultStatus.Succeeded));
    } else {
      try { url ? localStorage.setItem(SETTING_KEY, url) : localStorage.removeItem(SETTING_KEY); done(true); }
      catch (e) { done(false); }
    }
  }
  function wireSettings() {
    $('save-settings').addEventListener('click', () => {
      const url = $('flow-url').value.trim();
      if (url && !/^https:\/\/[^/]+\.(powerplatform\.com|logic\.azure\.com)(:\d+)?\//i.test(url)) {
        $('settings-help').textContent = 'That doesn\'t look like a Power Automate flow link. Copy the HTTP URL from the flow\'s trigger.';
        return;
      }
      saveFlowUrl(url, ok => {
        $('settings-help').textContent = ok
          ? (url ? 'Saved. You can now file emails.' : 'Link removed.')
          : 'Could not save the setting. Try again.';
        if (state.result) renderFileView();
      });
    });
  }

  // ---------- Start-up ----------

  function start(inOutlook) {
    state.mode = inOutlook ? 'outlook' : 'preview';
    wireTabs();
    wireFileView();
    wireSearchView();
    wireSettings();

    if (inOutlook) {
      state.projects = [];
      state.filedLog = [];
      $('tab-settings').hidden = false;
      $('folder-block').hidden = true;
      $('flow-url').value = getFlowUrl();
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
    else M.extractHdNumbers(email.subject).forEach(addStub);
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
    const name = project.jobName ? `${project.hdNumber} – ${project.jobName}` : project.hdNumber;
    const sub = [project.client, why].filter(Boolean).join(' · ');
    return `<button type="button" class="option${sel}" data-hd="${esc(project.hdNumber)}">
      <span class="name">${esc(name)}</span>
      <span class="why">${esc(sub)}${closed}</span>
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
    if (state.mode === 'outlook') {
      const hasLink = !!getFlowUrl();
      btn.disabled = !project || !hasLink;
      $('file-help').textContent = !hasLink
        ? 'Add the filing flow link in Settings first.'
        : project
          ? (/-\d+$/.test(project.hdNumber)
            ? `Saves to the ${project.hdNumber} variation folder › 01 Emails on the project site (or Emails if there's no variation folder).`
            : `Saves to the ${project.hdNumber} project site › Emails.`)
          : '';
      return;
    }
    $('file-help').textContent = project
      ? `Saves to ${project.siteUrl.replace('https://', '')} › ${folderSelect.value}`
      : '';
  }

  function renderProjectSearch() {
    const q = $('project-search').value.trim().toLowerCase();
    const out = $('project-results');
    if (q.length < 2) { out.innerHTML = ''; return; }
    if (state.mode === 'outlook') {
      // No project list in Outlook yet: accept a typed HD number.
      const hds = M.extractHdNumbers(q.replace(/^(\d)/, 'HD$1'));
      out.innerHTML = hds.length
        ? hds.map(hd => optionHtml(stubProject(hd), 'Typed HD number')).join('')
        : '<p class="empty">Type the HD number, e.g. HD4156 or HD4156-1.</p>';
      out.querySelectorAll('.option').forEach(btn => btn.addEventListener('click', () => {
        addStub(btn.dataset.hd);
        state.selectedHd = btn.dataset.hd;
        renderFileView();
        renderProjectSearch();
      }));
      return;
    }
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
      fileFromOutlook(project.hdNumber);
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

  // Outlook: get the whole email (with attachments) as an .eml file and send it to the flow.
  function fileFromOutlook(hd) {
    const item = Office.context.mailbox.item;
    const btn = $('file-btn');
    const say = (cls, msg) => { $('file-notice').innerHTML = `<div class="notice ${cls}">${msg}</div>`; };
    const flowUrl = getFlowUrl();
    if (!flowUrl) { say('warn', 'Add the filing flow link in Settings first.'); return; }
    if (!Office.context.requirements.isSetSupported('Mailbox', '1.14') || !item.getAsFileAsync) {
      say('err', 'This version of Outlook can\'t hand the email over as a file. Use new Outlook or Outlook on the web.');
      return;
    }
    btn.disabled = true;
    btn.textContent = 'Filing…';
    const email = state.email;
    item.getAsFileAsync(res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        say('err', 'Could not read the email from Outlook: ' + esc(res.error && res.error.message));
        renderButtonOnly(hd);
        return;
      }
      fetch(flowUrl, {
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
          if (ok) {
            let info = {};
            try { info = JSON.parse(text); } catch (e) { /* plain text reply */ }
            const where = (info.Path || '').replace(/^\/Shared Documents\//, '').replace(/\/[^/]*$/, '').replace(/\//g, ' › ');
            say('ok', `Filed to ${esc(hd)}${where ? ' › ' + esc(where) : ''}.<br><small>${esc(info.Name || '')}</small>`);
            state.filedLog.push({
              messageId: email.messageId, conversationId: email.conversationId, hdNumber: hd,
              folder: where || 'Emails', subject: email.subject, from: email.from, fromName: email.fromName,
              client: '', date: (email.date || '').slice(0, 10), preview: '',
            });
          } else {
            say('err', esc(text || `Not filed (error ${status}).`));
          }
        })
        .catch(() => say('err', 'Could not reach the filing flow. Check your connection and the link in Settings.'))
        .finally(() => renderButtonOnly(hd));
    });
  }

  function renderButtonOnly(hd) {
    const btn = $('file-btn');
    btn.disabled = false;
    btn.textContent = `File to ${hd}`;
  }

  // ---------- Search view ----------

  function wireTabs() {
    const show = which => {
      ['file', 'search', 'settings'].forEach(v => {
        $('view-' + v).hidden = v !== which;
        $('tab-' + v).classList.toggle('active', v === which);
        $('tab-' + v).setAttribute('aria-selected', v === which);
      });
      if (which === 'search') runSearch();
    };
    $('tab-file').addEventListener('click', () => show('file'));
    $('tab-search').addEventListener('click', () => show('search'));
    $('tab-settings').addEventListener('click', () => show('settings'));
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
