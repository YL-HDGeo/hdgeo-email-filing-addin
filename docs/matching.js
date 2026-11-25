// Project matching rules for the HD Geo email filing add-in.
// Pure functions only: no Outlook or SharePoint calls, so the same rules run
// in the browser preview and inside Outlook.

const HDGEO_DOMAIN = 'hdgeo.co.nz';

// Free/public email domains. A shared domain like gmail.com says nothing
// about which client sent the email, so these never count as a domain match.
const PUBLIC_DOMAINS = [
  'gmail.com', 'googlemail.com', 'outlook.com', 'hotmail.com', 'live.com',
  'yahoo.com', 'yahoo.co.nz', 'icloud.com', 'me.com', 'xtra.co.nz',
  'slingshot.co.nz', 'orcon.net.nz', 'vodafone.co.nz', 'spark.co.nz',
];

// Scoro job statuses that mean the project is finished. Finished projects can
// still be filed to by HD number or thread, but are never guessed from the sender.
const CLOSED_STATUSES = ['completed', 'cancelled', 'closed'];

// Finds HD numbers such as HD4203, HD 4203 and dash jobs such as HD926-5.
// The full dash number is kept: HD926-5 must never become HD926.
function extractHdNumbers(text) {
  if (!text) return [];
  const found = [];
  const re = /\bHD\s?(\d{3,5}(?:-\d{1,3})?)(?!\d)/gi;
  let m;
  while ((m = re.exec(text)) !== null) {
    const hd = 'HD' + m[1];
    if (!found.includes(hd)) found.push(hd);
  }
  return found;
}

function domainOf(address) {
  const at = (address || '').lastIndexOf('@');
  return at === -1 ? '' : address.slice(at + 1).toLowerCase();
}

function isCalendarItem(email) {
  const cls = (email.itemClass || '').toLowerCase();
  return email.itemType === 'appointment'
    || cls.startsWith('ipm.schedule.meeting')
    || cls.startsWith('ipm.appointment');
}

function isActive(project) {
  return !CLOSED_STATUSES.includes((project.status || '').toLowerCase());
}

// The outside people on an email: the sender for incoming mail, the
// recipients for mail sent by HD Geo staff.
function externalAddresses(email) {
  const all = [email.from, ...(email.to || []), ...(email.cc || [])]
    .filter(Boolean)
    .map(a => a.toLowerCase());
  return [...new Set(all.filter(a => domainOf(a) !== HDGEO_DOMAIN))];
}

// Works out which projects an email most likely belongs to.
// Checks, in order: HD number in the subject, same email thread already
// filed, sender is a known client contact, sender's company domain.
// Returns { blocked, alreadyFiled, suggestions, autoSelect }.
function suggestProjects(email, projects, filedLog) {
  if (isCalendarItem(email)) {
    return { blocked: 'calendar', alreadyFiled: null, suggestions: [], autoSelect: null };
  }

  const byHd = new Map(projects.map(p => [p.hdNumber.toUpperCase(), p]));
  const alreadyFiled = filedLog.find(f => f.messageId === email.messageId) || null;
  const suggestions = [];
  const add = (project, reason, strength) => {
    if (project && !suggestions.some(s => s.project.hdNumber === project.hdNumber)) {
      suggestions.push({ project, reason, strength });
    }
  };

  // 1. HD number in the subject
  for (const hd of extractHdNumbers(email.subject)) {
    add(byHd.get(hd.toUpperCase()), 'HD number in the subject', 'strong');
  }

  // 2. Same email thread as an email that has already been filed
  if (email.conversationId) {
    for (const f of filedLog) {
      if (f.conversationId === email.conversationId) {
        add(byHd.get(f.hdNumber.toUpperCase()), 'Same email thread as an email already filed', 'strong');
      }
    }
  }

  // 3 and 4. Sender (or recipients) matched to the client on active projects
  const outside = externalAddresses(email);
  const active = projects.filter(isActive);
  for (const p of active) {
    const contacts = (p.clientContacts || []).map(c => c.toLowerCase());
    if (outside.some(a => contacts.includes(a))) {
      add(p, 'Client contact on this project', 'medium');
    }
  }
  const domains = outside.map(domainOf).filter(d => d && !PUBLIC_DOMAINS.includes(d));
  for (const p of active) {
    const projectDomains = (p.clientContacts || []).map(domainOf);
    if (domains.some(d => projectDomains.includes(d))) {
      add(p, 'Same company as the client on this project', 'weak');
    }
  }

  // Only pre-select when there is one clear answer. If the client has several
  // active projects, the person chooses: the add-in never guesses.
  let autoSelect = null;
  const strong = suggestions.filter(s => s.strength === 'strong');
  if (strong.length === 1) autoSelect = strong[0].project.hdNumber;
  else if (strong.length === 0 && suggestions.length === 1) autoSelect = suggestions[0].project.hdNumber;

  return { blocked: null, alreadyFiled, suggestions, autoSelect };
}

// ---------- Plain-sentence search ----------
// Turns a sentence such as "Sarah from Clientco about retaining wall levels
// last month" into filters: sender, client, HD number, date range, keywords.
// Fixed rules, no AI.

const STOP_WORDS = new Set((
  'a an and any all about are at be by can did do email emails mail mails message ' +
  'messages find for from get had has have he her him his i in is it its me my of on ' +
  'or our re regarding said say says sent send show she that the their them they this ' +
  'to us was we were what when where which who with wrote you your client clients job ' +
  'project projects ltd limited please look looking search').split(' '));

const COMPANY_FILLER = new Set(['ltd', 'limited', 'the', 'nz', 'group', 'and', 'co', 'company', 'inc']);

const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july',
  'august', 'september', 'october', 'november', 'december'];

const iso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const nz = isoDate => isoDate.split('-').reverse().join('/');

// Finds a time period in the sentence. Returns { from, to, label, words } or null.
function parseDateRange(text, today) {
  const t = text.toLowerCase();
  const day = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  const mondayOf = d => addDays(d, -((d.getDay() + 6) % 7));
  const range = (from, to, label, words) => ({ from: iso(from), to: iso(to), label, words });
  let m;

  if (/\btoday\b/.test(t)) return range(day, day, 'Today', ['today']);
  if (/\byesterday\b/.test(t)) return range(addDays(day, -1), addDays(day, -1), 'Yesterday', ['yesterday']);
  if (/\bthis week\b/.test(t)) return range(mondayOf(day), day, 'This week', ['this', 'week']);
  if (/\blast week\b/.test(t)) {
    const mon = addDays(mondayOf(day), -7);
    return range(mon, addDays(mon, 6), 'Last week', ['last', 'week']);
  }
  if (/\bthis month\b/.test(t)) {
    return range(new Date(day.getFullYear(), day.getMonth(), 1), day, 'This month', ['this', 'month']);
  }
  if (/\blast month\b/.test(t)) {
    const from = new Date(day.getFullYear(), day.getMonth() - 1, 1);
    const to = new Date(day.getFullYear(), day.getMonth(), 0);
    return range(from, to, 'Last month', ['last', 'month']);
  }
  if ((m = t.match(/\b(?:last|past) (\d{1,3}) (day|week|month)s?\b/))) {
    const n = Number(m[1]);
    const from = m[2] === 'day' ? addDays(day, -n)
      : m[2] === 'week' ? addDays(day, -7 * n)
        : new Date(day.getFullYear(), day.getMonth() - n, day.getDate());
    return range(from, day, `Last ${n} ${m[2]}${n === 1 ? '' : 's'}`, ['last', 'past', m[1], m[2], m[2] + 's']);
  }
  for (let i = 0; i < 12; i++) {
    const name = MONTHS[i];
    if (new RegExp(`\\b(${name}|${name.slice(0, 3)})\\b`).test(t) && !(name === 'may' && !/\bin may\b/.test(t))) {
      // The most recent past occurrence of that month.
      const year = i > day.getMonth() ? day.getFullYear() - 1 : day.getFullYear();
      const label = name[0].toUpperCase() + name.slice(1) + ' ' + year;
      return range(new Date(year, i, 1), new Date(year, i + 1, 0), label, [name, name.slice(0, 3)]);
    }
  }
  return null;
}

// Light word stem so "levels" finds "level" and "walls" finds "wall".
const stem = w => (w.length > 4 && w.endsWith('s') && !w.endsWith('ss') ? w.slice(0, -1) : w);
const wordsOf = s => (s || '').toLowerCase().match(/[a-z0-9]+(?:-[0-9]+)?/g) || [];

function parseSearch(text, projects, filedLog, today) {
  const query = { text, hdNumbers: [], client: null, sender: null, dates: null, keywords: [] };
  if (!text || !text.trim()) return query;
  const used = new Set();

  // HD numbers
  query.hdNumbers = extractHdNumbers(text);
  query.hdNumbers.forEach(hd => used.add(hd.toLowerCase()));

  const words = wordsOf(text);

  // Client: any distinctive word of a client's name, e.g. "clientco" or "waikato builders"
  const clients = [...new Set(projects.map(p => p.client).filter(Boolean))];
  let best = null;
  for (const c of clients) {
    const nameWords = wordsOf(c).filter(w => !COMPANY_FILLER.has(w));
    const hits = nameWords.filter(w => words.includes(w));
    if (hits.length && (!best || hits.length > best.hits.length)) best = { client: c, hits };
  }
  if (best) { query.client = best.client; best.hits.forEach(w => used.add(w)); }

  // Sender: a first name, surname or email name of someone who appears in filed emails
  const people = new Map();
  filedLog.forEach(f => {
    const keys = wordsOf(f.fromName).concat(wordsOf((f.from || '').split('@')[0]));
    keys.forEach(k => { if (k.length > 2 && !people.has(k)) people.set(k, f.fromName || f.from); });
  });
  for (const w of words) {
    if (people.has(w) && !used.has(w)) { query.sender = { word: w, name: people.get(w) }; used.add(w); break; }
  }

  // Date range
  query.dates = parseDateRange(text, today);
  if (query.dates) query.dates.words.forEach(w => used.add(w));

  // Whatever is left are keywords
  query.keywords = [...new Set(words.filter(w => !used.has(w) && !STOP_WORDS.has(w) && w.length > 1))];
  return query;
}

// Runs a parsed search. Client, sender, HD number and dates must all match.
// Keywords rank the results: an email must match at least one keyword, and
// emails matching more keywords come first.
function searchFiled(filedLog, projects, query) {
  const byHd = new Map(projects.map(p => [p.hdNumber.toUpperCase(), p]));
  const kws = query.keywords.map(stem);
  return filedLog
    .map(f => {
      const p = byHd.get((f.hdNumber || '').toUpperCase());
      const client = f.client || (p && p.client) || '';
      if (query.client && client !== query.client) return null;
      if (query.hdNumbers.length && !query.hdNumbers.some(hd => (f.hdNumber || '').toUpperCase() === hd.toUpperCase()
        || (f.hdNumber || '').toUpperCase().startsWith(hd.toUpperCase() + '-'))) return null;
      if (query.sender) {
        const who = wordsOf(f.fromName).concat(wordsOf((f.from || '').split('@')[0]));
        if (!who.includes(query.sender.word)) return null;
      }
      const date = (f.date || '').slice(0, 10);
      if (query.dates && (date < query.dates.from || date > query.dates.to)) return null;

      const hay = new Set(wordsOf([f.subject, f.preview, p && p.jobName].join(' ')).map(stem));
      const score = kws.filter(k => hay.has(k)).length;
      if (kws.length && score === 0) return null;
      return { ...f, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score || (b.date || '').localeCompare(a.date || ''));
}

// Readable summary of how the sentence was understood, for the chips under the box.
function describeSearch(q) {
  const parts = [];
  if (q.sender) parts.push(['From', q.sender.name]);
  if (q.client) parts.push(['Client', q.client]);
  if (q.hdNumbers.length) parts.push(['Project', q.hdNumbers.join(', ')]);
  if (q.dates) parts.push(['When', `${q.dates.label} (${nz(q.dates.from)}–${nz(q.dates.to)})`]);
  if (q.keywords.length) parts.push(['Words', q.keywords.join(', ')]);
  return parts;
}

// Exposed for the panel script and for Node-free tests in the browser.
window.Matching = {
  extractHdNumbers, suggestProjects, parseSearch, searchFiled, describeSearch, isCalendarItem, domainOf,
};
