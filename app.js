// Renders the CV from cv-data.json. Edit that file + redeploy to update — that is
// the whole "easy to update" story. No build step, no framework.

const ICON = {
  email: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/>',
  orcid: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><circle cx="12" cy="8.2" r=".5" fill="currentColor"/>',
  rg: '<circle cx="12" cy="12" r="9"/><path d="M9.5 16V9H12a2 2 0 0 1 0 4H9.5m2.6 0 2.4 3"/>',
  link: '<path d="M9.5 14.5 14.5 9.5"/><path d="M11 6.5 12.5 5a4 4 0 0 1 5.6 5.6L16.5 12"/><path d="M13 17.5 11.5 19a4 4 0 0 1-5.6-5.6L7.5 12"/>',
  pdf: '<path d="M14 3v5h5"/><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M9 13h6M9 16h4"/>',
  preprint: '<path d="M14 3v5h5"/><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/>',
  scholar: '<circle cx="11" cy="11" r="6"/><path d="m20 20-3.6-3.6"/>',
};

const svg = (name) =>
  `<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">${ICON[name]}</svg>`;

const esc = (s) =>
  String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function boldName(authors, name) {
  const safe = esc(authors);
  if (!name) return safe;
  return safe.split(esc(name)).join(`<b>${esc(name)}</b>`);
}

function safeUrl(value) {
  try {
    if (typeof value !== 'string' || /[\s\\]/.test(value)) return '';
    const url = new URL(value);
    return ['https:', 'http:'].includes(url.protocol) && !url.username && !url.password ? url.href : '';
  } catch { return ''; }
}

function safeAssetPath(value) {
  return typeof value === 'string' &&
    /^[A-Za-z0-9._/-]+$/.test(value) &&
    !value.includes('..') &&
    !value.startsWith('/')
    ? value
    : '';
}

function doiUrl(doi) {
  return /^10\.\d{4,9}\//.test(doi) ? safeUrl(`https://doi.org/${doi}`) : safeUrl(doi);
}

function pubLinks(pub) {
  const out = [];
  const L = pub.links || {};
  if (L.doi) out.push(['link', doiUrl(L.doi), 'DOI / article']);
  if (L.pdf) out.push(['pdf', L.pdf, 'PDF']);
  if (L.preprint) out.push(['preprint', L.preprint, 'Preprint']);
  // Always-available fallback so every entry has a working "find it" action.
  out.push(['scholar', `https://scholar.google.com/scholar?q=${encodeURIComponent(pub.title)}`, 'Find on Google Scholar']);
  return `<span class="links">${out
    .filter(([, href]) => safeUrl(href))
    .map(([ic, href, t]) => `<a class="iconlink" href="${esc(safeUrl(href))}" target="_blank" rel="noopener noreferrer" title="${esc(t)}" aria-label="${esc(t)}">${svg(ic)}</a>`)
    .join('')}</span>`;
}

function pubItem(pub, i) {
  const info = pub.info ? `, ${esc(pub.info)}` : '';
  const eq = pub.equalContrib ? ' <span class="meta">#&nbsp;equal contribution.</span>' : '';
  const badge = pub.metrics ? ` <span class="badge">${esc(pub.metrics)}</span>` : '';
  return `<div class="pub"><span class="num">${i}.</span><span class="body">
    <span class="authors">${boldName(pub.authors, window.__HL)}</span> (${esc(pub.year)}).
    <span class="title">${esc(pub.title)}</span>.
    <span class="venue">${esc(pub.venue)}</span>${info}.${eq}${badge}${pubLinks(pub)}
  </span></div>`;
}

function section(title, inner) {
  return `<section><h2 class="sec-title">${title}</h2>${inner}</section>`;
}

function render(cv) {
  window.__HL = cv.highlightName;
  document.title = cv.name ? `${cv.name} - CV` : 'Curriculum Vitae';
  document.querySelector('meta[name="description"]').content = [cv.name, cv.title, cv.affiliation].filter(Boolean).join(' - ');
  const c = [];

  const chips = [];
  if (cv.email) chips.push(`<a class="chip" href="mailto:${esc(cv.email)}">${svg('email')} ${esc(cv.email)}</a>`);
  if (safeUrl(cv.orcid)) chips.push(`<a class="chip" href="${esc(safeUrl(cv.orcid))}" target="_blank" rel="noopener noreferrer">${svg('orcid')} ORCID</a>`);
  if (safeUrl(cv.researchgate)) chips.push(`<a class="chip" href="${esc(safeUrl(cv.researchgate))}" target="_blank" rel="noopener noreferrer">${svg('rg')} ResearchGate</a>`);
  const pdfHref = safeAssetPath(cv.site?.pdf);
  if (pdfHref) chips.push(`<a class="chip" href="${esc(pdfHref)}" target="_blank" rel="noopener noreferrer">${svg('pdf')} CV PDF</a>`);

  c.push(`<header class="head">
    <h1 class="name">${esc(cv.name)}${cv.nameZh ? `<span class="zh">${esc(cv.nameZh)}</span>` : ''}</h1>
    ${cv.title ? `<div class="subtitle">${esc(cv.title)}</div>` : ''}
    ${cv.affiliation ? `<div class="affil">${esc(cv.affiliation)}</div>` : ''}
    ${cv.location ? `<div class="loc">${esc(cv.location)}</div>` : ''}
    <div class="contacts">${chips.join('')}</div>
  </header>`);

  // Education
  c.push(section('Education', cv.education.map((e) => `<div class="entry">
    <div class="when">${esc(e.period)}</div>
    <div><div class="deg">${esc(e.degree)}</div>
      <div class="org">${esc(e.institution)}${e.location ? ` <span class="place">· ${esc(e.location)}</span>` : ''}</div>
      ${e.detail ? `<div class="detail">${esc(e.detail)}</div>` : ''}
      ${e.supervisor ? `<div class="detail">Supervisor: ${esc(e.supervisor)}</div>` : ''}</div>
  </div>`).join('')));

  // Research interests
  if (cv.interests?.length) c.push(section('Research Interests',
    `<div class="tags">${cv.interests.map((t) => `<span class="tag">${esc(t)}</span>`).join('')}</div>`));

  // Skills
  if (cv.skills?.length) c.push(section('Skills',
    `<ul class="bullets">${cv.skills.map((s) => `<li>${esc(s)}</li>`).join('')}</ul>`));

  // Languages
  if (cv.languages?.length) c.push(section('Languages',
    cv.languages.map((l) => `<div class="lang-row"><b>${esc(l.name)}</b><span class="lvl">${esc(l.level)}${l.note ? ` · ${esc(l.note)}` : ''}</span></div>`).join('')));

  // Awards
  if (cv.awards?.length) c.push(section('Awards &amp; Honours',
    cv.awards.map((a) => `<div class="award"><span class="yr">${esc(a.year)}</span><span>${esc(a.text)}</span></div>`).join('')));

  // Publications
  const pubs = cv.publications || {};
  let pubHtml = '';
  if (pubs.lead?.length) pubHtml += `<div class="pub-group"><h3>Lead-author publications</h3>${pubs.lead.map((p, i) => pubItem(p, i + 1)).join('')}</div>`;
  if (pubs.collaborative?.length) pubHtml += `<div class="pub-group"><h3>Collaborative publications</h3>${pubs.collaborative.map((p, i) => pubItem(p, i + 1)).join('')}</div>`;
  if (pubHtml) c.push(section('Publications', pubHtml));

  // Footer with share QR
  if (safeUrl(cv.site?.url)) c.push(`<div class="foot"><img src="qr.png" alt="QR code to this CV" /><span>${esc(cv.site.url)}</span></div>`);

  document.getElementById('cv').innerHTML = c.join('');
}

document.getElementById('export').addEventListener('click', () => window.print());

fetch('cv-data.json')
  .then((r) => r.json())
  .then(render)
  .catch((e) => {
    document.getElementById('cv').innerHTML = `<p class="loading">Could not load cv-data.json (${esc(e.message)}). Serve this folder over http.</p>`;
  });
