  const FIELDS = [
    { key: 'Name:',                  label: 'Name:',                  comments: false },
    { key: 'Swim Level:',            label: 'Swim Level:',            comments: false },
    { key: 'Instructor Comments:',   label: 'Instructor Comments:',   comments: true  },
    { key: 'Next Time Register In:', label: 'Next Time Register In:', comments: false },
    { key: 'Instructor Name:',       label: 'Instructor Name:',       comments: false },
    { key: 'Date:',                  label: 'Date:',                  comments: false },
  ];

  const FOOTER_URL = 'recreation.utoronto.ca';

  // ── Drag & Drop ──
  const dropZone = document.getElementById('drop-zone');
  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
  dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('dragover'); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); });
  document.getElementById('file-input').addEventListener('change', e => { if (e.target.files[0]) processFile(e.target.files[0]); });

  function processFile(file) {
    document.getElementById('error-box').style.display = 'none';
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb   = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        const ws   = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (!rows.length) { showError('The spreadsheet appears to be empty.'); return; }
        const missing = FIELDS.map(f => f.key).filter(k => !(k in rows[0]));
        if (missing.length) { showError('Missing columns: ' + missing.join(', ') + '. Please check headers match exactly (including colons).'); return; }
        renderCards(rows);
      } catch(err) { showError('Could not read the file. Please make sure it is a valid .xlsx or .xls file.'); }
    };
    reader.readAsArrayBuffer(file);
  }

  function formatDate(val) {
    if (!val) return '';
    if (val instanceof Date) { return val.getFullYear() + '-' + String(val.getMonth()+1).padStart(2,'0') + '-' + String(val.getDate()).padStart(2,'0'); }
    return String(val);
  }

  function escapeHtml(str) { return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  function makeCard(row) {
    const card   = document.createElement('div');
    card.className = 'swim-card';

    const header = document.createElement('div');
    header.className = 'card-header';
    header.innerHTML = '<div class="org">Junior Blues Aquatics</div><div class="card-title">Swim Progress Comment Card</div>';
    card.appendChild(header);

    const body = document.createElement('div');
    body.className = 'card-body';

    FIELDS.forEach(f => {
      const field = document.createElement('div');
      let value   = row[f.key];
      if (f.key === 'Date:') value = formatDate(value);
      else value = (value === null || value === undefined) ? '' : String(value);

      field.className = f.comments ? 'card-field field-comments' : 'card-field inline';
      field.innerHTML = `<div class="f-label">${f.label}</div><div class="f-value">${escapeHtml(value)}</div>`;
      body.appendChild(field);
    });

    card.appendChild(body);

    const footer = document.createElement('div');
    footer.className = 'card-footer';
    footer.innerHTML = `Registration for Spring/Summer 2026 Junior Blues programs will be April 1<sup>st</sup> at 7am.<br>Please use our online registration service at: <a href="https://${FOOTER_URL}">${FOOTER_URL}</a>`;
    card.appendChild(footer);

    return card;
  }

  function renderCards(rows) {
    const container = document.getElementById('cards-container');
    container.innerHTML = '';
    rows.forEach(row => container.appendChild(makeCard(row)));
    document.getElementById('card-count').textContent = rows.length + ' card' + (rows.length !== 1 ? 's' : '');
    document.getElementById('upload-screen').style.display = 'none';
    document.getElementById('cards-screen').style.display = 'block';
  }

  /* ── Font auto-shrink ──────────────────────────────────────────────────
     Measures each comment value box and shrinks font until text fits.
     Uses a binary search between MIN and the computed default size for
     speed, then does a final fine-tune pass to avoid overshooting.
  ──────────────────────────────────────────────────────────────────────── */
  const FONT_MIN  = 0.5;   // rem floor
  const FONT_DEFAULT = 1.0; // rem — reset target before measuring

  function fitAllComments() {
    document.querySelectorAll('.field-comments .f-value').forEach(fitOne);
  }

  function fitOne(el) {
    // 1. Reset to default so we start from a known state
    el.style.fontSize = FONT_DEFAULT + 'rem';

    // 2. If it already fits, nothing to do
    if (el.scrollHeight <= el.clientHeight) return;

    // 3. Binary search for the largest size that fits
    let lo = FONT_MIN, hi = FONT_DEFAULT;
    for (let i = 0; i < 20; i++) {           // 20 iterations → precision ~0.001rem
      const mid = (lo + hi) / 2;
      el.style.fontSize = mid + 'rem';
      if (el.scrollHeight <= el.clientHeight) {
        lo = mid;   // fits — try larger
      } else {
        hi = mid;   // overflows — try smaller
      }
    }
    // Land on lo (the largest confirmed-fitting size)
    el.style.fontSize = lo + 'rem';
  }

  /* ── Print ─────────────────────────────────────────────────────────────
     Strategy:
       1. Restructure DOM into .print-page pairs (each fills one page).
       2. Wait for a full repaint so print-layout heights are resolved.
       3. Run fitAllComments() against the now-correctly-sized print cards.
       4. Call window.print().
       5. After the print dialog closes, restore the flat screen layout
          and re-fit for screen sizes.
  ──────────────────────────────────────────────────────────────────────── */
  function doPrint() {
    const container = document.getElementById('cards-container');
    const cards     = Array.from(container.children);

    // Build print pages
    container.innerHTML = '';
    for (let i = 0; i < cards.length; i += 2) {
      const page = document.createElement('div');
      page.className = 'print-page';
      page.appendChild(cards[i].cloneNode(true));
      if (cards[i + 1]) page.appendChild(cards[i + 1].cloneNode(true));
      container.appendChild(page);
    }

    // Two rAF passes ensure the browser has fully laid out the print DOM
    // before we measure for font-fitting.
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        fitAllComments();
        window.print();
        // Restore flat screen layout after printing
        container.innerHTML = '';
        cards.forEach(c => container.appendChild(c));
      });
    });
  }

  function showError(msg) { const b = document.getElementById('error-box'); b.textContent = '⚠ ' + msg; b.style.display = 'block'; }

  function resetToUpload() {
    document.getElementById('cards-screen').style.display = 'none';
    document.getElementById('upload-screen').style.display = 'flex';
    document.getElementById('file-input').value = '';
    document.getElementById('error-box').style.display = 'none';
    document.getElementById('cards-container').innerHTML = '';
  }