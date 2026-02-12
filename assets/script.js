
    // ========== FAST XLSX PARSER USING JSZIP ==========
    // This actually works with real XLSX files

    // ----- THEME -----
    const root = document.documentElement;
    const metaThemeColor = document.getElementById('theme-color-meta');
    const lightBtn = document.getElementById('themeLightBtn');
    const darkBtn = document.getElementById('themeDarkBtn');

    function setTheme(theme, save = true) {
      if (theme === 'dark') {
        root.setAttribute('data-theme', 'dark');
        metaThemeColor.setAttribute('content', '#0a0c10');
        darkBtn.classList.add('active');
        lightBtn.classList.remove('active');
      } else {
        root.setAttribute('data-theme', 'light');
        metaThemeColor.setAttribute('content', '#2563eb');
        lightBtn.classList.add('active');
        darkBtn.classList.remove('active');
      }
      if (save) localStorage.setItem('csv-theme', theme);
    }

    const saved = localStorage.getItem('csv-theme');
    if (saved) setTheme(saved, false);
    else setTheme(window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light', false);

    lightBtn.onclick = () => setTheme('light');
    darkBtn.onclick = () => setTheme('dark');

    // ----- ONBOARDING -----
    const guide = document.getElementById('onboardingGuide');
    if (localStorage.getItem('guideDismissed') === 'true') guide.style.display = 'none';

    document.getElementById('dismissGuideBtn').onclick = () => {
      guide.style.display = 'none';
      localStorage.setItem('guideDismissed', 'true');
    };

    // ----- TOAST -----
    const toast = document.getElementById('feedbackToast');
    function showToast(msg, duration = 2000) {
      toast.textContent = msg;
      toast.classList.add('show');
      setTimeout(() => toast.classList.remove('show'), duration);
    }

    // ----- CORE VARIABLES -----
    const fileInput = document.getElementById('fileInput');
    const uploadBox = document.getElementById('uploadBox');
    const tableWrap = document.getElementById('tableWrap');
    const searchInput = document.getElementById('search');
    const controls = document.getElementById('controls');
    let tableData = [];
    let currentFilter = '';

    // ========== PROPER XLSX PARSER ==========
    async function parseXLSX(arrayBuffer) {
      try {
        // Load ZIP with JSZip
        const zip = await JSZip.loadAsync(arrayBuffer);

        // Find shared strings (for cell values)
        let sharedStrings = [];
        const sharedStringFile = zip.file(/xl\/sharedStrings\.xml/);
        if (sharedStringFile.length) {
          const content = await sharedStringFile[0].async('text');
          const matches = content.match(/<t[^>]*>([^<]+)<\/t>/g);
          if (matches) {
            sharedStrings = matches.map(m => m.replace(/<\/?t[^>]*>/g, ''));
          }
        }

        // Get the first sheet
        const sheetFile = zip.file(/xl\/worksheets\/sheet\d+\.xml/);
        if (!sheetFile.length) {
          throw new Error('No sheets found');
        }

        const sheetContent = await sheetFile[0].async('text');

        // Parse rows
        const rows = [];
        const rowMatches = sheetContent.match(/<row[^>]*>.*?<\/row>/g) || [];

        for (let rowXml of rowMatches) {
          const cells = [];
          const cellMatches = rowXml.match(/<c[^>]*>(.*?)<\/c>/g) || [];

          for (let cellXml of cellMatches) {
            let value = '';

            // Check if it's a shared string (t="s")
            if (cellXml.includes('t="s"')) {
              const vMatch = cellXml.match(/<v>([^<]+)<\/v>/);
              if (vMatch && vMatch[1]) {
                const index = parseInt(vMatch[1]);
                value = sharedStrings[index] || '';
              }
            } else {
              // Inline value
              const vMatch = cellXml.match(/<v>([^<]+)<\/v>/);
              if (vMatch) value = vMatch[1];
            }
            cells.push(value);
          }

          if (cells.length > 0) {
            rows.push(cells);
          }
        }

        return rows.length > 0 ? rows : [['Empty XLSX file']];
      } catch (err) {
        console.error('XLSX parse error:', err);
        throw err;
      }
    }

    // ----- FILE LOADING -----
    async function loadFile(file) {
      const ext = file.name.split('.').pop().toLowerCase();

      try {
        if (ext === 'xlsx') {
          showToast('📊 Reading XLSX...', 1000);
          const buffer = await file.arrayBuffer();
          tableData = await parseXLSX(buffer);
          renderTable(tableData);
          controls.style.display = 'flex';
          showToast('✅ XLSX loaded successfully');
        } else if (ext === 'csv') {
          const text = await file.text();
          tableData = text.trim().split(/\r?\n/).map(r =>
            r.split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(c => c.replace(/^"|"$/g, '').trim())
          );
          renderTable(tableData);
          controls.style.display = 'flex';
          showToast('✅ CSV loaded');
        } else if (ext === 'xls') {
          showToast('⚠️ Old .xls not supported, save as .xlsx or .csv', 3000);
        } else {
          showToast('❌ Unsupported format', 2000);
        }
      } catch (err) {
        console.error(err);
        showToast('❌ Error loading file', 2000);
      }
    }

    // ----- RENDER -----
    function renderTable(data) {
      if (!data || data.length === 0) {
        tableWrap.innerHTML = '<div class="empty-state"><svg class="icon" viewBox="0 0 24 24" width="40" height="40"><rect x="2" y="4" width="20" height="16" rx="2" ry="2"/></svg><p><strong>No data</strong></p></div>';
        return;
      }

      // Apply filter
      let displayData = data;
      if (currentFilter) {
        displayData = [data[0], ...data.slice(1).filter(r =>
          r.join(' ').toLowerCase().includes(currentFilter)
        )];
      }

      // Build HTML
      let html = '<table><thead><tr>';

      // Handle empty header row
      if (displayData[0].length === 0) displayData[0] = ['Column 1', 'Column 2'];

      for (let h of displayData[0]) {
        html += `<th>${escapeHTML(h || '')}</th>`;
      }
      html += '</tr></thead><tbody>';

      for (let i = 1; i < displayData.length; i++) {
        html += '<tr>';
        for (let j = 0; j < displayData[0].length; j++) {
          const val = displayData[i]?.[j] !== undefined ? displayData[i][j] : '';
          html += `<td contenteditable oninput="updateCell(${i}, ${j}, this.innerText)">${escapeHTML(val)}</td>`;
        }
        html += '</tr>';
      }
      html += '</tbody></table>';

      tableWrap.innerHTML = html;
    }

    // Global cell updater
    window.updateCell = function (rowIdx, colIdx, val) {
      if (!tableData[rowIdx]) tableData[rowIdx] = [];
      tableData[rowIdx][colIdx] = val;
    };

    // Escape helper
    function escapeHTML(s) {
      if (s === undefined || s === null) return '';
      return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // ----- SEARCH -----
    let searchTimer;
    searchInput.oninput = () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        currentFilter = searchInput.value.toLowerCase();
        if (tableData.length) renderTable(tableData);
      }, 200);
    };

    // ----- EXPORTS -----
    window.exportCSV = function () {
      if (!tableData.length) return showToast('❌ No data to export', 1500);
      const csv = tableData.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
      download(csv, 'edited.csv', 'text/csv');
      showToast('⬇️ CSV exported');
    };

    window.exportExcel = function () {
      if (!tableData.length) return showToast('❌ No data to export', 1500);
      const html = `<table>${tableData.map(r => '<tr>' + r.map(c => `<td>${c}</td>`).join('') + '</tr>').join('')}</table>`;
      download(html, 'edited.xls', 'application/vnd.ms-excel');
      showToast('⬇️ XLS exported');
    };

    window.exportPDF = function () {
      if (!tableData.length) return showToast('❌ No data to export', 1500);
      const w = window.open('');
      w.document.write(`<html><head><style>
        table{border-collapse:collapse;width:100%;}
        td,th{border:1px solid #000;padding:4px;}
        @page{size:A4 landscape;margin:8mm;}
      </style></head><body>${tableWrap.innerHTML}</body></html>`);
      w.document.close();
      setTimeout(() => { w.print(); w.close(); showToast('📄 PDF ready'); }, 300);
    };

    function download(data, name, mime) {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([data], { type: mime }));
      a.download = name;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 100);
    }

    // ----- EVENT LISTENERS -----
    uploadBox.onclick = () => fileInput.click();
    fileInput.onchange = e => e.target.files[0] && loadFile(e.target.files[0]);

    uploadBox.ondragover = e => { e.preventDefault(); uploadBox.style.borderColor = 'var(--accent2)'; };
    uploadBox.ondragleave = () => uploadBox.style.borderColor = 'var(--upload-dash)';
    uploadBox.ondrop = e => {
      e.preventDefault();
      uploadBox.style.borderColor = 'var(--upload-dash)';
      if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
    };

    // First edit hint
    document.addEventListener('input', e => {
      if (e.target.matches('td[contenteditable]') && !window.editHintShown) {
        showToast('✏️ Cell updated');
        window.editHintShown = true;
      }
    }, { once: true });

    // PWA install
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('service-worker.js').catch(() => { });
    }
    let deferredPrompt;
    window.addEventListener('beforeinstallprompt', e => {
      e.preventDefault();
      deferredPrompt = e;
      document.getElementById('installBtn').hidden = false;
    });
    document.getElementById('installBtn').onclick = async () => {
      deferredPrompt.prompt();
      await deferredPrompt.userChoice;
      deferredPrompt = null;
      document.getElementById('installBtn').hidden = true;
    };
