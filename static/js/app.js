(() => {
  const dropZone   = document.getElementById('drop-zone');
  const fileInput  = document.getElementById('file-input');
  const fileNameEl = document.getElementById('file-name');
  const convertBtn = document.getElementById('convert-btn');
  const btnText    = document.getElementById('btn-text');
  const btnSpinner = document.getElementById('btn-spinner');

  const uploadCard  = document.getElementById('upload-card');
  const progressCard= document.getElementById('progress-card');
  const resultCard  = document.getElementById('result-card');
  const errorCard   = document.getElementById('error-card');

  const progressBar  = document.getElementById('progress-bar');
  const progressLabel= document.getElementById('progress-label');
  const statsGrid    = document.getElementById('stats-grid');
  const downloadLink = document.getElementById('download-link');
  const errorMsg     = document.getElementById('error-msg');

  let selectedFile = null;

  // ── File selection ─────────────────────────────────────
  function selectFile(file) {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      alert('Por favor, selecione um arquivo PDF.');
      return;
    }
    selectedFile = file;
    fileNameEl.textContent = `📎 ${file.name} (${(file.size/1024).toFixed(1)} KB)`;
    convertBtn.disabled = false;
  }

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) selectFile(fileInput.files[0]);
  });

  dropZone.addEventListener('click', () => fileInput.click());

  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('over');
  });

  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));

  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('over');
    const file = e.dataTransfer.files[0];
    if (file) selectFile(file);
  });

  // ── Progress animation ──────────────────────────────────
  const steps = [
    { pct: 15, label: 'Lendo o PDF…' },
    { pct: 35, label: 'Extraindo produtos e pedidos…' },
    { pct: 60, label: 'Organizando os dados…' },
    { pct: 80, label: 'Gerando planilha Excel…' },
    { pct: 95, label: 'Finalizando…' },
  ];

  function animateProgress() {
    let i = 0;
    function next() {
      if (i >= steps.length) return;
      const s = steps[i++];
      progressBar.style.width  = s.pct + '%';
      progressLabel.textContent = s.label;
      setTimeout(next, 700 + Math.random() * 400);
    }
    next();
  }

  // ── Show / hide cards ───────────────────────────────────
  function showOnly(card) {
    [uploadCard, progressCard, resultCard, errorCard].forEach(c => {
      c.classList.toggle('hidden', c !== card);
    });
  }

  // ── Convert ─────────────────────────────────────────────
  convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    showOnly(progressCard);
    progressBar.style.width = '0%';
    animateProgress();

    const form = new FormData();
    form.append('file', selectedFile);

    try {
      const res  = await fetch('/convert', { method: 'POST', body: form });
      const data = await res.json();

      progressBar.style.width  = '100%';
      progressLabel.textContent = 'Concluído!';

      await new Promise(r => setTimeout(r, 500));

      if (data.success) {
        // Build stats
        const s = data.stats;
        statsGrid.innerHTML = `
          <div class="stat-box">
            <div class="stat-label">Linhas de pedido</div>
            <div class="stat-value">${s.rows.toLocaleString('pt-BR')}</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Produtos únicos</div>
            <div class="stat-value">${s.products.toLocaleString('pt-BR')}</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Clientes únicos</div>
            <div class="stat-value">${s.clients.toLocaleString('pt-BR')}</div>
          </div>
          <div class="stat-box">
            <div class="stat-label">Total Geral</div>
            <div class="stat-value" style="font-size:1rem">${s.total}</div>
          </div>
        `;
        downloadLink.href = `/download/${data.filename}`;
        showOnly(resultCard);
      } else {
        errorMsg.textContent = data.error || 'Erro desconhecido.';
        showOnly(errorCard);
      }
    } catch (err) {
      errorMsg.textContent = 'Falha na comunicação com o servidor: ' + err.message;
      showOnly(errorCard);
    }
  });

  // ── Reset ────────────────────────────────────────────────
  function reset() {
    selectedFile = null;
    fileInput.value = '';
    fileNameEl.textContent = 'Nenhum arquivo selecionado';
    convertBtn.disabled = true;
    showOnly(uploadCard);
  }

  document.getElementById('new-btn').addEventListener('click', reset);
  document.getElementById('retry-btn').addEventListener('click', reset);
})();
