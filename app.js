// ============================================================================
// CHOCONOTA - SISTEMA DE CONTROLE DE DEVEDORES (TRUFAS E CONES)
// ============================================================================

const STORAGE_KEYS = {
  RECORDS: 'choconota_records_v1',
  PRICES: 'choconota_prices_v1',
  WEBHOOK_URL: 'choconota_sheets_url_v1',
  AUTH_HASH: 'choconota_auth_hash_v1',
  AUTH_SESSION: 'choconota_auth_session_v1'
};

// URL da Planilha Google criptografada (só é descriptografada ao digitar o login e senha corretos)
// Assim, mesmo com o código público no GitHub, ninguém consegue ver o link da sua planilha!
const ENCRYPTED_DEFAULT_URL =
  'a4dac6f206e65ac744026f6c7aa495fdd92147519758ac757d59d852b9aa9adeac886a11d200c7598ce5138294b9a06dc3a269a2f7d9931bf4b1986a0f7573d87471ab43c91a50d872d40ad127920f84d77fc36c0de415832b03c4524c15b55833b4c2d69a7a9731fa181bef4b7a4a07';

// Hash SHA-256 padrão para autenticação (a senha nunca fica exposta em texto puro)
const DEFAULT_AUTH_HASH =
  'fbed72fdc17a21906cb23d6929f7395a337ecc102a1853b614ca3d22b56e695e';

// Preços iniciais solicitados: Trufa = R$ 5,00 | Cone = R$ 8,00
const DEFAULT_PRICES = {
  trufa: 5.00,
  cone: 8.00
};

let state = {
  isAuthenticated: false,
  authHash: DEFAULT_AUTH_HASH,
  records: [],
  prices: { ...DEFAULT_PRICES },
  webhookUrl: '',
  isSyncing: false,
  personFilter: 'PENDENTE',
  personSearch: '',
  historyFilter: 'TODOS',
  historySearch: '',
  customTotalTouched: false
};

// Código do Google Apps Script embutido para o botão "Copiar Código"
const APPS_SCRIPT_CODE = `const SHEET_VENDAS = 'Vendas';
const SHEET_RESUMO = 'Resumo Devedores';

function getOrCreateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetVendas = ss.getSheetByName(SHEET_VENDAS);
  if (!sheetVendas) sheetVendas = ss.insertSheet(SHEET_VENDAS);
  if (sheetVendas.getLastRow() === 0) {
    const headers = ['ID','Data','Horário','Nome da Pessoa','Descrição Produtos','Qtd Trufas','Qtd Cones','Quantidade Total','Sabor / Observação','Valor Total (R$)','Status'];
    sheetVendas.appendRow(headers);
    sheetVendas.getRange(1, 1, 1, headers.length).setBackground('#4E2618').setFontColor('#FFFFFF').setFontWeight('bold');
    sheetVendas.setFrozenRows(1);
  }
  let sheetResumo = ss.getSheetByName(SHEET_RESUMO);
  if (!sheetResumo) sheetResumo = ss.insertSheet(SHEET_RESUMO);
  return { ss, sheetVendas, sheetResumo };
}

function doGet() {
  try {
    const { sheetVendas } = getOrCreateSheets();
    const rows = sheetVendas.getDataRange().getDisplayValues();
    const records = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      records.push({
        id: String(r[0]), date: String(r[1]), time: String(r[2]),
        customerName: String(r[3]), productDesc: String(r[4]),
        qtyTrufa: Number(r[5]) || 0, qtyCone: Number(r[6]) || 0,
        totalQty: Number(r[7]) || 0, flavorNotes: String(r[8] || ''),
        totalValue: parseFloat(String(r[9]).replace('R$','').replace(/\\./g,'').replace(',','.').trim()) || 0,
        status: String(r[10] || 'PENDENTE')
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, records })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const payload = JSON.parse(e.postData.contents);
    const { sheetVendas, sheetResumo } = getOrCreateSheets();
    if (payload.action === 'SYNC_ALL' && Array.isArray(payload.records)) {
      writeAllRecords(sheetVendas, payload.records);
      rebuildResumoSheet(sheetResumo, payload.records);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function formatRowValues(rec) {
  return [rec.id, rec.date, rec.time, rec.customerName, rec.productDesc, Number(rec.qtyTrufa)||0, Number(rec.qtyCone)||0, Number(rec.totalQty)||0, rec.flavorNotes||'', Number(rec.totalValue)||0, rec.status||'PENDENTE'];
}

function writeAllRecords(sheet, records) {
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).clearContent();
  if (records.length === 0) return;
  const rows = records.map(formatRowValues);
  sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  sheet.getRange(2, 10, rows.length, 1).setNumberFormat('R$ #,##0.00');
  sheet.autoResizeColumns(1, 11);
}

function rebuildResumoSheet(sheetResumo, records) {
  sheetResumo.clear();
  const headers = ['Nome da Pessoa','Total Pendente (Devedor)','Total Já Pago','Qtd Trufas Pendentes','Qtd Cones Pendentes','Situação'];
  sheetResumo.appendRow(headers);
  sheetResumo.getRange(1, 1, 1, headers.length).setBackground('#4E2618').setFontColor('#FFFFFF').setFontWeight('bold');
  sheetResumo.setFrozenRows(1);
  const map = {};
  records.forEach(r => {
    const key = (r.customerName || 'Sem Nome').trim().toLowerCase();
    if (!map[key]) map[key] = { name: (r.customerName || 'Sem Nome').trim(), pendingValue: 0, paidValue: 0, pendingTrufas: 0, pendingCones: 0 };
    if (r.status === 'PAGO') {
      map[key].paidValue += Number(r.totalValue) || 0;
    } else {
      map[key].pendingValue += Number(r.totalValue) || 0;
      map[key].pendingTrufas += Number(r.qtyTrufa) || 0;
      map[key].pendingCones += Number(r.qtyCone) || 0;
    }
  });
  const rows = Object.values(map).sort((a,b) => b.pendingValue - a.pendingValue).map(p => [p.name, p.pendingValue, p.paidValue, p.pendingTrufas, p.pendingCones, p.pendingValue > 0 ? 'PENDENTE' : 'QUITADO']);
  if (rows.length > 0) {
    sheetResumo.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheetResumo.getRange(2, 2, rows.length, 2).setNumberFormat('R$ #,##0.00');
    sheetResumo.autoResizeColumns(1, headers.length);
  }
}`;

// ============================================================================
// INICIALIZAÇÃO E AUTENTICAÇÃO
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  loadLocalState();
  bindAuthEvents();
  bindEvents();
  initClock();
  checkInitialAuth();
});

function loadLocalState() {
  try {
    const savedHash = localStorage.getItem(STORAGE_KEYS.AUTH_HASH);
    if (savedHash) {
      state.authHash = savedHash;
    }

    const savedRecords = localStorage.getItem(STORAGE_KEYS.RECORDS);
    if (savedRecords) {
      state.records = JSON.parse(savedRecords);
    }

    const savedPrices = localStorage.getItem(STORAGE_KEYS.PRICES);
    if (savedPrices) {
      const parsed = JSON.parse(savedPrices);
      state.prices = {
        trufa: Number(parsed.trufa) >= 0 ? Number(parsed.trufa) : DEFAULT_PRICES.trufa,
        cone: Number(parsed.cone) >= 0 ? Number(parsed.cone) : DEFAULT_PRICES.cone
      };
    }

    const savedUrl = localStorage.getItem(STORAGE_KEYS.WEBHOOK_URL);
    if (savedUrl && savedUrl.trim()) {
      state.webhookUrl = savedUrl.trim();
    }
  } catch (e) {
    console.error('Erro ao carregar dados locais:', e);
  }
}

function decryptUrl(hex, keyStreamSeed) {
  let keyHex = '';
  let counter = 0;
  while (keyHex.length < hex.length) {
    keyHex += sha256(keyStreamSeed + ':' + counter++);
  }
  let out = '';
  for (let i = 0; i < hex.length; i += 2) {
    const k = parseInt(keyHex.substr(i, 2), 16);
    const c = parseInt(hex.substr(i, 2), 16) ^ k;
    out += String.fromCharCode(c);
  }
  return out;
}

function checkInitialAuth() {
  const savedSession =
    localStorage.getItem(STORAGE_KEYS.AUTH_SESSION) ||
    sessionStorage.getItem(STORAGE_KEYS.AUTH_SESSION);

  if (savedSession && savedSession === state.authHash) {
    unlockApp();
  } else {
    lockApp();
  }
}

function bindAuthEvents() {
  const loginForm = document.getElementById('loginForm');
  const btnTogglePass = document.getElementById('btnTogglePass');
  const passInput = document.getElementById('loginPassword');
  const btnLogout = document.getElementById('btnLogout');

  btnTogglePass.addEventListener('click', () => {
    const isPassword = passInput.type === 'password';
    passInput.type = isPassword ? 'text' : 'password';
    btnTogglePass.textContent = isPassword ? '🙈' : '👁️';
  });

  loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const user = document.getElementById('loginUsername').value.trim().toLowerCase();
    const pass = document.getElementById('loginPassword').value;
    const remember = document.getElementById('rememberMe').checked;
    const errorEl = document.getElementById('loginErrorMsg');

    const seed = `${user}:${pass}`;
    const candidateHash = sha256(seed);

    if (candidateHash === state.authHash) {
      errorEl.classList.add('hidden');

      // Descriptografa a URL da Planilha Google usando o login e senha corretos
      if (!state.webhookUrl && candidateHash === DEFAULT_AUTH_HASH) {
        const unlockedUrl = decryptUrl(ENCRYPTED_DEFAULT_URL, seed);
        if (unlockedUrl.startsWith('https://')) {
          state.webhookUrl = unlockedUrl;
          localStorage.setItem(STORAGE_KEYS.WEBHOOK_URL, unlockedUrl);
        }
      }

      if (remember) {
        localStorage.setItem(STORAGE_KEYS.AUTH_SESSION, candidateHash);
      } else {
        sessionStorage.setItem(STORAGE_KEYS.AUTH_SESSION, candidateHash);
      }
      document.getElementById('loginPassword').value = '';
      unlockApp();
      showToast('🔓 Bem-vindo ao ChocoNota!');
    } else {
      errorEl.classList.remove('hidden');
    }
  });

  btnLogout.addEventListener('click', () => {
    localStorage.removeItem(STORAGE_KEYS.AUTH_SESSION);
    sessionStorage.removeItem(STORAGE_KEYS.AUTH_SESSION);
    lockApp();
    showToast('🔒 Sistema bloqueado com segurança.');
  });
}

function unlockApp() {
  state.isAuthenticated = true;
  document.getElementById('loginScreen').classList.add('hidden');
  document.getElementById('appShell').classList.remove('hidden');

  updatePriceDisplays();
  recalculateFormTotal();
  renderAll();

  if (state.webhookUrl) {
    pullFromGoogleSheets(false);
  }
}

function lockApp() {
  state.isAuthenticated = false;
  document.getElementById('appShell').classList.add('hidden');
  document.getElementById('settingsModal').classList.add('hidden');
  document.getElementById('loginScreen').classList.remove('hidden');
  document.getElementById('loginUsername').focus();
}

function saveLocalRecords() {
  localStorage.setItem(STORAGE_KEYS.RECORDS, JSON.stringify(state.records));
}

function initClock() {
  const clockEl = document.getElementById('liveClock');
  const update = () => {
    const now = new Date();
    const dateStr = now.toLocaleDateString('pt-BR');
    const timeStr = now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    clockEl.textContent = `${dateStr} • ${timeStr}`;
  };
  update();
  setInterval(update, 15000);
}

// ============================================================================
// EVENTOS DA INTERFACE
// ============================================================================
function bindEvents() {
  // Navegação entre abas
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      switchTab(btn.dataset.tab);
    });
  });

  // Botões + e - de quantidade de Trufa e Cone
  document.querySelectorAll('.stepper-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const product = btn.dataset.product;
      const isPlus = btn.classList.contains('plus');
      const input = document.getElementById(product === 'trufa' ? 'qtyTrufa' : 'qtyCone');
      let current = parseInt(input.value, 10) || 0;
      current = isPlus ? current + 1 : Math.max(0, current - 1);
      input.value = current;
      state.customTotalTouched = false;
      recalculateFormTotal();
    });
  });

  // Se o usuário editar manualmente o valor total (ex: desconto)
  const customTotalInput = document.getElementById('customTotalValue');
  customTotalInput.addEventListener('input', () => {
    state.customTotalTouched = true;
  });

  // Chips de sabores rápidos
  document.querySelectorAll('#flavorChips .chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const flavor = chip.dataset.flavor;
      const input = document.getElementById('flavorNotes');
      const current = input.value.trim();
      if (!current) {
        input.value = flavor;
      } else if (!current.toLowerCase().includes(flavor.toLowerCase())) {
        input.value = `${current}, ${flavor}`;
      }
      input.focus();
    });
  });

  // Autocomplete de Nome da Pessoa
  const nameInput = document.getElementById('customerName');
  const suggestionsBox = document.getElementById('customerSuggestions');

  nameInput.addEventListener('input', () => {
    renderNameSuggestions(nameInput.value);
  });

  nameInput.addEventListener('focus', () => {
    if (nameInput.value.trim().length > 0) {
      renderNameSuggestions(nameInput.value);
    }
  });

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.autocomplete-wrapper')) {
      suggestionsBox.classList.add('hidden');
    }
  });

  // Submissão do formulário de Nova Anotação
  document.getElementById('saleForm').addEventListener('submit', handleFormSubmit);

  // Filtros na aba "Por Pessoa"
  document.getElementById('searchPersonInput').addEventListener('input', (e) => {
    state.personSearch = e.target.value;
    renderPeopleList();
  });

  document.querySelectorAll('#personFilterGroup .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#personFilterGroup .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.personFilter = btn.dataset.filter;
      renderPeopleList();
    });
  });

  // Filtros na aba "Histórico"
  document.getElementById('searchHistoryInput').addEventListener('input', (e) => {
    state.historySearch = e.target.value;
    renderHistoryList();
  });

  document.querySelectorAll('#historyFilterGroup .seg-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#historyFilterGroup .seg-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.historyFilter = btn.dataset.filter;
      renderHistoryList();
    });
  });

  // Botões de Modal de Configurações e Preços
  document.getElementById('btnOpenSettings').addEventListener('click', openSettingsModal);
  document.getElementById('btnSyncStatus').addEventListener('click', openSettingsModal);
  document.getElementById('btnEditPricesInline').addEventListener('click', openSettingsModal);
  document.getElementById('btnCloseSettings').addEventListener('click', closeSettingsModal);
  document.getElementById('settingsModal').addEventListener('click', (e) => {
    if (e.target.id === 'settingsModal') closeSettingsModal();
  });

  // Salvar preços
  document.getElementById('btnSavePrices').addEventListener('click', () => {
    const newTrufa = parseFloat(document.getElementById('settingPriceTrufa').value);
    const newCone = parseFloat(document.getElementById('settingPriceCone').value);
    if (isNaN(newTrufa) || newTrufa < 0 || isNaN(newCone) || newCone < 0) {
      showToast('⚠️ Digite valores válidos para os preços.');
      return;
    }
    state.prices.trufa = newTrufa;
    state.prices.cone = newCone;
    localStorage.setItem(STORAGE_KEYS.PRICES, JSON.stringify(state.prices));
    updatePriceDisplays();
    state.customTotalTouched = false;
    recalculateFormTotal();
    showToast('✅ Preços atualizados com sucesso!');
  });

  // Alterar Usuário e Senha
  document.getElementById('btnSaveCredentials').addEventListener('click', () => {
    const newUser = document.getElementById('settingNewUser').value.trim().toLowerCase();
    const newPass = document.getElementById('settingNewPass').value;

    if (!newUser || !newPass) {
      showToast('⚠️ Preencha o novo usuário e a nova senha.');
      return;
    }

    const newHash = sha256(`${newUser}:${newPass}`);
    state.authHash = newHash;
    localStorage.setItem(STORAGE_KEYS.AUTH_HASH, newHash);
    localStorage.setItem(STORAGE_KEYS.AUTH_SESSION, newHash);

    document.getElementById('settingNewUser').value = '';
    document.getElementById('settingNewPass').value = '';
    showToast(`🔐 Senha atualizada para o usuário "${newUser}"!`);
  });

  // Salvar URL da Planilha Google
  document.getElementById('btnSaveWebhook').addEventListener('click', async () => {
    const url = document.getElementById('sheetsWebhookUrl').value.trim();
    state.webhookUrl = url;
    localStorage.setItem(STORAGE_KEYS.WEBHOOK_URL, url);
    updateSyncStatusUI();

    if (!url) {
      showToast('Modo local ativado (sem link da planilha).');
      return;
    }
    await syncWithGoogleSheets(true);
    closeSettingsModal();
  });

  // Mostrar código do Google Apps Script
  document.getElementById('btnShowTutorial').addEventListener('click', () => {
    const box = document.getElementById('tutorialAccordion');
    box.classList.toggle('hidden');
  });

  // Copiar código do Google Apps Script
  document.getElementById('btnCopyAppsScript').addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(APPS_SCRIPT_CODE);
      showToast('📋 Código copiado! Agora cole no Apps Script da Planilha.');
    } catch {
      showToast('⚠️ Não foi possível copiar automaticamente. Veja o arquivo Code.gs.');
    }
  });

  // Botões de sincronizar e exportar CSV no Histórico
  document.getElementById('btnSyncNow').addEventListener('click', () => {
    if (!state.webhookUrl) {
      openSettingsModal();
      showToast('Configure o link da sua Planilha Google primeiro.');
      return;
    }
    syncWithGoogleSheets(true);
  });

  document.getElementById('btnExportCSV').addEventListener('click', exportToCSV);
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-btn').forEach(b => {
    const active = b.dataset.tab === tabId;
    b.classList.toggle('active', active);
    b.setAttribute('aria-selected', active ? 'true' : 'false');
  });
  document.querySelectorAll('.tab-panel').forEach(p => {
    p.classList.toggle('active', p.id === tabId);
  });
}

// ============================================================================
// CÁLCULOS E FORMULÁRIO DE NOVA VENDA
// ============================================================================
function updatePriceDisplays() {
  document.getElementById('displayPriceTrufa').textContent = `${formatCurrency(state.prices.trufa)} un.`;
  document.getElementById('displayPriceCone').textContent = `${formatCurrency(state.prices.cone)} un.`;
  document.getElementById('settingPriceTrufa').value = state.prices.trufa.toFixed(2);
  document.getElementById('settingPriceCone').value = state.prices.cone.toFixed(2);
}

function recalculateFormTotal() {
  const qtyTrufa = parseInt(document.getElementById('qtyTrufa').value, 10) || 0;
  const qtyCone = parseInt(document.getElementById('qtyCone').value, 10) || 0;

  document.getElementById('cardTrufa').classList.toggle('has-qty', qtyTrufa > 0);
  document.getElementById('cardCone').classList.toggle('has-qty', qtyCone > 0);

  const calculatedTotal = (qtyTrufa * state.prices.trufa) + (qtyCone * state.prices.cone);
  const totalQty = qtyTrufa + qtyCone;

  const hintEl = document.getElementById('totalItemsHint');
  if (totalQty === 0) {
    hintEl.textContent = 'Selecione as quantidades acima';
  } else {
    const parts = [];
    if (qtyTrufa > 0) parts.push(`${qtyTrufa} ${qtyTrufa === 1 ? 'Trufa' : 'Trufas'}`);
    if (qtyCone > 0) parts.push(`${qtyCone} ${qtyCone === 1 ? 'Cone' : 'Cones'}`);
    hintEl.textContent = parts.join(' + ');
  }

  if (!state.customTotalTouched) {
    document.getElementById('customTotalValue').value = calculatedTotal.toFixed(2);
  }
}

function handleFormSubmit(e) {
  e.preventDefault();
  if (!state.isAuthenticated) return;

  const customerName = document.getElementById('customerName').value.trim();
  const qtyTrufa = parseInt(document.getElementById('qtyTrufa').value, 10) || 0;
  const qtyCone = parseInt(document.getElementById('qtyCone').value, 10) || 0;
  const totalQty = qtyTrufa + qtyCone;
  const flavorNotes = document.getElementById('flavorNotes').value.trim();
  const totalValue = parseFloat(document.getElementById('customTotalValue').value) || 0;
  const statusRadio = document.querySelector('input[name="saleStatus"]:checked');
  const status = statusRadio ? statusRadio.value : 'PENDENTE';

  if (!customerName) {
    showToast('⚠️ Informe o nome da pessoa.');
    document.getElementById('customerName').focus();
    return;
  }

  if (totalQty <= 0) {
    showToast('⚠️ Adicione pelo menos 1 Trufa ou 1 Cone.');
    return;
  }

  const now = new Date();
  const dateStr = now.toLocaleDateString('pt-BR');
  const timeStr = now.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

  const descParts = [];
  if (qtyTrufa > 0) descParts.push(`${qtyTrufa}x Trufa`);
  if (qtyCone > 0) descParts.push(`${qtyCone}x Cone`);

  const newRecord = {
    id: 'choco_' + Date.now() + '_' + Math.random().toString(36).substring(2, 6),
    date: dateStr,
    time: timeStr,
    customerName,
    productDesc: descParts.join(' + '),
    qtyTrufa,
    qtyCone,
    totalQty,
    flavorNotes,
    totalValue,
    status
  };

  state.records.unshift(newRecord);
  saveLocalRecords();
  renderAll();

  // Limpar formulário para a próxima anotação
  document.getElementById('customerName').value = '';
  document.getElementById('qtyTrufa').value = '0';
  document.getElementById('qtyCone').value = '0';
  document.getElementById('flavorNotes').value = '';
  document.querySelector('input[name="saleStatus"][value="PENDENTE"]').checked = true;
  state.customTotalTouched = false;
  recalculateFormTotal();

  showToast(`🍫 Anotado para ${customerName} (${formatCurrency(totalValue)})!`);

  // Sincronizar em background com a planilha Google
  syncWithGoogleSheets(false);
}

// ============================================================================
// AUTOCOMPLETE E CHIPS DE CLIENTES FREQUENTES
// ============================================================================
function getGroupedPeopleMap() {
  const map = {};
  state.records.forEach(r => {
    const cleanName = (r.customerName || 'Sem Nome').trim();
    const key = cleanName.toLowerCase();
    if (!map[key]) {
      map[key] = {
        key,
        name: cleanName,
        pendingValue: 0,
        paidValue: 0,
        pendingTrufas: 0,
        pendingCones: 0,
        items: []
      };
    }
    map[key].items.push(r);
    if (r.status === 'PAGO') {
      map[key].paidValue += Number(r.totalValue) || 0;
    } else {
      map[key].pendingValue += Number(r.totalValue) || 0;
      map[key].pendingTrufas += Number(r.qtyTrufa) || 0;
      map[key].pendingCones += Number(r.qtyCone) || 0;
    }
  });
  return map;
}

function renderNameSuggestions(query) {
  const box = document.getElementById('customerSuggestions');
  const q = query.trim().toLowerCase();
  if (!q) {
    box.classList.add('hidden');
    return;
  }

  const people = Object.values(getGroupedPeopleMap()).filter(p =>
    p.name.toLowerCase().includes(q)
  );

  if (people.length === 0) {
    box.classList.add('hidden');
    return;
  }

  box.innerHTML = people.slice(0, 6).map(p => `
    <div class="suggestion-item" data-name="${escapeHtml(p.name)}">
      <span>${escapeHtml(p.name)}</span>
      ${p.pendingValue > 0
        ? `<span class="suggestion-debt">Deve ${formatCurrency(p.pendingValue)}</span>`
        : `<span style="font-size:0.75rem;color:var(--paid-text)">Em dia</span>`
      }
    </div>
  `).join('');

  box.classList.remove('hidden');

  box.querySelectorAll('.suggestion-item').forEach(item => {
    item.addEventListener('click', () => {
      document.getElementById('customerName').value = item.dataset.name;
      box.classList.add('hidden');
    });
  });
}

function renderQuickCustomersBar() {
  const bar = document.getElementById('QuickCustomersBar');
  const people = Object.values(getGroupedPeopleMap())
    .sort((a, b) => b.pendingValue - a.pendingValue || b.items.length - a.items.length)
    .slice(0, 5);

  if (people.length === 0) {
    bar.innerHTML = '';
    return;
  }

  bar.innerHTML = people.map(p => `
    <button type="button" class="chip quick-person-chip" data-name="${escapeHtml(p.name)}">
      👤 ${escapeHtml(p.name)}
    </button>
  `).join('');

  bar.querySelectorAll('.quick-person-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.getElementById('customerName').value = chip.dataset.name;
    });
  });
}

// ============================================================================
// RENDERIZAÇÃO GERAL (CARDS DE RESUMO, POR PESSOA, HISTÓRICO)
// ============================================================================
function renderAll() {
  if (!state.isAuthenticated) return;
  renderTopSummary();
  renderQuickCustomersBar();
  renderPeopleList();
  renderHistoryList();
  updateSyncStatusUI();
}

function renderTopSummary() {
  let pendingTotal = 0;
  let paidTotal = 0;
  let totalItems = 0;
  const debtorsSet = new Set();

  state.records.forEach(r => {
    const val = Number(r.totalValue) || 0;
    totalItems += Number(r.totalQty) || 0;
    if (r.status === 'PAGO') {
      paidTotal += val;
    } else {
      pendingTotal += val;
      if (r.customerName) debtorsSet.add(r.customerName.trim().toLowerCase());
    }
  });

  document.getElementById('totalPendingValue').textContent = formatCurrency(pendingTotal);
  const debtorCount = debtorsSet.size;
  document.getElementById('totalPendingCount').textContent =
    debtorCount === 1 ? '1 pessoa devendo' : `${debtorCount} pessoas devendo`;

  document.getElementById('totalPaidValue').textContent = formatCurrency(paidTotal);
  document.getElementById('totalItemsSold').textContent =
    totalItems === 1 ? '1 chocolate vendido' : `${totalItems} chocolates vendidos`;

  document.getElementById('badgePessoas').textContent = debtorCount;
}

function renderPeopleList() {
  const container = document.getElementById('peopleListContainer');
  const map = getGroupedPeopleMap();
  let list = Object.values(map);

  if (state.personFilter === 'PENDENTE') {
    list = list.filter(p => p.pendingValue > 0);
  } else if (state.personFilter === 'PAGO') {
    list = list.filter(p => p.pendingValue === 0 && p.paidValue > 0);
  }

  const q = state.personSearch.trim().toLowerCase();
  if (q) {
    list = list.filter(p => p.name.toLowerCase().includes(q));
  }

  list.sort((a, b) => b.pendingValue - a.pendingValue || a.name.localeCompare(b.name));

  if (list.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">🍫</div>
        <p><strong>Nenhuma pessoa encontrada neste filtro.</strong></p>
        <p style="font-size:0.82rem;margin-top:4px;">Use a aba "Anotar" para registrar vendas de trufas e cones.</p>
      </div>
    `;
    return;
  }

  container.innerHTML = list.map(person => {
    const hasDebt = person.pendingValue > 0;
    const itemSummaryParts = [];
    if (person.pendingTrufas > 0) itemSummaryParts.push(`${person.pendingTrufas} Trufa(s)`);
    if (person.pendingCones > 0) itemSummaryParts.push(`${person.pendingCones} Cone(s)`);
    const pendingItemsText = itemSummaryParts.length > 0
      ? `Pendente: ${itemSummaryParts.join(' e ')}`
      : `${person.items.length} anotação(ões) no histórico`;

    return `
      <div class="person-card ${hasDebt ? 'has-debt' : 'all-paid'}">
        <div class="person-summary">
          <div class="person-info">
            <h3>${escapeHtml(person.name)}</h3>
            <span class="person-meta">${pendingItemsText}</span>
          </div>
          <div class="person-amounts">
            ${hasDebt
              ? `<span class="person-debt-badge">${formatCurrency(person.pendingValue)}</span>`
              : `<span class="person-paid-badge">✅ Quitado (${formatCurrency(person.paidValue)})</span>`
            }
          </div>
        </div>

        <div class="person-actions-row">
          ${hasDebt ? `
            <button type="button" class="btn-pay-all" data-person-key="${escapeHtml(person.key)}">
              ✅ Receber Tudo (${formatCurrency(person.pendingValue)})
            </button>
          ` : ''}
          <button type="button" class="btn-add-more" data-person-name="${escapeHtml(person.name)}">
            ➕ Anotar mais p/ ${escapeHtml(person.name.split(' ')[0])}
          </button>
        </div>

        <div class="person-items-list">
          ${person.items.map(item => `
            <div class="person-item-row">
              <div>
                <strong>${escapeHtml(item.productDesc)}</strong>
                ${item.flavorNotes ? `<span class="history-flavor"> • ${escapeHtml(item.flavorNotes)}</span>` : ''}
                <div class="history-datetime">📅 ${escapeHtml(item.date)} às ${escapeHtml(item.time)}</div>
              </div>
              <div style="display:flex;align-items:center;gap:8px;">
                <strong>${formatCurrency(item.totalValue)}</strong>
                <button
                  type="button"
                  class="btn-toggle-status ${item.status === 'PENDENTE' ? 'mark-paid' : ''}"
                  data-id="${escapeHtml(item.id)}"
                >
                  ${item.status === 'PENDENTE' ? '✅ Pagar' : '↩️ Pago'}
                </button>
              </div>
            </div>
          `).join('')}
        </div>
      </div>
    `;
  }).join('');

  container.querySelectorAll('.btn-pay-all').forEach(btn => {
    btn.addEventListener('click', () => {
      markAllPersonPaid(btn.dataset.personKey);
    });
  });

  container.querySelectorAll('.btn-add-more').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.dataset.personName;
      switchTab('tab-novo');
      const nameInput = document.getElementById('customerName');
      nameInput.value = name;
      window.scrollTo({ top: 0, behavior: 'smooth' });
      showToast(`Anotando nova venda para ${name}`);
    });
  });

  container.querySelectorAll('.btn-toggle-status').forEach(btn => {
    btn.addEventListener('click', () => {
      toggleRecordStatus(btn.dataset.id);
    });
  });
}

function renderHistoryList() {
  const container = document.getElementById('historyListContainer');
  let list = [...state.records];

  if (state.historyFilter !== 'TODOS') {
    list = list.filter(r => r.status === state.historyFilter);
  }

  const q = state.historySearch.trim().toLowerCase();
  if (q) {
    list = list.filter(r =>
      (r.customerName || '').toLowerCase().includes(q) ||
      (r.flavorNotes || '').toLowerCase().includes(q) ||
      (r.date || '').toLowerCase().includes(q) ||
      (r.productDesc || '').toLowerCase().includes(q)
    );
  }

  document.getElementById('historyCountLabel').textContent =
    list.length === 1 ? '1 registro' : `${list.length} registros`;

  if (list.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state-icon">📋</div>
        <p><strong>Nenhum lançamento encontrado.</strong></p>
      </div>
    `;
    return;
  }

  container.innerHTML = list.map(item => `
    <div class="history-card">
      <div class="history-top">
        <span class="history-person">${escapeHtml(item.customerName)}</span>
        <span class="history-datetime">📅 ${escapeHtml(item.date)} • ⏰ ${escapeHtml(item.time)}</span>
      </div>
      <div class="history-middle">
        <div>
          <div class="history-products">${escapeHtml(item.productDesc)} (Qtd: ${item.totalQty})</div>
          ${item.flavorNotes ? `<div class="history-flavor">OBS/Sabor: ${escapeHtml(item.flavorNotes)}</div>` : ''}
        </div>
        <div class="history-price">${formatCurrency(item.totalValue)}</div>
      </div>
      <div class="history-bottom">
        <span class="status-badge ${item.status === 'PAGO' ? 'pago' : 'pendente'}">
          ${item.status === 'PAGO' ? '✅ Pago' : '⏳ Pendente'}
        </span>
        <div class="item-btns">
          <button
            type="button"
            class="btn-toggle-status ${item.status === 'PENDENTE' ? 'mark-paid' : ''}"
            data-id="${escapeHtml(item.id)}"
          >
            ${item.status === 'PENDENTE' ? 'Marcar Pago' : 'Voltar p/ Pendente'}
          </button>
          <button type="button" class="btn-delete" data-id="${escapeHtml(item.id)}" title="Excluir anotação">
            🗑️
          </button>
        </div>
      </div>
    </div>
  `).join('');

  container.querySelectorAll('.btn-toggle-status').forEach(btn => {
    btn.addEventListener('click', () => toggleRecordStatus(btn.dataset.id));
  });

  container.querySelectorAll('.btn-delete').forEach(btn => {
    btn.addEventListener('click', () => deleteRecord(btn.dataset.id));
  });
}

// ============================================================================
// AÇÕES DE STATUS (DAR BAIXA / EXCLUIR)
// ============================================================================
function toggleRecordStatus(id) {
  if (!state.isAuthenticated) return;
  const rec = state.records.find(r => r.id === id);
  if (!rec) return;

  rec.status = rec.status === 'PAGO' ? 'PENDENTE' : 'PAGO';
  saveLocalRecords();
  renderAll();

  showToast(
    rec.status === 'PAGO'
      ? `✅ Pagamento de ${rec.customerName} registrado!`
      : `⏳ Anotação de ${rec.customerName} voltou para Pendente.`
  );
  syncWithGoogleSheets(false);
}

function markAllPersonPaid(personKey) {
  if (!state.isAuthenticated) return;
  let count = 0;
  let personName = '';
  state.records.forEach(r => {
    if ((r.customerName || '').trim().toLowerCase() === personKey && r.status !== 'PAGO') {
      r.status = 'PAGO';
      personName = r.customerName;
      count++;
    }
  });

  if (count > 0) {
    saveLocalRecords();
    renderAll();
    showToast(`🎉 Todas as compras de ${personName} foram quitadas!`);
    syncWithGoogleSheets(false);
  }
}

function deleteRecord(id) {
  if (!state.isAuthenticated) return;
  const rec = state.records.find(r => r.id === id);
  if (!rec) return;

  const confirmDelete = window.confirm(
    `Deseja excluir a anotação de ${rec.customerName} (${rec.productDesc} - ${formatCurrency(rec.totalValue)})?`
  );
  if (!confirmDelete) return;

  state.records = state.records.filter(r => r.id !== id);
  saveLocalRecords();
  renderAll();
  showToast('🗑️ Anotação removida.');
  syncWithGoogleSheets(false);
}

// ============================================================================
// INTEGRAÇÃO COM GOOGLE PLANILHAS (GOOGLE SHEETS)
// ============================================================================
function updateSyncStatusUI() {
  const dot = document.getElementById('syncDot');
  const text = document.getElementById('syncText');
  const modalBadge = document.getElementById('modalConnBadge');
  const urlInput = document.getElementById('sheetsWebhookUrl');

  if (urlInput && document.activeElement !== urlInput) {
    urlInput.value = state.webhookUrl;
  }

  if (state.isSyncing) {
    dot.className = 'sync-dot syncing';
    text.textContent = 'Salvando...';
    return;
  }

  if (state.webhookUrl) {
    dot.className = 'sync-dot connected';
    text.textContent = 'Planilha OK';
    modalBadge.textContent = 'Conectado';
    modalBadge.className = 'conn-badge connected';
  } else {
    dot.className = 'sync-dot';
    text.textContent = 'Conectar Planilha';
    modalBadge.textContent = 'Salvo no Celular';
    modalBadge.className = 'conn-badge';
  }
}

async function syncWithGoogleSheets(showFeedback = false) {
  if (!state.webhookUrl || !state.isAuthenticated) return;

  state.isSyncing = true;
  updateSyncStatusUI();

  try {
    await fetch(state.webhookUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify({
        action: 'SYNC_ALL',
        records: state.records
      })
    });

    if (showFeedback) {
      showToast('📊 Planilha Google sincronizada com sucesso!');
    }
  } catch (err) {
    console.error('Erro ao sincronizar com Google Sheets:', err);
    if (showFeedback) {
      showToast('⚠️ Sem conexão agora. Os dados estão salvos no celular.');
    }
  } finally {
    state.isSyncing = false;
    updateSyncStatusUI();
  }
}

async function pullFromGoogleSheets(showFeedback = false) {
  if (!state.webhookUrl || !state.isAuthenticated) return;

  state.isSyncing = true;
  updateSyncStatusUI();

  try {
    const response = await fetch(state.webhookUrl, { method: 'GET' });
    const data = await response.json();
    if (data && data.ok && Array.isArray(data.records)) {
      if (data.records.length > 0 || state.records.length === 0) {
        state.records = data.records;
        saveLocalRecords();
        renderAll();
      } else if (state.records.length > 0 && data.records.length === 0) {
        await syncWithGoogleSheets(false);
      }
      if (showFeedback) showToast('📊 Dados atualizados da Planilha Google!');
    }
  } catch (err) {
    console.warn('Pull inicial ignorado ou offline:', err);
  } finally {
    state.isSyncing = false;
    updateSyncStatusUI();
  }
}

function exportToCSV() {
  if (!state.isAuthenticated) return;
  if (state.records.length === 0) {
    showToast('⚠️ Não há registros para exportar.');
    return;
  }

  const headers = [
    'Data',
    'Horario',
    'Nome da Pessoa',
    'Produtos',
    'Qtd Trufas',
    'Qtd Cones',
    'Quantidade Total',
    'Sabor / Observacao',
    'Valor Total (R$)',
    'Status'
  ];

  const rows = state.records.map(r => [
    r.date,
    r.time,
    `"${(r.customerName || '').replace(/"/g, '""')}"`,
    `"${(r.productDesc || '').replace(/"/g, '""')}"`,
    r.qtyTrufa || 0,
    r.qtyCone || 0,
    r.totalQty || 0,
    `"${(r.flavorNotes || '').replace(/"/g, '""')}"`,
    (Number(r.totalValue) || 0).toFixed(2).replace('.', ','),
    r.status
  ]);

  const csvContent = '\uFEFF' + [headers.join(';'), ...rows.map(e => e.join(';'))].join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', `choconota_devedores_${new Date().toISOString().slice(0, 10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('📥 Planilha baixada!');
}

// ============================================================================
// CRIPTOGRAFIA SHA-256 (FUNCIONA TANTO NO GITHUB HTTPS QUANTO EM ARQUIVO LOCAL)
// ============================================================================
function sha256(ascii) {
  function rightRotate(value, amount) {
    return (value >>> amount) | (value << (32 - amount));
  }

  // Converte string UTF-8 para bytes ASCII
  const utf8 = unescape(encodeURIComponent(ascii));
  const mathPow = Math.pow;
  const maxWord = mathPow(2, 32);
  const lengthProperty = 'length';
  let i, j;
  let result = '';

  const words = [];
  const asciiBitLength = utf8[lengthProperty] * 8;

  let hash = (sha256.h = sha256.h || []);
  const k = (sha256.k = sha256.k || []);
  let primeCounter = k[lengthProperty];

  const isComposite = {};
  for (let candidate = 2; primeCounter < 64; candidate++) {
    if (!isComposite[candidate]) {
      for (i = 0; i < 313; i += candidate) {
        isComposite[i] = candidate;
      }
      hash[primeCounter] = (mathPow(candidate, 0.5) * maxWord) | 0;
      k[primeCounter++] = (mathPow(candidate, 1 / 3) * maxWord) | 0;
    }
  }

  let padded = utf8 + '\x80';
  while ((padded[lengthProperty] % 64) - 56) padded += '\x00';
  for (i = 0; i < padded[lengthProperty]; i++) {
    j = padded.charCodeAt(i);
    if (j >> 8) return '';
    words[i >> 2] |= j << ((3 - (i % 4)) * 8);
  }
  words[words[lengthProperty]] = (asciiBitLength / maxWord) | 0;
  words[words[lengthProperty]] = asciiBitLength;

  let h = hash.slice(0, 8);
  for (j = 0; j < words[lengthProperty]; ) {
    const w = words.slice(j, (j += 16));
    const oldHash = h.slice(0);

    for (i = 0; i < 64; i++) {
      const i2 = i + j;
      const w15 = w[i - 15],
        w2 = w[i - 2];

      const a = h[0],
        e = h[4];
      const temp1 =
        h[7] +
        (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25)) +
        ((e & h[5]) ^ (~e & h[6])) +
        k[i] +
        (w[i] =
          i < 16
            ? w[i]
            : (w[i - 16] +
                (rightRotate(w15, 7) ^ rightRotate(w15, 18) ^ (w15 >>> 3)) +
                w[i - 7] +
                (rightRotate(w2, 17) ^ rightRotate(w2, 19) ^ (w2 >>> 10))) |
              0);
      const temp2 =
        (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22)) +
        ((a & h[1]) ^ (a & h[2]) ^ (h[1] & h[2]));

      h = [(temp1 + temp2) | 0].concat(h);
      h[4] = (h[4] + temp1) | 0;
      h.pop();
    }

    for (i = 0; i < 8; i++) {
      h[i] = (h[i] + oldHash[i]) | 0;
    }
  }

  for (i = 0; i < 8; i++) {
    for (j = 3; j + 1; j--) {
      const b = (h[i] >> (j * 8)) & 255;
      result += (b < 16 ? 0 : '') + b.toString(16);
    }
  }
  return result;
}

// ============================================================================
// UTILITÁRIOS E MODAIS
// ============================================================================
function openSettingsModal() {
  if (!state.isAuthenticated) return;
  updatePriceDisplays();
  updateSyncStatusUI();
  document.getElementById('settingsModal').classList.remove('hidden');
}

function closeSettingsModal() {
  document.getElementById('settingsModal').classList.add('hidden');
}

function formatCurrency(value) {
  const num = Number(value) || 0;
  return num.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

let toastTimer = null;
function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.classList.remove('hidden');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.add('hidden');
  }, 3000);
}
