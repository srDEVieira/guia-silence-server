const state = {
  apiReady: false,
  mode: 'ambos',
  maxItemRows: 21,
  patrimonyLength: 5,
  licenseBlocked: false,
  licenseMessage: '',
  licenseConnected: false,
  licenseHeartbeatId: null,
  profiles: [],
  activeProfileId: '',
  themeTransitionTimer: null,
  autosaveTimer: null,
  collapsedCards: new Set(),
  printHistory: [],
  rows: {
    delivery: [],
    receipt: [],
  },
  lookupTimers: new WeakMap(),
};

const els = {};
const DRAFT_STORAGE_KEY = 'guide-draft-v2';
const HISTORY_STORAGE_KEY = 'guide-print-history-v1';
const HISTORY_LIMIT = 8;
const PROFILE_STORAGE_KEY = 'guide-active-profile-v1';

window.addEventListener('pywebviewready', () => {
  state.apiReady = true;
});

window.addEventListener('DOMContentLoaded', async () => {
  cacheElements();
  ensureStartupVideoPlayback();
  bindEvents();
  if (window.__INITIAL_STATE) {
    applyInitialState(window.__INITIAL_STATE);
  } else {
    loadInitialState();
  }
  startLicenseHeartbeat();
  window.addEventListener('resize', syncModeThumb);
  hideStartupOverlay();
});

function cacheElements() {
  els.body = document.body;
  els.startupOverlay = document.getElementById('startup-overlay');
  els.startupVideo = document.querySelector('.startup-overlay__video');
  els.modeSwitcher = document.getElementById('mode-switcher');
  els.modeThumb = els.modeSwitcher ? els.modeSwitcher.querySelector('.segmented__thumb') : null;
  els.inventoryStatus = document.getElementById('inventory-status');
  els.printCard = document.getElementById('print-card');
  els.modeCard = document.getElementById('mode-card');
  els.printerSelect = document.getElementById('printer-select');
  els.refreshPrinters = document.getElementById('refresh-printers');
  els.copiesInput = document.getElementById('copies-input');
  els.allowModifiableGuides = document.getElementById('allow-modifiable-guides');
  els.deliveryReceiverUnit = document.getElementById('delivery-receiver-unit');
  els.deliveryRoom = document.getElementById('delivery-room');
  els.receiptSenderUnit = document.getElementById('receipt-sender-unit');
  els.receiptRoom = document.getElementById('receipt-room');
  els.modeButtons = Array.from(document.querySelectorAll('[data-mode]'));
  els.deliveryBlock = document.getElementById('delivery-block');
  els.receiptBlock = document.getElementById('receipt-block');
  els.deliveryItemsBlock = document.getElementById('delivery-items-block');
  els.receiptItemsBlock = document.getElementById('receipt-items-block');
  els.deliveryModeBlocks = [els.deliveryBlock, els.deliveryItemsBlock];
  els.receiptModeBlocks = [els.receiptBlock, els.receiptItemsBlock];
  els.deliveryItems = document.getElementById('delivery-items');
  els.receiptItems = document.getElementById('receipt-items');
  els.deliveryCount = document.getElementById('delivery-count');
  els.receiptCount = document.getElementById('receipt-count');
  els.deliveryAdd = document.getElementById('delivery-add');
  els.deliveryRemove = document.getElementById('delivery-remove');
  els.receiptAdd = document.getElementById('receipt-add');
  els.receiptRemove = document.getElementById('receipt-remove');
  els.statusText = document.getElementById('status-text');
  els.clearItems = document.getElementById('clear-items');
  els.printButton = document.getElementById('print-button');
  els.toast = document.getElementById('toast');
  els.printReadiness = document.getElementById('print-readiness');
  els.readinessItems = els.printReadiness
    ? {
      printer: els.printReadiness.querySelector('[data-check="printer"]'),
      delivery: els.printReadiness.querySelector('[data-check="delivery"]'),
      receipt: els.printReadiness.querySelector('[data-check="receipt"]'),
      items: els.printReadiness.querySelector('[data-check="items"]'),
    }
    : {};
  els.printHistoryList = document.getElementById('print-history-list');
  els.clearHistory = document.getElementById('clear-history');
  els.heroGuideSteps = {
    flow: document.querySelector('[data-guide-step="flow"]'),
    core: document.querySelector('[data-guide-step="core"]'),
    items: document.querySelector('[data-guide-step="items"]'),
  };
  els.profileSwitch = document.getElementById('profile-switch');
  els.profileList = document.getElementById('profile-list');
  els.collapseButtons = Array.from(document.querySelectorAll('[data-collapse-target]'));
}

function bindEvents() {
  els.refreshPrinters.addEventListener('click', refreshPrinters);
  els.deliveryAdd.addEventListener('click', () => addRow('delivery'));
  els.deliveryRemove.addEventListener('click', () => removeRow('delivery'));
  els.receiptAdd.addEventListener('click', () => addRow('receipt'));
  els.receiptRemove.addEventListener('click', () => removeRow('receipt'));
  els.clearItems.addEventListener('click', clearItems);
  els.printButton.addEventListener('click', printDocuments);
  if (els.clearHistory) {
    els.clearHistory.addEventListener('click', clearHistory);
  }
  if (els.printHistoryList) {
    els.printHistoryList.addEventListener('click', handleHistoryAction);
  }

  els.collapseButtons.forEach((button) => {
    button.addEventListener('click', () => toggleCardCollapse(button.dataset.collapseTarget));
  });

  els.modeButtons.forEach((button) => {
    button.addEventListener('click', () => {
      state.mode = button.dataset.mode;
      updateModeUI();
      applyModeTheme();
      updatePrintReadiness();
      scheduleDraftSave();
    });
  });

  [
    els.deliveryReceiverUnit,
    els.deliveryRoom,
    els.receiptSenderUnit,
    els.receiptRoom,
    els.copiesInput,
    els.printerSelect,
    els.allowModifiableGuides,
  ].forEach((input) => input.addEventListener('input', () => {
    renderPreview();
    updatePrintReadiness();
    scheduleDraftSave();
  }));
}

async function loadInitialState() {
  setStatus('Carregando configurações...');
  hydrateUiPreferences();
  const apiReady = await waitForPywebviewApi();
  if (!apiReady) {
    setStatus('Falha ao conectar com o backend.');
    toast('Não foi possível conectar com o backend.');
    return;
  }
  const result = await window.pywebview.api.get_initial_state();
  applyInitialState(result);
}

async function waitForPywebviewApi(timeoutMs = 8000) {
  if (window.pywebview && window.pywebview.api) {
    state.apiReady = true;
    return true;
  }

  return new Promise((resolve) => {
    let done = false;

    const finish = (ok) => {
      if (done) return;
      done = true;
      window.removeEventListener('pywebviewready', onReady);
      resolve(ok);
    };

    const onReady = () => {
      state.apiReady = true;
      finish(true);
    };

    window.addEventListener('pywebviewready', onReady, { once: true });
    window.setTimeout(() => finish(Boolean(window.pywebview && window.pywebview.api)), timeoutMs);
  });
}

function applyInitialState(result) {
  if (!result.ok) {
    toast('Não foi possível carregar a interface.');
    setStatus('Falha ao carregar configurações.');
    return;
  }

  state.maxItemRows = result.maxItemRows;
  state.patrimonyLength = result.patrimonyLength || 5;
  syncInventoryStatus(result);
  syncLicenseStatus(result);

  fillPrinters(result.printers || [], result.defaultPrinter || '');
  hydratePrintHistory();

  for (let i = 0; i < 1; i += 1) {
    addRow('delivery', false);
    addRow('receipt', false);
  }

  const restored = restoreDraftIfAvailable();
  if (restored) {
    setStatus('Rascunho restaurado automaticamente.');
  }

  updateModeUI(false);
  updateCounters();
  applyModeTheme();
  syncCollapsedCards();
  updatePrintReadiness();
  saveDraftNow();
  renderPrintHistory();
  if (!restored) {
    setStatus('Sistema pronto para preenchimento e impressão.');
  }
  refreshPrintersSilently();
  loadProfiles();
}

async function loadProfiles() {
  if (!state.apiReady || !window.pywebview || !window.pywebview.api || !window.pywebview.api.get_profiles) {
    return;
  }
  try {
    const response = await window.pywebview.api.get_profiles();
    const profiles = Array.isArray(response && response.profiles) ? response.profiles : [];
    state.profiles = profiles.filter((p) => p && p.profile_id && p.display_name);
    renderProfiles();
  } catch {
    // keep default profile button
  }
}

function renderProfiles() {
  if (!els.profileList) return;
  if (!state.profiles.length) {
    applyProfileTheme(null);
    return;
  }

  const savedProfileId = window.localStorage.getItem(PROFILE_STORAGE_KEY) || '';
  let activeId = state.activeProfileId || savedProfileId;
  if (!state.profiles.some((p) => p.profile_id === activeId)) {
    activeId = state.profiles[0].profile_id;
  }
  state.activeProfileId = activeId;

  els.profileList.innerHTML = '';
  state.profiles.forEach((profile) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'profile-pill';
    if (profile.profile_id === activeId) {
      button.classList.add('is-active');
    }
    button.textContent = profile.display_name;
    button.addEventListener('click', () => {
      state.activeProfileId = profile.profile_id;
      try {
        window.localStorage.setItem(PROFILE_STORAGE_KEY, state.activeProfileId);
      } catch {}
      renderProfiles();
      applyProfileTheme(profile);
    });
    els.profileList.appendChild(button);
  });

  const activeProfile = state.profiles.find((p) => p.profile_id === activeId) || null;
  applyProfileTheme(activeProfile);
}

function applyProfileTheme(profile) {
  startThemeTransition();
  const defaultHero = 'url("./assets/tai_lung__kung_fu_panda____wallpaper___g3anzera_by_reysosen_dfz4k15.png")';
  const heroBgUrl = profile && profile.hero_bg_url ? profile.hero_bg_url : '';
  const accentColor = normalizeHexColor(profile && profile.accent_color ? profile.accent_color : '#114c78');
  const accentRgb = hexToRgb(accentColor);
  const accentStrong = rgbToHex(adjustLightness(accentRgb, -0.18));
  const accentSoft = rgbToHex(adjustLightness(accentRgb, 0.35));

  const accentHsl = rgbToHsl(accentRgb);
  const deliveryRgb = hslToRgb(
    (accentHsl.h - 42 + 360) % 360,
    clamp01(accentHsl.s + 0.08),
    clamp01(accentHsl.l + 0.02)
  );
  const receiptRgb = hslToRgb(
    (accentHsl.h + 36) % 360,
    clamp01(Math.max(0.35, accentHsl.s - 0.02)),
    clamp01(Math.min(0.52, accentHsl.l - 0.04))
  );
  const deliveryHex = rgbToHex(deliveryRgb);
  const receiptHex = rgbToHex(receiptRgb);

  document.documentElement.style.setProperty(
    '--hero-profile-image',
    heroBgUrl ? `url("${heroBgUrl.replaceAll('"', '\\"')}")` : defaultHero
  );
  document.documentElement.style.setProperty('--profile-accent', accentColor);
  document.documentElement.style.setProperty('--accent', accentColor);
  document.documentElement.style.setProperty('--accent-strong', accentStrong);
  document.documentElement.style.setProperty('--accent-soft', accentSoft);
  document.documentElement.style.setProperty('--delivery', deliveryHex);
  document.documentElement.style.setProperty('--delivery-soft', `rgba(${deliveryRgb.r}, ${deliveryRgb.g}, ${deliveryRgb.b}, 0.14)`);
  document.documentElement.style.setProperty('--receipt', receiptHex);
  document.documentElement.style.setProperty('--receipt-soft', `rgba(${receiptRgb.r}, ${receiptRgb.g}, ${receiptRgb.b}, 0.14)`);
  document.documentElement.style.setProperty('--accent-rgb', `${accentRgb.r}, ${accentRgb.g}, ${accentRgb.b}`);
  document.documentElement.style.setProperty('--accent-strong-rgb', `${hexToRgb(accentStrong).r}, ${hexToRgb(accentStrong).g}, ${hexToRgb(accentStrong).b}`);
  document.documentElement.style.setProperty('--delivery-rgb', `${deliveryRgb.r}, ${deliveryRgb.g}, ${deliveryRgb.b}`);
  document.documentElement.style.setProperty('--receipt-rgb', `${receiptRgb.r}, ${receiptRgb.g}, ${receiptRgb.b}`);
}

function startThemeTransition() {
  document.body.classList.add('theme-transition');
  if (state.themeTransitionTimer) {
    window.clearTimeout(state.themeTransitionTimer);
  }
  state.themeTransitionTimer = window.setTimeout(() => {
    document.body.classList.remove('theme-transition');
    state.themeTransitionTimer = null;
  }, 520);
}

function normalizeHexColor(value) {
  const raw = String(value || '').trim();
  if (!raw) return '#114c78';
  if (/^#[0-9a-fA-F]{6}$/.test(raw)) return raw.toLowerCase();
  if (/^#[0-9a-fA-F]{3}$/.test(raw)) {
    const r = raw[1];
    const g = raw[2];
    const b = raw[3];
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  return '#114c78';
}

function hexToRgb(hex) {
  const normalized = normalizeHexColor(hex);
  return {
    r: Number.parseInt(normalized.slice(1, 3), 16),
    g: Number.parseInt(normalized.slice(3, 5), 16),
    b: Number.parseInt(normalized.slice(5, 7), 16),
  };
}

function rgbToHex(rgb) {
  const toHex = (n) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, '0');
  return `#${toHex(rgb.r)}${toHex(rgb.g)}${toHex(rgb.b)}`;
}

function rgbToHsl(rgb) {
  const r = rgb.r / 255;
  const g = rgb.g / 255;
  const b = rgb.b / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) {
    return { h: 0, s: 0, l };
  }
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  switch (max) {
    case r: h = ((g - b) / d + (g < b ? 6 : 0)); break;
    case g: h = ((b - r) / d + 2); break;
    default: h = ((r - g) / d + 4); break;
  }
  return { h: h * 60, s, l };
}

function hslToRgb(h, s, l) {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const hp = h / 60;
  const x = c * (1 - Math.abs((hp % 2) - 1));
  let r1 = 0;
  let g1 = 0;
  let b1 = 0;
  if (hp >= 0 && hp < 1) { r1 = c; g1 = x; b1 = 0; }
  else if (hp < 2) { r1 = x; g1 = c; b1 = 0; }
  else if (hp < 3) { r1 = 0; g1 = c; b1 = x; }
  else if (hp < 4) { r1 = 0; g1 = x; b1 = c; }
  else if (hp < 5) { r1 = x; g1 = 0; b1 = c; }
  else { r1 = c; g1 = 0; b1 = x; }
  const m = l - c / 2;
  return {
    r: (r1 + m) * 255,
    g: (g1 + m) * 255,
    b: (b1 + m) * 255,
  };
}

function adjustLightness(rgb, delta) {
  const hsl = rgbToHsl(rgb);
  return hslToRgb(hsl.h, hsl.s, clamp01(hsl.l + delta));
}

function clamp01(value) {
  return Math.max(0, Math.min(1, value));
}

function fillPrinters(printers, defaultPrinter) {
  els.printerSelect.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = printers.length ? 'Selecione uma impressora' : 'Nenhuma impressora encontrada';
  els.printerSelect.appendChild(placeholder);

  printers.forEach((printer) => {
    const option = document.createElement('option');
    option.value = printer;
    option.textContent = printer;
    els.printerSelect.appendChild(option);
  });

  if (defaultPrinter && printers.includes(defaultPrinter)) {
    els.printerSelect.value = defaultPrinter;
  } else if (printers.length) {
    els.printerSelect.value = printers[0];
  }
}

async function refreshPrinters() {
  if (!state.apiReady) {
    toast('A interface ainda está conectando com o backend.');
    return;
  }
  setStatus('Atualizando impressoras...');
  const result = await window.pywebview.api.refresh_printers();
  fillPrinters(result.printers || [], result.defaultPrinter || '');
  syncInventoryStatus(result);
  syncLicenseStatus(result);
  updatePrintReadiness();
  scheduleDraftSave();
  setStatus('Impressoras atualizadas.');
}

async function refreshPrintersSilently() {
  if (!state.apiReady) return;
  try {
    const result = await window.pywebview.api.refresh_printers();
    fillPrinters(result.printers || [], result.defaultPrinter || '');
    syncInventoryStatus(result);
    syncLicenseStatus(result);
    updatePrintReadiness();
  } catch {
    // Keep UI responsive even if refresh fails.
  }
}

function syncInventoryStatus(result) {
  const baseFound = Boolean(result && result.inventoryBaseFound);
  const ready = Boolean(result && result.inventoryReady);
  const count = Number(result && result.inventoryCount ? result.inventoryCount : 0);

  if (!baseFound) {
    els.inventoryStatus.textContent = 'Base não encontrada';
    return;
  }
  if (!ready) {
    els.inventoryStatus.textContent = 'Carregando base...';
    return;
  }
  els.inventoryStatus.textContent = `${count} chapas carregadas`;
}

function syncLicenseStatus(result) {
  const prevBlocked = state.licenseBlocked;
  const prevMessage = state.licenseMessage;
  const blocked = Boolean(result && result.licenseBlocked);
  const connected = Boolean(result && result.licenseConnected);
  const message = String((result && result.licenseMessage) || '').trim();
  state.licenseBlocked = blocked;
  state.licenseConnected = connected;
  state.licenseMessage = message;

  if (blocked && (!prevBlocked || prevMessage !== message)) {
    setStatus(message || 'Licenca bloqueada pelo administrador.');
    toast(message || 'Licenca bloqueada pelo administrador.');
    return;
  }

  if (prevBlocked && !blocked) {
    setStatus('Licenca validada.');
    toast('Licenca liberada pelo administrador.');
  }
}

function startLicenseHeartbeat() {
  if (state.licenseHeartbeatId) return;

  const tick = async () => {
    if (!state.apiReady || !window.pywebview || !window.pywebview.api || !window.pywebview.api.get_license_status) {
      return;
    }
    try {
      const result = await window.pywebview.api.get_license_status();
      syncLicenseStatus(result);
      updatePrintReadiness();
    } catch {
      // Keep UI stable when connection blips.
    }
  };

  tick();
  state.licenseHeartbeatId = window.setInterval(tick, 5000);
}

function updateModeUI(animate = true) {
  els.modeButtons.forEach((button) => button.classList.toggle('is-active', button.dataset.mode === state.mode));
  syncModeThumb();
  const showDelivery = state.mode === 'entrega' || state.mode === 'ambos';
  const showReceipt = state.mode === 'recebimento' || state.mode === 'ambos';
  setModeGroupState(els.deliveryModeBlocks, showDelivery, animate, 'delivery');
  setModeGroupState(els.receiptModeBlocks, showReceipt, animate, 'receipt');
}

function syncModeThumb() {
  if (!els.modeSwitcher || !els.modeThumb) return;

  const activeButton = els.modeButtons.find((button) => button.dataset.mode === state.mode);
  if (!activeButton) return;

  const switcherRect = els.modeSwitcher.getBoundingClientRect();
  const buttonRect = activeButton.getBoundingClientRect();
  const offsetX = buttonRect.left - switcherRect.left;

  els.modeThumb.style.width = `${buttonRect.width}px`;
  els.modeThumb.style.transform = `translateX(${offsetX}px)`;
}

function setModeGroupState(elements, shouldShow, animate, kind) {
  const blocks = elements.filter(Boolean);
  if (!blocks.length) return;

  blocks.forEach((element) => {
    if (element._modeVisibilityTimer) {
      window.clearTimeout(element._modeVisibilityTimer);
      element._modeVisibilityTimer = null;
    }
    element.classList.remove('is-animating-in', 'is-animating-out', 'is-inactive', 'mode-block--delivery', 'mode-block--receipt');
    element.classList.add(kind === 'delivery' ? 'mode-block--delivery' : 'mode-block--receipt');
  });

  if (!animate) {
    blocks.forEach((element) => {
      element.classList.toggle('is-inactive', !shouldShow);
      element.style.pointerEvents = shouldShow ? '' : 'none';
    });
    return;
  }

  if (shouldShow) {
    blocks.forEach((element) => {
      element.style.pointerEvents = '';
      element.classList.remove('is-inactive');
    });
    void blocks[0].offsetWidth;
    blocks.forEach((element, index) => {
      element.style.animationDelay = `${index * 35}ms`;
      element.classList.add('is-animating-in');
      element._modeVisibilityTimer = window.setTimeout(() => {
        element.classList.remove('is-animating-in');
        element.style.animationDelay = '';
        element._modeVisibilityTimer = null;
      }, 320 + index * 35);
    });
    return;
  }

  blocks.forEach((element, index) => {
    element.style.pointerEvents = 'none';
    element.style.animationDelay = `${index * 24}ms`;
    element.classList.add('is-animating-out');
    element._modeVisibilityTimer = window.setTimeout(() => {
      element.classList.add('is-inactive');
      element.classList.remove('is-animating-out');
      element.style.animationDelay = '';
      element._modeVisibilityTimer = null;
    }, 210 + index * 24);
  });
}

function applyModeTheme() {
  els.body.classList.remove('mode-entrega', 'mode-recebimento', 'mode-ambos');
  if (state.mode === 'entrega') {
    els.body.classList.add('mode-entrega');
  } else if (state.mode === 'recebimento') {
    els.body.classList.add('mode-recebimento');
  } else {
    els.body.classList.add('mode-ambos');
  }
}

function hydrateUiPreferences() {
  try {
    const savedCollapsed = JSON.parse(window.localStorage.getItem('guide-collapsed-cards') || '[]');
    if (Array.isArray(savedCollapsed)) {
      state.collapsedCards = new Set(savedCollapsed);
    }
  } catch {}
}

function toggleCardCollapse(cardId) {
  if (state.collapsedCards.has(cardId)) {
    state.collapsedCards.delete(cardId);
  } else {
    state.collapsedCards.add(cardId);
  }
  try {
    window.localStorage.setItem('guide-collapsed-cards', JSON.stringify([...state.collapsedCards]));
  } catch {}
  syncCollapsedCards();
}

function syncCollapsedCards() {
  els.collapseButtons.forEach((button) => {
    const cardId = button.dataset.collapseTarget;
    const card = document.getElementById(cardId);
    if (!card) return;
    const collapsed = state.collapsedCards.has(cardId);
    card.classList.toggle('card--collapsed', collapsed);
    button.textContent = collapsed ? 'Expandir' : 'Ocultar';
    button.setAttribute('aria-expanded', String(!collapsed));
  });
}

function scheduleDraftSave() {
  if (state.autosaveTimer) {
    window.clearTimeout(state.autosaveTimer);
  }
  state.autosaveTimer = window.setTimeout(() => {
    saveDraftNow();
    state.autosaveTimer = null;
  }, 260);
}

function saveDraftNow() {
  const payload = buildPayload();
  const snapshot = {
    mode: payload.mode,
    printerName: payload.printerName,
    copies: payload.copies,
    allowModifiableGuides: payload.allowModifiableGuides,
    deliveryReceiverUnit: payload.deliveryReceiverUnit,
    receiptSenderUnit: payload.receiptSenderUnit,
    deliveryRoom: payload.deliveryRoom,
    receiptRoom: payload.receiptRoom,
    deliveryItems: collectRows('delivery'),
    receiptItems: collectRows('receipt'),
    savedAt: Date.now(),
  };
  try {
    window.localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(snapshot));
  } catch {}
}

function restoreDraftIfAvailable() {
  try {
    const raw = window.localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!raw) return false;
    const draft = JSON.parse(raw);
    if (!draft || typeof draft !== 'object') return false;

    const mode = String(draft.mode || '').trim();
    if (['entrega', 'recebimento', 'ambos'].includes(mode)) {
      state.mode = mode;
    }

    els.deliveryReceiverUnit.value = String(draft.deliveryReceiverUnit || '');
    els.receiptSenderUnit.value = String(draft.receiptSenderUnit || '');
    els.deliveryRoom.value = String(draft.deliveryRoom || '');
    els.receiptRoom.value = String(draft.receiptRoom || '');
    els.copiesInput.value = String(draft.copies || '1');
    els.allowModifiableGuides.checked = Boolean(draft.allowModifiableGuides);

    const preferredPrinter = String(draft.printerName || '').trim();
    if (preferredPrinter) {
      const optionExists = Array.from(els.printerSelect.options).some((option) => option.value === preferredPrinter);
      if (optionExists) {
        els.printerSelect.value = preferredPrinter;
      }
    }

    setRowsFromSnapshot('delivery', Array.isArray(draft.deliveryItems) ? draft.deliveryItems : []);
    setRowsFromSnapshot('receipt', Array.isArray(draft.receiptItems) ? draft.receiptItems : []);
    return true;
  } catch {
    return false;
  }
}

function setRowsFromSnapshot(kind, items) {
  state.rows[kind].forEach((row) => row.element.remove());
  state.rows[kind] = [];

  const source = Array.isArray(items) ? items : [];
  const desiredCount = clamp(source.length || 1, 1, state.maxItemRows);

  for (let i = 0; i < desiredCount; i += 1) {
    addRow(kind, false);
  }

  state.rows[kind].forEach((row, index) => {
    const data = source[index] || {};
    row.patrimonyInput.value = String(data.patrimony || '').trim();
    row.descriptionInput.value = String(data.description || '').trim();
    updateDescriptionIcon(row.descriptionInput, row.descriptionIcon);
    row.patrimonyInput.classList.remove('input--invalid');
  });
}

function buildPayload() {
  return {
    mode: state.mode,
    printerName: els.printerSelect.value,
    copies: els.copiesInput.value,
    allowModifiableGuides: els.allowModifiableGuides.checked,
    deliveryReceiverUnit: els.deliveryReceiverUnit.value,
    receiptSenderUnit: els.receiptSenderUnit.value,
    deliveryRoom: els.deliveryRoom.value,
    receiptRoom: els.receiptRoom.value,
    deliveryItems: collectRows('delivery'),
    receiptItems: collectRows('receipt'),
  };
}

function computeReadiness(payload) {
  const p = payload || buildPayload();
  const requiresDelivery = p.mode === 'entrega' || p.mode === 'ambos';
  const requiresReceipt = p.mode === 'recebimento' || p.mode === 'ambos';
  const copies = Number.parseInt(String(p.copies || '1'), 10);

  const hasPrinter = Boolean(String(p.printerName || '').trim()) && copies >= 1;
  const hasDeliveryCore = !requiresDelivery || (Boolean(String(p.deliveryReceiverUnit || '').trim()) && Boolean(String(p.deliveryRoom || '').trim()));
  const hasReceiptCore = !requiresReceipt || (Boolean(String(p.receiptSenderUnit || '').trim()) && Boolean(String(p.receiptRoom || '').trim()));

  const deliveryValid = !requiresDelivery || validateItems(p.deliveryItems || [], p.allowModifiableGuides);
  const receiptValid = !requiresReceipt || validateItems(p.receiptItems || [], p.allowModifiableGuides);
  const hasActiveItems = (!requiresDelivery || (p.deliveryItems || []).length > 0) && (!requiresReceipt || (p.receiptItems || []).length > 0);

  const itemsValid = hasActiveItems && deliveryValid && receiptValid;
  const ready = hasPrinter && hasDeliveryCore && hasReceiptCore && itemsValid;

  return {
    ready,
    printer: hasPrinter,
    delivery: hasDeliveryCore,
    receipt: hasReceiptCore,
    items: itemsValid,
  };
}

function validateItems(items, allowModifiable) {
  if (!Array.isArray(items) || !items.length) return false;
  return items.every((item) => {
    const patrimony = String(item.patrimony || '').trim();
    const description = String(item.description || '').trim();
    if (!description) return false;
    if (!patrimony && !allowModifiable) return false;
    return true;
  });
}

function updatePrintReadiness() {
  const readiness = computeReadiness();
  const keys = ['printer', 'delivery', 'receipt', 'items'];
  keys.forEach((key) => {
    const element = els.readinessItems[key];
    if (!element) return;
    element.classList.toggle('is-ok', Boolean(readiness[key]));
    element.classList.toggle('is-fail', !readiness[key]);
  });

  if (els.printButton) {
    const locked = state.licenseBlocked;
    els.printButton.disabled = locked || !readiness.ready;
    els.printButton.title = locked
      ? (state.licenseMessage || 'Licenca bloqueada pelo administrador.')
      : (readiness.ready ? '' : 'Preencha os requisitos para liberar a impressão.');
  }

  updateFloatingGuide(readiness);
}

function updateFloatingGuide(readinessInput) {
  const readiness = readinessInput || computeReadiness();
  const steps = els.heroGuideSteps || {};
  const flowOk = ['entrega', 'recebimento', 'ambos'].includes(state.mode);
  const coreOk = readiness.delivery && readiness.receipt;
  const itemsOk = readiness.items;

  const applyStepState = (element, ok) => {
    if (!element) return;
    element.classList.toggle('is-done', ok);
    element.classList.toggle('is-pending', !ok);
  };

  applyStepState(steps.flow, flowOk);
  applyStepState(steps.core, coreOk);
  applyStepState(steps.items, itemsOk);
}

function hydratePrintHistory() {
  try {
    const raw = window.localStorage.getItem(HISTORY_STORAGE_KEY);
    const parsed = JSON.parse(raw || '[]');
    state.printHistory = Array.isArray(parsed) ? parsed.slice(0, HISTORY_LIMIT) : [];
  } catch {
    state.printHistory = [];
  }
}

function persistPrintHistory() {
  try {
    window.localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(state.printHistory.slice(0, HISTORY_LIMIT)));
  } catch {}
}

function addPrintHistoryEntry(payload) {
  const entry = {
    id: `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
    createdAt: Date.now(),
    mode: payload.mode,
    printerName: payload.printerName,
    copies: payload.copies,
    allowModifiableGuides: Boolean(payload.allowModifiableGuides),
    deliveryReceiverUnit: payload.deliveryReceiverUnit,
    receiptSenderUnit: payload.receiptSenderUnit,
    deliveryRoom: payload.deliveryRoom,
    receiptRoom: payload.receiptRoom,
    deliveryItems: payload.deliveryItems || [],
    receiptItems: payload.receiptItems || [],
  };

  state.printHistory = [entry, ...state.printHistory].slice(0, HISTORY_LIMIT);
  persistPrintHistory();
  renderPrintHistory();
}

function renderPrintHistory() {
  if (!els.printHistoryList) return;
  if (!state.printHistory.length) {
    els.printHistoryList.innerHTML = '<div class="history-empty">Nenhuma impressão registrada ainda.</div>';
    return;
  }

  els.printHistoryList.innerHTML = state.printHistory.map((entry) => {
    const date = new Date(entry.createdAt || Date.now()).toLocaleString('pt-BR');
    const labelMode = entry.mode === 'ambos'
      ? 'Entrega + Recebimento'
      : entry.mode === 'entrega'
        ? 'Somente entrega'
        : 'Somente recebimento';
    const deliveryCount = Array.isArray(entry.deliveryItems) ? entry.deliveryItems.length : 0;
    const receiptCount = Array.isArray(entry.receiptItems) ? entry.receiptItems.length : 0;
    return `
      <div class="history-item">
        <div class="history-item__main">
          <strong>${labelMode}</strong>
          <span>${date}</span>
          <span>${deliveryCount} itens entrega • ${receiptCount} itens recebimento</span>
        </div>
        <button class="btn btn--ghost history-item__btn" type="button" data-history-apply="${entry.id}">Repetir</button>
      </div>
    `;
  }).join('');
}

function handleHistoryAction(event) {
  const button = event.target.closest('[data-history-apply]');
  if (!button) return;
  const entryId = button.dataset.historyApply;
  const entry = state.printHistory.find((item) => item.id === entryId);
  if (!entry) return;

  state.mode = ['entrega', 'recebimento', 'ambos'].includes(entry.mode) ? entry.mode : 'ambos';
  els.copiesInput.value = String(entry.copies || '1');
  els.allowModifiableGuides.checked = Boolean(entry.allowModifiableGuides);
  els.deliveryReceiverUnit.value = String(entry.deliveryReceiverUnit || '');
  els.receiptSenderUnit.value = String(entry.receiptSenderUnit || '');
  els.deliveryRoom.value = String(entry.deliveryRoom || '');
  els.receiptRoom.value = String(entry.receiptRoom || '');

  const printer = String(entry.printerName || '').trim();
  if (printer) {
    const optionExists = Array.from(els.printerSelect.options).some((option) => option.value === printer);
    if (optionExists) {
      els.printerSelect.value = printer;
    }
  }

  setRowsFromSnapshot('delivery', Array.isArray(entry.deliveryItems) ? entry.deliveryItems : []);
  setRowsFromSnapshot('receipt', Array.isArray(entry.receiptItems) ? entry.receiptItems : []);

  updateCounters();
  updateModeUI();
  applyModeTheme();
  updatePrintReadiness();
  scheduleDraftSave();
  setStatus('Configuração restaurada do histórico.');
  toast('Dados da última impressão aplicados.');
}

function clearHistory() {
  state.printHistory = [];
  persistPrintHistory();
  renderPrintHistory();
  setStatus('Histórico de impressões limpo.');
}

function addRow(kind, rerender = true) {
  if (state.rows[kind].length >= state.maxItemRows) {
    toast(`Cada guia aceita no máximo ${state.maxItemRows} linhas.`);
    return;
  }

  const row = createRow(kind, state.rows[kind].length);
  state.rows[kind].push(row);
  getRowsContainer(kind).appendChild(row.element);
  updateCounters();
  if (rerender) {
    renderPreview();
    updatePrintReadiness();
    scheduleDraftSave();
  }
}

function removeRow(kind) {
  if (!state.rows[kind].length) return;
  const row = state.rows[kind].pop();
  row.element.remove();
  updateCounters();
  renderPreview();
  updatePrintReadiness();
  scheduleDraftSave();
}

function createRow(kind, index) {
  const element = document.createElement('div');
  element.className = 'item-row item-row--enter';

  const patrimonyInput = document.createElement('input');
  patrimonyInput.type = 'text';
  patrimonyInput.placeholder = `Chapa ${index + 1} (até ${state.patrimonyLength} dígitos)`;

  const descriptionField = document.createElement('div');
  descriptionField.className = 'item-field item-field--description';

  const descriptionInput = document.createElement('input');
  descriptionInput.type = 'text';
  descriptionInput.placeholder = 'Descrição do item';

  const descriptionIcon = document.createElement('span');
  descriptionIcon.className = 'item-type-icon';
  descriptionIcon.setAttribute('aria-hidden', 'true');

  patrimonyInput.addEventListener('input', () => {
    scheduleLookup(patrimonyInput, descriptionInput);
    renderPreview();
    updatePrintReadiness();
    scheduleDraftSave();
  });
  descriptionInput.addEventListener('input', () => {
    updateDescriptionIcon(descriptionInput, descriptionIcon);
    renderPreview();
    updatePrintReadiness();
    scheduleDraftSave();
  });

  descriptionField.append(descriptionInput, descriptionIcon);
  element.append(patrimonyInput, descriptionField);
  requestAnimationFrame(() => {
    element.classList.remove('item-row--enter');
  });
  return { element, patrimonyInput, descriptionInput, descriptionIcon };
}

function getRowsContainer(kind) {
  return kind === 'delivery' ? els.deliveryItems : els.receiptItems;
}

function updateCounters() {
  els.deliveryCount.textContent = `${state.rows.delivery.length}/${state.maxItemRows}`;
  els.receiptCount.textContent = `${state.rows.receipt.length}/${state.maxItemRows}`;
}

function clearItems() {
  [...state.rows.delivery, ...state.rows.receipt].forEach((row) => {
    row.patrimonyInput.value = '';
    row.descriptionInput.value = '';
    row.patrimonyInput.classList.remove('input--invalid');
    updateDescriptionIcon(row.descriptionInput, row.descriptionIcon);
  });
  renderPreview();
  updatePrintReadiness();
  scheduleDraftSave();
  setStatus('Linhas de itens limpas.');
}

function scheduleLookup(patrimonyInput, descriptionInput) {
  const existing = state.lookupTimers.get(patrimonyInput);
  if (existing) clearTimeout(existing);

  const timer = setTimeout(async () => {
    const raw = patrimonyInput.value.trim();
    const digits = raw.replace(/\D/g, '');

    if (!raw) {
      patrimonyInput.classList.remove('input--invalid');
      descriptionInput.value = '';
      updateDescriptionIcon(descriptionInput);
      renderPreview();
      return;
    }

    if (digits.length > state.patrimonyLength) {
      patrimonyInput.classList.add('input--invalid');
      descriptionInput.value = '';
      updateDescriptionIcon(descriptionInput);
      renderPreview();
      return;
    }

    if (!state.apiReady) return;
    const result = await window.pywebview.api.lookup_item(digits);

    const currentDigits = patrimonyInput.value.trim().replace(/\D/g, '');
    if (currentDigits !== digits) {
      return;
    }

    if (result.ok && result.description) {
      descriptionInput.value = result.description;
      patrimonyInput.classList.remove('input--invalid');
    } else {
      descriptionInput.value = '';
      if (digits.length === state.patrimonyLength) {
        patrimonyInput.classList.add('input--invalid');
      } else {
        patrimonyInput.classList.remove('input--invalid');
      }
    }
    updateDescriptionIcon(descriptionInput);
    renderPreview();
    updatePrintReadiness();
    scheduleDraftSave();
  }, 220);

  state.lookupTimers.set(patrimonyInput, timer);
}

function collectRows(kind) {
  return state.rows[kind]
    .map((row) => ({
      patrimony: row.patrimonyInput.value.trim(),
      description: row.descriptionInput.value.trim(),
    }))
    .filter((item) => item.patrimony || item.description);
}

function renderPreview() {
  if (!els.previewStage) {
    return;
  }
  const pages = [];
  if (state.mode === 'entrega' || state.mode === 'ambos') {
    pages.push(renderPage({
      topLeftLabel: 'UA Remetente:',
      topLeftValue: 'DIVISÃO DE PATRIMÔNIO',
      topRightLabel: 'Sala:',
      topRightValue: 'ESTOQUE',
      bottomLeftLabel: 'UA Receptora:',
      bottomLeftValue: els.deliveryReceiverUnit.value.trim(),
      bottomRightLabel: 'Sala:',
      bottomRightValue: els.deliveryRoom.value.trim(),
      items: collectRows('delivery'),
      signatureSide: 'left',
    }));
  }
  if (state.mode === 'recebimento' || state.mode === 'ambos') {
    pages.push(renderPage({
      topLeftLabel: 'UA Remetente:',
      topLeftValue: els.receiptSenderUnit.value.trim(),
      topRightLabel: 'Sala:',
      topRightValue: els.receiptRoom.value.trim(),
      bottomLeftLabel: 'UA Receptora:',
      bottomLeftValue: 'DIVISÃO DE PATRIMÔNIO',
      bottomRightLabel: 'Sala:',
      bottomRightValue: 'ESTOQUE',
      items: collectRows('receipt'),
      signatureSide: 'right',
    }));
  }
  els.previewStage.innerHTML = pages.join('');
}

function renderPage(data) {
  const rows = [...data.items];
  while (rows.length < 21) rows.push({ patrimony: '', description: '' });
  const today = new Date().toLocaleDateString('pt-BR');
  const leftSigned = data.signatureSide === 'left';
  const rightSigned = data.signatureSide === 'right';

  return `
    <article class="page">
      <div class="page-top">
        <div class="brand">
          <div class="brand__logo">Alesp</div>
          <div class="brand__sub">ASSEMBLEIA LEGISLATIVA<br>DO ESTADO DE SÃO PAULO</div>
        </div>
        <div class="page-waves"></div>
      </div>
      <div class="page-title">GUIA DE TRANSFERENCIA DE BENS PATRIMONIAIS - MOVEIS</div>
      <table class="info-table">
        <tr>
          <td>${data.topLeftLabel}</td>
          <td>${escapeHtml(data.topLeftValue || '')}</td>
          <td>${data.topRightLabel}</td>
          <td>${escapeHtml(data.topRightValue || '')}</td>
        </tr>
        <tr>
          <td>${data.bottomLeftLabel}</td>
          <td>${escapeHtml(data.bottomLeftValue || '')}</td>
          <td>${data.bottomRightLabel}</td>
          <td>${escapeHtml(data.bottomRightValue || '')}</td>
        </tr>
      </table>
      <table class="items-table">
        <thead>
          <tr>
            <th style="width:19%;">Nº Patrimônio</th>
            <th>Descrição do bem</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map((row) => `<tr><td>${escapeHtml(row.patrimony || '')}</td><td>${escapeHtml(row.description || '')}</td></tr>`).join('')}
        </tbody>
      </table>
      <div class="signature-grid">
        ${renderSignatureBox(leftSigned ? `Data de envio: ${today}` : 'Data de envio: ____/____/______', leftSigned)}
        ${renderSignatureBox(rightSigned ? `Data de recebimento: ${today}` : 'Data de recebimento: ____/____/______', rightSigned, true)}
      </div>
      <div class="page-footer">PATRIMÔNIO: SALA T16 - TÉRREO - RAMAIS: 6188/6524/6528</div>
    </article>
  `;
}

function renderSignatureBox(title, signed, receipt = false) {
  return `
    <section class="signature-box">
      <div>${title}</div>
      <div>Matrícula: ${signed ? '31356' : '______________________'}</div>
      <div>Nome: ${signed ? 'Felipe Solano Silva Lyra' : '__________________________'}</div>
      <div class="signature-box__sign">${signed ? 'Felipe Solano' : ''}</div>
      <div class="signature-box__footer">${receipt ? 'Assinatura - Receptor' : 'Assinatura - Remetente'}</div>
    </section>
  `;
}

async function printDocuments() {
  if (!state.apiReady) {
    toast('O backend ainda está iniciando. Tente novamente em instantes.');
    return;
  }
  const payload = buildPayload();
  const readiness = computeReadiness(payload);
  if (!readiness.ready) {
    updatePrintReadiness();
    toast('Complete os campos obrigatórios para liberar a impressão.');
    setStatus('Pendências encontradas antes da impressão.');
    return;
  }

  els.printButton.disabled = true;
  setStatus('Validando e enviando para a fila de impressão...');
  try {
    const result = await window.pywebview.api.print_guides(payload);
    if (!result.ok) {
      toast(result.error || 'Não foi possível imprimir.');
      setStatus('Falha ao imprimir.');
      return;
    }
    toast(result.message || 'Impressão enviada para a fila.');
    setStatus('Impressão enviada para a fila.');
    addPrintHistoryEntry(payload);
    saveDraftNow();
  } finally {
    updatePrintReadiness();
  }
}

function setStatus(text) {
  els.statusText.textContent = text;
}

function toast(message) {
  els.toast.textContent = message;
  els.toast.classList.add('is-visible');
  clearTimeout(els.toast.timerId);
  els.toast.timerId = setTimeout(() => els.toast.classList.remove('is-visible'), 3200);
}

function hideStartupOverlay() {
  if (!els.startupOverlay) return;
  window.setTimeout(() => {
    els.startupOverlay.classList.add('is-hidden');
    window.setTimeout(() => {
      if (els.startupOverlay && els.startupOverlay.parentNode) {
        els.startupOverlay.parentNode.removeChild(els.startupOverlay);
      }
    }, 520);
  }, 9500);
}

async function ensureStartupVideoPlayback() {
  if (!els.startupVideo) return;
  els.startupVideo.muted = true;
  els.startupVideo.defaultMuted = true;
  els.startupVideo.setAttribute('muted', 'muted');
  els.startupVideo.playsInline = true;
  els.startupVideo.setAttribute('playsinline', 'playsinline');

  const playWithTimeout = async (timeoutMs = 1600) => {
    const playPromise = els.startupVideo.play();
    const timeoutPromise = new Promise((_, reject) => {
      window.setTimeout(() => reject(new Error('play-timeout')), timeoutMs);
    });
    if (playPromise && typeof playPromise.then === 'function') {
      await Promise.race([playPromise, timeoutPromise]);
    }
  };

  try {
    await playWithTimeout();
  } catch {
    const fallback = els.startupVideo.dataset.fallback || '';
    if (fallback) {
      try {
        els.startupVideo.src = fallback;
        els.startupVideo.load();
        await playWithTimeout(1800);
        return;
      } catch {}
    }
    // If playback remains blocked, keep loading screen style without crashing.
    els.startupVideo.classList.add('startup-overlay__video--blocked');
  }
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function updateDescriptionIcon(descriptionInput, explicitIconElement) {
  const row = explicitIconElement
    ? { descriptionIcon: explicitIconElement }
    : [...state.rows.delivery, ...state.rows.receipt].find((entry) => entry.descriptionInput === descriptionInput);
  if (!row || !row.descriptionIcon) return;

  const icon = inferItemIcon(descriptionInput.value);
  row.descriptionIcon.innerHTML = icon.svg;
  row.descriptionIcon.dataset.kind = icon.kind;
  row.descriptionIcon.title = icon.label;
  row.descriptionIcon.classList.toggle('is-empty', icon.kind === 'empty');
}

function inferItemIcon(text) {
  const value = String(text || '').toLowerCase();
  const map = [
    { kind: 'chair', label: 'Cadeira', terms: ['cadeira', 'poltrona', 'banco'], svg: makeIconSvg('chair') },
    { kind: 'table', label: 'Mesa', terms: ['mesa', 'escrivaninha', 'bancada'], svg: makeIconSvg('table') },
    { kind: 'cabinet', label: 'Armário', terms: ['armario', 'arquivo', 'gaveteiro', 'estante', 'balcao'], svg: makeIconSvg('cabinet') },
    { kind: 'screen', label: 'Monitor', terms: ['monitor', 'tv', 'tela'], svg: makeIconSvg('screen') },
    { kind: 'computer', label: 'Computador', terms: ['cpu', 'computador', 'notebook', 'laptop'], svg: makeIconSvg('computer') },
    { kind: 'printer', label: 'Impressora', terms: ['impressora', 'scanner', 'multifuncional'], svg: makeIconSvg('printer') },
    { kind: 'phone', label: 'Telefone', terms: ['telefone', 'ramal', 'aparelho'], svg: makeIconSvg('phone') },
  ];

  for (const item of map) {
    if (item.terms.some((term) => value.includes(term))) {
      return item;
    }
  }

  if (value.trim()) {
    return { kind: 'generic', label: 'Item patrimonial', svg: makeIconSvg('generic') };
  }

  return { kind: 'empty', label: '', svg: '' };
}

function makeIconSvg(kind) {
  const palette = {
    chair: '#0f7c78',
    table: '#1d5f91',
    cabinet: '#8c5a2e',
    screen: '#3558a8',
    computer: '#245d7a',
    printer: '#6a5acd',
    phone: '#9b5b2a',
    generic: '#6b7d8a',
  };

  const color = palette[kind] || palette.generic;
  const paths = {
    chair: '<path d="M8 3h8v5H8zM7 9h10v4H7zM8 13v4M16 13v4M7 17h10" />',
    table: '<path d="M4 7h16M6 7v3M18 7v3M7 10v7M17 10v7" />',
    cabinet: '<path d="M6 4h12v16H6zM12 4v16M9 9h1M14 9h1M9 14h1M14 14h1" />',
    screen: '<path d="M4 5h16v11H4zM10 19h4M12 16v3" />',
    computer: '<path d="M5 6h9v8H5zM15 8h4v9h-4M9 18h4" />',
    printer: '<path d="M7 5h10v4H7zM5 10h14v6H5zM7 13h10M8 16h8" />',
    phone: '<path d="M9 5h6v14H9zM11 7h2M11 17h2" />',
    generic: '<path d="M6 6h12v12H6zM9 9h6M9 12h6M9 15h4" />',
  };

  return `
    <svg viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">
      ${paths[kind] || paths.generic}
    </svg>
  `;
}
