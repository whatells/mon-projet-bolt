// Performance-focused endpoints and helpers for CRM UI
const BOOTSTRAP_KEY = 'BOOTSTRAP_JSON_v1';
const STOCK_HEADERS = ['Date entrée','SKU','Titre','Photos','Catégorie','Marque','Taille','État','Prix achat (link)','Statut','Plateforme','Réf Achat','Favoris','Offres','Notes'];
const VENTES_HEADERS = ['Date','Plateforme','Titre','Prix','Frais/Comm','Frais port','Acheteur','SKU','Marge brute','Marge nette'];
const DASHBOARD_CACHE_KEY = 'cache:dashboard:v1';
const STOCK_CACHE_KEY = 'cache:stock_all:v1';
const VENTES_CACHE_KEY = 'cache:ventes_all:v1';
const BOOTSTRAP_CACHE_KEY = 'cache:bootstrap_json:v1';
const CACHE_TTL_SECONDS = 600;
const STOCK_PAGE_SIZE = 20;
const VENTES_PAGE_SIZE = 20;

function timed(label, fn) {
  const started = Date.now();
  try {
    return fn();
  } finally {
    const elapsed = Date.now() - started;
    console.log(`[perf] ${label} server ms=${elapsed}`);
  }
}

function openCRM() {
  return timed('openCRM', () => {
    const html = buildUiHtml_()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    SpreadsheetApp.getUi().showModalDialog(html, '🚀 CRM - Interface Principale');
  });
}

function buildUiHtml_() {
  const payloadString = getBootstrapJson_();
  const t = HtmlService.createTemplateFromFile('Index');
  t.BOOTSTRAP_JSON = payloadString || 'null';
  return t.evaluate()
    .setWidth(1200)
    .setHeight(700)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cronRecomputeBootstrap() {
  return timed('cronRecomputeBootstrap', () => {
    const payload = buildBootstrapPayload_();
    const json = JSON.stringify(payload);
    setBootstrapJson_(json);
    purgeStockCache_();
    purgeVentesCache_();
    softExpireDashboard_();
    return { ok: true, ts: payload.ts };
  });
}

function buildBootstrapPayload_() {
  const ts = Date.now();
  const ss = SpreadsheetApp.getActive();
  const dashboard = timed('bootstrap:dashboard', () => buildDashboardData_(ss));
  const stock = timed('bootstrap:stock', () => buildStockBootstrap_(ss));
  const ventes = timed('bootstrap:ventes', () => buildVentesBootstrap_(ss));
  const config = timed('bootstrap:config', () => buildConfigBootstrap_());
  const logs = timed('bootstrap:logs', () => buildLogsBootstrap_(ss));

  return {
    ts: ts,
    kpis: dashboard.kpis,
    stock: stock,
    ventes: ventes,
    config: config,
    logs: logs
  };
}

function buildDashboardData_(ss) {
  const sh = ss.getSheetByName('Dashboard');
  if (!sh) return { kpis: [] };
  const values = sh.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return { kpis: [] };
  }
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const key = values[i][0];
    if (!key) continue;
    rows.push([key, values[i][1]]);
  }
  return { kpis: rows };
}

function buildStockBootstrap_(ss) {
  const sh = ss.getSheetByName('Stock');
  if (!sh) {
    return { total: 0, pageSize: STOCK_PAGE_SIZE, page1: [] };
  }
  const values = sh.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return { total: 0, pageSize: STOCK_PAGE_SIZE, page1: [] };
  }
  const headers = values[0].map(h => String(h || '').trim());
  const pick = makePicker_(headers, STOCK_HEADERS);
  const body = values.slice(1);
  const mapped = body.map(pick);
  return {
    total: mapped.length,
    pageSize: STOCK_PAGE_SIZE,
    page1: mapped.slice(0, STOCK_PAGE_SIZE)
  };
}

function buildVentesBootstrap_(ss) {
  const sh = ss.getSheetByName('Ventes');
  if (!sh) {
    return { total: 0, pageSize: VENTES_PAGE_SIZE, page1: [] };
  }
  const values = sh.getDataRange().getValues();
  if (!values || values.length <= 1) {
    return { total: 0, pageSize: VENTES_PAGE_SIZE, page1: [] };
  }
  const headers = values[0].map(h => String(h || '').trim());
  const pick = makePicker_(headers, VENTES_HEADERS);
  const body = values.slice(1);
  const mapped = body.map(pick);
  return {
    total: mapped.length,
    pageSize: VENTES_PAGE_SIZE,
    page1: mapped.slice(0, VENTES_PAGE_SIZE)
  };
}

function buildConfigBootstrap_() {
  if (typeof getKnownConfig === 'function') {
    try {
      return getKnownConfig();
    } catch (err) {
      console.warn('getKnownConfig failed for bootstrap', err);
    }
  }
  return [];
}

function buildLogsBootstrap_(ss) {
  const sh = ss.getSheetByName('Logs');
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const take = Math.min(50, last - 1);
  const start = last - take + 1;
  const rows = sh.getRange(start, 1, take, 5).getValues();
  return rows;
}

function getBootstrapJson_() {
  const cache = CacheService.getDocumentCache();
  const cached = cache ? cache.get(BOOTSTRAP_CACHE_KEY) : null;
  if (cached) {
    return cached;
  }
  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty(BOOTSTRAP_KEY);
  if (stored) {
    cachePut_(cache, BOOTSTRAP_CACHE_KEY, stored, CACHE_TTL_SECONDS);
    return stored;
  }
  const payload = buildBootstrapPayload_();
  const json = JSON.stringify(payload);
  setBootstrapJson_(json);
  cachePut_(cache, BOOTSTRAP_CACHE_KEY, json, CACHE_TTL_SECONDS);
  return json;
}

function setBootstrapJson_(json) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(BOOTSTRAP_KEY, json);
  const cache = CacheService.getDocumentCache();
  cachePut_(cache, BOOTSTRAP_CACHE_KEY, json, CACHE_TTL_SECONDS);
  try {
    const parsed = JSON.parse(json);
    const kpis = parsed && parsed.kpis ? parsed.kpis : [];
    cachePut_(cache, DASHBOARD_CACHE_KEY, JSON.stringify({ kpis: kpis }), CACHE_TTL_SECONDS);
  } catch (err) {
    console.warn('Unable to seed dashboard cache', err);
  }
}

function cachePut_(cache, key, value, ttl) {
  if (!cache || typeof value !== 'string') return;
  if (value.length > 90000) return;
  cache.put(key, value, ttl);
}

function makePicker_(headers, wanted) {
  const map = {};
  wanted.forEach((name, idx) => {
    const pos = headers.findIndex(h => h === name);
    map[idx] = pos;
  });
  return function (row) {
    return wanted.map((_, idx) => {
      const pos = map[idx];
      return pos >= 0 ? row[pos] : '';
    });
  };
}

function toNum_(value) {
  if (value == null || value === '') return 0;
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function fmtEuro_(value) {
  const num = toNum_(value);
  if (!num) return '';
  return Utilities.formatString('%s €', num.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
}

function getStockAllRows_() {
  const cache = CacheService.getDocumentCache();
  const cached = cache ? cache.get(STOCK_CACHE_KEY) : null;
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      if (cache) cache.remove(STOCK_CACHE_KEY);
    }
  }
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Stock');
  if (!sh) {
    cachePut_(cache, STOCK_CACHE_KEY, '[]', CACHE_TTL_SECONDS);
    return [];
  }
  const values = sh.getDataRange().getValues();
  if (!values || values.length <= 1) {
    cachePut_(cache, STOCK_CACHE_KEY, '[]', CACHE_TTL_SECONDS);
    return [];
  }
  const headers = values[0].map(h => String(h || '').trim());
  const pick = makePicker_(headers, STOCK_HEADERS);
  const rows = values.slice(1).map(pick);
  cachePut_(cache, STOCK_CACHE_KEY, JSON.stringify(rows), CACHE_TTL_SECONDS);
  return rows;
}

function getVentesAllRows_() {
  const cache = CacheService.getDocumentCache();
  const cached = cache ? cache.get(VENTES_CACHE_KEY) : null;
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      if (cache) cache.remove(VENTES_CACHE_KEY);
    }
  }
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Ventes');
  if (!sh) {
    cachePut_(cache, VENTES_CACHE_KEY, '[]', CACHE_TTL_SECONDS);
    return [];
  }
  const values = sh.getDataRange().getValues();
  if (!values || values.length <= 1) {
    cachePut_(cache, VENTES_CACHE_KEY, '[]', CACHE_TTL_SECONDS);
    return [];
  }
  const headers = values[0].map(h => String(h || '').trim());
  const pick = makePicker_(headers, VENTES_HEADERS);
  const rows = values.slice(1).map(pick);
  cachePut_(cache, VENTES_CACHE_KEY, JSON.stringify(rows), CACHE_TTL_SECONDS);
  return rows;
}

function getDashboardCached_() {
  const cache = CacheService.getDocumentCache();
  const cached = cache ? cache.get(DASHBOARD_CACHE_KEY) : null;
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      if (cache) cache.remove(DASHBOARD_CACHE_KEY);
    }
  }
  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty(BOOTSTRAP_KEY);
  if (stored) {
    try {
      const parsed = JSON.parse(stored);
      const res = { kpis: parsed && parsed.kpis ? parsed.kpis : [] };
      cachePut_(cache, DASHBOARD_CACHE_KEY, JSON.stringify(res), CACHE_TTL_SECONDS);
      return res;
    } catch (err) {
      console.warn('Failed to parse stored bootstrap JSON', err);
    }
  }
  const computed = buildDashboardData_(SpreadsheetApp.getActive());
  cachePut_(cache, DASHBOARD_CACHE_KEY, JSON.stringify(computed), CACHE_TTL_SECONDS);
  return computed;
}

function softExpireDashboard_() {
  const cache = CacheService.getDocumentCache();
  if (cache) cache.remove(DASHBOARD_CACHE_KEY);
}

function purgeStockCache_() {
  const cache = CacheService.getDocumentCache();
  if (cache) cache.remove(STOCK_CACHE_KEY);
}

function purgeVentesCache_() {
  const cache = CacheService.getDocumentCache();
  if (cache) cache.remove(VENTES_CACHE_KEY);
}

function setupTriggers() {
  ScriptApp.newTrigger('cronRecomputeBootstrap').timeBased().everyHours(1).create();
}
