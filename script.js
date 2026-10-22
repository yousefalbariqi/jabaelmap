/* =========================================================
   جمعية مساجد — خريطة المساجد — المنطق الكامل
   ========================================================= */
'use strict';

/* ---------------------------------------------------------
   1) الإعدادات (config-driven — إضافة/حذف مصدر = سطر واحد)
   --------------------------------------------------------- */
const SOURCES = [{ id: 'association', file: 'association-mosques.xlsx', role: 'mosque', priority: 1 }];

const ALLOWED_STATUS = ['لم يبدأ', 'تم الإنجاز'];
const KSA = { latMin: 16, latMax: 33, lngMin: 34, lngMax: 56 };

const CAT_LABEL = { mosque: 'مساجد', jamie: 'جوامع', land: 'أرضي مخصصة', all: 'الكل' };
const CAT_LABEL_ONE = { mosque: 'مسجد', jamie: 'جامع', land: 'أرض مخصصة' };

/* ---------------------------------------------------------
   2) مساعدات عامة
   --------------------------------------------------------- */
const normStr = v => String(v == null ? '' : v).replace(/﻿/g, '').replace(/\s+/g, ' ').trim();

// طيّ الأحرف مع الحفاظ على الطول (للبحث + التظليل)
const fold = s => normStr(s).toLowerCase()
  .replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/[ىي]/g, 'ي')
  .replace(/ؤ/g, 'و').replace(/ئ/g, 'ي');

// تطبيع كامل للبحث (يزيل التشكيل)
const searchNorm = s => fold(s).replace(/[ً-ْٰ]/g, '');

const escapeHtml = s => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// أرقام لاتينية (إنجليزية) في كل الواجهة
const fmtNum = n => Number(n ?? 0).toLocaleString('en-US');

function pick(row, keys) {
  for (const k of keys) {
    const val = normStr(row[k]);
    if (val) return val;
  }
  return '';
}

function parseCoord(v, kind) {
  const n = Number(normStr(v));
  if (!Number.isFinite(n)) return null;
  if (kind === 'lat' && (n < KSA.latMin || n > KSA.latMax)) return null;
  if (kind === 'lng' && (n < KSA.lngMin || n > KSA.lngMax)) return null;
  return n;
}

function cleanNumber(v) {
  const s = normStr(v);
  return /^\d+(?:\.\d+)?$/.test(s) ? Number(s) : null;
}

function pickArea(row) {
  const cands = [row['space'], row['المساحة التقريبية'], row['مساحة المسجد']]
    .map(cleanNumber)
    .filter(n => n != null && n >= 20);
  return cands.length ? Math.max(...cands) : null;
}

const normalizeTag = v => ({ 'مسجد': 'مسجد', 'جامع': 'جامع' }[normStr(v)] || '');

function classify(role, type) {
  if (role === 'land') return 'land';
  return normStr(type) === 'جامع' ? 'jamie' : 'mosque';
}

/* ---------------------------------------------------------
   3) تطبيع السجل → الشكل الموحّد
   --------------------------------------------------------- */
// كادر المسجد — الخطيب يظهر للجوامع فقط
const STAFF_ROLES = [
  {
    key: 'khateeb', label: 'الخطيب', icon: 'fa-microphone-lines', jamieOnly: true,
    cols: {
      name:     ['اسم الخطيب', 'الخطيب'],
      phone:    ['رقم التواصل الخطيب', 'رقم تواصل الخطيب', 'جوال الخطيب', 'رقم الخطيب', 'هاتف الخطيب'],
      contract: ['نوع عقد الخطيب', 'نوع العقد الخطيب', 'عقد الخطيب'],
    },
  },
  {
    key: 'imam', label: 'الإمام', icon: 'fa-person-praying',
    cols: {
      name:     ['اسم الإمام', 'اسم الامام', 'الإمام', 'الامام'],
      phone:    ['رقم التواصل الإمام', 'رقم تواصل الإمام', 'جوال الإمام', 'جوال الامام', 'رقم الإمام', 'رقم الامام', 'هاتف الإمام'],
      contract: ['نوع عقد الإمام', 'نوع عقد الامام', 'نوع العقد الإمام', 'عقد الإمام', 'عقد الامام'],
    },
  },
  {
    key: 'muezzin', label: 'المؤذن', icon: 'fa-bullhorn',
    cols: {
      name:     ['اسم المؤذن', 'اسم الموذن', 'المؤذن', 'الموذن'],
      phone:    ['رقم التواصل المؤذن', 'رقم تواصل المؤذن', 'جوال المؤذن', 'جوال الموذن', 'رقم المؤذن', 'رقم الموذن', 'هاتف المؤذن'],
      contract: ['نوع عقد المؤذن', 'نوع عقد الموذن', 'نوع العقد المؤذن', 'عقد المؤذن', 'عقد الموذن'],
    },
  },
];
const STAFF_FIELDS = ['name', 'phone', 'contract'];

function pickStaff(row) {
  const out = {};
  for (const role of STAFF_ROLES) {
    out[role.key] = {
      name:     pick(row, role.cols.name),
      phone:    pick(row, role.cols.phone),
      contract: pick(row, role.cols.contract),
    };
  }
  return out;
}

function normalizeRecord(row, source) {
  const name = normStr(row['اسم المسجد']);
  if (!name) return null;
  const lat = parseCoord(row['خط العرض'], 'lat');
  const lng = parseCoord(row['خط الطول'], 'lng');
  return {
    id: normStr(row['رقم المسجد']), _role: source.role, _priority: source.priority,
    category: name.startsWith('جامع') ? 'jamie' : 'mosque',
    type: name.startsWith('جامع') ? 'جامع' : 'مسجد', name,
    neighborhood: normStr(row['الحي']), governorate: '', lat, lng,
    hasCoords: lat != null && lng != null,
    workStatus: normStr(row['حالة الإنجاز']),
    maintenanceItem: String(row['الأعمال المنجزة'] ?? '').trim(),
    facilities: {}, staff: {},
  };
}

/* ---------------------------------------------------------
   4) إزالة التكرار + الدمج على مستوى الحقل
   --------------------------------------------------------- */
function dedupeKey(r) {
  if (r.id) return 'id:' + r.id;
  if (r.hasCoords) return `xy:${r.lat.toFixed(5)}|${r.lng.toFixed(5)}`;
  return 'name:' + fold(r.name);
}

function mergeGroup(group) {
  const sorted = group.slice().sort((a, b) => a._priority - b._priority);
  const base = sorted[0];
  const out = { ...base, facilities: { ...base.facilities }, staff: {} };
  STAFF_ROLES.forEach(role => { out.staff[role.key] = { ...(base.staff?.[role.key] || {}) }; });

  const strFields = ['type', 'name', 'neighborhood', 'governorate', 'yearBuilt',
    'workStatus', 'maintenanceItem', 'projectLink', 'landStatus', 'notes'];
  const numFields = ['lat', 'lng', 'capacity', 'area', 'intendedCapacity'];

  for (const rec of sorted) {
    strFields.forEach(f => { if (!out[f] && rec[f]) out[f] = rec[f]; });
    numFields.forEach(f => { if (out[f] == null && rec[f] != null) out[f] = rec[f]; });
    Object.keys(out.facilities).forEach(f => {
      if (!out.facilities[f] && rec.facilities[f]) out.facilities[f] = rec.facilities[f];
    });
    STAFF_ROLES.forEach(role => {
      const dst = out.staff[role.key];
      const src = rec.staff?.[role.key];
      if (!src) return;
      STAFF_FIELDS.forEach(f => { if (!dst[f] && src[f]) dst[f] = src[f]; });
    });
  }
  out.hasCoords = out.lat != null && out.lng != null;
  out.category = classify(base._role, out.type);
  return out;
}

function dedupeAndMerge(records) {
  const groups = new Map();
  for (const r of records) {
    const k = dedupeKey(r);
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(r);
  }
  return [...groups.values()].map(mergeGroup);
}

/* ---------------------------------------------------------
   5) تحميل البيانات
   --------------------------------------------------------- */
function parseCsv(text) {
  return Papa.parse(text, {
    header: true,
    skipEmptyLines: 'greedy',
    transformHeader: h => h.replace(/^﻿/, '').replace(/\s+/g, ' ').trim(),
    transform: v => (v == null ? '' : String(v).replace(/^﻿/, '').trim()),
  }).data;
}

const DATA_COLUMNS = ['رقم المسجد', 'اسم المسجد', 'الحي', 'حالة الإنجاز', 'الأعمال المنجزة', 'خط الطول', 'خط العرض'];
async function loadData() {
  const src = SOURCES[0];
  const res = await fetch(src.file);
  if (!res.ok) throw new Error('تعذّر تحميل ملف بيانات الجمعية');
  const workbook = XLSX.read(await res.arrayBuffer(), { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const table = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  const headers = (table[0] || []).map(normStr);
  const missing = DATA_COLUMNS.filter(h => !headers.includes(h));
  if (missing.length) throw new Error('أعمدة مفقودة: ' + missing.join('، '));
  const ids = new Set();
  const records = table.slice(1).filter(row => row.some(v => normStr(v))).map((values, index) => {
    const row = Object.fromEntries(headers.map((h, i) => [h, values[i] ?? '']));
    const rec = normalizeRecord(row, src);
    if (!rec || !/^\d+$/.test(rec.id)) throw new Error(`بيانات غير مكتملة في الصف ${index + 2}`);
    if (ids.has(rec.id)) throw new Error(`رقم المسجد مكرر: ${rec.id}`);
    ids.add(rec.id);
    return rec;
  });
  return { records, mosques: records, lands: [] };
}

/* ---------------------------------------------------------
   6) الإحصائيات الحيّة
   --------------------------------------------------------- */
function computeStats(records) {
  const isPlace = r => r.category !== 'land';
  const neigh = new Set(records.filter(isPlace).map(r => fold(r.neighborhood)).filter(Boolean));
  return {
    total:   records.filter(isPlace).length,
    mosque:  records.filter(r => r.category === 'mosque').length,
    jamie:   records.filter(r => r.category === 'jamie').length,
    land:    records.filter(r => r.category === 'land').length,
    landAvailable: records.filter(r => r.category === 'land' && r.landStatus === 'متاحة').length,
    neighborhoods: neigh.size,
    worked:  records.filter(r => r.workStatus === 'تم الإنجاز').length,
  };
}

function renderStats(stats, activeCat, scopeMode) {
  const scope = document.getElementById('statsScope');
  scope.innerHTML = 'تعرض: <b>' + escapeHtml(CAT_LABEL[activeCat] || 'الكل') + '</b>' +
    (scopeMode === 'viewport' ? ' · <b>النطاق المرئي</b>' : '');

  const tiles = [
    { key: 'total',  cat: '',       icon: 'fa-location-dot',  label: 'إجمالي المواقع', wide: true },
    { key: 'mosque', cat: 'mosque', icon: 'fa-mosque',        label: 'مساجد' },
    { key: 'jamie',  cat: 'jamie',  icon: 'fa-place-of-worship', label: 'جوامع' },
    
    
    { key: 'neighborhoods', cat: '', icon: 'fa-city',         label: 'الأحياء المخدومة' },
    { key: 'worked', cat: '',       icon: 'fa-screwdriver-wrench', label: 'تم الإنجاز', wide: true },
  ];

  const dim = c => activeCat !== 'all' && c && c !== activeCat;
  document.getElementById('statsBody').innerHTML = tiles.map(t => `
    <div class="kpi ${t.wide ? 'kpi--wide' : ''} ${dim(t.cat) ? 'is-dim' : ''}" ${t.cat ? `data-cat="${t.cat}"` : ''}>
      <div class="kpi__icon"><i class="fas ${t.icon}"></i></div>
      <div class="kpi__value">${fmtNum(stats[t.key] ?? 0)}</div>
      <div class="kpi__label">${t.label}</div>
    </div>`).join('');
}

/* ---------------------------------------------------------
   7) الخريطة + الخرائط الأساسية + الثيم
   --------------------------------------------------------- */
const map = L.map('map', { center: [26.966, 49.645], zoom: 10, minZoom: 5, maxZoom: 20, zoomControl: false });
L.control.zoom({ position: 'bottomleft' }).addTo(map);

const BASEMAPS = {
  light: L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png?key=cb1_4l7z_1_4b3e8578ddcc25143578052d', {
    attribution: '© OpenStreetMap © CARTO', maxZoom: 20, subdomains: 'abcd',
  }),
  dark: L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png?key=cb1_4l7z_1_4b3e8578ddcc25143578052d', {
    attribution: '© OpenStreetMap © CARTO', maxZoom: 20, subdomains: 'abcd',
  }),
  satellite: L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    attribution: 'Tiles © Esri — World Imagery', maxZoom: 19,
  }),
};

let currentBase = null;
let basemapMode = 'auto'; // 'auto' (يتبع الثيم) | 'satellite'

function resolveBase() {
  if (basemapMode === 'satellite') return 'satellite';
  return document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
}
function applyBasemap() {
  const which = resolveBase();
  if (currentBase === BASEMAPS[which]) return;
  Object.values(BASEMAPS).forEach(l => map.hasLayer(l) && map.removeLayer(l));
  currentBase = BASEMAPS[which];
  currentBase.addTo(map);
}

function initTheme() {
  const saved = localStorage.getItem('theme');
  const theme = saved || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  document.documentElement.dataset.theme = theme;
  updateThemeIcon(theme);
}
function updateThemeIcon(theme) {
  const i = document.querySelector('#themeToggle i');
  if (i) i.className = theme === 'dark' ? 'fas fa-sun' : 'fas fa-moon';
}
function toggleTheme() {
  const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  localStorage.setItem('theme', next);
  updateThemeIcon(next);
  applyBasemap();
}
function toggleBasemap() {
  basemapMode = basemapMode === 'satellite' ? 'auto' : 'satellite';
  document.getElementById('basemapToggle').classList.toggle('is-active', basemapMode === 'satellite');
  applyBasemap();
  showToast(basemapMode === 'satellite' ? 'الخريطة: قمر صناعي' : 'الخريطة: مبسّطة', 'fa-layer-group');
}

/* ---------------------------------------------------------
   8) العلامات + التجميع
   --------------------------------------------------------- */
const PIN_PATH = 'M17 43 C 17 43 3 26 3 15.5 A 14 14 0 1 1 31 15.5 C 31 26 17 43 17 43 Z';

// رموز SVG مخصّصة لكل فئة (بلون الفئة على الرأس الأبيض)
const GLYPH_SVG = {
  mosque: `<g class="pin-glyph">
      <rect x="16.5" y="8.4" width="1" height="2.4" rx="0.5"/>
      <circle cx="17" cy="7.8" r="1"/>
      <path d="M12 19.4 V16.8 a5 5 0 0 1 10 0 V19.4 Z"/>
      <rect x="11.2" y="19.2" width="11.6" height="1.7" rx="0.6"/>
    </g>`,
  jamie: `<g class="pin-glyph">
      <rect x="9.8" y="12.3" width="1.9" height="8.6" rx="0.7"/>
      <circle cx="10.75" cy="11.6" r="1.15"/>
      <rect x="22.3" y="12.3" width="1.9" height="8.6" rx="0.7"/>
      <circle cx="23.25" cy="11.6" r="1.15"/>
      <path d="M13.2 19.6 V17.4 a3.8 3.8 0 0 1 7.6 0 V19.6 Z"/>
      <rect x="12.4" y="19.4" width="9.2" height="1.6" rx="0.6"/>
    </g>`,
  land: `<g class="pin-glyph">
      <rect x="13.6" y="8.6" width="1.3" height="12.2" rx="0.5"/>
      <path d="M14.9 9.1 L21.8 11.2 L14.9 13.4 Z"/>
    </g>`,
};

function markerHtml(record) {
  const cat = record.category;
  const big = cat === 'jamie';
  return `<div class="marker-pin ${big ? 'marker-pin--jamie' : ''}" data-cat="${cat}">
      <span class="marker-label">${escapeHtml(record.name)}</span>
      <svg viewBox="0 0 34 44" aria-hidden="true">
        <path class="pin-body" d="${PIN_PATH}"/>
        <circle cx="17" cy="15.5" r="8.6" fill="#fff"/>
        ${GLYPH_SVG[cat] || GLYPH_SVG.mosque}
      </svg>
    </div>`;
}

function makeMarker(record) {
  const big = record.category === 'jamie';
  const marker = L.marker([record.lat, record.lng], {
    icon: L.divIcon({
      className: 'marker-wrap',
      html: markerHtml(record),
      iconSize: big ? [37, 47] : [34, 44],
      iconAnchor: big ? [18, 47] : [17, 44],
    }),
    category: record.category,
    record,
    keyboard: false,
    title: record.name,
  });
  marker.on('click', () => openDetail(record, marker));
  return marker;
}

const clusterGroup = L.markerClusterGroup({
  showCoverageOnHover: false,
  spiderfyOnMaxZoom: true,
  chunkedLoading: true,
  maxClusterRadius: 55,
  disableClusteringAtZoom: 15,
  iconCreateFunction: cluster => {
    const children = cluster.getAllChildMarkers();
    const count = children.length;
    const by = { mosque: 0, jamie: 0, land: 0 };
    children.forEach(m => { by[m.options.category] = (by[m.options.category] || 0) + 1; });

    const size = count < 10 ? 38 : count < 100 ? 46 : 54;
    const colorVar = { mosque: 'var(--cat-mosque)', jamie: 'var(--cat-jamie)', land: 'var(--cat-land)' };
    let acc = 0; const stops = [];
    ['mosque', 'jamie', 'land'].forEach(c => {
      if (!by[c]) return;
      const from = (acc / count) * 100; acc += by[c];
      const to = (acc / count) * 100;
      stops.push(`${colorVar[c]} ${from}% ${to}%`);
    });
    const ring = stops.length ? `conic-gradient(${stops.join(',')})` : 'var(--accent)';

    return L.divIcon({
      className: 'cluster-wrap',
      html: `<div class="cluster-bubble" style="width:${size}px;height:${size}px;background:${ring}">
               <div class="cluster-bubble__inner">${fmtNum(count)}</div>
             </div>`,
      iconSize: [size, size],
    });
  },
});
map.addLayer(clusterGroup);

/* ---------------------------------------------------------
   9) الحالة + العرض + الفلترة
   --------------------------------------------------------- */
const state = { records: [], mosques: [], lands: [], activeCat: 'all', activeMarker: null,
  displayed: [], statsScope: 'all', landsFilter: 'all' };
const markerById = new Map();

function getFiltered() {
  return state.activeCat === 'all'
    ? state.records
    : state.records.filter(r => r.category === state.activeCat);
}

function renderMarkers(records, fit = true) {
  state.displayed = records;
  clusterGroup.clearLayers();
  markerById.clear();
  const withCoords = records.filter(r => r.hasCoords);
  const markers = withCoords.map(r => {
    const m = makeMarker(r);
    markerById.set(r, m);
    return m;
  });
  clusterGroup.addLayers(markers);

  if (fit && markers.length) {
    map.fitBounds(clusterGroup.getBounds().pad(0.12), { animate: false });
  }
}

function applyFilter(cat, fit = true) {
  state.activeCat = cat;
  // مزامنة الأزرار
  document.querySelectorAll('.seg-btn').forEach(b => {
    const on = b.dataset.cat === cat;
    b.classList.toggle('is-active', on);
    b.setAttribute('aria-checked', on);
  });
  document.querySelectorAll('.legend-item').forEach(b => {
    b.setAttribute('aria-pressed', b.dataset.cat === cat);
  });
  renderMarkers(getFiltered(), fit);
  refreshStats();
}

// أساس الإحصائيات (كل المعروض أو ضمن النطاق المرئي)
function statsBase() {
  let base = state.displayed;
  if (state.statsScope === 'viewport') {
    const b = map.getBounds();
    base = base.filter(r => r.hasCoords && b.contains([r.lat, r.lng]));
  }
  return base;
}
function refreshStats() {
  if (statsCard.hidden) return;
  renderStats(computeStats(statsBase()), state.activeCat, state.statsScope);
}

/* ---------------------------------------------------------
   10) لوحة التفاصيل (تخطيطان)
   --------------------------------------------------------- */
const panel = document.getElementById('detailPanel');

const FACILITIES = [
  { key: 'mensPrayer',   label: 'مصلى الرجال',       icon: 'fa-person-praying', always: true },
  { key: 'womensPrayer', label: 'مصلى النساء',        icon: 'fa-person-praying' },
  { key: 'mensWC',       label: 'دورات مياه الرجال',  icon: 'fa-restroom' },
  { key: 'womensWC',     label: 'دورات مياه النساء',  icon: 'fa-restroom' },
  { key: 'guardRoom',    label: 'غرفة الحارس',        icon: 'fa-user-shield' },
  { key: 'storage',      label: 'مستودع',             icon: 'fa-box-archive' },
  { key: 'quranCircles', label: 'حلقات التحفيظ',      icon: 'fa-book-quran' },
  { key: 'imamHousing',  label: 'سكن الإمام',         icon: 'fa-house' },
  { key: 'muezzinHousing', label: 'سكن المؤذن',       icon: 'fa-house-chimney' },
];

const STATUS_MAP = {
  'تم العمل عليه':  { cls: 'done', label: 'تم العمل عليه' },
  'جاري العمل':     { cls: 'prog', label: 'جاري العمل' },
  'سيتم العمل عليه':{ cls: 'plan', label: 'سيتم العمل عليه' },
};
const LAND_STATUS_MAP = {
  'متاحة':        { cls: 'avail',    label: 'متاحة للبناء' },
  'محجوزة':       { cls: 'reserved', label: 'محجوزة' },
  'قيد التخصيص':  { cls: 'pending',  label: 'قيد التخصيص' },
};

const num = v => (v == null || v === '' ? null : v);
const orDash = v => (num(v) == null ? '—' : escapeHtml(String(v)));

function headerHtml(record) {
  const loc = [record.neighborhood, record.governorate].filter(Boolean).map(escapeHtml).join(' · ') || '—';
  return `
    <div class="detail-header">
      <div class="detail-header__main">
        <span class="cat-chip" data-cat="${record.category}">
          <span class="legend-dot" data-cat="${record.category}"></span>${CAT_LABEL_ONE[record.category]}
        </span>
        <h2 class="detail-title" id="dpTitle">${escapeHtml(record.name)}</h2>
        <p class="detail-sub"><i class="fas fa-location-dot"></i> ${loc}</p>
      </div>
      <button class="detail-close" id="detailClose" aria-label="إغلاق"><i class="fas fa-times"></i></button>
    </div>`;
}

// زر الموقع → يفتح خرائط جوجل على إحداثيات المسجد مباشرة
function ctaHtml(record) {
  const label = record.category === 'jamie' ? 'موقع الجامع' : 'موقع المسجد';
  if (!record.hasCoords) {
    return `<span class="detail-cta is-disabled" aria-disabled="true">
        <i class="fas fa-location-dot"></i> الموقع غير متوفر
      </span>`;
  }
  const url = `https://www.google.com/maps/search/?api=1&query=${record.lat},${record.lng}`;
  return `<a class="detail-cta" href="${url}" target="_blank" rel="noopener">
      <i class="fas fa-location-dot"></i> ${label}
      <i class="fas fa-arrow-up-right-from-square detail-cta__ext" aria-hidden="true"></i>
    </a>`;
}

// كادر المسجد — الخطيب للجوامع فقط
function staffHtml(record) {
  const roles = STAFF_ROLES.filter(r => !r.jamieOnly || record.category === 'jamie');

  const cards = roles.map(role => {
    const person = record.staff?.[role.key] || {};
    const rows = [
      { icon: 'fa-user',           label: 'الاسم',       value: person.name },
      { icon: 'fa-phone',          label: 'رقم التواصل', value: person.phone, tel: true },
      { icon: 'fa-file-signature', label: 'نوع العقد',   value: person.contract },
    ].map(row => {
      let value;
      if (num(row.value) == null) {
        value = `<span class="staff-row__value is-empty">—</span>`;
      } else if (row.tel) {
        const dial = String(row.value).replace(/[^\d+]/g, '');
        value = dial
          ? `<a class="staff-row__value staff-row__value--tel" href="tel:${escapeHtml(dial)}">${escapeHtml(row.value)}</a>`
          : `<span class="staff-row__value staff-row__value--tel">${escapeHtml(row.value)}</span>`;
      } else {
        value = `<span class="staff-row__value">${escapeHtml(row.value)}</span>`;
      }
      return `<div class="staff-row">
          <span class="staff-row__label"><i class="fas ${row.icon}"></i> ${row.label}</span>
          ${value}
        </div>`;
    }).join('');

    const filled = STAFF_FIELDS.some(f => num(person[f]) != null);
    return `<div class="staff-card" data-filled="${filled}">
        <div class="staff-card__head">
          <span class="staff-card__icon"><i class="fas ${role.icon}"></i></span>
          <span class="staff-card__role">${role.label}</span>
        </div>
        <div class="staff-rows">${rows}</div>
      </div>`;
  }).join('');

  const title = record.category === 'jamie'
  ? 'منسوبي الجامع'
  : 'منسوبي المسجد';
  return `<h3 class="detail-section-title"><i class="fas fa-users"></i> ${title}</h3>
    <div class="staff-cards">${cards}</div>`;
}

function renderMosqueDetail(record) {
  const cls = record.workStatus === 'تم الإنجاز' ? 'done' : 'plan';
  return headerHtml(record) + `
    <div class="status-row"><span class="status-row__label">حالة الإنجاز:</span>
      <span class="status-chip" data-status="${cls}">${escapeHtml(record.workStatus || 'غير محدد')}</span>
    </div>
    <div class="detail-kpis"><div class="detail-kpi"><div class="detail-kpi__value">${escapeHtml(record.id)}</div><div class="detail-kpi__label">رقم المسجد</div></div></div>
    <h3 class="detail-section-title"><i class="fas fa-screwdriver-wrench"></i> الأعمال المنجزة</h3>
    <div class="notes-box association-works">${escapeHtml(record.maintenanceItem || 'لم تُسجل أعمال منجزة.')}</div>
    ${ctaHtml(record)}`;
}

function renderLandDetail(record) {
  const st = LAND_STATUS_MAP[record.landStatus] || { cls: 'none', label: 'غير محدد' };
  const intended = record.intendedType ? CAT_LABEL_ONE[classify('mosque', record.intendedType)] : '—';
  const notes = record.notes
    ? `<h3 class="detail-section-title"><i class="fas fa-note-sticky"></i> ملاحظات</h3>
       <div class="notes-box">${escapeHtml(record.notes)}</div>` : '';
  const cta = record.projectLink
    ? `<a class="detail-cta" href="${escapeHtml(record.projectLink)}" target="_blank" rel="noopener"><i class="fas fa-link"></i> رابط المشروع</a>` : '';

  return headerHtml(record) + `
    <div class="status-row">
      <span class="status-row__label">حالة الأرض:</span>
      <span class="status-chip" data-status="${st.cls}">${st.label}</span>
    </div>
    <div class="detail-kpis">
      <div class="detail-kpi"><div class="detail-kpi__value">${orDash(record.area)} <small>م²</small></div><div class="detail-kpi__label">المساحة التقريبية</div></div>
      <div class="detail-kpi"><div class="detail-kpi__value">${orDash(record.intendedCapacity)} <small>مصلٍّ</small></div><div class="detail-kpi__label">السعة المقترحة</div></div>
      <div class="detail-kpi"><div class="detail-kpi__value">${escapeHtml(intended)}</div><div class="detail-kpi__label">النوع المقترح</div></div>
    </div>
    ${notes}
    ${cta}`;
}

let lastFocused = null;
function openDetail(record, marker) {
  document.getElementById('detailBody').innerHTML =
    record.category === 'land' ? renderLandDetail(record) : renderMosqueDetail(record);

  document.getElementById('detailClose').addEventListener('click', closeDetail);
  panel.hidden = false;
  requestAnimationFrame(() => panel.classList.add('is-open'));
  document.body.classList.add('panel-open');

  // العلامة النشطة
  setActiveMarker(marker || markerById.get(record) || null);
  lastFocused = document.activeElement;
  panel.querySelector('.detail-close').focus();

  if (record.hasCoords) {
    const fly = prefersReducedMotion() ? 'setView' : 'flyTo';
    map[fly]([record.lat, record.lng], Math.max(map.getZoom(), 16), { duration: 0.6 });
  }
}

function closeDetail() {
  panel.classList.remove('is-open');
  document.body.classList.remove('panel-open');
  setActiveMarker(null);
  setTimeout(() => { panel.hidden = true; }, 260);
  if (lastFocused && lastFocused.focus) lastFocused.focus();
  setTimeout(() => map.invalidateSize(), 280);
}

function setActiveMarker(marker) {
  if (state.activeMarker) {
    const el = state.activeMarker.getElement();
    el && el.querySelector('.marker-pin')?.classList.remove('marker-pin--active');
  }
  state.activeMarker = marker;
  if (marker) {
    const el = marker.getElement();
    el && el.querySelector('.marker-pin')?.classList.add('marker-pin--active');
  }
}

const prefersReducedMotion = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;

/* ---------------------------------------------------------
   11) البحث الذكي + الاقتراحات
   --------------------------------------------------------- */
const searchInput = document.getElementById('searchInput');
const acList = document.getElementById('acList');
const searchClear = document.getElementById('searchClear');
const searchPill = document.querySelector('.search-pill');
let acItems = [];        // العناصر القابلة للاختيار
let acActiveIdx = -1;

function searchMatches(query) {
  const q = searchNorm(query);
  if (!q) return [];
  return state.records.filter(r =>
    searchNorm(r.name).includes(q) ||
    searchNorm(r.neighborhood).includes(q) ||
    searchNorm(r.governorate).includes(q)
  );
}

function highlight(text, query) {
  const fName = fold(text), fQ = fold(query);
  const idx = fQ ? fName.indexOf(fQ) : -1;
  if (idx < 0) return escapeHtml(text);
  return escapeHtml(text.slice(0, idx)) +
    '<mark>' + escapeHtml(text.slice(idx, idx + query.length)) + '</mark>' +
    escapeHtml(text.slice(idx + query.length));
}

function renderAutocomplete(query) {
  const results = searchMatches(query);
  acItems = [];
  acActiveIdx = -1;

  if (!results.length) {
    acList.innerHTML = `<li class="ac-empty">لا توجد نتائج مطابقة</li>`;
    openAc();
    return;
  }

  const cats = ['mosque', 'jamie', 'land'];
  let html = '';
  cats.forEach(cat => {
    const group = results.filter(r => r.category === cat);
    if (!group.length) return;
    html += `<li class="ac-group__label">${CAT_LABEL[cat]} (${fmtNum(group.length)})</li>`;
    group.slice(0, 5).forEach(r => {
      const i = acItems.push(r) - 1;
      const sub = [r.neighborhood, r.governorate].filter(Boolean).map(escapeHtml).join(' · ');
      html += `<li class="ac-item" role="option" id="ac-${i}" data-idx="${i}" data-cat="${cat}">
          <span class="ac-dot"></span>
          <span class="ac-name">${highlight(r.name, query)}</span>
          <span class="ac-sub">${sub}</span>
          <span class="ac-chip" data-cat="${cat}">${CAT_LABEL_ONE[cat]}</span>
        </li>`;
    });
    if (group.length > 5) {
      html += `<li class="ac-more" data-cat="${cat}">عرض كل نتائج ${CAT_LABEL[cat]} (${fmtNum(group.length)})</li>`;
    }
  });
  acList.innerHTML = html;
  openAc();
}

function openAc() { acList.hidden = false; searchPill.setAttribute('aria-expanded', 'true'); }
function closeAc() { acList.hidden = true; searchPill.setAttribute('aria-expanded', 'false'); acActiveIdx = -1; }

function setAcActive(idx) {
  const nodes = acList.querySelectorAll('.ac-item');
  nodes.forEach(n => n.classList.remove('is-active'));
  if (idx < 0 || idx >= nodes.length) { acActiveIdx = -1; searchInput.removeAttribute('aria-activedescendant'); return; }
  acActiveIdx = idx;
  const node = acList.querySelector(`.ac-item[data-idx="${idx}"]`);
  if (node) { node.classList.add('is-active'); node.scrollIntoView({ block: 'nearest' }); searchInput.setAttribute('aria-activedescendant', 'ac-' + idx); }
}

function chooseRecord(record) {
  closeAc();
  searchInput.value = record.name;
  searchClear.hidden = false;
  openDetail(record);
}

let searchTimer = null;
searchInput.addEventListener('input', () => {
  const q = searchInput.value.trim();
  searchClear.hidden = !q;
  clearTimeout(searchTimer);
  if (!q) { closeAc(); return; }
  searchTimer = setTimeout(() => renderAutocomplete(q), 180);
});

searchInput.addEventListener('keydown', e => {
  const nodes = acList.querySelectorAll('.ac-item');
  if (e.key === 'ArrowDown') { e.preventDefault(); if (!acList.hidden) setAcActive(Math.min(acActiveIdx + 1, nodes.length - 1)); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); setAcActive(Math.max(acActiveIdx - 1, 0)); }
  else if (e.key === 'Enter') {
    e.preventDefault();
    if (acActiveIdx >= 0 && acItems[acActiveIdx]) chooseRecord(acItems[acActiveIdx]);
    else runSearch();
  } else if (e.key === 'Escape') {
    if (!acList.hidden) closeAc();
    else { searchInput.value = ''; searchClear.hidden = true; resetSearch(); }
  }
});

acList.addEventListener('click', e => {
  const item = e.target.closest('.ac-item');
  if (item) { chooseRecord(acItems[+item.dataset.idx]); return; }
  const more = e.target.closest('.ac-more');
  if (more) { closeAc(); applyFilter(more.dataset.cat); }
});

function runSearch() {
  const results = searchMatches(searchInput.value.trim());
  closeAc();
  if (!results.length) { showToast('لم يتم العثور على نتائج مطابقة', 'fa-circle-exclamation', true); return; }
  if (results.length === 1) { chooseRecord(results[0]); return; }
  const withCoords = results.filter(r => r.hasCoords);
  renderMarkers(results, false);
  if (withCoords.length) map.fitBounds(clusterGroup.getBounds().pad(0.12));
  showToast(`${fmtNum(results.length)} نتيجة`, 'fa-magnifying-glass');
}

function resetSearch() {
  closeAc();
  applyFilter(state.activeCat);
}

searchClear.addEventListener('click', () => {
  searchInput.value = '';
  searchClear.hidden = true;
  searchInput.focus();
  resetSearch();
});

/* ---------------------------------------------------------
   12) عناصر التحكم (الفلتر، المفتاح، الأدوات)
   --------------------------------------------------------- */
document.getElementById('categoryFilter').addEventListener('click', e => {
  const btn = e.target.closest('.seg-btn');
  if (btn) applyFilter(btn.dataset.cat);
});

document.getElementById('legendItems').addEventListener('click', e => {
  const btn = e.target.closest('.legend-item');
  if (!btn) return;
  const cat = btn.dataset.cat;
  applyFilter(state.activeCat === cat ? 'all' : cat);
});

document.getElementById('legendToggle').addEventListener('click', () => {
  const legend = document.getElementById('legend');
  const collapsed = legend.classList.toggle('is-collapsed');
  document.getElementById('legendToggle').setAttribute('aria-expanded', !collapsed);
});

document.getElementById('themeToggle').addEventListener('click', toggleTheme);
document.getElementById('basemapToggle').addEventListener('click', toggleBasemap);
document.getElementById('resetView').addEventListener('click', () => {
  searchInput.value = ''; searchClear.hidden = true; closeDetail();
  applyFilter('all', true);
});

// الإحصائيات
const statsCard = document.getElementById('statsCard');
const statsToggle = document.getElementById('statsToggle');
function toggleStats(show) {
  const open = show ?? statsCard.hidden;
  statsCard.hidden = !open;
  statsToggle.classList.toggle('is-active', open);
  statsToggle.setAttribute('aria-expanded', open);
  if (open) { closeLands(); refreshStats(); }
}
statsToggle.addEventListener('click', () => toggleStats());
document.getElementById('statsClose').addEventListener('click', () => toggleStats(false));

// تبديل النطاق المرئي
const statsScopeToggle = document.getElementById('statsScopeToggle');
statsScopeToggle.addEventListener('click', () => {
  state.statsScope = state.statsScope === 'viewport' ? 'all' : 'viewport';
  const on = state.statsScope === 'viewport';
  statsScopeToggle.setAttribute('aria-pressed', on);
  statsScopeToggle.classList.toggle('is-active', on);
  refreshStats();
});
map.on('moveend', () => { if (state.statsScope === 'viewport') refreshStats(); });

// تصدير CSV
document.getElementById('exportBtn').addEventListener('click', exportCSV);
function exportCSV() {
  const rows = state.displayed;
  if (!rows.length) { showToast('لا توجد مواقع للتصدير', 'fa-circle-exclamation', true); return; }
  const headers = DATA_COLUMNS;
  const esc = v => {
    let value = String(v ?? '');
    if (/^[=+@\-\t\r]/.test(value)) value = "'" + value;
    return '"' + value.replace(/"/g, '""') + '"';
  };
  const lines = [headers.map(esc).join(',')];
  rows.forEach(r => lines.push([r.id, r.name, r.neighborhood, r.workStatus,
    r.maintenanceItem, r.lng ?? '', r.lat ?? ''].map(esc).join(',')));
  const csv = '﻿' + lines.join('\r\n'); // BOM لتوافق Excel مع العربي
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `مواقع-${CAT_LABEL[state.activeCat] || 'الكل'}.csv`;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  showToast(`تم تصدير ${fmtNum(rows.length)} موقعًا`, 'fa-file-csv');
}

// لوحة الأراضي المخصصة
const landsPanel = document.getElementById('landsPanel');
const landsToggle = document.getElementById('landsToggle');
function landsFiltered() {
  return state.landsFilter === 'all' ? state.lands : state.lands.filter(l => l.landStatus === state.landsFilter);
}
function renderLands() {
  const list = document.getElementById('landsList');
  const lands = landsFiltered();
  if (!lands.length) { list.innerHTML = '<div class="land-card__empty">لا توجد أراضٍ بهذا التصنيف</div>'; return; }
  list.innerHTML = lands.map((l, i) => {
    const st = LAND_STATUS_MAP[l.landStatus] || { cls: 'none', label: 'غير محدد' };
    const loc = [l.neighborhood, l.governorate].filter(Boolean).map(escapeHtml).join(' · ') || '—';
    return `<button class="land-card" data-idx="${i}">
        <div class="land-card__top">
          <span class="land-card__name">${escapeHtml(l.name)}</span>
          <span class="status-chip" data-status="${st.cls}">${st.label}</span>
        </div>
        <div class="land-card__meta">
          <span><i class="fas fa-location-dot"></i> ${loc}</span>
          <span><i class="fas fa-ruler-combined"></i> ${l.area != null ? l.area + ' م²' : '—'}</span>
          <span><i class="fas fa-users"></i> ${l.intendedCapacity != null ? l.intendedCapacity + ' مصلٍّ' : '—'}</span>
        </div>
      </button>`;
  }).join('');
}
function toggleLands(show) {
  const open = show ?? landsPanel.hidden;
  landsPanel.hidden = !open;
  landsToggle.classList.toggle('is-active', open);
  landsToggle.setAttribute('aria-expanded', open);
  if (open) { toggleStats(false); renderLands(); }
}
function closeLands() {
  if (landsPanel.hidden) return;
  landsPanel.hidden = true;
  landsToggle.classList.remove('is-active');
  landsToggle.setAttribute('aria-expanded', false);
}
landsToggle.addEventListener('click', () => toggleLands());
document.getElementById('landsClose').addEventListener('click', () => toggleLands(false));
document.getElementById('landsFilter').addEventListener('click', e => {
  const btn = e.target.closest('.chip-btn'); if (!btn) return;
  state.landsFilter = btn.dataset.ls;
  document.querySelectorAll('#landsFilter .chip-btn').forEach(b => b.classList.toggle('is-active', b === btn));
  renderLands();
});
document.getElementById('landsList').addEventListener('click', e => {
  const card = e.target.closest('.land-card'); if (!card) return;
  const land = landsFiltered()[+card.dataset.idx];
  if (land) openDetail(land);
});

/* ---------------------------------------------------------
   13) أحداث عامة (Esc، نقر خارجي)
   --------------------------------------------------------- */
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    if (!panel.hidden) closeDetail();
    else if (!landsPanel.hidden) toggleLands(false);
    else if (!statsCard.hidden) toggleStats(false);
  }
});

document.addEventListener('click', e => {
  if (!acList.hidden && !searchPill.contains(e.target) && !acList.contains(e.target)) closeAc();
}, true);

map.on('click', () => { if (!panel.hidden) closeDetail(); });

/* ---------------------------------------------------------
   14) التنبيهات (toast)
   --------------------------------------------------------- */
function showToast(msg, icon = 'fa-circle-info', isError = false) {
  const wrap = document.getElementById('toastWrap');
  const el = document.createElement('div');
  el.className = 'toast' + (isError ? ' toast--error' : '');
  el.setAttribute('role', 'status');
  el.innerHTML = `<i class="fas ${icon}"></i> ${escapeHtml(msg)}`;
  wrap.appendChild(el);
  setTimeout(() => {
    el.style.opacity = '0';
    setTimeout(() => el.remove(), 300);
  }, 3000);
}

/* ---------------------------------------------------------
   15) الإقلاع
   --------------------------------------------------------- */
function hideLoading() {
  const ov = document.getElementById('loadingOverlay');
  ov.classList.add('is-hidden');
  setTimeout(() => { ov.hidden = true; }, 300);
}

(async function init() {
  initTheme();
  applyBasemap();

  try {
    const data = await loadData();
    Object.assign(state, data);
    applyFilter('all', true);
    hideLoading();
    const n = state.records.filter(r => r.hasCoords).length;
    showToast(`تم تحميل ${fmtNum(n)} موقعًا`, 'fa-circle-check');
  } catch (err) {
    console.error(err);
    document.getElementById('loadingText').textContent = 'تعذّر تحميل البيانات: ' + err.message;
    showToast('حدث خطأ أثناء تحميل البيانات', 'fa-triangle-exclamation', true);
  }
})();
