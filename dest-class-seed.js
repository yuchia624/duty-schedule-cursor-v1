/**
 * 航線分類規則與夏季班表種子（長榮立榮 2026 夏季班表）
 */
(function (global) {
  const SEED_VERSION = '2026-summer-homeline-v4';

  /** 區域分類主檔（表格第二欄標籤） */
  const REGION_CATALOG = [
    { id: 'sea', label: '東南亞', sortOrder: 0, tone: 'purple' },
    { id: 'cnhk', label: '港澳大陸', sortOrder: 1, tone: 'green' },
    { id: 'am', label: '美洲', sortOrder: 2, tone: 'blue' },
    { id: 'oc', label: '大洋洲', sortOrder: 3, tone: 'orange' },
    { id: 'eu', label: '歐洲', sortOrder: 4, tone: 'indigo' },
    { id: 'nea', label: '東北亞', sortOrder: 5, tone: 'teal' }
  ];

  /** 航線群 id → 區域分類 id */
  const GROUP_REGION_MAP = {
    vn: 'sea', ph: 'sea', th: 'sea', sg: 'sea', kul: 'sea', id: 'sea', kh: 'sea', sea: 'sea',
    cn: 'cnhk', hkmo: 'cnhk',
    us: 'am', ca: 'am',
    au: 'oc',
    eu: 'eu',
    jp: 'nea', kr: 'nea', nea: 'nea'
  };

  const GROUP_RULES = [
    { id: 'vn', label: '越南線', icon: '🇻🇳', sortOrder: 0, regionId: 'sea', airports: ['HAN', 'SGN', 'DAD', 'CXR', 'PQC', 'VII', 'HUI'] },
    { id: 'cn', label: '大陸線', icon: '🇨🇳', sortOrder: 1, regionId: 'cnhk', airports: ['PVG', 'PEK', 'CAN', 'HKG', 'XMN', 'SZX', 'CTU', 'CKG', 'WUH', 'NKG', 'HGH', 'SHE', 'TAO', 'TSN', 'DLC', 'HRB', 'KMG', 'FOC', 'WNZ', 'CSX', 'CGQ', 'TNA', 'LHW', 'URC', 'HAK', 'NNG'] },
    { id: 'hkmo', label: '港澳線', icon: '🇭🇰', sortOrder: 2, regionId: 'cnhk', airports: ['HKG', 'MFM'] },
    { id: 'us', label: '美國線', icon: '🇺🇸', sortOrder: 3, regionId: 'am', airports: ['LAX', 'SFO', 'JFK', 'SEA', 'ORD', 'DFW', 'IAH', 'ANC', 'ATL', 'BOS', 'EWR', 'HNL', 'IAD', 'LAS', 'MIA', 'MSP', 'PHX', 'PDX', 'SAN', 'SLC', 'YVR'] },
    { id: 'ca', label: '加拿大線', icon: '🇨🇦', sortOrder: 4, regionId: 'am', airports: ['YVR', 'YYZ', 'YUL', 'YEG', 'YWG', 'YOW'] },
    { id: 'au', label: '澳洲線', icon: '🇦🇺', sortOrder: 5, regionId: 'oc', airports: ['SYD', 'MEL', 'BNE', 'PER', 'ADL', 'CBR', 'OOL', 'CNS', 'DRW', 'HBA'] },
    { id: 'eu', label: '歐洲線', icon: '🇪🇺', sortOrder: 6, regionId: 'eu', airports: ['LHR', 'CDG', 'FRA', 'AMS', 'FCO', 'MUC', 'MXP', 'VIE', 'BCN', 'MAD', 'ZRH', 'BRU', 'CPH', 'DUB', 'GVA', 'IST', 'LIS', 'MAN', 'OSL', 'PRG', 'STO', 'VCE', 'WAW', 'ATH', 'BER', 'BUD', 'HEL'] },
    { id: 'jp', label: '日本線', icon: '🇯🇵', sortOrder: 7, regionId: 'nea', airports: ['NRT', 'HND', 'KIX', 'FUK', 'OKA', 'CTS', 'SDJ', 'AOJ', 'KMQ', 'NGO', 'HIJ', 'KOJ', 'KMJ', 'MYJ', 'OIT', 'TAK', 'TKS', 'UBJ', 'IZO', 'MMJ', 'KIJ', 'HKD', 'AXT', 'FKS'] },
    { id: 'kr', label: '韓國線', icon: '🇰🇷', sortOrder: 8, regionId: 'nea', airports: ['ICN', 'PUS', 'GMP', 'CJU', 'TAE', 'KWJ', 'CJJ', 'MWX', 'RSU', 'USN', 'WJU'] },
    { id: 'ph', label: '菲律賓線', icon: '🇵🇭', sortOrder: 9, regionId: 'sea', airports: ['MNL', 'CEB', 'CRK', 'DVO', 'ILO', 'BCD', 'CGY', 'KLO', 'LGP', 'PPS', 'TAC'] },
    { id: 'th', label: '泰國線', icon: '🇹🇭', sortOrder: 10, regionId: 'sea', airports: ['BKK', 'DMK', 'CNX', 'HKT', 'UTP', 'KBV', 'HDY', 'USM', 'CEI'] },
    { id: 'sg', label: '新加坡線', icon: '🇸🇬', sortOrder: 11, regionId: 'sea', airports: ['SIN'] },
    { id: 'kul', label: '吉隆坡線', icon: '🇲🇾', sortOrder: 12, regionId: 'sea', airports: ['KUL', 'SZB'] },
    { id: 'id', label: '印尼線', icon: '🇮🇩', sortOrder: 13, regionId: 'sea', airports: ['CGK', 'DPS', 'SUB', 'UPG', 'MDC', 'LOP', 'YIA', 'BPN', 'PLM'] },
    { id: 'kh', label: '柬埔寨線', icon: '🇰🇭', sortOrder: 14, regionId: 'sea', airports: ['PNH', 'REP', 'SAI', 'KTI'] },
    { id: 'sea', label: '東南亞線', icon: '🌏', sortOrder: 15, regionId: 'sea', airports: ['SIN', 'KUL', 'BKK', 'MNL', 'CGK', 'DPS', 'HAN', 'SGN', 'DAD', 'CNX', 'CRK', 'CEB', 'VTE', 'PNH', 'RGN', 'BWN', 'KTI'] },
    { id: 'nea', label: '東北亞線', icon: '🗾', sortOrder: 16, regionId: 'nea', airports: ['TPE', 'ICN', 'NRT', 'KIX', 'PVG', 'PEK', 'HND', 'PUS', 'CTS', 'FUK', 'OKA', 'SZX', 'XMN', 'HKG', 'MFM', 'SDJ', 'AOJ', 'KMQ'] }
  ];

  const HOMELINE_SUMMER_2026_AIRPORTS = [
    'DAD', 'HAN', 'SGN', 'HKG', 'PEK', 'PVG', 'SZX', 'XMN', 'MFM',
    'ANC', 'DFW', 'IAH', 'JFK', 'LAX', 'ORD', 'SEA', 'SFO', 'YVR', 'YYZ',
    'BNE', 'CDG', 'MUC', 'MXP', 'VIE',
    'AOJ', 'CTS', 'FUK', 'KIX', 'KMQ', 'NRT', 'OKA', 'SDJ',
    'ICN', 'PUS', 'CEB', 'CRK', 'MNL', 'BKK', 'CNX',
    'SIN', 'KUL', 'CGK', 'DPS', 'KTI'
  ];

  function normalizeIata(code) {
    const s = String(code ?? '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
    return s.length === 3 ? s : '';
  }

  function buildGroupsFromAirports(airportList) {
    const set = new Set((airportList || []).map(normalizeIata).filter(Boolean));
    const groups = [];
    GROUP_RULES.forEach(rule => {
      const ruleSet = new Set((rule.airports || []).map(normalizeIata).filter(Boolean));
      const matched = [...set].filter(iata => ruleSet.has(iata)).sort();
      if (!matched.length) return;
      groups.push({
        id: rule.id,
        label: rule.label,
        icon: rule.icon,
        regionId: rule.regionId || GROUP_REGION_MAP[rule.id] || '',
        sortOrder: rule.sortOrder,
        airports: matched
      });
    });
    return groups;
  }

  function formatAirportPreview(airports) {
    const list = (airports || []).slice().sort();
    if (!list.length) return '—';
    return list.join(', ');
  }

  const HOMELINE_SUMMER_2026_GROUPS = buildGroupsFromAirports(HOMELINE_SUMMER_2026_AIRPORTS);

  function listRegions() {
    return REGION_CATALOG.slice().sort((a, b) => (a.sortOrder || 0) - (b.sortOrder || 0));
  }

  function getRegion(regionId) {
    const key = String(regionId ?? '').trim();
    if (!key) return null;
    return REGION_CATALOG.find(r => r.id === key) || null;
  }

  function resolveRegionId(groupId) {
    const key = String(groupId ?? '').trim();
    if (!key) return '';
    return GROUP_REGION_MAP[key] || '';
  }

  global.DestClassSeed = {
    SEED_VERSION,
    REGION_CATALOG,
    GROUP_REGION_MAP,
    GROUP_RULES,
    HOMELINE_SUMMER_2026_AIRPORTS,
    HOMELINE_SUMMER_2026_GROUPS,
    buildGroupsFromAirports,
    formatAirportPreview,
    normalizeIata,
    listRegions,
    getRegion,
    resolveRegionId
  };
})(typeof window !== 'undefined' ? window : globalThis);
