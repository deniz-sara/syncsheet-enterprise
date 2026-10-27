import { UniverWorkbookData } from './excelZeroUpload';
import { colIndexToLetter } from './collaboration';

export type PivotAggFunc = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';
export type PivotSortOrder = 'desc' | 'asc' | 'alpha';

export interface PivotConfig {
  sourceSheetId: string;
  rowColIdx: number;
  colColIdx: number | null; // null ise tek boyutlu, sayı ise çapraz (cross-tab) sütun kırılımı
  valColIdx: number;
  aggFunc: PivotAggFunc;
  filterColIdx: number | null;
  filterValue: string;
  sortOrder?: PivotSortOrder;
}

export interface PivotHeaderItem {
  colIndex: number;
  colLetter: string;
  title: string;
  rawName: string;
  isNumeric: boolean;
}

export interface PivotResultRow {
  rowKey: string;
  colValues: Record<string, number>;
  rowTotal: number;
  rowCount: number;
}

export interface PivotComputationResult {
  sourceSheetId: string;
  sourceSheetName: string;
  headers: PivotHeaderItem[];
  uniqueColKeys: string[];
  rows: PivotResultRow[];
  colTotals: Record<string, number>;
  grandTotal: number;
  totalRecordCount: number;
  filterOptions: string[];
}

export function isPivotSheetEntry(sheetId?: string, sheetObj?: any): boolean {
  if (!sheetId && !sheetObj) return false;
  const sid = String(sheetId || sheetObj?.id || '');
  const sname = String(sheetObj?.name || '');
  return (
    sid.startsWith('sheet_pivot_') ||
    sname.startsWith('Pivot_Tablo') ||
    Boolean(sheetObj?.pivotConfig) ||
    Boolean(sheetObj?.cellData?.[0]?.[0]?.custom?.pivotConfig)
  );
}

export function getSavedPivotConfigFromSheet(
  workbookData: UniverWorkbookData,
  pivotSheetId: string
): PivotConfig | null {
  const sheet = workbookData?.sheets?.[pivotSheetId];
  if (!sheet) return null;
  const cfg = sheet.pivotConfig || sheet.cellData?.[0]?.[0]?.custom?.pivotConfig;
  if (cfg && typeof cfg === 'object' && typeof cfg.rowColIdx === 'number') {
    return { ...cfg };
  }
  return null;
}

export function resolveSourceSheetId(
  workbookData: UniverWorkbookData,
  preferredSheetId?: string
): string {
  const sheetIds =
    workbookData.sheetOrder && workbookData.sheetOrder.length > 0
      ? workbookData.sheetOrder
      : Object.keys(workbookData.sheets || {});

  if (
    preferredSheetId &&
    workbookData.sheets?.[preferredSheetId] &&
    !isPivotSheetEntry(preferredSheetId, workbookData.sheets[preferredSheetId])
  ) {
    return preferredSheetId;
  }

  const firstNonPivot = sheetIds.find(
    (id) => workbookData.sheets?.[id] && !isPivotSheetEntry(id, workbookData.sheets[id])
  );
  return firstNonPivot || sheetIds[0] || 'sheet_1';
}

export function extractSheetHeaders(
  workbookData: UniverWorkbookData,
  sheetId?: string
): PivotHeaderItem[] {
  const sid = resolveSourceSheetId(workbookData, sheetId);
  const sheet = workbookData.sheets?.[sid];
  if (!sheet || !sheet.cellData) return [];

  const cellData = sheet.cellData;
  const headerRow = cellData[0] || {};

  // İlk 40 satırı tarayarak dolu sütunları ve sayısal olup olmadıklarını tespit et
  const colIndicesSet = new Set<number>();
  Object.keys(headerRow).forEach((k) => {
    const n = Number(k);
    if (!isNaN(n) && headerRow[n]?.v !== undefined && String(headerRow[n]?.v).trim() !== '') {
      colIndicesSet.add(n);
    }
  });

  const sampleRowKeys = Object.keys(cellData)
    .map(Number)
    .filter((r) => !isNaN(r) && r > 0)
    .sort((a, b) => a - b)
    .slice(0, 40);

  const numericStats: Record<number, { num: number; text: number }> = {};

  for (const r of sampleRowKeys) {
    const rowObj = cellData[r];
    if (!rowObj) continue;
    for (const cKey of Object.keys(rowObj)) {
      const c = Number(cKey);
      if (isNaN(c)) continue;
      const val = rowObj[c]?.v;
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        colIndicesSet.add(c);
        if (!numericStats[c]) numericStats[c] = { num: 0, text: 0 };
        const parsed =
          typeof val === 'number'
            ? val
            : Number(String(val).replace(/[^0-9.-]/g, ''));
        if (typeof val === 'number' || (!isNaN(parsed) && /[0-9]/.test(String(val)))) {
          numericStats[c].num++;
        } else {
          numericStats[c].text++;
        }
      }
    }
  }

  // Eğer sayfa tamamen boşsa bile A, B, C, D sütunlarını göster ki kullanıcı seçim yapabilsin
  if (colIndicesSet.size === 0) {
    return [0, 1, 2, 3].map((c) => {
      const letter = colIndexToLetter(c);
      return {
        colIndex: c,
        colLetter: letter,
        rawName: `${letter} Sütunu`,
        title: `${letter} — ${letter} Sütunu`,
        isNumeric: c >= 1,
      };
    });
  }

  const sortedCols = Array.from(colIndicesSet).sort((a, b) => a - b);
  const result: PivotHeaderItem[] = [];

  for (const c of sortedCols) {
    const raw = headerRow[c]?.v;
    const letter = colIndexToLetter(c);
    const rawName =
      raw !== undefined && raw !== null && String(raw).trim() !== ''
        ? String(raw).trim()
        : `${letter} Sütunu`;
    const stat = numericStats[c];
    const isNumeric = stat ? stat.num > 0 && stat.num >= stat.text : false;

    result.push({
      colIndex: c,
      colLetter: letter,
      rawName,
      title: `${letter} — ${rawName}`,
      isNumeric,
    });
  }

  return result;
}

export function computePivotTable(
  workbookData: UniverWorkbookData,
  config: PivotConfig
): PivotComputationResult {
  const sid = resolveSourceSheetId(workbookData, config.sourceSheetId);
  const sourceSheetName = workbookData.sheets?.[sid]?.name || 'Sayfa 1';
  const headers = extractSheetHeaders(workbookData, sid);

  const validRowColIdx = headers.some((h) => h.colIndex === config.rowColIdx)
    ? config.rowColIdx
    : headers[0]?.colIndex ?? 0;

  const validValColIdx = headers.some((h) => h.colIndex === config.valColIdx)
    ? config.valColIdx
    : headers.find((h) => h.isNumeric && h.colIndex !== validRowColIdx)?.colIndex ??
      headers[Math.min(1, headers.length - 1)]?.colIndex ??
      0;

  const cellData = workbookData.sheets?.[sid]?.cellData || {};
  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((r) => !isNaN(r) && r > 0)
    .sort((a, b) => a - b)
    .slice(0, 5000); // İlk 5000 satır üzerinden ultra-hızlı canlı özet

  const filterSet = new Set<string>();
  const colKeySet = new Set<string>();

  // groupMap[rowKey][colKey] = number[]
  const groupMap: Record<string, Record<string, number[]>> = {};
  let totalRecordCount = 0;

  for (const r of rowKeys) {
    const rowObj = cellData[r];
    if (!rowObj) continue;

    // Eğer satır "GENEL TOPLAM" satırıysa pivot hesabına iki kez girmemesi için atla
    const firstVal = String(rowObj[0]?.v || '').toUpperCase();
    if (firstVal.includes('GENEL TOPLAM')) continue;

    // Boş satır kontrolü
    const hasAnyValue = Object.values(rowObj).some(
      (cell: any) => cell && cell.v !== undefined && cell.v !== ''
    );
    if (!hasAnyValue) continue;

    if (config.filterColIdx !== null && config.filterColIdx >= 0) {
      const fVal = String(rowObj[config.filterColIdx]?.v ?? '(Boş)').trim() || '(Boş)';
      filterSet.add(fVal);
      if (config.filterValue && config.filterValue !== 'ALL' && fVal !== config.filterValue) {
        continue;
      }
    }

    const rKey = String(rowObj[validRowColIdx]?.v ?? '(Belirtilmemiş)').trim() || '(Boş)';
    const cKey =
      config.colColIdx !== null && config.colColIdx >= 0
        ? String(rowObj[config.colColIdx]?.v ?? 'Toplam').trim() || '(Boş)'
        : 'Değer';

    colKeySet.add(cKey);

    const rawVal = rowObj[validValColIdx]?.v;
    const numVal =
      typeof rawVal === 'number'
        ? rawVal
        : Number(String(rawVal || '').replace(/[^0-9.-]/g, ''));
    const validNum = !isNaN(numVal) ? numVal : 0;

    if (!groupMap[rKey]) groupMap[rKey] = {};
    if (!groupMap[rKey][cKey]) groupMap[rKey][cKey] = [];
    groupMap[rKey][cKey].push(config.aggFunc === 'COUNT' ? 1 : validNum);
    totalRecordCount++;
  }

  const uniqueColKeys =
    config.colColIdx !== null && config.colColIdx >= 0
      ? Array.from(colKeySet).slice(0, 12)
      : ['Değer'];

  const aggregateArray = (arr: number[], fn: PivotAggFunc): number => {
    if (!arr || arr.length === 0) return 0;
    if (fn === 'COUNT') return arr.length;
    const sum = arr.reduce((a, b) => a + b, 0);
    if (fn === 'SUM') return Math.round(sum * 100) / 100;
    if (fn === 'AVERAGE') return Math.round((sum / arr.length) * 100) / 100;
    if (fn === 'MAX') return Math.max(...arr);
    if (fn === 'MIN') return Math.min(...arr);
    return sum;
  };

  const rows: PivotResultRow[] = [];
  const colRawArrays: Record<string, number[]> = {};
  const allRawValues: number[] = [];

  Object.entries(groupMap).forEach(([rKey, colMap]) => {
    const colValues: Record<string, number> = {};
    const rowRaw: number[] = [];

    uniqueColKeys.forEach((cKey) => {
      const arr = colMap[cKey] || [];
      colValues[cKey] = aggregateArray(arr, config.aggFunc);
      rowRaw.push(...arr);
      if (!colRawArrays[cKey]) colRawArrays[cKey] = [];
      colRawArrays[cKey].push(...arr);
      allRawValues.push(...arr);
    });

    rows.push({
      rowKey: rKey,
      colValues,
      rowTotal: aggregateArray(rowRaw, config.aggFunc),
      rowCount: rowRaw.length,
    });
  });

  const sortOrder = config.sortOrder || 'desc';
  if (sortOrder === 'asc') {
    rows.sort((a, b) => a.rowTotal - b.rowTotal);
  } else if (sortOrder === 'alpha') {
    rows.sort((a, b) => a.rowKey.localeCompare(b.rowKey, 'tr'));
  } else {
    rows.sort((a, b) => b.rowTotal - a.rowTotal);
  }

  const colTotals: Record<string, number> = {};
  uniqueColKeys.forEach((cKey) => {
    colTotals[cKey] = aggregateArray(colRawArrays[cKey] || [], config.aggFunc);
  });

  const grandTotal = aggregateArray(allRawValues, config.aggFunc);

  return {
    sourceSheetId: sid,
    sourceSheetName,
    headers,
    uniqueColKeys,
    rows,
    colTotals,
    grandTotal,
    totalRecordCount,
    filterOptions: Array.from(filterSet),
  };
}

/**
 * Oluşturulan veya güncellenen Pivot Tabloyu Excel Çalışma Kitabının içine gerçek bir çalışma sayfası (Sheet)
 * olarak ekler veya mevcut Pivot sayfasını (existingPivotSheetId) ANLIK OLARAK yerinde günceller!
 */
export function insertPivotSheetIntoWorkbook(
  workbookData: UniverWorkbookData,
  pivotResult: PivotComputationResult,
  config: PivotConfig,
  existingPivotSheetId?: string | null
): { updatedWorkbook: UniverWorkbookData; newSheetName: string; pivotSheetId: string } {
  const clone: UniverWorkbookData = JSON.parse(JSON.stringify(workbookData));
  const existingPivotSheets = Object.keys(clone.sheets || {}).filter((k) =>
    isPivotSheetEntry(k, clone.sheets[k])
  );

  const reuseId =
    existingPivotSheetId && clone.sheets?.[existingPivotSheetId]
      ? existingPivotSheetId
      : null;

  const existingNames = new Set(
    Object.values(clone.sheets || {}).map((s: any) => String(s?.name || '').trim())
  );
  let nextNum = existingPivotSheets.length + 1;
  let candidateName = `Pivot_Tablo_${nextNum}`;
  while (existingNames.has(candidateName)) {
    nextNum++;
    candidateName = `Pivot_Tablo_${nextNum}`;
  }

  const newSheetId =
    reuseId || `sheet_pivot_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  const newSheetName = reuseId
    ? clone.sheets[reuseId].name || 'Pivot_Tablo_1'
    : candidateName;

  const rowHeaderTitle =
    pivotResult.headers.find((h) => h.colIndex === config.rowColIdx)?.title ||
    pivotResult.headers[0]?.title ||
    'Satır Etiketleri';
  const colHeaderTitle =
    config.colColIdx !== null
      ? pivotResult.headers.find((h) => h.colIndex === config.colColIdx)?.title || 'Sütunlar'
      : 'Tek Boyut';
  const valHeaderTitle =
    pivotResult.headers.find((h) => h.colIndex === config.valColIdx)?.title ||
    pivotResult.headers[Math.min(1, pivotResult.headers.length - 1)]?.title ||
    'Değerler';
  const filterHeaderTitle =
    config.filterColIdx !== null
      ? pivotResult.headers.find((h) => h.colIndex === config.filterColIdx)?.title || 'Filtre'
      : 'Yok';

  const aggLabelMap: Record<PivotAggFunc, string> = {
    SUM: 'Toplam',
    AVERAGE: 'Ortalama',
    COUNT: 'Sayım (Adet)',
    MAX: 'Maksimum',
    MIN: 'Minimum',
  };

  clone.styles = {
    ...(clone.styles || {}),
    pivot_header: {
      bl: 1,
      bg: { rgb: '#1e3a8a' },
      cl: { rgb: '#ffffff' },
      fs: 11,
    },
    pivot_submeta: {
      bl: 1,
      bg: { rgb: '#eff6ff' },
      cl: { rgb: '#1d4ed8' },
      fs: 10,
    },
    pivot_total: {
      bl: 1,
      bg: { rgb: '#dbeafe' },
      cl: { rgb: '#1e3a8a' },
      fs: 11,
    },
  };

  const cellData: Record<number, Record<number, any>> = {};

  // Üst bilgi satırı (Satır 0) - custom.pivotConfig içinde bu Pivot sekmesinin kendi bağımsız ayarlarını saklar!
  cellData[0] = {
    0: {
      v: `${newSheetName} (${pivotResult.sourceSheetName}): ${rowHeaderTitle} bazında ${aggLabelMap[config.aggFunc]} (${valHeaderTitle})`,
      t: 1,
      s: 'pivot_total',
      custom: {
        pivotConfig: { ...config, sourceSheetId: pivotResult.sourceSheetId },
      },
    },
  };

  // Satır 1: Aktif Pivot Alanları özeti (Kullanıcı sağdaki menüden değiştirdikçe anında güncellenir)
  cellData[1] = {
    0: {
      v: `Kaynak: ${pivotResult.sourceSheetName}`,
      t: 1,
      s: 'pivot_submeta',
    },
    1: {
      v: `Satır: ${rowHeaderTitle}`,
      t: 1,
      s: 'pivot_submeta',
    },
    2: {
      v: `Sütun: ${colHeaderTitle}`,
      t: 1,
      s: 'pivot_submeta',
    },
    3: {
      v: `Değer: ${aggLabelMap[config.aggFunc]} (${valHeaderTitle})`,
      t: 1,
      s: 'pivot_submeta',
    },
    4: {
      v:
        config.filterColIdx !== null
          ? `Filtre: ${filterHeaderTitle} = ${config.filterValue === 'ALL' ? 'Tümü' : config.filterValue}`
          : 'Filtre: Tümü (Filtre Yok)',
      t: 1,
      s: 'pivot_submeta',
    },
  };

  // Tablo Başlıkları (Satır 2)
  cellData[2] = {
    0: { v: rowHeaderTitle, t: 1, s: 'pivot_header' },
  };

  pivotResult.uniqueColKeys.forEach((cKey, idx) => {
    cellData[2][idx + 1] = {
      v: cKey === 'Değer' ? `${aggLabelMap[config.aggFunc]} - ${valHeaderTitle}` : cKey,
      t: 1,
      s: 'pivot_header',
    };
  });

  const showRowTotalCol = pivotResult.uniqueColKeys.length > 1;
  if (showRowTotalCol) {
    cellData[2][pivotResult.uniqueColKeys.length + 1] = {
      v: 'GENEL TOPLAM',
      t: 1,
      s: 'pivot_header',
    };
  }

  cellData[2][showRowTotalCol ? pivotResult.uniqueColKeys.length + 2 : 2] = {
    v: 'KAYIT ADEDİ',
    t: 1,
    s: 'pivot_header',
  };

  // Veri Satırları (Satır 3'ten itibaren)
  let currentRow = 3;
  pivotResult.rows.forEach((rItem) => {
    cellData[currentRow] = {
      0: { v: rItem.rowKey, t: 1 },
    };
    pivotResult.uniqueColKeys.forEach((cKey, idx) => {
      cellData[currentRow][idx + 1] = {
        v: rItem.colValues[cKey] ?? 0,
        t: 2,
      };
    });
    if (showRowTotalCol) {
      cellData[currentRow][pivotResult.uniqueColKeys.length + 1] = {
        v: rItem.rowTotal,
        t: 2,
        s: 'pivot_total',
      };
    }
    cellData[currentRow][showRowTotalCol ? pivotResult.uniqueColKeys.length + 2 : 2] = {
      v: rItem.rowCount,
      t: 2,
    };
    currentRow++;
  });

  // Alt Genel Toplam Satırı
  cellData[currentRow] = {
    0: { v: 'GENEL TOPLAM', t: 1, s: 'pivot_total' },
  };
  pivotResult.uniqueColKeys.forEach((cKey, idx) => {
    cellData[currentRow][idx + 1] = {
      v: pivotResult.colTotals[cKey] ?? 0,
      t: 2,
      s: 'pivot_total',
    };
  });
  if (showRowTotalCol) {
    cellData[currentRow][pivotResult.uniqueColKeys.length + 1] = {
      v: pivotResult.grandTotal,
      t: 2,
      s: 'pivot_total',
    };
  }
  cellData[currentRow][showRowTotalCol ? pivotResult.uniqueColKeys.length + 2 : 2] = {
    v: pivotResult.totalRecordCount,
    t: 2,
    s: 'pivot_total',
  };

  clone.sheets[newSheetId] = {
    id: newSheetId,
    name: newSheetName,
    tabColor: '#2563eb',
    rowCount: Math.max(currentRow + 100, 500),
    columnCount: 20,
    defaultColumnWidth: 175,
    defaultRowHeight: 26,
    cellData,
    pivotConfig: { ...config, sourceSheetId: pivotResult.sourceSheetId },
  };

  // Oluşturulan veya güncellenen Pivot sekmesini ilk sıraya (aktif görünür sekmeye) koyalım
  const existingOrder = (clone.sheetOrder || Object.keys(clone.sheets)).filter(
    (id) => id !== newSheetId
  );
  clone.sheetOrder = [newSheetId, ...existingOrder];

  return { updatedWorkbook: clone, newSheetName, pivotSheetId: newSheetId };
}

/**
 * KOŞULLU BİÇİMLENDİRME (CONDITIONAL FORMATTING) MOTORU
 */
export type ConditionalRuleOperator =
  | 'gt'
  | 'lt'
  | 'between'
  | 'eq'
  | 'contains'
  | 'above_avg'
  | 'below_avg'
  | 'duplicates';

export type ConditionalColorPreset = 'green' | 'red' | 'yellow' | 'blue' | 'purple';

export interface ConditionalRuleItem {
  id: string;
  sheetId: string;
  colIndex: number; // -1 = tüm sütunlar
  colLabel: string;
  operator: ConditionalRuleOperator;
  operatorLabel: string;
  value1: string;
  value2?: string;
  colorPreset: ConditionalColorPreset;
  matchedCount: number;
}

export const COLOR_PRESET_MAP: Record<
  ConditionalColorPreset,
  { label: string; bg: string; cl: string; badgeClass: string }
> = {
  green: {
    label: 'Yeşil Dolgu ve Koyu Yeşil Metin',
    bg: '#dcfce7',
    cl: '#166534',
    badgeClass: 'bg-emerald-100 text-emerald-800 border-emerald-300',
  },
  red: {
    label: 'Açık Kırmızı Dolgu ve Koyu Kırmızı Metin',
    bg: '#fee2e2',
    cl: '#991b1b',
    badgeClass: 'bg-rose-100 text-rose-800 border-rose-300',
  },
  yellow: {
    label: 'Sarı Dolgu ve Koyu Sarı Metin',
    bg: '#fef9c3',
    cl: '#854d0e',
    badgeClass: 'bg-amber-100 text-amber-800 border-amber-300',
  },
  blue: {
    label: 'Mavi Dolgu ve Koyu Mavi Metin',
    bg: '#dbeafe',
    cl: '#1e40af',
    badgeClass: 'bg-blue-100 text-blue-800 border-blue-300',
  },
  purple: {
    label: 'Mor Dolgu ve Koyu Mor Metin',
    bg: '#f3e8ff',
    cl: '#6b21a8',
    badgeClass: 'bg-purple-100 text-purple-800 border-purple-300',
  },
};

export function applyConditionalFormattingRule(
  workbookData: UniverWorkbookData,
  sheetId: string,
  colIndex: number,
  operator: ConditionalRuleOperator,
  value1: string,
  value2: string,
  colorPreset: ConditionalColorPreset
): { updatedWorkbook: UniverWorkbookData; matchedCount: number } {
  const clone: UniverWorkbookData = JSON.parse(JSON.stringify(workbookData));
  const sid =
    sheetId && clone.sheets?.[sheetId]
      ? sheetId
      : clone.sheetOrder?.[0] || Object.keys(clone.sheets || {})[0];
  const sheet = clone.sheets?.[sid];
  if (!sheet || !sheet.cellData) {
    return { updatedWorkbook: clone, matchedCount: 0 };
  }

  const styleKey = `cf_${colorPreset}_style`;
  const preset = COLOR_PRESET_MAP[colorPreset];
  clone.styles = {
    ...(clone.styles || {}),
    [styleKey]: {
      bl: 1,
      bg: { rgb: preset.bg },
      cl: { rgb: preset.cl },
    },
  };

  const cellData = sheet.cellData;
  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((r) => !isNaN(r) && r > 0);

  // Ortalama veya yinelenen değerler için ön tarama
  const numericValues: number[] = [];
  const freqMap: Record<string, number> = {};

  rowKeys.forEach((r) => {
    const rowObj = cellData[r];
    if (!rowObj) return;
    const firstCell = String(rowObj[0]?.v || '').toUpperCase();
    if (firstCell.includes('GENEL TOPLAM')) return;

    const colsToCheck = colIndex >= 0 ? [colIndex] : Object.keys(rowObj).map(Number);
    colsToCheck.forEach((c) => {
      const v = rowObj[c]?.v;
      if (v === undefined || v === null || v === '') return;
      const num = Number(v);
      if (!isNaN(num)) numericValues.push(num);
      const strKey = String(v).trim().toLowerCase();
      freqMap[strKey] = (freqMap[strKey] || 0) + 1;
    });
  });

  const avg =
    numericValues.length > 0
      ? numericValues.reduce((a, b) => a + b, 0) / numericValues.length
      : 0;

  let matchedCount = 0;
  const numVal1 = Number(value1);
  const numVal2 = Number(value2);

  rowKeys.forEach((r) => {
    const rowObj = cellData[r];
    if (!rowObj) return;
    const firstCell = String(rowObj[0]?.v || '').toUpperCase();
    if (firstCell.includes('GENEL TOPLAM')) return;

    const colsToCheck = colIndex >= 0 ? [colIndex] : Object.keys(rowObj).map(Number);
    colsToCheck.forEach((c) => {
      const cell = rowObj[c];
      if (!cell || cell.v === undefined || cell.v === null || cell.v === '') return;

      const rawVal = cell.v;
      const num = Number(rawVal);
      const strVal = String(rawVal).trim().toLowerCase();
      let isMatch = false;

      if (operator === 'gt' && !isNaN(num) && !isNaN(numVal1)) {
        isMatch = num > numVal1;
      } else if (operator === 'lt' && !isNaN(num) && !isNaN(numVal1)) {
        isMatch = num < numVal1;
      } else if (operator === 'between' && !isNaN(num) && !isNaN(numVal1) && !isNaN(numVal2)) {
        const minV = Math.min(numVal1, numVal2);
        const maxV = Math.max(numVal1, numVal2);
        isMatch = num >= minV && num <= maxV;
      } else if (operator === 'eq') {
        isMatch = strVal === value1.trim().toLowerCase() || (!isNaN(num) && num === numVal1);
      } else if (operator === 'contains') {
        isMatch = strVal.includes(value1.trim().toLowerCase());
      } else if (operator === 'above_avg' && !isNaN(num)) {
        isMatch = num > avg;
      } else if (operator === 'below_avg' && !isNaN(num)) {
        isMatch = num < avg;
      } else if (operator === 'duplicates') {
        isMatch = (freqMap[strVal] || 0) > 1;
      }

      if (isMatch) {
        cell.s = styleKey;
        matchedCount++;
      }
    });
  });

  return { updatedWorkbook: clone, matchedCount };
}

/**
 * Çoklu Çalışma Kitabı (Multi-Workbook) geçişlerini hemen test edebilmek için
 * 2. hazır kurumsal çalışma kitabı: Satış & CRM Hedef Tablosu
 */
export function getSalesCrmWorkbook(): UniverWorkbookData {
  return {
    id: 'wb_sales_crm_2026',
    name: '2026_Satis_ve_CRM_Hedefleri',
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {
      header_style: {
        bl: 1,
        bg: { rgb: '#1e3a8a' },
        cl: { rgb: '#ffffff' },
        fs: 11,
      },
      formula_cell: {
        bl: 1,
        bg: { rgb: '#eff6ff' },
        cl: { rgb: '#1d4ed8' },
      },
      status_ok: {
        bl: 1,
        bg: { rgb: '#dcfce7' },
        cl: { rgb: '#15803d' },
      },
    },
    sheetOrder: ['sheet_sales'],
    sheets: {
      sheet_sales: {
        id: 'sheet_sales',
        name: 'Bolgesel_Satislar',
        tabColor: '#2563eb',
        rowCount: 1500,
        columnCount: 20,
        defaultColumnWidth: 150,
        defaultRowHeight: 26,
        freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 },
        cellData: {
          0: {
            0: { v: 'MÜŞTERİ KODU', t: 1, s: 'header_style' },
            1: { v: 'KURUMSAL MÜŞTERİ', t: 1, s: 'header_style' },
            2: { v: 'BÖLGE', t: 1, s: 'header_style' },
            3: { v: 'TEMSİLCİ', t: 1, s: 'header_style' },
            4: { v: 'HEDEF CİRO (₺)', t: 1, s: 'header_style' },
            5: { v: 'GERÇEKLEŞEN (₺)', t: 1, s: 'header_style' },
            6: { v: 'FARK (FORMÜL)', t: 1, s: 'header_style' },
            7: { v: 'DURUM', t: 1, s: 'header_style' },
          },
          1: {
            0: { v: 'CRM-101', t: 1 },
            1: { v: 'Anadolu Holding A.Ş.', t: 1 },
            2: { v: 'Marmara', t: 1 },
            3: { v: 'Ahmet Kaya', t: 1 },
            4: { v: 850000, t: 2 },
            5: { v: 920000, t: 2 },
            6: { f: '=F2-E2', v: 70000, t: 2, s: 'formula_cell' },
            7: { v: 'Hedef Aşıldı', t: 1, s: 'status_ok' },
          },
          2: {
            0: { v: 'CRM-102', t: 1 },
            1: { v: 'Ege Lojistik Grubu', t: 1 },
            2: { v: 'Ege', t: 1 },
            3: { v: 'Selin Yılmaz', t: 1 },
            4: { v: 640000, t: 2 },
            5: { v: 610000, t: 2 },
            6: { f: '=F3-E3', v: -30000, t: 2, s: 'formula_cell' },
            7: { v: 'Takip Ediliyor', t: 1 },
          },
          3: {
            0: { v: 'CRM-103', t: 1 },
            1: { v: 'Başkent Enerji Sanayi', t: 1 },
            2: { v: 'İç Anadolu', t: 1 },
            3: { v: 'Deniz Sara', t: 1 },
            4: { v: 1200000, t: 2 },
            5: { v: 1350000, t: 2 },
            6: { f: '=F4-E4', v: 150000, t: 2, s: 'formula_cell' },
            7: { v: 'Hedef Aşıldı', t: 1, s: 'status_ok' },
          },
          4: {
            0: { v: 'CRM-104', t: 1 },
            1: { v: 'Akdeniz Turizm Yatırım', t: 1 },
            2: { v: 'Akdeniz', t: 1 },
            3: { v: 'Mert Demir', t: 1 },
            4: { v: 500000, t: 2 },
            5: { v: 540000, t: 2 },
            6: { f: '=F5-E5', v: 40000, t: 2, s: 'formula_cell' },
            7: { v: 'Hedef Aşıldı', t: 1, s: 'status_ok' },
          },
        },
      },
    },
  };
}

export function getBlankWorkbook(indexNum: number): UniverWorkbookData {
  const sheetId = 'sheet_1';
  return {
    id: `wb_blank_${Date.now()}`,
    name: `Yeni_Calisma_Kitabi_${indexNum}`,
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: 'Sayfa1',
        tabColor: '#10b981',
        rowCount: 1500,
        columnCount: 26,
        defaultColumnWidth: 130,
        defaultRowHeight: 25,
        cellData: {},
      },
    },
  };
}
