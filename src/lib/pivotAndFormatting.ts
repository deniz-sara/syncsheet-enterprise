import { UniverWorkbookData } from './excelZeroUpload';
import { colIndexToLetter } from './collaboration';

export type PivotAggFunc = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';

export interface PivotConfig {
  sourceSheetId: string;
  rowColIdx: number;
  colColIdx: number | null; // null ise tek boyutlu, sayı ise çapraz (cross-tab) sütun kırılımı
  valColIdx: number;
  aggFunc: PivotAggFunc;
  filterColIdx: number | null;
  filterValue: string;
}

export interface PivotResultRow {
  rowKey: string;
  colValues: Record<string, number>;
  rowTotal: number;
  rowCount: number;
}

export interface PivotComputationResult {
  headers: { colIndex: number; colLetter: string; title: string }[];
  uniqueColKeys: string[];
  rows: PivotResultRow[];
  colTotals: Record<string, number>;
  grandTotal: number;
  totalRecordCount: number;
  filterOptions: string[];
}

export function extractSheetHeaders(
  workbookData: UniverWorkbookData,
  sheetId?: string
): { colIndex: number; colLetter: string; title: string }[] {
  const sid =
    sheetId && workbookData.sheets?.[sheetId]
      ? sheetId
      : workbookData.sheetOrder?.[0] || Object.keys(workbookData.sheets || {})[0];
  const sheet = workbookData.sheets?.[sid];
  if (!sheet || !sheet.cellData) return [];

  const headerRow = sheet.cellData[0] || {};
  const maxCols = Math.max(
    ...Object.keys(headerRow)
      .map(Number)
      .filter((n) => !isNaN(n)),
    7
  );

  const result: { colIndex: number; colLetter: string; title: string }[] = [];
  for (let c = 0; c <= maxCols; c++) {
    const raw = headerRow[c]?.v;
    const letter = colIndexToLetter(c);
    if (raw !== undefined && raw !== null && String(raw).trim() !== '') {
      result.push({
        colIndex: c,
        colLetter: letter,
        title: `${letter} — ${String(raw).trim()}`,
      });
    }
  }
  return result;
}

export function computePivotTable(
  workbookData: UniverWorkbookData,
  config: PivotConfig
): PivotComputationResult {
  const headers = extractSheetHeaders(workbookData, config.sourceSheetId);
  const sid =
    config.sourceSheetId && workbookData.sheets?.[config.sourceSheetId]
      ? config.sourceSheetId
      : workbookData.sheetOrder?.[0] || Object.keys(workbookData.sheets || {})[0];

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
      const fVal = String(rowObj[config.filterColIdx]?.v ?? '(Boş)').trim();
      filterSet.add(fVal);
      if (config.filterValue && config.filterValue !== 'ALL' && fVal !== config.filterValue) {
        continue;
      }
    }

    const rKey = String(rowObj[config.rowColIdx]?.v ?? '(Belirtilmemiş)').trim() || '(Boş)';
    const cKey =
      config.colColIdx !== null && config.colColIdx >= 0
        ? String(rowObj[config.colColIdx]?.v ?? 'Toplam').trim() || '(Boş)'
        : 'Değer';

    colKeySet.add(cKey);

    const rawVal = rowObj[config.valColIdx]?.v;
    const numVal = typeof rawVal === 'number' ? rawVal : Number(String(rawVal || '').replace(/[^0-9.-]/g, ''));
    const validNum = !isNaN(numVal) ? numVal : 0;

    if (!groupMap[rKey]) groupMap[rKey] = {};
    if (!groupMap[rKey][cKey]) groupMap[rKey][cKey] = [];
    groupMap[rKey][cKey].push(config.aggFunc === 'COUNT' ? 1 : validNum);
    totalRecordCount++;
  }

  const uniqueColKeys =
    config.colColIdx !== null && config.colColIdx >= 0
      ? Array.from(colKeySet).slice(0, 10)
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

  rows.sort((a, b) => b.rowTotal - a.rowTotal);

  const colTotals: Record<string, number> = {};
  uniqueColKeys.forEach((cKey) => {
    colTotals[cKey] = aggregateArray(colRawArrays[cKey] || [], config.aggFunc);
  });

  const grandTotal = aggregateArray(allRawValues, config.aggFunc);

  return {
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
 * Oluşturulan Pivot Tabloyu Excel Çalışma Kitabının içine gerçek bir yeni çalışma sayfası (Sheet)
 * olarak ekler ve yeni workbookData döndürür.
 */
export function insertPivotSheetIntoWorkbook(
  workbookData: UniverWorkbookData,
  pivotResult: PivotComputationResult,
  config: PivotConfig
): { updatedWorkbook: UniverWorkbookData; newSheetName: string } {
  const clone: UniverWorkbookData = JSON.parse(JSON.stringify(workbookData));
  const existingPivotCount = Object.keys(clone.sheets || {}).filter((k) =>
    k.startsWith('sheet_pivot_')
  ).length;

  const newSheetId = `sheet_pivot_${Date.now()}`;
  const newSheetName = `Pivot_Tablo_${existingPivotCount + 1}`;

  const rowHeaderTitle =
    pivotResult.headers.find((h) => h.colIndex === config.rowColIdx)?.title || 'Satır Etiketleri';
  const valHeaderTitle =
    pivotResult.headers.find((h) => h.colIndex === config.valColIdx)?.title || 'Değerler';

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
    pivot_total: {
      bl: 1,
      bg: { rgb: '#dbeafe' },
      cl: { rgb: '#1e3a8a' },
      fs: 11,
    },
  };

  const cellData: Record<number, Record<number, any>> = {};

  // Üst bilgi satırı (Satır 0)
  cellData[0] = {
    0: {
      v: `PİVOT RAPORU: ${rowHeaderTitle} bazında ${aggLabelMap[config.aggFunc]} (${valHeaderTitle})`,
      t: 1,
      s: 'pivot_total',
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
    defaultColumnWidth: 165,
    defaultRowHeight: 26,
    cellData,
  };

  // Yeni oluşturulan Pivot sekmesini ilk sıraya veya aktif görünür sıraya koyalım
  clone.sheetOrder = [newSheetId, ...(clone.sheetOrder || [])];

  return { updatedWorkbook: clone, newSheetName };
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
    styles: {
      header_style: {
        bl: 1,
        bg: { rgb: '#f1f5f9' },
        cl: { rgb: '#0f172a' },
        fs: 11,
      },
    },
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
        cellData: {
          0: {
            0: { v: 'BAŞLIK 1', t: 1, s: 'header_style' },
            1: { v: 'BAŞLIK 2', t: 1, s: 'header_style' },
            2: { v: 'TUTAR', t: 1, s: 'header_style' },
          },
        },
      },
    },
  };
}
