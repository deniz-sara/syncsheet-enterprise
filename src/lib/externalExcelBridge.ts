import * as XLSX from 'xlsx';
import { UniverWorkbookData } from './excelZeroUpload';
import { colIndexToLetter } from './collaboration';
import { evaluateFormulaPreview } from './excelFormulas';

export interface LinkedDesktopExcelInfo {
  id: string;
  fileName: string;
  linkedSheetId: string;
  linkedSheetName: string;
  lastModified: number;
  lastSyncedAt: string;
  rowCount: number;
  colCount: number;
  headers: Array<{
    colIndex: number;
    colLetter: string;
    headerName: string;
    sampleValue: any;
  }>;
  isLiveWatching: boolean;
}

export function sanitizeLinkedSheetName(fileName: string): string {
  const base = fileName
    .replace(/\.(xlsx|xls|csv|xlsm)$/i, '')
    .replace(/[^a-zA-Z0-9ğüşıöçĞÜŞİÖÇ_]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
    .slice(0, 18);

  return `DisExcel_${base || 'Tablo'}`;
}

export function extractExternalSheetMetadata(
  sheet: any
): {
  rowCount: number;
  colCount: number;
  headers: Array<{
    colIndex: number;
    colLetter: string;
    headerName: string;
    sampleValue: any;
  }>;
} {
  const cellData = sheet?.cellData || {};
  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((n) => !isNaN(n))
    .sort((a, b) => a - b);

  const maxRow = rowKeys.length > 0 ? Math.max(...rowKeys) : 0;
  const headerRow = cellData[0] || {};
  const colIndices = Object.keys(headerRow)
    .map(Number)
    .filter((n) => !isNaN(n))
    .sort((a, b) => a - b);

  const headers = colIndices.map((cIdx) => {
    const colLetter = colIndexToLetter(cIdx);
    const headerName = String(headerRow[cIdx]?.v || `${colLetter} Sütunu`);
    const sampleValue = cellData[1]?.[cIdx]?.v ?? '';
    return {
      colIndex: cIdx,
      colLetter,
      headerName,
      sampleValue,
    };
  });

  return {
    rowCount: maxRow + 1,
    colCount: headers.length,
    headers,
  };
}

/**
 * Bilgisayardaki gerçek Excel dosyasından okunan sayfayı mevcut SyncSheet çalışma kitabına
 * bağlı referans sekmesi (Örn: DisExcel_Musteriler) olarak ekler veya günceller.
 * Ardından ana sayfadaki VLOOKUP / XLOOKUP formüllerini yeni veriye göre yeniden hesaplar.
 */
export function mergeExternalExcelIntoWorkbook(
  currentWorkbook: UniverWorkbookData,
  externalWorkbook: UniverWorkbookData,
  fileName: string,
  isLiveWatching: boolean,
  lastModified: number
): {
  updatedWorkbook: UniverWorkbookData;
  linkedInfo: LinkedDesktopExcelInfo;
  recalculatedCellCount: number;
} {
  const cloned: UniverWorkbookData = JSON.parse(JSON.stringify(currentWorkbook));
  const extSheetIds = externalWorkbook.sheetOrder?.length
    ? externalWorkbook.sheetOrder
    : Object.keys(externalWorkbook.sheets || {});
  const primaryExtSheetId = extSheetIds[0];
  const extSheet = externalWorkbook.sheets?.[primaryExtSheetId];

  const linkedSheetName = sanitizeLinkedSheetName(fileName);
  const linkedSheetId = `linked_${linkedSheetName.toLowerCase()}`;

  const cleanCellData = extSheet?.cellData ? JSON.parse(JSON.stringify(extSheet.cellData)) : {};

  cloned.sheets[linkedSheetId] = {
    id: linkedSheetId,
    name: linkedSheetName,
    rowCount: Math.max(extSheet?.rowCount || 1000, 1000),
    columnCount: Math.max(extSheet?.columnCount || 26, 26),
    cellData: cleanCellData,
  };

  if (!cloned.sheetOrder.includes(linkedSheetId)) {
    cloned.sheetOrder.push(linkedSheetId);
  }

  const meta = extractExternalSheetMetadata(cloned.sheets[linkedSheetId]);

  // Ana sayfalardaki mevcut VLOOKUP / XLOOKUP formüllerini yeni dış Excel verisiyle yeniden hesapla
  let recalculatedCellCount = 0;
  for (const [sid, sheet] of Object.entries(cloned.sheets)) {
    if (sid === linkedSheetId || !sheet?.cellData) continue;
    for (const rKey of Object.keys(sheet.cellData)) {
      const r = Number(rKey);
      const rowObj = sheet.cellData[r];
      if (!rowObj) continue;
      for (const cKey of Object.keys(rowObj)) {
        const c = Number(cKey);
        const cell = rowObj[c];
        if (!cell?.f || typeof cell.f !== 'string') continue;

        const formulaUpper = cell.f.toUpperCase();
        if (
          formulaUpper.includes(linkedSheetName.toUpperCase()) ||
          formulaUpper.includes('VLOOKUP(') ||
          formulaUpper.includes('XLOOKUP(')
        ) {
          const evalResult = evaluateRawFormulaOnWorkbook(cell.f, cloned, sid);
          if (evalResult !== undefined) {
            cell.v = evalResult;
            cell.t = typeof evalResult === 'number' ? 2 : 1;
            recalculatedCellCount++;
          }
        }
      }
    }
  }

  const linkedInfo: LinkedDesktopExcelInfo = {
    id: linkedSheetId,
    fileName,
    linkedSheetId,
    linkedSheetName,
    lastModified,
    lastSyncedAt: new Date().toLocaleTimeString('tr-TR'),
    rowCount: meta.rowCount,
    colCount: meta.colCount,
    headers: meta.headers,
    isLiveWatching,
  };

  return {
    updatedWorkbook: cloned,
    linkedInfo,
    recalculatedCellCount,
  };
}

/**
 * Hücre içindeki ham =VLOOKUP(...) veya =XLOOKUP(...) formül dizgesini ayrıştırıp
 * güncel çalışma kitabı üzerinde hesaplar.
 */
export function evaluateRawFormulaOnWorkbook(
  formulaStr: string,
  workbookData: UniverWorkbookData,
  activeSheetId: string
): any {
  const trimmed = formulaStr.trim();
  const match = trimmed.match(/^=\s*(VLOOKUP|XLOOKUP)\s*\((.*)\)$/i);
  if (!match) return undefined;

  const fnName = match[1].toUpperCase();
  const argsRaw = splitFormulaTopLevelArgs(match[2]);

  if (fnName === 'VLOOKUP' && argsRaw.length >= 3) {
    const { computedValue } = evaluateFormulaPreview(
      'VLOOKUP',
      {
        lookup_value: argsRaw[0] || '',
        table_array: argsRaw[1] || '',
        col_index_num: argsRaw[2] || '2',
        range_lookup: argsRaw[3] || '0',
      },
      workbookData,
      activeSheetId
    );
    return computedValue;
  }

  if (fnName === 'XLOOKUP' && argsRaw.length >= 3) {
    const { computedValue } = evaluateFormulaPreview(
      'XLOOKUP',
      {
        lookup_value: argsRaw[0] || '',
        lookup_array: argsRaw[1] || '',
        return_array: argsRaw[2] || '',
        if_not_found: argsRaw[3] || '"Bulunamadı"',
      },
      workbookData,
      activeSheetId
    );
    return computedValue;
  }

  return undefined;
}

function splitFormulaTopLevelArgs(inner: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  let depth = 0;

  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if (ch === '"') {
      inQuotes = !inQuotes;
      current += ch;
      continue;
    }
    if (!inQuotes) {
      if (ch === '(') depth++;
      else if (ch === ')') depth = Math.max(0, depth - 1);
      else if ((ch === ',' || ch === ';') && depth === 0) {
        result.push(current.trim());
        current = '';
        continue;
      }
    }
    current += ch;
  }
  if (current.trim() !== '') {
    result.push(current.trim());
  }
  return result;
}

/**
 * Bağlı masaüstü Excel tablosundan aktif SyncSheet tablosundaki tüm satırlara
 * toplu olarak canlı =VLOOKUP veya =XLOOKUP formülü + hesaplanmış değer yazar.
 */
export function applyBulkCrossWorkbookLookup(params: {
  workbookData: UniverWorkbookData;
  activeSheetId?: string;
  linkedInfo: LinkedDesktopExcelInfo;
  formulaType: 'VLOOKUP' | 'XLOOKUP';
  sourceKeyColIndex: number; // SyncSheet'teki aranan anahtar sütunu (Örn: 0 -> A sütunu)
  externalKeyColIndex: number; // Masaüstü Excel'deki anahtar sütunu (Örn: 0 -> A sütunu)
  externalReturnColIndex: number; // Masaüstü Excel'den getirilecek sütun (Örn: 2 -> C sütunu)
  targetColIndex: number | 'NEW_COL'; // SyncSheet'te yazılacak hedef sütun
  newColumnHeader?: string;
}): {
  updatedWorkbook: UniverWorkbookData;
  matchedCount: number;
  totalRowsProcessed: number;
  targetColLetter: string;
  sampleFormula: string;
} {
  const {
    workbookData,
    activeSheetId,
    linkedInfo,
    formulaType,
    sourceKeyColIndex,
    externalKeyColIndex,
    externalReturnColIndex,
    targetColIndex,
    newColumnHeader,
  } = params;

  const cloned: UniverWorkbookData = JSON.parse(JSON.stringify(workbookData));
  const sid =
    activeSheetId && cloned.sheets[activeSheetId]
      ? activeSheetId
      : cloned.sheetOrder.find((id) => id !== linkedInfo.linkedSheetId) || cloned.sheetOrder[0];

  const sheet = cloned.sheets[sid];
  if (!sheet || !sheet.cellData) {
    return {
      updatedWorkbook: cloned,
      matchedCount: 0,
      totalRowsProcessed: 0,
      targetColLetter: 'A',
      sampleFormula: '',
    };
  }

  const cellData = sheet.cellData;
  const headerRow = cellData[0] || {};
  const existingCols = Object.keys(headerRow)
    .map(Number)
    .filter((n) => !isNaN(n));
  const maxExistingCol = existingCols.length > 0 ? Math.max(...existingCols) : 0;

  const finalTargetCol =
    targetColIndex === 'NEW_COL' ? maxExistingCol + 1 : Number(targetColIndex);

  const returnHeaderObj = linkedInfo.headers.find((h) => h.colIndex === externalReturnColIndex);
  const headerTitle =
    newColumnHeader?.trim() ||
    (returnHeaderObj ? `${returnHeaderObj.headerName} (Excel)` : 'Masaüstü Excel Verisi');

  if (!cellData[0]) cellData[0] = {};
  if (targetColIndex === 'NEW_COL' || !cellData[0][finalTargetCol]?.v) {
    cellData[0][finalTargetCol] = {
      v: headerTitle,
      t: 1,
      s: {
        bg: { rgb: '#ecfdf5' },
        cl: { rgb: '#065f46' },
        bl: 1,
      },
    };
  }

  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((n) => !isNaN(n) && n > 0)
    .sort((a, b) => a - b);

  const srcKeyLetter = colIndexToLetter(sourceKeyColIndex);
  const extKeyLetter = colIndexToLetter(externalKeyColIndex);
  const extRetLetter = colIndexToLetter(externalReturnColIndex);
  const extMaxRow = Math.max(linkedInfo.rowCount, 2);
  const extSheetName = linkedInfo.linkedSheetName;

  // VLOOKUP için tablo aralığı बाईं (sol) anahtar sütundan sağdaki hedef sütuna kadar olmalı
  const startColIdx = Math.min(externalKeyColIndex, externalReturnColIndex);
  const endColIdx = Math.max(externalKeyColIndex, externalReturnColIndex);
  const vlookupStartLetter = colIndexToLetter(startColIdx);
  const vlookupEndLetter = colIndexToLetter(endColIdx);
  const vlookupColNum = externalReturnColIndex - startColIdx + 1;

  let matchedCount = 0;
  let totalRowsProcessed = 0;
  let sampleFormula = '';

  for (const r of rowKeys) {
    const keyVal = cellData[r]?.[sourceKeyColIndex]?.v;
    if (keyVal === undefined || keyVal === null || String(keyVal).trim() === '') continue;
    const keyUpper = String(keyVal).toUpperCase();
    if (keyUpper.includes('GENEL TOPLAM') || keyUpper === 'TOPLAM') continue;

    const excelRowNum = r + 1;
    const lookupCellRef = `${srcKeyLetter}${excelRowNum}`;

    let formulaStr = '';
    let computedValue: any = '#N/A';

    if (formulaType === 'VLOOKUP' && externalReturnColIndex >= externalKeyColIndex) {
      const tableArrayRef = `${extSheetName}!${vlookupStartLetter}2:${vlookupEndLetter}${extMaxRow}`;
      const res = evaluateFormulaPreview(
        'VLOOKUP',
        {
          lookup_value: lookupCellRef,
          table_array: tableArrayRef,
          col_index_num: String(vlookupColNum),
          range_lookup: '0',
        },
        cloned,
        sid
      );
      formulaStr = res.formulaString;
      computedValue = res.computedValue;
    } else {
      const lookupArrayRef = `${extSheetName}!${extKeyLetter}2:${extKeyLetter}${extMaxRow}`;
      const returnArrayRef = `${extSheetName}!${extRetLetter}2:${extRetLetter}${extMaxRow}`;
      const res = evaluateFormulaPreview(
        'XLOOKUP',
        {
          lookup_value: lookupCellRef,
          lookup_array: lookupArrayRef,
          return_array: returnArrayRef,
          if_not_found: '"Eşleşmedi"',
        },
        cloned,
        sid
      );
      formulaStr = res.formulaString;
      computedValue = res.computedValue;
    }

    if (!sampleFormula) sampleFormula = formulaStr;
    totalRowsProcessed++;
    if (computedValue !== '#N/A' && computedValue !== 'Eşleşmedi') {
      matchedCount++;
    }

    if (!cellData[r]) cellData[r] = {};
    cellData[r][finalTargetCol] = {
      v: computedValue,
      f: formulaStr,
      t: typeof computedValue === 'number' ? 2 : 1,
      s: {
        bg: { rgb: '#f0fdf4' },
        cl: { rgb: '#065f46' },
        bl: 1,
      },
    };
  }

  return {
    updatedWorkbook: cloned,
    matchedCount,
    totalRowsProcessed,
    targetColLetter: colIndexToLetter(finalTargetCol),
    sampleFormula,
  };
}

/**
 * Kullanıcının kendi bilgisayarında gerçek Microsoft Excel ile hemen test edebilmesi için
 * mevcut SyncSheet tablosundaki birincil anahtarlarla (A sütunu) eşleşen örnek bir
 * gerçek .xlsx dosyası üretip indirir.
 */
export function downloadSampleDesktopExcelForVlookup(workbookData: UniverWorkbookData): void {
  const sid = workbookData.sheetOrder?.[0] || Object.keys(workbookData.sheets || {})[0];
  const cellData = workbookData.sheets?.[sid]?.cellData || {};

  const primaryHeader = String(cellData[0]?.[0]?.v || 'PROJE KODU');
  const keys: string[] = [];

  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((n) => !isNaN(n) && n > 0)
    .sort((a, b) => a - b);

  for (const r of rowKeys) {
    const val = String(cellData[r]?.[0]?.v || '').trim();
    if (!val || val.toUpperCase().includes('TOPLAM')) continue;
    keys.push(val);
    if (keys.length >= 20) break;
  }

  if (keys.length === 0) {
    keys.push('PRJ-2026-01', 'PRJ-2026-02', 'PRJ-2026-03', 'PRJ-2026-04', 'PRJ-2026-05');
  }

  const vendors = [
    'KoçSistem Kurumsal A.Ş.',
    'Aselsan Teknoloji Çözümleri',
    'Logo Yazılım Kurumsal ERP',
    'Turkcell Dijital İş Servisleri',
    'SabancıDx Bulut Hizmetleri',
    'Havelsan Siber Güvenlik',
  ];
  const costCenters = ['TR-IST-101', 'TR-ANK-204', 'TR-IZM-309', 'TR-IST-412', 'TR-BRS-501'];
  const riskLevels = ['Düşük Risk', 'Orta Risk', 'Kritik Öncelik', 'Onaylı - Güvenli'];

  const aoa: any[][] = [
    [
      primaryHeader,
      'ERP MALIYET MERKEZI',
      'TEDARIKCI FIRMA',
      'EK ONAYLI LIMIT (TL)',
      'DENETIM DURUMU',
    ],
  ];

  keys.forEach((k, idx) => {
    aoa.push([
      k,
      costCenters[idx % costCenters.length],
      vendors[idx % vendors.length],
      (idx + 2) * 175000,
      riskLevels[idx % riskLevels.length],
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [
    { wch: 18 },
    { wch: 22 },
    { wch: 30 },
    { wch: 22 },
    { wch: 20 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'ERP_Dis_Veri');
  XLSX.writeFile(wb, 'Masaustu_ERP_Vlookup_Tablosu.xlsx');
}
