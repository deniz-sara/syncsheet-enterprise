import { colIndexToLetter } from './collaboration';

export interface SheetColumnMeta {
  colIndex: number;
  colLetter: string;
  header: string;
  hasTextHeader: boolean;
  isNumeric: boolean;
  sampleValues: Array<string | number>;
  firstDataRow: number; // 1-based Excel row (başlık yoksa 1, varsa 2)
  lastDataRow: number; // 1-based Excel row
}

export interface CopilotRowWrite {
  row: number; // 0-based
  col: number; // 0-based
  formula: string;
  computedValue: string | number;
}

export interface CopilotResult {
  title: string;
  explanation: string;
  formula: string;
  computedPreview: string | number;
  targetCellRef: string;
  canApplyToColumn: boolean;
  columnHeaderSuggestion?: string;
  columnWrites?: CopilotRowWrite[];
  category: 'sum' | 'average' | 'minmax' | 'count' | 'if' | 'math' | 'lookup' | 'text';
}

export interface CopilotSuggestion {
  label: string;
  prompt: string;
  badge: string;
}

function letterToColIndex(letters: string): number {
  let col = 0;
  const clean = letters.toUpperCase().trim();
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return Math.max(0, col - 1);
}

function parseNumberSafe(raw: any): number | null {
  if (raw === undefined || raw === null || String(raw).trim() === '') return null;
  if (typeof raw === 'number' && !Number.isNaN(raw)) return raw;
  const cleaned = String(raw)
    .replace(/[₺$€%\s]/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const n = Number(cleaned);
  return Number.isNaN(n) ? null : n;
}

function readCellNumber(cellData: any, row0: number, col0: number): number {
  const n = parseNumberSafe(cellData?.[row0]?.[col0]?.v);
  return n !== null ? n : 0;
}

function readCellRaw(cellData: any, row0: number, col0: number): any {
  return cellData?.[row0]?.[col0]?.v;
}

export function inspectActiveSheetColumns(workbookData: any, sheetId?: string): {
  sheetId: string;
  sheetName: string;
  columns: SheetColumnMeta[];
  maxDataRow: number; // 1-based
} {
  const sheets = workbookData?.sheets || {};
  const targetId =
    sheetId && sheets[sheetId] ? sheetId : Object.keys(sheets)[0] || 'sheet_1';
  const sheet = sheets[targetId] || { name: 'Sayfa1', cellData: {} };
  const cellData = sheet.cellData || {};

  let maxRowIdx = -1;
  let maxColIdx = -1;

  Object.keys(cellData).forEach((rKey) => {
    const r = Number(rKey);
    if (Number.isNaN(r)) return;
    const rowObj = cellData[r] || {};
    Object.keys(rowObj).forEach((cKey) => {
      const c = Number(cKey);
      if (Number.isNaN(c)) return;
      const val = rowObj[c]?.v;
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        if (r > maxRowIdx) maxRowIdx = r;
        if (c > maxColIdx) maxColIdx = c;
      }
    });
  });

  if (maxRowIdx === -1 || maxColIdx === -1) {
    return {
      sheetId: targetId,
      sheetName: sheet.name || 'Sayfa1',
      columns: [],
      maxDataRow: 10,
    };
  }

  const maxDataRow = Math.max(1, maxRowIdx + 1);
  const columns: SheetColumnMeta[] = [];

  const scanCols = Math.min(maxColIdx + 1, 26);
  for (let c = 0; c < scanCols; c++) {
    const colLetter = colIndexToLetter(c);
    const rawRow0 = cellData[0]?.[c]?.v;
    const row0Num = parseNumberSafe(rawRow0);
    const hasTextHeader =
      rawRow0 !== undefined &&
      rawRow0 !== null &&
      String(rawRow0).trim() !== '' &&
      row0Num === null;

    const header = hasTextHeader
      ? String(rawRow0).trim()
      : `${colLetter} Sütunu`;

    const startRow0 = hasTextHeader ? 1 : 0;
    const samples: Array<string | number> = [];
    let numCount = 0;
    let totalCount = 0;
    let lastNonEmptyRow1Based = hasTextHeader ? 2 : 1;

    for (let r = startRow0; r <= maxRowIdx; r++) {
      const v = cellData[r]?.[c]?.v;
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        totalCount++;
        lastNonEmptyRow1Based = r + 1;
        const parsed = parseNumberSafe(v);
        if (parsed !== null) {
          numCount++;
          if (samples.length < 4) samples.push(parsed);
        } else {
          if (samples.length < 4) samples.push(String(v));
        }
      }
    }

    if (totalCount > 0 || hasTextHeader) {
      columns.push({
        colIndex: c,
        colLetter,
        header,
        hasTextHeader,
        isNumeric: totalCount > 0 && numCount >= Math.ceil(totalCount * 0.5),
        sampleValues: samples,
        firstDataRow: hasTextHeader ? 2 : 1,
        lastDataRow: Math.max(hasTextHeader ? 2 : 1, lastNonEmptyRow1Based),
      });
    }
  }

  return {
    sheetId: targetId,
    sheetName: sheet.name || 'Sayfa1',
    columns,
    maxDataRow,
  };
}

function readNumericValuesInRange(
  cellData: any,
  startCol: number,
  startRow1Based: number,
  endCol: number,
  endRow1Based: number
): number[] {
  const nums: number[] = [];
  const r1 = Math.min(startRow1Based, endRow1Based) - 1;
  const r2 = Math.max(startRow1Based, endRow1Based) - 1;
  const c1 = Math.min(startCol, endCol);
  const c2 = Math.max(startCol, endCol);

  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const n = parseNumberSafe(cellData?.[r]?.[c]?.v);
      if (n !== null) nums.push(n);
    }
  }
  return nums;
}

function normalizeTr(str: string): string {
  return str
    .toLowerCase()
    .replace(/ı/g, 'i')
    .replace(/ğ/g, 'g')
    .replace(/ü/g, 'u')
    .replace(/ş/g, 's')
    .replace(/ö/g, 'o')
    .replace(/ç/g, 'c')
    .trim();
}

interface ParsedCellRef {
  ref: string; // Örn: "A1"
  colLetter: string; // Örn: "A"
  colIndex: number; // 0-based
  row1Based: number; // 1-based
  row0: number; // 0-based
}

function extractExplicitCellRefs(prompt: string): ParsedCellRef[] {
  const refs: ParsedCellRef[] = [];
  const regex = /\b([A-Za-z]{1,2})(\d{1,5})\b/g;
  let m: RegExpExecArray | null;
  while ((m = regex.exec(prompt)) !== null) {
    const colLetter = m[1].toUpperCase();
    const row1Based = Number(m[2]);
    if (row1Based >= 1 && row1Based <= 100000) {
      const colIndex = letterToColIndex(colLetter);
      if (colIndex >= 0 && colIndex < 52) {
        refs.push({
          ref: `${colLetter}${row1Based}`,
          colLetter,
          colIndex,
          row1Based,
          row0: row1Based - 1,
        });
      }
    }
  }
  return refs;
}

function extractExplicitRange(prompt: string): {
  rangeStr: string;
  start: ParsedCellRef;
  end: ParsedCellRef;
} | null {
  // Matches "A1:A10", "A1:B5", or Turkish "A1'den A10'a kadar", "A1 ile A10 arası"
  const colonMatch = prompt.match(/\b([A-Za-z]{1,2})(\d{1,5})\s*:\s*([A-Za-z]{1,2})(\d{1,5})\b/);
  if (colonMatch) {
    const c1 = colonMatch[1].toUpperCase();
    const r1 = Number(colonMatch[2]);
    const c2 = colonMatch[3].toUpperCase();
    const r2 = Number(colonMatch[4]);
    return {
      rangeStr: `${c1}${r1}:${c2}${r2}`,
      start: { ref: `${c1}${r1}`, colLetter: c1, colIndex: letterToColIndex(c1), row1Based: r1, row0: r1 - 1 },
      end: { ref: `${c2}${r2}`, colLetter: c2, colIndex: letterToColIndex(c2), row1Based: r2, row0: r2 - 1 },
    };
  }

  const norm = normalizeTr(prompt);
  const rangeTrMatch = norm.match(
    /\b([a-z]{1,2})(\d{1,5})(?:['’]?[a-z]*)?\s+(?:ile|den|dan|-)\s+([a-z]{1,2})(\d{1,5})(?:['’]?[a-z]*)?\s*(?:arasi|arasini|kadar)/
  );
  if (rangeTrMatch) {
    const c1 = rangeTrMatch[1].toUpperCase();
    const r1 = Number(rangeTrMatch[2]);
    const c2 = rangeTrMatch[3].toUpperCase();
    const r2 = Number(rangeTrMatch[4]);
    return {
      rangeStr: `${c1}${r1}:${c2}${r2}`,
      start: { ref: `${c1}${r1}`, colLetter: c1, colIndex: letterToColIndex(c1), row1Based: r1, row0: r1 - 1 },
      end: { ref: `${c2}${r2}`, colLetter: c2, colIndex: letterToColIndex(c2), row1Based: r2, row0: r2 - 1 },
    };
  }

  return null;
}

function extractColumnMentions(
  prompt: string,
  columns: SheetColumnMeta[],
  maxDataRow: number
): SheetColumnMeta[] {
  const matched: SheetColumnMeta[] = [];
  const norm = normalizeTr(prompt);

  // 1) Explicit "A sütunu", "B kolonu", "C sütunundaki"
  const colRegex = /\b([a-z])\s*(?:sutun|kolon)/gi;
  let m: RegExpExecArray | null;
  while ((m = colRegex.exec(norm)) !== null) {
    const letter = m[1].toUpperCase();
    const colIndex = letterToColIndex(letter);
    const existing = columns.find((c) => c.colLetter === letter) || {
      colIndex,
      colLetter: letter,
      header: `${letter} Sütunu`,
      hasTextHeader: false,
      isNumeric: true,
      sampleValues: [],
      firstDataRow: 1,
      lastDataRow: Math.max(2, maxDataRow),
    };
    if (!matched.some((x) => x.colIndex === existing.colIndex)) {
      matched.push(existing);
    }
  }

  // 2) Match by real text header name in the sheet
  columns.forEach((col) => {
    if (!col.hasTextHeader) return;
    const normHeader = normalizeTr(col.header);
    if (normHeader && normHeader.length >= 2 && norm.includes(normHeader)) {
      if (!matched.some((x) => x.colIndex === col.colIndex)) {
        matched.push(col);
      }
    }
  });

  return matched;
}

export function getCopilotSmartSuggestions(
  workbookData: any,
  sheetId?: string
): CopilotSuggestion[] {
  const { columns } = inspectActiveSheetColumns(workbookData, sheetId);
  const namedCols = columns.filter((c) => c.hasTextHeader);

  // Eğer tabloda kullanıcının kendi yazdığı gerçek sütun başlıkları varsa onlara göre öner,
  // yoksa net, anlaşılır standart hücre/sütun örnekleri göster.
  if (namedCols.length >= 2) {
    const numCols = namedCols.filter((c) => c.isNumeric);
    const c1 = numCols[0] || namedCols[0];
    const c2 = numCols[1] || namedCols.find((x) => x.colIndex !== c1.colIndex) || c1;
    const rStart = c1.firstDataRow;
    const rEnd = Math.max(rStart + 1, c1.lastDataRow);

    return [
      {
        badge: 'TOPLA',
        label: `${c1.colLetter} sütununu (${c1.header}) topla`,
        prompt: `${c1.colLetter}${rStart}:${c1.colLetter}${rEnd} aralığını topla`,
      },
      {
        badge: 'ORTALAMA',
        label: `${c1.colLetter} sütunu (${c1.header}) ortalaması`,
        prompt: `${c1.colLetter}${rStart}:${c1.colLetter}${rEnd} ortalamasını al`,
      },
      {
        badge: 'ÇARPIM',
        label: `${c1.colLetter}2 ile ${c2.colLetter}2 hücresini çarp`,
        prompt: `${c1.colLetter}2 ile ${c2.colLetter}2 hücresini çarp`,
      },
      {
        badge: 'KOŞUL (EĞER)',
        label: `Eğer ${c1.colLetter}2 > 100 ise Onaylı yaz`,
        prompt: `Eğer ${c1.colLetter}2 100'den büyükse Onaylı değilse Beklemede yaz`,
      },
      {
        badge: 'YÜZDE / KDV',
        label: `${c1.colLetter}2 hücresinin %20'sini hesapla`,
        prompt: `${c1.colLetter}2 hücresinin %20'sini hesapla`,
      },
      {
        badge: 'DÜŞEYARA',
        label: `A2 değerini ara, ${c1.colLetter} sütununu getir`,
        prompt: `A2 değerini A sütununda ara ${c1.colLetter} sütununu getir`,
      },
    ];
  }

  return [
    {
      badge: 'HÜCRE TOPLA',
      label: 'A1 ile B1 hücrelerini topla (=A1+B1)',
      prompt: 'A1 ile B1 hücresini topla',
    },
    {
      badge: 'ARALIK TOPLA',
      label: 'A1:A10 aralığının toplamını al (=SUM)',
      prompt: 'A1:A10 aralığını topla',
    },
    {
      badge: 'ÇARPMA / BÖLME',
      label: 'A1 ile B1 hücresini çarp (=A1*B1)',
      prompt: 'A1 ile B1 hücresini çarp',
    },
    {
      badge: 'ORTALAMA',
      label: 'A1:A10 aralığının ortalamasını bul (=AVERAGE)',
      prompt: 'A1:A10 ortalamasını hesapla',
    },
    {
      badge: 'YÜZDE / KDV',
      label: 'A1 hücresine %20 ekle (=A1*1.2)',
      prompt: 'A1 hücresine %20 ekle',
    },
    {
      badge: 'EĞER (IF)',
      label: 'A1 > 50 ise "Geçti" değilse "Kaldı" yaz',
      prompt: 'Eğer A1 50 den büyükse Geçti değilse Kaldı yaz',
    },
  ];
}

export function generateCopilotSolution(
  prompt: string,
  workbookData: any,
  activeRow: number = 0,
  activeCol: number = 0,
  sheetId?: string
): CopilotResult {
  const { sheetId: targetSheetId, columns, maxDataRow } = inspectActiveSheetColumns(
    workbookData,
    sheetId
  );
  const cellData = workbookData?.sheets?.[targetSheetId]?.cellData || {};
  const trimmed = prompt.trim();
  const norm = normalizeTr(trimmed);
  const targetCellRef = `${colIndexToLetter(activeCol)}${activeRow + 1}`;

  // 0) Kullanıcı doğrudan "=" ile başlayan veya "A1+B1" gibi matematiksel bir ifade yazdıysa
  const cleanMathExpr = trimmed.replace(/^=/, '').replace(/\s+/g, '').toUpperCase();
  if (/^[A-Z]{1,2}\d+(?:[+\-*/][A-Z0-9.,]+)+$/.test(cleanMathExpr)) {
    const refs = extractExplicitCellRefs(cleanMathExpr);
    let evalExpr = cleanMathExpr;
    refs.forEach((r) => {
      const val = readCellNumber(cellData, r.row0, r.colIndex);
      evalExpr = evalExpr.replace(new RegExp(`\\b${r.ref}\\b`, 'g'), String(val));
    });
    let computed: number | string = 0;
    try {
      // Sadece güvenli sayı ve dört işlem karakterleri içeriyorsa hesapla
      if (/^[0-9.+\-*/()]+$/.test(evalExpr)) {
        computed = Math.round(Number(Function(`"use strict"; return (${evalExpr})`)()) * 100) / 100;
      }
    } catch {
      computed = 0;
    }
    return {
      title: `Hücre Hesaplama Formülü (= ${cleanMathExpr})`,
      explanation: `Yazdığınız ${cleanMathExpr} matematiksel işlemini doğrudan Excel formülüne dönüştürür.`,
      formula: `=${cleanMathExpr}`,
      computedPreview: computed,
      targetCellRef,
      canApplyToColumn: false,
      category: 'math',
    };
  }

  const explicitRange = extractExplicitRange(trimmed);
  const explicitCells = extractExplicitCellRefs(trimmed);
  const mentionedCols = extractColumnMentions(trimmed, columns, maxDataRow);

  const primaryCol =
    mentionedCols[0] ||
    columns.find((c) => c.isNumeric) ||
    columns[0] || {
      colIndex: 0,
      colLetter: 'A',
      header: 'A Sütunu',
      hasTextHeader: false,
      isNumeric: true,
      sampleValues: [],
      firstDataRow: 1,
      lastDataRow: Math.max(2, maxDataRow),
    };

  const secondaryCol =
    mentionedCols[1] ||
    columns.find((c) => c.colIndex !== primaryCol.colIndex) || {
      colIndex: primaryCol.colIndex + 1,
      colLetter: colIndexToLetter(primaryCol.colIndex + 1),
      header: `${colIndexToLetter(primaryCol.colIndex + 1)} Sütunu`,
      hasTextHeader: false,
      isNumeric: true,
      sampleValues: [],
      firstDataRow: primaryCol.firstDataRow,
      lastDataRow: primaryCol.lastDataRow,
    };

  const nextEmptyColIdx =
    columns.length > 0 ? Math.max(...columns.map((c) => c.colIndex)) + 1 : activeCol;

  // 1) DÜŞEYARA / XLOOKUP / VLOOKUP
  if (
    norm.includes('duseyara') ||
    norm.includes('xlookup') ||
    norm.includes('vlookup') ||
    (norm.includes('ara') && norm.includes('getir'))
  ) {
    const lookupCell = explicitCells[0]?.ref || 'A2';
    const searchCol = mentionedCols[0] || columns[0] || primaryCol;
    const returnCol = mentionedCols[1] || columns[1] || secondaryCol;
    const endR = Math.max(searchCol.lastDataRow, returnCol.lastDataRow, 10);
    const formula = `=XLOOKUP(${lookupCell}, ${searchCol.colLetter}1:${searchCol.colLetter}${endR}, ${returnCol.colLetter}1:${returnCol.colLetter}${endR}, "Bulunamadı")`;
    const previewVal =
      readCellRaw(cellData, 1, returnCol.colIndex) ??
      readCellRaw(cellData, 0, returnCol.colIndex) ??
      'Eşleşen Değer';

    return {
      title: `Düşeyara (XLOOKUP): ${lookupCell} → ${returnCol.colLetter} Sütunu`,
      explanation: `${lookupCell} hücresindeki değeri ${searchCol.colLetter} sütununda arar ve karşılık gelen ${returnCol.colLetter} sütunu değerini getirir.`,
      formula,
      computedPreview: previewVal,
      targetCellRef,
      canApplyToColumn: false,
      category: 'lookup',
    };
  }

  // 2) EĞER / KOŞULLU İFADE (IF)
  if (
    norm.includes('eger') ||
    norm.includes('buyukse') ||
    norm.includes('kucukse') ||
    norm.includes('esitse') ||
    norm.includes('degilse') ||
    norm.includes('ise ')
  ) {
    const targetRefCell = explicitCells[0] || {
      ref: `${primaryCol.colLetter}${primaryCol.firstDataRow}`,
      colLetter: primaryCol.colLetter,
      colIndex: primaryCol.colIndex,
      row1Based: primaryCol.firstDataRow,
      row0: primaryCol.firstDataRow - 1,
    };

    // Hücre referansı içindeki satır numarasını hariç tutarak eşik sayısını bul
    const promptWithoutCellRefs = trimmed.replace(/\b[A-Za-z]{1,2}\d{1,5}\b/g, '');
    const thresholdMatch = promptWithoutCellRefs.match(/(\d+(?:[.,]\d+)?)/);
    const threshold = thresholdMatch
      ? Number(thresholdMatch[1].replace(/\./g, '').replace(',', '.'))
      : 100;

    const isLess = norm.includes('kucuk') && !norm.includes('buyuk');
    const isEqual = norm.includes('esit') && !norm.includes('buyuk') && !norm.includes('kucuk');
    const op = isEqual ? '=' : isLess ? '<' : '>';

    let trueLabel = 'Onaylı';
    let falseLabel = 'Beklemede';
    const quoteMatches = trimmed.match(/['"]([^'"]+)['"]/g);
    if (quoteMatches && quoteMatches.length >= 2) {
      trueLabel = quoteMatches[0].replace(/['"]/g, '');
      falseLabel = quoteMatches[1].replace(/['"]/g, '');
    } else {
      // "... büyükse X değilse Y yaz" kalıbını yakala
      const patternMatch = trimmed.match(
        /(?:büyükse|buyukse|küçükse|kucukse|eşitse|esitse|ise)\s+([^\s,]+)\s+(?:değilse|degilse|aksi halde|yoksa)\s+([^\s,]+?)(?:\s+yaz|\s*$)/i
      );
      if (patternMatch) {
        trueLabel = patternMatch[1].replace(/['"]/g, '');
        falseLabel = patternMatch[2].replace(/['"]/g, '');
      }
    }

    const formula = `=IF(${targetRefCell.ref}${op}${threshold}, "${trueLabel}", "${falseLabel}")`;
    const sampleVal = readCellNumber(cellData, targetRefCell.row0, targetRefCell.colIndex);
    const cond = isEqual
      ? sampleVal === threshold
      : isLess
        ? sampleVal < threshold
        : sampleVal > threshold;

    const writes: CopilotRowWrite[] = [];
    for (let r = primaryCol.firstDataRow - 1; r < primaryCol.lastDataRow; r++) {
      const v = readCellNumber(cellData, r, targetRefCell.colIndex);
      const rowCond = isEqual ? v === threshold : isLess ? v < threshold : v > threshold;
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=IF(${targetRefCell.colLetter}${r + 1}${op}${threshold}, "${trueLabel}", "${falseLabel}")`,
        computedValue: rowCond ? trueLabel : falseLabel,
      });
    }

    return {
      title: `Koşullu Formül (EĞER / IF) — ${targetRefCell.ref} ${op} ${threshold}`,
      explanation: `${targetRefCell.ref} hücresi ${threshold} değerinden ${isEqual ? 'eşitse' : isLess ? 'küçükse' : 'büyükse'} "${trueLabel}", değilse "${falseLabel}" sonucunu verir.`,
      formula,
      computedPreview: cond ? trueLabel : falseLabel,
      targetCellRef,
      canApplyToColumn: writes.length > 1,
      columnHeaderSuggestion: 'Durum',
      columnWrites: writes,
      category: 'if',
    };
  }

  // 3) YÜZDE / KDV / İNDİRİM / ZAM HESAPLAMA
  if (
    norm.includes('kdv') ||
    norm.includes('yuzde') ||
    norm.includes('%') ||
    norm.includes('indirim') ||
    norm.includes('zam') ||
    norm.includes('ekle')
  ) {
    const promptWithoutRefs = trimmed.replace(/\b[A-Za-z]{1,2}\d{1,5}\b/g, '');
    const pctMatch = promptWithoutRefs.match(/(\d+(?:[.,]\d+)?)/);
    const pct = pctMatch ? Number(pctMatch[1].replace(',', '.')) : 20;
    const isAdd =
      norm.includes('dahil') || norm.includes('ekle') || norm.includes('zam') || norm.includes('artir');
    const isDiscount = norm.includes('indirim') || norm.includes('dus');

    const multiplier = isAdd
      ? Number((1 + pct / 100).toFixed(4))
      : isDiscount
        ? Number((1 - pct / 100).toFixed(4))
        : Number((pct / 100).toFixed(4));

    const refCell = explicitCells[0] || {
      ref: `${primaryCol.colLetter}${primaryCol.firstDataRow}`,
      colLetter: primaryCol.colLetter,
      colIndex: primaryCol.colIndex,
      row1Based: primaryCol.firstDataRow,
      row0: primaryCol.firstDataRow - 1,
    };

    const formula = `=${refCell.ref}*${multiplier}`;
    const baseVal = readCellNumber(cellData, refCell.row0, refCell.colIndex);
    const preview = Math.round(baseVal * multiplier * 100) / 100;

    const writes: CopilotRowWrite[] = [];
    for (let r = primaryCol.firstDataRow - 1; r < primaryCol.lastDataRow; r++) {
      const rv = readCellNumber(cellData, r, refCell.colIndex);
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=${refCell.colLetter}${r + 1}*${multiplier}`,
        computedValue: Math.round(rv * multiplier * 100) / 100,
      });
    }

    return {
      title: `${refCell.ref} İçin %${pct} ${isAdd ? 'Artış / KDV Dahil' : isDiscount ? 'İndirimli' : 'Yüzde'} Hesabı`,
      explanation: `${refCell.ref} hücresini ${multiplier} ile çarpar (Mevcut değer: ${baseVal} ➔ Sonuç: ${preview}).`,
      formula,
      computedPreview: preview,
      targetCellRef,
      canApplyToColumn: writes.length > 1,
      columnHeaderSuggestion: `%${pct} Hesap`,
      columnWrites: writes,
      category: 'math',
    };
  }

  // 4) BELİRLİ HÜCRELER ARASI DÖRT İŞLEM (Örn: "A1 ile B1'i topla", "A2 ile B2 çarp", "B1'den C1'i çıkar", "A1'i B1'e böl")
  if (!explicitRange && explicitCells.length >= 2) {
    const isMultiply = norm.includes('carp') || norm.includes('*');
    const isSubtract = norm.includes('cikar') || norm.includes('fark') || norm.includes('eksi');
    const isDivide = norm.includes('bol') || norm.includes('oran') || norm.includes('/');
    const isAverage = norm.includes('ortalama') || norm.includes('average');

    if (isAverage) {
      const refList = explicitCells.map((c) => c.ref).join(', ');
      const vals = explicitCells.map((c) => readCellNumber(cellData, c.row0, c.colIndex));
      const avg =
        vals.length > 0
          ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100) / 100
          : 0;
      return {
        title: `Seçili Hücrelerin Ortalaması (${refList})`,
        explanation: `${refList} hücrelerinin aritmetik ortalamasını hesaplar.`,
        formula: `=AVERAGE(${refList})`,
        computedPreview: avg,
        targetCellRef,
        canApplyToColumn: false,
        category: 'average',
      };
    }

    const op = isMultiply ? '*' : isSubtract ? '-' : isDivide ? '/' : '+';
    const opName =
      op === '*' ? 'Çarpımı' : op === '-' ? 'Farkı' : op === '/' ? 'Bölümü' : 'Toplamı';

    const formula =
      op === '+' && explicitCells.length > 2
        ? `=SUM(${explicitCells.map((c) => c.ref).join(', ')})`
        : `=${explicitCells.map((c) => c.ref).join(op)}`;

    const values = explicitCells.map((c) => readCellNumber(cellData, c.row0, c.colIndex));
    let previewResult = values[0] || 0;
    for (let i = 1; i < values.length; i++) {
      if (op === '+') previewResult += values[i];
      else if (op === '-') previewResult -= values[i];
      else if (op === '*') previewResult *= values[i];
      else if (op === '/') previewResult = values[i] !== 0 ? previewResult / values[i] : 0;
    }
    previewResult = Math.round(previewResult * 100) / 100;

    return {
      title: `${explicitCells.map((c) => c.ref).join(` ${op} `)} Hücrelerinin ${opName}`,
      explanation: `Belirttiğiniz hücrelerin (${explicitCells
        .map((c, idx) => `${c.ref}=${values[idx]}`)
        .join(', ')}) ${opName.toLowerCase()}nı hesaplar.`,
      formula,
      computedPreview: previewResult,
      targetCellRef,
      canApplyToColumn: false,
      category: op === '+' ? 'sum' : 'math',
    };
  }

  // 5) İKİ SÜTUN ARASI ÇARPMA / ÇIKARMA / BÖLME (Örn: "A sütunu ile B sütununu çarp")
  if (
    norm.includes('carp') ||
    norm.includes('cikar') ||
    norm.includes('fark') ||
    norm.includes('bol')
  ) {
    const op =
      norm.includes('cikar') || norm.includes('fark')
        ? '-'
        : norm.includes('bol')
          ? '/'
          : '*';
    const r1 = primaryCol.firstDataRow;
    const maxR = Math.max(primaryCol.lastDataRow, secondaryCol.lastDataRow, r1);
    const formula = `=${primaryCol.colLetter}${r1}${op}${secondaryCol.colLetter}${r1}`;

    const writes: CopilotRowWrite[] = [];
    for (let r = r1 - 1; r < maxR; r++) {
      const a = readCellNumber(cellData, r, primaryCol.colIndex);
      const b = readCellNumber(cellData, r, secondaryCol.colIndex);
      const val =
        op === '*' ? a * b : op === '-' ? a - b : b !== 0 ? a / b : 0;
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=${primaryCol.colLetter}${r + 1}${op}${secondaryCol.colLetter}${r + 1}`,
        computedValue: Math.round(val * 100) / 100,
      });
    }

    const firstA = readCellNumber(cellData, r1 - 1, primaryCol.colIndex);
    const firstB = readCellNumber(cellData, r1 - 1, secondaryCol.colIndex);
    const firstRes =
      op === '*' ? firstA * firstB : op === '-' ? firstA - firstB : firstB !== 0 ? firstA / firstB : 0;

    return {
      title: `${primaryCol.colLetter} ve ${secondaryCol.colLetter} Sütunları İşlemi (${op})`,
      explanation: `${primaryCol.colLetter}${r1} (${firstA}) ve ${secondaryCol.colLetter}${r1} (${firstB}) hücrelerine ${op} işlemini uygular.`,
      formula,
      computedPreview: Math.round(firstRes * 100) / 100,
      targetCellRef,
      canApplyToColumn: writes.length > 1,
      columnHeaderSuggestion: `${primaryCol.colLetter}${op}${secondaryCol.colLetter}`,
      columnWrites: writes,
      category: 'math',
    };
  }

  // 6) ARALIK VEYA SÜTUN BELİRLEME (Örn: "A1:A10" veya "A sütunu")
  const effectiveRange = explicitRange
    ? explicitRange.rangeStr
    : `${primaryCol.colLetter}${primaryCol.firstDataRow}:${primaryCol.colLetter}${Math.max(
        primaryCol.firstDataRow + 1,
        primaryCol.lastDataRow
      )}`;

  const rangeNums = explicitRange
    ? readNumericValuesInRange(
        cellData,
        explicitRange.start.colIndex,
        explicitRange.start.row1Based,
        explicitRange.end.colIndex,
        explicitRange.end.row1Based
      )
    : readNumericValuesInRange(
        cellData,
        primaryCol.colIndex,
        primaryCol.firstDataRow,
        primaryCol.colIndex,
        Math.max(primaryCol.firstDataRow + 1, primaryCol.lastDataRow)
      );

  // ORTALAMA (AVERAGE)
  if (norm.includes('ortalama') || norm.includes('average') || norm.includes('ort')) {
    const sum = rangeNums.reduce((a, b) => a + b, 0);
    const avg = rangeNums.length > 0 ? Math.round((sum / rangeNums.length) * 100) / 100 : 0;
    return {
      title: `${effectiveRange} Aralığının Ortalaması (AVERAGE)`,
      explanation: `${effectiveRange} aralığındaki ${rangeNums.length} sayısal hücrenin ortalamasını hesaplar.`,
      formula: `=AVERAGE(${effectiveRange})`,
      computedPreview: avg,
      targetCellRef,
      canApplyToColumn: false,
      category: 'average',
    };
  }

  // EN BÜYÜK (MAX) / EN KÜÇÜK (MIN)
  if (
    norm.includes('en buyuk') ||
    norm.includes('en yuksek') ||
    norm.includes('maksimum') ||
    norm.includes('max')
  ) {
    const maxVal = rangeNums.length > 0 ? Math.max(...rangeNums) : 0;
    return {
      title: `${effectiveRange} En Büyük Değer (MAX)`,
      explanation: `${effectiveRange} aralığındaki en yüksek sayısal değeri bulur.`,
      formula: `=MAX(${effectiveRange})`,
      computedPreview: maxVal,
      targetCellRef,
      canApplyToColumn: false,
      category: 'minmax',
    };
  }

  if (
    norm.includes('en kucuk') ||
    norm.includes('en dusuk') ||
    norm.includes('minimum') ||
    norm.includes('min')
  ) {
    const minVal = rangeNums.length > 0 ? Math.min(...rangeNums) : 0;
    return {
      title: `${effectiveRange} En Küçük Değer (MIN)`,
      explanation: `${effectiveRange} aralığındaki en düşük sayısal değeri bulur.`,
      formula: `=MIN(${effectiveRange})`,
      computedPreview: minVal,
      targetCellRef,
      canApplyToColumn: false,
      category: 'minmax',
    };
  }

  // SAYI / ADET (COUNTA)
  if (
    norm.includes('kac adet') ||
    norm.includes('kac tane') ||
    norm.includes('sayisi') ||
    norm.includes('count')
  ) {
    return {
      title: `${effectiveRange} Dolu Hücre Sayısı (COUNTA)`,
      explanation: `${effectiveRange} aralığındaki dolu hücre adedini sayar.`,
      formula: `=COUNTA(${effectiveRange})`,
      computedPreview: rangeNums.length,
      targetCellRef,
      canApplyToColumn: false,
      category: 'count',
    };
  }

  // VARSAYILAN: TOPLAM (SUM)
  const totalSum = Math.round(rangeNums.reduce((a, b) => a + b, 0) * 100) / 100;
  return {
    title: `${effectiveRange} Aralığının Toplamı (SUM)`,
    explanation: `${effectiveRange} aralığındaki sayısal değerlerin toplamını alır (${rangeNums.length} dolu sayısal hücre bulundu).`,
    formula: `=SUM(${effectiveRange})`,
    computedPreview: totalSum,
    targetCellRef,
    canApplyToColumn: false,
    category: 'sum',
  };
}
