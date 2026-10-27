import { colIndexToLetter } from './collaboration';

export interface SheetColumnMeta {
  colIndex: number;
  colLetter: string;
  header: string;
  isNumeric: boolean;
  sampleValues: Array<string | number>;
  firstDataRow: number; // 1-based Excel row (genelde 2)
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

export function inspectActiveSheetColumns(workbookData: any, sheetId?: string): {
  sheetId: string;
  sheetName: string;
  columns: SheetColumnMeta[];
  maxDataRow: number; // 1-based
} {
  const sheets = workbookData?.sheets || {};
  const targetId =
    sheetId && sheets[sheetId] ? sheetId : Object.keys(sheets)[0] || 'sheet_main';
  const sheet = sheets[targetId] || { name: 'Sayfa1', cellData: {} };
  const cellData = sheet.cellData || {};

  let maxRowIdx = 0;
  let maxColIdx = 0;

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

  const maxDataRow = Math.max(2, maxRowIdx + 1);
  const columns: SheetColumnMeta[] = [];

  const scanCols = Math.min(Math.max(maxColIdx + 1, 6), 26);
  for (let c = 0; c < scanCols; c++) {
    const colLetter = colIndexToLetter(c);
    const rawHeader = cellData[0]?.[c]?.v;
    const header =
      rawHeader !== undefined && rawHeader !== null && String(rawHeader).trim() !== ''
        ? String(rawHeader).trim()
        : `${colLetter} Sütunu`;

    const samples: Array<string | number> = [];
    let numCount = 0;
    let totalCount = 0;
    let lastNonEmptyRow = 1;

    for (let r = 1; r <= maxRowIdx; r++) {
      const v = cellData[r]?.[c]?.v;
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        totalCount++;
        lastNonEmptyRow = r + 1;
        const parsed =
          typeof v === 'number'
            ? v
            : Number(
                String(v)
                  .replace(/[₺$€%\s]/g, '')
                  .replace(/\./g, '')
                  .replace(',', '.')
              );
        if (!Number.isNaN(parsed) && String(v).trim() !== '') {
          numCount++;
          if (samples.length < 4) samples.push(parsed);
        } else {
          if (samples.length < 4) samples.push(String(v));
        }
      }
    }

    if (totalCount > 0 || (rawHeader !== undefined && String(rawHeader).trim() !== '')) {
      columns.push({
        colIndex: c,
        colLetter,
        header,
        isNumeric: totalCount > 0 && numCount >= Math.ceil(totalCount * 0.5),
        sampleValues: samples,
        firstDataRow: 2,
        lastDataRow: Math.max(2, lastNonEmptyRow),
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

function readNumericValuesInColumn(
  cellData: any,
  colIndex: number,
  startRow1Based: number,
  endRow1Based: number
): number[] {
  const nums: number[] = [];
  for (let r = startRow1Based - 1; r <= endRow1Based - 1; r++) {
    const raw = cellData?.[r]?.[colIndex]?.v;
    if (raw === undefined || raw === null || String(raw).trim() === '') continue;
    if (typeof raw === 'number' && !Number.isNaN(raw)) {
      nums.push(raw);
    } else {
      const cleaned = String(raw)
        .replace(/[₺$€%\s]/g, '')
        .replace(/\./g, '')
        .replace(',', '.');
      const n = Number(cleaned);
      if (!Number.isNaN(n)) nums.push(n);
    }
  }
  return nums;
}

function readCellNumber(cellData: any, row0: number, col0: number): number {
  const raw = cellData?.[row0]?.[col0]?.v;
  if (typeof raw === 'number') return raw;
  if (!raw) return 0;
  const cleaned = String(raw)
    .replace(/[₺$€%\s]/g, '')
    .replace(/\./g, '')
    .replace(',', '.');
  const n = Number(cleaned);
  return Number.isNaN(n) ? 0 : n;
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

function matchColumnsFromPrompt(
  prompt: string,
  columns: SheetColumnMeta[]
): SheetColumnMeta[] {
  const matched: SheetColumnMeta[] = [];
  const normPrompt = normalizeTr(prompt);

  // 1) Explicit column letters like "A sütunu", "B kolonu", "C2:C10", "B ile C"
  const letterRegex = /\b([a-zA-Z])\s*(?:sutun|kolon|hucr|\d|\b)/gi;
  let m: RegExpExecArray | null;
  while ((m = letterRegex.exec(normPrompt)) !== null) {
    const letter = m[1].toUpperCase();
    const found = columns.find((c) => c.colLetter === letter);
    if (found && !matched.some((x) => x.colIndex === found.colIndex)) {
      matched.push(found);
    }
  }

  // 2) Match by column header name
  columns.forEach((col) => {
    const normHeader = normalizeTr(col.header);
    if (normHeader && normHeader.length >= 2 && normPrompt.includes(normHeader)) {
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
  const numericCols = columns.filter((c) => c.isNumeric);
  const firstNum = numericCols[0] || columns[1] || columns[0];
  const secondNum = numericCols[1] || columns[2] || firstNum;
  const firstText = columns.find((c) => !c.isNumeric) || columns[0];

  const colAName = firstNum ? `${firstNum.colLetter} (${firstNum.header})` : 'B sütunu';
  const colALetter = firstNum ? firstNum.colLetter : 'B';
  const colBLetter = secondNum ? secondNum.colLetter : 'C';
  const textLetter = firstText ? firstText.colLetter : 'A';

  return [
    {
      badge: 'TOPLAM',
      label: `${colAName} toplamını hesapla`,
      prompt: `${colALetter} sütunundaki değerlerin toplamını al`,
    },
    {
      badge: 'ORTALAMA',
      label: `${colAName} ortalamasını bul`,
      prompt: `${colALetter} sütunundaki değerlerin ortalamasını hesapla`,
    },
    {
      badge: 'TÜM SATIRLAR',
      label: `${colALetter} sütununa %20 KDV hesapla`,
      prompt: `${colALetter} sütunundaki değerlerin %20 KDV tutarını tüm satırlara hesapla`,
    },
    {
      badge: 'ÇARPIM',
      label: `${colALetter} ve ${colBLetter} sütunlarını çarp`,
      prompt: `${colALetter} sütunu ile ${colBLetter} sütununu çarp ve tüm satırlara uygula`,
    },
    {
      badge: 'KOŞUL (EĞER)',
      label: `Koşullu Durum (Onaylı / Beklemede)`,
      prompt: `Eğer ${colALetter}2 10000'den büyükse Onaylı değilse Beklemede yaz`,
    },
    {
      badge: 'DÜŞEYARA',
      label: `${textLetter} sütununa göre Düşeyara (XLOOKUP)`,
      prompt: `${textLetter}2 değerini ${textLetter} sütununda ara ve ${colALetter} sütunundaki karşılığını getir`,
    },
  ];
}

export function generateCopilotSolution(
  prompt: string,
  workbookData: any,
  activeRow: number = 1,
  activeCol: number = 0,
  sheetId?: string
): CopilotResult {
  const { sheetId: targetSheetId, columns, maxDataRow } = inspectActiveSheetColumns(
    workbookData,
    sheetId
  );
  const cellData = workbookData?.sheets?.[targetSheetId]?.cellData || {};
  const norm = normalizeTr(prompt);
  const matchedCols = matchColumnsFromPrompt(prompt, columns);
  const numericCols = columns.filter((c) => c.isNumeric);

  const primaryCol =
    matchedCols[0] ||
    numericCols[0] ||
    columns[Math.min(activeCol, Math.max(0, columns.length - 1))] || {
      colIndex: 1,
      colLetter: 'B',
      header: 'B Sütunu',
      isNumeric: true,
      sampleValues: [],
      firstDataRow: 2,
      lastDataRow: Math.max(2, maxDataRow),
    };

  const secondaryCol =
    matchedCols[1] ||
    numericCols.find((c) => c.colIndex !== primaryCol.colIndex) ||
    columns.find((c) => c.colIndex !== primaryCol.colIndex) || {
      colIndex: primaryCol.colIndex + 1,
      colLetter: colIndexToLetter(primaryCol.colIndex + 1),
      header: `${colIndexToLetter(primaryCol.colIndex + 1)} Sütunu`,
      isNumeric: true,
      sampleValues: [],
      firstDataRow: 2,
      lastDataRow: Math.max(2, maxDataRow),
    };

  const targetCellRef = `${colIndexToLetter(activeCol)}${activeRow + 1}`;
  const rangeStr = `${primaryCol.colLetter}2:${primaryCol.colLetter}${primaryCol.lastDataRow}`;
  const nums = readNumericValuesInColumn(
    cellData,
    primaryCol.colIndex,
    2,
    primaryCol.lastDataRow
  );

  // Next empty column for bulk column operations
  const nextEmptyColIdx =
    columns.length > 0 ? Math.max(...columns.map((c) => c.colIndex)) + 1 : activeCol;

  // 1) DÜŞEYARA / XLOOKUP / VLOOKUP
  if (
    norm.includes('duseyara') ||
    norm.includes('xlookup') ||
    norm.includes('vlookup') ||
    (norm.includes('ara') && norm.includes('getir'))
  ) {
    const lookupCol = matchedCols[0] || columns[0] || primaryCol;
    const returnCol = matchedCols[1] || numericCols[0] || secondaryCol;
    const endR = Math.max(lookupCol.lastDataRow, returnCol.lastDataRow, 10);
    const formula = `=XLOOKUP(${lookupCol.colLetter}2, ${lookupCol.colLetter}2:${lookupCol.colLetter}${endR}, ${returnCol.colLetter}2:${returnCol.colLetter}${endR}, "Bulunamadı")`;
    const previewVal =
      cellData?.[1]?.[returnCol.colIndex]?.v ?? 'Örnek Eşleşen Değer';

    return {
      title: `Düşeyara (XLOOKUP) Formülü — ${lookupCol.header} → ${returnCol.header}`,
      explanation: `${lookupCol.colLetter} sütunundaki anahtar değeri arar ve aynı satırdaki ${returnCol.colLetter} (${returnCol.header}) değerini otomatik getirir.`,
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
    norm.includes('degilse')
  ) {
    const thresholdMatch = prompt.match(/(\d+(?:[.,]\d+)?)/);
    const threshold = thresholdMatch
      ? Number(thresholdMatch[1].replace('.', '').replace(',', '.'))
      : 10000;
    const isLess = norm.includes('kucukse') && !norm.includes('buyukse');
    const op = isLess ? '<' : '>';

    let trueLabel = 'Onaylı';
    let falseLabel = 'Beklemede';
    const quoteMatches = prompt.match(/['"]([^'"]+)['"]/g);
    if (quoteMatches && quoteMatches.length >= 2) {
      trueLabel = quoteMatches[0].replace(/['"]/g, '');
      falseLabel = quoteMatches[1].replace(/['"]/g, '');
    } else if (norm.includes('kar') && norm.includes('zarar')) {
      trueLabel = 'Kâr';
      falseLabel = 'Zarar';
    } else if (norm.includes('yuksek') && norm.includes('dusuk')) {
      trueLabel = 'Yüksek';
      falseLabel = 'Düşük';
    }

    const row1Based = Math.max(2, activeRow + 1);
    const formula = `=IF(${primaryCol.colLetter}${row1Based}${op}${threshold}, "${trueLabel}", "${falseLabel}")`;

    const writes: CopilotRowWrite[] = [];
    for (let r = 1; r < primaryCol.lastDataRow; r++) {
      const val = readCellNumber(cellData, r, primaryCol.colIndex);
      const cond = isLess ? val < threshold : val > threshold;
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=IF(${primaryCol.colLetter}${r + 1}${op}${threshold}, "${trueLabel}", "${falseLabel}")`,
        computedValue: cond ? trueLabel : falseLabel,
      });
    }

    const sampleVal = readCellNumber(cellData, row1Based - 1, primaryCol.colIndex);
    const preview = (isLess ? sampleVal < threshold : sampleVal > threshold)
      ? trueLabel
      : falseLabel;

    return {
      title: `Koşullu Karar (EĞER / IF) — ${primaryCol.header}`,
      explanation: `${primaryCol.colLetter} sütunundaki değer ${threshold.toLocaleString('tr-TR')} sınırından ${isLess ? 'küçükse' : 'büyükse'} "${trueLabel}", aksi halde "${falseLabel}" yazar.`,
      formula,
      computedPreview: preview,
      targetCellRef,
      canApplyToColumn: writes.length > 0,
      columnHeaderSuggestion: `${primaryCol.header} Durumu`,
      columnWrites: writes,
      category: 'if',
    };
  }

  // 3) KDV / YÜZDE / İNDİRİM / ARTIŞ HESAPLAMA
  if (
    norm.includes('kdv') ||
    norm.includes('yuzde') ||
    norm.includes('%') ||
    norm.includes('indirim') ||
    norm.includes('zam') ||
    norm.includes('artis')
  ) {
    const pctMatch = prompt.match(/(?:%|yüzde|yuzde)\s*(\d+(?:[.,]\d+)?)|(\d+(?:[.,]\d+)?)\s*%/i);
    const pct = pctMatch ? Number((pctMatch[1] || pctMatch[2]).replace(',', '.')) : 20;
    const isAdd =
      norm.includes('dahil') || norm.includes('ekle') || norm.includes('zam') || norm.includes('artis');
    const isDiscount = norm.includes('indirim') || norm.includes('dus');

    const multiplier = isAdd
      ? Number((1 + pct / 100).toFixed(4))
      : isDiscount
        ? Number((1 - pct / 100).toFixed(4))
        : Number((pct / 100).toFixed(4));

    const row1Based = Math.max(2, activeRow + 1);
    const formula = `=${primaryCol.colLetter}${row1Based}*${multiplier}`;

    const writes: CopilotRowWrite[] = [];
    for (let r = 1; r < primaryCol.lastDataRow; r++) {
      const baseVal = readCellNumber(cellData, r, primaryCol.colIndex);
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=${primaryCol.colLetter}${r + 1}*${multiplier}`,
        computedValue: Math.round(baseVal * multiplier * 100) / 100,
      });
    }

    const firstVal = readCellNumber(cellData, row1Based - 1, primaryCol.colIndex);
    return {
      title: `%${pct} ${isAdd ? 'Artışlı / KDV Dahil' : isDiscount ? 'İndirimli' : 'Oran / KDV'} Hesaplama`,
      explanation: `${primaryCol.header} (${primaryCol.colLetter}) değerini ${multiplier} katsayısı ile çarpar. İsterseniz tek tıkla tüm satırlara yeni sütun olarak uygulayabilirsiniz.`,
      formula,
      computedPreview: Math.round(firstVal * multiplier * 100) / 100,
      targetCellRef,
      canApplyToColumn: writes.length > 0,
      columnHeaderSuggestion: isAdd
        ? `${primaryCol.header} (+%${pct})`
        : isDiscount
          ? `${primaryCol.header} (-%${pct})`
          : `%${pct} KDV / Oran`,
      columnWrites: writes,
      category: 'math',
    };
  }

  // 4) İKİ SÜTUN ARİTMETİĞİ (ÇARP, ÇIKAR, BÖL, SATIR BAZLI TOPLA)
  if (
    norm.includes('carp') ||
    norm.includes('cikar') ||
    norm.includes('fark') ||
    norm.includes('bol') ||
    (matchedCols.length >= 2 && norm.includes('tum satir'))
  ) {
    const op = norm.includes('cikar') || norm.includes('fark')
      ? '-'
      : norm.includes('bol')
        ? '/'
        : norm.includes('topla')
          ? '+'
          : '*';
    const opLabel =
      op === '*' ? 'Çarpımı' : op === '-' ? 'Farkı' : op === '/' ? 'Oranı' : 'Toplamı';
    const row1Based = Math.max(2, activeRow + 1);
    const maxR = Math.max(primaryCol.lastDataRow, secondaryCol.lastDataRow);
    const formula = `=${primaryCol.colLetter}${row1Based}${op}${secondaryCol.colLetter}${row1Based}`;

    const writes: CopilotRowWrite[] = [];
    for (let r = 1; r < maxR; r++) {
      const a = readCellNumber(cellData, r, primaryCol.colIndex);
      const b = readCellNumber(cellData, r, secondaryCol.colIndex);
      const res =
        op === '*'
          ? a * b
          : op === '-'
            ? a - b
            : op === '/'
              ? b !== 0
                ? Math.round((a / b) * 100) / 100
                : 0
              : a + b;
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=${primaryCol.colLetter}${r + 1}${op}${secondaryCol.colLetter}${r + 1}`,
        computedValue: Math.round(res * 100) / 100,
      });
    }

    const sampleA = readCellNumber(cellData, row1Based - 1, primaryCol.colIndex);
    const sampleB = readCellNumber(cellData, row1Based - 1, secondaryCol.colIndex);
    const sampleRes =
      op === '*'
        ? sampleA * sampleB
        : op === '-'
          ? sampleA - sampleB
          : op === '/'
            ? sampleB !== 0
              ? Math.round((sampleA / sampleB) * 100) / 100
              : 0
            : sampleA + sampleB;

    return {
      title: `${primaryCol.header} (${primaryCol.colLetter}) ${op} ${secondaryCol.header} (${secondaryCol.colLetter}) ${opLabel}`,
      explanation: `Her satır için ${primaryCol.colLetter} ve ${secondaryCol.colLetter} hücrelerini (${op}) işlemine sokar.`,
      formula,
      computedPreview: sampleRes,
      targetCellRef,
      canApplyToColumn: writes.length > 0,
      columnHeaderSuggestion: `${primaryCol.header} ${op} ${secondaryCol.header}`,
      columnWrites: writes,
      category: 'math',
    };
  }

  // 5) ORTALAMA (AVERAGE)
  if (norm.includes('ortalama') || norm.includes('average') || norm.includes('ort')) {
    const sum = nums.reduce((a, b) => a + b, 0);
    const avg = nums.length > 0 ? Math.round((sum / nums.length) * 100) / 100 : 0;
    return {
      title: `${primaryCol.header} (${primaryCol.colLetter}) Sütunu Ortalaması`,
      explanation: `${rangeStr} aralığındaki ${nums.length} sayısal hücrenin aritmetik ortalamasını alır.`,
      formula: `=AVERAGE(${rangeStr})`,
      computedPreview: avg,
      targetCellRef,
      canApplyToColumn: false,
      category: 'average',
    };
  }

  // 6) EN BÜYÜK / MAKSİMUM (MAX) veya EN KÜÇÜK / MİNİMUM (MIN)
  if (
    norm.includes('en buyuk') ||
    norm.includes('en yuksek') ||
    norm.includes('maksimum') ||
    norm.includes('max')
  ) {
    const maxVal = nums.length > 0 ? Math.max(...nums) : 0;
    return {
      title: `${primaryCol.header} (${primaryCol.colLetter}) En Yüksek Değer (MAX)`,
      explanation: `${rangeStr} aralığındaki en büyük sayısal değeri bulur.`,
      formula: `=MAX(${rangeStr})`,
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
    const minVal = nums.length > 0 ? Math.min(...nums) : 0;
    return {
      title: `${primaryCol.header} (${primaryCol.colLetter}) En Düşük Değer (MIN)`,
      explanation: `${rangeStr} aralığındaki en küçük sayısal değeri bulur.`,
      formula: `=MIN(${rangeStr})`,
      computedPreview: minVal,
      targetCellRef,
      canApplyToColumn: false,
      category: 'minmax',
    };
  }

  // 7) ADET / SAYI (COUNTA / COUNT)
  if (
    norm.includes('kac adet') ||
    norm.includes('kac tane') ||
    norm.includes('sayisi') ||
    norm.includes('dolu hucre') ||
    norm.includes('count')
  ) {
    let nonEmpty = 0;
    for (let r = 1; r < primaryCol.lastDataRow; r++) {
      const v = cellData?.[r]?.[primaryCol.colIndex]?.v;
      if (v !== undefined && v !== null && String(v).trim() !== '') nonEmpty++;
    }
    return {
      title: `${primaryCol.header} (${primaryCol.colLetter}) Kayıt Sayısı (COUNTA)`,
      explanation: `${rangeStr} aralığındaki dolu satırların toplam adedini sayar.`,
      formula: `=COUNTA(${rangeStr})`,
      computedPreview: nonEmpty,
      targetCellRef,
      canApplyToColumn: false,
      category: 'count',
    };
  }

  // 8) METİN BİRLEŞTİRME (CONCAT)
  if (norm.includes('birlestir') || norm.includes('concat')) {
    const row1Based = Math.max(2, activeRow + 1);
    const formula = `=CONCAT(${primaryCol.colLetter}${row1Based}, " - ", ${secondaryCol.colLetter}${row1Based})`;
    const writes: CopilotRowWrite[] = [];
    const maxR = Math.max(primaryCol.lastDataRow, secondaryCol.lastDataRow);
    for (let r = 1; r < maxR; r++) {
      const a = String(cellData?.[r]?.[primaryCol.colIndex]?.v ?? '').trim();
      const b = String(cellData?.[r]?.[secondaryCol.colIndex]?.v ?? '').trim();
      writes.push({
        row: r,
        col: nextEmptyColIdx,
        formula: `=CONCAT(${primaryCol.colLetter}${r + 1}, " - ", ${secondaryCol.colLetter}${r + 1})`,
        computedValue: `${a} - ${b}`,
      });
    }
    const pA = String(cellData?.[row1Based - 1]?.[primaryCol.colIndex]?.v ?? '');
    const pB = String(cellData?.[row1Based - 1]?.[secondaryCol.colIndex]?.v ?? '');
    return {
      title: `${primaryCol.header} ve ${secondaryCol.header} Metin Birleştirme`,
      explanation: `${primaryCol.colLetter} ve ${secondaryCol.colLetter} sütunlarındaki değerleri arasına tire koyarak birleştirir.`,
      formula,
      computedPreview: `${pA} - ${pB}`,
      targetCellRef,
      canApplyToColumn: writes.length > 0,
      columnHeaderSuggestion: `${primaryCol.header} & ${secondaryCol.header}`,
      columnWrites: writes,
      category: 'text',
    };
  }

  // 9) VARSAYILAN: TOPLAM (SUM)
  const totalSum = Math.round(nums.reduce((a, b) => a + b, 0) * 100) / 100;
  return {
    title: `${primaryCol.header} (${primaryCol.colLetter}) Genel Toplamı (SUM)`,
    explanation: `${rangeStr} aralığındaki tüm sayısal değerleri toplar (${nums.length} satır).`,
    formula: `=SUM(${rangeStr})`,
    computedPreview: totalSum,
    targetCellRef,
    canApplyToColumn: false,
    category: 'sum',
  };
}
