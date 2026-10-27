import { UniverWorkbookData } from './excelZeroUpload';
import { colIndexToLetter } from './collaboration';

export interface FormulaArgDef {
  name: string;
  labelTr: string;
  descriptionTr: string;
  placeholder: string;
  required: boolean;
  isRange?: boolean;
  options?: { value: string; label: string }[];
}

export interface ExcelFormulaDef {
  name: string; // İngilizce Excel formül adı (Örn: XLOOKUP, VLOOKUP, SUM, IF)
  category: 'lookup' | 'math' | 'logical' | 'stats' | 'text';
  categoryLabelTr: string;
  shortDescTr: string;
  syntax: string;
  example: string;
  args: FormulaArgDef[];
}

export const EXCEL_FORMULA_CATALOG: ExcelFormulaDef[] = [
  {
    name: 'XLOOKUP',
    category: 'lookup',
    categoryLabelTr: 'Arama ve Başvuru',
    shortDescTr:
      'Bir sütunda veya aralıkta bir değeri arar ve başka bir sütundaki karşılık gelen değeri döndürür (DÜŞEYARA / VLOOKUP’ın modern ve kolay hali).',
    syntax: '=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])',
    example: '=XLOOKUP(A2, A2:A50, G2:G50, "Bulunamadı", 0)',
    args: [
      {
        name: 'lookup_value',
        labelTr: 'Aranan Değer (lookup_value)',
        descriptionTr: 'Tabloda aramak istediğiniz hücre veya değer (Örn: A2 veya "PRJ-2026-01")',
        placeholder: 'A2 veya "PRJ-2026-01"',
        required: true,
      },
      {
        name: 'lookup_array',
        labelTr: 'Arama Yapılacak Sütun / Aralık (lookup_array)',
        descriptionTr: 'Aranan değerin bulunduğu sütun aralığı (Örn: A2:A7)',
        placeholder: 'A2:A7',
        required: true,
        isRange: true,
      },
      {
        name: 'return_array',
        labelTr: 'Getirilecek Sonuç Sütunu / Aralık (return_array)',
        descriptionTr: 'Eşleşme bulunduğunda değerin alınacağı sütun aralığı (Örn: G2:G7)',
        placeholder: 'G2:G7',
        required: true,
        isRange: true,
      },
      {
        name: 'if_not_found',
        labelTr: 'Bulunamazsa Yazılacak Değer ([if_not_found])',
        descriptionTr: 'Aranan değer tabloda yoksa gösterilecek metin (Örn: "Bulunamadı")',
        placeholder: '"Bulunamadı"',
        required: false,
      },
      {
        name: 'match_mode',
        labelTr: 'Eşleşme Modu ([match_mode])',
        descriptionTr: '0: Tam eşleşme (önerilen), -1: Tam veya bir küçük, 1: Tam veya bir büyük',
        placeholder: '0',
        required: false,
        options: [
          { value: '0', label: '0 - Tam Eşleşme (Varsayılan)' },
          { value: '-1', label: '-1 - Tam Eşleşme veya Bir Küçük Öğe' },
          { value: '1', label: '1 - Tam Eşleşme veya Bir Büyük Öğe' },
          { value: '2', label: '2 - Joker Karakter Eşleşmesi (*, ?)' },
        ],
      },
    ],
  },
  {
    name: 'VLOOKUP',
    category: 'lookup',
    categoryLabelTr: 'Arama ve Başvuru',
    shortDescTr:
      'Tablonun ilk sütununda bir değer arar ve aynı satırdaki belirtilen sütun numarasından değeri getirir (Klasik DÜŞEYARA).',
    syntax: '=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    example: '=VLOOKUP(A2, A2:H7, 7, FALSE)',
    args: [
      {
        name: 'lookup_value',
        labelTr: 'Aranan Değer (lookup_value)',
        descriptionTr: 'Tablonun ilk sütununda aranacak değer veya hücre (Örn: A2)',
        placeholder: 'A2',
        required: true,
      },
      {
        name: 'table_array',
        labelTr: 'Tablo Aralığı (table_array)',
        descriptionTr: 'Aramanın yapılacağı tüm tablo aralığı (Örn: A2:H7)',
        placeholder: 'A2:H7',
        required: true,
        isRange: true,
      },
      {
        name: 'col_index_num',
        labelTr: 'Sütun İndis Sayısı (col_index_num)',
        descriptionTr: 'Tablo aralığındaki kaçıncı sütunun getirileceği (1 = ilk sütun, 2 = ikinci sütun...)',
        placeholder: '7',
        required: true,
      },
      {
        name: 'range_lookup',
        labelTr: 'Eşleşme Türü ([range_lookup])',
        descriptionTr: 'FALSE (0): Tam eşleşme, TRUE (1): Yaklaşık eşleşme',
        placeholder: 'FALSE',
        required: false,
        options: [
          { value: 'FALSE', label: 'FALSE (0) - Tam Eşleşme (Önerilen)' },
          { value: 'TRUE', label: 'TRUE (1) - Yaklaşık Eşleşme' },
        ],
      },
    ],
  },
  {
    name: 'SUM',
    category: 'math',
    categoryLabelTr: 'Matematik ve Toplama',
    shortDescTr: 'Seçilen hücre aralığındaki tüm sayıları toplar (TOPLA).',
    syntax: '=SUM(number1, [number2], ...)',
    example: '=SUM(D2:F2)',
    args: [
      {
        name: 'range1',
        labelTr: 'Toplanacak Hücre Aralığı (number1)',
        descriptionTr: 'Toplanmasını istediğiniz hücre aralığı (Örn: D2:D6 veya D2:F2)',
        placeholder: 'D2:D6',
        required: true,
        isRange: true,
      },
      {
        name: 'range2',
        labelTr: 'Ek Hücre / Aralık ([number2])',
        descriptionTr: 'İsteğe bağlı ikinci hücre aralığı veya sayı',
        placeholder: '',
        required: false,
        isRange: true,
      },
    ],
  },
  {
    name: 'SUMIF',
    category: 'math',
    categoryLabelTr: 'Matematik ve Toplama',
    shortDescTr: 'Belirli bir koşula (ölçüte) uyan hücreleri toplar (ETOPLA).',
    syntax: '=SUMIF(range, criteria, [sum_range])',
    example: '=SUMIF(H2:H6, "Aktif Onaylı", G2:G6)',
    args: [
      {
        name: 'range',
        labelTr: 'Koşulun Aranacağı Aralık (range)',
        descriptionTr: 'Koşulun kontrol edileceği sütun aralığı (Örn: H2:H6)',
        placeholder: 'H2:H6',
        required: true,
        isRange: true,
      },
      {
        name: 'criteria',
        labelTr: 'Koşul / Ölçüt (criteria)',
        descriptionTr: 'Hangi değerlerin toplanacağını belirleyen ölçüt (Örn: "Aktif Onaylı" veya ">400000")',
        placeholder: '"Aktif Onaylı"',
        required: true,
      },
      {
        name: 'sum_range',
        labelTr: 'Toplanacak Sütun Aralığı ([sum_range])',
        descriptionTr: 'Koşul sağlandığında toplanacak sayısal sütun (Örn: G2:G6)',
        placeholder: 'G2:G6',
        required: false,
        isRange: true,
      },
    ],
  },
  {
    name: 'AVERAGE',
    category: 'stats',
    categoryLabelTr: 'İstatistik',
    shortDescTr: 'Seçilen hücre aralığındaki sayıların aritmetik ortalamasını alır (ORTALAMA).',
    syntax: '=AVERAGE(number1, [number2], ...)',
    example: '=AVERAGE(D2:D6)',
    args: [
      {
        name: 'range1',
        labelTr: 'Ortalaması Alınacak Aralık',
        descriptionTr: 'Sayısal hücre aralığı (Örn: D2:D6)',
        placeholder: 'D2:D6',
        required: true,
        isRange: true,
      },
    ],
  },
  {
    name: 'IF',
    category: 'logical',
    categoryLabelTr: 'Mantıksal',
    shortDescTr:
      'Bir koşulun doğru veya yanlış olmasına göre iki farklı sonuçtan birini döndürür (EĞER).',
    syntax: '=IF(logical_test, value_if_true, value_if_false)',
    example: '=IF(G2>1500000, "Yüksek Bütçe", "Normal Bütçe")',
    args: [
      {
        name: 'logical_test',
        labelTr: 'Mantıksal Koşul (logical_test)',
        descriptionTr: 'Test edilecek karşılaştırma (Örn: G2>1500000 veya D2>=500000)',
        placeholder: 'G2>1500000',
        required: true,
      },
      {
        name: 'value_if_true',
        labelTr: 'Koşul Doğruysa (value_if_true)',
        descriptionTr: 'Koşul sağlandığında hücreye yazılacak değer (Örn: "Yüksek Bütçe")',
        placeholder: '"Yüksek Bütçe"',
        required: true,
      },
      {
        name: 'value_if_false',
        labelTr: 'Koşul Yanlışsa (value_if_false)',
        descriptionTr: 'Koşul sağlanmadığında hücreye yazılacak değer (Örn: "Normal Bütçe")',
        placeholder: '"Normal Bütçe"',
        required: true,
      },
    ],
  },
  {
    name: 'COUNTIF',
    category: 'stats',
    categoryLabelTr: 'İstatistik',
    shortDescTr: 'Belirli bir koşulu sağlayan hücrelerin sayısını verir (EĞERSAY).',
    syntax: '=COUNTIF(range, criteria)',
    example: '=COUNTIF(H2:H6, "Aktif Onaylı")',
    args: [
      {
        name: 'range',
        labelTr: 'Sayılacak Hücre Aralığı (range)',
        descriptionTr: 'Kontrol edilecek hücre aralığı (Örn: H2:H6)',
        placeholder: 'H2:H6',
        required: true,
        isRange: true,
      },
      {
        name: 'criteria',
        labelTr: 'Ölçüt (criteria)',
        descriptionTr: 'Sayılacak değer veya koşul (Örn: "Aktif Onaylı" veya ">500000")',
        placeholder: '"Aktif Onaylı"',
        required: true,
      },
    ],
  },
  {
    name: 'MAX',
    category: 'stats',
    categoryLabelTr: 'İstatistik',
    shortDescTr: 'Seçilen aralıktaki en büyük (maksimum) sayıyı döndürür (MAK).',
    syntax: '=MAX(number1, [number2], ...)',
    example: '=MAX(G2:G6)',
    args: [
      {
        name: 'range1',
        labelTr: 'Hücre Aralığı',
        descriptionTr: 'En büyük değerin aranacağı aralık (Örn: G2:G6)',
        placeholder: 'G2:G6',
        required: true,
        isRange: true,
      },
    ],
  },
  {
    name: 'MIN',
    category: 'stats',
    categoryLabelTr: 'İstatistik',
    shortDescTr: 'Seçilen aralıktaki en küçük (minimum) sayıyı döndürür (MİN).',
    syntax: '=MIN(number1, [number2], ...)',
    example: '=MIN(G2:G6)',
    args: [
      {
        name: 'range1',
        labelTr: 'Hücre Aralığı',
        descriptionTr: 'En küçük değerin aranacağı aralık (Örn: G2:G6)',
        placeholder: 'G2:G6',
        required: true,
        isRange: true,
      },
    ],
  },
  {
    name: 'INDEX',
    category: 'lookup',
    categoryLabelTr: 'Arama ve Başvuru',
    shortDescTr: 'Bir tablo veya aralıktaki satır ve sütun kesişimindeki hücre değerini döndürür (İNDİS).',
    syntax: '=INDEX(array, row_num, [column_num])',
    example: '=INDEX(A2:H6, 2, 7)',
    args: [
      {
        name: 'array',
        labelTr: 'Tablo Aralığı (array)',
        descriptionTr: 'Hücre aralığı (Örn: A2:H6)',
        placeholder: 'A2:H6',
        required: true,
        isRange: true,
      },
      {
        name: 'row_num',
        labelTr: 'Satır Numarası (row_num)',
        descriptionTr: 'Aralık içindeki kaçıncı satır (Örn: 2)',
        placeholder: '2',
        required: true,
      },
      {
        name: 'column_num',
        labelTr: 'Sütun Numarası ([column_num])',
        descriptionTr: 'Aralık içindeki kaçıncı sütun (Örn: 7)',
        placeholder: '7',
        required: false,
      },
    ],
  },
  {
    name: 'MATCH',
    category: 'lookup',
    categoryLabelTr: 'Arama ve Başvuru',
    shortDescTr: 'Aranan bir öğenin aralık içindeki göreli sıra numarasını döndürür (KAÇINCI).',
    syntax: '=MATCH(lookup_value, lookup_array, [match_type])',
    example: '=MATCH("PRJ-2026-02", A2:A6, 0)',
    args: [
      {
        name: 'lookup_value',
        labelTr: 'Aranan Değer (lookup_value)',
        descriptionTr: 'Sırası bulunacak değer veya hücre (Örn: "PRJ-2026-02" veya A3)',
        placeholder: '"PRJ-2026-02"',
        required: true,
      },
      {
        name: 'lookup_array',
        labelTr: 'Aranan Aralık (lookup_array)',
        descriptionTr: 'Tek satır veya tek sütunluk aralık (Örn: A2:A6)',
        placeholder: 'A2:A6',
        required: true,
        isRange: true,
      },
      {
        name: 'match_type',
        labelTr: 'Eşleşme Türü ([match_type])',
        descriptionTr: '0: Tam eşleşme',
        placeholder: '0',
        required: false,
      },
    ],
  },
  {
    name: 'IFERROR',
    category: 'logical',
    categoryLabelTr: 'Mantıksal',
    shortDescTr: 'Formül hata (#N/A, #VALUE! vb.) üretirse alternatif bir değer döndürür (EĞERHATA).',
    syntax: '=IFERROR(value, value_if_error)',
    example: '=IFERROR(D2/E2, 0)',
    args: [
      {
        name: 'value',
        labelTr: 'Kontrol Edilecek Formül/Değer (value)',
        descriptionTr: 'Hesaplanacak ifade (Örn: D2/E2)',
        placeholder: 'D2/E2',
        required: true,
      },
      {
        name: 'value_if_error',
        labelTr: 'Hata Varsa Gösterilecek Değer (value_if_error)',
        descriptionTr: 'Hata durumunda yazılacak metin veya sayı (Örn: 0 veya "Hata Yok")',
        placeholder: '0',
        required: true,
      },
    ],
  },
  {
    name: 'ROUND',
    category: 'math',
    categoryLabelTr: 'Matematik ve Toplama',
    shortDescTr: 'Bir sayıyı belirtilen basamak sayısına yuvarlar (YUVARLA).',
    syntax: '=ROUND(number, num_digits)',
    example: '=ROUND(G2/12, 2)',
    args: [
      {
        name: 'number',
        labelTr: 'Sayı veya Hücre (number)',
        descriptionTr: 'Yuvarlanacak sayı veya hücre (Örn: G2/12)',
        placeholder: 'G2/12',
        required: true,
      },
      {
        name: 'num_digits',
        labelTr: 'Ondalık Basamak Sayısı (num_digits)',
        descriptionTr: 'Virgülden sonra kaç basamak kalacağı (Örn: 2 veya 0)',
        placeholder: '2',
        required: true,
      },
    ],
  },
  {
    name: 'CONCAT',
    category: 'text',
    categoryLabelTr: 'Metin',
    shortDescTr: 'Birden fazla hücredeki metinleri tek bir hücrede birleştirir (BİRLEŞTİR).',
    syntax: '=CONCAT(text1, [text2], ...)',
    example: '=CONCAT(A2, " - ", B2)',
    args: [
      {
        name: 'text1',
        labelTr: 'Birinci Metin / Hücre (text1)',
        descriptionTr: 'İlk metin veya hücre (Örn: A2)',
        placeholder: 'A2',
        required: true,
      },
      {
        name: 'text2',
        labelTr: 'İkinci Metin / Hücre (text2)',
        descriptionTr: 'Araya eklenecek ayraç veya ikinci hücre (Örn: " - ")',
        placeholder: '" - "',
        required: false,
      },
      {
        name: 'text3',
        labelTr: 'Üçüncü Metin / Hücre (text3)',
        descriptionTr: 'Üçüncü metin veya hücre (Örn: B2)',
        placeholder: 'B2',
        required: false,
      },
    ],
  },
];

const RECENT_FORMULAS_KEY = 'syncsheet_recent_formulas_v1';

export function getRecentFormulaNames(): string[] {
  if (typeof window === 'undefined') {
    return ['XLOOKUP', 'SUM', 'IF', 'VLOOKUP', 'AVERAGE', 'COUNTIF'];
  }
  try {
    const raw = localStorage.getItem(RECENT_FORMULAS_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch (e) {
    // ignore
  }
  return ['XLOOKUP', 'SUM', 'IF', 'VLOOKUP', 'AVERAGE', 'COUNTIF'];
}

export function recordFormulaUsage(formulaNameOrExpression: string): string[] {
  const cleaned = formulaNameOrExpression
    .trim()
    .replace(/^=/, '')
    .split('(')[0]
    .trim()
    .toUpperCase();

  if (!cleaned) return getRecentFormulaNames();

  const known = EXCEL_FORMULA_CATALOG.some((f) => f.name === cleaned);
  if (!known && !/^[A-Z0-9_.]+$/.test(cleaned)) {
    return getRecentFormulaNames();
  }

  const current = getRecentFormulaNames().filter((item) => item !== cleaned);
  const next = [cleaned, ...current].slice(0, 10);

  if (typeof window !== 'undefined') {
    try {
      localStorage.setItem(RECENT_FORMULAS_KEY, JSON.stringify(next));
    } catch (e) {
      // ignore
    }
  }
  return next;
}

/**
 * Aktif sayfadaki sütun başlıklarını ve önerilen veri aralıklarını döndürür
 * (Örn: A Sütunu: PROJE KODU -> A2:A7)
 */
export function getSheetColumnRanges(
  workbookData: UniverWorkbookData,
  activeSheetId?: string
): Array<{
  colIndex: number;
  colLetter: string;
  headerName: string;
  rangeRef: string;
  sampleCellRef: string;
  sampleValue: any;
}> {
  const sid =
    activeSheetId && workbookData.sheets?.[activeSheetId]
      ? activeSheetId
      : workbookData.sheetOrder?.[0] || Object.keys(workbookData.sheets || {})[0];
  const sheet = workbookData.sheets?.[sid];
  if (!sheet || !sheet.cellData) return [];

  const cellData = sheet.cellData;
  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((n) => !isNaN(n))
    .sort((a, b) => a - b);

  // Son dolu veri satırını bulalım ("GENEL TOPLAM" satırı varsa onu hariç tutabiliriz)
  let maxDataRow = 1;
  for (const r of rowKeys) {
    if (r === 0) continue;
    const rowObj = cellData[r];
    if (!rowObj) continue;
    const firstCellVal = String(rowObj[0]?.v || '').toUpperCase();
    if (firstCellVal.includes('GENEL TOPLAM') || firstCellVal.includes('TOPLAM')) {
      continue;
    }
    if (Object.keys(rowObj).length > 0) {
      maxDataRow = Math.max(maxDataRow, r);
    }
  }

  const headerRow = cellData[0] || {};
  const colIndices = Object.keys(headerRow)
    .map(Number)
    .filter((n) => !isNaN(n))
    .sort((a, b) => a - b);

  return colIndices.map((cIdx) => {
    const letter = colIndexToLetter(cIdx);
    const headerName = String(headerRow[cIdx]?.v || `${letter} Sütunu`);
    const sampleVal = cellData[1]?.[cIdx]?.v ?? '';
    return {
      colIndex: cIdx,
      colLetter: letter,
      headerName,
      rangeRef: `${letter}2:${letter}${maxDataRow + 1}`,
      sampleCellRef: `${letter}2`,
      sampleValue: sampleVal,
    };
  });
}

function parseCellCoord(ref: string): { r: number; c: number } | null {
  const m = ref.trim().toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  const letters = m[1];
  const rowNum = parseInt(m[2], 10) - 1;
  let colNum = 0;
  for (let i = 0; i < letters.length; i++) {
    colNum = colNum * 26 + (letters.charCodeAt(i) - 64);
  }
  return { r: rowNum, c: colNum - 1 };
}

function resolveSheetCellData(
  workbookData: UniverWorkbookData | undefined,
  defaultCellData: Record<number, Record<number, any>>,
  rawRangeOrRef: string
): { targetCellData: Record<number, Record<number, any>>; cleanRef: string } {
  const trimmed = rawRangeOrRef.trim();
  const bangIdx = trimmed.indexOf('!');
  if (bangIdx === -1 || !workbookData?.sheets) {
    return { targetCellData: defaultCellData, cleanRef: trimmed };
  }

  const rawSheetName = trimmed
    .slice(0, bangIdx)
    .trim()
    .replace(/^'+|'+$/g, '')
    .replace(/^\[.*?\]/, ''); // [Kitap1.xlsx]Sayfa1!A2:B10 formatını da destekler
  const cleanRef = trimmed.slice(bangIdx + 1).trim();

  for (const [sid, sheet] of Object.entries(workbookData.sheets)) {
    if (
      sid.toLowerCase() === rawSheetName.toLowerCase() ||
      (sheet.name && sheet.name.toLowerCase() === rawSheetName.toLowerCase())
    ) {
      return { targetCellData: sheet.cellData || {}, cleanRef };
    }
  }

  return { targetCellData: defaultCellData, cleanRef };
}

function getRangeValues(
  cellData: Record<number, Record<number, any>>,
  rangeStr: string,
  workbookData?: UniverWorkbookData
): any[][] {
  const { targetCellData, cleanRef } = resolveSheetCellData(workbookData, cellData, rangeStr);
  const parts = cleanRef.trim().split(':');
  if (parts.length === 1) {
    const coord = parseCellCoord(parts[0]);
    if (!coord) return [[cleanRef.replace(/^"|"$/g, '')]];
    return [[targetCellData[coord.r]?.[coord.c]?.v]];
  }
  const start = parseCellCoord(parts[0]);
  const end = parseCellCoord(parts[1]);
  if (!start || !end) return [];

  const r1 = Math.min(start.r, end.r);
  const r2 = Math.max(start.r, end.r);
  const c1 = Math.min(start.c, end.c);
  const c2 = Math.max(start.c, end.c);

  const matrix: any[][] = [];
  for (let r = r1; r <= r2; r++) {
    const rowArr: any[] = [];
    for (let c = c1; c <= c2; c++) {
      rowArr.push(targetCellData[r]?.[c]?.v);
    }
    matrix.push(rowArr);
  }
  return matrix;
}

function resolveSingleArgValue(
  cellData: Record<number, Record<number, any>>,
  rawArg: string,
  workbookData?: UniverWorkbookData
): any {
  const { targetCellData, cleanRef } = resolveSheetCellData(workbookData, cellData, rawArg);
  const trimmed = cleanRef.trim();
  if (!trimmed) return '';
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1);
  }
  const coord = parseCellCoord(trimmed);
  if (coord) {
    return targetCellData[coord.r]?.[coord.c]?.v;
  }
  const num = Number(trimmed);
  if (!isNaN(num)) return num;
  return trimmed;
}

/**
 * Kullanıcının fx sihirbazında oluşturduğu formülü (XLOOKUP, VLOOKUP, SUM, IF, vb.)
 * aktif tablo verisi üzerinde canlı olarak hesaplar ve önizleme/değer döndürür.
 */
export function evaluateFormulaPreview(
  formulaName: string,
  argValues: Record<string, string>,
  workbookData: UniverWorkbookData,
  activeSheetId?: string
): { formulaString: string; computedValue: any } {
  const def = EXCEL_FORMULA_CATALOG.find((f) => f.name === formulaName);
  if (!def) {
    return { formulaString: `=${formulaName}()`, computedValue: '' };
  }

  const orderedArgs: string[] = [];
  def.args.forEach((argDef) => {
    const val = (argValues[argDef.name] || '').trim();
    if (val !== '') {
      orderedArgs.push(val);
    }
  });

  const formulaString = `=${formulaName}(${orderedArgs.join(', ')})`;

  const sid =
    activeSheetId && workbookData.sheets?.[activeSheetId]
      ? activeSheetId
      : workbookData.sheetOrder?.[0] || Object.keys(workbookData.sheets || {})[0];
  const cellData = workbookData.sheets?.[sid]?.cellData || {};

  try {
    if (formulaName === 'XLOOKUP') {
      const lookupVal = resolveSingleArgValue(cellData, argValues.lookup_value || '', workbookData);
      const lookupMat = getRangeValues(cellData, argValues.lookup_array || '', workbookData).flat();
      const returnMat = getRangeValues(cellData, argValues.return_array || '', workbookData).flat();
      const notFoundVal = argValues.if_not_found
        ? resolveSingleArgValue(cellData, argValues.if_not_found, workbookData)
        : '#N/A';

      const idx = lookupMat.findIndex(
        (item) =>
          item !== undefined &&
          String(item).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()
      );
      if (idx !== -1 && idx < returnMat.length) {
        return { formulaString, computedValue: returnMat[idx] };
      }
      return { formulaString, computedValue: notFoundVal };
    }

    if (formulaName === 'VLOOKUP') {
      const lookupVal = resolveSingleArgValue(cellData, argValues.lookup_value || '', workbookData);
      const tableMat = getRangeValues(cellData, argValues.table_array || '', workbookData);
      const colIdx = Math.max(1, Number(argValues.col_index_num || 1)) - 1;

      for (const row of tableMat) {
        if (
          row[0] !== undefined &&
          String(row[0]).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()
        ) {
          return { formulaString, computedValue: row[colIdx] ?? '#REF!' };
        }
      }
      return { formulaString, computedValue: '#N/A' };
    }

    if (formulaName === 'SUM') {
      const vals1 = getRangeValues(cellData, argValues.range1 || '', workbookData).flat();
      const vals2 = argValues.range2
        ? getRangeValues(cellData, argValues.range2, workbookData).flat()
        : [];
      const sum = [...vals1, ...vals2].reduce((acc, v) => {
        const n = Number(v);
        return !isNaN(n) ? acc + n : acc;
      }, 0);
      return { formulaString, computedValue: sum };
    }

    if (formulaName === 'AVERAGE') {
      const vals = getRangeValues(cellData, argValues.range1 || '', workbookData)
        .flat()
        .map(Number)
        .filter((n) => !isNaN(n));
      const avg =
        vals.length > 0
          ? Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 100) / 100
          : 0;
      return { formulaString, computedValue: avg };
    }

    if (formulaName === 'MAX') {
      const vals = getRangeValues(cellData, argValues.range1 || '', workbookData)
        .flat()
        .map(Number)
        .filter((n) => !isNaN(n));
      return { formulaString, computedValue: vals.length > 0 ? Math.max(...vals) : 0 };
    }

    if (formulaName === 'MIN') {
      const vals = getRangeValues(cellData, argValues.range1 || '', workbookData)
        .flat()
        .map(Number)
        .filter((n) => !isNaN(n));
      return { formulaString, computedValue: vals.length > 0 ? Math.min(...vals) : 0 };
    }

    if (formulaName === 'COUNTIF') {
      const vals = getRangeValues(cellData, argValues.range || '', workbookData).flat();
      const crit = String(
        resolveSingleArgValue(cellData, argValues.criteria || '', workbookData)
      ).toLowerCase();
      const count = vals.filter((v) => String(v ?? '').toLowerCase() === crit).length;
      return { formulaString, computedValue: count };
    }

    if (formulaName === 'SUMIF') {
      const checkVals = getRangeValues(cellData, argValues.range || '', workbookData).flat();
      const sumVals = argValues.sum_range
        ? getRangeValues(cellData, argValues.sum_range, workbookData).flat()
        : checkVals;
      const crit = String(
        resolveSingleArgValue(cellData, argValues.criteria || '', workbookData)
      ).toLowerCase();
      let total = 0;
      checkVals.forEach((v, i) => {
        if (String(v ?? '').toLowerCase() === crit) {
          const n = Number(sumVals[i]);
          if (!isNaN(n)) total += n;
        }
      });
      return { formulaString, computedValue: total };
    }

    if (formulaName === 'IF') {
      const testStr = (argValues.logical_test || '').trim();
      const valTrue = resolveSingleArgValue(cellData, argValues.value_if_true || '', workbookData);
      const valFalse = resolveSingleArgValue(
        cellData,
        argValues.value_if_false || '',
        workbookData
      );
      const match = testStr.match(/^([A-Z]+\d+)\s*(>=|<=|>|<|=)\s*(.+)$/i);
      if (match) {
        const left = resolveSingleArgValue(cellData, match[1], workbookData);
        const op = match[2];
        const right = resolveSingleArgValue(cellData, match[3], workbookData);
        let passed = false;
        if (op === '>') passed = Number(left) > Number(right);
        else if (op === '<') passed = Number(left) < Number(right);
        else if (op === '>=') passed = Number(left) >= Number(right);
        else if (op === '<=') passed = Number(left) <= Number(right);
        else if (op === '=') passed = String(left) === String(right);
        return { formulaString, computedValue: passed ? valTrue : valFalse };
      }
      return { formulaString, computedValue: valTrue };
    }
  } catch (e) {
    // ignore preview errors
  }

  return { formulaString, computedValue: '' };
}
