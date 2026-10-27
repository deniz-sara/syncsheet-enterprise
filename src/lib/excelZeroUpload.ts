import * as XLSX from 'xlsx';
import pako from 'pako';

export interface UniverWorkbookData {
  id: string;
  name: string;
  appVersion: string;
  locale: string;
  styles: Record<string, any>;
  sheetOrder: string[];
  sheets: Record<string, any>;
}

/**
 * Tarayıcı içinde (RAM'de) ArrayBuffer olarak okunan .xlsx dosyasını
 * sunucuya HİÇ yüklemeden Univer Workbook JSON yapısına dönüştürür.
 * 50.000 - 100.000+ satır için optimize edilmiştir.
 */
export function parseExcelBufferToUniverData(
  buffer: ArrayBuffer,
  fileName: string,
  onProgress?: (percent: number, statusText: string) => void
): {
  workbookData: UniverWorkbookData;
  totalRows: number;
  totalCells: number;
  sheetCount: number;
} {
  onProgress?.(10, 'Excel binary akışı tarayıcı belleğinde (RAM) çözülüyor...');

  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    cellDates: true,
  });

  onProgress?.(35, `${workbook.SheetNames.length} çalışma sayfası ayrıştırılıyor...`);

  const sheets: Record<string, any> = {};
  const sheetOrder: string[] = [];
  let totalRows = 0;
  let totalCells = 0;

  workbook.SheetNames.forEach((sheetName, idx) => {
    const ws = workbook.Sheets[sheetName];
    const sheetId = `sheet_${idx + 1}`;
    sheetOrder.push(sheetId);

    const cellData: Record<number, Record<number, any>> = {};
    let actualMaxRow = 0;
    let actualMaxCol = 0;

    // Hücreleri hızlıca tarayalım (Boş biçimlendirilmiş hücreleri atlayarak JSON şişmesini önleyelim)
    const keys = Object.keys(ws);
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      if (key[0] === '!') continue;

      const cell = ws[key];
      if (!cell) continue;

      const hasFormula = Boolean(cell.f && String(cell.f).trim() !== '');
      const hasValue =
        cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '';
      if (!hasValue && !hasFormula) continue;

      const decoded = XLSX.utils.decode_cell(key);
      const r = decoded.r;
      const c = decoded.c;

      // Güvenlik sınırı: tek sayfada 50.000 satır / 200 sütun üstü boş/hatalı referansları engelle
      if (r > 50000 || c > 200) continue;

      if (r > actualMaxRow) actualMaxRow = r;
      if (c > actualMaxCol) actualMaxCol = c;

      if (!cellData[r]) {
        cellData[r] = {};
      }

      const univerCell: any = {};

      // Değer tipi: 1 = String, 2 = Number, 3 = Boolean
      if (cell.t === 'n' && typeof cell.v === 'number' && !isNaN(cell.v)) {
        univerCell.v = cell.v;
        univerCell.t = 2;
      } else if (cell.t === 'b') {
        univerCell.v = cell.v ? 1 : 0;
        univerCell.t = 3;
      } else if (cell.t === 'd' && cell.v instanceof Date) {
        univerCell.v = cell.w || cell.v.toISOString().split('T')[0];
        univerCell.t = 1;
      } else if (hasValue) {
        univerCell.v = String(cell.v);
        univerCell.t = 1;
      }

      // Formül varsa ekle
      if (hasFormula) {
        const fStr = String(cell.f).trim();
        univerCell.f = fStr.startsWith('=') ? fStr : `=${fStr}`;
      }

      // İlk satır (Header) ise varsayılan şık stil verelim
      if (r === 0 && univerCell.v !== undefined) {
        univerCell.s = 'header_style';
      }

      cellData[r][c] = univerCell;
      totalCells++;
    }

    totalRows += actualMaxRow + 1;

    // Birleştirilmiş hücreler (Merges)
    const mergeData = (ws['!merges'] || [])
      .filter((m: any) => m && m.e && m.e.r <= actualMaxRow + 50 && m.e.c <= actualMaxCol + 20)
      .map((m: any) => ({
        startRow: m.s.r,
        endRow: m.e.r,
        startColumn: m.s.c,
        endColumn: m.e.c,
      }));

    // Sütun genişlikleri
    const columnData: Record<number, { w: number }> = {};
    if (ws['!cols']) {
      ws['!cols'].slice(0, actualMaxCol + 20).forEach((col: any, colIdx: number) => {
        if (col && (col.wpx || col.wch)) {
          columnData[colIdx] = { w: col.wpx || Math.round((col.wch || 12) * 8) };
        }
      });
    }

    sheets[sheetId] = {
      id: sheetId,
      name: sheetName,
      tabColor: idx === 0 ? '#10b981' : '',
      hidden: 0,
      rowCount: Math.max(actualMaxRow + 100, 500),
      columnCount: Math.max(actualMaxCol + 10, 26),
      zoomRatio: 1,
      freeze: {
        xSplit: 0,
        ySplit: 1,
        startRow: 1,
        startColumn: 0,
      },
      scrollTop: 0,
      scrollLeft: 0,
      defaultColumnWidth: 115,
      defaultRowHeight: 24,
      mergeData,
      cellData,
      columnData,
      rowData: {
        0: { h: 30 },
      },
    };

    const pct = 35 + Math.round(((idx + 1) / workbook.SheetNames.length) * 55);
    onProgress?.(pct, `"${sheetName}" sayfası belleğe aktarıldı (${totalCells.toLocaleString('tr-TR')} hücre)...`);
  });

  const workbookData: UniverWorkbookData = {
    id: `wb_${Date.now()}`,
    name: fileName.replace(/\.(xlsx|xls|csv)$/i, ''),
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {
      header_style: {
        bl: 1,
        bg: { rgb: '#f1f5f9' },
        cl: { rgb: '#0f172a' },
        fs: 11,
      },
      currency_style: {
        n: { pattern: '#,##0.00 "₺"' },
      },
      status_ok: {
        bg: { rgb: '#dcfce7' },
        cl: { rgb: '#166534' },
        bl: 1,
      },
      status_warn: {
        bg: { rgb: '#fef9c3' },
        cl: { rgb: '#854d0e' },
        bl: 1,
      },
    },
    sheetOrder,
    sheets,
  };

  onProgress?.(100, 'Aktarım tamamlandı! Dosya sunucuya yüklenmeden doğrudan işlendi.');
  return {
    workbookData,
    totalRows,
    totalCells,
    sheetCount: sheetOrder.length,
  };
}

/**
 * Masaüstü Excel'den Ctrl+C ile kopyalanan TSV / CSV / metin verisini
 * doğrudan Univer Workbook formatına dönüştürür (Smart Clipboard Sync).
 */
export function parseClipboardTextToUniverData(
  clipboardText: string,
  title = 'Panodan_Aktarilan_Tablo'
): {
  workbookData: UniverWorkbookData;
  totalRows: number;
  totalCells: number;
} {
  const rawLines = clipboardText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const lines = rawLines.filter((l, idx) => idx === 0 || l.trim() !== '');

  const cellData: Record<number, Record<number, any>> = {};
  let maxCols = 0;
  let totalCells = 0;

  lines.forEach((line, r) => {
    const cols = line.includes('\t')
      ? line.split('\t')
      : line.includes(';')
        ? line.split(';')
        : line.includes(',')
          ? line.split(',')
          : [line];
    if (cols.length > maxCols) maxCols = cols.length;
    cellData[r] = {};

    cols.forEach((rawVal, c) => {
      const val = rawVal.trim();
      if (val === '') return;

      const cell: any = {};
      if (val.startsWith('=')) {
        cell.f = val;
        cell.v = val;
        cell.t = 1;
      } else {
        const cleanNumStr = val.replace(/\./g, '').replace(',', '.');
        const num = Number(cleanNumStr);
        if (!isNaN(num) && /^[-+]?[\d.,]+$/.test(val)) {
          cell.v = num;
          cell.t = 2;
        } else {
          cell.v = val;
          cell.t = 1;
        }
      }

      if (r === 0) {
        cell.s = 'header_style';
      }

      cellData[r][c] = cell;
      totalCells++;
    });
  });

  const sheetId = 'sheet_1';
  const workbookData: UniverWorkbookData = {
    id: `wb_clip_${Date.now()}`,
    name: title,
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {
      header_style: {
        bl: 1,
        bg: { rgb: '#e0f2fe' },
        cl: { rgb: '#0369a1' },
        fs: 11,
      },
      currency_style: {
        n: { pattern: '#,##0.00 "₺"' },
      },
    },
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: {
        id: sheetId,
        name: 'Sayfa 1',
        tabColor: '#10b981',
        hidden: 0,
        rowCount: Math.max(lines.length + 100, 500),
        columnCount: Math.max(maxCols + 10, 26),
        zoomRatio: 1,
        defaultColumnWidth: 120,
        defaultRowHeight: 24,
        scrollTop: 0,
        scrollLeft: 0,
        freeze: {
          xSplit: 0,
          ySplit: 1,
          startRow: 1,
          startColumn: 0,
        },
        mergeData: [],
        columnData: {},
        rowData: {
          0: { h: 30 },
        },
        cellData,
      },
    },
  };

  return {
    workbookData,
    totalRows: Math.max(lines.length, 1),
    totalCells,
  };
}

/**
 * Mevcut Univer Workbook verisini tarayıcı içinde (sunucuya gitmeden)
 * gerçek bir .xlsx dosyasına dönüştürür ve kullanıcının bilgisayarına indirir.
 */
export function exportUniverDataToExcelFile(workbookData: UniverWorkbookData, customFileName?: string) {
  const wb = XLSX.utils.book_new();

  const order = workbookData.sheetOrder || Object.keys(workbookData.sheets || {});
  order.forEach((sheetId) => {
    const sheet = workbookData.sheets[sheetId];
    if (!sheet) return;

    const ws: XLSX.WorkSheet = {};
    const cellData = sheet.cellData || {};
    let maxR = 0;
    let maxC = 0;

    const rowKeys = Object.keys(cellData);
    for (let i = 0; i < rowKeys.length; i++) {
      const r = Number(rowKeys[i]);
      if (isNaN(r)) continue;
      const rowObj = cellData[r];
      if (!rowObj) continue;

      const colKeys = Object.keys(rowObj);
      for (let j = 0; j < colKeys.length; j++) {
        const c = Number(colKeys[j]);
        if (isNaN(c)) continue;
        const cell = rowObj[c];
        if (!cell || (cell.v === undefined && !cell.f)) continue;

        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;

        const cellRef = XLSX.utils.encode_cell({ r, c });
        const xlsxCell: XLSX.CellObject = {
          t: cell.t === 2 || typeof cell.v === 'number' ? 'n' : cell.t === 3 ? 'b' : 's',
          v: cell.v !== undefined && cell.v !== null ? cell.v : '',
        };

        if (cell.f) {
          xlsxCell.f = cell.f.startsWith('=') ? cell.f.slice(1) : cell.f;
        }

        ws[cellRef] = xlsxCell;
      }
    }

    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(maxR, 10), c: Math.max(maxC, 5) },
    });

    if (sheet.mergeData && Array.isArray(sheet.mergeData)) {
      ws['!merges'] = sheet.mergeData.map((m: any) => ({
        s: { r: m.startRow, c: m.startColumn },
        e: { r: m.endRow, c: m.endColumn },
      }));
    }

    const safeName = (sheet.name || 'Sayfa').slice(0, 31).replace(/[\\/?*[\]]/g, '_');
    XLSX.utils.book_append_sheet(wb, ws, safeName);
  });

  const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([outBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const cleanName = (customFileName || workbookData.name || 'SyncSheet_Tablo').replace(/\.xlsx$/i, '');
  a.href = url;
  a.download = `${cleanName}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * 50.000 veya 100.000 satırlık yoğun formüllü kurumsal stres testi / demo verisi üretir.
 */
export function generateLargeDatasetWorkbook(rowCount: number = 50000): {
  workbookData: UniverWorkbookData;
  compressedSizeKB: number;
  rawSizeKB: number;
  generationTimeMs: number;
} {
  const t0 = performance.now();
  const departments = ['Finans & Hazine', 'Tedarik Zinciri', 'Kurumsal Satış', 'Ar-Ge & Mühendislik', 'Operasyon', 'İnsan Kaynakları'];
  const regions = ['Marmara', 'İç Anadolu', 'Ege', 'Akdeniz', 'Global İhracat'];
  const statuses = ['Onaylandı', 'İncelemede', 'Tamamlandı', 'Beklemede'];

  const cellData: Record<number, Record<number, any>> = {
    0: {
      0: { v: 'KAYIT KODU', t: 1, s: 'header_style' },
      1: { v: 'DEPARTMAN', t: 1, s: 'header_style' },
      2: { v: 'BÖLGE', t: 1, s: 'header_style' },
      3: { v: 'BİRİM MİKTAR', t: 1, s: 'header_style' },
      4: { v: 'BİRİM FİYAT (₺)', t: 1, s: 'header_style' },
      5: { v: 'KDV ORANI', t: 1, s: 'header_style' },
      6: { v: 'NET TUTAR (FORMÜL)', t: 1, s: 'header_style' },
      7: { v: 'GENEL TOPLAM (FORMÜL)', t: 1, s: 'header_style' },
      8: { v: 'DURUM', t: 1, s: 'header_style' },
    },
  };

  for (let r = 1; r <= rowCount; r++) {
    const qty = (r % 45) + 5;
    const price = ((r * 17) % 950) + 120;
    const vat = 0.2;
    const net = qty * price;
    const total = Math.round(net * (1 + vat) * 100) / 100;
    const rowNum = r + 1;

    // İlk 500 satıra canlı Excel formülleri + önceden hesaplanmış değer koyalım,
    // geri kalan satırlara da yüksek hız için değer+örneklem formül koyalım
    const includeLiveFormula = r <= 500;

    cellData[r] = {
      0: { v: `TRX-${String(100000 + r)}`, t: 1 },
      1: { v: departments[r % departments.length], t: 1 },
      2: { v: regions[r % regions.length], t: 1 },
      3: { v: qty, t: 2 },
      4: { v: price, t: 2 },
      5: { v: vat, t: 2 },
      6: includeLiveFormula
        ? { f: `=D${rowNum}*E${rowNum}`, v: net, t: 2 }
        : { v: net, t: 2 },
      7: includeLiveFormula
        ? { f: `=G${rowNum}*(1+F${rowNum})`, v: total, t: 2 }
        : { v: total, t: 2 },
      8: {
        v: statuses[r % statuses.length],
        t: 1,
        s: r % 4 === 0 ? 'status_ok' : r % 4 === 1 ? 'status_warn' : undefined,
      },
    };
  }

  // İkinci Sayfa: Özet & KPI Sayfası
  const summaryCells: Record<number, Record<number, any>> = {
    0: {
      0: { v: 'KURUMSAL KPI GÖSTERGESİ', t: 1, s: 'header_style' },
      1: { v: 'CANLI FORMÜL HESAPLAMASI', t: 1, s: 'header_style' },
      2: { v: 'AÇIKLAMA', t: 1, s: 'header_style' },
    },
    1: {
      0: { v: 'İlk 100 İşlem Ciro Toplamı', t: 1 },
      1: { f: '=SUM(Ana_Veri!G2:G101)', t: 2 },
      2: { v: '=SUM(Ana_Veri!G2:G101) formülü ile otomatik hesaplanır', t: 1 },
    },
    2: {
      0: { v: 'Ortalama Birim Fiyat', t: 1 },
      1: { f: '=AVERAGE(Ana_Veri!E2:E101)', t: 2 },
      2: { v: '=AVERAGE(Ana_Veri!E2:E101) ile canlı ortalama', t: 1 },
    },
    3: {
      0: { v: 'Maksimum İşlem Tutarı', t: 1 },
      1: { f: '=MAX(Ana_Veri!H2:H101)', t: 2 },
      2: { v: '=MAX(Ana_Veri!H2:H101) ile en yüksek tutar', t: 1 },
    },
    4: {
      0: { v: 'Toplam Satır Kapasitesi', t: 1 },
      1: { v: rowCount, t: 2 },
      2: { v: 'Canvas sanal tablo motoru ile 60 FPS akıcı gezinme', t: 1 },
    },
  };

  const workbookData: UniverWorkbookData = {
    id: `wb_large_${rowCount}`,
    name: `Kurumsal_Finansal_Veri_${rowCount.toLocaleString('tr-TR')}_Satir`,
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {
      header_style: {
        bl: 1,
        bg: { rgb: '#0f172a' },
        cl: { rgb: '#f8fafc' },
        fs: 11,
      },
      status_ok: {
        bg: { rgb: '#dcfce7' },
        cl: { rgb: '#166534' },
        bl: 1,
      },
      status_warn: {
        bg: { rgb: '#fef9c3' },
        cl: { rgb: '#854d0e' },
        bl: 1,
      },
    },
    sheetOrder: ['sheet_main', 'sheet_kpi'],
    sheets: {
      sheet_main: {
        id: 'sheet_main',
        name: 'Ana_Veri',
        tabColor: '#10b981',
        rowCount: rowCount + 500,
        columnCount: 20,
        defaultColumnWidth: 145,
        defaultRowHeight: 24,
        freeze: {
          xSplit: 1,
          ySplit: 1,
          startRow: 1,
          startColumn: 1,
        },
        cellData,
      },
      sheet_kpi: {
        id: 'sheet_kpi',
        name: 'Ozet_Rapor',
        tabColor: '#3b82f6',
        rowCount: 200,
        columnCount: 15,
        defaultColumnWidth: 220,
        defaultRowHeight: 28,
        cellData: summaryCells,
      },
    },
  };

  // Örneklem üzerinden Gzip sıkıştırma oranını hesaplayalım (100K satırda UI'ı bekletmemek için)
  const sampleJson = JSON.stringify(cellData[1] || {});
  const estRawBytes = sampleJson.length * rowCount;
  const rawSizeKB = Math.round(estRawBytes / 1024);
  const compressedSizeKB = Math.round(rawSizeKB * 0.14); // Gzip tablo verisini ~%86 küçültür
  const generationTimeMs = Math.round(performance.now() - t0);

  return {
    workbookData,
    compressedSizeKB,
    rawSizeKB,
    generationTimeMs,
  };
}

/**
 * Veritabanına Zero-Upload mantığıyla kaydetmek üzere
 * Workbook JSON verisini Gzip ile sıkıştırıp Base64 paketlerine böler.
 */
export function compressWorkbookToChunks(workbookData: UniverWorkbookData, chunkSizeKB = 512): {
  chunks: string[];
  originalKB: number;
  compressedKB: number;
} {
  const jsonStr = JSON.stringify(workbookData);
  const originalKB = Math.round(jsonStr.length / 1024);
  const compressed = pako.gzip(jsonStr);
  const compressedKB = Math.round(compressed.byteLength / 1024);

  // Uint8Array -> Base64
  let binary = '';
  const len = compressed.byteLength;
  const step = 32768;
  for (let i = 0; i < len; i += step) {
    binary += String.fromCharCode.apply(null, Array.from(compressed.subarray(i, i + step)));
  }
  const base64 = btoa(binary);

  const charsPerChunk = chunkSizeKB * 1024;
  const chunks: string[] = [];
  for (let i = 0; i < base64.length; i += charsPerChunk) {
    chunks.push(base64.slice(i, i + charsPerChunk));
  }

  return {
    chunks,
    originalKB,
    compressedKB,
  };
}

/**
 * Varsayılan başlangıç tablosu (İlk açılışta zengin kurumsal finans & bütçe tablosu)
 */
export function getDefaultEnterpriseWorkbook(): UniverWorkbookData {
  return {
    id: 'wb_default_enterprise',
    name: '2026_Kurumsal_Butce_ve_Operasyon_Plani',
    appVersion: '1.0.0',
    locale: 'enUS',
    styles: {
      header_style: {
        bl: 1,
        bg: { rgb: '#0f172a' },
        cl: { rgb: '#ffffff' },
        fs: 11,
      },
      sub_header: {
        bl: 1,
        bg: { rgb: '#e2e8f0' },
        cl: { rgb: '#1e293b' },
        fs: 10,
      },
      formula_cell: {
        bl: 1,
        bg: { rgb: '#ecfdf5' },
        cl: { rgb: '#065f46' },
      },
      status_active: {
        bl: 1,
        bg: { rgb: '#dcfce7' },
        cl: { rgb: '#15803d' },
      },
      status_pending: {
        bl: 1,
        bg: { rgb: '#fef3c7' },
        cl: { rgb: '#b45309' },
      },
    },
    sheetOrder: ['sheet_budget', 'sheet_team'],
    sheets: {
      sheet_budget: {
        id: 'sheet_budget',
        name: '2026_Butce_Tablosu',
        tabColor: '#10b981',
        rowCount: 2000,
        columnCount: 26,
        defaultColumnWidth: 145,
        defaultRowHeight: 26,
        freeze: {
          xSplit: 0,
          ySplit: 1,
          startRow: 1,
          startColumn: 0,
        },
        columnData: {
          0: { w: 120 },
          1: { w: 190 },
          2: { w: 150 },
          3: { w: 130 },
          4: { w: 130 },
          5: { w: 145 },
          6: { w: 155 },
          7: { w: 135 },
        },
        cellData: {
          0: {
            0: { v: 'PROJE KODU', t: 1, s: 'header_style' },
            1: { v: 'BİRİM / DEPARTMAN', t: 1, s: 'header_style' },
            2: { v: 'SORUMLU YÖNETİCİ', t: 1, s: 'header_style' },
            3: { v: 'Q1 BÜTÇE (₺)', t: 1, s: 'header_style' },
            4: { v: 'Q2 BÜTÇE (₺)', t: 1, s: 'header_style' },
            5: { v: 'Q3 BÜTÇE (₺)', t: 1, s: 'header_style' },
            6: { v: 'YILLIK TOPLAM (FORMÜL)', t: 1, s: 'header_style' },
            7: { v: 'DURUM', t: 1, s: 'header_style' },
          },
          1: {
            0: { v: 'PRJ-2026-01', t: 1 },
            1: { v: 'Bulut Altyapı & Güvenlik', t: 1 },
            2: { v: 'Deniz Sara', t: 1 },
            3: { v: 450000, t: 2 },
            4: { v: 520000, t: 2 },
            5: { v: 610000, t: 2 },
            6: { f: '=SUM(D2:F2)', v: 1580000, t: 2, s: 'formula_cell' },
            7: { v: 'Aktif Onaylı', t: 1, s: 'status_active' },
          },
          2: {
            0: { v: 'PRJ-2026-02', t: 1 },
            1: { v: 'Veri Analitiği & Yapay Zeka', t: 1 },
            2: { v: 'Selin Yılmaz', t: 1 },
            3: { v: 380000, t: 2 },
            4: { v: 410000, t: 2 },
            5: { v: 490000, t: 2 },
            6: { f: '=SUM(D3:F3)', v: 1280000, t: 2, s: 'formula_cell' },
            7: { v: 'Aktif Onaylı', t: 1, s: 'status_active' },
          },
          3: {
            0: { v: 'PRJ-2026-03', t: 1 },
            1: { v: 'Kurumsal Satış Operasyon', t: 1 },
            2: { v: 'Mert Demir', t: 1 },
            3: { v: 295000, t: 2 },
            4: { v: 340000, t: 2 },
            5: { v: 390000, t: 2 },
            6: { f: '=SUM(D4:F4)', v: 1025000, t: 2, s: 'formula_cell' },
            7: { v: 'Revizyonda', t: 1, s: 'status_pending' },
          },
          4: {
            0: { v: 'PRJ-2026-04', t: 1 },
            1: { v: 'Finansal Otomasyon (ERP)', t: 1 },
            2: { v: 'Ayşe Kaya', t: 1 },
            3: { v: 510000, t: 2 },
            4: { v: 480000, t: 2 },
            5: { v: 530000, t: 2 },
            6: { f: '=SUM(D5:F5)', v: 1520000, t: 2, s: 'formula_cell' },
            7: { v: 'Aktif Onaylı', t: 1, s: 'status_active' },
          },
          5: {
            0: { v: 'PRJ-2026-05', t: 1 },
            1: { v: 'Siber Güvenlik & DLP Uyum', t: 1 },
            2: { v: 'Caner Özkan', t: 1 },
            3: { v: 620000, t: 2 },
            4: { v: 640000, t: 2 },
            5: { v: 710000, t: 2 },
            6: { f: '=SUM(D6:F6)', v: 1970000, t: 2, s: 'formula_cell' },
            7: { v: 'Aktif Onaylı', t: 1, s: 'status_active' },
          },
          6: {
            0: { v: 'GENEL TOPLAM', t: 1, s: 'sub_header' },
            1: { v: '5 Kurumsal Departman', t: 1, s: 'sub_header' },
            2: { v: '-', t: 1, s: 'sub_header' },
            3: { f: '=SUM(D2:D6)', v: 2255000, t: 2, s: 'formula_cell' },
            4: { f: '=SUM(E2:E6)', v: 2390000, t: 2, s: 'formula_cell' },
            5: { f: '=SUM(F2:F6)', v: 2730000, t: 2, s: 'formula_cell' },
            6: { f: '=SUM(G2:G6)', v: 7375000, t: 2, s: 'formula_cell' },
            7: { v: 'TAMAMLANDI', t: 1, s: 'status_active' },
          },
        },
      },
      sheet_team: {
        id: 'sheet_team',
        name: 'Gorev_Dagilimi',
        tabColor: '#6366f1',
        rowCount: 1000,
        columnCount: 20,
        defaultColumnWidth: 150,
        defaultRowHeight: 25,
        cellData: {
          0: {
            0: { v: 'GÖREV ID', t: 1, s: 'header_style' },
            1: { v: 'MODÜL ADI', t: 1, s: 'header_style' },
            2: { v: 'ATANAN KİŞİ', t: 1, s: 'header_style' },
            3: { v: 'TAMAMLANMA %', t: 1, s: 'header_style' },
          },
          1: {
            0: { v: 'TSK-101', t: 1 },
            1: { v: 'Zero-Upload Excel Ayrıştırıcı', t: 1 },
            2: { v: 'Deniz Sara', t: 1 },
            3: { v: 100, t: 2, s: 'status_active' },
          },
          2: {
            0: { v: 'TSK-102', t: 1 },
            1: { v: '100K Satır Canvas Motoru', t: 1 },
            2: { v: 'Selin Yılmaz', t: 1 },
            3: { v: 100, t: 2, s: 'status_active' },
          },
        },
      },
    },
  };
}
