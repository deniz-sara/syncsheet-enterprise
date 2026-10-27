import type { NextApiRequest, NextApiResponse } from 'next';
import { dbGetRoom } from '@/lib/supabaseStorage';

function extractSheetRows(workbookData: any, requestedSheetId?: string): {
  sheetName: string;
  rows: string[][];
} {
  if (!workbookData || !workbookData.sheets) {
    return {
      sheetName: 'Bos_Sayfa',
      rows: [['Kayitli oda verisi bulunamadi. Once SyncSheet uzerinde Kaydet butonuna basin.']],
    };
  }

  const sheetIds = workbookData.sheetOrder?.length
    ? workbookData.sheetOrder
    : Object.keys(workbookData.sheets);

  const targetId =
    requestedSheetId && workbookData.sheets[requestedSheetId]
      ? requestedSheetId
      : sheetIds[0];

  const sheet = workbookData.sheets[targetId];
  if (!sheet || !sheet.cellData) {
    return { sheetName: sheet?.name || 'Sayfa1', rows: [] };
  }

  const cellData = sheet.cellData;
  const rowKeys = Object.keys(cellData)
    .map(Number)
    .filter((n) => !isNaN(n));

  if (rowKeys.length === 0) {
    return { sheetName: sheet.name || 'Sayfa1', rows: [] };
  }

  const maxRow = Math.min(Math.max(...rowKeys), 5000);
  let maxCol = 0;

  for (const r of rowKeys) {
    if (r > maxRow) continue;
    const cols = Object.keys(cellData[r] || {})
      .map(Number)
      .filter((n) => !isNaN(n));
    if (cols.length > 0) {
      maxCol = Math.max(maxCol, ...cols);
    }
  }
  maxCol = Math.min(maxCol, 50);

  const rows: string[][] = [];
  for (let r = 0; r <= maxRow; r++) {
    const rowArr: string[] = [];
    let hasAnyValue = false;
    for (let c = 0; c <= maxCol; c++) {
      const cell = cellData[r]?.[c];
      const val = cell?.v !== undefined && cell?.v !== null ? String(cell.v) : '';
      if (val !== '') hasAnyValue = true;
      rowArr.push(val);
    }
    if (hasAnyValue || r === 0) {
      rows.push(rowArr);
    }
  }

  return {
    sheetName: sheet.name || 'Sayfa1',
    rows,
  };
}

function escapeCsvCell(val: string): string {
  if (val.includes('"') || val.includes(',') || val.includes(';') || val.includes('\n')) {
    return `"${val.replace(/"/g, '""')}"`;
  }
  return val;
}

function escapeHtml(val: string): string {
  return val
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');

  const roomCode = (req.query.room as string) || 'XLS-849-201';
  const format = ((req.query.format as string) || 'html').toLowerCase();
  const sheetId = req.query.sheet as string | undefined;

  const room = await dbGetRoom(roomCode);
  const { sheetName, rows } = extractSheetRows(room?.workbookData, sheetId);

  // 1) .iqy (Microsoft Excel Web Query file)
  if (format === 'iqy') {
    const host = req.headers.host || 'localhost:3005';
    const protocol = host.includes('localhost') ? 'http' : 'https';
    const liveHtmlUrl = `${protocol}://${host}/api/excel-bridge?room=${encodeURIComponent(roomCode)}&format=html${sheetId ? `&sheet=${encodeURIComponent(sheetId)}` : ''}`;

    const iqyContent = [
      'WEB',
      '1',
      liveHtmlUrl,
      '',
      'Selection=EntirePage',
      'Formatting=None',
      'PreFormattedTextToColumns=True',
      'ConsecutiveDelimitersAsOne=True',
      'SingleBlockTextImport=False',
      'DisableDateRecognition=False',
      'DisableRedirections=False',
    ].join('\r\n');

    res.setHeader('Content-Type', 'text/x-ms-iqy; charset=utf-8');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="SyncSheet_Canli_Baglanti_${roomCode}.iqy"`
    );
    return res.status(200).send(iqyContent);
  }

  // 2) CSV feed (Power Query / Data -> From Web in Excel)
  if (format === 'csv') {
    const delimiter = req.query.sep === ';' ? ';' : ',';
    const csvLines = rows.map((row) => row.map(escapeCsvCell).join(delimiter));
    // UTF-8 BOM so Turkish characters open cleanly in Microsoft Excel
    const csvContent = '\uFEFF' + csvLines.join('\r\n');
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader(
      'Content-Disposition',
      `inline; filename="SyncSheet_${roomCode}_${sheetName}.csv"`
    );
    return res.status(200).send(csvContent);
  }

  // 3) JSON feed
  if (format === 'json') {
    return res.status(200).json({
      roomCode,
      sheetName,
      updatedAt: room?.updatedAt || new Date().toISOString(),
      rowCount: rows.length,
      rows,
    });
  }

  // 4) Default: Clean HTML Table for Microsoft Excel Web Query (.iqy / From Web)
  const headerRow = rows[0] || [];
  const bodyRows = rows.slice(1);

  const html = `<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8" />
  <title>SyncSheet Canlı Veri Köprüsü - ${escapeHtml(roomCode)} (${escapeHtml(sheetName)})</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; margin: 16px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: left; }
    th { background-color: #ecfdf5; color: #065f46; font-weight: bold; }
  </style>
</head>
<body>
  <table id="SyncSheetLiveTable">
    <thead>
      <tr>
        ${headerRow.map((h) => `<th>${escapeHtml(h)}</th>`).join('')}
      </tr>
    </thead>
    <tbody>
      ${bodyRows
        .map(
          (r) =>
            `<tr>${r
              .map((cell) => {
                const num = Number(cell);
                const isNum = cell.trim() !== '' && !isNaN(num);
                return `<td${isNum ? ` data-type="number"` : ''}>${escapeHtml(cell)}</td>`;
              })
              .join('')}</tr>`
        )
        .join('\n      ')}
    </tbody>
  </table>
</body>
</html>`;

  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  return res.status(200).send(html);
}
