'use client';

import React, { useEffect, useRef } from 'react';
import ReactDOM from 'react-dom';
import { createRoot } from 'react-dom/client';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/presets/lib/styles/preset-sheets-filter.css';
import '@univerjs/presets/lib/styles/preset-sheets-sort.css';
import '@univerjs/presets/lib/styles/preset-sheets-data-validation.css';
import '@univerjs/presets/lib/styles/preset-sheets-conditional-formatting.css';
import '@univerjs/presets/lib/styles/preset-sheets-find-replace.css';

import { UniverWorkbookData } from '@/lib/excelZeroUpload';
import { formatCellRef } from '@/lib/collaboration';

function ensureReact18InternalsShim() {
  const r = React as any;
  if (r && !r.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED) {
    r.__SECRET_INTERNALS_DO_NOT_USE_OR_YOU_WILL_BE_FIRED = {
      ReactCurrentOwner: { current: null },
      ReactCurrentDispatcher: { current: null },
      ReactCurrentBatchConfig: { transition: null },
      ReactCurrentActQueue: { current: null, isBatchingLegacy: false, didScheduleLegacyUpdate: false },
    };
  }

  const dom = ReactDOM as any;
  if (dom) {
    if (!dom.createRoot) {
      dom.createRoot = createRoot;
    }
    if (!dom.render) {
      dom.render = (element: any, container: any) => {
        if (!container) return;
        const root = container.__univer_react_root__ || createRoot(container);
        container.__univer_react_root__ = root;
        root.render(element);
      };
    }
    if (!dom.unmountComponentAtNode) {
      dom.unmountComponentAtNode = (container: any) => {
        if (container && container.__univer_react_root__) {
          try {
            container.__univer_react_root__.unmount();
          } catch (e) {
            // ignore
          }
          delete container.__univer_react_root__;
          return true;
        }
        return false;
      };
    }
  }
}

ensureReact18InternalsShim();

function mergeLocales(...locales: Record<string, any>[]): Record<string, any> {
  const result: Record<string, any> = {};
  for (const loc of locales) {
    if (!loc) continue;
    for (const key of Object.keys(loc)) {
      if (
        loc[key] &&
        typeof loc[key] === 'object' &&
        !Array.isArray(loc[key]) &&
        result[key] &&
        typeof result[key] === 'object'
      ) {
        result[key] = mergeLocales(result[key], loc[key]);
      } else {
        result[key] = loc[key];
      }
    }
  }
  return result;
}

/**
 * Formüller İngilizce (SUM, XLOOKUP, VLOOKUP, IF, AVERAGE vb.) kalırken
 * Menü Bar, Şerit (Ribbon), Sağ Tık Menüsü ve Araç İpuçlarını Türkçeleştiren sözlük:
 */
const TURKISH_MENU_LOCALE_OVERRIDES = {
  spreadsheetLabel: 'Çalışma Sayfası',
  spreadsheetRightLabel: 'Diğer Sayfalar',
  ribbon: {
    start: 'Giriş',
    insert: 'Ekle',
    formulas: 'Formüller',
    data: 'Veri',
    view: 'Görünüm',
    others: 'Diğer',
    more: 'Daha Fazla',
  },
  toolbar: {
    undo: 'Geri Al (Ctrl+Z)',
    redo: 'Yinele (Ctrl+Y)',
    formatPainter: 'Biçim Boyacısı',
    currencyFormat: 'Para Birimi Biçimi (₺)',
    percentageFormat: 'Yüzde Biçimi (%)',
    numberDecrease: 'Ondalık Basamak Azalt',
    numberIncrease: 'Ondalık Basamak Artır',
    moreFormats: 'Sayı Biçimleri',
    font: 'Yazı Tipi',
    fontSize: 'Yazı Boyutu',
    bold: 'Kalın (Ctrl+B)',
    italic: 'İtalik (Ctrl+I)',
    strikethrough: 'Üstü Çizili',
    subscript: 'Alt Simge',
    superscript: 'Üst Simge',
    underline: 'Altı Çizili (Ctrl+U)',
    textColor: {
      main: 'Yazı Rengi',
      right: 'Renk Seçin',
    },
    resetColor: 'Sıfırla',
    customColor: 'ÖZEL RENK',
    alternatingColors: 'Alternatif Satır Renkleri',
    confirmColor: 'Tamam',
    cancelColor: 'İptal',
    collapse: 'Daralt',
    fillColor: {
      main: 'Hücre Dolgu Rengi',
      right: 'Renk Seçin',
    },
    border: {
      main: 'Kenarlıklar',
      right: 'Kenarlık Stili',
    },
    mergeCell: {
      main: 'Birleştir ve Ortala',
      right: 'Birleştirme Türü',
    },
    horizontalAlignMode: {
      main: 'Yatay Hizalama',
      right: 'Hizalama',
    },
    verticalAlignMode: {
      main: 'Dikey Hizalama',
      right: 'Hizalama',
    },
    textWrapMode: {
      main: 'Metni Kaydır',
      right: 'Metin Kaydırma Modu',
    },
    textRotateMode: {
      main: 'Metni Döndür',
      right: 'Döndürme Açısı',
    },
    freezeTopRow: 'Üst Satırı Dondur',
    sortAndFilter: 'Sırala ve Filtrele',
    findAndReplace: 'Bul ve Değiştir',
    sum: 'SUM (Topla)',
    autoSum: 'Otomatik Toplam (SUM)',
    moreFunction: 'Fonksiyon Ekle (fx)',
    conditionalFormatting: 'Koşullu Biçimlendirme',
    comment: 'Açıklama Ekle',
    pivotTable: 'Pivot Tablo',
    chart: 'Grafik',
    dataValidation: 'Veri Doğrulama',
    protection: 'Sayfayı Koru',
    clearText: 'Rengi Temizle',
    toolMore: 'Daha Fazla Araç',
    toolMoreTip: 'Diğer Özellikler',
    hideGridlines: 'Kılavuz Çizgilerini Gizle',
    showGridlines: 'Kılavuz Çizgilerini Göster',
    toggleGridlines: 'Kılavuz Çizgileri Aç/Kapat',
  },
  align: {
    left: 'Sola Hizala',
    center: 'Ortala',
    right: 'Sağa Hizala',
    top: 'Üste Hizala',
    middle: 'Dikey Ortala',
    bottom: 'Alta Hizala',
  },
  textWrap: {
    overflow: 'Taşır',
    wrap: 'Metni Kaydır',
    clip: 'Kırp',
  },
  merge: {
    all: 'Tümünü Birleştir',
    vertical: 'Dikey Birleştir',
    horizontal: 'Yatay Birleştir',
    cancel: 'Birleştirmeyi Kaldır',
  },
  sheetConfig: {
    delete: 'Sayfayı Sil',
    copy: 'Sayfayı Kopyala',
    rename: 'Sayfayı Yeniden Adlandır',
    changeColor: 'Sekme Rengini Değiştir',
    hide: 'Sayfayı Gizle',
    unhide: 'Gizli Sayfaları Göster',
    moveLeft: 'Sola Taşı',
    moveRight: 'Sağa Taşı',
    resetColor: 'Rengi Sıfırla',
    cancelText: 'İptal',
    chooseText: 'Onayla',
  },
  rightClick: {
    copy: 'Kopyala (Ctrl+C)',
    copyAs: 'Farklı Kopyala',
    cut: 'Kes (Ctrl+X)',
    paste: 'Yapıştır (Ctrl+V)',
    pasteSpecial: 'Özel Yapıştır',
    pasteValue: 'Sadece Değerleri Yapıştır',
    pasteFormat: 'Sadece Biçimi Yapıştır',
    pasteColWidth: 'Sütun Genişliğini Yapıştır',
    insert: 'Ekle',
    insertRow: 'Üste Satır Ekle',
    insertRowBefore: 'Öncesine Satır Ekle',
    insertColumn: 'Sola Sütun Ekle',
    insertColumnBefore: 'Öncesine Sütun Ekle',
    delete: 'Sil',
    deleteCell: 'Hücre Sil',
    insertCell: 'Hücre Ekle',
    deleteSelectedRow: 'Seçili Satırı Sil',
    deleteSelectedColumn: 'Seçili Sütunu Sil',
    hideSelectedRow: 'Seçili Satırı Gizle',
    showHideRow: 'Gizli Satırları Göster',
    rowHeight: 'Satır Yüksekliği',
    hideSelectedColumn: 'Seçili Sütunu Gizle',
    showHideColumn: 'Gizli Sütunları Göster',
    columnWidth: 'Sütun Genişliği',
    orderAZ: "A'dan Z'ye Sırala (Artan)",
    orderZA: "Z'den A'ya Sırala (Azalan)",
    clearSelection: 'Temizle',
    clearContent: 'İçeriği Temizle',
    clearFormat: 'Biçimlendirmeyi Temizle',
    clearAll: 'Tümünü Temizle',
    sortSelection: 'Sırala',
    filterSelection: 'Filtrele',
  },
  button: {
    confirm: 'Tamam',
    cancel: 'İptal',
    close: 'Kapat',
    update: 'Güncelle',
    delete: 'Sil',
    insert: 'Ekle',
  },
};

export interface UniverSpreadsheetHandle {
  getCurrentSnapshot: () => UniverWorkbookData;
  loadWorkbookSnapshot: (wb: UniverWorkbookData) => void;
  applyRemoteCellUpdate: (sheetId: string, row: number, col: number, value: any, formula?: string) => void;
  setFormulaOrValueOnCell: (
    sheetId: string,
    row: number,
    col: number,
    formula: string | undefined,
    computedValue: any
  ) => void;
  navigateToCell: (row: number, col: number) => void;
  runScriptOnActiveSheet: (scriptFn: (sheetApi: any) => void) => void;
}

interface Props {
  workbookData: UniverWorkbookData;
  onSelectionChange?: (info: {
    sheetId: string;
    sheetName: string;
    row: number;
    col: number;
    cellRef: string;
    value: any;
    formula?: string;
  }) => void;
  onCellEdited?: (info: {
    sheetId: string;
    row: number;
    col: number;
    value: any;
    formula?: string;
    previousValue?: any;
    previousFormula?: string;
  }) => boolean | void;
  onUserActivity?: () => void;
  onFxIconClicked?: () => void;
  onReady?: (handle: UniverSpreadsheetHandle) => void;
}

export default function UniverSpreadsheet({
  workbookData,
  onSelectionChange,
  onCellEdited,
  onUserActivity,
  onFxIconClicked,
  onReady,
}: Props) {
  const wrapperRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const univerAPIRef = useRef<any>(null);
  const currentUnitIdRef = useRef<string | null>(null);

  // Canlı olarak her hücre düzenlemesinde güncellenen tam Workbook kopyası:
  // Böylece versiyon geçmişine kayıt yaparken veya geri dönerken HİÇBİR değişiklik kaybolmaz!
  const currentWorkbookDataRef = useRef<UniverWorkbookData>(
    JSON.parse(JSON.stringify(workbookData))
  );
  const isApplyingRemoteRef = useRef<boolean>(false);

  const cellCacheRef = useRef<Record<string, { v: any; f?: string }>>({});

  const onSelectionChangeRef = useRef(onSelectionChange);
  const onCellEditedRef = useRef(onCellEdited);
  const onUserActivityRef = useRef(onUserActivity);
  const onFxIconClickedRef = useRef(onFxIconClicked);
  const onReadyRef = useRef(onReady);

  useEffect(() => {
    onSelectionChangeRef.current = onSelectionChange;
    onCellEditedRef.current = onCellEdited;
    onUserActivityRef.current = onUserActivity;
    onFxIconClickedRef.current = onFxIconClicked;
    onReadyRef.current = onReady;
  }, [onSelectionChange, onCellEdited, onUserActivity, onFxIconClicked, onReady]);

  const populateCacheFromWorkbook = (wb: UniverWorkbookData) => {
    const cache: Record<string, { v: any; f?: string }> = {};
    if (wb && wb.sheets) {
      Object.values(wb.sheets).forEach((sheet: any) => {
        const sid = sheet.id;
        const cData = sheet.cellData || {};
        Object.keys(cData).slice(0, 1500).forEach((rKey) => {
          const rowObj = cData[rKey];
          if (!rowObj) return;
          Object.keys(rowObj).forEach((cKey) => {
            const cell = rowObj[cKey];
            if (cell) {
              cache[`${sid}_${rKey}_${cKey}`] = { v: cell.v, f: cell.f };
            }
          });
        });
      });
    }
    cellCacheRef.current = cache;
  };

  // Yardımcı: Bellekteki currentWorkbookDataRef içine yapılan hücre değişikliğini anında işle
  const syncCellIntoMemoryWorkbook = (
    sheetId: string,
    row: number,
    col: number,
    val: any,
    formula?: string
  ) => {
    const wb = currentWorkbookDataRef.current;
    if (!wb || !wb.sheets) return;
    const targetSid = wb.sheets[sheetId]
      ? sheetId
      : wb.sheetOrder?.[0] || Object.keys(wb.sheets)[0];
    const sheet = wb.sheets[targetSid];
    if (!sheet) return;
    if (!sheet.cellData) sheet.cellData = {};
    if (!sheet.cellData[row]) sheet.cellData[row] = {};
    const existingCell = sheet.cellData[row][col] || {};
    sheet.cellData[row][col] = {
      ...existingCell,
      v: val,
      t: typeof val === 'number' ? 2 : 1,
      f: formula || undefined,
    };
  };

  // Yardımcı: Univer içine yeni bir workbook nesnesini güvenli şekilde yükle (Çift yüklemeyi ve donmayı önler)
  const lastMountedRef = useRef<any>(null);
  const mountWorkbookIntoUniver = (wbToLoad: UniverWorkbookData) => {
    const api = univerAPIRef.current;
    if (!api || !wbToLoad) return;
    if (lastMountedRef.current === wbToLoad) return;
    lastMountedRef.current = wbToLoad;

    try {
      isApplyingRemoteRef.current = true;
      const cleanClone: UniverWorkbookData = JSON.parse(JSON.stringify(wbToLoad));
      currentWorkbookDataRef.current = cleanClone;
      populateCacheFromWorkbook(cleanClone);

      const oldUnitId = currentUnitIdRef.current;
      const newUnitId = `wb_unit_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      currentUnitIdRef.current = newUnitId;

      if (oldUnitId && oldUnitId !== newUnitId && typeof api.disposeUnit === 'function') {
        try {
          api.disposeUnit(oldUnitId);
        } catch (e) {
          // ignore
        }
      }

      api.createWorkbook({
        ...cleanClone,
        id: newUnitId,
      });
    } catch (err) {
      console.error('Workbook mount error:', err);
    } finally {
      setTimeout(() => {
        isApplyingRemoteRef.current = false;
      }, 80);
    }
  };

  useEffect(() => {
    let isCancelled = false;
    let disposable: any = null;

    async function bootUniverEngineOnce() {
      if (!wrapperRef.current || univerAPIRef.current) return;

      ensureReact18InternalsShim();

      const [
        presetsMod,
        coreMod,
        coreLocMod,
        filterMod,
        filterLocMod,
        sortMod,
        sortLocMod,
        dvMod,
        dvLocMod,
        cfMod,
        cfLocMod,
        frMod,
        frLocMod,
      ] = await Promise.all([
        import('@univerjs/presets'),
        import('@univerjs/preset-sheets-core'),
        import('@univerjs/preset-sheets-core/locales/en-US'),
        import('@univerjs/preset-sheets-filter'),
        import('@univerjs/preset-sheets-filter/locales/en-US'),
        import('@univerjs/preset-sheets-sort'),
        import('@univerjs/preset-sheets-sort/locales/en-US'),
        import('@univerjs/preset-sheets-data-validation'),
        import('@univerjs/preset-sheets-data-validation/locales/en-US'),
        import('@univerjs/preset-sheets-conditional-formatting'),
        import('@univerjs/preset-sheets-conditional-formatting/locales/en-US'),
        import('@univerjs/preset-sheets-find-replace'),
        import('@univerjs/preset-sheets-find-replace/locales/en-US'),
      ]);

      if (isCancelled || !wrapperRef.current) return;

      const { createUniver, defaultTheme, LocaleType } = presetsMod;
      const { UniverSheetsCorePreset } = coreMod;
      const { UniverSheetsFilterPreset } = filterMod;
      const { UniverSheetsSortPreset } = sortMod;
      const { UniverSheetsDataValidationPreset } = dvMod;
      const { UniverSheetsConditionalFormattingPreset } = cfMod;
      const { UniverSheetsFindReplacePreset } = frMod;

      const { univer, univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: mergeLocales(
            coreLocMod.default || coreLocMod,
            filterLocMod.default || filterLocMod,
            sortLocMod.default || sortLocMod,
            dvLocMod.default || dvLocMod,
            cfLocMod.default || cfLocMod,
            frLocMod.default || frLocMod,
            TURKISH_MENU_LOCALE_OVERRIDES
          ),
        },
        theme: defaultTheme,
        presets: [
          UniverSheetsCorePreset({
            container: wrapperRef.current,
            header: true,
            toolbar: true,
            formulaBar: true,
            footer: true,
            contextMenu: true,
          }),
          UniverSheetsFilterPreset(),
          UniverSheetsSortPreset(),
          UniverSheetsDataValidationPreset(),
          UniverSheetsConditionalFormattingPreset(),
          UniverSheetsFindReplacePreset(),
        ],
      });

      univerRef.current = univer;
      univerAPIRef.current = univerAPI;

      const initialData = currentWorkbookDataRef.current;
      populateCacheFromWorkbook(initialData);

      const uniqueId = `wb_unit_${Date.now()}`;
      currentUnitIdRef.current = uniqueId;
      univerAPI.createWorkbook({
        ...initialData,
        id: uniqueId,
      } as any);

      disposable = univerAPI.onCommandExecuted((command: any) => {
        if (isApplyingRemoteRef.current) return;

        const id: string = command?.id || '';

        if (id.includes('set-selections') || id.includes('selection')) {
          try {
            const activeWorkbook = univerAPI.getActiveWorkbook();
            const activeSheet = activeWorkbook?.getActiveSheet();
            const selections = command?.params?.selections;
            if (activeSheet && selections && selections.length > 0) {
              const primary = selections[0]?.primary || selections[0]?.range;
              const r = primary?.actualRow ?? primary?.startRow ?? 0;
              const c = primary?.actualColumn ?? primary?.startColumn ?? 0;
              const sid = activeSheet.getSheetId();
              const range = activeSheet.getRange(r, c, 1, 1);
              const val = range?.getValue?.();
              const formula = (range as any)?.getFormula?.() || range?.getFormulas?.()?.[0]?.[0];

              cellCacheRef.current[`${sid}_${r}_${c}`] = { v: val, f: formula || undefined };

              onSelectionChangeRef.current?.({
                sheetId: sid,
                sheetName: activeSheet.getSheetName(),
                row: r,
                col: c,
                cellRef: formatCellRef(r, c),
                value: val,
                formula: formula || undefined,
              });
            }
          } catch (err) {
            // ignore
          }
        }

        if (id.includes('set-range-values') && command?.params?.cellValue) {
          try {
            const cellValueMatrix = command.params.cellValue;
            const activeWb = univerAPI.getActiveWorkbook();
            const activeSheet = activeWb?.getActiveSheet();
            const subUnitId =
              command.params.subUnitId ||
              activeSheet?.getSheetId() ||
              'sheet_1';

            Object.keys(cellValueMatrix).forEach((rKey) => {
              const r = Number(rKey);
              const rowObj = cellValueMatrix[rKey];
              if (!rowObj) return;
              Object.keys(rowObj).forEach((cKey) => {
                const c = Number(cKey);
                const cell = rowObj[cKey];
                if (cell) {
                  const cacheKey = `${subUnitId}_${r}_${c}`;
                  const prev = cellCacheRef.current[cacheKey];

                  const allowed = onCellEditedRef.current?.({
                    sheetId: subUnitId,
                    row: r,
                    col: c,
                    value: cell.v,
                    formula: cell.f,
                    previousValue: prev?.v,
                    previousFormula: prev?.f,
                  });

                  // Hücre kilitliyse eski değerine geri al
                  if (allowed === false) {
                    isApplyingRemoteRef.current = true;
                    const targetSheet =
                      activeWb?.getSheetBySheetId(subUnitId) || activeSheet;
                    if (targetSheet) {
                      const rng = targetSheet.getRange(r, c, 1, 1);
                      if (prev?.f) {
                        rng.setValue(prev.f);
                      } else {
                        rng.setValue(prev?.v !== undefined ? prev.v : '');
                      }
                    }
                    setTimeout(() => {
                      isApplyingRemoteRef.current = false;
                    }, 60);
                  } else {
                    cellCacheRef.current[cacheKey] = {
                      v: cell.v,
                      f: cell.f,
                    };
                    // Bellekteki tam Workbook nesnesini de anında güncelle!
                    syncCellIntoMemoryWorkbook(subUnitId, r, c, cell.v, cell.f);
                  }
                }
              });
            });
          } catch (err) {
            // ignore
          }
        }
      });

      const handle: UniverSpreadsheetHandle = {
        getCurrentSnapshot: () => {
          return JSON.parse(JSON.stringify(currentWorkbookDataRef.current));
        },
        loadWorkbookSnapshot: (wb: UniverWorkbookData) => {
          mountWorkbookIntoUniver(wb);
        },
        applyRemoteCellUpdate: (sheetId, row, col, value, formula) => {
          try {
            isApplyingRemoteRef.current = true;
            cellCacheRef.current[`${sheetId}_${row}_${col}`] = {
              v: value,
              f: formula,
            };
            syncCellIntoMemoryWorkbook(sheetId, row, col, value, formula);

            const activeWb = univerAPIRef.current?.getActiveWorkbook();
            const sheet = activeWb?.getSheetBySheetId(sheetId) || activeWb?.getActiveSheet();
            if (sheet) {
              const range = sheet.getRange(row, col, 1, 1);
              if (formula) {
                range.setValue(formula);
              } else {
                range.setValue(value);
              }
            }
          } catch (e) {
            // ignore
          } finally {
            setTimeout(() => {
              isApplyingRemoteRef.current = false;
            }, 50);
          }
        },
        setFormulaOrValueOnCell: (sheetId, row, col, formula, computedValue) => {
          try {
            isApplyingRemoteRef.current = true;
            cellCacheRef.current[`${sheetId}_${row}_${col}`] = {
              v: computedValue,
              f: formula,
            };
            syncCellIntoMemoryWorkbook(sheetId, row, col, computedValue, formula);

            const activeWb = univerAPIRef.current?.getActiveWorkbook();
            const sheet = activeWb?.getSheetBySheetId(sheetId) || activeWb?.getActiveSheet();
            if (sheet) {
              const range = sheet.getRange(row, col, 1, 1);
              if (formula) {
                range.setValue(formula);
              } else {
                range.setValue(computedValue);
              }
            }
          } catch (e) {
            // ignore
          } finally {
            setTimeout(() => {
              isApplyingRemoteRef.current = false;
            }, 50);
          }
        },
        navigateToCell: (row, col) => {
          try {
            const activeWb = univerAPIRef.current?.getActiveWorkbook();
            const sheet = activeWb?.getActiveSheet();
            if (sheet && typeof sheet.setActiveSelection === 'function') {
              const range = sheet.getRange(row, col, 1, 1);
              sheet.setActiveSelection(range);
            }
          } catch (e) {
            // ignore
          }
        },
        runScriptOnActiveSheet: (scriptFn) => {
          try {
            const activeWb = univerAPIRef.current?.getActiveWorkbook();
            const sheet = activeWb?.getActiveSheet();
            if (sheet) {
              scriptFn(sheet);
            }
          } catch (e) {
            console.error('Makro çalıştırma hatası:', e);
          }
        },
      };

      onReadyRef.current?.(handle);
    }

    bootUniverEngineOnce();

    return () => {
      isCancelled = true;
      try {
        disposable?.dispose?.();
      } catch (e) {
        // ignore
      }
    };
  }, []);

  useEffect(() => {
    if (workbookData) {
      currentWorkbookDataRef.current = JSON.parse(JSON.stringify(workbookData));
      populateCacheFromWorkbook(currentWorkbookDataRef.current);
    }
    if (!univerAPIRef.current) return;
    mountWorkbookIntoUniver(workbookData);
  }, [workbookData]);

  return (
    <div
      ref={wrapperRef}
      onKeyDownCapture={() => onUserActivityRef.current?.()}
      onClickCapture={(e) => {
        // Eğer kullanıcı Univer formül çubuğundaki "fx" ikonuna tıklarsa kolay fx sihirbazını da açalım
        const target = e.target as HTMLElement | null;
        if (target) {
          const text = (target.textContent || '').trim().toLowerCase();
          const cls = String(target.className || '');
          if (text === 'fx' || cls.includes('formula-icon') || target.closest('[class*="formula-bar"] svg')) {
            // Sadece sol taraftaki fx butonuna tıklandığında tetikle
            const rect = target.getBoundingClientRect();
            if (rect.left < 220 && rect.top < 260) {
              onFxIconClickedRef.current?.();
            }
          }
        }
      }}
      className="relative w-full h-full flex-1 bg-white overflow-hidden"
    />
  );
}
