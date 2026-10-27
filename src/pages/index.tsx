import React, { useState, useEffect, useRef, useMemo } from 'react';
import dynamic from 'next/dynamic';
import {
  FileSpreadsheet,
  Upload,
  Download,
  ShieldCheck,
  Sparkles,
  Share2,
  ClipboardPaste,
  Database,
  Eye,
  Lock,
  Unlock,
  Check,
  Copy,
  Layers,
  Zap,
  Terminal,
  Filter,
  FolderOpen,
  X,
  History,
  BarChart3,
  RotateCcw,
  ChevronDown,
  Clock,
  Plus,
  FunctionSquare,
  PaintBucket,
  Table2,
  LogOut,
  UserCheck,
  Save,
  Trash2,
  Star,
  Search,
  Link2,
  RefreshCw,
  Radio,
  ExternalLink,
} from 'lucide-react';

import {
  UniverWorkbookData,
  parseExcelBufferToUniverData,
  parseClipboardTextToUniverData,
  exportUniverDataToExcelFile,
  generateLargeDatasetWorkbook,
  getDefaultEnterpriseWorkbook,
} from '@/lib/excelZeroUpload';

import {
  CollaboratorPresence,
  RealtimeRoomClient,
  getRandomUserColor,
  formatCellRef,
  colIndexToLetter,
  getActiveCellLockOwner,
} from '@/lib/collaboration';

import {
  EXCEL_FORMULA_CATALOG,
  getRecentFormulaNames,
  recordFormulaUsage,
  getSheetColumnRanges,
  evaluateFormulaPreview,
} from '@/lib/excelFormulas';

import {
  PivotConfig,
  PivotAggFunc,
  computePivotTable,
  insertPivotSheetIntoWorkbook,
  extractSheetHeaders,
  ConditionalRuleOperator,
  ConditionalColorPreset,
  ConditionalRuleItem,
  COLOR_PRESET_MAP,
  applyConditionalFormattingRule,
  getSalesCrmWorkbook,
  getBlankWorkbook,
} from '@/lib/pivotAndFormatting';

import {
  LinkedDesktopExcelInfo,
  mergeExternalExcelIntoWorkbook,
  applyBulkCrossWorkbookLookup,
  downloadSampleDesktopExcelForVlookup,
} from '@/lib/externalExcelBridge';

import type { UniverSpreadsheetHandle } from '@/components/UniverSpreadsheet';

const UniverSpreadsheet = dynamic(() => import('@/components/UniverSpreadsheet'), {
  ssr: false,
  loading: () => (
    <div className="flex-1 flex flex-col items-center justify-center bg-slate-50 text-slate-600 gap-3">
      <div className="w-9 h-9 border-3 border-emerald-600 border-t-transparent rounded-full animate-spin" />
      <p className="text-sm font-medium text-slate-700">Excel Çalışma Alanı Hazırlanıyor...</p>
    </div>
  ),
});

interface AuditEvent {
  id: string;
  roomCode?: string; // Yalnızca ilgili Excel odasında gözükmesi için
  userId: string;
  userName: string;
  userEmail?: string;
  userDepartment?: string;
  userColor: string;
  action: string;
  cellRef?: string;
  sheetName?: string;
  oldValue?: any;
  newValue?: any;
  timestamp: number;
  snapshotId?: string; // Bu değişikliğe tek tıkla geri dönebilmek için
}

interface SnapshotMeta {
  id: string;
  roomCode?: string; // Yalnızca ilgili Excel odasında gözükmesi için
  label: string;
  createdBy: string;
  createdByDepartment?: string;
  createdAt: string;
  timestamp: number; // Kesin sıralama (en son yapılan en üstte) için
  rowCount: number;
  isAutoChange?: boolean;
}

interface OpenWorkbookTab {
  id: string;
  roomCode?: string;
  name: string;
  workbookData: UniverWorkbookData;
  updatedAt: number;
}

interface UserSavedRoomItem {
  roomCode: string;
  title: string;
  pin: string;
  role: 'owner' | 'editor' | 'viewer';
  lastAccessedAt: string;
  updatedAt?: string;
}

const AUTH_STORAGE_KEY = 'syncsheet_auth_user_v2';

export default function SyncSheetApp() {
  // 1) GERÇEK KULLANICI ADI & ŞİFRE GİRİŞ / KAYIT SİSTEMİ VE KULLANICIYA ÖZEL ODA-PIN GEÇMİŞİ
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [authMode, setAuthMode] = useState<'login' | 'register'>('login');
  const [authError, setAuthError] = useState<string | null>(null);
  const [authLoading, setAuthLoading] = useState<boolean>(false);

  const [currentUser, setCurrentUser] = useState<{
    id: string;
    username: string;
    name: string;
    email: string;
    department: string;
    color: string;
    role: 'editor' | 'viewer';
  }>({
    id: 'usr_denizsara',
    username: 'denizsara',
    name: 'Deniz Sara',
    email: 'denizsara@kurumsal.local',
    department: 'Yönetim & Strateji',
    color: '#059669',
    role: 'editor',
  });

  const [authForm, setAuthForm] = useState({
    username: '',
    password: '',
    name: '',
    department: '',
  });

  // Giriş yapan kullanıcının geçmişte girdiği veya oluşturduğu tüm Excel Odaları + PIN'leri
  const [userSavedRooms, setUserSavedRooms] = useState<UserSavedRoomItem[]>([
    {
      roomCode: 'XLS-849-201',
      title: '2026_Kurumsal_Butce_Plani',
      pin: '',
      role: 'owner',
      lastAccessedAt: new Date().toISOString(),
    },
  ]);

  // Yeni Excel Odası Oluşturma Formu State'i
  const [newRoomForm, setNewRoomForm] = useState({
    title: '',
    pin: '',
    template: 'blank' as 'blank' | 'budget' | 'crm',
  });
  const [revealedPins, setRevealedPins] = useState<Record<string, boolean>>({});

  // 2) ÇOKLU ÇALIŞMA KİTABI SEKME SİSTEMİ (MULTI-WORKBOOK TABS)
  const [openWorkbooks, setOpenWorkbooks] = useState<OpenWorkbookTab[]>(() => {
    const wb1 = getDefaultEnterpriseWorkbook();
    const wb2 = getSalesCrmWorkbook();
    return [
      {
        id: wb1.id,
        roomCode: 'XLS-849-201',
        name: wb1.name,
        workbookData: wb1,
        updatedAt: Date.now(),
      },
      {
        id: wb2.id,
        roomCode: 'XLS-702-418',
        name: wb2.name,
        workbookData: wb2,
        updatedAt: Date.now() - 30000,
      },
    ];
  });
  const [activeWorkbookTabId, setActiveWorkbookTabId] = useState<string>('wb_default_enterprise');

  const [workbookData, setWorkbookData] = useState<UniverWorkbookData>(() =>
    getDefaultEnterpriseWorkbook()
  );

  const [roomCode, setRoomCode] = useState<string>('XLS-849-201');
  const [roomPin, setRoomPin] = useState<string>('');
  const [defaultRoomRole, setDefaultRoomRole] = useState<'editor' | 'viewer'>('editor');
  const [protectedColumns, setProtectedColumns] = useState<string[]>(['A']);
  const [roomVersion, setRoomVersion] = useState<number>(1);
  const [isRoomOwner, setIsRoomOwner] = useState<boolean>(true);

  // Hareketsizlik (Idle) sonrası hücre kilidinin otomatik kalkma süresi (Varsayılan: 30 sn)
  const [lockTimeoutSec, setLockTimeoutSec] = useState<number>(30);
  const [nowTick, setNowTick] = useState<number>(Date.now());

  const [stats, setStats] = useState({
    totalRows: 7,
    totalCells: 56,
    sheetCount: 2,
    compressedKB: 4,
    lastSavedTime: new Date().toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit' }),
    lastSyncText: 'Otomatik Kaydedildi',
    zeroUploadVerified: true,
  });

  const [collaborators, setCollaborators] = useState<Record<string, CollaboratorPresence>>({});
  const [followingUserId, setFollowingUserId] = useState<string | null>(null);
  const [activeSelection, setActiveSelection] = useState({
    sheetId: 'sheet_budget',
    sheetName: '2026_Butce_Tablosu',
    cellRef: 'G2',
    row: 1,
    col: 6,
    value: 1580000 as any,
    formula: '=SUM(D2:F2)' as string | undefined,
  });

  // 3) VERSİYON GEÇMİŞİ VE SON DEĞİŞİKLİKLER (HER ZAMAN EN YENİ EN ÜSTTE VE SADECE AKTİF ODAYA AİT!)
  const initialTimestampRef = useRef<number>(Date.now() - 120000);
  const [auditLog, setAuditLog] = useState<AuditEvent[]>([
    {
      id: 'init_1',
      roomCode: 'XLS-849-201',
      userId: 'usr_deniz',
      userName: 'Deniz Sara',
      userEmail: 'deniz.sara@sirket.com.tr',
      userDepartment: 'Yönetim & Strateji',
      userColor: '#059669',
      action: 'Çalışma kitabını başlattı (İlk Versiyon)',
      timestamp: initialTimestampRef.current,
      snapshotId: 'snap_initial',
    },
  ]);

  const [snapshots, setSnapshots] = useState<SnapshotMeta[]>([
    {
      id: 'snap_initial',
      roomCode: 'XLS-849-201',
      label: 'İlk Başlangıç Versiyonu (v1)',
      createdBy: 'Deniz Sara',
      createdByDepartment: 'Yönetim & Strateji',
      createdAt: new Date(initialTimestampRef.current).toISOString(),
      timestamp: initialTimestampRef.current,
      rowCount: 7,
    },
  ]);

  const [localSnapshotsMap, setLocalSnapshotsMap] = useState<Record<string, UniverWorkbookData>>(
    () => ({
      snap_initial: JSON.parse(JSON.stringify(getDefaultEnterpriseWorkbook())),
    })
  );

  // 4) FX FONKSİYON SİHİRBAZI & EN SON KULLANILAN FONKSİYONLAR STATE'İ
  const [recentFormulas, setRecentFormulas] = useState<string[]>([
    'XLOOKUP',
    'SUM',
    'IF',
    'VLOOKUP',
    'AVERAGE',
    'COUNTIF',
  ]);
  const [fxCategory, setFxCategory] = useState<
    'recent' | 'all' | 'lookup' | 'math' | 'logical' | 'stats' | 'text'
  >('recent');
  const [fxSearchQuery, setFxSearchQuery] = useState<string>('');
  const [selectedFormulaName, setSelectedFormulaName] = useState<string>('XLOOKUP');
  const [fxArgValues, setFxArgValues] = useState<Record<string, string>>({
    lookup_value: 'A2',
    lookup_array: 'A2:A6',
    return_array: 'G2:G6',
    if_not_found: '"Bulunamadı"',
    match_mode: '0',
  });

  // 5) PİVOT TABLO OLUŞTURUCU STATE'İ
  const [pivotConfig, setPivotConfig] = useState<PivotConfig>({
    sourceSheetId: 'sheet_budget',
    rowColIdx: 1, // BİRİM / DEPARTMAN
    colColIdx: null, // Tek boyutlu veya çapraz
    valColIdx: 6, // YILLIK TOPLAM
    aggFunc: 'SUM',
    filterColIdx: null,
    filterValue: 'ALL',
  });

  // 6) KOŞULLU BİÇİMLENDİRME STATE'İ
  const [cfColIndex, setCfColIndex] = useState<number>(6); // G Sütunu (YILLIK TOPLAM)
  const [cfOperator, setCfOperator] = useState<ConditionalRuleOperator>('gt');
  const [cfValue1, setCfValue1] = useState<string>('1400000');
  const [cfValue2, setCfValue2] = useState<string>('2000000');
  const [cfColorPreset, setCfColorPreset] = useState<ConditionalColorPreset>('green');
  const [activeCfRules, setActiveCfRules] = useState<ConditionalRuleItem[]>([]);

  const [activeDrawer, setActiveDrawer] = useState<
    'none' | 'pivot' | 'history' | 'protection' | 'conditional'
  >('none');

  const [openMenu, setOpenMenu] = useState<'none' | 'file' | 'tools'>('none');
  const [personalFilterMode, setPersonalFilterMode] = useState<boolean>(true);
  const [simulateBotCollab, setSimulateBotCollab] = useState<boolean>(false);

  const [activeModal, setActiveModal] = useState<
    | null
    | 'fx'
    | 'pivot_modal'
    | 'clipboard'
    | 'share'
    | 'join'
    | 'macro'
    | 'save_info'
    | 'excel_bridge'
    | 'my_rooms'
  >(null);

  // 7) CANLI MASAÜSTÜ EXCEL VLOOKUP / XLOOKUP KÖPRÜSÜ STATE'İ
  const [bridgeTab, setBridgeTab] = useState<'desktop_to_web' | 'web_to_desktop'>('desktop_to_web');
  const [linkedDesktopExcel, setLinkedDesktopExcel] = useState<LinkedDesktopExcelInfo | null>(null);
  const [bridgeFormulaType, setBridgeFormulaType] = useState<'XLOOKUP' | 'VLOOKUP'>('XLOOKUP');
  const [bridgeSourceKeyCol, setBridgeSourceKeyCol] = useState<number>(0);
  const [bridgeExtKeyCol, setBridgeExtKeyCol] = useState<number>(0);
  const [bridgeExtReturnCol, setBridgeExtReturnCol] = useState<number>(2);
  const [bridgeTargetMode, setBridgeTargetMode] = useState<'NEW_COL' | 'ACTIVE_CELL'>('NEW_COL');
  const [bridgeCustomColHeader, setBridgeCustomColHeader] = useState<string>('');
  const [copiedBridgeUrl, setCopiedBridgeUrl] = useState<string | null>(null);

  const [importProgress, setImportProgress] = useState<{
    active: boolean;
    percent: number;
    text: string;
  }>({ active: false, percent: 0, text: '' });

  const [clipboardText, setClipboardText] = useState<string>('');
  const [joinCodeInput, setJoinCodeInput] = useState<string>('');
  const [joinPinInput, setJoinPinInput] = useState<string>('');
  const [savedRoomsList, setSavedRoomsList] = useState<any[]>([]);
  const [statusToast, setStatusToast] = useState<string | null>(null);
  const [copiedLink, setCopiedLink] = useState<boolean>(false);
  const [newSnapshotLabel, setNewSnapshotLabel] = useState<string>('');
  const [isDraggingOver, setIsDraggingOver] = useState<boolean>(false);

  const [macroCode, setMacroCode] = useState<string>(
    `// Aktif sayfadaki Q1-Q3 bütçe hücrelerini %10 artırır:
for (let r = 1; r <= 5; r++) {
  for (let c = 3; c <= 5; c++) {
    const cell = sheet.getRange(r, c, 1, 1);
    const val = Number(cell.getValue() || 0);
    if (val > 0) {
      cell.setValue(Math.round(val * 1.10));
    }
  }
}`
  );

  const sheetHandleRef = useRef<UniverSpreadsheetHandle | null>(null);
  const roomClientRef = useRef<RealtimeRoomClient | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const bridgeFileInputRef = useRef<HTMLInputElement | null>(null);
  const desktopFileHandleRef = useRef<any>(null);
  const lastDesktopFileModifiedRef = useRef<number>(0);
  const roomVersionRef = useRef<number>(1);
  const lastActivitySentRef = useRef<number>(0);
  const isHydratedFromStorageRef = useRef<boolean>(false);
  const latestRestoredWorkbookRef = useRef<UniverWorkbookData | null>(null);

  const WORKSPACE_STORAGE_KEY = 'syncsheet_workspace_state_v2';

  const persistWorkspaceToBrowser = (
    overrideWb?: UniverWorkbookData,
    overrideTabs?: OpenWorkbookTab[],
    overrideActiveTabId?: string,
    overrideRoomCode?: string
  ) => {
    if (typeof window === 'undefined' || !isHydratedFromStorageRef.current) return;
    try {
      const wbToSave = overrideWb || sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
      // Eğer tablo çok büyükse (>2.5MB) tarayıcı localStorage'ını dondurmamak için yalnız sunucuya (Supabase) bırak
      const serializedWb = JSON.stringify(wbToSave);
      if (serializedWb.length > 2500000) return;

      const tabIdToSave = overrideActiveTabId || activeWorkbookTabId;
      const codeToSave = overrideRoomCode || roomCode;
      const tabsToSave = (overrideTabs || openWorkbooks)
        .slice(0, 8)
        .map((t) =>
          t.id === tabIdToSave
            ? { ...t, roomCode: codeToSave, name: wbToSave.name, workbookData: wbToSave }
            : t
        );
      const payload = {
        savedAt: Date.now(),
        roomCode: codeToSave,
        activeWorkbookTabId: tabIdToSave,
        workbookData: wbToSave,
        openWorkbooks: tabsToSave,
        snapshots: snapshots.slice(0, 15),
        auditLog: auditLog.slice(0, 30),
      };
      localStorage.setItem(WORKSPACE_STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
      // 100K satır gibi çok büyük tablolarda localStorage kotası aşılırsa sunucu (Supabase) devrededir
    }
  };

  useEffect(() => {
    roomVersionRef.current = roomVersion;
  }, [roomVersion]);

  useEffect(() => {
    if (Object.keys(collaborators).length === 0) return;
    const timer = setInterval(() => {
      setNowTick(Date.now());
    }, 1000);
    return () => clearInterval(timer);
  }, [collaborators]);

  const showToast = (msg: string) => {
    setStatusToast(msg);
    setTimeout(() => {
      setStatusToast((prev) => (prev === msg ? null : prev));
    }, 3800);
  };

  // Kullanıcının girdiği veya oluşturduğu odayı + PIN bilgisini kendi hesabına kaydet
  const recordRoomInUserHistory = async (
    roomCodeToSave: string,
    titleToSave: string,
    pinToSave?: string,
    roleToSave: 'owner' | 'editor' | 'viewer' = 'editor',
    overrideUsername?: string
  ) => {
    const cleanCode = roomCodeToSave.trim().toUpperCase();
    const uname = (overrideUsername || currentUser.username || '').trim().toLowerCase();
    if (!cleanCode) return;

    const nowIso = new Date().toISOString();
    setUserSavedRooms((prev) => {
      const existing = prev.find((r) => r.roomCode === cleanCode);
      const updatedItem: UserSavedRoomItem = {
        roomCode: cleanCode,
        title: titleToSave || existing?.title || `Excel_${cleanCode}`,
        pin: pinToSave !== undefined ? pinToSave : existing?.pin || '',
        role: existing?.role === 'owner' ? 'owner' : roleToSave,
        lastAccessedAt: nowIso,
      };
      const rest = prev.filter((r) => r.roomCode !== cleanCode);
      const nextList = [updatedItem, ...rest];
      if (typeof window !== 'undefined') {
        try {
          const rawAuth = localStorage.getItem(AUTH_STORAGE_KEY);
          if (rawAuth) {
            const parsed = JSON.parse(rawAuth);
            parsed.savedRooms = nextList;
            localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(parsed));
          }
        } catch (e) {
          // ignore
        }
      }
      return nextList;
    });

    if (!uname) return;
    try {
      const res = await fetch('/api/auth', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'UPSERT_USER_ROOM',
          username: uname,
          roomCode: cleanCode,
          title: titleToSave,
          pin: pinToSave,
          role: roleToSave,
        }),
      });
      const data = await res.json();
      if (res.ok && Array.isArray(data.savedRooms)) {
        setUserSavedRooms(data.savedRooms);
      }
    } catch (e) {
      // ignore
    }
  };

  const handleRemoveRoomFromUserHistory = async (roomCodeToRemove: string) => {
    const cleanCode = roomCodeToRemove.trim().toUpperCase();
    setUserSavedRooms((prev) => prev.filter((r) => r.roomCode !== cleanCode));
    if (!currentUser.username) return;
    try {
      const res = await fetch('/api/auth', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'REMOVE_USER_ROOM',
          username: currentUser.username,
          roomCode: cleanCode,
        }),
      });
      const data = await res.json();
      if (res.ok && Array.isArray(data.savedRooms)) {
        setUserSavedRooms(data.savedRooms);
      }
      showToast(`🗑️ ${cleanCode} odası geçmiş listenizden kaldırıldı.`);
    } catch (e) {
      // ignore
    }
  };

  // KULLANICI GİRİŞ YAP / KAYIT OL İŞLEMİ
  const handleAuthSubmit = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    setAuthError(null);

    const cleanUsername = authForm.username.trim().toLowerCase();
    const cleanPassword = authForm.password.trim();

    if (!cleanUsername || cleanUsername.length < 3) {
      setAuthError('Lütfen en az 3 karakterli bir kullanıcı adı girin.');
      return;
    }
    if (!cleanPassword || cleanPassword.length < 3) {
      setAuthError('Lütfen en az 3 karakterli bir şifre girin.');
      return;
    }

    setAuthLoading(true);
    try {
      const res = await fetch('/api/auth', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: authMode === 'register' ? 'REGISTER' : 'LOGIN',
          username: cleanUsername,
          password: cleanPassword,
          name: authForm.name.trim() || cleanUsername,
          department: authForm.department.trim() || 'Genel Müdürlük',
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        setAuthError(data.error || 'Giriş yapılamadı.');
        setAuthLoading(false);
        return;
      }

      const loggedUser = data.user;
      setCurrentUser({
        id: loggedUser.id,
        username: loggedUser.username,
        name: loggedUser.name,
        email: loggedUser.email,
        department: loggedUser.department,
        color: loggedUser.color,
        role: 'editor',
      });
      const userRooms: UserSavedRoomItem[] = Array.isArray(loggedUser.savedRooms)
        ? loggedUser.savedRooms
        : [];
      setUserSavedRooms(userRooms);
      localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(loggedUser));
      setIsAuthenticated(true);
      setAuthLoading(false);

      // Eğer URL'de paylaşılan bir oda kodu (?room=XLS-...) varsa doğrudan o odayı aç, yoksa son odayı yükle
      const urlParams =
        typeof window !== 'undefined' ? new URLSearchParams(window.location.search) : null;
      const sharedRoomParam = urlParams?.get('room')?.trim().toUpperCase();
      const sharedPinParam = urlParams?.get('pin')?.trim() || '';

      if (sharedRoomParam) {
        const savedPin =
          sharedPinParam ||
          userRooms.find((r) => r.roomCode === sharedRoomParam)?.pin ||
          '';
        setRoomPin(savedPin);
        await loadRoomFromDatabase(sharedRoomParam, savedPin, false, 0);
      } else if (userRooms.length > 0) {
        const lastRoom = userRooms[0];
        setRoomPin(lastRoom.pin || '');
        await loadRoomFromDatabase(lastRoom.roomCode, lastRoom.pin || '', true, 0);
      }

      showToast(
        authMode === 'register'
          ? `🎉 Hesabınız oluşturuldu! Hoş geldiniz, ${loggedUser.name}`
          : `👋 Tekrar hoş geldiniz, ${loggedUser.name} (${userRooms.length} kayıtlı Excel odanız yüklendi)`
      );
    } catch (err) {
      setAuthError('Sunucu ile bağlantı kurulamadı.');
      setAuthLoading(false);
    }
  };

  const handleLogout = () => {
    if (typeof window !== 'undefined') {
      localStorage.removeItem(AUTH_STORAGE_KEY);
    }
    setIsAuthenticated(false);
    setAuthForm({ username: currentUser.username, password: '', name: '', department: '' });
    setAuthError(null);
  };

  // Başlangıçta kullanıcı oturumunu, geçmiş odalarını/PIN'lerini ve KAYITLI TABLOYU yükle
  useEffect(() => {
    if (typeof window === 'undefined') return;

    setRecentFormulas(getRecentFormulaNames());

    let restoredUsername = '';
    let restoredUserRooms: UserSavedRoomItem[] = [];

    const savedAuth = localStorage.getItem(AUTH_STORAGE_KEY);
    if (savedAuth) {
      try {
        const parsed = JSON.parse(savedAuth);
        if (parsed && parsed.username) {
          restoredUsername = parsed.username;
          setCurrentUser({
            id: parsed.id || `usr_${parsed.username}`,
            username: parsed.username,
            name: parsed.name || parsed.username,
            email: parsed.email || `${parsed.username}@kurumsal.local`,
            department: parsed.department || 'Genel Müdürlük',
            color: parsed.color || '#059669',
            role: parsed.role || 'editor',
          });
          if (Array.isArray(parsed.savedRooms) && parsed.savedRooms.length > 0) {
            restoredUserRooms = parsed.savedRooms;
            setUserSavedRooms(parsed.savedRooms);
          }
          setIsAuthenticated(true);

          // Sunucudan kullanıcının güncel oda & PIN geçmişini çek
          fetch(`/api/auth?username=${encodeURIComponent(parsed.username)}`)
            .then((r) => r.json())
            .then((d) => {
              if (d?.user?.savedRooms) {
                setUserSavedRooms(d.user.savedRooms);
              }
            })
            .catch(() => {});
        }
      } catch (e) {
        // ignore
      }
    }

    const params = new URLSearchParams(window.location.search);
    const roomParam = params.get('room');
    const pinParam = params.get('pin');
    const roleParam = params.get('role');
    if (roleParam === 'viewer') {
      setCurrentUser((prev) => ({ ...prev, role: 'viewer' }));
      setIsRoomOwner(false);
    }

    let localSavedAt = 0;
    let targetRoom = roomParam
      ? roomParam.toUpperCase()
      : restoredUserRooms[0]?.roomCode || 'XLS-849-201';
    let targetPin =
      pinParam ||
      restoredUserRooms.find((r) => r.roomCode === targetRoom)?.pin ||
      '';

    // 1) Önce tarayıcı hafızasındaki (localStorage) son kaydedilen çalışma kitaplarını ve tabloyu geri yükle
    if (!roomParam) {
      try {
        const rawWorkspace = localStorage.getItem(WORKSPACE_STORAGE_KEY);
        if (rawWorkspace) {
          const parsedWs = JSON.parse(rawWorkspace);
          if (parsedWs && parsedWs.workbookData && parsedWs.workbookData.sheets) {
            localSavedAt = parsedWs.savedAt || Date.now();
            if (parsedWs.roomCode) {
              targetRoom = parsedWs.roomCode;
              targetPin =
                restoredUserRooms.find((r) => r.roomCode === targetRoom)?.pin || '';
            }
            if (Array.isArray(parsedWs.openWorkbooks) && parsedWs.openWorkbooks.length > 0) {
              setOpenWorkbooks(parsedWs.openWorkbooks);
            }
            if (parsedWs.activeWorkbookTabId) {
              setActiveWorkbookTabId(parsedWs.activeWorkbookTabId);
            }
            if (Array.isArray(parsedWs.snapshots) && parsedWs.snapshots.length > 0) {
              setSnapshots(parsedWs.snapshots);
            }
            if (Array.isArray(parsedWs.auditLog) && parsedWs.auditLog.length > 0) {
              setAuditLog(parsedWs.auditLog);
            }
            const clonedWb = JSON.parse(JSON.stringify(parsedWs.workbookData));
            latestRestoredWorkbookRef.current = clonedWb;
            setWorkbookData(clonedWb);
            sheetHandleRef.current?.loadWorkbookSnapshot(clonedWb);
          }
        }
      } catch (e) {
        // ignore
      }
    }

    isHydratedFromStorageRef.current = true;
    setRoomCode(targetRoom);
    if (targetPin) setRoomPin(targetPin);

    // 2) Ardından sunucu veritabanındaki (.data/rooms.json) kayıtlı odayı kontrol et
    loadRoomFromDatabase(targetRoom, targetPin, !roomParam, localSavedAt);
  }, []);

  // Tablo, sekmeler veya versiyonlar değiştikçe tarayıcı hafızasını otomatik güncelle
  useEffect(() => {
    if (!isHydratedFromStorageRef.current) return;
    persistWorkspaceToBrowser(workbookData, openWorkbooks, activeWorkbookTabId, roomCode);
  }, [roomCode, workbookData, openWorkbooks, activeWorkbookTabId, snapshots, auditLog]);

  // Seçilen fonksiyon değiştiğinde varsayılan örnek argümanları akıllıca doldur
  const handleSelectFormulaInWizard = (fnName: string) => {
    setSelectedFormulaName(fnName);
    const cols = getSheetColumnRanges(workbookData, activeSelection.sheetId);
    const firstCol = cols[0]?.rangeRef || 'A2:A6';
    const firstCell = cols[0]?.sampleCellRef || 'A2';
    const lastNumCol = cols[cols.length - 2]?.rangeRef || 'G2:G6';

    if (fnName === 'XLOOKUP') {
      setFxArgValues({
        lookup_value: firstCell,
        lookup_array: firstCol,
        return_array: lastNumCol,
        if_not_found: '"Bulunamadı"',
        match_mode: '0',
      });
    } else if (fnName === 'VLOOKUP') {
      setFxArgValues({
        lookup_value: firstCell,
        table_array: `A2:${cols[cols.length - 1]?.colLetter || 'H'}6`,
        col_index_num: String(Math.max(2, cols.length - 1)),
        range_lookup: 'FALSE',
      });
    } else if (fnName === 'SUM' || fnName === 'AVERAGE' || fnName === 'MAX' || fnName === 'MIN') {
      setFxArgValues({
        range1: 'D2:F2',
      });
    } else if (fnName === 'IF') {
      setFxArgValues({
        logical_test: 'G2>1500000',
        value_if_true: '"Yüksek Bütçe"',
        value_if_false: '"Standart Bütçe"',
      });
    } else if (fnName === 'COUNTIF') {
      setFxArgValues({
        range: 'H2:H6',
        criteria: '"Aktif Onaylı"',
      });
    } else if (fnName === 'SUMIF') {
      setFxArgValues({
        range: 'H2:H6',
        criteria: '"Aktif Onaylı"',
        sum_range: 'G2:G6',
      });
    } else {
      setFxArgValues({});
    }
  };

  // Çoklu Çalışma Kitabı Sekmesi Değiştirme (Mevcut kitabı kaybetmeden geçiş yapar ve o kitaba ait odayı aktif eder)
  const handleSwitchWorkbookTab = (targetTabId: string) => {
    if (targetTabId === activeWorkbookTabId) return;

    // Önce mevcut açık çalışma kitabının son halini kaydet
    const currentLiveSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    const updatedList = openWorkbooks.map((tab) =>
      tab.id === activeWorkbookTabId
        ? {
            ...tab,
            roomCode: tab.roomCode || roomCode,
            name: currentLiveSnap.name || tab.name,
            workbookData: JSON.parse(JSON.stringify(currentLiveSnap)),
            updatedAt: Date.now(),
          }
        : tab
    );

    const targetTab = updatedList.find((t) => t.id === targetTabId);
    if (!targetTab) return;

    setOpenWorkbooks(updatedList);
    setActiveWorkbookTabId(targetTab.id);

    const targetRoomCode =
      targetTab.roomCode ||
      (targetTab.id === 'wb_default_enterprise'
        ? 'XLS-849-201'
        : targetTab.id === 'wb_sales_crm'
          ? 'XLS-702-418'
          : targetTab.id.startsWith('wb_XLS-')
            ? targetTab.id.replace(/^wb_/, '')
            : roomCode);
    if (targetRoomCode && targetRoomCode !== roomCode) {
      setRoomCode(targetRoomCode);
    }

    const clonedTarget = JSON.parse(JSON.stringify(targetTab.workbookData));
    latestRestoredWorkbookRef.current = clonedTarget;
    setWorkbookData(clonedTarget);
    sheetHandleRef.current?.loadWorkbookSnapshot(clonedTarget);

    const firstSheetId =
      clonedTarget.sheetOrder?.[0] || Object.keys(clonedTarget.sheets || {})[0] || 'sheet_1';
    setPivotConfig((prev) => ({
      ...prev,
      sourceSheetId: firstSheetId,
      rowColIdx: 1,
      valColIdx: 5,
    }));

    showToast(`📂 "${targetTab.name}" (${targetRoomCode}) çalışma kitabına geçildi.`);
  };

  const handleCreateNewWorkbookTab = () => {
    const currentLiveSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    const nextIdx = openWorkbooks.length + 1;
    const newWb = getBlankWorkbook(nextIdx);
    const r1 = Math.floor(100 + Math.random() * 900);
    const r2 = Math.floor(100 + Math.random() * 900);
    const newTabRoomCode = `XLS-${r1}-${r2}`;
    const newTab: OpenWorkbookTab = {
      id: newWb.id,
      roomCode: newTabRoomCode,
      name: newWb.name,
      workbookData: newWb,
      updatedAt: Date.now(),
    };

    setOpenWorkbooks((prev) => [
      ...prev.map((t) =>
        t.id === activeWorkbookTabId
          ? {
              ...t,
              roomCode: t.roomCode || roomCode,
              workbookData: JSON.parse(JSON.stringify(currentLiveSnap)),
            }
          : t
      ),
      newTab,
    ]);
    setActiveWorkbookTabId(newTab.id);
    setRoomCode(newTabRoomCode);
    latestRestoredWorkbookRef.current = newWb;
    setWorkbookData(newWb);
    sheetHandleRef.current?.loadWorkbookSnapshot(newWb);
    showToast(`✨ "${newWb.name}" (${newTabRoomCode}) oluşturuldu ve açıldı.`);
  };

  const handleCloseWorkbookTab = (tabIdToClose: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (openWorkbooks.length <= 1) {
      showToast('En az bir çalışma kitabı açık kalmalıdır.');
      return;
    }
    const remaining = openWorkbooks.filter((t) => t.id !== tabIdToClose);
    setOpenWorkbooks(remaining);
    if (activeWorkbookTabId === tabIdToClose) {
      const nextActive = remaining[0];
      setActiveWorkbookTabId(nextActive.id);
      if (nextActive.roomCode) {
        setRoomCode(nextActive.roomCode);
      }
      const cloned = JSON.parse(JSON.stringify(nextActive.workbookData));
      latestRestoredWorkbookRef.current = cloned;
      setWorkbookData(cloned);
      sheetHandleRef.current?.loadWorkbookSnapshot(cloned);
    }
  };

  useEffect(() => {
    const client = new RealtimeRoomClient(roomCode, (msg) => {
      if (msg.type === 'PRESENCE_UPDATE') {
        if (msg.presence.userId === currentUser.id) return;
        setCollaborators((prev) => ({
          ...prev,
          [msg.presence.userId]: msg.presence,
        }));

        if (followingUserId === msg.presence.userId) {
          sheetHandleRef.current?.navigateToCell(msg.presence.row, msg.presence.col);
        }
      } else if (msg.type === 'CELL_UPDATE') {
        if (msg.senderId === currentUser.id) return;
        sheetHandleRef.current?.applyRemoteCellUpdate(
          msg.sheetId,
          msg.row,
          msg.col,
          msg.value,
          msg.formula
        );
        const refStr = formatCellRef(msg.row, msg.col);
        const nowTs = Date.now();
        const snapId = `snap_remote_${nowTs}`;
        const updatedSnap = sheetHandleRef.current?.getCurrentSnapshot();
        if (updatedSnap) {
          setLocalSnapshotsMap((prev) => ({
            ...prev,
            [snapId]: JSON.parse(JSON.stringify(updatedSnap)),
          }));
        }

        setStats((prev) => ({
          ...prev,
          lastSavedTime: new Date(nowTs).toLocaleTimeString('tr-TR', {
            hour: '2-digit',
            minute: '2-digit',
          }),
          lastSyncText: `${msg.senderName} ${refStr} hücresini güncelledi`,
        }));
        setAuditLog((prev) =>
          [
            {
              id: `aud_${nowTs}`,
              roomCode,
              userId: msg.senderId,
              userName: msg.senderName,
              userColor: '#2563eb',
              action: `${refStr} hücresini güncelledi`,
              cellRef: refStr,
              newValue: msg.formula || msg.value,
              timestamp: nowTs,
              snapshotId: snapId,
            },
            ...prev,
          ].sort((a, b) => b.timestamp - a.timestamp)
        );
      } else if (msg.type === 'WORKBOOK_SYNC') {
        if (msg.senderId === currentUser.id) return;
        const cloned = JSON.parse(JSON.stringify(msg.workbookData));
        setWorkbookData(cloned);
        sheetHandleRef.current?.loadWorkbookSnapshot(cloned);
        showToast(`${msg.senderName} tabloyu güncelledi`);
      }
    });

    roomClientRef.current = client;

    const pollInterval = setInterval(async () => {
      try {
        const res = await fetch(
          `/api/rooms?code=${encodeURIComponent(roomCode)}&pin=${encodeURIComponent(roomPin)}&mode=poll&sinceVersion=${roomVersionRef.current}`
        );
        if (!res.ok) return;
        const data = await res.json();

        if (data.deltas && data.deltas.length > 0) {
          data.deltas.forEach((d: any) => {
            if (d.senderId !== currentUser.id) {
              sheetHandleRef.current?.applyRemoteCellUpdate(
                d.sheetId,
                d.row,
                d.col,
                d.value,
                d.formula
              );
            }
          });
          setRoomVersion(data.version);
        }

        if (data.presenceMap) {
          setCollaborators(() => {
            const next: Record<string, CollaboratorPresence> = {};
            Object.values(data.presenceMap).forEach((p: any) => {
              if (
                p.userId &&
                p.userId !== currentUser.id &&
                !String(p.userId).startsWith('bot_')
              ) {
                next[p.userId] = p;
              }
            });
            return next;
          });
        }
      } catch (e) {
        // ignore
      }
    }, 3000);

    return () => {
      client.destroy();
      clearInterval(pollInterval);
    };
  }, [roomCode, currentUser.id, followingUserId]);

  // Bot simülasyonu: Selin Yılmaz D2 hücresinde işlem yapıyor
  const botStartRef = useRef<number>(Date.now());
  useEffect(() => {
    if (!simulateBotCollab) {
      setCollaborators((prev) => {
        const copy = { ...prev };
        delete copy['bot_selin'];
        delete copy['bot_mert'];
        return copy;
      });
      return;
    }

    botStartRef.current = Date.now();
    setCollaborators((prev) => ({
      ...prev,
      bot_selin: {
        userId: 'bot_selin',
        userName: 'Selin Yılmaz',
        userColor: '#2563eb',
        role: 'editor',
        activeSheetId: 'sheet_budget',
        activeSheetName: '2026_Butce_Tablosu',
        activeCellRef: 'D2',
        row: 1,
        col: 3,
        lastUpdated: Date.now(),
        lastActivityAt: botStartRef.current,
      },
      bot_mert: {
        userId: 'bot_mert',
        userName: 'Mert Demir',
        userColor: '#7c3aed',
        role: 'viewer',
        activeSheetId: 'sheet_budget',
        activeSheetName: '2026_Butce_Tablosu',
        activeCellRef: 'E3',
        row: 2,
        col: 4,
        lastUpdated: Date.now(),
        lastActivityAt: botStartRef.current,
      },
    }));
  }, [simulateBotCollab]);

  const processExcelFile = async (file: File) => {
    setImportProgress({
      active: true,
      percent: 10,
      text: `"${file.name}" okunuyor...`,
    });

    try {
      const buffer = await file.arrayBuffer();
      setTimeout(async () => {
        try {
          const result = parseExcelBufferToUniverData(buffer, file.name, (percent, statusText) => {
            setImportProgress({ active: true, percent, text: statusText });
          });

          // Yüklenen Excel dosyası için kullanıcıya özel yeni bir oda kodu oluştur ve geçmişine kaydet
          const r1 = Math.floor(100 + Math.random() * 900);
          const r2 = Math.floor(100 + Math.random() * 900);
          const importedRoomCode = `XLS-${r1}-${r2}`;
          result.workbookData.id = `wb_${importedRoomCode}`;

          const firstSheetId =
            result.workbookData.sheetOrder?.[0] ||
            Object.keys(result.workbookData.sheets || {})[0] ||
            'sheet_1';
          const firstSheetName =
            result.workbookData.sheets?.[firstSheetId]?.name || 'Sayfa 1';

          // Mevcut açık kitabı kaydet ve yeni yüklenen Excel'i YENİ SEKME olarak çalışma kitaplarına ekle!
          const currentLiveSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
          const newTab: OpenWorkbookTab = {
            id: result.workbookData.id,
            roomCode: importedRoomCode,
            name: result.workbookData.name,
            workbookData: result.workbookData,
            updatedAt: Date.now(),
          };

          setOpenWorkbooks((prev) => [
            ...prev.map((t) =>
              t.id === activeWorkbookTabId
                ? {
                    ...t,
                    roomCode: t.roomCode || roomCode,
                    workbookData: currentLiveSnap,
                  }
                : t
            ),
            newTab,
          ]);
          setActiveWorkbookTabId(newTab.id);

          setActiveSelection({
            sheetId: firstSheetId,
            sheetName: firstSheetName,
            cellRef: 'A1',
            row: 0,
            col: 0,
            value: '',
            formula: undefined,
          });
          setPivotConfig((prev) => ({
            ...prev,
            sourceSheetId: firstSheetId,
            rowColIdx: 0,
            valColIdx: 1,
          }));

          latestRestoredWorkbookRef.current = result.workbookData;
          setWorkbookData(result.workbookData);
          sheetHandleRef.current?.loadWorkbookSnapshot(result.workbookData);

          const nowTs = Date.now();
          const compKB = Math.max(4, Math.round(result.totalCells * 0.012));
          setStats({
            totalRows: result.totalRows,
            totalCells: result.totalCells,
            sheetCount: result.sheetCount,
            compressedKB: compKB,
            lastSavedTime: new Date(nowTs).toLocaleTimeString('tr-TR', {
              hour: '2-digit',
              minute: '2-digit',
            }),
            lastSyncText: 'Otomatik Kaydedildi',
            zeroUploadVerified: true,
          });

          setRoomCode(importedRoomCode);
          setRoomPin('');
          setIsRoomOwner(true);

          const snapId = `snap_${nowTs}`;
          const newSnapMeta: SnapshotMeta = {
            id: snapId,
            roomCode: importedRoomCode,
            label: `Excel Yüklendi: ${file.name}`,
            createdBy: currentUser.name,
            createdByDepartment: currentUser.department,
            createdAt: new Date(nowTs).toISOString(),
            timestamp: nowTs,
            rowCount: result.totalRows,
          };

          setSnapshots((prev) =>
            [newSnapMeta, ...prev].sort((a, b) => b.timestamp - a.timestamp)
          );
          setLocalSnapshotsMap((prev) => ({
            ...prev,
            [snapId]: result.workbookData,
          }));

          setAuditLog((prev) =>
            [
              {
                id: `aud_${nowTs}`,
                roomCode: importedRoomCode,
                userId: currentUser.id,
                userName: currentUser.name,
                userDepartment: currentUser.department,
                userColor: currentUser.color,
                action: `"${file.name}" dosyasını yeni çalışma kitabı olarak açtı`,
                timestamp: nowTs,
                snapshotId: snapId,
              },
              ...prev,
            ].sort((a, b) => b.timestamp - a.timestamp)
          );

          setImportProgress({
            active: true,
            percent: 95,
            text: `Oda (${importedRoomCode}) veritabanına kaydediliyor...`,
          });

          await fetch('/api/rooms', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              actionType: 'SAVE_WORKBOOK',
              roomCode: importedRoomCode,
              defaultRole: 'editor',
              protectedColumns: [],
              title: result.workbookData.name,
              ownerId: currentUser.id,
              ownerName: currentUser.name,
              ownerEmail: currentUser.email,
              workbookData: result.workbookData,
              stats: {
                totalRows: result.totalRows,
                totalCells: result.totalCells,
                compressedKB: compKB,
              },
            }),
          }).catch(() => {});

          await recordRoomInUserHistory(
            importedRoomCode,
            result.workbookData.name,
            '',
            'owner'
          );

          setImportProgress({ active: false, percent: 0, text: '' });
          setActiveModal(null);
          showToast(
            `✅ "${file.name}" (${importedRoomCode}) odası olarak kaydedildi ve paylaşıma hazır!`
          );
        } catch (innerErr: any) {
          setImportProgress({ active: false, percent: 0, text: '' });
          showToast(`Hata: ${innerErr?.message || 'Excel işlenemedi'}`);
        }
      }, 40);
    } catch (err: any) {
      setImportProgress({ active: false, percent: 0, text: '' });
      showToast(`Hata: ${err?.message || 'Excel okunamadı'}`);
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleLocalExcelFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) await processExcelFile(file);
  };

  const handleDropExcelFile = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDraggingOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file && /\.(xlsx|xls|csv)$/i.test(file.name)) {
      await processExcelFile(file);
    } else if (file) {
      showToast('Lütfen geçerli bir Excel (.xlsx, .xls, .csv) dosyası sürükleyin.');
    }
  };

  const handleClipboardImport = async () => {
    if (!clipboardText.trim()) return;
    const r1 = Math.floor(100 + Math.random() * 900);
    const r2 = Math.floor(100 + Math.random() * 900);
    const clippedRoomCode = `XLS-${r1}-${r2}`;
    const tableTitle = `Panodan_Tablo_${r1}`;

    const res = parseClipboardTextToUniverData(clipboardText, tableTitle);
    res.workbookData.id = `wb_${clippedRoomCode}`;

    const nowTs = Date.now();
    const snapId = `snap_clip_${nowTs}`;
    const compKB = Math.max(2, Math.round(res.totalCells * 0.012));

    const currentLiveSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    const newTab: OpenWorkbookTab = {
      id: res.workbookData.id,
      roomCode: clippedRoomCode,
      name: res.workbookData.name,
      workbookData: res.workbookData,
      updatedAt: nowTs,
    };
    setOpenWorkbooks((prev) => [
      ...prev.map((t) =>
        t.id === activeWorkbookTabId
          ? { ...t, roomCode: t.roomCode || roomCode, workbookData: currentLiveSnap }
          : t
      ),
      newTab,
    ]);
    setActiveWorkbookTabId(newTab.id);

    setActiveSelection({
      sheetId: 'sheet_1',
      sheetName: 'Sayfa 1',
      cellRef: 'A1',
      row: 0,
      col: 0,
      value: '',
      formula: undefined,
    });
    setPivotConfig((prev) => ({
      ...prev,
      sourceSheetId: 'sheet_1',
      rowColIdx: 0,
      valColIdx: 1,
    }));

    setRoomCode(clippedRoomCode);
    setRoomPin('');
    setIsRoomOwner(true);

    latestRestoredWorkbookRef.current = res.workbookData;
    setWorkbookData(res.workbookData);
    sheetHandleRef.current?.loadWorkbookSnapshot(res.workbookData);

    setStats({
      totalRows: res.totalRows,
      totalCells: res.totalCells,
      sheetCount: 1,
      compressedKB: compKB,
      lastSavedTime: new Date(nowTs).toLocaleTimeString('tr-TR', {
        hour: '2-digit',
        minute: '2-digit',
      }),
      lastSyncText: 'Otomatik Kaydedildi',
      zeroUploadVerified: true,
    });

    setLocalSnapshotsMap((prev) => ({
      ...prev,
      [snapId]: res.workbookData,
    }));
    setSnapshots((prev) =>
      [
        {
          id: snapId,
          roomCode: clippedRoomCode,
          label: 'Panodan Aktarılan Tablo',
          createdBy: currentUser.name,
          createdByDepartment: currentUser.department,
          createdAt: new Date(nowTs).toISOString(),
          timestamp: nowTs,
          rowCount: res.totalRows,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );
    setAuditLog((prev) =>
      [
        {
          id: `aud_clip_${nowTs}`,
          roomCode: clippedRoomCode,
          userId: currentUser.id,
          userName: currentUser.name,
          userDepartment: currentUser.department,
          userColor: currentUser.color,
          action: `Panodan ${res.totalRows} satırlık yeni çalışma kitabı (${clippedRoomCode}) oluşturdu`,
          timestamp: nowTs,
          snapshotId: snapId,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setClipboardText('');
    setActiveModal(null);

    await fetch('/api/rooms', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        actionType: 'SAVE_WORKBOOK',
        roomCode: clippedRoomCode,
        defaultRole: 'editor',
        protectedColumns: [],
        title: res.workbookData.name,
        ownerId: currentUser.id,
        ownerName: currentUser.name,
        ownerEmail: currentUser.email,
        workbookData: res.workbookData,
        stats: {
          totalRows: res.totalRows,
          totalCells: res.totalCells,
          compressedKB: compKB,
        },
      }),
    }).catch(() => {});

    await recordRoomInUserHistory(
      clippedRoomCode,
      res.workbookData.name,
      '',
      'owner'
    );

    showToast(
      `✅ ${res.totalRows} satır panodan "${tableTitle}" (${clippedRoomCode}) odasına aktarıldı ve kaydedildi!`
    );
  };

  const handleLoadLargeDataset = (rowCount: number) => {
    setOpenMenu('none');
    showToast(`⚡ ${rowCount.toLocaleString('tr-TR')} satırlık tablo hazırlanıyor...`);
    setTimeout(() => {
      const res = generateLargeDatasetWorkbook(rowCount);
      const newTab: OpenWorkbookTab = {
        id: res.workbookData.id,
        name: res.workbookData.name,
        workbookData: res.workbookData,
        updatedAt: Date.now(),
      };
      setOpenWorkbooks((prev) => [...prev, newTab]);
      setActiveWorkbookTabId(newTab.id);
      setWorkbookData(res.workbookData);
      sheetHandleRef.current?.loadWorkbookSnapshot(res.workbookData);

      setStats({
        totalRows: rowCount,
        totalCells: rowCount * 9,
        sheetCount: 2,
        compressedKB: res.compressedSizeKB,
        lastSavedTime: new Date().toLocaleTimeString('tr-TR', {
          hour: '2-digit',
          minute: '2-digit',
        }),
        lastSyncText: `${rowCount.toLocaleString('tr-TR')} satır hazır (${res.generationTimeMs}ms)`,
        zeroUploadVerified: true,
      });
      showToast(`🚀 ${rowCount.toLocaleString('tr-TR')} satırlık tablo yeni sekmede açıldı!`);
    }, 50);
  };

  const handleExportToExcel = () => {
    setOpenMenu('none');
    const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    exportUniverDataToExcelFile(currentSnap, currentSnap.name || 'SyncSheet_Tablo');
    showToast('📥 Excel (.xlsx) dosyası bilgisayarınıza indirildi.');
  };

  const saveRoomToDatabase = async () => {
    setOpenMenu('none');
    try {
      const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
      const nowTs = Date.now();
      const snapId = `snap_save_${nowTs}`;
      const clonedSnap = JSON.parse(JSON.stringify(currentSnap));
      latestRestoredWorkbookRef.current = clonedSnap;

      // Yerel versiyon önbelleğine de tam kopyasını kaydet
      setLocalSnapshotsMap((prev) => ({
        ...prev,
        [snapId]: clonedSnap,
      }));

      // Açık çalışma kitapları listesini ve tarayıcı hafızasını (localStorage) anında güncelle
      setOpenWorkbooks((prev) => {
        const nextTabs = prev.map((t) =>
          t.id === activeWorkbookTabId
            ? { ...t, name: clonedSnap.name, workbookData: clonedSnap }
            : t
        );
        persistWorkspaceToBrowser(clonedSnap, nextTabs, activeWorkbookTabId);
        return nextTabs;
      });

      const res = await fetch('/api/rooms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          actionType: 'SAVE_WORKBOOK',
          roomCode,
          pin: roomPin || undefined,
          defaultRole: defaultRoomRole,
          protectedColumns,
          title: currentSnap.name,
          ownerId: currentUser.id,
          ownerName: currentUser.name,
          ownerEmail: currentUser.email,
          workbookData: currentSnap,
          stats: {
            totalRows: stats.totalRows,
            totalCells: stats.totalCells,
            compressedKB: stats.compressedKB,
          },
        }),
      });
      const data = await res.json();
      if (res.ok) {
        const nextVer = data.version || roomVersion + 1;
        setRoomVersion(nextVer);
        const timeStr = new Date(nowTs).toLocaleTimeString('tr-TR', {
          hour: '2-digit',
          minute: '2-digit',
        });

        setStats((prev) => ({
          ...prev,
          lastSavedTime: timeStr,
          lastSyncText: `Kaydedildi (${timeStr})`,
        }));

        const newSnap: SnapshotMeta = {
          id: snapId,
          roomCode,
          label: `Kayıt Noktası v${nextVer} (${timeStr})`,
          createdBy: currentUser.name,
          createdByDepartment: currentUser.department,
          createdAt: new Date(nowTs).toISOString(),
          timestamp: nowTs,
          rowCount: stats.totalRows,
        };
        setSnapshots((prev) => [newSnap, ...prev].sort((a, b) => b.timestamp - a.timestamp));

        // Kullanıcının kişisel oda ve PIN geçmişine de kaydet
        recordRoomInUserHistory(
          roomCode,
          currentSnap.name,
          roomPin,
          isRoomOwner ? 'owner' : 'editor'
        );

        showToast(
          `💾 Tablonuz sisteme ve ortak çalışma odasına (${roomCode}) başarıyla kaydedildi!`
        );
      } else {
        showToast(`Hata: ${data.error}`);
      }
    } catch (e: any) {
      showToast('Kayıt sırasında hata oluştu');
    }
  };

  // MANUEL VEYA OTOMATİK VERSİYON OLUŞTURMA (En yeni her zaman en üstte!)
  const handleCreateSnapshot = async () => {
    const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    const nowTs = Date.now();
    const timeStr = new Date(nowTs).toLocaleTimeString('tr-TR', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    });
    const label = newSnapshotLabel.trim() || `Manuel Versiyon (${timeStr})`;
    const snapId = `snap_${nowTs}`;

    const deepClonedSnap = JSON.parse(JSON.stringify(currentSnap));

    const newMeta: SnapshotMeta = {
      id: snapId,
      roomCode,
      label,
      createdBy: currentUser.name,
      createdByDepartment: currentUser.department,
      createdAt: new Date(nowTs).toISOString(),
      timestamp: nowTs,
      rowCount: stats.totalRows,
    };

    setLocalSnapshotsMap((prev) => ({
      ...prev,
      [snapId]: deepClonedSnap,
    }));

    setSnapshots((prev) => [newMeta, ...prev].sort((a, b) => b.timestamp - a.timestamp));

    setAuditLog((prev) =>
      [
        {
          id: `aud_${nowTs}`,
          roomCode,
          userId: currentUser.id,
          userName: currentUser.name,
          userDepartment: currentUser.department,
          userColor: currentUser.color,
          action: `"${label}" adıyla versiyon yedeği aldı`,
          timestamp: nowTs,
          snapshotId: snapId,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setNewSnapshotLabel('');
    showToast(`⏱️ "${label}" versiyonu en üste kaydedildi.`);
  };

  // VERSİYONA GERİ DÖNME (%100 GARANTİLİ DEEP CLONE + UNIVER RE-MOUNT)
  const handleRestoreSnapshotById = async (snapId: string, labelText: string) => {
    const localTarget = localSnapshotsMap[snapId];
    if (localTarget) {
      const deepRestored: UniverWorkbookData = JSON.parse(JSON.stringify(localTarget));
      latestRestoredWorkbookRef.current = deepRestored;
      setWorkbookData(deepRestored);
      sheetHandleRef.current?.loadWorkbookSnapshot(deepRestored);

      // Açık çalışma kitabı sekmesini de güncelle
      setOpenWorkbooks((prev) => {
        const nextTabs = prev.map((t) =>
          t.id === activeWorkbookTabId
            ? { ...t, workbookData: JSON.parse(JSON.stringify(deepRestored)) }
            : t
        );
        persistWorkspaceToBrowser(deepRestored, nextTabs, activeWorkbookTabId);
        return nextTabs;
      });

      roomClientRef.current?.sendWorkbookSync(currentUser.id, currentUser.name, deepRestored);

      const nowTs = Date.now();
      setAuditLog((prev) =>
        [
          {
            id: `aud_restore_${nowTs}`,
            roomCode,
            userId: currentUser.id,
            userName: currentUser.name,
            userDepartment: currentUser.department,
            userColor: currentUser.color,
            action: `⏪ "${labelText}" durumuna geri döndü`,
            timestamp: nowTs,
            snapshotId: snapId,
          },
          ...prev,
        ].sort((a, b) => b.timestamp - a.timestamp)
      );

      showToast(`⏪ "${labelText}" versiyonuna eksiksiz geri dönüldü!`);
      return;
    }

    try {
      const res = await fetch('/api/rooms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          actionType: 'RESTORE_SNAPSHOT',
          roomCode,
          snapshotId: snapId,
          userId: currentUser.id,
          userName: currentUser.name,
        }),
      });
      const data = await res.json();
      if (res.ok && data.room?.workbookData) {
        const deepRestored = JSON.parse(JSON.stringify(data.room.workbookData));
        latestRestoredWorkbookRef.current = deepRestored;
        setWorkbookData(deepRestored);
        sheetHandleRef.current?.loadWorkbookSnapshot(deepRestored);
        roomClientRef.current?.sendWorkbookSync(currentUser.id, currentUser.name, deepRestored);
        showToast(`⏪ "${labelText}" versiyonuna geri dönüldü!`);
      } else {
        showToast('Bu versiyonun yerel verisi bulunamadı.');
      }
    } catch (e) {
      showToast('Versiyon geri yüklenemedi.');
    }
  };

  // FX SİHİRBAZINDAN HÜCREYE FORMÜL UYGULAMA
  const handleInsertFormulaToActiveCell = () => {
    const preview = evaluateFormulaPreview(
      selectedFormulaName,
      fxArgValues,
      sheetHandleRef.current?.getCurrentSnapshot() || workbookData,
      activeSelection.sheetId
    );

    const updatedRecents = recordFormulaUsage(selectedFormulaName);
    setRecentFormulas(updatedRecents);

    sheetHandleRef.current?.setFormulaOrValueOnCell(
      activeSelection.sheetId,
      activeSelection.row,
      activeSelection.col,
      preview.formulaString,
      preview.computedValue
    );

    const nowTs = Date.now();
    const snapId = `snap_fx_${nowTs}`;
    const latestWb = sheetHandleRef.current?.getCurrentSnapshot();
    if (latestWb) {
      const clonedWb = JSON.parse(JSON.stringify(latestWb));
      latestRestoredWorkbookRef.current = clonedWb;
      setLocalSnapshotsMap((prev) => ({
        ...prev,
        [snapId]: clonedWb,
      }));
      persistWorkspaceToBrowser(clonedWb);
    }

    setSnapshots((prev) =>
      [
        {
          id: snapId,
          roomCode,
          label: `${activeSelection.cellRef} hücresine ${selectedFormulaName} formülü eklendi`,
          createdBy: currentUser.name,
          createdByDepartment: currentUser.department,
          createdAt: new Date(nowTs).toISOString(),
          timestamp: nowTs,
          rowCount: stats.totalRows,
          isAutoChange: true,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setAuditLog((prev) =>
      [
        {
          id: `aud_fx_${nowTs}`,
          roomCode,
          userId: currentUser.id,
          userName: currentUser.name,
          userDepartment: currentUser.department,
          userColor: currentUser.color,
          action: `${activeSelection.cellRef} hücresine ${selectedFormulaName} formülü ekledi`,
          cellRef: activeSelection.cellRef,
          sheetName: activeSelection.sheetName,
          newValue: `${preview.formulaString} ➔ ${preview.computedValue}`,
          timestamp: nowTs,
          snapshotId: snapId,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setActiveModal(null);
    showToast(
      `✅ ${activeSelection.cellRef} hücresine ${preview.formulaString} formülü uygulandı!`
    );
  };

  // PİVOT TABLOYU YENİ SEKME (SHEET) OLARAK EKLEME
  const handleCreatePivotSheet = () => {
    const liveWb = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    const pivotRes = computePivotTable(liveWb, pivotConfig);
    const { updatedWorkbook, newSheetName } = insertPivotSheetIntoWorkbook(
      liveWb,
      pivotRes,
      pivotConfig
    );

    const nowTs = Date.now();
    const snapId = `snap_pivot_${nowTs}`;
    setLocalSnapshotsMap((prev) => ({
      ...prev,
      [snapId]: JSON.parse(JSON.stringify(updatedWorkbook)),
    }));
    setSnapshots((prev) =>
      [
        {
          id: snapId,
          roomCode,
          label: `Pivot Tablo Eklendi (${newSheetName})`,
          createdBy: currentUser.name,
          createdByDepartment: currentUser.department,
          createdAt: new Date(nowTs).toISOString(),
          timestamp: nowTs,
          rowCount: pivotRes.rows.length,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setAuditLog((prev) =>
      [
        {
          id: `aud_piv_${nowTs}`,
          roomCode,
          userId: currentUser.id,
          userName: currentUser.name,
          userDepartment: currentUser.department,
          userColor: currentUser.color,
          action: `"${newSheetName}" adıyla yeni Pivot Tablo sayfası oluşturdu`,
          timestamp: nowTs,
          snapshotId: snapId,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setWorkbookData(updatedWorkbook);
    sheetHandleRef.current?.loadWorkbookSnapshot(updatedWorkbook);
    setActiveModal(null);
    showToast(`📊 "${newSheetName}" çalışma kitabınıza yeni sayfa olarak eklendi!`);
  };

  // KOŞULLU BİÇİMLENDİRME KURALI UYGULAMA
  const handleApplyConditionalRule = () => {
    const liveWb = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
    const { updatedWorkbook, matchedCount } = applyConditionalFormattingRule(
      liveWb,
      activeSelection.sheetId,
      cfColIndex,
      cfOperator,
      cfValue1,
      cfValue2,
      cfColorPreset
    );

    const headers = extractSheetHeaders(liveWb, activeSelection.sheetId);
    const colLabel =
      cfColIndex === -1
        ? 'Tüm Tablo'
        : headers.find((h) => h.colIndex === cfColIndex)?.title ||
          `${colIndexToLetter(cfColIndex)} Sütunu`;

    const opLabels: Record<ConditionalRuleOperator, string> = {
      gt: `Büyüktür (> ${cfValue1})`,
      lt: `Küçüktür (< ${cfValue1})`,
      between: `Arasında (${cfValue1} - ${cfValue2})`,
      eq: `Eşittir (= ${cfValue1})`,
      contains: `İçerdiği Metin ("${cfValue1}")`,
      above_avg: 'Ortalamanın Üstünde',
      below_avg: 'Ortalamanın Altında',
      duplicates: 'Yinelenen Değerler',
    };

    const newRule: ConditionalRuleItem = {
      id: `cf_${Date.now()}`,
      sheetId: activeSelection.sheetId,
      colIndex: cfColIndex,
      colLabel,
      operator: cfOperator,
      operatorLabel: opLabels[cfOperator],
      value1: cfValue1,
      value2: cfValue2,
      colorPreset: cfColorPreset,
      matchedCount,
    };

    setActiveCfRules((prev) => [newRule, ...prev]);

    const nowTs = Date.now();
    const snapId = `snap_cf_${nowTs}`;
    setLocalSnapshotsMap((prev) => ({
      ...prev,
      [snapId]: JSON.parse(JSON.stringify(updatedWorkbook)),
    }));
    setSnapshots((prev) =>
      [
        {
          id: snapId,
          roomCode,
          label: `Koşullu Biçimlendirme (${colLabel})`,
          createdBy: currentUser.name,
          createdByDepartment: currentUser.department,
          createdAt: new Date(nowTs).toISOString(),
          timestamp: nowTs,
          rowCount: stats.totalRows,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setAuditLog((prev) =>
      [
        {
          id: `aud_cf_${nowTs}`,
          roomCode,
          userId: currentUser.id,
          userName: currentUser.name,
          userDepartment: currentUser.department,
          userColor: currentUser.color,
          action: `${colLabel} için Koşullu Biçimlendirme uyguladı (${matchedCount} hücre vurgulandı)`,
          timestamp: nowTs,
          snapshotId: snapId,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setWorkbookData(updatedWorkbook);
    sheetHandleRef.current?.loadWorkbookSnapshot(updatedWorkbook);
    showToast(`🎨 Koşullu biçimlendirme uygulandı: ${matchedCount} hücre renklendirildi!`);
  };

  // --- CANLI MASAÜSTÜ EXCEL VLOOKUP / XLOOKUP KÖPRÜSÜ FONKSİYONLARI ---
  const syncLocalDesktopFileIntoBridge = async (
    file: File,
    isLiveWatching: boolean,
    isAutoReload = false
  ) => {
    try {
      const buffer = await file.arrayBuffer();
      const parsedExternal = await parseExcelBufferToUniverData(buffer, file.name);
      const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;

      const { updatedWorkbook, linkedInfo, recalculatedCellCount } =
        mergeExternalExcelIntoWorkbook(
          currentSnap,
          parsedExternal.workbookData,
          file.name,
          isLiveWatching,
          file.lastModified
        );

      lastDesktopFileModifiedRef.current = file.lastModified;
      setLinkedDesktopExcel(linkedInfo);

      if (linkedInfo.headers.length > 1 && !isAutoReload) {
        setBridgeExtKeyCol(0);
        setBridgeExtReturnCol(Math.min(2, linkedInfo.headers.length - 1));
      }

      setWorkbookData(updatedWorkbook);
      sheetHandleRef.current?.loadWorkbookSnapshot(updatedWorkbook);

      if (isAutoReload) {
        showToast(
          `🔄 Masaüstü Excel (${file.name}) kaydedildi! ${
            recalculatedCellCount > 0
              ? `${recalculatedCellCount} VLOOKUP/XLOOKUP hücresi canlı güncellendi.`
              : 'Bağlı tablo verisi yenilendi.'
          }`
        );
      } else {
        showToast(
          `🔗 "${file.name}" bağlandı (${linkedInfo.linkedSheetName} sekmesi eklendi). Şimdi VLOOKUP/XLOOKUP kurabilirsiniz!`
        );
      }
    } catch (err) {
      showToast('⚠️ Masaüstü Excel dosyası okunurken hata oluştu.');
    }
  };

  const handlePickLiveDesktopExcelFile = async () => {
    const win = window as any;
    if (typeof win.showOpenFilePicker === 'function') {
      try {
        const [fileHandle] = await win.showOpenFilePicker({
          multiple: false,
          types: [
            {
              description: 'Microsoft Excel Dosyaları',
              accept: {
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                'application/vnd.ms-excel': ['.xls'],
                'text/csv': ['.csv'],
              },
            },
          ],
        });
        if (fileHandle) {
          desktopFileHandleRef.current = fileHandle;
          const file: File = await fileHandle.getFile();
          await syncLocalDesktopFileIntoBridge(file, true, false);
        }
        return;
      } catch (err: any) {
        if (err?.name === 'AbortError') return;
      }
    }
    // Tarayıcı File System Access API desteklemiyorsa (Safari/Firefox) standart seçiciyi aç
    bridgeFileInputRef.current?.click();
  };

  const handleFallbackBridgeFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    desktopFileHandleRef.current = null;
    await syncLocalDesktopFileIntoBridge(file, false, false);
    e.target.value = '';
  };

  const handleRefreshDesktopFileNow = async () => {
    if (desktopFileHandleRef.current) {
      try {
        const file: File = await desktopFileHandleRef.current.getFile();
        await syncLocalDesktopFileIntoBridge(file, true, true);
        return;
      } catch (e) {
        // ignore
      }
    }
    bridgeFileInputRef.current?.click();
  };

  // Masaüstündeki gerçek Excel dosyasında Ctrl+S yapıldığında otomatik algılayan Canlı İzleyici (2 saniyede bir kontrol)
  useEffect(() => {
    if (!linkedDesktopExcel?.isLiveWatching) return;
    const timer = setInterval(async () => {
      const handle = desktopFileHandleRef.current;
      if (!handle) return;
      try {
        const file: File = await handle.getFile();
        if (file.lastModified > lastDesktopFileModifiedRef.current) {
          await syncLocalDesktopFileIntoBridge(file, true, true);
        }
      } catch (e) {
        // Dosya Excel tarafından o salise yazılıyor olabilir, bir sonraki döngüde okur
      }
    }, 2000);
    return () => clearInterval(timer);
  }, [linkedDesktopExcel?.isLiveWatching, workbookData]);

  const handleApplyBridgeVlookup = () => {
    if (!linkedDesktopExcel) {
      showToast('⚠️ Önce bilgisayarınızdan bağlanacak Excel dosyasını seçin.');
      return;
    }

    const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;

    if (bridgeTargetMode === 'ACTIVE_CELL') {
      // Sadece aktif seçili hücreye canlı formülü yaz
      const srcKeyLetter = colIndexToLetter(bridgeSourceKeyCol);
      const extKeyLetter = colIndexToLetter(bridgeExtKeyCol);
      const extRetLetter = colIndexToLetter(bridgeExtReturnCol);
      const extMaxRow = Math.max(linkedDesktopExcel.rowCount, 2);
      const extSheetName = linkedDesktopExcel.linkedSheetName;
      const targetRowNum = Math.max(activeSelection.row + 1, 2);
      const lookupRef = `${srcKeyLetter}${targetRowNum}`;

      let formulaString = '';
      let computedValue: any = '';

      if (bridgeFormulaType === 'VLOOKUP' && bridgeExtReturnCol >= bridgeExtKeyCol) {
        const startLetter = colIndexToLetter(Math.min(bridgeExtKeyCol, bridgeExtReturnCol));
        const endLetter = colIndexToLetter(Math.max(bridgeExtKeyCol, bridgeExtReturnCol));
        const colIdxNum = bridgeExtReturnCol - bridgeExtKeyCol + 1;
        const res = evaluateFormulaPreview(
          'VLOOKUP',
          {
            lookup_value: lookupRef,
            table_array: `${extSheetName}!${startLetter}2:${endLetter}${extMaxRow}`,
            col_index_num: String(colIdxNum),
            range_lookup: '0',
          },
          currentSnap,
          activeSelection.sheetId
        );
        formulaString = res.formulaString;
        computedValue = res.computedValue;
      } else {
        const res = evaluateFormulaPreview(
          'XLOOKUP',
          {
            lookup_value: lookupRef,
            lookup_array: `${extSheetName}!${extKeyLetter}2:${extKeyLetter}${extMaxRow}`,
            return_array: `${extSheetName}!${extRetLetter}2:${extRetLetter}${extMaxRow}`,
            if_not_found: '"Eşleşmedi"',
          },
          currentSnap,
          activeSelection.sheetId
        );
        formulaString = res.formulaString;
        computedValue = res.computedValue;
      }

      sheetHandleRef.current?.setFormulaOrValueOnCell(
        activeSelection.sheetId,
        activeSelection.row,
        activeSelection.col,
        formulaString,
        computedValue
      );
      setRecentFormulas(recordFormulaUsage(bridgeFormulaType));
      setActiveModal(null);
      showToast(
        `✅ ${activeSelection.cellRef} hücresine masaüstü Excel formülü (${formulaString}) uygulandı!`
      );
      return;
    }

    // TÜM SÜTUNA TOPLU VLOOKUP / XLOOKUP UYGULA
    const {
      updatedWorkbook,
      matchedCount,
      totalRowsProcessed,
      targetColLetter,
      sampleFormula,
    } = applyBulkCrossWorkbookLookup({
      workbookData: currentSnap,
      activeSheetId: activeSelection.sheetId,
      linkedInfo: linkedDesktopExcel,
      formulaType: bridgeFormulaType,
      sourceKeyColIndex: bridgeSourceKeyCol,
      externalKeyColIndex: bridgeExtKeyCol,
      externalReturnColIndex: bridgeExtReturnCol,
      targetColIndex: 'NEW_COL',
      newColumnHeader: bridgeCustomColHeader,
    });

    const nowTs = Date.now();
    const snapId = `snap_bridge_${nowTs}`;
    setLocalSnapshotsMap((prev) => ({
      ...prev,
      [snapId]: JSON.parse(JSON.stringify(updatedWorkbook)),
    }));
    setSnapshots((prev) =>
      [
        {
          id: snapId,
          roomCode,
          label: `Masaüstü Excel ${bridgeFormulaType} (${linkedDesktopExcel.fileName} -> ${targetColLetter} Sütunu)`,
          createdBy: currentUser.name,
          createdByDepartment: currentUser.department,
          createdAt: new Date(nowTs).toISOString(),
          timestamp: nowTs,
          rowCount: stats.totalRows,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setAuditLog((prev) =>
      [
        {
          id: `aud_bridge_${nowTs}`,
          roomCode,
          userId: currentUser.id,
          userName: currentUser.name,
          userDepartment: currentUser.department,
          userColor: currentUser.color,
          action: `Masaüstü Excel (${linkedDesktopExcel.fileName}) ile ${targetColLetter} sütununa canlı ${bridgeFormulaType} kurdu (${matchedCount}/${totalRowsProcessed} satır eşleşti)`,
          cellRef: `${targetColLetter}2`,
          newValue: sampleFormula,
          timestamp: nowTs,
          snapshotId: snapId,
        },
        ...prev,
      ].sort((a, b) => b.timestamp - a.timestamp)
    );

    setRecentFormulas(recordFormulaUsage(bridgeFormulaType));
    setWorkbookData(updatedWorkbook);
    sheetHandleRef.current?.loadWorkbookSnapshot(updatedWorkbook);
    setActiveModal(null);
    showToast(
      `🚀 Masaüstü Excel'den ${matchedCount}/${totalRowsProcessed} satır ${targetColLetter} sütununa ${bridgeFormulaType} ile bağlandı!`
    );
  };

  const loadRoomFromDatabase = async (
    codeToLoad: string,
    pinToUse: string,
    silentIfNotFound = false,
    localSavedAt = 0
  ) => {
    try {
      const clean = codeToLoad.toUpperCase().trim();
      const res = await fetch(
        `/api/rooms?code=${encodeURIComponent(clean)}&pin=${encodeURIComponent(pinToUse)}`
      );
      const data = await res.json();
      if (!res.ok) {
        if (data?.requiresPin) {
          setJoinCodeInput(clean);
          setActiveModal('join');
          showToast('🔒 Bu oda PIN şifresi ile korunuyor. Lütfen PIN şifresini girin.');
          return;
        }
        if (!silentIfNotFound) {
          showToast(`⚠️ ${data.error || 'Oda bulunamadı'}`);
        }
        return;
      }
      if (data.room && data.room.workbookData) {
        const roomUpdatedTs = data.room.updatedAt ? new Date(data.room.updatedAt).getTime() : 0;

        // Eğer sayfa ilk açıldığında (silentIfNotFound=true) tarayıcı hafızasında daha yeni bir kayıt varsa
        // onu ezmemek için sadece oda versiyonunu ve snapshot geçmişini senkronize et
        if (silentIfNotFound && localSavedAt > 0 && localSavedAt > roomUpdatedTs + 2000) {
          setRoomVersion(data.room.version || 1);
          return;
        }

        setRoomCode(data.room.roomCode);
        if (pinToUse !== undefined) {
          setRoomPin(pinToUse);
        }
        const cloned = JSON.parse(JSON.stringify(data.room.workbookData));
        latestRestoredWorkbookRef.current = cloned;
        setWorkbookData(cloned);
        sheetHandleRef.current?.loadWorkbookSnapshot(cloned);

        // Kullanıcının geçmişte girdiği odalar & PIN listesine bu odayı kaydet
        recordRoomInUserHistory(
          data.room.roomCode,
          cloned.name || data.room.title || `Excel_${data.room.roomCode}`,
          pinToUse || undefined,
          data.room.ownerId === currentUser.id ? 'owner' : 'editor'
        );

        // Açık çalışma kitapları sekmesindeki aktif kitabı da sunucudan gelen son kayıtla eşitle
        setOpenWorkbooks((prev) => {
          const exists = prev.some(
            (t) => t.id === cloned.id || t.roomCode === data.room.roomCode
          );
          if (exists) {
            return prev.map((t) =>
              t.id === cloned.id || t.roomCode === data.room.roomCode
                ? {
                    ...t,
                    id: cloned.id || `wb_${data.room.roomCode}`,
                    roomCode: data.room.roomCode,
                    name: cloned.name,
                    workbookData: cloned,
                  }
                : t
            );
          }
          return [
            {
              id: cloned.id || `wb_${data.room.roomCode}`,
              roomCode: data.room.roomCode,
              name: cloned.name || data.room.title || 'Kayitli_Tablo',
              workbookData: cloned,
              updatedAt: Date.now(),
            },
            ...prev.filter((t) => t.id !== 'wb_default_enterprise'),
          ];
        });
        setActiveWorkbookTabId(cloned.id || `wb_${data.room.roomCode}`);

        setRoomVersion(data.room.version || 1);
        setProtectedColumns(data.room.protectedColumns || []);

        const loadedRoomCode = data.room.roomCode;
        const roomAuditList: AuditEvent[] = Array.isArray(data.room.auditLog)
          ? data.room.auditLog.map((ev: any) => ({
              ...ev,
              roomCode: loadedRoomCode,
            }))
          : [];
        setAuditLog((prev) =>
          [
            ...roomAuditList,
            ...prev.filter((a) => (a.roomCode || 'XLS-849-201') !== loadedRoomCode),
          ].sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0))
        );

        const snapMap: Record<string, UniverWorkbookData> = {};
        const roomSnapshots: SnapshotMeta[] = Array.isArray(data.room.snapshots)
          ? data.room.snapshots.map((s: any) => {
              const ts = s.createdAt ? new Date(s.createdAt).getTime() : Date.now();
              if (s.workbookData) {
                snapMap[s.id] = s.workbookData;
              }
              return {
                id: s.id,
                roomCode: loadedRoomCode,
                label: s.label,
                createdBy: s.createdBy,
                createdAt: s.createdAt,
                timestamp: ts,
                rowCount: s.rowCount || 10,
                isAutoChange: s.isAutoChange,
              };
            })
          : [];
        if (Object.keys(snapMap).length > 0) {
          setLocalSnapshotsMap((prev) => ({ ...prev, ...snapMap }));
        }
        setSnapshots((prev) =>
          [
            ...roomSnapshots,
            ...prev.filter((s) => (s.roomCode || 'XLS-849-201') !== loadedRoomCode),
          ].sort((a, b) => b.timestamp - a.timestamp)
        );

        setActiveModal(null);
        if (!silentIfNotFound) {
          showToast(
            `📂 "${cloned.name || data.room.title}" (${data.room.roomCode}) odasına geçiş yapıldı!`
          );
        }
      }
    } catch (e) {
      if (!silentIfNotFound) {
        showToast('Oda yüklenirken hata oluştu.');
      }
    }
  };

  // Kullanıcının geçmişte girdiği odalar/Exceller arasında tek tıkla geçiş yapması (Önce mevcut odayı kaydeder!)
  const handleSwitchToUserRoom = async (targetCode: string, savedPinForRoom = '') => {
    const cleanTarget = targetCode.trim().toUpperCase();
    if (cleanTarget === roomCode) {
      setActiveModal(null);
      return;
    }

    // Geçiş yapmadan önce üzerinde çalışılan mevcut odayı sessizce kaydet
    try {
      const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
      await fetch('/api/rooms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          actionType: 'SAVE_WORKBOOK',
          roomCode,
          pin: roomPin || undefined,
          defaultRole: defaultRoomRole,
          protectedColumns,
          title: currentSnap.name,
          ownerId: currentUser.id,
          ownerName: currentUser.name,
          ownerEmail: currentUser.email,
          workbookData: currentSnap,
          silentSwitchSave: true,
          stats: {
            totalRows: stats.totalRows,
            totalCells: stats.totalCells,
            compressedKB: stats.compressedKB,
          },
        }),
      });
    } catch (e) {
      // ignore
    }

    await loadRoomFromDatabase(cleanTarget, savedPinForRoom, false, 0);
  };

  // Kullanıcının yeni bir Excel odası (PIN'li veya PIN'siz) oluşturup geçmişine eklemesi
  const handleCreateNewUserExcelRoom = async () => {
    const randomPart1 = Math.floor(100 + Math.random() * 900);
    const randomPart2 = Math.floor(100 + Math.random() * 900);
    const generatedCode = `XLS-${randomPart1}-${randomPart2}`;

    let baseWb: UniverWorkbookData;
    if (newRoomForm.template === 'budget') {
      baseWb = getDefaultEnterpriseWorkbook();
    } else if (newRoomForm.template === 'crm') {
      baseWb = getSalesCrmWorkbook();
    } else {
      baseWb = getBlankWorkbook(userSavedRooms.length + 1);
    }

    const finalTitle =
      newRoomForm.title.trim().replace(/\s+/g, '_') ||
      `Yeni_Excel_${userSavedRooms.length + 1}`;
    const finalPin = newRoomForm.pin.trim();

    baseWb.id = `wb_${generatedCode}`;
    baseWb.name = finalTitle;

    try {
      // Önce mevcut odayı sessizce kaydet
      const currentSnap = sheetHandleRef.current?.getCurrentSnapshot() || workbookData;
      await fetch('/api/rooms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          actionType: 'SAVE_WORKBOOK',
          roomCode,
          pin: roomPin || undefined,
          title: currentSnap.name,
          ownerId: currentUser.id,
          ownerName: currentUser.name,
          workbookData: currentSnap,
          silentSwitchSave: true,
        }),
      });

      // Yeni odayı veritabanında oluştur
      const createRes = await fetch('/api/rooms', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          actionType: 'SAVE_WORKBOOK',
          roomCode: generatedCode,
          pin: finalPin || undefined,
          defaultRole: 'editor',
          protectedColumns: [],
          title: finalTitle,
          ownerId: currentUser.id,
          ownerName: currentUser.name,
          ownerEmail: currentUser.email,
          workbookData: baseWb,
          stats: {
            totalRows: 20,
            totalCells: 100,
            compressedKB: 2,
          },
        }),
      });
      const createData = await createRes.json().catch(() => ({}));

      const nowTs = Date.now();
      const initSnapId = createData?.room?.snapshots?.[0]?.id || `snap_init_${nowTs}`;
      setLocalSnapshotsMap((prev) => ({
        ...prev,
        [initSnapId]: JSON.parse(JSON.stringify(baseWb)),
      }));
      setSnapshots((prev) =>
        [
          {
            id: initSnapId,
            roomCode: generatedCode,
            label: `Oda Oluşturuldu (${finalTitle})`,
            createdBy: currentUser.name,
            createdByDepartment: currentUser.department,
            createdAt: new Date(nowTs).toISOString(),
            timestamp: nowTs,
            rowCount: 20,
          },
          ...prev.filter((s) => (s.roomCode || 'XLS-849-201') !== generatedCode),
        ].sort((a, b) => b.timestamp - a.timestamp)
      );
      setAuditLog((prev) =>
        [
          {
            id: `aud_init_${nowTs}`,
            roomCode: generatedCode,
            userId: currentUser.id,
            userName: currentUser.name,
            userDepartment: currentUser.department,
            userColor: currentUser.color,
            action: `"${finalTitle}" (${generatedCode}) Excel odasını oluşturdu`,
            timestamp: nowTs,
            snapshotId: initSnapId,
          },
          ...prev.filter((a) => (a.roomCode || 'XLS-849-201') !== generatedCode),
        ].sort((a, b) => b.timestamp - a.timestamp)
      );

      setRoomCode(generatedCode);
      setRoomPin(finalPin);
      setIsRoomOwner(true);
      latestRestoredWorkbookRef.current = baseWb;
      setWorkbookData(baseWb);
      sheetHandleRef.current?.loadWorkbookSnapshot(baseWb);

      setOpenWorkbooks((prev) => [
        {
          id: baseWb.id,
          roomCode: generatedCode,
          name: finalTitle,
          workbookData: baseWb,
          updatedAt: Date.now(),
        },
        ...prev,
      ]);
      setActiveWorkbookTabId(baseWb.id);

      await recordRoomInUserHistory(generatedCode, finalTitle, finalPin, 'owner');
      setNewRoomForm({ title: '', pin: '', template: 'blank' });
      setActiveModal(null);
      showToast(
        `✨ Yeni Excel odanız "${finalTitle}" (${generatedCode}${finalPin ? ` • PIN: ${finalPin}` : ''}) oluşturuldu!`
      );
    } catch (e) {
      showToast('Yeni oda oluşturulurken hata oluştu.');
    }
  };

  const fetchSavedRooms = async () => {
    try {
      const res = await fetch('/api/rooms');
      const data = await res.json();
      if (data.rooms) setSavedRoomsList(data.rooms);
    } catch (e) {
      // ignore
    }
  };

  // Canlı Pivot ve Sütun Hesaplamalarını Memoize Et (Büyük Excel dosyalarında her saniye yeniden hesaplanıp ekranı dondurmasını önler)
  const livePivotResult = useMemo(
    () => computePivotTable(workbookData, pivotConfig),
    [workbookData, pivotConfig]
  );

  const currentSheetColumns = useMemo(
    () => getSheetColumnRanges(workbookData, activeSelection.sheetId),
    [workbookData, activeSelection.sheetId]
  );

  const selectedFormulaDef =
    EXCEL_FORMULA_CATALOG.find((f) => f.name === selectedFormulaName) || EXCEL_FORMULA_CATALOG[0];

  const liveFormulaPreview = useMemo(() => {
    if (activeModal !== 'fx') {
      return { formulaString: `=${selectedFormulaName}()`, computedValue: '', explanation: '' };
    }
    return evaluateFormulaPreview(
      selectedFormulaName,
      fxArgValues,
      workbookData,
      activeSelection.sheetId
    );
  }, [selectedFormulaName, fxArgValues, workbookData, activeSelection.sheetId, activeModal]);

  // Versiyon ve İşlem listesini SADECE AKTİF BULUNULAN ODADAKİ EXCEL İÇİN filtrele ve en yeni en üstte sırala
  const sortedSnapshots = snapshots
    .filter((s) => (s.roomCode || 'XLS-849-201') === roomCode)
    .sort((a, b) => b.timestamp - a.timestamp);
  const sortedAuditLog = auditLog
    .filter((a) => (a.roomCode || 'XLS-849-201') === roomCode)
    .sort((a, b) => b.timestamp - a.timestamp);

  const collaboratorList = Object.values(collaborators);

  const activeCellLockInfo = getActiveCellLockOwner(
    collaborators,
    currentUser.id,
    activeSelection.sheetId,
    activeSelection.row,
    activeSelection.col,
    nowTick,
    lockTimeoutSec * 1000
  );

  // EĞER KULLANICI GİRİŞ YAPMADIYSA GERÇEK KULLANICI ADI & ŞİFRE GİRİŞ / KAYIT EKRANINI GÖSTER
  if (!isAuthenticated) {
    return (
      <div className="min-h-screen w-screen bg-slate-900 flex items-center justify-center p-4">
        <div className="bg-white border border-slate-200 rounded-2xl max-w-md w-full p-7 shadow-2xl">
          <div className="flex items-center gap-3 mb-5">
            <div className="w-11 h-11 rounded-xl bg-[#107c41] flex items-center justify-center shadow-sm">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-900">SyncSheet Kurumsal Excel</h1>
              <p className="text-xs text-slate-500">
                Kişisel Oda, PIN ve Çalışma Kitabı Yönetimli Kullanıcı Girişi
              </p>
            </div>
          </div>

          {/* Giriş Yap / Hesap Oluştur Sekmeleri */}
          <div className="grid grid-cols-2 gap-1 p-1 bg-slate-100 rounded-xl mb-5 border border-slate-200">
            <button
              type="button"
              onClick={() => {
                setAuthMode('login');
                setAuthError(null);
              }}
              className={`py-2 rounded-lg text-xs font-bold transition cursor-pointer ${
                authMode === 'login'
                  ? 'bg-white text-slate-900 shadow-2xs'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              Giriş Yap
            </button>
            <button
              type="button"
              onClick={() => {
                setAuthMode('register');
                setAuthError(null);
              }}
              className={`py-2 rounded-lg text-xs font-bold transition cursor-pointer ${
                authMode === 'register'
                  ? 'bg-white text-emerald-700 shadow-2xs'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              Hesap Oluştur (Kayıt Ol)
            </button>
          </div>

          {authError && (
            <div className="mb-4 p-3 rounded-xl bg-rose-50 border border-rose-200 text-xs text-rose-700 font-medium">
              ⚠️ {authError}
            </div>
          )}

          <form onSubmit={handleAuthSubmit} className="space-y-3.5">
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1">
                Kullanıcı Adı <span className="text-rose-500">*</span>
              </label>
              <input
                type="text"
                required
                value={authForm.username}
                onChange={(e) => setAuthForm({ ...authForm, username: e.target.value })}
                placeholder="Örn: denizsara"
                className="w-full bg-white border border-slate-300 focus:border-emerald-600 focus:outline-none rounded-xl px-3.5 py-2.5 text-xs text-slate-900"
              />
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1">
                Şifre <span className="text-rose-500">*</span>
              </label>
              <input
                type="password"
                required
                value={authForm.password}
                onChange={(e) => setAuthForm({ ...authForm, password: e.target.value })}
                placeholder="Şifrenizi girin..."
                className="w-full bg-white border border-slate-300 focus:border-emerald-600 focus:outline-none rounded-xl px-3.5 py-2.5 text-xs text-slate-900"
              />
            </div>

            {authMode === 'register' && (
              <div className="grid grid-cols-2 gap-2.5 pt-1">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Ad Soyad
                  </label>
                  <input
                    type="text"
                    value={authForm.name}
                    onChange={(e) => setAuthForm({ ...authForm, name: e.target.value })}
                    placeholder="Örn: Deniz Sara"
                    className="w-full bg-white border border-slate-300 focus:border-emerald-600 focus:outline-none rounded-xl px-3 py-2 text-xs text-slate-900"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Departman / Birim
                  </label>
                  <input
                    type="text"
                    value={authForm.department}
                    onChange={(e) => setAuthForm({ ...authForm, department: e.target.value })}
                    placeholder="Örn: Finans / Yönetim"
                    className="w-full bg-white border border-slate-300 focus:border-emerald-600 focus:outline-none rounded-xl px-3 py-2 text-xs text-slate-900"
                  />
                </div>
              </div>
            )}

            <button
              type="submit"
              disabled={authLoading}
              className="w-full mt-2 py-3 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] disabled:opacity-60 text-white text-xs font-bold shadow-sm transition cursor-pointer"
            >
              {authLoading
                ? 'İşlem Yapılıyor...'
                : authMode === 'register'
                ? 'Hesabımı Oluştur ve Excellerimi Aç'
                : 'Giriş Yap ve Kayıtlı Excellerimi Getir'}
            </button>
          </form>

          <div className="mt-5 pt-4 border-t border-slate-100 text-[11px] text-slate-500 leading-relaxed">
            🔒 Hesabınızla giriş yaptığınızda <strong>geçmişte girdiğiniz tüm Excel odaları, oda kodları ve PIN şifreleri</strong> hesabınızda otomatik saklanır; istediğiniz zaman tek tıkla odalarınız arasında geçiş yapabilirsiniz.
          </div>
        </div>
      </div>
    );
  }

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setIsDraggingOver(true);
      }}
      onDragLeave={() => setIsDraggingOver(false)}
      onDrop={handleDropExcelFile}
      onClick={() => {
        if (openMenu !== 'none') setOpenMenu('none');
      }}
      className="flex flex-col w-screen h-screen bg-white text-slate-800 select-none overflow-hidden"
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleLocalExcelFileSelect}
        className="hidden"
      />
      <input
        ref={bridgeFileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleFallbackBridgeFileSelect}
        className="hidden"
      />

      {isDraggingOver && (
        <div className="fixed inset-0 z-50 bg-emerald-600/15 backdrop-blur-xs border-4 border-dashed border-emerald-600 flex flex-col items-center justify-center pointer-events-none">
          <div className="bg-white px-8 py-6 rounded-2xl shadow-xl text-center">
            <Upload className="w-12 h-12 text-emerald-600 mx-auto mb-2 animate-bounce" />
            <p className="text-lg font-bold text-slate-800">
              Excel Dosyasını Buraya Bırakın (Excellerim Listenize Eklenir)
            </p>
            <p className="text-xs text-slate-500 mt-1">
              Sunucuya dosya yüklenmeden (Zero-Upload) doğrudan hesabınızdaki odalara kaydedilir
            </p>
          </div>
        </div>
      )}

      {/* 1. KATMAN: KULLANICININ GEÇMİŞTE GİRDİĞİ ODALAR / EXCELLER VE PIN YÖNETİM ÇUBUĞU */}
      <div className="h-10 bg-slate-900 text-slate-200 px-3 flex items-center justify-between gap-2 text-xs shrink-0 border-b border-slate-800">
        <div className="flex items-center gap-1.5 overflow-x-auto no-scrollbar py-1">
          <button
            onClick={() => {
              fetchSavedRooms();
              setActiveModal('my_rooms');
            }}
            className="text-[11px] font-bold bg-emerald-600 hover:bg-emerald-500 text-white px-2.5 py-1 rounded-lg flex items-center gap-1.5 shrink-0 cursor-pointer transition shadow-2xs"
            title="Geçmişte girdiğiniz tüm Excel odalarını, oda kodlarını ve PIN şifrelerini görüntüleyin"
          >
            <FolderOpen className="w-3.5 h-3.5" />
            <span>Excellerim & PIN&apos;lerim ({userSavedRooms.length})</span>
          </button>

          <span className="text-slate-700 mx-0.5">|</span>

          {/* Kullanıcının Kayıtlı / Geçmişte Girdiği Excel Odaları Sekmeleri */}
          {userSavedRooms.map((rm) => {
            const isActiveRoom = rm.roomCode === roomCode;
            return (
              <div
                key={rm.roomCode}
                onClick={() => handleSwitchToUserRoom(rm.roomCode, rm.pin)}
                className={`group flex items-center gap-1.5 px-3 py-1 rounded-t-lg text-xs font-medium cursor-pointer transition shrink-0 border-b-2 ${
                  isActiveRoom
                    ? 'bg-white text-slate-900 font-bold border-emerald-600 shadow-xs'
                    : 'bg-slate-800/90 text-slate-300 hover:bg-slate-700 border-transparent'
                }`}
                title={`Oda Kodu: ${rm.roomCode}${rm.pin ? ` • Kayıtlı PIN: ${rm.pin}` : ' • Şifresiz Oda'} (Geçiş yapmak için tıklayın)`}
              >
                <FileSpreadsheet
                  className={`w-3.5 h-3.5 shrink-0 ${
                    isActiveRoom ? 'text-emerald-600' : 'text-emerald-400'
                  }`}
                />
                <span className="truncate max-w-[165px]">{rm.title}</span>
                <span
                  className={`text-[10px] font-mono px-1.5 py-0.2 rounded ${
                    isActiveRoom
                      ? 'bg-emerald-100 text-emerald-800 font-bold'
                      : 'bg-slate-700 text-slate-300'
                  }`}
                >
                  {rm.roomCode}
                </span>
                {rm.pin && (
                  <span
                    className={`text-[10px] font-mono px-1.5 py-0.2 rounded flex items-center gap-0.5 ${
                      isActiveRoom
                        ? 'bg-amber-100 text-amber-800 font-bold'
                        : 'bg-amber-500/20 text-amber-300'
                    }`}
                    title={`Bu odanın kayıtlı PIN şifresi: ${rm.pin}`}
                  >
                    <Lock className="w-2.5 h-2.5" />
                    PIN:{rm.pin}
                  </span>
                )}
              </div>
            );
          })}

          <button
            onClick={() => {
              fetchSavedRooms();
              setActiveModal('my_rooms');
            }}
            className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-emerald-400 text-[11px] font-semibold transition shrink-0 cursor-pointer"
            title="Yeni Excel Odası Oluştur veya Oda Kodu + PIN ile Başka Bir Excel'e Katıl"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Yeni Oda / Odaya Katıl</span>
          </button>

          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/30 text-[11px] font-semibold transition shrink-0 cursor-pointer"
            title="Bilgisayardan bir Excel (.xlsx) dosyası yükleyip kendi odalarınıza ekleyin"
          >
            <Upload className="w-3 h-3" />
            <span>Bilgisayardan Excel Ekle (.xlsx)</span>
          </button>
        </div>

        {/* Sağ Üst: Giriş Yapan Kullanıcı Hesabı & Oturumu Kapat */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="hidden md:flex items-center gap-1.5 bg-slate-800 px-2.5 py-0.5 rounded-full border border-slate-700">
            <UserCheck className="w-3.5 h-3.5 text-emerald-400" />
            <span className="text-[11px] font-bold text-white">{currentUser.name}</span>
            <span className="text-[10px] font-mono text-emerald-400">
              (@{currentUser.username})
            </span>
          </div>
          <button
            onClick={handleLogout}
            className="flex items-center gap-1 px-2.5 py-0.5 rounded bg-slate-800 hover:bg-rose-950 hover:text-rose-300 text-slate-300 border border-slate-700 text-[11px] transition cursor-pointer"
            title="Hesabınızdan çıkış yapın"
          >
            <LogOut className="w-3 h-3" />
            <span>Çıkış Yap</span>
          </button>
        </div>
      </div>

      {/* 2. KATMAN: ANA ÜST BAR (AKTİF EXCEL ADI, ODA KODU, KAYIT DURUMU VE ORTAK ÇALIŞANLAR) */}
      <header className="h-13 bg-white border-b border-slate-200 px-4 flex items-center justify-between gap-3 z-30 shrink-0">
        <div className="flex items-center gap-3 min-w-0">
          <div className="flex items-center gap-2">
            <input
              type="text"
              value={workbookData.name}
              onChange={(e) => {
                const newName = e.target.value;
                setWorkbookData((prev) => ({ ...prev, name: newName }));
                setUserSavedRooms((prev) =>
                  prev.map((r) => (r.roomCode === roomCode ? { ...r, title: newName } : r))
                );
                setOpenWorkbooks((prev) =>
                  prev.map((t) => (t.id === activeWorkbookTabId ? { ...t, name: newName } : t))
                );
              }}
              className="font-bold text-sm text-slate-900 hover:bg-slate-100 focus:bg-white focus:ring-2 focus:ring-emerald-600 focus:outline-none px-2 py-1 rounded transition truncate max-w-[240px]"
              title="Excel çalışma kitabı adını değiştirmek için tıklayın"
            />
            <span className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200 font-mono text-[11px] font-bold text-slate-700">
              Oda: {roomCode}
            </span>
            {roomPin && (
              <span className="px-2 py-0.5 rounded bg-amber-50 border border-amber-200 font-mono text-[11px] font-bold text-amber-800 flex items-center gap-1">
                <Lock className="w-3 h-3" />
                PIN: {roomPin}
              </span>
            )}
          </div>

          {/* KAYIT VE İNDİRME BUTONLARI (NET VE ANLAŞILIR) */}
          <div className="flex items-center gap-1.5">
            <button
              onClick={saveRoomToDatabase}
              className="inline-flex items-center gap-1.5 text-xs font-semibold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-3 py-1.5 rounded-lg cursor-pointer transition"
              title="Değişiklikleri aktif Excel odasına ve hesabınızın oda geçmişine kaydeder"
            >
              <Save className="w-3.5 h-3.5 text-emerald-600" />
              <span>Kaydet</span>
              <span className="text-[10px] font-normal text-emerald-600 hidden lg:inline">
                ({stats.lastSyncText})
              </span>
            </button>

            <button
              onClick={handleExportToExcel}
              className="inline-flex items-center gap-1.5 text-xs font-semibold text-slate-700 bg-slate-100 hover:bg-slate-200 border border-slate-200 px-3 py-1.5 rounded-lg cursor-pointer transition"
              title="Tabloyu bilgisayarınıza normal .xlsx Excel dosyası olarak indirir"
            >
              <Download className="w-3.5 h-3.5 text-slate-700" />
              <span>Bilgisayara İndir (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* SAĞ: Aktif Kullanıcılar, Odalarım & PIN'lerim ve Paylaşım */}
        <div className="flex items-center gap-2 shrink-0">
          <div className="flex items-center gap-1 bg-slate-100 px-2 py-1 rounded-full border border-slate-200">
            {collaboratorList.map((collab) => {
              const isFollowing = followingUserId === collab.userId;
              const activityTime = collab.lastActivityAt || collab.lastUpdated || 0;
              const elapsedSec = Math.floor((nowTick - activityTime) / 1000);
              const isCellStillLocked =
                collab.role === 'editor' && elapsedSec >= 0 && elapsedSec < lockTimeoutSec;
              const remSec = Math.max(0, lockTimeoutSec - elapsedSec);

              return (
                <button
                  key={collab.userId}
                  onClick={() => {
                    if (isFollowing) {
                      setFollowingUserId(null);
                      showToast(`${collab.userName} takibi durduruldu.`);
                    } else {
                      setFollowingUserId(collab.userId);
                      sheetHandleRef.current?.navigateToCell(collab.row, collab.col);
                      showToast(`👁️ ${collab.userName} (${collab.activeCellRef}) takip ediliyor`);
                    }
                  }}
                  className={`flex items-center gap-1 pl-1 pr-2 py-0.5 rounded-full text-xs transition cursor-pointer ${
                    isFollowing
                      ? 'bg-blue-600 text-white shadow-xs'
                      : 'hover:bg-white text-slate-700'
                  }`}
                  title={
                    isCellStillLocked
                      ? `🔒 ${collab.userName} şu an ${collab.activeCellRef} hücresinde işlem yapıyor (${remSec} sn hareketsiz kalırsa kilit kalkar).`
                      : `🔓 ${collab.userName} ${collab.activeCellRef} hücresinde (Hareketsizlik süresi dolduğu için kilit kalktı).`
                  }
                >
                  <span
                    className="w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold text-white"
                    style={{ backgroundColor: collab.userColor }}
                  >
                    {collab.userName.slice(0, 2).toUpperCase()}
                  </span>
                  <span className="font-mono text-[11px] font-semibold flex items-center gap-0.5">
                    {collab.activeCellRef}
                    {isCellStillLocked ? (
                      <Lock className="w-2.5 h-2.5 text-amber-600" />
                    ) : (
                      <Unlock className="w-2.5 h-2.5 text-slate-400" />
                    )}
                  </span>
                </button>
              );
            })}

            <div
              className="flex items-center gap-1 pl-1 pr-2 py-0.5 rounded-full bg-white border border-slate-200 text-xs font-semibold text-slate-800"
              title={`Aktif Kullanıcı: ${currentUser.name} (@${currentUser.username})`}
            >
              <span
                className="w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold text-white"
                style={{ backgroundColor: currentUser.color }}
              >
                {currentUser.name.slice(0, 2).toUpperCase()}
              </span>
              <span className="text-[11px] hidden sm:inline">Siz ({activeSelection.cellRef})</span>
            </div>
          </div>

          <button
            onClick={() => {
              fetchSavedRooms();
              setActiveModal('my_rooms');
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-300 transition cursor-pointer"
          >
            <FolderOpen className="w-3.5 h-3.5 text-emerald-700" />
            <span>Excellerim & Geçmiş Odalarım</span>
          </button>

          <button
            onClick={() => setActiveModal('share')}
            className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold bg-[#107c41] hover:bg-[#0c5e31] text-white shadow-xs transition cursor-pointer"
          >
            <Share2 className="w-3.5 h-3.5" />
            <span>Paylaş & PIN Ayarla ({roomCode})</span>
          </button>
        </div>
      </header>

      {/* 3. KATMAN: TÜRKÇE EXCEL ÖZELLİK ŞERİDİ (RIBBON BAR - FX, MASAÜSTÜ EXCEL KÖPRÜSÜ, PİVOT, KOŞULLU BİÇİMLENDİRME) */}
      <div className="bg-slate-50 border-b border-slate-200 px-4 py-1.5 flex items-center justify-between gap-2 flex-wrap z-20 shrink-0">
        <div className="flex items-center gap-1.5 flex-wrap">
          {/* FX FONKSİYON SİHİRBAZI BUTONU */}
          <button
            onClick={() => {
              handleSelectFormulaInWizard(recentFormulas[0] || 'XLOOKUP');
              setActiveModal('fx');
            }}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white shadow-2xs transition cursor-pointer"
            title="Kolay XLOOKUP, VLOOKUP, IF, SUM ve Sık Kullanılan Fonksiyon Sihirbazı"
          >
            <FunctionSquare className="w-4 h-4" />
            <span>fx Fonksiyon Ekle (XLOOKUP)</span>
          </button>

          {/* MASAÜSTÜ GERÇEK EXCEL <-> SYNCSHEET CANLI VLOOKUP KÖPRÜSÜ BUTONU */}
          <button
            onClick={() => setActiveModal('excel_bridge')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold border shadow-2xs transition cursor-pointer ${
              linkedDesktopExcel
                ? 'bg-emerald-600 hover:bg-emerald-700 text-white border-emerald-700'
                : 'bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border-emerald-300'
            }`}
            title="Bilgisayarınızda açık olan gerçek Microsoft Excel ile bu sistem arasında canlı VLOOKUP / XLOOKUP bağlantısı kurun"
          >
            <Link2 className="w-4 h-4" />
            <span>Masaüstü Excel VLOOKUP Köprüsü</span>
            {linkedDesktopExcel && (
              <span className="px-1.5 py-0.2 rounded-full bg-white text-emerald-800 text-[10px] font-extrabold flex items-center gap-1">
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                Canlı Bağlı
              </span>
            )}
          </button>

          {/* PİVOT TABLO OLUŞTUR BUTONU */}
          <button
            onClick={() => setActiveModal('pivot_modal')}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white hover:bg-blue-50 text-blue-700 border border-blue-200 transition cursor-pointer"
            title="Excel tarzı satır, sütun ve değer seçerek Pivot Tablo oluşturun"
          >
            <Table2 className="w-4 h-4 text-blue-600" />
            <span>Pivot Tablo Oluştur</span>
          </button>

          {/* KOŞULLU BİÇİMLENDİRME BUTONU */}
          <button
            onClick={() =>
              setActiveDrawer(activeDrawer === 'conditional' ? 'none' : 'conditional')
            }
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
              activeDrawer === 'conditional'
                ? 'bg-rose-50 border-rose-300 text-rose-700'
                : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-100'
            }`}
            title="Büyüktür, Küçüktür, İçerdiği Metin ve Renk Ölçekleri ile hücreleri vurgulayın"
          >
            <PaintBucket className="w-4 h-4 text-rose-600" />
            <span>Koşullu Biçimlendirme</span>
            {activeCfRules.length > 0 && (
              <span className="px-1.5 py-0.2 rounded-full bg-rose-600 text-white text-[10px]">
                {activeCfRules.length}
              </span>
            )}
          </button>

          <span className="text-slate-300 mx-0.5">|</span>

          {/* HIZLI FORMÜL KISAYOLLARI (EN SON KULLANILANLAR) */}
          <div className="hidden lg:flex items-center gap-1">
            <span className="text-[11px] font-semibold text-slate-400 mr-1">
              Son Kullanılanlar:
            </span>
            {recentFormulas.slice(0, 4).map((fnName) => (
              <button
                key={fnName}
                onClick={() => {
                  handleSelectFormulaInWizard(fnName);
                  setActiveModal('fx');
                }}
                className="px-2 py-1 rounded bg-white hover:bg-indigo-50 text-indigo-700 border border-slate-200 text-[11px] font-mono font-semibold transition cursor-pointer"
                title={`${fnName} fonksiyon sihirbazını aç`}
              >
                ={fnName}()
              </button>
            ))}
          </div>

          <span className="text-slate-300 mx-0.5 hidden lg:inline">|</span>

          {/* KİŞİSEL FİLTRE VE VERİ ARAÇLARI */}
          <button
            onClick={() => {
              setPersonalFilterMode(!personalFilterMode);
              showToast(
                !personalFilterMode
                  ? 'Kişisel Filtre Açık: Filtreleriniz sadece sizin ekranınızda görünür.'
                  : 'Genel Filtre Açık: Filtreler herkesin ekranında uygulanır.'
              );
            }}
            className={`px-2.5 py-1 rounded-lg text-xs flex items-center gap-1 border transition cursor-pointer ${
              personalFilterMode
                ? 'text-emerald-700 bg-emerald-50 border-emerald-200 font-semibold'
                : 'text-slate-600 bg-white border-slate-200 hover:bg-slate-100'
            }`}
          >
            <Filter className="w-3.5 h-3.5" />
            <span>{personalFilterMode ? 'Kişisel Filtre' : 'Genel Filtre'}</span>
          </button>

          <button
            onClick={() => setActiveModal('clipboard')}
            className="px-2.5 py-1 rounded-lg text-xs font-medium bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 flex items-center gap-1 cursor-pointer"
          >
            <ClipboardPaste className="w-3.5 h-3.5 text-blue-600" />
            <span className="hidden sm:inline">Panodan Yapıştır</span>
          </button>

          {/* Büyük Veri Test Menüsü */}
          <div className="relative" onClick={(e) => e.stopPropagation()}>
            <button
              onClick={() => setOpenMenu(openMenu === 'tools' ? 'none' : 'tools')}
              className="px-2.5 py-1 rounded-lg text-xs font-medium bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 flex items-center gap-1 cursor-pointer"
            >
              <Zap className="w-3.5 h-3.5 text-amber-600" />
              <span>Veri & Makro</span>
              <ChevronDown className="w-3 h-3 text-slate-400" />
            </button>
            {openMenu === 'tools' && (
              <div className="absolute left-0 top-8 w-60 bg-white rounded-xl shadow-xl border border-slate-200 py-1.5 z-50 text-xs">
                <button
                  onClick={() => {
                    setOpenMenu('none');
                    setActiveModal('macro');
                  }}
                  className="w-full px-3.5 py-2 text-left hover:bg-slate-50 flex items-center gap-2 text-slate-700"
                >
                  <Terminal className="w-4 h-4 text-emerald-600" />
                  <span>Makro / Otomasyon Çalıştır</span>
                </button>
                <div className="my-1 border-t border-slate-100" />
                <button
                  onClick={() => handleLoadLargeDataset(50000)}
                  className="w-full px-3.5 py-1.5 text-left hover:bg-slate-50 flex items-center gap-2 text-slate-700"
                >
                  <Zap className="w-3.5 h-3.5 text-purple-600" />
                  <span>50.000 Satır Örnek Veri Aç</span>
                </button>
                <button
                  onClick={() => handleLoadLargeDataset(100000)}
                  className="w-full px-3.5 py-1.5 text-left hover:bg-slate-50 flex items-center gap-2 text-slate-700"
                >
                  <Layers className="w-3.5 h-3.5 text-pink-600" />
                  <span>100.000 Satır Örnek Veri Aç</span>
                </button>
              </div>
            )}
          </div>
        </div>

        {/* SAĞ TARAF: VERSİYON GEÇMİŞİ VE HÜCRE KİLİT AYARLARI */}
        <div className="flex items-center gap-1.5">
          <button
            onClick={() => setActiveDrawer(activeDrawer === 'history' ? 'none' : 'history')}
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
              activeDrawer === 'history'
                ? 'bg-amber-50 border-amber-300 text-amber-800'
                : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-100'
            }`}
          >
            <History className="w-4 h-4 text-amber-600" />
            <span>Versiyon Geçmişi ({sortedSnapshots.length})</span>
          </button>

          <button
            onClick={() =>
              setActiveDrawer(activeDrawer === 'protection' ? 'none' : 'protection')
            }
            className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
              activeDrawer === 'protection'
                ? 'bg-rose-50 border-rose-300 text-rose-800'
                : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-100'
            }`}
          >
            <Lock className="w-3.5 h-3.5 text-rose-600" />
            <span>Hücre Kilidi ({lockTimeoutSec}sn)</span>
          </button>
        </div>
      </div>

      {/* AKILLI HÜCRE KİLİDİ BANNERI */}
      {activeCellLockInfo && (
        <div className="bg-amber-50 border-b border-amber-200 px-4 py-1.5 flex items-center justify-between text-xs text-amber-900 z-20">
          <div className="flex items-center gap-2">
            <Lock className="w-4 h-4 text-amber-600 shrink-0" />
            <span>
              <strong>{activeSelection.cellRef}</strong> hücresinde şu anda{' '}
              <strong>{activeCellLockInfo.lockedBy.userName}</strong> işlem yapıyor. Çakışmayı önlemek için bu hücre geçici olarak kilitlidir.
            </span>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-amber-100 border border-amber-300 font-mono font-semibold text-amber-800">
              <Clock className="w-3 h-3" />
              Hareketsiz kalırsa {activeCellLockInfo.remainingSeconds} sn sonra açılır
            </span>
          </div>
        </div>
      )}

      {/* Dosya Okuma İlerleme Çubuğu */}
      {importProgress.active && (
        <div className="bg-emerald-50 border-b border-emerald-200 px-4 py-2 flex items-center justify-between text-xs text-emerald-900 z-20">
          <div className="flex items-center gap-2 font-medium">
            <div className="w-3.5 h-3.5 border-2 border-emerald-600 border-t-transparent rounded-full animate-spin" />
            <span>{importProgress.text}</span>
          </div>
          <span className="font-mono font-bold">%{importProgress.percent}</span>
        </div>
      )}

      {/* Takip Modu Çubuğu */}
      {followingUserId && collaborators[followingUserId] && (
        <div className="bg-blue-600 text-white px-4 py-1.5 flex items-center justify-between text-xs z-20">
          <div className="flex items-center gap-2">
            <Eye className="w-4 h-4" />
            <span>
              Şu anda <strong>{collaborators[followingUserId].userName}</strong> adlı kullanıcının çalıştığı hücreyi ({collaborators[followingUserId].activeCellRef}) canlı takip ediyorsunuz.
            </span>
          </div>
          <button
            onClick={() => setFollowingUserId(null)}
            className="px-2.5 py-0.5 rounded bg-white/20 hover:bg-white/30 font-semibold"
          >
            Takibi Bırak
          </button>
        </div>
      )}

      {/* BİLDİRİM TOAST */}
      {statusToast && (
        <div className="fixed bottom-10 right-6 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-lg flex items-center gap-2.5 text-xs font-medium">
          <Sparkles className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{statusToast}</span>
        </div>
      )}

      {/* ANA EXCEL ÇALIŞMA ALANI */}
      <div className="flex-1 flex w-full h-full overflow-hidden relative bg-white">
        <main className="flex-1 flex flex-col h-full overflow-hidden relative">
          <UniverSpreadsheet
            workbookData={workbookData}
            onReady={(handle) => {
              sheetHandleRef.current = handle;
              if (latestRestoredWorkbookRef.current) {
                handle.loadWorkbookSnapshot(latestRestoredWorkbookRef.current);
              }
            }}
            onFxIconClicked={() => {
              handleSelectFormulaInWizard(recentFormulas[0] || 'XLOOKUP');
              setActiveModal('fx');
            }}
            onUserActivity={() => {
              const now = Date.now();
              if (now - lastActivitySentRef.current > 2000) {
                lastActivitySentRef.current = now;
                roomClientRef.current?.sendPresence({
                  userId: currentUser.id,
                  userName: currentUser.name,
                  userColor: currentUser.color,
                  role: currentUser.role,
                  activeSheetId: activeSelection.sheetId,
                  activeSheetName: activeSelection.sheetName,
                  activeCellRef: activeSelection.cellRef,
                  row: activeSelection.row,
                  col: activeSelection.col,
                  lastUpdated: now,
                  lastActivityAt: now,
                  isTyping: true,
                });
              }
            }}
            onSelectionChange={(sel) => {
              const now = Date.now();
              lastActivitySentRef.current = now;
              setActiveSelection({
                sheetId: sel.sheetId,
                sheetName: sel.sheetName,
                cellRef: sel.cellRef,
                row: sel.row,
                col: sel.col,
                value: sel.value,
                formula: sel.formula,
              });
              roomClientRef.current?.sendPresence({
                userId: currentUser.id,
                userName: currentUser.name,
                userColor: currentUser.color,
                role: currentUser.role,
                activeSheetId: sel.sheetId,
                activeSheetName: sel.sheetName,
                activeCellRef: sel.cellRef,
                row: sel.row,
                col: sel.col,
                lastUpdated: now,
                lastActivityAt: now,
              });
            }}
            onCellEdited={(edit) => {
              const colLetter = colIndexToLetter(edit.col);
              const cellRef = formatCellRef(edit.row, edit.col);

              if (currentUser.role === 'viewer') {
                showToast('🔒 Salt Okunur moddasınız, değişiklik yapamazsınız.');
                return false;
              }

              const lockOwner = getActiveCellLockOwner(
                collaborators,
                currentUser.id,
                edit.sheetId,
                edit.row,
                edit.col,
                Date.now(),
                lockTimeoutSec * 1000
              );

              if (lockOwner) {
                showToast(
                  `🔒 ${cellRef} hücresinde şu an ${lockOwner.lockedBy.userName} işlem yapıyor! (${lockOwner.remainingSeconds} sn sonra kilidi açılır)`
                );
                return false;
              }

              if (protectedColumns.includes(colLetter) && !isRoomOwner) {
                showToast(`🛡️ ${colLetter} sütunu kilitlidir, sadece oda sahibi değiştirebilir.`);
                return false;
              }

              // Eğer kullanıcı hücreye doğrudan formül yazdıysa (örn: =XLOOKUP veya =SUM),
              // bu fonksiyonu otomatik olarak "En Son Kullanılanlar" listesinin en üstüne ekle!
              if (edit.formula && edit.formula.startsWith('=')) {
                const updatedRecents = recordFormulaUsage(edit.formula);
                setRecentFormulas(updatedRecents);
              }

              const now = Date.now();
              const autoSnapId = `snap_edit_${now}`;

              // Hücre düzenlendikten hemen sonra tam Workbook görüntüsünü hem tarayıcıya (localStorage)
              // hem de sunucu odasına (.data/rooms.json) kaydet! Böylece F5 yapıldığında asla kaybolmaz!
              setTimeout(() => {
                const latestSnap = sheetHandleRef.current?.getCurrentSnapshot();
                if (latestSnap) {
                  const cloned = JSON.parse(JSON.stringify(latestSnap));
                  latestRestoredWorkbookRef.current = cloned;
                  setLocalSnapshotsMap((prev) => ({
                    ...prev,
                    [autoSnapId]: cloned,
                  }));
                  setOpenWorkbooks((prev) => {
                    const nextTabs = prev.map((t) =>
                      t.id === activeWorkbookTabId ? { ...t, workbookData: cloned } : t
                    );
                    persistWorkspaceToBrowser(cloned, nextTabs, activeWorkbookTabId);
                    return nextTabs;
                  });

                  fetch('/api/rooms', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                      actionType: 'DELTA_SYNC',
                      roomCode,
                      snapshotId: autoSnapId,
                      workbookData: cloned,
                      presence: {
                        userId: currentUser.id,
                        userName: currentUser.name,
                        userColor: currentUser.color,
                        role: currentUser.role,
                        activeSheetId: edit.sheetId,
                        activeSheetName: activeSelection.sheetName,
                        activeCellRef: cellRef,
                        row: edit.row,
                        col: edit.col,
                        lastActivityAt: now,
                      },
                      cellDelta: {
                        senderId: currentUser.id,
                        senderName: currentUser.name,
                        senderColor: currentUser.color,
                        sheetId: edit.sheetId,
                        sheetName: activeSelection.sheetName,
                        row: edit.row,
                        col: edit.col,
                        cellRef,
                        value: edit.value,
                        formula: edit.formula,
                      },
                    }),
                  }).catch(() => {});
                }
              }, 25);

              const timeStr = new Date(now).toLocaleTimeString('tr-TR', {
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
              });

              const autoMeta: SnapshotMeta = {
                id: autoSnapId,
                roomCode,
                label: `${cellRef} düzenlendi (${edit.formula || edit.value})`,
                createdBy: currentUser.name,
                createdByDepartment: currentUser.department,
                createdAt: new Date(now).toISOString(),
                timestamp: now,
                rowCount: stats.totalRows,
                isAutoChange: true,
              };

              setSnapshots((prev) =>
                [autoMeta, ...prev].sort((a, b) => b.timestamp - a.timestamp).slice(0, 60)
              );

              setAuditLog((prev) =>
                [
                  {
                    id: `aud_${now}`,
                    roomCode,
                    userId: currentUser.id,
                    userName: currentUser.name,
                    userEmail: currentUser.email,
                    userDepartment: currentUser.department,
                    userColor: currentUser.color,
                    action: `${cellRef} hücresini düzenledi`,
                    cellRef,
                    sheetName: activeSelection.sheetName,
                    oldValue: edit.previousFormula || edit.previousValue,
                    newValue: edit.formula || edit.value,
                    timestamp: now,
                    snapshotId: autoSnapId,
                  },
                  ...prev,
                ]
                  .sort((a, b) => b.timestamp - a.timestamp)
                  .slice(0, 80)
              );

              roomClientRef.current?.sendCellUpdate({
                senderId: currentUser.id,
                senderName: currentUser.name,
                sheetId: edit.sheetId,
                row: edit.row,
                col: edit.col,
                value: edit.value,
                formula: edit.formula,
              });

              roomClientRef.current?.sendPresence({
                userId: currentUser.id,
                userName: currentUser.name,
                userColor: currentUser.color,
                role: currentUser.role,
                activeSheetId: edit.sheetId,
                activeSheetName: activeSelection.sheetName,
                activeCellRef: cellRef,
                row: edit.row,
                col: edit.col,
                lastUpdated: now,
                lastActivityAt: now,
              });

              setStats((prev) => ({
                ...prev,
                lastSavedTime: timeStr,
                lastSyncText: `${cellRef} otomatik kaydedildi (${timeStr})`,
              }));

              return true;
            }}
          />
        </main>

        {/* SAĞ YAN PANEL 1: KOŞULLU BİÇİMLENDİRME (EXCEL CONDITIONAL FORMATTING) */}
        {activeDrawer === 'conditional' && (
          <aside className="w-96 bg-white border-l border-slate-200 flex flex-col h-full z-20 shrink-0 shadow-xl">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-rose-50/40">
              <div className="flex items-center gap-2">
                <PaintBucket className="w-5 h-5 text-rose-600" />
                <div>
                  <h3 className="text-sm font-bold text-slate-900">Koşullu Biçimlendirme</h3>
                  <p className="text-[11px] text-slate-500">
                    Hücre kurallarını vurgulayın ve renklendirin
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveDrawer('none')}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-4 space-y-4 overflow-y-auto flex-1 text-xs">
              <div className="space-y-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    1. Uygulanacak Sütun / Aralık
                  </label>
                  <select
                    value={cfColIndex}
                    onChange={(e) => setCfColIndex(Number(e.target.value))}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-slate-800 font-medium"
                  >
                    <option value={-1}>Tüm Tablo Hücreleri</option>
                    {extractSheetHeaders(workbookData, activeSelection.sheetId).map((h) => (
                      <option key={h.colIndex} value={h.colIndex}>
                        {h.title} Sütunu
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    2. Vurgulama Kuralı (Koşul)
                  </label>
                  <select
                    value={cfOperator}
                    onChange={(e) => setCfOperator(e.target.value as ConditionalRuleOperator)}
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-slate-800 font-medium"
                  >
                    <option value="gt">Büyüktür (&gt;)</option>
                    <option value="lt">Küçüktür (&lt;)</option>
                    <option value="between">Arasında (İki Değer Arası)</option>
                    <option value="eq">Eşittir (=)</option>
                    <option value="contains">İçerdiği Metin (Metin İçeren)</option>
                    <option value="above_avg">Ortalamanın Üstünde Olanlar</option>
                    <option value="below_avg">Ortalamanın Altında Olanlar</option>
                    <option value="duplicates">Yinelenen (Tekrar Eden) Değerler</option>
                  </select>
                </div>

                {!['above_avg', 'below_avg', 'duplicates'].includes(cfOperator) && (
                  <div className="grid grid-cols-1 gap-2">
                    <div>
                      <label className="block font-semibold text-slate-700 mb-1">
                        {cfOperator === 'between' ? 'Başlangıç Değeri' : 'Koşul Değeri'}
                      </label>
                      <input
                        type="text"
                        value={cfValue1}
                        onChange={(e) => setCfValue1(e.target.value)}
                        placeholder="Örn: 500000 veya Aktif"
                        className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900"
                      />
                    </div>
                    {cfOperator === 'between' && (
                      <div>
                        <label className="block font-semibold text-slate-700 mb-1">
                          Bitiş Değeri
                        </label>
                        <input
                          type="text"
                          value={cfValue2}
                          onChange={(e) => setCfValue2(e.target.value)}
                          placeholder="Örn: 1500000"
                          className="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-slate-900"
                        />
                      </div>
                    )}
                  </div>
                )}

                <div>
                  <label className="block font-semibold text-slate-700 mb-1.5">
                    3. Biçimlendirme Renk Stili
                  </label>
                  <div className="space-y-1.5">
                    {(Object.keys(COLOR_PRESET_MAP) as ConditionalColorPreset[]).map((presetKey) => {
                      const p = COLOR_PRESET_MAP[presetKey];
                      return (
                        <button
                          key={presetKey}
                          type="button"
                          onClick={() => setCfColorPreset(presetKey)}
                          className={`w-full px-3 py-2 rounded-lg border text-left flex items-center justify-between transition cursor-pointer ${
                            cfColorPreset === presetKey
                              ? 'ring-2 ring-slate-900 border-slate-900 font-bold'
                              : 'border-slate-200 opacity-85 hover:opacity-100'
                          }`}
                          style={{ backgroundColor: p.bg, color: p.cl }}
                        >
                          <span>{p.label}</span>
                          {cfColorPreset === presetKey && <Check className="w-4 h-4" />}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <button
                  onClick={handleApplyConditionalRule}
                  className="w-full py-2.5 rounded-xl bg-rose-600 hover:bg-rose-700 text-white font-bold shadow-xs transition cursor-pointer"
                >
                  Koşullu Biçimlendirmeyi Tabloya Uygula
                </button>
              </div>

              {/* Hızlı Hazır Kurallar */}
              <div>
                <h4 className="font-bold text-slate-700 mb-2">Tek Tıkla Hazır Excel Kuralları</h4>
                <div className="grid grid-cols-1 gap-1.5">
                  <button
                    onClick={() => {
                      setCfColIndex(6);
                      setCfOperator('above_avg');
                      setCfColorPreset('green');
                      setTimeout(handleApplyConditionalRule, 30);
                    }}
                    className="p-2.5 rounded-lg border border-emerald-200 bg-emerald-50/60 hover:bg-emerald-100/70 text-emerald-900 text-left font-medium transition cursor-pointer"
                  >
                    🟢 Ortalamanın Üstündeki Bütçeleri Yeşil Vurgula
                  </button>
                  <button
                    onClick={() => {
                      setCfColIndex(6);
                      setCfOperator('below_avg');
                      setCfColorPreset('red');
                      setTimeout(handleApplyConditionalRule, 30);
                    }}
                    className="p-2.5 rounded-lg border border-rose-200 bg-rose-50/60 hover:bg-rose-100/70 text-rose-900 text-left font-medium transition cursor-pointer"
                  >
                    🔴 Ortalamanın Altındaki Bütçeleri Kırmızı Vurgula
                  </button>
                </div>
              </div>

              {activeCfRules.length > 0 && (
                <div>
                  <h4 className="font-bold text-slate-700 mb-2">
                    Uygulanan Kurallar ({activeCfRules.length})
                  </h4>
                  <div className="space-y-2">
                    {activeCfRules.map((rule) => (
                      <div
                        key={rule.id}
                        className="p-2.5 rounded-xl border border-slate-200 bg-slate-50 flex items-center justify-between gap-2"
                      >
                        <div>
                          <div className="font-bold text-slate-800">{rule.colLabel}</div>
                          <div className="text-[11px] text-slate-500">
                            {rule.operatorLabel} • {rule.matchedCount} hücre
                          </div>
                        </div>
                        <button
                          onClick={() =>
                            setActiveCfRules((prev) => prev.filter((r) => r.id !== rule.id))
                          }
                          className="p-1 text-slate-400 hover:text-rose-600"
                          title="Kuralı listeden kaldır"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </aside>
        )}

        {/* SAĞ YAN PANEL 2: VERSİYON GEÇMİŞİ (SON YAPILAN DEĞİŞİKLİK HER ZAMAN EN ÜSTTE!) */}
        {activeDrawer === 'history' && (
          <aside className="w-96 bg-white border-l border-slate-200 flex flex-col h-full z-20 shrink-0 shadow-xl">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-amber-50/50">
              <div className="flex items-center gap-2 min-w-0">
                <History className="w-5 h-5 text-amber-600 shrink-0" />
                <div className="min-w-0">
                  <h3 className="text-sm font-bold text-slate-900">
                    Versiyon & Değişiklik Geçmişi
                  </h3>
                  <p className="text-[11px] text-slate-500 truncate">
                    Yalnızca aktif oda ({roomCode}) • {workbookData.name}
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveDrawer('none')}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-4 space-y-4 overflow-y-auto flex-1">
              {/* Aktif Oda Rozeti */}
              <div className="bg-amber-50/70 border border-amber-200/80 rounded-xl px-3 py-2 flex items-center justify-between gap-2 text-[11px]">
                <span className="text-amber-900 font-semibold truncate">
                  📂 {workbookData.name}
                </span>
                <span className="font-mono font-bold bg-white text-amber-800 border border-amber-300 px-2 py-0.5 rounded-md shrink-0">
                  {roomCode}
                </span>
              </div>

              {/* Manuel Versiyon Oluşturma */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-2">
                <div className="text-xs font-bold text-slate-700">
                  Mevcut Durumu İsimlendirerek Yedekle
                </div>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={newSnapshotLabel}
                    onChange={(e) => setNewSnapshotLabel(e.target.value)}
                    placeholder="Örn: Toplantı Öncesi Final"
                    className="flex-1 bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-emerald-600"
                  />
                  <button
                    onClick={handleCreateSnapshot}
                    className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shrink-0 cursor-pointer"
                  >
                    Yedek Al
                  </button>
                </div>
              </div>

              {/* 1. EN ÜSTTE: SON YAPILAN DEĞİŞİKLİKLER VE KAYITLI VERSİYONLAR (YENİDEN ESKİYE SIRALI) */}
              <div>
                <div className="flex items-center justify-between mb-2">
                  <h4 className="text-xs font-bold text-slate-800">
                    Bu Odadaki Versiyonlar ({sortedSnapshots.length})
                  </h4>
                  <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full">
                    En Yeni En Üstte
                  </span>
                </div>

                {sortedSnapshots.length === 0 ? (
                  <div className="p-3 rounded-xl bg-slate-50 border border-dashed border-slate-200 text-[11px] text-slate-500 text-center">
                    Bu Excel odasında ({roomCode}) henüz kayıtlı bir versiyon bulunmuyor. Hücre düzenlediğinizde veya &quot;Yedek Al&quot; dediğinizde burada listelenir.
                  </div>
                ) : (
                  <div className="space-y-2">
                    {sortedSnapshots.map((snap, index) => (
                      <div
                        key={snap.id}
                        className={`p-3 rounded-xl border flex items-center justify-between gap-2 transition ${
                          index === 0
                            ? 'bg-emerald-50/70 border-emerald-300 shadow-2xs'
                            : 'bg-slate-50 border-slate-200 hover:bg-slate-100/70'
                        }`}
                      >
                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5">
                            {index === 0 && (
                              <span className="px-1.5 py-0.5 rounded bg-emerald-600 text-white text-[9px] font-bold uppercase shrink-0">
                                EN SON
                              </span>
                            )}
                            <span className="text-xs font-bold text-slate-800 truncate">
                              {snap.label}
                            </span>
                          </div>
                          <div className="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1.5">
                            <span className="font-semibold text-slate-700">{snap.createdBy}</span>
                            {snap.createdByDepartment && (
                              <span className="text-[10px] text-slate-400">
                                ({snap.createdByDepartment})
                              </span>
                            )}
                            <span>•</span>
                            <span className="font-mono text-[10px]">
                              {new Date(snap.timestamp).toLocaleTimeString('tr-TR', {
                                hour: '2-digit',
                                minute: '2-digit',
                                second: '2-digit',
                              })}
                            </span>
                          </div>
                        </div>

                        <button
                          onClick={() => handleRestoreSnapshotById(snap.id, snap.label)}
                          className="px-2.5 py-1.5 rounded-lg bg-white hover:bg-amber-50 text-amber-800 border border-slate-300 hover:border-amber-400 text-[11px] font-bold flex items-center gap-1 shrink-0 shadow-2xs cursor-pointer transition"
                          title="Tabloyu eksiksiz olarak bu versiyondaki haline geri döndür"
                        >
                          <RotateCcw className="w-3.5 h-3.5 text-amber-600" />
                          <span>Bu Haline Dön</span>
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* 2. KULLANICI İŞLEM LOGU (KİMLİK TAKİPLİ) */}
              <div className="pt-2 border-t border-slate-200">
                <h4 className="text-xs font-bold text-slate-700 mb-2">
                  Bu Odadaki Hücre İşlem Kayıtları ({sortedAuditLog.length})
                </h4>
                {sortedAuditLog.length === 0 ? (
                  <div className="p-3 rounded-xl bg-slate-50 border border-dashed border-slate-200 text-[11px] text-slate-500 text-center">
                    Bu odada ({roomCode}) henüz hücre değişikliği kaydı yok.
                  </div>
                ) : (
                  <div className="space-y-2">
                    {sortedAuditLog.map((ev, idx) => (
                      <div
                        key={ev.id}
                        className="p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs space-y-1"
                      >
                        <div className="flex items-center justify-between">
                          <span className="font-bold text-slate-800 flex items-center gap-1.5">
                            <span
                              className="w-2.5 h-2.5 rounded-full"
                              style={{ backgroundColor: ev.userColor }}
                            />
                            {ev.userName}
                            {idx === 0 && (
                              <span className="text-[9px] bg-blue-100 text-blue-700 px-1.5 py-0.2 rounded font-bold">
                                Az Önce
                              </span>
                            )}
                          </span>
                          <span className="text-[10px] font-mono text-slate-400">
                            {new Date(ev.timestamp).toLocaleTimeString('tr-TR')}
                          </span>
                        </div>
                        <p className="text-slate-600 text-[11px]">{ev.action}</p>
                        {ev.newValue !== undefined && (
                          <div className="flex items-center justify-between gap-2 pt-0.5">
                            <div className="text-[11px] font-mono bg-white border border-slate-200 px-2 py-0.5 rounded text-emerald-700 truncate flex-1">
                              {ev.oldValue !== undefined ? `${ev.oldValue} ➔ ` : ''}
                              {String(ev.newValue)}
                            </div>
                            {ev.snapshotId && (
                              <button
                                onClick={() =>
                                  handleRestoreSnapshotById(ev.snapshotId!, ev.action)
                                }
                                className="text-[10px] font-semibold text-amber-700 hover:underline shrink-0 cursor-pointer"
                              >
                                Buraya Dön
                              </button>
                            )}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </aside>
        )}

        {/* SAĞ YAN PANEL 3: AKILLI HÜCRE KİLİDİ SÜRESİ & SÜTUN KİLİTLEME */}
        {activeDrawer === 'protection' && (
          <aside className="w-88 bg-white border-l border-slate-200 flex flex-col h-full z-20 shrink-0 shadow-lg">
            <div className="p-4 border-b border-slate-100 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Lock className="w-5 h-5 text-rose-600" />
                <div>
                  <h3 className="text-sm font-bold text-slate-800">Hücre & Sütun Kilidi</h3>
                  <p className="text-[11px] text-slate-500">
                    Çakışma koruması ve unutulan ekran zaman aşımı
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveDrawer('none')}
                className="text-slate-400 hover:text-slate-700 p-1 rounded-lg hover:bg-slate-100"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-4 space-y-5 overflow-y-auto flex-1">
              <div className="bg-amber-50/70 border border-amber-200 rounded-xl p-3.5 space-y-2.5">
                <div className="flex items-center gap-1.5 text-xs font-bold text-amber-900">
                  <Clock className="w-4 h-4 text-amber-600" />
                  <span>Unutulan Hücre Kilidi Açılma Süresi</span>
                </div>
                <p className="text-[11px] text-amber-800 leading-relaxed">
                  Bir kullanıcı bir hücrede işlem yaparken diğer kullanıcılar o hücreyi değiştiremez. Kişi ekranını o hücrede açık unutursa aşağıdaki süre sonunda kilit otomatik düşer:
                </p>
                <div className="grid grid-cols-3 gap-2 pt-1">
                  {[15, 30, 60].map((sec) => (
                    <button
                      key={sec}
                      onClick={() => {
                        setLockTimeoutSec(sec);
                        showToast(`⏱️ Hücre hareketsizlik kilit süresi ${sec} saniye olarak ayarlandı.`);
                      }}
                      className={`py-1.5 rounded-lg text-xs font-semibold border transition cursor-pointer ${
                        lockTimeoutSec === sec
                          ? 'bg-amber-600 text-white border-amber-600'
                          : 'bg-white text-slate-700 border-amber-200 hover:bg-amber-100/50'
                      }`}
                    >
                      {sec} Saniye
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <h4 className="text-xs font-semibold text-slate-700 mb-2">
                  Sadece Oda Sahibinin Değiştirebileceği Sütunlar
                </h4>
                <div className="grid grid-cols-4 gap-2">
                  {['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'].map((col) => {
                    const isLocked = protectedColumns.includes(col);
                    return (
                      <button
                        key={col}
                        onClick={() => {
                          if (isLocked) {
                            setProtectedColumns(protectedColumns.filter((c) => c !== col));
                          } else {
                            setProtectedColumns([...protectedColumns, col]);
                          }
                        }}
                        className={`py-2 px-3 rounded-lg font-mono text-xs font-bold border flex items-center justify-center gap-1 transition cursor-pointer ${
                          isLocked
                            ? 'bg-rose-50 border-rose-300 text-rose-700'
                            : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        {isLocked && <Lock className="w-3 h-3" />}
                        <span>{col}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </aside>
        )}
      </div>

      {/* MODAL 1: EXCEL FX FONKSİYON SİHİRBAZI (KOLAY XLOOKUP + EN SON KULLANILANLAR) */}
      {activeModal === 'fx' && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-4xl w-full max-h-[90vh] flex flex-col shadow-2xl overflow-hidden">
            {/* Üst Başlık */}
            <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-9 h-9 rounded-xl bg-indigo-600 flex items-center justify-center font-mono font-extrabold text-base">
                  fx
                </div>
                <div>
                  <h3 className="text-base font-bold">
                    Excel Fonksiyon Sihirbazı (Kolay XLOOKUP & Formül Oluşturucu)
                  </h3>
                  <p className="text-xs text-slate-300">
                    Hedef Hücre:{' '}
                    <span className="font-mono font-bold text-emerald-400">
                      {activeSelection.cellRef}
                    </span>{' '}
                    • Formüller İngilizce Excel standardındadır, açıklamalar Türkçedir.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-white p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="flex-1 grid grid-cols-1 md:grid-cols-12 overflow-hidden">
              {/* SOL SÜTUN: KATEGORİLER VE FONKSİYON LİSTESİ (EN SON KULLANILANLAR DAHİL) */}
              <div className="md:col-span-5 border-r border-slate-200 bg-slate-50 p-4 flex flex-col overflow-y-auto">
                {/* Arama Kutusu */}
                <div className="relative mb-3">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2.5" />
                  <input
                    type="text"
                    value={fxSearchQuery}
                    onChange={(e) => {
                      setFxSearchQuery(e.target.value);
                      if (e.target.value.trim()) setFxCategory('all');
                    }}
                    placeholder="Fonksiyon ara (XLOOKUP, SUM, IF...)"
                    className="w-full bg-white border border-slate-300 rounded-xl pl-9 pr-3 py-2 text-xs text-slate-800 focus:outline-none focus:border-indigo-600"
                  />
                </div>

                {/* Kategori Sekmeleri */}
                <div className="flex flex-wrap gap-1 mb-3">
                  {[
                    { id: 'recent', label: '⭐ En Son Kullanılanlar' },
                    { id: 'lookup', label: '🔍 Arama (XLOOKUP)' },
                    { id: 'math', label: '🧮 Matematik' },
                    { id: 'logical', label: '⚖️ Mantıksal' },
                    { id: 'stats', label: '📊 İstatistik' },
                    { id: 'all', label: 'Tüm Fonksiyonlar' },
                  ].map((cat) => (
                    <button
                      key={cat.id}
                      onClick={() => setFxCategory(cat.id as any)}
                      className={`px-2.5 py-1 rounded-lg text-[11px] font-semibold transition cursor-pointer ${
                        fxCategory === cat.id
                          ? 'bg-indigo-600 text-white'
                          : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
                      }`}
                    >
                      {cat.label}
                    </button>
                  ))}
                </div>

                {/* Fonksiyon Listesi */}
                <div className="space-y-1.5 flex-1 overflow-y-auto pr-1">
                  {(() => {
                    let list = EXCEL_FORMULA_CATALOG;
                    if (fxSearchQuery.trim()) {
                      const q = fxSearchQuery.toLowerCase();
                      list = EXCEL_FORMULA_CATALOG.filter(
                        (f) =>
                          f.name.toLowerCase().includes(q) ||
                          f.shortDescTr.toLowerCase().includes(q)
                      );
                    } else if (fxCategory === 'recent') {
                      list = recentFormulas
                        .map((rName) => EXCEL_FORMULA_CATALOG.find((f) => f.name === rName))
                        .filter(Boolean) as typeof EXCEL_FORMULA_CATALOG;
                    } else if (fxCategory !== 'all') {
                      list = EXCEL_FORMULA_CATALOG.filter((f) => f.category === fxCategory);
                    }

                    return list.map((fn) => {
                      const isSelected = fn.name === selectedFormulaName;
                      const isRecent = recentFormulas.includes(fn.name);
                      return (
                        <button
                          key={fn.name}
                          onClick={() => handleSelectFormulaInWizard(fn.name)}
                          className={`w-full p-2.5 rounded-xl text-left border transition cursor-pointer ${
                            isSelected
                              ? 'bg-indigo-50 border-indigo-500 ring-1 ring-indigo-500'
                              : 'bg-white border-slate-200 hover:border-indigo-300'
                          }`}
                        >
                          <div className="flex items-center justify-between">
                            <span className="font-mono font-bold text-xs text-indigo-950 flex items-center gap-1.5">
                              {fn.name}
                              {fn.name === 'XLOOKUP' && (
                                <span className="px-1.5 py-0.2 rounded bg-emerald-100 text-emerald-800 font-sans text-[10px] font-bold">
                                  Önerilen
                                </span>
                              )}
                            </span>
                            {isRecent && (
                              <span className="text-[10px] text-amber-600 flex items-center gap-0.5 font-medium">
                                <Star className="w-3 h-3 fill-amber-400 text-amber-500" />
                                Sık Kullanılan
                              </span>
                            )}
                          </div>
                          <p className="text-[11px] text-slate-600 mt-1 line-clamp-2">
                            {fn.shortDescTr}
                          </p>
                        </button>
                      );
                    });
                  })()}
                </div>
              </div>

              {/* SAĞ SÜTUN: SEÇİLEN FONKSİYONUN ADIM ADIM KOLAY PARAMETRE SİHİRBAZI */}
              <div className="md:col-span-7 p-5 flex flex-col justify-between overflow-y-auto bg-white">
                <div className="space-y-4">
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-sm font-mono font-extrabold text-indigo-700">
                        {selectedFormulaDef.syntax}
                      </span>
                      <span className="text-[11px] bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-full font-semibold">
                        {selectedFormulaDef.categoryLabelTr}
                      </span>
                    </div>
                    <p className="text-xs text-slate-600 mt-1.5">
                      {selectedFormulaDef.shortDescTr}
                    </p>
                  </div>

                  {(selectedFormulaName === 'XLOOKUP' || selectedFormulaName === 'VLOOKUP') && (
                    <div className="bg-emerald-50 border border-emerald-300 rounded-xl p-3 flex items-center justify-between gap-3">
                      <div className="text-xs text-emerald-950">
                        <strong className="font-bold flex items-center gap-1.5">
                          <Link2 className="w-4 h-4 text-emerald-700" />
                          Bilgisayarınızda Açık Gerçek Excel Dosyası ile {selectedFormulaName} Yapmak İster misiniz?
                        </strong>
                        <span className="text-[11px] text-emerald-800 block mt-0.5">
                          Masaüstünüzdeki açık <code>.xlsx</code> dosyasını canlı bağlayıp (Excel&apos;de Ctrl+S yaptıkça otomatik güncellenir) sütunları tek tıkla eşleştirebilirsiniz.
                        </span>
                      </div>
                      <button
                        onClick={() => {
                          setBridgeFormulaType(selectedFormulaName as 'XLOOKUP' | 'VLOOKUP');
                          setActiveModal('excel_bridge');
                        }}
                        className="px-3 py-2 rounded-xl bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold shrink-0 cursor-pointer transition"
                      >
                        Masaüstü Excel Köprüsünü Aç →
                      </button>
                    </div>
                  )}

                  {/* Sütun Hızlı Seçim Yardımı */}
                  {currentSheetColumns.length > 0 && (
                    <div className="bg-emerald-50/60 border border-emerald-200 rounded-xl p-3">
                      <div className="text-[11px] font-bold text-emerald-900 mb-1.5">
                        💡 Tablonuzdaki Sütunlar (Kutulara tıklamak yerine aşağıdaki sütun butonlarıyla hızlıca seçebilirsiniz):
                      </div>
                      <div className="flex flex-wrap gap-1.5">
                        {currentSheetColumns.map((col) => (
                          <span
                            key={col.colIndex}
                            className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-white border border-emerald-300 text-[11px] text-emerald-900 font-medium"
                          >
                            <strong className="font-mono">{col.colLetter}:</strong> {col.headerName}{' '}
                            <span className="text-slate-400 font-mono">({col.rangeRef})</span>
                          </span>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Parametre Giriş Alanları */}
                  <div className="space-y-3">
                    {selectedFormulaDef.args.map((arg) => (
                      <div
                        key={arg.name}
                        className="p-3 rounded-xl border border-slate-200 bg-slate-50/50 space-y-1.5"
                      >
                        <div className="flex items-center justify-between">
                          <label className="text-xs font-bold text-slate-800">
                            {arg.labelTr}{' '}
                            {arg.required ? (
                              <span className="text-rose-600">*</span>
                            ) : (
                              <span className="text-slate-400 font-normal">(İsteğe Bağlı)</span>
                            )}
                          </label>
                        </div>
                        <p className="text-[11px] text-slate-500">{arg.descriptionTr}</p>

                        {arg.options ? (
                          <select
                            value={fxArgValues[arg.name] || ''}
                            onChange={(e) =>
                              setFxArgValues((prev) => ({ ...prev, [arg.name]: e.target.value }))
                            }
                            className="w-full bg-white border border-slate-300 rounded-lg px-3 py-2 text-xs font-mono text-slate-900"
                          >
                            {arg.options.map((opt) => (
                              <option key={opt.value} value={opt.value}>
                                {opt.label}
                              </option>
                            ))}
                          </select>
                        ) : (
                          <input
                            type="text"
                            value={fxArgValues[arg.name] || ''}
                            onChange={(e) =>
                              setFxArgValues((prev) => ({ ...prev, [arg.name]: e.target.value }))
                            }
                            placeholder={arg.placeholder}
                            className="w-full bg-white border border-slate-300 rounded-lg px-3 py-2 text-xs font-mono text-slate-900 focus:outline-none focus:border-indigo-600"
                          />
                        )}

                        {/* Sütun Aralığı veya Hücre Hızlı Seçim Butonları */}
                        {currentSheetColumns.length > 0 && (
                          <div className="flex items-center gap-1 flex-wrap pt-1">
                            <span className="text-[10px] text-slate-400 mr-1">Hızlı Seç:</span>
                            {currentSheetColumns.map((col) => (
                              <button
                                key={col.colIndex}
                                type="button"
                                onClick={() =>
                                  setFxArgValues((prev) => ({
                                    ...prev,
                                    [arg.name]: arg.isRange ? col.rangeRef : col.sampleCellRef,
                                  }))
                                }
                                className="px-2 py-0.5 rounded bg-white hover:bg-indigo-50 text-slate-700 hover:text-indigo-700 border border-slate-200 text-[10px] font-mono transition cursor-pointer"
                                title={`${col.headerName} sütununu seç`}
                              >
                                {col.colLetter} ({col.headerName.slice(0, 12)})
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>

                {/* Alt Canlı Önizleme ve Uygula Butonu */}
                <div className="mt-5 pt-4 border-t border-slate-200 space-y-3">
                  <div className="bg-slate-900 text-white p-3.5 rounded-xl flex items-center justify-between gap-4">
                    <div className="min-w-0">
                      <div className="text-[10px] uppercase tracking-wider text-slate-400 font-semibold">
                        Oluşturulan İngilizce Excel Formülü
                      </div>
                      <div className="text-sm font-mono font-bold text-emerald-400 truncate mt-0.5">
                        {liveFormulaPreview.formulaString}
                      </div>
                    </div>
                    <div className="text-right shrink-0 bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-700">
                      <div className="text-[10px] text-slate-400">Canlı Hesaplama Sonucu</div>
                      <div className="text-sm font-mono font-bold text-white">
                        {liveFormulaPreview.computedValue !== ''
                          ? typeof liveFormulaPreview.computedValue === 'number'
                            ? liveFormulaPreview.computedValue.toLocaleString('tr-TR')
                            : String(liveFormulaPreview.computedValue)
                          : '—'}
                      </div>
                    </div>
                  </div>

                  <div className="flex justify-end gap-2">
                    <button
                      onClick={() => setActiveModal(null)}
                      className="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold cursor-pointer"
                    >
                      İptal
                    </button>
                    <button
                      onClick={handleInsertFormulaToActiveCell}
                      className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold shadow-sm cursor-pointer"
                    >
                      ✅ {activeSelection.cellRef} Hücresine Formülü Uygula
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: EXCEL GİBİ PİVOT TABLO OLUŞTURUCU */}
      {activeModal === 'pivot_modal' && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-5xl w-full max-h-[90vh] flex flex-col shadow-2xl overflow-hidden">
            <div className="px-6 py-4 bg-[#1e3a8a] text-white flex items-center justify-between">
              <div className="flex items-center gap-3">
                <Table2 className="w-6 h-6 text-blue-200" />
                <div>
                  <h3 className="text-base font-bold">
                    Excel Pivot Tablo Oluşturucu (Özet Tablo & Çapraz Analiz)
                  </h3>
                  <p className="text-xs text-blue-200">
                    Satır, Sütun, Değer ve Filtre alanlarını seçin; ister canlı inceleyin ister yeni çalışma sayfası (sekme) olarak ekleyin.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveModal(null)}
                className="text-blue-200 hover:text-white p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="flex-1 grid grid-cols-1 md:grid-cols-12 overflow-hidden">
              {/* SOL: PİVOT ALAN SEÇİMLERİ (EXCEL PIVOTTABLE FIELDS) */}
              <div className="md:col-span-4 border-r border-slate-200 bg-slate-50 p-4 space-y-4 overflow-y-auto text-xs">
                <div className="font-bold text-slate-800 text-sm border-b border-slate-200 pb-2">
                  Pivot Tablo Alanlarını Seçin
                </div>

                {/* 1. Satır Alanı */}
                <div className="bg-white p-3 rounded-xl border border-slate-200 space-y-1.5">
                  <label className="block font-bold text-slate-800">
                    1. Satır Alanı (Gruplanacak Sütun)
                  </label>
                  <p className="text-[11px] text-slate-500">
                    Verilerin hangi sütuna göre gruplanacağını seçin
                  </p>
                  <select
                    value={pivotConfig.rowColIdx}
                    onChange={(e) =>
                      setPivotConfig((prev) => ({ ...prev, rowColIdx: Number(e.target.value) }))
                    }
                    className="w-full bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-2 text-slate-900 font-semibold"
                  >
                    {livePivotResult.headers.map((h) => (
                      <option key={h.colIndex} value={h.colIndex}>
                        {h.title}
                      </option>
                    ))}
                  </select>
                </div>

                {/* 2. Sütun Kırılım Alanı */}
                <div className="bg-white p-3 rounded-xl border border-slate-200 space-y-1.5">
                  <label className="block font-bold text-slate-800">
                    2. Sütun Kırılımı (İsteğe Bağlı Matris)
                  </label>
                  <p className="text-[11px] text-slate-500">
                    Değerleri yan yana sütunlara bölmek için seçin (Örn: DURUM veya BÖLGE)
                  </p>
                  <select
                    value={pivotConfig.colColIdx === null ? -1 : pivotConfig.colColIdx}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      setPivotConfig((prev) => ({
                        ...prev,
                        colColIdx: val === -1 ? null : val,
                      }));
                    }}
                    className="w-full bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-2 text-slate-900 font-semibold"
                  >
                    <option value={-1}>(Yok — Tek Boyutlu Özet Tablo)</option>
                    {livePivotResult.headers.map((h) => (
                      <option key={h.colIndex} value={h.colIndex}>
                        {h.title}
                      </option>
                    ))}
                  </select>
                </div>

                {/* 3. Değer Alanı ve Fonksiyon */}
                <div className="bg-white p-3 rounded-xl border border-slate-200 space-y-2">
                  <label className="block font-bold text-slate-800">
                    3. Hesaplanacak Değer Sütunu ($\Sigma$ Değerler)
                  </label>
                  <select
                    value={pivotConfig.valColIdx}
                    onChange={(e) =>
                      setPivotConfig((prev) => ({ ...prev, valColIdx: Number(e.target.value) }))
                    }
                    className="w-full bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-2 text-slate-900 font-semibold"
                  >
                    {livePivotResult.headers.map((h) => (
                      <option key={h.colIndex} value={h.colIndex}>
                        {h.title}
                      </option>
                    ))}
                  </select>

                  <label className="block font-bold text-slate-700 pt-1">
                    Hesaplama Türü (Fonksiyon)
                  </label>
                  <div className="grid grid-cols-3 gap-1.5">
                    {(
                      [
                        { id: 'SUM', label: 'Toplam' },
                        { id: 'AVERAGE', label: 'Ortalama' },
                        { id: 'COUNT', label: 'Sayım (Adet)' },
                        { id: 'MAX', label: 'Maksimum' },
                        { id: 'MIN', label: 'Minimum' },
                      ] as { id: PivotAggFunc; label: string }[]
                    ).map((fn) => (
                      <button
                        key={fn.id}
                        type="button"
                        onClick={() => setPivotConfig((prev) => ({ ...prev, aggFunc: fn.id }))}
                        className={`py-1.5 px-2 rounded-lg text-[11px] font-bold border transition cursor-pointer ${
                          pivotConfig.aggFunc === fn.id
                            ? 'bg-blue-600 text-white border-blue-600'
                            : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                        }`}
                      >
                        {fn.label}
                      </button>
                    ))}
                  </div>
                </div>

                {/* 4. Filtre Alanı */}
                <div className="bg-white p-3 rounded-xl border border-slate-200 space-y-2">
                  <label className="block font-bold text-slate-800">
                    4. Üst Filtre (İsteğe Bağlı)
                  </label>
                  <select
                    value={pivotConfig.filterColIdx === null ? -1 : pivotConfig.filterColIdx}
                    onChange={(e) => {
                      const val = Number(e.target.value);
                      setPivotConfig((prev) => ({
                        ...prev,
                        filterColIdx: val === -1 ? null : val,
                        filterValue: 'ALL',
                      }));
                    }}
                    className="w-full bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-slate-800"
                  >
                    <option value={-1}>(Filtre Yok — Tüm Kayıtlar)</option>
                    {livePivotResult.headers.map((h) => (
                      <option key={h.colIndex} value={h.colIndex}>
                        {h.title}
                      </option>
                    ))}
                  </select>

                  {pivotConfig.filterColIdx !== null && (
                    <select
                      value={pivotConfig.filterValue}
                      onChange={(e) =>
                        setPivotConfig((prev) => ({ ...prev, filterValue: e.target.value }))
                      }
                      className="w-full bg-blue-50 border border-blue-300 rounded-lg px-2.5 py-1.5 text-blue-900 font-semibold"
                    >
                      <option value="ALL">Tüm Değerleri Göster</option>
                      {livePivotResult.filterOptions.map((opt) => (
                        <option key={opt} value={opt}>
                          Sadece: {opt}
                        </option>
                      ))}
                    </select>
                  )}
                </div>
              </div>

              {/* SAĞ: CANLI PİVOT TABLO ÖNİZLEMESİ VE YENİ SAYFAYA EKLEME */}
              <div className="md:col-span-8 p-5 flex flex-col justify-between overflow-y-auto bg-white">
                <div className="space-y-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <h4 className="text-sm font-bold text-slate-900">
                        Canlı Pivot Tablo Önizlemesi ({livePivotResult.rows.length} Grup,{' '}
                        {livePivotResult.totalRecordCount} Kayıt)
                      </h4>
                      <p className="text-xs text-slate-500">
                        Genel Sonuç:{' '}
                        <strong className="text-blue-700 font-mono">
                          {livePivotResult.grandTotal.toLocaleString('tr-TR')}
                        </strong>
                      </p>
                    </div>
                    <button
                      onClick={handleCreatePivotSheet}
                      className="px-4 py-2 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold flex items-center gap-1.5 shadow-sm transition cursor-pointer"
                    >
                      <Plus className="w-4 h-4" />
                      <span>Yeni Sayfa Sekmesi Olarak Tabloya Ekle</span>
                    </button>
                  </div>

                  {/* Pivot Tablo Matrisi */}
                  <div className="border border-slate-200 rounded-xl overflow-x-auto max-h-72">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead className="bg-slate-900 text-white sticky top-0">
                        <tr>
                          <th className="py-2.5 px-3 font-bold border-r border-slate-700">
                            {livePivotResult.headers.find((h) => h.colIndex === pivotConfig.rowColIdx)
                              ?.title || 'Satır Etiketleri'}
                          </th>
                          {livePivotResult.uniqueColKeys.map((cKey) => (
                            <th
                              key={cKey}
                              className="py-2.5 px-3 text-right font-bold border-r border-slate-700"
                            >
                              {cKey}
                            </th>
                          ))}
                          {livePivotResult.uniqueColKeys.length > 1 && (
                            <th className="py-2.5 px-3 text-right font-bold bg-blue-900">
                              Genel Toplam
                            </th>
                          )}
                          <th className="py-2.5 px-3 text-right font-semibold text-slate-300">
                            Kayıt Adedi
                          </th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {livePivotResult.rows.map((rItem) => (
                          <tr key={rItem.rowKey} className="hover:bg-blue-50/40">
                            <td className="py-2 px-3 font-semibold text-slate-800 border-r border-slate-100">
                              {rItem.rowKey}
                            </td>
                            {livePivotResult.uniqueColKeys.map((cKey) => (
                              <td
                                key={cKey}
                                className="py-2 px-3 text-right font-mono text-slate-700 border-r border-slate-100"
                              >
                                {(rItem.colValues[cKey] ?? 0).toLocaleString('tr-TR')}
                              </td>
                            ))}
                            {livePivotResult.uniqueColKeys.length > 1 && (
                              <td className="py-2 px-3 text-right font-mono font-bold text-blue-800 bg-blue-50/50">
                                {rItem.rowTotal.toLocaleString('tr-TR')}
                              </td>
                            )}
                            <td className="py-2 px-3 text-right font-mono text-slate-500">
                              {rItem.rowCount}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot className="bg-blue-50 border-t-2 border-blue-200 font-bold text-blue-950">
                        <tr>
                          <td className="py-2.5 px-3">GENEL TOPLAM</td>
                          {livePivotResult.uniqueColKeys.map((cKey) => (
                            <td key={cKey} className="py-2.5 px-3 text-right font-mono">
                              {(livePivotResult.colTotals[cKey] ?? 0).toLocaleString('tr-TR')}
                            </td>
                          ))}
                          {livePivotResult.uniqueColKeys.length > 1 && (
                            <td className="py-2.5 px-3 text-right font-mono text-blue-700">
                              {livePivotResult.grandTotal.toLocaleString('tr-TR')}
                            </td>
                          )}
                          <td className="py-2.5 px-3 text-right font-mono">
                            {livePivotResult.totalRecordCount}
                          </td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  {/* Görsel Dağılım Grafiği */}
                  <div className="bg-slate-50 border border-slate-200 rounded-xl p-4">
                    <h5 className="text-xs font-bold text-slate-700 mb-3 flex items-center gap-1.5">
                      <BarChart3 className="w-4 h-4 text-blue-600" />
                      <span>Pivot Grafik Dağılımı</span>
                    </h5>
                    <div className="space-y-2">
                      {livePivotResult.rows.slice(0, 6).map((rItem, idx) => {
                        const maxVal = Math.max(
                          ...livePivotResult.rows.map((x) => Math.abs(x.rowTotal)),
                          1
                        );
                        const pct = Math.max(
                          6,
                          Math.round((Math.abs(rItem.rowTotal) / maxVal) * 100)
                        );
                        const barColors = [
                          'bg-blue-600',
                          'bg-emerald-600',
                          'bg-indigo-600',
                          'bg-amber-500',
                          'bg-rose-500',
                          'bg-purple-600',
                        ];
                        return (
                          <div key={rItem.rowKey} className="space-y-1">
                            <div className="flex justify-between text-[11px]">
                              <span className="font-semibold text-slate-700 truncate max-w-[260px]">
                                {rItem.rowKey}
                              </span>
                              <span className="font-mono font-bold text-slate-800">
                                {rItem.rowTotal.toLocaleString('tr-TR')}
                              </span>
                            </div>
                            <div className="w-full h-2.5 bg-slate-200 rounded-full overflow-hidden">
                              <div
                                className={`${barColors[idx % barColors.length]} h-full rounded-full`}
                                style={{ width: `${pct}%` }}
                              />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>

                <div className="pt-4 border-t border-slate-200 flex justify-end gap-2">
                  <button
                    onClick={() => setActiveModal(null)}
                    className="px-4 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold cursor-pointer"
                  >
                    Kapat
                  </button>
                  <button
                    onClick={handleCreatePivotSheet}
                    className="px-5 py-2 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold cursor-pointer"
                  >
                    ➕ Yeni Sayfa Sekmesi Olarak Tabloya Ekle
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: NASIL KAYDEDİLİYOR BİLGİLENDİRME MODALI */}
      {activeModal === 'save_info' && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-md w-full p-6 shadow-2xl">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2.5">
                <div className="p-2 rounded-xl bg-emerald-50 text-emerald-600">
                  <ShieldCheck className="w-5 h-5" />
                </div>
                <h3 className="text-base font-bold text-slate-900">
                  Kayıt Sistemi Nasıl Çalışır?
                </h3>
              </div>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="space-y-3 text-xs text-slate-600 leading-relaxed">
              <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                <strong className="text-slate-900 block mb-1">
                  1. Otomatik Sistem & Oda Kaydı (Kaydet Butonu):
                </strong>
                Tabloda yaptığınız her hücre değişikliği anında ortak çalışma odasına (
                <code className="font-mono text-emerald-700">{roomCode}</code>) ve tarayıcı hafızasına otomatik kaydedilir. Böylece sayfayı yenileseniz bile verileriniz kaybolmaz ve aynı odadaki ekip arkadaşlarınız değişiklikleri anında görür.
              </div>
              <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                <strong className="text-slate-900 block mb-1">
                  2. Bilgisayara Excel İndir (.xlsx Butonu):
                </strong>
                Çalışmanız bittiğinde dosyayı kendi bilgisayarınızın Masaüstü / İndirilenler klasörüne standart Microsoft Excel (<code>.xlsx</code>) dosyası olarak kaydetmek için üstteki <strong>&quot;Bilgisayara İndir (.xlsx)&quot;</strong> butonuna tıklamanız yeterlidir.
              </div>
            </div>
            <button
              onClick={() => setActiveModal(null)}
              className="w-full mt-5 py-2.5 rounded-xl bg-[#107c41] text-white text-xs font-bold cursor-pointer"
            >
              Anladım, Kapat
            </button>
          </div>
        </div>
      )}

      {/* MODAL: PAYLAŞ */}
      {activeModal === 'share' && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-md w-full p-6 shadow-2xl">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2.5">
                <div className="p-2 rounded-xl bg-emerald-50 text-emerald-600">
                  <Share2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-900">
                    Bu Excel Tablosunu Paylaş
                  </h3>
                  <p className="text-xs text-slate-500">
                    Ekip arkadaşlarınızla aynı anda birlikte çalışın
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-4">
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-center">
                <p className="text-[11px] uppercase tracking-wider font-semibold text-slate-400 mb-1">
                  Oda Paylaşım Kodu
                </p>
                <div className="text-2xl font-mono font-extrabold text-emerald-700 tracking-widest">
                  {roomCode}
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Katılımcı Yetkisi
                  </label>
                  <select
                    value={defaultRoomRole}
                    onChange={(e) =>
                      setDefaultRoomRole(e.target.value as 'editor' | 'viewer')
                    }
                    className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs text-slate-800"
                  >
                    <option value="editor">Düzenleyebilir (Editör)</option>
                    <option value="viewer">Sadece Okuyabilir</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    Oda Şifresi (İsteğe Bağlı)
                  </label>
                  <input
                    type="password"
                    value={roomPin}
                    onChange={(e) => setRoomPin(e.target.value)}
                    placeholder="Şifresiz"
                    className="w-full bg-white border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-semibold text-slate-700 mb-1">
                  Doğrudan Davet Linki
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    readOnly
                    value={
                      typeof window !== 'undefined'
                        ? `${window.location.origin}/?room=${roomCode}&role=${defaultRoomRole}${roomPin ? `&pin=${encodeURIComponent(roomPin)}` : ''}`
                        : `http://localhost:3005/?room=${roomCode}&role=${defaultRoomRole}${roomPin ? `&pin=${encodeURIComponent(roomPin)}` : ''}`
                    }
                    className="flex-1 bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs font-mono text-slate-600"
                  />
                  <button
                    onClick={() => {
                      const url = `${window.location.origin}/?room=${roomCode}&role=${defaultRoomRole}${roomPin ? `&pin=${encodeURIComponent(roomPin)}` : ''}`;
                      navigator.clipboard.writeText(url);
                      setCopiedLink(true);
                      saveRoomToDatabase();
                      setTimeout(() => setCopiedLink(false), 2000);
                    }}
                    className="px-3.5 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold flex items-center gap-1 cursor-pointer"
                  >
                    {copiedLink ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                    <span>{copiedLink ? 'Kopyalandı' : 'Kopyala'}</span>
                  </button>
                </div>
              </div>

              <div className="pt-3 flex justify-end gap-2 border-t border-slate-100">
                <button
                  onClick={() => setActiveModal(null)}
                  className="px-4 py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold"
                >
                  Kapat
                </button>
                <button
                  onClick={() => {
                    saveRoomToDatabase();
                    setActiveModal(null);
                  }}
                  className="px-4 py-2 rounded-lg bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-semibold"
                >
                  Kaydet ve Paylaşıma Aç
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: KULLANICININ GEÇMİŞTE GİRDİĞİ ODALAR, EXCELLER VE PIN YÖNETİMİ */}
      {(activeModal === 'my_rooms' || activeModal === 'join') && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-4xl w-full max-h-[88vh] flex flex-col shadow-2xl overflow-hidden">
            {/* Başlık */}
            <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between shrink-0">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-emerald-600 flex items-center justify-center">
                  <FolderOpen className="w-5 h-5 text-white" />
                </div>
                <div>
                  <h3 className="text-base font-bold flex items-center gap-2">
                    Excellerim, Geçmişte Girdiğim Odalar ve Kayıtlı PIN&apos;lerim
                    <span className="px-2 py-0.5 rounded-full bg-emerald-500/20 border border-emerald-400/30 text-emerald-300 text-[10px] font-mono">
                      @{currentUser.username}
                    </span>
                  </h3>
                  <p className="text-xs text-slate-300">
                    Hesabınızla geçmişte girdiğiniz veya oluşturduğunuz tüm Excel tablolarını, oda kodlarını ve PIN şifrelerini görüntüleyin; istediğinize tek tıkla geçiş yapıp çalışın.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* İçerik */}
            <div className="p-6 overflow-y-auto grid grid-cols-1 lg:grid-cols-12 gap-6 flex-1">
              {/* SOL TARAF (7 Sütun): KULLANICININ KAYITLI EXCEL ODALARI VE PIN LİSTESİ */}
              <div className="lg:col-span-7 flex flex-col space-y-3">
                <div className="flex items-center justify-between">
                  <h4 className="text-xs font-extrabold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                    Hesabımdaki Kayıtlı Excel Odaları ({userSavedRooms.length})
                  </h4>
                  <span className="text-[11px] text-slate-400">
                    Geçiş yapmak istediğiniz Excel&apos;e tıklayın
                  </span>
                </div>

                <div className="space-y-2.5 overflow-y-auto max-h-[55vh] pr-1">
                  {userSavedRooms.map((rm) => {
                    const isCurrent = rm.roomCode === roomCode;
                    const isPinShown = revealedPins[rm.roomCode] ?? true;
                    return (
                      <div
                        key={rm.roomCode}
                        className={`p-3.5 rounded-2xl border transition flex items-center justify-between gap-3 ${
                          isCurrent
                            ? 'bg-emerald-50/70 border-emerald-500 ring-1 ring-emerald-500'
                            : 'bg-white hover:bg-slate-50 border-slate-200'
                        }`}
                      >
                        <div className="min-w-0 space-y-1.5">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="text-sm font-bold text-slate-900 truncate">
                              {rm.title}.xlsx
                            </span>
                            {isCurrent && (
                              <span className="px-2 py-0.5 rounded-full bg-emerald-600 text-white text-[10px] font-extrabold">
                                ŞU AN AÇIK
                              </span>
                            )}
                          </div>

                          <div className="flex items-center gap-2 flex-wrap text-xs">
                            {/* Oda Kodu */}
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-slate-100 border border-slate-200 font-mono font-bold text-slate-800">
                              Oda Kodu: <span className="text-emerald-700">{rm.roomCode}</span>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  navigator.clipboard.writeText(rm.roomCode);
                                  showToast(`📋 Oda kodu (${rm.roomCode}) kopyalandı!`);
                                }}
                                className="text-slate-400 hover:text-slate-700 ml-1 cursor-pointer"
                                title="Oda kodunu kopyala"
                              >
                                <Copy className="w-3 h-3" />
                              </button>
                            </span>

                            {/* Kayıtlı PIN Şifresi */}
                            {rm.pin ? (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md bg-amber-50 border border-amber-300 font-mono font-bold text-amber-900">
                                <Lock className="w-3 h-3 text-amber-600" />
                                <span>
                                  PIN: {isPinShown ? rm.pin : '••••'}
                                </span>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setRevealedPins((prev) => ({
                                      ...prev,
                                      [rm.roomCode]: !isPinShown,
                                    }));
                                  }}
                                  className="text-[10px] underline text-amber-700 hover:text-amber-950 ml-0.5 cursor-pointer font-sans"
                                >
                                  {isPinShown ? 'Gizle' : 'Göster'}
                                </button>
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-slate-100 text-slate-500 text-[11px]">
                                <Unlock className="w-3 h-3" />
                                Şifresiz Oda
                              </span>
                            )}
                          </div>

                          <div className="text-[11px] text-slate-400">
                            Son erişim:{' '}
                            {rm.lastAccessedAt
                              ? new Date(rm.lastAccessedAt).toLocaleString('tr-TR')
                              : 'Yeni'}
                          </div>
                        </div>

                        <div className="flex items-center gap-1.5 shrink-0">
                          {!isCurrent ? (
                            <button
                              onClick={() => handleSwitchToUserRoom(rm.roomCode, rm.pin)}
                              className="px-3.5 py-2 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold cursor-pointer transition shadow-2xs"
                            >
                              Bu Excel&apos;e Geç ve Çalış →
                            </button>
                          ) : (
                            <span className="px-3 py-2 rounded-xl bg-emerald-100 text-emerald-800 text-xs font-bold">
                              Aktif Çalışılıyor
                            </span>
                          )}

                          {userSavedRooms.length > 1 && (
                            <button
                              onClick={() => handleRemoveRoomFromUserHistory(rm.roomCode)}
                              className="p-2 rounded-xl text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition cursor-pointer"
                              title="Bu odayı geçmiş listemden kaldır"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* SAĞ TARAF (5 Sütun): YENİ EXCEL ODASI OLUŞTUR VEYA BAŞKA ODAYA KOD+PIN İLE KATIL */}
              <div className="lg:col-span-5 space-y-4">
                {/* KART 1: YENİ EXCEL ODASI + PIN OLUŞTUR */}
                <div className="p-4 rounded-2xl bg-emerald-50/60 border border-emerald-200 space-y-3">
                  <div className="flex items-center gap-2">
                    <Plus className="w-4 h-4 text-emerald-700" />
                    <h4 className="text-xs font-extrabold uppercase tracking-wider text-emerald-950">
                      Yeni Excel Çalışma Odası Oluştur
                    </h4>
                  </div>
                  <p className="text-[11px] text-emerald-800">
                    Yeni bir Excel çalışma kitabı oluşturun, isterseniz PIN şifresi belirleyin. Hesabınıza otomatik kaydedilir.
                  </p>

                  <div className="space-y-2.5">
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">
                        Excel Dosya / Oda Adı
                      </label>
                      <input
                        type="text"
                        value={newRoomForm.title}
                        onChange={(e) =>
                          setNewRoomForm({ ...newRoomForm, title: e.target.value })
                        }
                        placeholder="Örn: 2026_Satis_Hedefleri"
                        className="w-full bg-white border border-emerald-300 rounded-xl px-3 py-2 text-xs text-slate-900"
                      />
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">
                          Oda PIN Şifresi (İsteğe Bağlı)
                        </label>
                        <input
                          type="text"
                          value={newRoomForm.pin}
                          onChange={(e) =>
                            setNewRoomForm({ ...newRoomForm, pin: e.target.value })
                          }
                          placeholder="Örn: 1453 veya Boş"
                          className="w-full bg-white border border-emerald-300 rounded-xl px-3 py-2 text-xs font-mono text-slate-900"
                        />
                      </div>

                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">
                          Başlangıç Şablonu
                        </label>
                        <select
                          value={newRoomForm.template}
                          onChange={(e) =>
                            setNewRoomForm({
                              ...newRoomForm,
                              template: e.target.value as any,
                            })
                          }
                          className="w-full bg-white border border-emerald-300 rounded-xl px-2.5 py-2 text-xs text-slate-900"
                        >
                          <option value="blank">Boş Excel Sayfası</option>
                          <option value="budget">Kurumsal Bütçe Şablonu</option>
                          <option value="crm">Satış & CRM Şablonu</option>
                        </select>
                      </div>
                    </div>

                    <button
                      onClick={handleCreateNewUserExcelRoom}
                      className="w-full py-2.5 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold shadow-xs cursor-pointer transition"
                    >
                      ➕ Yeni Excel Odası Oluştur ve Çalışmaya Başla
                    </button>
                  </div>
                </div>

                {/* KART 2: ODA KODU VE PIN İLE BAŞKA BİR EXCEL'E KATIL */}
                <div className="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center gap-2">
                    <Lock className="w-4 h-4 text-indigo-600" />
                    <h4 className="text-xs font-extrabold uppercase tracking-wider text-slate-800">
                      Oda Kodu & PIN ile Başka Bir Excel&apos;e Katıl
                    </h4>
                  </div>
                  <p className="text-[11px] text-slate-500">
                    Bir kez katıldığınızda oda kodu ve PIN şifresi hesabınıza kaydedilir; bir daha PIN yazmadan tek tıkla girebilirsiniz.
                  </p>

                  <div className="space-y-2">
                    <input
                      type="text"
                      value={joinCodeInput}
                      onChange={(e) => setJoinCodeInput(e.target.value.toUpperCase())}
                      placeholder="Oda Kodu (Örn: XLS-849-201)"
                      className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-mono text-slate-900"
                    />
                    <input
                      type="text"
                      value={joinPinInput}
                      onChange={(e) => setJoinPinInput(e.target.value)}
                      placeholder="Oda PIN Şifresi (Varsa)"
                      className="w-full bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs font-mono text-slate-900"
                    />
                    <button
                      onClick={() => {
                        if (!joinCodeInput.trim()) return;
                        handleSwitchToUserRoom(joinCodeInput, joinPinInput);
                      }}
                      className="w-full py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold cursor-pointer transition"
                    >
                      🔗 Odaya Katıl ve Hesabıma Kaydet
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: PANODAN YAPIŞTIR */}
      {activeModal === 'clipboard' && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-lg w-full p-6 shadow-2xl">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-base font-bold text-slate-900">
                Masaüstü Excel&apos;den Kopyala & Yapıştır
              </h3>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <textarea
              value={clipboardText}
              onChange={(e) => setClipboardText(e.target.value)}
              placeholder="Masaüstü Excel'de hücreleri kopyalayıp (Ctrl+C) buraya yapıştırın (Ctrl+V)..."
              rows={7}
              className="w-full bg-slate-50 border border-slate-300 rounded-xl p-3 text-xs font-mono text-slate-800 mb-4"
            />
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setActiveModal(null)}
                className="px-4 py-2 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700"
              >
                İptal
              </button>
              <button
                onClick={handleClipboardImport}
                className="px-4 py-2 rounded-lg text-xs font-semibold bg-[#107c41] text-white"
              >
                Yeni Çalışma Kitabı Sekmesine Aktar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: MAKRO */}
      {activeModal === 'macro' && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-lg w-full p-6 shadow-2xl">
            <div className="flex items-center justify-between mb-3">
              <h3 className="text-base font-bold text-slate-900">Makro / Otomasyon Çalıştır</h3>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-slate-700 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <textarea
              value={macroCode}
              onChange={(e) => setMacroCode(e.target.value)}
              rows={8}
              className="w-full bg-slate-900 text-emerald-300 rounded-xl p-3 text-xs font-mono mb-4"
            />
            <div className="flex justify-end gap-2">
              <button
                onClick={() => setActiveModal(null)}
                className="px-4 py-2 rounded-lg text-xs font-semibold bg-slate-100 text-slate-700"
              >
                İptal
              </button>
              <button
                onClick={() => {
                  sheetHandleRef.current?.runScriptOnActiveSheet((sheet) => {
                    const fn = new Function('sheet', macroCode);
                    fn(sheet);
                  });
                  setActiveModal(null);
                  showToast('⚡ Makro çalıştırıldı!');
                }}
                className="px-4 py-2 rounded-lg text-xs font-semibold bg-[#107c41] text-white"
              >
                Çalıştır
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: ÇİFT YÖNLÜ CANLI MASAÜSTÜ EXCEL <-> SYNCSHEET VLOOKUP / XLOOKUP KÖPRÜSÜ */}
      {activeModal === 'excel_bridge' && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-4xl w-full max-h-[90vh] flex flex-col shadow-2xl overflow-hidden">
            {/* Üst Başlık */}
            <div className="px-6 py-4 bg-slate-900 text-white flex items-center justify-between shrink-0">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-emerald-600 flex items-center justify-center shadow-sm">
                  <Link2 className="w-5 h-5 text-white" />
                </div>
                <div>
                  <h3 className="text-base font-bold flex items-center gap-2">
                    Canlı Masaüstü Excel ⇄ SyncSheet VLOOKUP / XLOOKUP Köprüsü
                    <span className="px-2 py-0.5 rounded-full bg-emerald-500/20 border border-emerald-400/30 text-emerald-300 text-[10px] font-semibold">
                      Çift Yönlü Canlı Senkronizasyon
                    </span>
                  </h3>
                  <p className="text-xs text-slate-300">
                    Bilgisayarınızda açık olan gerçek Microsoft Excel ile bu sistem arasında canlı DÜŞEYARA (VLOOKUP) ve XLOOKUP bağlantısı kurun
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveModal(null)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Yön Seçim Sekmeleri */}
            <div className="bg-slate-100 px-6 pt-3 border-b border-slate-200 flex gap-2 shrink-0">
              <button
                onClick={() => setBridgeTab('desktop_to_web')}
                className={`px-4 py-2.5 rounded-t-xl text-xs font-bold flex items-center gap-2 border-b-2 transition cursor-pointer ${
                  bridgeTab === 'desktop_to_web'
                    ? 'bg-white text-emerald-800 border-emerald-600 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900 border-transparent'
                }`}
              >
                <Upload className="w-4 h-4 text-emerald-600" />
                <span>
                  1. YÖN: Bilgisayardaki Açık Excel&apos;den → Bu Sisteme Canlı VLOOKUP / XLOOKUP Çek
                </span>
              </button>

              <button
                onClick={() => setBridgeTab('web_to_desktop')}
                className={`px-4 py-2.5 rounded-t-xl text-xs font-bold flex items-center gap-2 border-b-2 transition cursor-pointer ${
                  bridgeTab === 'web_to_desktop'
                    ? 'bg-white text-indigo-800 border-indigo-600 shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900 border-transparent'
                }`}
              >
                <ExternalLink className="w-4 h-4 text-indigo-600" />
                <span>
                  2. YÖN: Bilgisayardaki Gerçek Excel İçinden → Bu Sisteme VLOOKUP Yap (.iqy / Web Sorgusu)
                </span>
              </button>
            </div>

            {/* İçerik Gövdesi */}
            <div className="p-6 overflow-y-auto space-y-5 flex-1">
              {bridgeTab === 'desktop_to_web' ? (
                <>
                  {/* ADIM 1: MASAÜSTÜ EXCEL DOSYASINI BAĞLAMA VE CANLI İZLEME */}
                  <div className="p-4 rounded-2xl border border-emerald-200 bg-emerald-50/50 space-y-3">
                    <div className="flex items-start justify-between gap-4 flex-wrap">
                      <div>
                        <span className="inline-block px-2 py-0.5 rounded bg-emerald-700 text-white text-[10px] font-bold uppercase tracking-wider mb-1">
                          Adım 1 • Yerel Excel Dosyasını Canlı Bağla
                        </span>
                        <h4 className="text-sm font-bold text-slate-900">
                          Bilgisayarınızda Açık Olan Gerçek Excel (.xlsx) Dosyasını Seçin
                        </h4>
                        <p className="text-xs text-slate-600 mt-0.5">
                          Dosyayı seçtiğinizde sunucuya yüklenmeden (Zero-Upload) mevcut çalışma kitabınıza canlı referans sekmesi olarak bağlanır.{' '}
                          <strong>
                            Gerçek Excel&apos;de değişiklik yapıp Kaydet (Ctrl+S) dediğinizde buradaki VLOOKUP sonuçları 2 saniye içinde otomatik güncellenir!
                          </strong>
                        </p>
                      </div>

                      <div className="flex items-center gap-2 shrink-0 flex-wrap">
                        <button
                          onClick={() => downloadSampleDesktopExcelForVlookup(workbookData)}
                          className="px-3 py-2 rounded-xl bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 text-xs font-semibold flex items-center gap-1.5 cursor-pointer transition shadow-2xs"
                          title="Mevcut tablonuzdaki Proje Kodlarıyla birebir eşleşen örnek bir gerçek Excel dosyasını bilgisayarınıza indirir"
                        >
                          <Download className="w-3.5 h-3.5 text-indigo-600" />
                          <span>Test İçin Uyumlu Örnek Excel İndir (.xlsx)</span>
                        </button>

                        <button
                          onClick={handlePickLiveDesktopExcelFile}
                          className="px-4 py-2 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold flex items-center gap-1.5 cursor-pointer transition shadow-sm"
                        >
                          <FolderOpen className="w-4 h-4" />
                          <span>
                            {linkedDesktopExcel
                              ? 'Farklı Bir Masaüstü Excel Seç (.xlsx)'
                              : 'Bilgisayardaki Açık Excel Dosyasını Bağla (.xlsx)'}
                          </span>
                        </button>
                      </div>
                    </div>

                    {/* Bağlı Dosya Durum Kartı */}
                    {linkedDesktopExcel ? (
                      <div className="bg-white border border-emerald-300 rounded-xl p-3.5 flex items-center justify-between gap-3 flex-wrap">
                        <div className="flex items-center gap-3">
                          <div className="w-9 h-9 rounded-lg bg-emerald-100 text-emerald-800 flex items-center justify-center font-bold">
                            <FileSpreadsheet className="w-5 h-5" />
                          </div>
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="text-xs font-bold text-slate-900">
                                {linkedDesktopExcel.fileName}
                              </span>
                              <span className="px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-mono font-bold">
                                Sekme Adı: {linkedDesktopExcel.linkedSheetName}
                              </span>
                              {linkedDesktopExcel.isLiveWatching ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-600 text-white text-[10px] font-bold">
                                  <Radio className="w-3 h-3 animate-pulse" />
                                  Canlı İzleniyor (Ctrl+S Otomatik Yansır)
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 text-[10px] font-semibold">
                                  Manuel Yenileme Modu
                                </span>
                              )}
                            </div>
                            <p className="text-[11px] text-slate-500 mt-0.5">
                              {linkedDesktopExcel.rowCount} satır • {linkedDesktopExcel.colCount} sütun algılandı • Son Senkronizasyon:{' '}
                              <strong>{linkedDesktopExcel.lastSyncedAt}</strong>
                            </p>
                          </div>
                        </div>

                        <button
                          onClick={handleRefreshDesktopFileNow}
                          className="px-3 py-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-100 text-emerald-800 border border-emerald-200 text-xs font-semibold flex items-center gap-1.5 cursor-pointer"
                        >
                          <RefreshCw className="w-3.5 h-3.5" />
                          <span>Şimdi Yenile</span>
                        </button>
                      </div>
                    ) : (
                      <div className="bg-white/80 border border-dashed border-emerald-300 rounded-xl p-4 text-center text-xs text-slate-600">
                        💡 Henüz bilgisayarınızdan bir Excel dosyası bağlamadınız. Hemen denemek için önce{' '}
                        <strong>&quot;Test İçin Uyumlu Örnek Excel İndir (.xlsx)&quot;</strong> butonuna basıp dosyayı bilgisayarınızda gerçek Excel ile açın, ardından{' '}
                        <strong>&quot;Bilgisayardaki Açık Excel Dosyasını Bağla&quot;</strong> ile seçin!
                      </div>
                    )}
                  </div>

                  {/* ADIM 2: SÜTUN EŞLEŞTİRME VE CANLI VLOOKUP / XLOOKUP UYGULAMA */}
                  <div
                    className={`p-4 rounded-2xl border space-y-4 transition ${
                      linkedDesktopExcel
                        ? 'border-slate-200 bg-white'
                        : 'border-slate-200 bg-slate-50 opacity-60 pointer-events-none'
                    }`}
                  >
                    <div className="flex items-center justify-between flex-wrap gap-2">
                      <div>
                        <span className="inline-block px-2 py-0.5 rounded bg-indigo-600 text-white text-[10px] font-bold uppercase tracking-wider mb-1">
                          Adım 2 • Kolay VLOOKUP / XLOOKUP Sütun Eşleştirici
                        </span>
                        <h4 className="text-sm font-bold text-slate-900">
                          Hangi Sütunları Eşleştirip Veriyi Çekmek İstiyorsunuz?
                        </h4>
                      </div>

                      {/* Formül Tipi Seçimi */}
                      <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200">
                        <button
                          onClick={() => setBridgeFormulaType('XLOOKUP')}
                          className={`px-3 py-1 rounded-lg text-xs font-bold cursor-pointer transition ${
                            bridgeFormulaType === 'XLOOKUP'
                              ? 'bg-indigo-600 text-white shadow-2xs'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          =XLOOKUP (Önerilen)
                        </button>
                        <button
                          onClick={() => setBridgeFormulaType('VLOOKUP')}
                          className={`px-3 py-1 rounded-lg text-xs font-bold cursor-pointer transition ${
                            bridgeFormulaType === 'VLOOKUP'
                              ? 'bg-indigo-600 text-white shadow-2xs'
                              : 'text-slate-600 hover:text-slate-900'
                          }`}
                        >
                          =VLOOKUP (DÜŞEYARA)
                        </button>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                      {/* 1. Bu Sistemdeki Ortak Anahtar Sütunu */}
                      <div className="p-3 rounded-xl bg-slate-50 border border-slate-200">
                        <label className="block text-xs font-bold text-slate-800 mb-1">
                          1. Bu Sistemdeki (SyncSheet) Aranan Sütun
                        </label>
                        <p className="text-[11px] text-slate-500 mb-2">
                          Örn: Proje Kodu, Sicil No, Ürün Kodu
                        </p>
                        <select
                          value={bridgeSourceKeyCol}
                          onChange={(e) => setBridgeSourceKeyCol(Number(e.target.value))}
                          className="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-2 text-xs font-semibold text-slate-900"
                        >
                          {currentSheetColumns.map((col) => (
                            <option key={col.colIndex} value={col.colIndex}>
                              {col.colLetter} Sütunu: {col.headerName} (Örn: {String(col.sampleValue)})
                            </option>
                          ))}
                        </select>
                      </div>

                      {/* 2. Masaüstü Excel'deki Eşleşecek Sütun */}
                      <div className="p-3 rounded-xl bg-emerald-50/60 border border-emerald-200">
                        <label className="block text-xs font-bold text-emerald-950 mb-1">
                          2. Masaüstü Excel&apos;deki Eşleşecek Sütun
                        </label>
                        <p className="text-[11px] text-emerald-700 mb-2">
                          Masaüstü Excel dosyasında aramanın yapılacağı sütun
                        </p>
                        <select
                          value={bridgeExtKeyCol}
                          onChange={(e) => setBridgeExtKeyCol(Number(e.target.value))}
                          className="w-full bg-white border border-emerald-300 rounded-lg px-2.5 py-2 text-xs font-semibold text-slate-900"
                        >
                          {(linkedDesktopExcel?.headers || []).map((h) => (
                            <option key={h.colIndex} value={h.colIndex}>
                              {h.colLetter} Sütunu: {h.headerName} (Örn: {String(h.sampleValue)})
                            </option>
                          ))}
                        </select>
                      </div>

                      {/* 3. Masaüstü Excel'den Getirilecek Sütun */}
                      <div className="p-3 rounded-xl bg-indigo-50/60 border border-indigo-200">
                        <label className="block text-xs font-bold text-indigo-950 mb-1">
                          3. Masaüstü Excel&apos;den Getirilecek Veri Sütunu
                        </label>
                        <p className="text-[11px] text-indigo-700 mb-2">
                          Eşleşme bulunduğunda bu sisteme çekilecek sütun
                        </p>
                        <select
                          value={bridgeExtReturnCol}
                          onChange={(e) => setBridgeExtReturnCol(Number(e.target.value))}
                          className="w-full bg-white border border-indigo-300 rounded-lg px-2.5 py-2 text-xs font-semibold text-slate-900"
                        >
                          {(linkedDesktopExcel?.headers || []).map((h) => (
                            <option key={h.colIndex} value={h.colIndex}>
                              {h.colLetter} Sütunu: {h.headerName} (Örn: {String(h.sampleValue)})
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>

                    {/* Yazdırma Modu ve Canlı Formül Önizlemesi */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 items-center bg-slate-900 text-white p-4 rounded-xl">
                      <div className="space-y-2">
                        <label className="block text-[11px] font-bold uppercase tracking-wider text-emerald-400">
                          Uygulama Yöntemi Seçin:
                        </label>
                        <div className="flex flex-col gap-1.5 text-xs">
                          <label className="flex items-center gap-2 cursor-pointer">
                            <input
                              type="radio"
                              checked={bridgeTargetMode === 'NEW_COL'}
                              onChange={() => setBridgeTargetMode('NEW_COL')}
                              className="accent-emerald-500"
                            />
                            <span>
                              <strong>Tüm Satırlara Otomatik Uygula:</strong> Tablonun sağına yeni sütun açıp tüm satırlara formülü doldur
                            </span>
                          </label>
                          <label className="flex items-center gap-2 cursor-pointer">
                            <input
                              type="radio"
                              checked={bridgeTargetMode === 'ACTIVE_CELL'}
                              onChange={() => setBridgeTargetMode('ACTIVE_CELL')}
                              className="accent-emerald-500"
                            />
                            <span>
                              <strong>Sadece Seçili Hücreye ({activeSelection.cellRef}) Yaz:</strong> Yalnızca aktif hücreye formülü yerleştir
                            </span>
                          </label>
                        </div>
                      </div>

                      <div className="bg-slate-800 border border-slate-700 rounded-xl p-3">
                        <div className="text-[10px] font-semibold uppercase tracking-wider text-slate-400 mb-1">
                          Oluşturulacak Canlı Excel Formülü:
                        </div>
                        <div className="font-mono text-xs text-emerald-300 font-bold break-all">
                          {linkedDesktopExcel
                            ? bridgeFormulaType === 'VLOOKUP'
                              ? `=VLOOKUP(${colIndexToLetter(bridgeSourceKeyCol)}2, ${linkedDesktopExcel.linkedSheetName}!${colIndexToLetter(Math.min(bridgeExtKeyCol, bridgeExtReturnCol))}2:${colIndexToLetter(Math.max(bridgeExtKeyCol, bridgeExtReturnCol))}${linkedDesktopExcel.rowCount}, ${Math.abs(bridgeExtReturnCol - bridgeExtKeyCol) + 1}, 0)`
                              : `=XLOOKUP(${colIndexToLetter(bridgeSourceKeyCol)}2, ${linkedDesktopExcel.linkedSheetName}!${colIndexToLetter(bridgeExtKeyCol)}2:${colIndexToLetter(bridgeExtKeyCol)}${linkedDesktopExcel.rowCount}, ${linkedDesktopExcel.linkedSheetName}!${colIndexToLetter(bridgeExtReturnCol)}2:${colIndexToLetter(bridgeExtReturnCol)}${linkedDesktopExcel.rowCount}, "Eşleşmedi")`
                            : '=XLOOKUP(...)'}
                        </div>
                      </div>
                    </div>

                    <div className="flex justify-end gap-2 pt-1">
                      <button
                        onClick={() => setActiveModal(null)}
                        className="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold cursor-pointer"
                      >
                        Kapat
                      </button>
                      <button
                        onClick={handleApplyBridgeVlookup}
                        className="px-5 py-2.5 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold shadow-sm cursor-pointer flex items-center gap-2"
                      >
                        <Link2 className="w-4 h-4" />
                        <span>
                          {bridgeTargetMode === 'NEW_COL'
                            ? `Tüm Tabloya Canlı ${bridgeFormulaType} Uygula`
                            : `${activeSelection.cellRef} Hücresine Canlı ${bridgeFormulaType} Yaz`}
                        </span>
                      </button>
                    </div>
                  </div>
                </>
              ) : (
                /* 2. YÖN: GERÇEK MASAÜSTÜ EXCEL'DEN -> BU SİSTEME (SYNCSHEET) CANLI VLOOKUP YAPMAK */
                <div className="space-y-4">
                  <div className="p-4 rounded-2xl bg-indigo-50 border border-indigo-200 space-y-2">
                    <span className="inline-block px-2 py-0.5 rounded bg-indigo-700 text-white text-[10px] font-bold uppercase tracking-wider">
                      Gerçek Microsoft Excel → SyncSheet Canlı Veri Akışı
                    </span>
                    <h4 className="text-sm font-bold text-slate-900">
                      Bilgisayarınızdaki Gerçek Excel İçinden Bu Sistemdeki Tabloya VLOOKUP / DÜŞEYARA Yapın
                    </h4>
                    <p className="text-xs text-slate-600 leading-relaxed">
                      Bilgisayarınızda açık olan gerçek Microsoft Excel uygulamasında çalışırken, SyncSheet üzerindeki{' '}
                      <code className="font-mono font-bold text-indigo-800">{roomCode}</code> odasının güncel verilerini gerçek zamanlı bir tablo olarak Excel&apos;inize bağlayabilir ve kendi Excel dosyanızda{' '}
                      <code className="font-mono font-bold">=VLOOKUP(...)</code> veya{' '}
                      <code className="font-mono font-bold">=DÜŞEYARA(...)</code> yazabilirsiniz.
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* YÖNTEM A: TEK TIKLA .IQY (EXCEL WEB SORGUSU) İNDİR */}
                    <div className="p-4 rounded-2xl border border-slate-200 bg-white flex flex-col justify-between space-y-3">
                      <div className="space-y-2">
                        <div className="flex items-center gap-2">
                          <span className="w-6 h-6 rounded-full bg-emerald-100 text-emerald-800 font-bold text-xs flex items-center justify-center">
                            A
                          </span>
                          <h5 className="text-xs font-bold text-slate-900">
                            En Kolay Yöntem: Tek Tıkla Canlı Excel Bağlantı Dosyası (.iqy)
                          </h5>
                        </div>
                        <p className="text-xs text-slate-600 leading-relaxed">
                          Aşağıdaki butona tıkladığınızda bilgisayarınıza Microsoft Excel&apos;in resmi{' '}
                          <strong>Canlı Web Sorgusu (`.iqy`)</strong> dosyası iner. İnen dosyaya çift tıkladığınızda gerçek Excel otomatik açılır ve buradaki tabloyu canlı olarak Excel&apos;e çeker!
                        </p>
                        <ul className="text-[11px] text-slate-500 space-y-1 list-disc pl-4">
                          <li>Gerçek Excel&apos;de tabloya sağ tıklayıp <strong>&quot;Yenile (Refresh)&quot;</strong> dediğinizde buradaki son değişiklikler anında Excel&apos;inize gelir.</li>
                          <li>İlk bağlantidan önce güncel verinin yazılması için üstteki <strong>Kaydet</strong> butonuna basmanız önerilir.</li>
                        </ul>
                      </div>

                      <button
                        onClick={async () => {
                          await saveRoomToDatabase();
                          window.open(
                            `/api/excel-bridge?room=${encodeURIComponent(roomCode)}&format=iqy`,
                            '_blank'
                          );
                          showToast('📥 Gerçek Excel canlı bağlantı dosyası (.iqy) indirildi!');
                        }}
                        className="w-full py-2.5 px-4 rounded-xl bg-[#107c41] hover:bg-[#0c5e31] text-white text-xs font-bold flex items-center justify-center gap-2 cursor-pointer shadow-xs transition"
                      >
                        <Download className="w-4 h-4" />
                        <span>Gerçek Excel Canlı Bağlantı Dosyası İndir (.iqy)</span>
                      </button>
                    </div>

                    {/* YÖNTEM B: POWER QUERY / VERİ -> WEB'DEN CANLI URL */}
                    <div className="p-4 rounded-2xl border border-slate-200 bg-white flex flex-col justify-between space-y-3">
                      <div className="space-y-2">
                        <div className="flex items-center gap-2">
                          <span className="w-6 h-6 rounded-full bg-indigo-100 text-indigo-800 font-bold text-xs flex items-center justify-center">
                            B
                          </span>
                          <h5 className="text-xs font-bold text-slate-900">
                            Power Query Yöntemi: Excel → Veri → Web&apos;den (From Web)
                          </h5>
                        </div>
                        <p className="text-xs text-slate-600 leading-relaxed">
                          Gerçek Microsoft Excel&apos;de üst menüden{' '}
                          <strong>Veri (Data) → Web&apos;den (From Web)</strong> seçeneğine tıklayıp aşağıdaki canlı CSV veya HTML tablo adresini yapıştırarak canlı bağlantı kurabilirsiniz:
                        </p>

                        <div className="space-y-2 pt-1">
                          <div>
                            <label className="block text-[10px] font-bold uppercase text-slate-400 mb-1">
                              Canlı CSV Veri Akışı URL (Power Query İçin):
                            </label>
                            <div className="flex items-center gap-1.5">
                              <input
                                type="text"
                                readOnly
                                value={
                                  typeof window !== 'undefined'
                                    ? `${window.location.origin}/api/excel-bridge?room=${roomCode}&format=csv`
                                    : `http://localhost:3005/api/excel-bridge?room=${roomCode}&format=csv`
                                }
                                className="flex-1 bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-[11px] font-mono text-slate-700"
                              />
                              <button
                                onClick={() => {
                                  const url = `${window.location.origin}/api/excel-bridge?room=${roomCode}&format=csv`;
                                  navigator.clipboard.writeText(url);
                                  setCopiedBridgeUrl('csv');
                                  setTimeout(() => setCopiedBridgeUrl(null), 2000);
                                  showToast('📋 Canlı CSV bağlantı adresi kopyalandı!');
                                }}
                                className="px-3 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-semibold cursor-pointer"
                              >
                                {copiedBridgeUrl === 'csv' ? 'Kopyalandı' : 'Kopyala'}
                              </button>
                            </div>
                          </div>

                          <div>
                            <label className="block text-[10px] font-bold uppercase text-slate-400 mb-1">
                              Canlı HTML Tablo Akışı URL:
                            </label>
                            <div className="flex items-center gap-1.5">
                              <input
                                type="text"
                                readOnly
                                value={
                                  typeof window !== 'undefined'
                                    ? `${window.location.origin}/api/excel-bridge?room=${roomCode}&format=html`
                                    : `http://localhost:3005/api/excel-bridge?room=${roomCode}&format=html`
                                }
                                className="flex-1 bg-slate-50 border border-slate-300 rounded-lg px-2.5 py-1.5 text-[11px] font-mono text-slate-700"
                              />
                              <button
                                onClick={() => {
                                  const url = `${window.location.origin}/api/excel-bridge?room=${roomCode}&format=html`;
                                  navigator.clipboard.writeText(url);
                                  setCopiedBridgeUrl('html');
                                  setTimeout(() => setCopiedBridgeUrl(null), 2000);
                                  showToast('📋 Canlı HTML tablo adresi kopyalandı!');
                                }}
                                className="px-3 py-1.5 rounded-lg bg-slate-900 text-white text-xs font-semibold cursor-pointer"
                              >
                                {copiedBridgeUrl === 'html' ? 'Kopyalandı' : 'Kopyala'}
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Gerçek Excel'de Kullanılabilecek Hazır Formül Örnekleri */}
                  <div className="bg-slate-900 text-white rounded-2xl p-4 space-y-2">
                    <div className="text-[11px] font-bold text-emerald-400 uppercase tracking-wider">
                      💡 Gerçek Masaüstü Excel&apos;de Bağlantı Sonrası Yazabileceğiniz Örnek Formüller:
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2 text-xs font-mono">
                      <div className="bg-slate-800 p-2.5 rounded-xl border border-slate-700">
                        <span className="text-slate-400 block text-[10px] font-sans mb-0.5">
                          Türkçe Microsoft Excel (DÜŞEYARA / ÇAPRAZARA):
                        </span>
                        <code className="text-emerald-300">
                          =DÜŞEYARA(A2; SyncSheet_Tablo!A:H; 7; YANLIŞ)
                        </code>
                      </div>
                      <div className="bg-slate-800 p-2.5 rounded-xl border border-slate-700">
                        <span className="text-slate-400 block text-[10px] font-sans mb-0.5">
                          İngilizce Microsoft Excel (VLOOKUP / XLOOKUP):
                        </span>
                        <code className="text-indigo-300">
                          =XLOOKUP(A2, SyncSheet_Tablo!A:A, SyncSheet_Tablo!G:G)
                        </code>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
