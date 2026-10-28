import fs from 'fs';
import path from 'path';

const DATA_DIR = process.env.DATA_DIR || path.join(process.cwd(), '.data');
const USERS_FILE = path.join(DATA_DIR, 'users.json');
const ROOMS_FILE = path.join(DATA_DIR, 'rooms.json');

function getSupabaseConfig(): { url: string; key: string } | null {
  const url = (
    process.env.SUPABASE_URL ||
    process.env.NEXT_PUBLIC_SUPABASE_URL ||
    ''
  )
    .trim()
    .replace(/\/+$/, '');

  const key = (
    process.env.SUPABASE_SERVICE_ROLE_KEY ||
    process.env.SUPABASE_KEY ||
    process.env.SUPABASE_ANON_KEY ||
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||
    ''
  ).trim();

  if (!url || !key) {
    return null;
  }
  return { url, key };
}

export function isSupabaseConfigured(): boolean {
  return getSupabaseConfig() !== null;
}

function ensureLocalDataDir() {
  try {
    if (!fs.existsSync(DATA_DIR)) {
      fs.mkdirSync(DATA_DIR, { recursive: true });
    }
  } catch {
    // ignore on read-only fs
  }
}

export function readLocalUsersMap(): Record<string, any> {
  try {
    ensureLocalDataDir();
    if (!fs.existsSync(USERS_FILE)) return {};
    return JSON.parse(fs.readFileSync(USERS_FILE, 'utf-8') || '{}');
  } catch {
    return {};
  }
}

export function writeLocalUsersMap(users: Record<string, any>) {
  try {
    ensureLocalDataDir();
    fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2), 'utf-8');
  } catch {
    // ignore
  }
}

export function readLocalRoomsMap(): Record<string, any> {
  try {
    ensureLocalDataDir();
    if (!fs.existsSync(ROOMS_FILE)) return {};
    return JSON.parse(fs.readFileSync(ROOMS_FILE, 'utf-8') || '{}');
  } catch {
    return {};
  }
}

export function writeLocalRoomsMap(rooms: Record<string, any>) {
  try {
    ensureLocalDataDir();
    fs.writeFileSync(ROOMS_FILE, JSON.stringify(rooms), 'utf-8');
  } catch {
    // ignore
  }
}

function buildHeaders(key: string, extra?: Record<string, string>): Record<string, string> {
  return {
    apikey: key,
    Authorization: `Bearer ${key}`,
    'Content-Type': 'application/json',
    ...(extra || {}),
  };
}

// ============================================================================
// KULLANICI VERİTABANI İŞLEMLERİ (syncsheet_users - E-posta Bazlı Tekil Kayıt)
// ============================================================================

export async function dbGetUser(emailOrKey: string): Promise<any | null> {
  const clean = emailOrKey.trim().toLowerCase();
  if (!clean) return null;

  const cfg = getSupabaseConfig();
  if (cfg) {
    try {
      const endpoint = `${cfg.url}/rest/v1/syncsheet_users?username=eq.${encodeURIComponent(clean)}&select=data&limit=1`;
      const res = await fetch(endpoint, {
        method: 'GET',
        headers: buildHeaders(cfg.key),
      });
      if (res.ok) {
        const rows = await res.json();
        if (Array.isArray(rows) && rows.length > 0 && rows[0]?.data) {
          return rows[0].data;
        }
      }
    } catch (err) {
      console.error('Supabase dbGetUser error:', err);
    }
  }

  const localMap = readLocalUsersMap();
  return localMap[clean] || null;
}

export async function dbUpsertUser(user: any): Promise<void> {
  const clean = String(user?.email || user?.username || '')
    .trim()
    .toLowerCase();
  if (!clean) return;

  const normalizedUser = {
    ...user,
    email: clean,
    username: clean,
  };

  // Yerel önbelleği güncelle
  const localMap = readLocalUsersMap();
  localMap[clean] = normalizedUser;
  writeLocalUsersMap(localMap);

  // Supabase'e kalıcı olarak yaz
  const cfg = getSupabaseConfig();
  if (cfg) {
    try {
      const endpoint = `${cfg.url}/rest/v1/syncsheet_users?on_conflict=username`;
      await fetch(endpoint, {
        method: 'POST',
        headers: buildHeaders(cfg.key, {
          Prefer: 'resolution=merge-duplicates,return=minimal',
        }),
        body: JSON.stringify([
          {
            username: clean,
            data: normalizedUser,
            updated_at: new Date().toISOString(),
          },
        ]),
      });
    } catch (err) {
      console.error('Supabase dbUpsertUser error:', err);
    }
  }
}

// ============================================================================
// EXCEL ODALARI VERİTABANI İŞLEMLERİ (syncsheet_rooms)
// ============================================================================

export async function dbGetRoom(roomCode: string): Promise<any | null> {
  const clean = roomCode.trim().toUpperCase();
  if (!clean) return null;

  const cfg = getSupabaseConfig();
  if (cfg) {
    try {
      const endpoint = `${cfg.url}/rest/v1/syncsheet_rooms?room_code=eq.${encodeURIComponent(clean)}&select=data&limit=1`;
      const res = await fetch(endpoint, {
        method: 'GET',
        headers: buildHeaders(cfg.key),
      });
      if (res.ok) {
        const rows = await res.json();
        if (Array.isArray(rows) && rows.length > 0 && rows[0]?.data) {
          return rows[0].data;
        }
        return null;
      }
    } catch (err) {
      console.error('Supabase dbGetRoom error:', err);
    }
  }

  const localMap = readLocalRoomsMap();
  return localMap[clean] || null;
}

export async function dbGetAllRoomsMap(): Promise<Record<string, any>> {
  const cfg = getSupabaseConfig();

  if (cfg) {
    try {
      const endpoint = `${cfg.url}/rest/v1/syncsheet_rooms?select=room_code,data`;
      const res = await fetch(endpoint, {
        method: 'GET',
        headers: buildHeaders(cfg.key),
      });
      if (res.ok) {
        const rows = await res.json();
        if (Array.isArray(rows)) {
          const remoteMap: Record<string, any> = {};
          for (const row of rows) {
            if (row?.room_code && row?.data) {
              remoteMap[row.room_code] = row.data;
            }
          }
          writeLocalRoomsMap(remoteMap);
          return remoteMap;
        }
      }
    } catch (err) {
      console.error('Supabase dbGetAllRoomsMap error:', err);
    }
  }

  return readLocalRoomsMap();
}

export async function dbUpsertRoom(room: any): Promise<void> {
  const clean = String(room?.roomCode || '')
    .trim()
    .toUpperCase();
  if (!clean) return;

  const cloudRoom = {
    ...room,
    recentCellDeltas: Array.isArray(room.recentCellDeltas)
      ? room.recentCellDeltas.slice(-50)
      : [],
    auditLog: Array.isArray(room.auditLog) ? room.auditLog.slice(0, 50) : [],
    snapshots: Array.isArray(room.snapshots)
      ? room.snapshots.slice(0, 10).map((s: any) => {
          const snapSize = s?.workbookData ? JSON.stringify(s.workbookData).length : 0;
          // Büyük Excel / Makro dosyalarında (>250KB) snapshot içinde workbookData'yı ikinci kez kopyalayıp
          // veritabanı paket boyutunu ikiye katlamayı önle (ana workbookData zaten room.workbookData içinde kayıtlıdır)
          if (snapSize > 250000) {
            const { workbookData, ...metaOnly } = s;
            return metaOnly;
          }
          return s;
        })
      : [],
  };

  // Yerel önbelleği güncelle
  const localMap = readLocalRoomsMap();
  localMap[clean] = cloudRoom;
  writeLocalRoomsMap(localMap);

  // Supabase'e kalıcı olarak yaz
  const cfg = getSupabaseConfig();
  if (cfg) {
    try {
      const endpoint = `${cfg.url}/rest/v1/syncsheet_rooms?on_conflict=room_code`;
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 15000);
      await fetch(endpoint, {
        method: 'POST',
        headers: buildHeaders(cfg.key, {
          Prefer: 'resolution=merge-duplicates,return=minimal',
        }),
        body: JSON.stringify([
          {
            room_code: clean,
            data: cloudRoom,
            updated_at: new Date().toISOString(),
          },
        ]),
        signal: controller.signal,
      }).finally(() => clearTimeout(timeoutId));
    } catch (err) {
      console.error('Supabase dbUpsertRoom error:', err);
    }
  }
}

export async function dbDeleteRoom(roomCode: string): Promise<void> {
  const clean = String(roomCode || '')
    .trim()
    .toUpperCase();
  if (!clean) return;

  // 1) Yerel dosya önbelleğinden tamamen sil
  const localMap = readLocalRoomsMap();
  if (localMap[clean]) {
    delete localMap[clean];
    writeLocalRoomsMap(localMap);
  }

  // 2) Supabase veritabanından (syncsheet_rooms) kalıcı olarak sil (Çöplük oluşmasını engeller)
  const cfg = getSupabaseConfig();
  if (cfg) {
    try {
      const endpoint = `${cfg.url}/rest/v1/syncsheet_rooms?room_code=eq.${encodeURIComponent(clean)}`;
      await fetch(endpoint, {
        method: 'DELETE',
        headers: buildHeaders(cfg.key),
      });
    } catch (err) {
      console.error('Supabase dbDeleteRoom error:', err);
    }
  }
}
