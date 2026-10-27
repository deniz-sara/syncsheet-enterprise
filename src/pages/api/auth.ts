import type { NextApiRequest, NextApiResponse } from 'next';
import {
  dbGetAllRoomsMap,
  dbGetUser,
  dbUpsertUser,
  isSupabaseConfigured,
} from '@/lib/supabaseStorage';

export interface UserSavedRoomEntry {
  roomCode: string;
  title: string;
  pin: string;
  role: 'owner' | 'editor' | 'viewer';
  lastAccessedAt: string;
  updatedAt?: string;
}

export interface StoredUserRecord {
  id: string;
  username: string;
  password: string;
  name: string;
  email: string;
  department: string;
  color: string;
  createdAt: string;
  savedRooms: UserSavedRoomEntry[];
}

const USER_COLORS = [
  '#059669',
  '#2563eb',
  '#7c3aed',
  '#db2777',
  '#d97706',
  '#0891b2',
  '#4f46e5',
  '#be123c',
];

function pickColor(str: string): string {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  return USER_COLORS[Math.abs(hash) % USER_COLORS.length];
}

async function enrichUserRooms(user: StoredUserRecord): Promise<UserSavedRoomEntry[]> {
  const roomsDb = await dbGetAllRoomsMap();
  const list = Array.isArray(user.savedRooms) ? user.savedRooms : [];

  return list
    .map((entry) => {
      const liveRoom = roomsDb[entry.roomCode];
      return {
        ...entry,
        title:
          liveRoom?.workbookData?.name ||
          liveRoom?.title ||
          entry.title ||
          `Excel_${entry.roomCode}`,
        pin: entry.pin || liveRoom?.pin || '',
        updatedAt: liveRoom?.updatedAt || entry.lastAccessedAt,
      };
    })
    .sort(
      (a, b) =>
        new Date(b.lastAccessedAt || 0).getTime() -
        new Date(a.lastAccessedAt || 0).getTime()
    );
}

async function sanitizeUserForClient(user: StoredUserRecord) {
  return {
    id: user.id,
    username: user.username,
    name: user.name,
    email: user.email,
    department: user.department,
    color: user.color,
    role: 'editor' as const,
    savedRooms: await enrichUserRooms(user),
    storageEngine: isSupabaseConfigured() ? 'supabase' : 'local',
  };
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method === 'GET') {
    const username = String(req.query.username || '')
      .trim()
      .toLowerCase();
    if (!username) {
      return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
    }
    const existing = await dbGetUser(username);
    if (!existing) {
      return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
    }
    return res.status(200).json({ user: await sanitizeUserForClient(existing) });
  }

  if (req.method === 'POST') {
    const { action } = req.body || {};

    // 1) YENİ KULLANICI KAYDI (REGISTER)
    if (action === 'REGISTER') {
      const rawUsername = String(req.body.username || '')
        .trim()
        .toLowerCase();
      const password = String(req.body.password || '').trim();
      const name = String(req.body.name || '').trim() || rawUsername;
      const department = String(req.body.department || '').trim() || 'Genel Müdürlük';

      if (!rawUsername || rawUsername.length < 3) {
        return res
          .status(400)
          .json({ error: 'Kullanıcı adı en az 3 karakter olmalıdır.' });
      }
      if (!password || password.length < 3) {
        return res.status(400).json({ error: 'Şifre en az 3 karakter olmalıdır.' });
      }

      const existing = await dbGetUser(rawUsername);
      if (existing) {
        return res.status(409).json({
          error: 'Bu kullanıcı adı zaten kayıtlı. Lütfen giriş yapın veya farklı bir kullanıcı adı seçin.',
        });
      }

      const nowIso = new Date().toISOString();
      const newUser: StoredUserRecord = {
        id: `usr_${rawUsername.replace(/[^a-z0-9]/g, '_')}`,
        username: rawUsername,
        password,
        name,
        email: `${rawUsername}@kurumsal.local`,
        department,
        color: pickColor(rawUsername),
        createdAt: nowIso,
        savedRooms: [
          {
            roomCode: 'XLS-849-201',
            title: '2026_Kurumsal_Butce_Plani',
            pin: '',
            role: 'owner',
            lastAccessedAt: nowIso,
          },
        ],
      };

      await dbUpsertUser(newUser);

      return res.status(200).json({
        ok: true,
        user: await sanitizeUserForClient(newUser),
      });
    }

    // 2) KULLANICI GİRİŞİ (LOGIN)
    if (action === 'LOGIN') {
      const rawUsername = String(req.body.username || '')
        .trim()
        .toLowerCase();
      const password = String(req.body.password || '').trim();

      const existing = await dbGetUser(rawUsername);
      if (!existing) {
        return res.status(404).json({
          error: 'Bu kullanıcı adına ait hesap bulunamadı. "Hesap Oluştur (Kayıt Ol)" sekmesinden hemen oluşturabilirsiniz.',
        });
      }
      if (existing.password !== password) {
        return res.status(401).json({ error: 'Girdiğiniz şifre hatalı.' });
      }

      return res.status(200).json({
        ok: true,
        user: await sanitizeUserForClient(existing),
      });
    }

    // 3) KULLANICININ GİRDİĞİ VEYA OLUŞTURDUĞU ODAYI + PIN'İ GEÇMİŞİNE KAYDET (UPSERT_USER_ROOM)
    if (action === 'UPSERT_USER_ROOM') {
      const rawUsername = String(req.body.username || '')
        .trim()
        .toLowerCase();
      const roomCode = String(req.body.roomCode || '')
        .trim()
        .toUpperCase();
      const title = String(req.body.title || '').trim();
      const pin = req.body.pin !== undefined ? String(req.body.pin) : undefined;
      const role = (req.body.role as 'owner' | 'editor' | 'viewer') || 'editor';

      if (!rawUsername || !roomCode) {
        return res.status(400).json({ error: 'Eksik kullanıcı veya oda bilgisi.' });
      }

      const user = await dbGetUser(rawUsername);
      if (!user) {
        return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
      }

      if (!Array.isArray(user.savedRooms)) {
        user.savedRooms = [];
      }

      const nowIso = new Date().toISOString();
      const existingIdx = user.savedRooms.findIndex((r: any) => r.roomCode === roomCode);

      if (existingIdx !== -1) {
        const prev = user.savedRooms[existingIdx];
        user.savedRooms[existingIdx] = {
          ...prev,
          title: title || prev.title,
          pin: pin !== undefined ? pin : prev.pin,
          role: prev.role === 'owner' ? 'owner' : role,
          lastAccessedAt: nowIso,
        };
      } else {
        user.savedRooms.unshift({
          roomCode,
          title: title || `Excel_${roomCode}`,
          pin: pin || '',
          role,
          lastAccessedAt: nowIso,
        });
      }

      await dbUpsertUser(user);
      return res.status(200).json({
        ok: true,
        savedRooms: await enrichUserRooms(user),
      });
    }

    // 4) KULLANICININ ODA GEÇMİŞİNDEN BİR ODAYI KALDIR (REMOVE_USER_ROOM)
    if (action === 'REMOVE_USER_ROOM') {
      const rawUsername = String(req.body.username || '')
        .trim()
        .toLowerCase();
      const roomCode = String(req.body.roomCode || '')
        .trim()
        .toUpperCase();

      if (!rawUsername) {
        return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
      }

      const user = await dbGetUser(rawUsername);
      if (!user) {
        return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
      }

      user.savedRooms = (user.savedRooms || []).filter((r: any) => r.roomCode !== roomCode);
      await dbUpsertUser(user);

      return res.status(200).json({
        ok: true,
        savedRooms: await enrichUserRooms(user),
      });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
