import type { NextApiRequest, NextApiResponse } from 'next';
import {
  dbDeleteRoom,
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
  username: string; // E-posta adresi (küçük harf normalize edilmiş benzersiz anahtar)
  password: string;
  name: string; // Ad Soyad (Görünür isim)
  email: string; // E-posta adresi
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

function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

async function enrichUserRooms(user: StoredUserRecord): Promise<UserSavedRoomEntry[]> {
  const roomsDb = await dbGetAllRoomsMap();
  const list = Array.isArray(user.savedRooms) ? user.savedRooms : [];

  // Sadece veritabanında gerçekten var olan odaları döndür (silinen odalar çöplük yaratmasın)
  return list
    .filter((entry) => Boolean(entry?.roomCode && roomsDb[entry.roomCode]))
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
  const cleanEmail = (user.email || user.username || '').trim().toLowerCase();
  const enrichedRooms = await enrichUserRooms(user);

  // Eğer silinmiş odalar varsa kullanıcının kaydını da temiz tut
  if (Array.isArray(user.savedRooms) && enrichedRooms.length !== user.savedRooms.length) {
    user.savedRooms = enrichedRooms;
    await dbUpsertUser(user);
  }

  return {
    id: user.id,
    username: cleanEmail,
    name: user.name,
    email: cleanEmail,
    department: user.department,
    color: user.color,
    role: 'editor' as const,
    savedRooms: enrichedRooms,
    storageEngine: isSupabaseConfigured() ? 'supabase' : 'local',
  };
}

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method === 'GET') {
    const emailKey = String(req.query.email || req.query.username || '')
      .trim()
      .toLowerCase();
    if (!emailKey) {
      return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
    }
    const existing = await dbGetUser(emailKey);
    if (!existing) {
      return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
    }
    return res.status(200).json({ user: await sanitizeUserForClient(existing) });
  }

  if (req.method === 'POST') {
    const { action } = req.body || {};

    // 1) YENİ KULLANICI KAYDI (REGISTER) - E-Posta + Ad Soyad Zorunlu, Tekil E-Posta Kontrolü
    if (action === 'REGISTER') {
      const rawEmail = String(req.body.email || req.body.username || '')
        .trim()
        .toLowerCase();
      const password = String(req.body.password || '').trim();
      const name = String(req.body.name || '').trim();
      const department = String(req.body.department || '').trim() || 'Genel';

      if (!rawEmail || !isValidEmail(rawEmail)) {
        return res
          .status(400)
          .json({ error: 'Lütfen geçerli bir e-posta adresi girin (örn: ad.soyad@sirket.com).' });
      }
      if (!name || name.length < 2) {
        return res
          .status(400)
          .json({ error: 'Lütfen Ad Soyad bilginizi eksiksiz girin.' });
      }
      if (!password || password.length < 3) {
        return res.status(400).json({ error: 'Şifre en az 3 karakter olmalıdır.' });
      }

      const existing = await dbGetUser(rawEmail);
      if (existing) {
        return res.status(409).json({
          error:
            'Bu e-posta adresi ile kayıtlı bir kullanıcı zaten mevcut. Aynı e-posta adresi ile ikinci bir hesap oluşturulamaz.',
        });
      }

      const nowIso = new Date().toISOString();
      const newUser: StoredUserRecord = {
        id: `usr_${rawEmail.replace(/[^a-z0-9]/g, '_')}`,
        username: rawEmail,
        email: rawEmail,
        password,
        name,
        department,
        color: pickColor(rawEmail),
        createdAt: nowIso,
        savedRooms: [], // Varsayılan (default) hiçbir oda eklenmez!
      };

      await dbUpsertUser(newUser);

      return res.status(200).json({
        ok: true,
        user: await sanitizeUserForClient(newUser),
      });
    }

    // 2) KULLANICI GİRİŞİ (LOGIN) - E-Posta Adresi ve Şifre ile
    if (action === 'LOGIN') {
      const rawEmail = String(req.body.email || req.body.username || '')
        .trim()
        .toLowerCase();
      const password = String(req.body.password || '').trim();

      if (!rawEmail) {
        return res.status(400).json({ error: 'Lütfen e-posta adresinizi girin.' });
      }

      const existing = await dbGetUser(rawEmail);
      if (!existing) {
        return res.status(404).json({
          error:
            'Bu e-posta adresine ait hesap bulunamadı. "Kayıt Ol" sekmesinden yeni hesap oluşturabilirsiniz.',
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
      const rawEmail = String(req.body.email || req.body.username || '')
        .trim()
        .toLowerCase();
      const roomCode = String(req.body.roomCode || '')
        .trim()
        .toUpperCase();
      const title = String(req.body.title || '').trim();
      const pin = req.body.pin !== undefined ? String(req.body.pin) : undefined;
      const role = (req.body.role as 'owner' | 'editor' | 'viewer') || 'editor';

      if (!rawEmail || !roomCode) {
        return res.status(400).json({ error: 'Eksik kullanıcı veya oda bilgisi.' });
      }

      const user = await dbGetUser(rawEmail);
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

    // 4) ODAYI SİL / KALDIR — Veritabanından da (Supabase & Local) Kalıcı Olarak Siler!
    if (action === 'REMOVE_USER_ROOM') {
      const rawEmail = String(req.body.email || req.body.username || '')
        .trim()
        .toLowerCase();
      const roomCode = String(req.body.roomCode || '')
        .trim()
        .toUpperCase();

      if (!rawEmail) {
        return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
      }

      const user = await dbGetUser(rawEmail);
      if (!user) {
        return res.status(404).json({ error: 'Kullanıcı bulunamadı.' });
      }

      user.savedRooms = (user.savedRooms || []).filter((r: any) => r.roomCode !== roomCode);

      // Odayı veritabanından da (syncsheet_rooms ve .data/rooms.json) kalıcı olarak sil
      if (roomCode) {
        await dbDeleteRoom(roomCode);
      }

      await dbUpsertUser(user);

      return res.status(200).json({
        ok: true,
        savedRooms: await enrichUserRooms(user),
      });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
