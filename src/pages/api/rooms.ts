import type { NextApiRequest, NextApiResponse } from 'next';
import {
  dbDeleteRoom,
  dbGetAllRoomsMap,
  dbGetRoom,
  dbUpsertRoom,
  isSupabaseConfigured,
} from '@/lib/supabaseStorage';

export interface RoomSnapshot {
  id: string;
  roomCode?: string;
  label: string;
  createdBy: string;
  createdAt: string;
  rowCount: number;
  isAutoChange?: boolean;
  workbookData: any;
}

export interface RoomAuditEvent {
  id: string;
  roomCode?: string;
  userId: string;
  userName: string;
  userColor: string;
  action: string; // Örn: "D4 hücresini güncelledi", "Excel aktardı", "Versiyon oluşturdu"
  cellRef?: string;
  sheetName?: string;
  oldValue?: any;
  newValue?: any;
  timestamp: number;
  snapshotId?: string;
}

export interface RoomRecord {
  roomCode: string;
  pin?: string;
  defaultRole: 'editor' | 'viewer';
  protectedColumns: string[]; // Örn: ['A', 'G'] -> Sadece oda sahibi değiştirebilir
  title: string;
  ownerId: string;
  ownerName: string;
  ownerEmail?: string;
  updatedAt: string;
  version: number;
  workbookData: any;
  stats?: {
    totalRows: number;
    totalCells: number;
    compressedKB: number;
  };
  presenceMap: Record<string, any>;
  recentCellDeltas: Array<{
    version: number;
    senderId: string;
    senderName: string;
    sheetId: string;
    row: number;
    col: number;
    value: any;
    formula?: string;
    timestamp: number;
  }>;
  auditLog: RoomAuditEvent[];
  snapshots: RoomSnapshot[];
}

// Sunucu belleğinde canlı fare imleci (cursor) ve aktif hücre koordinatları için ultra-hızlı önbellek
const livePresenceMemory: Record<string, Record<string, any>> = {};

export const config = {
  api: {
    bodyParser: {
      sizeLimit: '50mb',
    },
  },
};

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method === 'GET') {
    const code = String(req.query.code || '')
      .toUpperCase()
      .trim();
    const pin = String(req.query.pin || '').trim();
    const sinceVersion = Number(req.query.sinceVersion || 0);
    const mode = String(req.query.mode || 'full'); // 'full' | 'poll'

    if (!code) {
      const rooms = await dbGetAllRoomsMap();
      const list = Object.values(rooms).map((r: any) => ({
        roomCode: r.roomCode,
        title: r.title,
        ownerName: r.ownerName,
        defaultRole: r.defaultRole || 'editor',
        hasPin: Boolean(r.pin),
        updatedAt: r.updatedAt,
        version: r.version || 1,
        stats: r.stats,
        activeUsersCount: Object.values({
          ...(r.presenceMap || {}),
          ...(livePresenceMemory[r.roomCode] || {}),
        }).filter((p: any) => Date.now() - (p.lastUpdated || 0) < 20000).length,
      }));
      return res.status(200).json({
        rooms: list,
        storageEngine: isSupabaseConfigured() ? 'supabase' : 'local',
      });
    }

    const room: RoomRecord | null = await dbGetRoom(code);
    if (!room) {
      return res.status(404).json({ error: 'Bu oda koduna ait kayıtlı bir tablo bulunamadı.' });
    }

    if (mode !== 'poll' && room.pin && room.pin !== pin) {
      return res
        .status(403)
        .json({ error: 'Oda PIN şifresi hatalı veya eksik.', requiresPin: true });
    }

    // 20 saniyeden eski presence kayıtlarını temizle ve canlı fare imleci önbelleği ile birleştir
    const activePresence: Record<string, any> = {};
    const mergedPresenceSource = {
      ...(room.presenceMap || {}),
      ...(livePresenceMemory[code] || {}),
    };
    Object.entries(mergedPresenceSource).forEach(([uid, p]: [string, any]) => {
      if (Date.now() - (p.lastUpdated || 0) < 20000) {
        activePresence[uid] = p;
      }
    });

    if (mode === 'poll') {
      const deltas = (room.recentCellDeltas || []).filter((d) => d.version > sinceVersion);
      return res.status(200).json({
        roomCode: room.roomCode,
        version: room.version || 1,
        defaultRole: room.defaultRole || 'editor',
        protectedColumns: room.protectedColumns || [],
        presenceMap: activePresence,
        deltas,
        auditLog: (room.auditLog || []).slice(0, 30),
        snapshotsMeta: (room.snapshots || []).map((s) => ({
          id: s.id,
          label: s.label,
          createdBy: s.createdBy,
          createdAt: s.createdAt,
          rowCount: s.rowCount,
        })),
      });
    }

    return res.status(200).json({
      room: {
        ...room,
        presenceMap: activePresence,
      },
      storageEngine: isSupabaseConfigured() ? 'supabase' : 'local',
    });
  }

  if (req.method === 'POST') {
    try {
      const body = req.body || {};
      const actionType = body.actionType || 'SAVE_WORKBOOK';
      const cleanCode = String(body.roomCode || '')
        .toUpperCase()
        .trim();

      if (!cleanCode) {
        return res.status(400).json({ error: 'roomCode zorunludur.' });
      }

      if (actionType === 'DELETE_ROOM') {
        delete livePresenceMemory[cleanCode];
        await dbDeleteRoom(cleanCode);
        return res.status(200).json({
          success: true,
          deletedRoomCode: cleanCode,
        });
      }

      // 0) Canlı Fare İmleci (Mouse Cursor) & Aktif Hücre Hızlı Senkronizasyonu (PRESENCE_SYNC)
      if (actionType === 'PRESENCE_SYNC') {
        if (body.presence && body.presence.userId) {
          livePresenceMemory[cleanCode] = livePresenceMemory[cleanCode] || {};
          livePresenceMemory[cleanCode][body.presence.userId] = {
            ...(livePresenceMemory[cleanCode][body.presence.userId] || {}),
            ...body.presence,
            lastUpdated: Date.now(),
          };
        }
        const activeMap: Record<string, any> = {};
        Object.entries(livePresenceMemory[cleanCode] || {}).forEach(([uid, p]: [string, any]) => {
          if (Date.now() - (p.lastUpdated || 0) < 20000) {
            activeMap[uid] = p;
          }
        });
        return res.status(200).json({
          success: true,
          presenceMap: activeMap,
        });
      }

      let existing: RoomRecord | null = await dbGetRoom(cleanCode);

      // 1) Canlı Hücre Değişikliği & İmleç Senkronizasyonu (DELTA_SYNC)
      if (actionType === 'DELTA_SYNC') {
        if (!existing) {
          existing = {
            roomCode: cleanCode,
            defaultRole: 'editor',
            protectedColumns: ['A'],
            title: body.workbookData?.name || '2026_Kurumsal_Butce_ve_Operasyon_Plani',
            ownerId: body.presence?.userId || 'usr_deniz',
            ownerName: body.presence?.userName || 'Deniz Sara',
            updatedAt: new Date().toISOString(),
            version: 1,
            workbookData: body.workbookData || null,
            presenceMap: {},
            recentCellDeltas: [],
            auditLog: [],
            snapshots: [],
          };
        } else if (body.workbookData && !existing.workbookData) {
          existing.workbookData = body.workbookData;
        }

        if (body.presence && body.presence.userId) {
          existing.presenceMap = existing.presenceMap || {};
          const mergedUserPres = {
            ...(livePresenceMemory[cleanCode]?.[body.presence.userId] || {}),
            ...(existing.presenceMap[body.presence.userId] || {}),
            ...body.presence,
            lastUpdated: Date.now(),
          };
          existing.presenceMap[body.presence.userId] = mergedUserPres;
          livePresenceMemory[cleanCode] = livePresenceMemory[cleanCode] || {};
          livePresenceMemory[cleanCode][body.presence.userId] = mergedUserPres;
        }

        if (body.cellDelta) {
          const {
            sheetId,
            row,
            col,
            value,
            formula,
            senderId,
            senderName,
            senderColor,
            cellRef,
            sheetName,
          } = body.cellDelta;

          existing.version = (existing.version || 1) + 1;
          existing.updatedAt = new Date().toISOString();

          // Eğer tam güncel workbookData gönderildiyse doğrudan onu kaydet, yoksa ilgili hücreyi mevcut değeri silmeden yama
          if (body.workbookData && body.workbookData.sheets) {
            existing.workbookData = body.workbookData;
          } else if (
            existing.workbookData &&
            existing.workbookData.sheets &&
            existing.workbookData.sheets[sheetId]
          ) {
            const targetSheet = existing.workbookData.sheets[sheetId];
            targetSheet.cellData = targetSheet.cellData || {};
            targetSheet.cellData[row] = targetSheet.cellData[row] || {};
            const oldCell = targetSheet.cellData[row][col];
            const nextVal = value !== undefined ? value : oldCell?.v;

            targetSheet.cellData[row][col] = {
              ...(oldCell || {}),
              ...(nextVal !== undefined
                ? {
                    v: nextVal,
                    t: typeof nextVal === 'number' ? 2 : 1,
                  }
                : {}),
              f: formula !== undefined ? formula || undefined : oldCell?.f,
            };
          }

          // Bu odanın kendi versiyon geçmişine (snapshots) hücre düzenleme kaydını ekle
          const snapId = body.snapshotId || `snap_edit_${Date.now()}`;
          const displaySummary =
            formula ||
            (value !== undefined && value !== '' ? String(value) : 'Biçimlendirildi');
          if (existing.workbookData) {
            existing.snapshots = existing.snapshots || [];
            existing.snapshots.unshift({
              id: snapId,
              roomCode: cleanCode,
              label: `${cellRef || `R${row + 1}C${col + 1}`} düzenlendi (${displaySummary})`,
              createdBy: senderName,
              createdAt: new Date().toISOString(),
              rowCount: existing.stats?.totalRows || 10,
              isAutoChange: true,
              workbookData: JSON.parse(JSON.stringify(existing.workbookData)),
            });
            if (existing.snapshots.length > 25) {
              existing.snapshots = existing.snapshots.slice(0, 25);
            }
          }

          // Audit log kaydı ekle (sadece bu odaya ait)
          existing.auditLog = existing.auditLog || [];
          existing.auditLog.unshift({
            id: `aud_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
            roomCode: cleanCode,
            userId: senderId,
            userName: senderName,
            userColor: senderColor || '#10b981',
            action: `${cellRef || `R${row + 1}C${col + 1}`} hücresini güncelledi`,
            cellRef,
            sheetName,
            newValue: displaySummary,
            timestamp: Date.now(),
            snapshotId: snapId,
          });
          if (existing.auditLog.length > 100) {
            existing.auditLog = existing.auditLog.slice(0, 100);
          }

          if (value !== undefined || formula !== undefined) {
            existing.recentCellDeltas = existing.recentCellDeltas || [];
            existing.recentCellDeltas.push({
              version: existing.version,
              senderId,
              senderName,
              sheetId,
              row,
              col,
              value,
              formula,
              timestamp: Date.now(),
            });
            if (existing.recentCellDeltas.length > 200) {
              existing.recentCellDeltas = existing.recentCellDeltas.slice(-200);
            }
          }
        }

        await dbUpsertRoom(existing);
        return res.status(200).json({
          success: true,
          version: existing.version,
        });
      }

      // 2) Zaman Makinesi Snapshot (Versiyon) Oluşturma
      if (actionType === 'CREATE_SNAPSHOT' && existing) {
        const snap: RoomSnapshot = {
          id: `snap_${Date.now()}`,
          roomCode: cleanCode,
          label: body.label || `Anlık Görüntü #${(existing.snapshots?.length || 0) + 1}`,
          createdBy: body.userName || existing.ownerName,
          createdAt: new Date().toISOString(),
          rowCount: existing.stats?.totalRows || 0,
          workbookData: body.workbookData || existing.workbookData,
        };
        existing.snapshots = existing.snapshots || [];
        existing.snapshots.unshift(snap);
        if (existing.snapshots.length > 15) {
          existing.snapshots = existing.snapshots.slice(0, 15);
        }

        existing.auditLog = existing.auditLog || [];
        existing.auditLog.unshift({
          id: `aud_${Date.now()}`,
          roomCode: cleanCode,
          userId: body.userId || 'usr',
          userName: body.userName || existing.ownerName,
          userColor: '#10b981',
          action: `"${snap.label}" versiyon kaydını oluşturdu`,
          timestamp: Date.now(),
          snapshotId: snap.id,
        });

        await dbUpsertRoom(existing);
        return res.status(200).json({
          success: true,
          snapshots: existing.snapshots,
        });
      }

      // 3) Snapshot Geri Yükleme (Restore)
      if (actionType === 'RESTORE_SNAPSHOT' && existing) {
        const targetSnap = (existing.snapshots || []).find((s) => s.id === body.snapshotId);
        if (!targetSnap) {
          return res.status(404).json({ error: 'Versiyon bulunamadı.' });
        }
        existing.workbookData = targetSnap.workbookData;
        existing.version = (existing.version || 1) + 1;
        existing.updatedAt = new Date().toISOString();

        existing.auditLog = existing.auditLog || [];
        existing.auditLog.unshift({
          id: `aud_${Date.now()}`,
          roomCode: cleanCode,
          userId: body.userId || 'usr',
          userName: body.userName || existing.ownerName,
          userColor: '#f59e0b',
          action: `"${targetSnap.label}" versiyonuna geri döndü`,
          timestamp: Date.now(),
          snapshotId: targetSnap.id,
        });

        await dbUpsertRoom(existing);
        return res.status(200).json({
          success: true,
          room: existing,
        });
      }

      // 4) Tam Oda ve Tablo Kaydı (SAVE_WORKBOOK)
      const {
        pin,
        defaultRole,
        protectedColumns,
        title,
        ownerId,
        ownerName,
        ownerEmail,
        workbookData,
        stats,
        silentSwitchSave,
      } = body;

      if (!workbookData) {
        return res.status(400).json({ error: 'workbookData zorunludur.' });
      }

      const nextVersion = silentSwitchSave
        ? existing?.version || 1
        : (existing?.version || 0) + 1;
      const snapId = `snap_${Date.now()}`;
      const initialSnap: RoomSnapshot = {
        id: snapId,
        roomCode: cleanCode,
        label: existing ? `Kayıt Noktası v${nextVersion}` : 'İlk Başlangıç Versiyonu (v1)',
        createdBy: ownerName || 'Deniz Sara',
        createdAt: new Date().toISOString(),
        rowCount: stats?.totalRows || 10,
        workbookData,
      };

      const shouldAddSaveEntry = !silentSwitchSave || !existing;

      const updatedRoom: RoomRecord = {
        roomCode: cleanCode,
        pin: pin !== undefined ? (pin ? String(pin).trim() : undefined) : existing?.pin,
        defaultRole: defaultRole || existing?.defaultRole || 'editor',
        protectedColumns: protectedColumns || existing?.protectedColumns || [],
        title: title || workbookData.name || 'Kurumsal Ortak Tablo',
        ownerId: ownerId || existing?.ownerId || 'usr_owner',
        ownerName: ownerName || existing?.ownerName || 'Deniz Sara',
        ownerEmail: ownerEmail || existing?.ownerEmail || 'deniz.sara@sirket.com.tr',
        updatedAt: new Date().toISOString(),
        version: nextVersion,
        workbookData,
        stats: stats || existing?.stats,
        presenceMap: existing?.presenceMap || {},
        recentCellDeltas: existing?.recentCellDeltas || [],
        auditLog: shouldAddSaveEntry
          ? [
              {
                id: `aud_${Date.now()}`,
                roomCode: cleanCode,
                userId: ownerId || 'usr_owner',
                userName: ownerName || 'Deniz Sara',
                userColor: '#10b981',
                action: existing
                  ? `Tabloyu kaydetti (v${nextVersion})`
                  : `"${title || workbookData.name || cleanCode}" çalışma odasını başlattı`,
                timestamp: Date.now(),
                snapshotId: snapId,
              },
              ...(existing?.auditLog || []).slice(0, 50),
            ]
          : existing?.auditLog || [],
        snapshots: shouldAddSaveEntry
          ? [initialSnap, ...(existing?.snapshots || []).slice(0, 20)]
          : existing?.snapshots || [],
      };

      await dbUpsertRoom(updatedRoom);

      return res.status(200).json({
        success: true,
        roomCode: cleanCode,
        version: nextVersion,
        updatedAt: updatedRoom.updatedAt,
      });
    } catch (err: any) {
      return res.status(500).json({ error: err?.message || 'Sunucu kayıt hatası' });
    }
  }

  return res.status(405).json({ error: 'Method not allowed' });
}
