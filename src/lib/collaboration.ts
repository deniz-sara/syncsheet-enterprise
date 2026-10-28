export interface CollaboratorPresence {
  userId: string;
  userName: string;
  userColor: string;
  role: 'editor' | 'viewer';
  activeSheetId: string;
  activeSheetName: string;
  activeCellRef: string; // Örn: "D4"
  row: number;
  col: number;
  lastUpdated: number;
  lastActivityAt: number; // Hücreyi seçtiği veya son tuşa bastığı an (Hareketsizlik kilidi için)
  isTyping?: boolean;
  cursorXPct?: number; // Tablo alanı içindeki yatay fare konumu (%0 - %100)
  cursorYPct?: number; // Tablo alanı içindeki dikey fare konumu (%0 - %100)
  cursorUpdatedAt?: number;
}

/**
 * Bir hücrede başka bir kullanıcı işlem yaparken o hücrenin kilitli kalacağı
 * maksimum hareketsizlik (idle) süresi: 30 Saniye (30.000 ms).
 * Kullanıcı başka hücreye geçerse kilit ANINDA kalkar;
 * hücrede ekranını açık unutursa 30 saniye sonra kilit otomatik düşer.
 */
export const CELL_LOCK_TIMEOUT_MS = 30000;

export function getActiveCellLockOwner(
  collaborators: Record<string, CollaboratorPresence>,
  currentUserId: string,
  sheetId: string,
  row: number,
  col: number,
  now: number = Date.now(),
  timeoutMs: number = CELL_LOCK_TIMEOUT_MS
): {
  lockedBy: CollaboratorPresence;
  remainingSeconds: number;
} | null {
  const list = Object.values(collaborators);
  for (const collab of list) {
    if (!collab || collab.userId === currentUserId) continue;
    if (collab.role === 'viewer') continue; // Salt okunur kullanıcılar hücre kilitlemez

    const sameSheet = !collab.activeSheetId || !sheetId || collab.activeSheetId === sheetId;
    if (sameSheet && collab.row === row && collab.col === col) {
      const activityTime = collab.lastActivityAt || collab.lastUpdated || 0;
      const elapsed = now - activityTime;
      if (elapsed >= 0 && elapsed < timeoutMs) {
        const remainingSeconds = Math.max(1, Math.ceil((timeoutMs - elapsed) / 1000));
        return {
          lockedBy: collab,
          remainingSeconds,
        };
      }
    }
  }
  return null;
}

export type CollabMessage =
  | {
      type: 'PRESENCE_UPDATE';
      roomCode: string;
      presence: CollaboratorPresence;
    }
  | {
      type: 'CELL_UPDATE';
      roomCode: string;
      senderId: string;
      senderName: string;
      sheetId: string;
      row: number;
      col: number;
      value: any;
      formula?: string;
      timestamp: number;
    }
  | {
      type: 'WORKBOOK_SYNC';
      roomCode: string;
      senderId: string;
      senderName: string;
      workbookData: any;
      timestamp: number;
    };

const USER_COLORS = [
  '#059669', // Emerald
  '#2563eb', // Blue
  '#7c3aed', // Violet
  '#d97706', // Amber
  '#db2777', // Pink
  '#0891b2', // Cyan
];

export function getRandomUserColor(seed: string): string {
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = seed.charCodeAt(i) + ((hash << 5) - hash);
  }
  return USER_COLORS[Math.abs(hash) % USER_COLORS.length];
}

export function colIndexToLetter(col: number): string {
  let temp = col;
  let letter = '';
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

export function formatCellRef(row: number, col: number): string {
  return `${colIndexToLetter(col)}${row + 1}`;
}

export class RealtimeRoomClient {
  private channel: BroadcastChannel | null = null;
  private roomCode: string;
  private onMessageCallback: (msg: CollabMessage) => void;

  constructor(roomCode: string, onMessage: (msg: CollabMessage) => void) {
    this.roomCode = roomCode;
    this.onMessageCallback = onMessage;

    if (typeof window !== 'undefined' && 'BroadcastChannel' in window) {
      this.channel = new BroadcastChannel(`syncsheet_room_${roomCode}`);
      this.channel.onmessage = (event) => {
        if (event.data && event.data.roomCode === this.roomCode) {
          this.onMessageCallback(event.data as CollabMessage);
        }
      };
    }
  }

  public sendPresence(presence: CollaboratorPresence) {
    const msg: CollabMessage = {
      type: 'PRESENCE_UPDATE',
      roomCode: this.roomCode,
      presence,
    };
    this.channel?.postMessage(msg);
  }

  public sendCellUpdate(payload: {
    senderId: string;
    senderName: string;
    sheetId: string;
    row: number;
    col: number;
    value: any;
    formula?: string;
  }) {
    const msg: CollabMessage = {
      type: 'CELL_UPDATE',
      roomCode: this.roomCode,
      ...payload,
      timestamp: Date.now(),
    };
    this.channel?.postMessage(msg);
  }

  public sendWorkbookSync(senderId: string, senderName: string, workbookData: any) {
    const msg: CollabMessage = {
      type: 'WORKBOOK_SYNC',
      roomCode: this.roomCode,
      senderId: string(senderId),
      senderName,
      workbookData,
      timestamp: Date.now(),
    };
    this.channel?.postMessage(msg);
  }

  public destroy() {
    this.channel?.close();
    this.channel = null;
  }
}

function string(val: any): string {
  return String(val || '');
}
