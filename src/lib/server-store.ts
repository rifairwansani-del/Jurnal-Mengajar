import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { AccountRole, PortalDatabase, UserAccount } from '../types.ts';

export const BOOTSTRAP_ADMIN_EMAIL = 'rifairwansani@sma1kudus.sch.id';
const VALID_ROLES: AccountRole[] = [
  'Admin',
  'Guru Mapel',
  'Wali Kelas',
  'Guru BK',
  'Kepala Sekolah',
];

const DATA_DIR = path.resolve(process.cwd(), 'data');
const CREDENTIALS_FILE = 'credentials.json';
const SNAPSHOT_FILE = 'portal-snapshot.json';

const SESSION_SECRET =
  (process.env.SESSION_SECRET && process.env.SESSION_SECRET.trim()) ||
  crypto.randomBytes(32).toString('hex');

export interface StoredCredential {
  accountId: string;
  username: string;
  email: string;
  name: string;
  role: AccountRole;
  nip: string;
  status: 'Aktif' | 'Nonaktif';
  passwordHash: string;
  salt: string;
  updatedAt: string;
}

export interface CredentialMeta {
  username: string;
  hasPassword: boolean;
}

export interface PortalSnapshotData {
  students: PortalDatabase['students'];
  teachers: PortalDatabase['teachers'];
  schedules: PortalDatabase['schedules'];
  accounts: UserAccount[];
  journals: PortalDatabase['journals'];
  grades: PortalDatabase['grades'];
  settings: PortalDatabase['settings'];
}

interface SnapshotEnvelope {
  database: PortalSnapshotData | null;
  meta: {
    updatedAt: number;
    lastLocalEditAt: number;
    lastFirestoreSyncAt: number;
  };
}

interface CredentialsEnvelope {
  credentials: Record<string, StoredCredential>;
}

export interface LocalTokenPayload {
  typ: 'local';
  accountId: string;
  username: string;
  email: string;
  name: string;
  role: AccountRole;
  iat: number;
  exp: number;
}

function resolveDataFilePath(filename: string): string {
  const safeName = path.basename(filename);
  const resolved = path.resolve(DATA_DIR, safeName);
  if (!resolved.startsWith(DATA_DIR + path.sep)) {
    throw new Error('Invalid storage path.');
  }
  return resolved;
}

function ensureDataDir(): void {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true, mode: 0o700 });
  }
}

function readJsonFile<T>(filename: string, fallback: T): T {
  try {
    const filePath = resolveDataFilePath(filename);
    if (!fs.existsSync(filePath)) return fallback;
    const raw = fs.readFileSync(filePath, 'utf-8');
    if (!raw.trim()) return fallback;
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

function writeJsonFileAtomic(filename: string, payload: unknown): void {
  ensureDataDir();
  const filePath = resolveDataFilePath(filename);
  const tmpPath = resolveDataFilePath(
    `${filename}.${process.pid}.${Date.now()}.tmp`
  );
  fs.writeFileSync(tmpPath, JSON.stringify(payload, null, 2), {
    encoding: 'utf-8',
    mode: 0o600,
  });
  fs.renameSync(tmpPath, filePath);
}

export function normalizeAccountId(raw: string): string {
  return String(raw || '')
    .trim()
    .toLowerCase()
    .replace(/\//g, '');
}

export function normalizeUsername(raw: string): string {
  return String(raw || '')
    .trim()
    .toLowerCase();
}

export function isValidUsernameFormat(username: string): boolean {
  return /^[a-z0-9._-]{3,40}$/.test(username);
}

export function hashPassword(password: string): { salt: string; hash: string } {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return { salt, hash };
}

export function verifyPassword(
  password: string,
  salt: string,
  expectedHash: string
): boolean {
  if (!password || !salt || !expectedHash) return false;
  try {
    const derived = crypto.scryptSync(password, salt, 64);
    const expected = Buffer.from(expectedHash, 'hex');
    if (derived.length !== expected.length) return false;
    return crypto.timingSafeEqual(derived, expected);
  } catch {
    return false;
  }
}

export function signLocalSessionToken(
  user: Omit<LocalTokenPayload, 'typ' | 'iat' | 'exp'>
): string {
  const nowSec = Math.floor(Date.now() / 1000);
  const payload: LocalTokenPayload = {
    typ: 'local',
    ...user,
    iat: nowSec,
    exp: nowSec + 60 * 60 * 24, // 24 hours
  };
  const payloadB64 = Buffer.from(JSON.stringify(payload), 'utf-8').toString(
    'base64url'
  );
  const sigB64 = crypto
    .createHmac('sha256', SESSION_SECRET)
    .update(payloadB64)
    .digest('base64url');
  return `local.${payloadB64}.${sigB64}`;
}

export function verifyLocalSessionToken(
  token: string
): LocalTokenPayload | null {
  if (!token || !token.startsWith('local.')) return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [, payloadB64, sigB64] = parts;
  if (!payloadB64 || !sigB64) return null;

  try {
    const expectedSig = crypto
      .createHmac('sha256', SESSION_SECRET)
      .update(payloadB64)
      .digest();
    const providedSig = Buffer.from(sigB64, 'base64url');
    if (expectedSig.length !== providedSig.length) return null;
    if (!crypto.timingSafeEqual(expectedSig, providedSig)) return null;

    const decoded = JSON.parse(
      Buffer.from(payloadB64, 'base64url').toString('utf-8')
    ) as LocalTokenPayload;
    if (decoded.typ !== 'local') return null;
    const nowSec = Math.floor(Date.now() / 1000);
    if (!decoded.exp || decoded.exp < nowSec) return null;
    return decoded;
  } catch {
    return null;
  }
}

function loadCredentialsEnvelope(): CredentialsEnvelope {
  const env = readJsonFile<CredentialsEnvelope>(CREDENTIALS_FILE, {
    credentials: {},
  });
  return {
    credentials:
      env && typeof env.credentials === 'object' && env.credentials !== null
        ? env.credentials
        : {},
  };
}

function loadSnapshotEnvelope(): SnapshotEnvelope {
  const env = readJsonFile<SnapshotEnvelope>(SNAPSHOT_FILE, {
    database: null,
    meta: { updatedAt: 0, lastLocalEditAt: 0, lastFirestoreSyncAt: 0 },
  });
  return {
    database: env?.database ?? null,
    meta: {
      updatedAt: Number(env?.meta?.updatedAt) || 0,
      lastLocalEditAt: Number(env?.meta?.lastLocalEditAt) || 0,
      lastFirestoreSyncAt: Number(env?.meta?.lastFirestoreSyncAt) || 0,
    },
  };
}

export function getCredentialMetaMap(): Record<string, CredentialMeta> {
  const { credentials } = loadCredentialsEnvelope();
  const out: Record<string, CredentialMeta> = {};
  Object.values(credentials).forEach((cred) => {
    const id = normalizeAccountId(cred.accountId);
    if (!id) return;
    out[id] = {
      username: cred.username || '',
      hasPassword: Boolean(cred.passwordHash && cred.salt),
    };
  });
  return out;
}

export function getAccountAccessInfo(accountId: string): {
  exists: boolean;
  status: 'Aktif' | 'Nonaktif';
  role: AccountRole;
  username: string;
  name: string;
  email: string;
} | null {
  const id = normalizeAccountId(accountId);
  if (!id) return null;

  const { credentials } = loadCredentialsEnvelope();
  const cred = credentials[id];

  const { database } = loadSnapshotEnvelope();
  const snapAcc = database?.accounts?.find(
    (a) => normalizeAccountId(a.id || a.email) === id
  );

  if (!cred && !snapAcc) return null;

  const status =
    (snapAcc?.status || cred?.status) === 'Nonaktif' ? 'Nonaktif' : 'Aktif';
  const rawRole = snapAcc?.role || cred?.role || 'Guru Mapel';
  const role: AccountRole = VALID_ROLES.includes(rawRole)
    ? rawRole
    : 'Guru Mapel';

  return {
    exists: true,
    status,
    role,
    username: cred?.username || snapAcc?.username || '',
    name: snapAcc?.name || cred?.name || id,
    email: snapAcc?.email || cred?.email || id,
  };
}

export function upsertAccountCredential(input: {
  accountId: string;
  previousId?: string;
  username?: string;
  password?: string;
  email: string;
  name: string;
  role: AccountRole;
  nip?: string;
  status: 'Aktif' | 'Nonaktif';
  signatureImage?: string;
  schoolStampImage?: string;
}): { ok: true; meta: CredentialMeta } | { ok: false; statusCode: number; error: string } {
  const accountId = normalizeAccountId(input.accountId || input.email);
  const emailLower = normalizeAccountId(input.email || accountId);
  const previousId = input.previousId
    ? normalizeAccountId(input.previousId)
    : undefined;
  const cleanUsername = normalizeUsername(input.username || '');
  const rawPassword = typeof input.password === 'string' ? input.password : '';

  if (!accountId || !emailLower.includes('@')) {
    return {
      ok: false,
      statusCode: 400,
      error: 'ID/Email akun tidak valid.',
    };
  }

  if (cleanUsername && !isValidUsernameFormat(cleanUsername)) {
    return {
      ok: false,
      statusCode: 400,
      error:
        'Username harus 3–40 karakter dan hanya boleh huruf kecil, angka, titik (.), garis bawah (_), atau tanda hubung (-).',
    };
  }

  if (rawPassword && rawPassword.length < 6) {
    return {
      ok: false,
      statusCode: 400,
      error: 'Password minimal harus 6 karakter.',
    };
  }

  const credEnv = loadCredentialsEnvelope();
  const existing =
    credEnv.credentials[accountId] ||
    (previousId ? credEnv.credentials[previousId] : undefined);

  if (cleanUsername) {
    const conflict = Object.values(credEnv.credentials).find(
      (c) =>
        c.username === cleanUsername &&
        c.accountId !== accountId &&
        c.accountId !== previousId
    );
    if (conflict) {
      return {
        ok: false,
        statusCode: 409,
        error: `Username "${cleanUsername}" sudah digunakan oleh akun lain.`,
      };
    }
  }

  if (previousId && previousId !== accountId && previousId !== BOOTSTRAP_ADMIN_EMAIL) {
    delete credEnv.credentials[previousId];
  }

  let passwordHash = existing?.passwordHash || '';
  let salt = existing?.salt || '';
  if (rawPassword) {
    const hashed = hashPassword(rawPassword);
    passwordHash = hashed.hash;
    salt = hashed.salt;
  }

  const validRole: AccountRole = VALID_ROLES.includes(input.role)
    ? input.role
    : 'Guru Mapel';
  const validStatus: 'Aktif' | 'Nonaktif' =
    accountId === BOOTSTRAP_ADMIN_EMAIL
      ? 'Aktif'
      : input.status === 'Nonaktif'
        ? 'Nonaktif'
        : 'Aktif';

  const record: StoredCredential = {
    accountId,
    username: cleanUsername,
    email: emailLower,
    name: String(input.name || emailLower).trim().slice(0, 160),
    role: accountId === BOOTSTRAP_ADMIN_EMAIL ? 'Admin' : validRole,
    nip: String(input.nip || '').trim().slice(0, 64),
    status: validStatus,
    passwordHash,
    salt,
    updatedAt: new Date().toISOString(),
  };

  credEnv.credentials[accountId] = record;
  writeJsonFileAtomic(CREDENTIALS_FILE, credEnv);

  // Keep snapshot accounts in sync immediately so username/password login works right away
  const snapEnv = loadSnapshotEnvelope();
  if (snapEnv.database) {
    const filtered = snapEnv.database.accounts.filter(
      (a) =>
        normalizeAccountId(a.id) !== accountId &&
        (!previousId || normalizeAccountId(a.id) !== previousId)
    );
    const prevSnapAcc = snapEnv.database.accounts.find(
      (a) =>
        normalizeAccountId(a.id) === accountId ||
        (previousId && normalizeAccountId(a.id) === previousId)
    );
    const nextAccount: UserAccount = {
      id: accountId,
      name: record.name,
      email: record.email,
      username: record.username,
      hasPassword: Boolean(record.passwordHash && record.salt),
      role: record.role,
      nip: record.nip,
      status: record.status,
      lastLogin: prevSnapAcc?.lastLogin || '-',
      signatureImage:
        typeof input.signatureImage === 'string'
          ? input.signatureImage.slice(0, 350000)
          : prevSnapAcc?.signatureImage || '',
      schoolStampImage:
        typeof input.schoolStampImage === 'string'
          ? input.schoolStampImage.slice(0, 350000)
          : prevSnapAcc?.schoolStampImage || '',
    };
    snapEnv.database.accounts = [...filtered, nextAccount];
    snapEnv.meta.updatedAt = Date.now();
    writeJsonFileAtomic(SNAPSHOT_FILE, snapEnv);
  }

  return {
    ok: true,
    meta: {
      username: record.username,
      hasPassword: Boolean(record.passwordHash && record.salt),
    },
  };
}

export function deleteAccountCredential(accountId: string): void {
  const id = normalizeAccountId(accountId);
  if (!id || id === BOOTSTRAP_ADMIN_EMAIL) return;

  const credEnv = loadCredentialsEnvelope();
  if (credEnv.credentials[id]) {
    delete credEnv.credentials[id];
    writeJsonFileAtomic(CREDENTIALS_FILE, credEnv);
  }

  const snapEnv = loadSnapshotEnvelope();
  if (snapEnv.database) {
    snapEnv.database.accounts = snapEnv.database.accounts.filter(
      (a) => normalizeAccountId(a.id) !== id
    );
    snapEnv.meta.updatedAt = Date.now();
    writeJsonFileAtomic(SNAPSHOT_FILE, snapEnv);
  }
}

// Simple in-memory rate limiter for login attempts
const loginAttempts = new Map<string, { count: number; resetAt: number }>();

export function checkLoginRateLimit(key: string): boolean {
  const now = Date.now();
  const entry = loginAttempts.get(key);
  if (!entry || entry.resetAt <= now) {
    return true;
  }
  return entry.count < 8;
}

export function recordFailedLoginAttempt(key: string): void {
  const now = Date.now();
  const entry = loginAttempts.get(key);
  if (!entry || entry.resetAt <= now) {
    loginAttempts.set(key, { count: 1, resetAt: now + 10 * 60 * 1000 });
    return;
  }
  entry.count += 1;
}

export function clearFailedLoginAttempts(key: string): void {
  loginAttempts.delete(key);
}

export function authenticateWithCredentials(
  usernameInput: string,
  passwordInput: string
):
  | {
      ok: true;
      token: string;
      user: {
        accountId: string;
        username: string;
        email: string;
        name: string;
        role: AccountRole;
        nip: string;
      };
      snapshot: SnapshotEnvelope;
    }
  | { ok: false; statusCode: number; error: string } {
  const cleanUsername = normalizeUsername(usernameInput);
  if (!cleanUsername || !passwordInput) {
    return {
      ok: false,
      statusCode: 400,
      error: 'Username dan password wajib diisi.',
    };
  }

  const credEnv = loadCredentialsEnvelope();
  const cred = Object.values(credEnv.credentials).find(
    (c) => c.username === cleanUsername
  );

  if (!cred || !cred.passwordHash || !cred.salt) {
    return {
      ok: false,
      statusCode: 401,
      error: 'Username atau password salah.',
    };
  }

  if (!verifyPassword(passwordInput, cred.salt, cred.passwordHash)) {
    return {
      ok: false,
      statusCode: 401,
      error: 'Username atau password salah.',
    };
  }

  const accessInfo = getAccountAccessInfo(cred.accountId);
  const currentStatus = accessInfo?.status || cred.status;
  if (currentStatus !== 'Aktif') {
    return {
      ok: false,
      statusCode: 403,
      error:
        'Akun Anda berstatus Nonaktif. Hubungi Admin sekolah untuk mengaktifkan kembali.',
    };
  }

  const nowStr = new Date().toISOString().slice(0, 16).replace('T', ' ');
  const resolvedRole =
    cred.accountId === BOOTSTRAP_ADMIN_EMAIL
      ? 'Admin'
      : accessInfo?.role || cred.role;
  const resolvedName = accessInfo?.name || cred.name;

  // Update lastLogin in snapshot if available
  const snapEnv = loadSnapshotEnvelope();
  if (snapEnv.database) {
    snapEnv.database.accounts = snapEnv.database.accounts.map((a) =>
      normalizeAccountId(a.id) === cred.accountId
        ? {
            ...a,
            lastLogin: nowStr,
            username: cred.username,
            hasPassword: true,
          }
        : a
    );
    snapEnv.meta.updatedAt = Date.now();
    writeJsonFileAtomic(SNAPSHOT_FILE, snapEnv);
  }

  const token = signLocalSessionToken({
    accountId: cred.accountId,
    username: cred.username,
    email: cred.email,
    name: resolvedName,
    role: resolvedRole,
  });

  return {
    ok: true,
    token,
    user: {
      accountId: cred.accountId,
      username: cred.username,
      email: cred.email,
      name: resolvedName,
      role: resolvedRole,
      nip: cred.nip || '',
    },
    snapshot: getSanitizedPortalSnapshot(),
  };
}

function stripPasswordFields(acc: UserAccount, meta?: CredentialMeta): UserAccount {
  const { ...safeAcc } = acc as UserAccount & { password?: unknown; passwordHash?: unknown };
  delete safeAcc.password;
  delete safeAcc.passwordHash;
  return {
    ...safeAcc,
    username: meta?.username || safeAcc.username || '',
    hasPassword: meta ? meta.hasPassword : Boolean(safeAcc.hasPassword),
  };
}

export function getSanitizedPortalSnapshot(): SnapshotEnvelope {
  const snapEnv = loadSnapshotEnvelope();
  const metaMap = getCredentialMetaMap();
  if (!snapEnv.database) {
    return snapEnv;
  }
  return {
    database: {
      ...snapEnv.database,
      accounts: Array.isArray(snapEnv.database.accounts)
        ? snapEnv.database.accounts.map((acc) => {
            const id = normalizeAccountId(acc.id || acc.email);
            return stripPasswordFields(acc, metaMap[id]);
          })
        : [],
    },
    meta: snapEnv.meta,
  };
}

export function savePortalSnapshot(
  incoming: PortalSnapshotData,
  options: {
    authType: 'google' | 'credentials';
    isAdmin: boolean;
    markFirestoreSynced?: boolean;
  }
): SnapshotEnvelope {
  const prev = loadSnapshotEnvelope();
  const metaMap = getCredentialMetaMap();
  const now = Date.now();

  const sanitizedIncomingAccounts = Array.isArray(incoming.accounts)
    ? incoming.accounts.map((acc) => {
        const id = normalizeAccountId(acc.id || acc.email);
        return stripPasswordFields(acc, metaMap[id]);
      })
    : [];

  // Preserve existing accounts list if non-admin calls PUT /api/portal/state
  const nextAccounts = options.isAdmin
    ? sanitizedIncomingAccounts
    : prev.database?.accounts && prev.database.accounts.length > 0
      ? prev.database.accounts
      : sanitizedIncomingAccounts;

  const nextDb: PortalSnapshotData = {
    students: Array.isArray(incoming.students) ? incoming.students : [],
    teachers: Array.isArray(incoming.teachers) ? incoming.teachers : [],
    schedules: Array.isArray(incoming.schedules) ? incoming.schedules : [],
    accounts: nextAccounts,
    journals: Array.isArray(incoming.journals) ? incoming.journals : [],
    grades: Array.isArray(incoming.grades) ? incoming.grades : [],
    settings: incoming.settings,
  };

  // Keep credentials.json metadata aligned when Admin syncs accounts
  if (options.isAdmin) {
    const credEnv = loadCredentialsEnvelope();
    let changed = false;
    nextAccounts.forEach((acc) => {
      const id = normalizeAccountId(acc.id || acc.email);
      const existing = credEnv.credentials[id];
      if (existing) {
        existing.name = acc.name;
        existing.role = acc.role;
        existing.nip = acc.nip;
        existing.status = acc.status;
        if (acc.username && isValidUsernameFormat(normalizeUsername(acc.username))) {
          existing.username = normalizeUsername(acc.username);
        }
        changed = true;
      }
    });
    if (changed) {
      writeJsonFileAtomic(CREDENTIALS_FILE, credEnv);
    }
  }

  const nextEnvelope: SnapshotEnvelope = {
    database: nextDb,
    meta: {
      updatedAt: now,
      lastLocalEditAt:
        options.authType === 'credentials' ? now : prev.meta.lastLocalEditAt,
      lastFirestoreSyncAt: options.markFirestoreSynced
        ? now
        : options.authType === 'google'
          ? now
          : prev.meta.lastFirestoreSyncAt,
    },
  };

  writeJsonFileAtomic(SNAPSHOT_FILE, nextEnvelope);
  return getSanitizedPortalSnapshot();
}
