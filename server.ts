import 'dotenv/config';
import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import { createServer as createViteServer } from 'vite';
import { firebaseProjectId, firestoreDatabaseId } from './src/lib/firebase-admin.ts';
import {
  authenticateWithCredentials,
  checkLoginRateLimit,
  clearFailedLoginAttempts,
  deleteAccountCredential,
  getCredentialMetaMap,
  getSanitizedPortalSnapshot,
  mergeSyncedCredentials,
  normalizeAccountId,
  normalizeUsername,
  recordFailedLoginAttempt,
  savePortalSnapshot,
  upsertAccountCredential,
  type PortalSnapshotData,
  type StoredCredential,
} from './src/lib/server-store.ts';
import { requireAuth, type AuthenticatedRequest } from './src/middleware/auth.ts';
import type { AccountRole } from './src/types.ts';

const PORT = Number(process.env.PORT) || 3000;
const isProd = process.env.NODE_ENV === 'production';

async function startServer() {
  const app = express();
  app.disable('x-powered-by');
  app.set('trust proxy', 1);
  app.use(express.json({ limit: '15mb' }));

  app.get('/api/health', (_req, res) => {
    res.json({
      ok: true,
      service: 'portal-eguru-api',
      projectId: firebaseProjectId,
      databaseId: firestoreDatabaseId,
      timestamp: new Date().toISOString(),
    });
  });

  app.get('/api/session', requireAuth, (req: AuthenticatedRequest, res) => {
    res.json({
      ok: true,
      user: req.user,
      serverTime: new Date().toISOString(),
    });
  });

  app.get('/api/auth/me', requireAuth, (req: AuthenticatedRequest, res) => {
    res.json({
      ok: true,
      user: req.user,
      serverTime: new Date().toISOString(),
    });
  });

  app.post('/api/auth/verify', requireAuth, (req: AuthenticatedRequest, res) => {
    const academicYear =
      typeof req.body?.academicYear === 'string' ? req.body.academicYear.trim() : '';
    const semester = req.body?.semester === 'Genap' ? 'Genap' : 'Ganjil';

    if (academicYear.length > 32) {
      res.status(400).json({ ok: false, error: 'Format tahun ajaran tidak valid.' });
      return;
    }

    res.json({
      ok: true,
      user: req.user,
      academicYear,
      semester,
      verifiedAt: new Date().toISOString(),
    });
  });

  app.post('/api/auth/login-credentials', (req, res) => {
    const username =
      typeof req.body?.username === 'string' ? normalizeUsername(req.body.username) : '';
    const password =
      typeof req.body?.password === 'string' ? req.body.password : '';

    const rateKey = `${req.ip || 'unknown'}:${username || 'empty'}`;
    if (!checkLoginRateLimit(rateKey)) {
      res.status(429).json({
        ok: false,
        error: 'Terlalu banyak percobaan login gagal. Coba lagi beberapa menit lagi.',
      });
      return;
    }

    const result = authenticateWithCredentials(username, password);
    if (!result.ok) {
      recordFailedLoginAttempt(rateKey);
      res.status(result.statusCode).json({ ok: false, error: result.error });
      return;
    }

    clearFailedLoginAttempts(rateKey);
    res.json({
      ok: true,
      token: result.token,
      user: result.user,
      snapshot: result.snapshot,
    });
  });

  app.get('/api/accounts/credentials-meta', requireAuth, (req: AuthenticatedRequest, res) => {
    if (!req.user?.isStaff) {
      res.status(403).json({ ok: false, error: 'Akses ditolak.' });
      return;
    }
    res.json({
      ok: true,
      items: getCredentialMetaMap(),
    });
  });

  app.post('/api/accounts/credentials', requireAuth, (req: AuthenticatedRequest, res) => {
    const caller = req.user;
    if (!caller || !caller.isStaff) {
      res.status(403).json({ ok: false, error: 'Akses ditolak.' });
      return;
    }

    const body = req.body as Record<string, unknown> | undefined;
    if (!body || typeof body !== 'object') {
      res.status(400).json({ ok: false, error: 'Payload tidak valid.' });
      return;
    }

    const accountId = normalizeAccountId(
      typeof body.accountId === 'string'
        ? body.accountId
        : typeof body.email === 'string'
          ? body.email
          : ''
    );
    const previousId =
      typeof body.previousId === 'string' ? normalizeAccountId(body.previousId) : undefined;

    if (!accountId) {
      res.status(400).json({ ok: false, error: 'ID akun wajib diisi.' });
      return;
    }

    const isSelf = caller.accountId === accountId;
    if (!caller.isAdmin && !isSelf) {
      res.status(403).json({
        ok: false,
        error: 'Hanya Admin yang dapat mengelola akun pengguna lain.',
      });
      return;
    }

    if (!caller.isAdmin && isSelf) {
      const password = typeof body.password === 'string' ? body.password : '';
      if (!password) {
        res.status(200).json({
          ok: true,
          meta: getCredentialMetaMap()[accountId] || { username: caller.username || '', hasPassword: false },
        });
        return;
      }
      // Non-admin self password update preserves existing role/status/username
      const snap = getSanitizedPortalSnapshot().database;
      const existingAcc = snap?.accounts.find((a) => normalizeAccountId(a.id) === accountId);
      const result = upsertAccountCredential({
        accountId,
        username: caller.username || existingAcc?.username || '',
        password,
        email: caller.email,
        name: existingAcc?.name || caller.name,
        role: caller.role,
        nip: existingAcc?.nip || '',
        status: 'Aktif',
      });
      if (!result.ok) {
        res.status(result.statusCode).json({ ok: false, error: result.error });
        return;
      }
      res.json({ ok: true, meta: result.meta });
      return;
    }

    const result = upsertAccountCredential({
      accountId,
      previousId,
      username: typeof body.username === 'string' ? body.username : '',
      password: typeof body.password === 'string' ? body.password : '',
      email: typeof body.email === 'string' ? body.email : accountId,
      name: typeof body.name === 'string' ? body.name : accountId,
      role: (typeof body.role === 'string' ? body.role : 'Guru Mapel') as AccountRole,
      nip: typeof body.nip === 'string' ? body.nip : '',
      status: body.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif',
      signatureImage: typeof body.signatureImage === 'string' ? body.signatureImage : undefined,
      schoolStampImage:
        typeof body.schoolStampImage === 'string' ? body.schoolStampImage : undefined,
    });

    if (!result.ok) {
      res.status(result.statusCode).json({ ok: false, error: result.error });
      return;
    }

    res.json({
      ok: true,
      meta: result.meta,
      credentialRecord: result.record,
    });
  });

  app.post('/api/accounts/credentials-sync', requireAuth, (req: AuthenticatedRequest, res) => {
    const caller = req.user;
    if (!caller || !caller.isAdmin) {
      res.status(403).json({ ok: false, error: 'Hanya Admin yang dapat menyinkronkan kredensial.' });
      return;
    }
    const body = req.body as { credentials?: StoredCredential[] } | undefined;
    const incoming = Array.isArray(body?.credentials) ? body.credentials : [];
    const merged = mergeSyncedCredentials(incoming);
    res.json({
      ok: true,
      credentials: merged,
    });
  });

  app.delete('/api/accounts/credentials/:accountId', requireAuth, (req: AuthenticatedRequest, res) => {
    const caller = req.user;
    if (!caller || !caller.isAdmin) {
      res.status(403).json({ ok: false, error: 'Hanya Admin yang dapat menghapus akun.' });
      return;
    }
    const targetId = normalizeAccountId(req.params.accountId || '');
    if (!targetId || targetId === caller.accountId) {
      res.status(400).json({ ok: false, error: 'Akun ini tidak dapat dihapus.' });
      return;
    }
    deleteAccountCredential(targetId);
    res.json({ ok: true });
  });

  app.get('/api/portal/state', requireAuth, (req: AuthenticatedRequest, res) => {
    if (!req.user?.isStaff) {
      res.status(403).json({ ok: false, error: 'Akun Anda belum aktif di sistem.' });
      return;
    }
    const snapshot = getSanitizedPortalSnapshot();
    res.json({ ok: true, ...snapshot });
  });

  app.put('/api/portal/state', requireAuth, (req: AuthenticatedRequest, res) => {
    const caller = req.user;
    if (!caller || !caller.isStaff) {
      res.status(403).json({ ok: false, error: 'Akun Anda belum aktif di sistem.' });
      return;
    }
    const body = req.body as { database?: PortalSnapshotData; markFirestoreSynced?: boolean } | undefined;
    if (!body?.database || typeof body.database !== 'object' || !body.database.settings) {
      res.status(400).json({ ok: false, error: 'Data portal tidak valid.' });
      return;
    }
    const saved = savePortalSnapshot(body.database, {
      authType: caller.authType,
      isAdmin: caller.isAdmin,
      markFirestoreSynced: Boolean(body.markFirestoreSynced),
    });
    res.json({ ok: true, ...saved });
  });

  app.use('/api', (_req, res) => {
    res.status(404).json({ ok: false, error: 'Endpoint API tidak ditemukan.' });
  });

  if (isProd) {
    const distDir = path.resolve(process.cwd(), 'dist');
    const indexHtml = path.join(distDir, 'index.html');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      if (fs.existsSync(indexHtml)) {
        res.sendFile(indexHtml);
      } else {
        res.status(404).send('Build output not found. Run npm run build first.');
      }
    });
  } else {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Portal E-Guru server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
