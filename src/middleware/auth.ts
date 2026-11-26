import type { NextFunction, Request, Response } from 'express';
import { adminAuth } from '../lib/firebase-admin.ts';
import {
  BOOTSTRAP_ADMIN_EMAIL,
  getAccountAccessInfo,
  normalizeAccountId,
  verifyLocalSessionToken,
} from '../lib/server-store.ts';
import type { AccountRole } from '../types.ts';

export { BOOTSTRAP_ADMIN_EMAIL };

export interface VerifiedSessionUser {
  authType: 'google' | 'credentials';
  uid: string;
  accountId: string;
  username?: string;
  email: string;
  emailVerified: boolean;
  name: string;
  role: AccountRole;
  picture?: string;
  isBootstrapAdmin: boolean;
  isAdmin: boolean;
  isStaff: boolean;
}

export interface AuthenticatedRequest extends Request {
  user?: VerifiedSessionUser;
}

export async function requireAuth(
  req: AuthenticatedRequest,
  res: Response,
  next: NextFunction
): Promise<void> {
  try {
    const header = req.headers.authorization;
    if (!header || !header.startsWith('Bearer ')) {
      res.status(401).json({ ok: false, error: 'Token otorisasi tidak ditemukan.' });
      return;
    }

    const token = header.slice('Bearer '.length).trim();
    if (!token) {
      res.status(401).json({ ok: false, error: 'Token otorisasi kosong.' });
      return;
    }

    if (token.startsWith('local.')) {
      const payload = verifyLocalSessionToken(token);
      if (!payload) {
        res.status(401).json({ ok: false, error: 'Sesi login tidak valid atau telah kedaluwarsa.' });
        return;
      }

      const accessInfo = getAccountAccessInfo(payload.accountId);
      if (!accessInfo || accessInfo.status !== 'Aktif') {
        res.status(403).json({
          ok: false,
          error: 'Akun tidak aktif atau telah dihapus dari sistem.',
        });
        return;
      }

      const accountId = normalizeAccountId(payload.accountId);
      const isBootstrapAdmin = accountId === BOOTSTRAP_ADMIN_EMAIL;
      const role = isBootstrapAdmin ? 'Admin' : accessInfo.role;

      req.user = {
        authType: 'credentials',
        uid: `local:${accountId}`,
        accountId,
        username: payload.username,
        email: accessInfo.email,
        emailVerified: true,
        name: accessInfo.name,
        role,
        isBootstrapAdmin,
        isAdmin: isBootstrapAdmin || role === 'Admin',
        isStaff: true,
      };
      next();
      return;
    }

    const decoded = await adminAuth.verifyIdToken(token);
    const email = normalizeAccountId(decoded.email || '');
    const emailVerified = Boolean(decoded.email_verified);

    if (!email || !emailVerified) {
      res.status(403).json({
        ok: false,
        error: 'Akun Google harus memiliki email yang terverifikasi.',
      });
      return;
    }

    const isBootstrapAdmin = email === BOOTSTRAP_ADMIN_EMAIL;
    const accessInfo = getAccountAccessInfo(email);
    const role: AccountRole = isBootstrapAdmin
      ? 'Admin'
      : accessInfo?.role || 'Guru Mapel';
    const isStaff = isBootstrapAdmin || Boolean(accessInfo && accessInfo.status === 'Aktif');
    const isAdmin = isBootstrapAdmin || (isStaff && role === 'Admin');

    req.user = {
      authType: 'google',
      uid: decoded.uid,
      accountId: email,
      username: accessInfo?.username || undefined,
      email,
      emailVerified,
      name:
        typeof decoded.name === 'string' && decoded.name.trim()
          ? decoded.name.trim()
          : accessInfo?.name || email,
      role,
      picture: typeof decoded.picture === 'string' ? decoded.picture : undefined,
      isBootstrapAdmin,
      isAdmin,
      isStaff,
    };

    next();
  } catch (err) {
    console.error('Auth token verification failed:', err instanceof Error ? err.message : err);
    res.status(401).json({ ok: false, error: 'Sesi login tidak valid atau telah kedaluwarsa.' });
  }
}
