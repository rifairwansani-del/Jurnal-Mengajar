import { useCallback, useEffect, useRef, useState } from 'react';
import { onAuthStateChanged, signInWithPopup, signOut, type User } from 'firebase/auth';
import {
  collection,
  deleteDoc,
  doc,
  getDoc,
  getDocs,
  onSnapshot,
  serverTimestamp,
  setDoc,
  updateDoc,
  writeBatch,
  type Unsubscribe,
} from 'firebase/firestore';
import { createDefaultSettings, createEmptyDatabase, DEFAULT_ADMIN_EMAIL } from '../data/seedData';
import {
  auth,
  db as firestoreDb,
  googleAuthProvider,
  handleFirestoreError,
  OperationType,
  testConnection,
} from '../lib/firebase';
import {
  GradeEntry,
  JournalEntry,
  PortalDatabase,
  ReportChartId,
  ReportPrintPreferences,
  ReportSectionId,
  SemesterType,
  SignaturePositionSettings,
  Student,
  SystemSettings,
  Teacher,
  TeachingSchedule,
  UserAccount,
} from '../types';
import {
  normalizeEmailId,
  normalizeKopSurat,
  normalizePhotoInputUrl,
  toFirestoreSafeId,
} from '../utils/portalHelpers';

const VALID_REPORT_SECTION_IDS: ReportSectionId[] = [
  'journal',
  'attendance',
  'attendance_monthly',
  'attendance_monitoring',
  'grades',
  'grades_remedial',
  'grades_leger',
  'students_roster',
  'teacher_activity',
  'charts',
];

const VALID_REPORT_CHART_IDS: ReportChartId[] = [
  'bar_class',
  'line_trend',
  'doughnut_attendance',
  'hbar_subject',
  'pie_kkm',
];

function clampStr(value: unknown, maxLen: number): string {
  if (typeof value !== 'string') return '';
  return value.length > maxLen ? value.slice(0, maxLen) : value;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== 'object') return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function stripUndefined<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map((item) => stripUndefined(item)) as unknown as T;
  }
  if (isPlainObject(value)) {
    const out: Record<string, unknown> = {};
    Object.entries(value).forEach(([k, v]) => {
      if (v !== undefined) {
        out[k] = stripUndefined(v);
      }
    });
    return out as T;
  }
  return value;
}

function stripServerMeta<T extends Record<string, unknown>>(raw: T): Omit<T, 'serverCreatedAt' | 'serverUpdatedAt'> {
  const { serverCreatedAt: _c, serverUpdatedAt: _u, ...rest } = raw;
  return rest;
}

function normalizeReportPrefs(
  raw: Partial<ReportPrintPreferences> | undefined,
  fallback: ReportPrintPreferences
): ReportPrintPreferences {
  const merged = { ...fallback, ...(raw || {}) };
  let selectedReportTypes: ReportSectionId[];
  if (Array.isArray(raw?.selectedReportTypes)) {
    selectedReportTypes = raw.selectedReportTypes.filter((id): id is ReportSectionId =>
      VALID_REPORT_SECTION_IDS.includes(id as ReportSectionId)
    );
  } else if (raw?.reportType === 'package') {
    selectedReportTypes = ['journal', 'attendance', 'grades'];
  } else if (
    raw?.reportType &&
    VALID_REPORT_SECTION_IDS.includes(raw.reportType as ReportSectionId)
  ) {
    selectedReportTypes = [raw.reportType as ReportSectionId];
  } else {
    selectedReportTypes = fallback.selectedReportTypes;
  }

  const selectedChartTypes = Array.isArray(raw?.selectedChartTypes)
    ? raw.selectedChartTypes.filter((id): id is ReportChartId =>
        VALID_REPORT_CHART_IDS.includes(id as ReportChartId)
      )
    : fallback.selectedChartTypes || VALID_REPORT_CHART_IDS;

  return {
    ...merged,
    selectedReportTypes,
    selectedChartTypes,
    showSignatureImage: raw?.showSignatureImage ?? fallback.showSignatureImage,
  };
}

function normalizeSignaturePosition(
  raw: Partial<SignaturePositionSettings> | undefined,
  fallback: SignaturePositionSettings
): SignaturePositionSettings {
  const legacySig = raw?.signatureSize;
  return {
    ...fallback,
    ...(raw || {}),
    principalSignatureSize:
      raw?.principalSignatureSize ?? legacySig ?? fallback.principalSignatureSize,
    teacherSignatureSize:
      raw?.teacherSignatureSize ?? legacySig ?? fallback.teacherSignatureSize,
    stampSize: raw?.stampSize ?? fallback.stampSize,
    principalSigOffsetX: raw?.principalSigOffsetX ?? fallback.principalSigOffsetX,
    principalSigOffsetY: raw?.principalSigOffsetY ?? fallback.principalSigOffsetY,
    teacherSigOffsetX: raw?.teacherSigOffsetX ?? fallback.teacherSigOffsetX,
    teacherSigOffsetY: raw?.teacherSigOffsetY ?? fallback.teacherSigOffsetY,
    stampOffsetX: raw?.stampOffsetX ?? fallback.stampOffsetX,
    stampOffsetY: raw?.stampOffsetY ?? fallback.stampOffsetY,
    stampTarget: 'kepsek',
  };
}

function normalizeSettings(raw: Partial<SystemSettings> | undefined): SystemSettings {
  const fallback = createDefaultSettings();
  if (!raw) return fallback;
  const cleanRaw = { ...raw } as Partial<SystemSettings> & { password?: unknown };
  delete cleanRaw.password;

  return {
    ...fallback,
    ...cleanRaw,
    academicYear: clampStr(cleanRaw.academicYear, 32) || fallback.academicYear,
    semester: cleanRaw.semester === 'Genap' ? 'Genap' : cleanRaw.semester === 'Ganjil' ? 'Ganjil' : fallback.semester,
    signatureLocation: clampStr(cleanRaw.signatureLocation ?? fallback.signatureLocation, 120),
    principalTitle: clampStr(cleanRaw.principalTitle ?? fallback.principalTitle, 120),
    principalName: clampStr(cleanRaw.principalName ?? fallback.principalName, 160),
    principalNip: clampStr(cleanRaw.principalNip ?? fallback.principalNip, 64),
    principalSignatureImage: clampStr(cleanRaw.principalSignatureImage ?? fallback.principalSignatureImage, 350000),
    defaultSchoolStampImage: clampStr(cleanRaw.defaultSchoolStampImage ?? fallback.defaultSchoolStampImage, 350000),
    teacherName: clampStr(cleanRaw.teacherName ?? fallback.teacherName, 160),
    teacherNip: clampStr(cleanRaw.teacherNip ?? fallback.teacherNip, 64),
    schoolName: clampStr(cleanRaw.schoolName ?? fallback.schoolName, 160),
    email: clampStr(cleanRaw.email ?? fallback.email, 160),
    themeId: cleanRaw.themeId || fallback.themeId,
    identity: {
      ...fallback.identity,
      ...(cleanRaw.identity || {}),
    },
    signaturePosition: normalizeSignaturePosition(
      cleanRaw.signaturePosition,
      fallback.signaturePosition
    ),
    reportPrintPreferences: normalizeReportPrefs(
      cleanRaw.reportPrintPreferences,
      fallback.reportPrintPreferences
    ),
    kopSurat: normalizeKopSurat(cleanRaw.kopSurat, fallback.kopSurat),
    masterClasses: Array.isArray(cleanRaw.masterClasses)
      ? cleanRaw.masterClasses.slice(0, 100)
      : fallback.masterClasses,
    masterSubjects: Array.isArray(cleanRaw.masterSubjects)
      ? cleanRaw.masterSubjects.slice(0, 100)
      : fallback.masterSubjects,
    masterExamTypes: Array.isArray(cleanRaw.masterExamTypes)
      ? cleanRaw.masterExamTypes.slice(0, 50)
      : fallback.masterExamTypes,
    masterAcademicYears: Array.isArray(cleanRaw.masterAcademicYears)
      ? cleanRaw.masterAcademicYears.slice(0, 30)
      : fallback.masterAcademicYears,
  };
}

function normalizeStudent(raw: Partial<Student>, fallbackIdx = 0): Student {
  const safeId = clampStr(toFirestoreSafeId(raw.id || `st-${Date.now()}-${fallbackIdx}`), 160);
  return {
    id: safeId || `st-${Date.now()}-${fallbackIdx}`,
    nis: clampStr(raw.nis, 64),
    nisn: clampStr(raw.nisn, 64),
    name: clampStr(raw.name, 160) || 'Tanpa Nama',
    className: clampStr(raw.className, 64) || 'Umum',
    absentNo: clampStr(raw.absentNo, 32),
    birthPlace: clampStr(raw.birthPlace, 120),
    birthDate: clampStr(raw.birthDate, 32),
    email: clampStr(raw.email, 160),
    photo: clampStr(normalizePhotoInputUrl(typeof raw.photo === 'string' ? raw.photo : ''), 350000),
    address: clampStr(raw.address, 400),
    phone: clampStr(raw.phone, 40),
    gender: raw.gender === 'P' ? 'P' : 'L',
    status:
      raw.status === 'Pindah' || raw.status === 'Lulus' || raw.status === 'Nonaktif'
        ? raw.status
        : 'Aktif',
    homeroomTeacher: clampStr(raw.homeroomTeacher, 160),
    counselorTeacher: clampStr(raw.counselorTeacher, 160),
    entryYear: clampStr(raw.entryYear, 32),
  };
}

function normalizeTeacher(raw: Partial<Teacher>, fallbackIdx = 0): Teacher {
  const safeId = clampStr(toFirestoreSafeId(raw.id || `tch-${Date.now()}-${fallbackIdx}`), 160);
  return {
    id: safeId || `tch-${Date.now()}-${fallbackIdx}`,
    nip: clampStr(raw.nip, 64),
    name: clampStr(raw.name, 160) || 'Tanpa Nama',
    position: clampStr(raw.position, 160),
    subject: clampStr(raw.subject, 120),
    email: clampStr(raw.email, 160),
    phone: clampStr(raw.phone, 40),
    gender: raw.gender === 'P' ? 'P' : 'L',
    employmentStatus:
      raw.employmentStatus === 'PPPK' ||
      raw.employmentStatus === 'Tetap Yayasan' ||
      raw.employmentStatus === 'Honorer'
        ? raw.employmentStatus
        : 'PNS',
    status: raw.status === 'Cuti' || raw.status === 'Nonaktif' ? raw.status : 'Aktif',
    address: clampStr(raw.address, 400),
    photo: clampStr(normalizePhotoInputUrl(typeof raw.photo === 'string' ? raw.photo : ''), 350000),
  };
}

function normalizeSchedule(raw: Partial<TeachingSchedule>, fallbackIdx = 0): TeachingSchedule {
  const safeId = clampStr(toFirestoreSafeId(raw.id || `sch-${Date.now()}-${fallbackIdx}`), 160);
  const validDays = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  return {
    id: safeId || `sch-${Date.now()}-${fallbackIdx}`,
    day: validDays.includes(raw.day || '') ? (raw.day as TeachingSchedule['day']) : 'Senin',
    jamKe: clampStr(raw.jamKe, 32),
    timeRange: clampStr(raw.timeRange, 64),
    className: clampStr(raw.className, 64) || 'Umum',
    subject: clampStr(raw.subject, 120) || 'Mapel',
    teacherName: clampStr(raw.teacherName, 160),
    room: clampStr(raw.room, 80),
    academicYear: clampStr(raw.academicYear, 32),
    semester: raw.semester === 'Genap' ? 'Genap' : 'Ganjil',
    notes: clampStr(raw.notes, 400),
  };
}

function normalizeJournal(raw: Partial<JournalEntry>, fallbackIdx = 0): JournalEntry {
  const safeId = clampStr(toFirestoreSafeId(raw.id || `jrn-${Date.now()}-${fallbackIdx}`), 160);
  return {
    id: safeId || `jrn-${Date.now()}-${fallbackIdx}`,
    date: clampStr(raw.date, 32) || new Date().toISOString().slice(0, 10),
    jamKe: clampStr(raw.jamKe, 32),
    className: clampStr(raw.className, 64) || 'Umum',
    subject: clampStr(raw.subject, 120),
    topic: clampStr(raw.topic, 300) || '-',
    activity: clampStr(raw.activity, 2000),
    academicYear: clampStr(raw.academicYear, 32),
    semester: raw.semester === 'Genap' ? 'Genap' : 'Ganjil',
    attendance: raw.attendance && typeof raw.attendance === 'object' ? raw.attendance : {},
    createdAt: clampStr(raw.createdAt, 64) || new Date().toISOString(),
  };
}

function normalizeGrade(raw: Partial<GradeEntry>, fallbackIdx = 0): GradeEntry {
  const safeId = clampStr(toFirestoreSafeId(raw.id || `grd-${Date.now()}-${fallbackIdx}`), 160);
  const validMarkers = ['neutral', 'green', 'yellow', 'red', 'blue'];
  const rawScore = typeof raw.score === 'number' && !Number.isNaN(raw.score)
    ? Math.min(100, Math.max(0, raw.score))
    : null;
  return {
    id: safeId || `grd-${Date.now()}-${fallbackIdx}`,
    className: clampStr(raw.className, 64) || 'Umum',
    subject: clampStr(raw.subject, 120) || 'Mapel',
    examType: clampStr(raw.examType, 64) || 'Tugas',
    studentId: clampStr(raw.studentId, 160) || 'unknown',
    score: rawScore,
    note: clampStr(raw.note, 500),
    marker: validMarkers.includes(raw.marker || '') ? (raw.marker as GradeEntry['marker']) : 'neutral',
    academicYear: clampStr(raw.academicYear, 32),
    semester: raw.semester === 'Genap' ? 'Genap' : 'Ganjil',
    updatedAt: clampStr(raw.updatedAt, 64) || new Date().toISOString(),
  };
}

function normalizeAccount(
  raw: Partial<UserAccount>,
  fallbackStamp: string
): UserAccount {
  const emailLower = clampStr(normalizeEmailId(raw.email || raw.id || ''), 160);
  const validRoles = ['Admin', 'Guru Mapel', 'Wali Kelas', 'Guru BK', 'Kepala Sekolah'];
  const resolvedRole: UserAccount['role'] = validRoles.includes(raw.role || '')
    ? (raw.role as UserAccount['role'])
    : 'Guru Mapel';
  const cleanUsername = clampStr(
    String(raw.username || '').trim().toLowerCase(),
    64
  );
  return {
    id: emailLower,
    name: clampStr(raw.name, 160) || emailLower,
    email: emailLower,
    username: cleanUsername,
    hasPassword: Boolean(raw.hasPassword),
    role: resolvedRole,
    nip: clampStr(raw.nip, 64),
    status: raw.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif',
    lastLogin: clampStr(raw.lastLogin, 64) || '-',
    signatureImage: clampStr(raw.signatureImage, 350000),
    schoolStampImage:
      resolvedRole === 'Kepala Sekolah'
        ? clampStr(raw.schoolStampImage || fallbackStamp, 350000)
        : '',
  };
}

interface BatchWriteOp {
  type: 'set' | 'update' | 'delete';
  collectionName: string;
  docId: string;
  data?: Record<string, unknown>;
}

async function commitBatchChunked(ops: BatchWriteOp[]): Promise<void> {
  const CHUNK_SIZE = 400;
  for (let i = 0; i < ops.length; i += CHUNK_SIZE) {
    const chunk = ops.slice(i, i + CHUNK_SIZE);
    const batch = writeBatch(firestoreDb);
    chunk.forEach((op) => {
      const ref = doc(firestoreDb, op.collectionName, op.docId);
      if (op.type === 'delete') {
        batch.delete(ref);
      } else if (op.type === 'set' && op.data) {
        batch.set(ref, stripUndefined(op.data));
      } else if (op.type === 'update' && op.data) {
        batch.update(ref, stripUndefined(op.data));
      }
    });
    await batch.commit();
  }
}

async function upsertDocSafe(
  collectionName: string,
  docId: string,
  payload: Record<string, unknown>
): Promise<void> {
  const docRef = doc(firestoreDb, collectionName, docId);
  const snap = await getDoc(docRef);
  if (snap.exists()) {
    await updateDoc(docRef, {
      ...stripUndefined(payload),
      serverUpdatedAt: serverTimestamp(),
    });
  } else {
    await setDoc(docRef, {
      ...stripUndefined(payload),
      serverCreatedAt: serverTimestamp(),
      serverUpdatedAt: serverTimestamp(),
    });
  }
}

interface CredentialSyncRecord {
  accountId: string;
  username: string;
  email: string;
  name: string;
  role: UserAccount['role'];
  nip: string;
  status: 'Aktif' | 'Nonaktif';
  passwordHash: string;
  salt: string;
  updatedAt: string;
}

export function usePortalStore() {
  const [db, setDb] = useState<PortalDatabase>(createEmptyDatabase);
  const [authLoading, setAuthLoading] = useState(true);
  const [signingIn, setSigningIn] = useState(false);
  const [dataLoading, setDataLoading] = useState(false);
  const [accessDeniedMessage, setAccessDeniedMessage] = useState<string | null>(null);
  const [pendingUserEmail, setPendingUserEmail] = useState<string | null>(null);
  const [fatalError, setFatalError] = useState<Error | null>(null);

  const dbRef = useRef<PortalDatabase>(db);
  dbRef.current = db;

  const sessionOverrideRef = useRef<{
    academicYear?: string;
    semester?: SemesterType;
    applyOverride?: boolean;
  }>({});
  const unsubsRef = useRef<Unsubscribe[]>([]);
  const authModeRef = useRef<'none' | 'google' | 'credentials'>('none');
  const localTokenRef = useRef<string | null>(null);
  const credentialsMetaRef = useRef<Record<string, { username: string; hasPassword: boolean }>>({});
  const initialServerSyncedRef = useRef(false);

  const cleanupListeners = useCallback(() => {
    unsubsRef.current.forEach((unsub) => unsub());
    unsubsRef.current = [];
  }, []);

  const reportFirestoreError = useCallback((err: unknown, op: OperationType, path: string) => {
    try {
      handleFirestoreError(err, op, path);
    } catch (wrapped) {
      setFatalError(wrapped instanceof Error ? wrapped : new Error(String(wrapped)));
    }
  }, []);

  const getActiveBearerToken = useCallback(async (): Promise<string | null> => {
    if (authModeRef.current === 'credentials') {
      return localTokenRef.current;
    }
    if (auth.currentUser) {
      return auth.currentUser.getIdToken();
    }
    return null;
  }, []);

  const pushSnapshotToServer = useCallback(
    async (snapshotDb: PortalDatabase, markFirestoreSynced = false): Promise<void> => {
      try {
        const token = await getActiveBearerToken();
        if (!token) return;
        await fetch('/api/portal/state', {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            markFirestoreSynced,
            database: {
              students: snapshotDb.students,
              teachers: snapshotDb.teachers,
              schedules: snapshotDb.schedules,
              accounts: snapshotDb.accounts,
              journals: snapshotDb.journals,
              grades: snapshotDb.grades,
              settings: snapshotDb.settings,
            },
          }),
        });
      } catch {
        // Non-fatal network hiccup
      }
    },
    [getActiveBearerToken]
  );

  useEffect(() => {
    void testConnection();
  }, []);

  useEffect(() => {
    const unsubscribeAuth = onAuthStateChanged(auth, async (firebaseUser: User | null) => {
      if (authModeRef.current === 'credentials' && localTokenRef.current && !firebaseUser) {
        return;
      }

      cleanupListeners();

      if (!firebaseUser) {
        authModeRef.current = 'none';
        initialServerSyncedRef.current = false;
        setDb(createEmptyDatabase());
        setAccessDeniedMessage(null);
        setPendingUserEmail(null);
        setDataLoading(false);
        setAuthLoading(false);
        return;
      }

      authModeRef.current = 'google';
      localTokenRef.current = null;
      initialServerSyncedRef.current = false;

      const rawEmail = firebaseUser.email || '';
      const emailLower = normalizeEmailId(rawEmail);

      if (!emailLower || !firebaseUser.emailVerified) {
        setPendingUserEmail(emailLower || null);
        setAccessDeniedMessage(
          'Akun Google Anda tidak memiliki alamat email yang terverifikasi.'
        );
        setDataLoading(false);
        setAuthLoading(false);
        return;
      }

      setDataLoading(true);
      try {
        // Verify Firebase ID token with Express backend
        try {
          const idToken = await firebaseUser.getIdToken();
          await fetch('/api/auth/verify', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              Authorization: `Bearer ${idToken}`,
            },
            body: JSON.stringify({
              academicYear: sessionOverrideRef.current.academicYear || '',
              semester: sessionOverrideRef.current.semester || 'Ganjil',
            }),
          });
        } catch {
          // Non-fatal if preview backend restarts briefly; Firestore security rules remain enforced
        }

        const isBootstrapAdmin = emailLower === DEFAULT_ADMIN_EMAIL;
        const userDocRef = doc(firestoreDb, 'users', firebaseUser.uid);

        try {
          const userSnap = await getDoc(userDocRef);
          if (!userSnap.exists()) {
            await setDoc(userDocRef, {
              uid: firebaseUser.uid,
              email: emailLower,
              displayName: firebaseUser.displayName || emailLower,
              role: isBootstrapAdmin ? 'Admin' : 'Pending',
              createdAt: serverTimestamp(),
              lastLoginAt: serverTimestamp(),
            });
          } else {
            const existingData = userSnap.data();
            await updateDoc(userDocRef, {
              uid: firebaseUser.uid,
              email: emailLower,
              displayName: firebaseUser.displayName || existingData.displayName || emailLower,
              role: isBootstrapAdmin ? 'Admin' : existingData.role || 'Pending',
              lastLoginAt: serverTimestamp(),
            });
          }
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, `users/${firebaseUser.uid}`);
          setDataLoading(false);
          setAuthLoading(false);
          return;
        }

        const accountDocRef = doc(firestoreDb, 'accounts', emailLower);
        let accountRecord: UserAccount | null = null;

        const ALLOWED_ACCOUNT_KEYS = new Set([
          'id',
          'name',
          'email',
          'username',
          'hasPassword',
          'role',
          'nip',
          'status',
          'lastLogin',
          'signatureImage',
          'schoolStampImage',
          'serverCreatedAt',
          'serverUpdatedAt',
        ]);

        let existingRawAccountData: Record<string, unknown> | null = null;
        try {
          const accSnap = await getDoc(accountDocRef);
          if (accSnap.exists()) {
            existingRawAccountData = accSnap.data() as Record<string, unknown>;
            accountRecord = normalizeAccount(existingRawAccountData as Partial<UserAccount>, '');
          }
        } catch (err) {
          reportFirestoreError(err, OperationType.GET, `accounts/${emailLower}`);
          setDataLoading(false);
          setAuthLoading(false);
          return;
        }

        if (!isBootstrapAdmin && (!accountRecord || accountRecord.status !== 'Aktif')) {
          setPendingUserEmail(emailLower);
          setAccessDeniedMessage(
            accountRecord?.status === 'Nonaktif'
              ? 'Akun Anda saat ini berstatus Nonaktif. Hubungi Admin sekolah untuk mengaktifkan kembali.'
              : `Email Google Anda belum terdaftar pada daftar akun sekolah. Hubungi Admin (${DEFAULT_ADMIN_EMAIL}) untuk mendaftarkan email ini di menu Manajemen Akun.`
          );
          setDb((prev) => ({
            ...prev,
            session: { ...prev.session, isLoggedIn: false },
          }));
          setDataLoading(false);
          setAuthLoading(false);
          return;
        }

        setAccessDeniedMessage(null);
        setPendingUserEmail(null);

        const nowStr = new Date().toISOString().slice(0, 16).replace('T', ' ');
        const defaultSettings = createDefaultSettings();

        if (!accountRecord && isBootstrapAdmin) {
          const initialAdminAccount: UserAccount = {
            id: emailLower,
            name: firebaseUser.displayName || defaultSettings.teacherName,
            email: emailLower,
            username: '',
            hasPassword: false,
            role: 'Admin',
            nip: defaultSettings.teacherNip,
            status: 'Aktif',
            lastLogin: nowStr,
            signatureImage: '',
            schoolStampImage: '',
          };
          try {
            await setDoc(accountDocRef, {
              ...initialAdminAccount,
              serverCreatedAt: serverTimestamp(),
              serverUpdatedAt: serverTimestamp(),
            });
            accountRecord = initialAdminAccount;
          } catch (err) {
            reportFirestoreError(err, OperationType.CREATE, `accounts/${emailLower}`);
          }
        } else if (accountRecord) {
          try {
            const hasUnexpectedKeys =
              existingRawAccountData !== null &&
              Object.keys(existingRawAccountData).some((k) => !ALLOWED_ACCOUNT_KEYS.has(k));

            if (isBootstrapAdmin && hasUnexpectedKeys) {
              // Cleanly repair any legacy fields on the admin account document
              const repairedAccount: UserAccount = {
                ...accountRecord,
                id: emailLower,
                email: emailLower,
                role: 'Admin',
                status: 'Aktif',
                lastLogin: nowStr,
              };
              const payload: Record<string, unknown> = {
                ...repairedAccount,
                serverUpdatedAt: serverTimestamp(),
              };
              if (existingRawAccountData?.serverCreatedAt !== undefined) {
                payload.serverCreatedAt = existingRawAccountData.serverCreatedAt;
              }
              await setDoc(accountDocRef, payload);
            } else {
              await updateDoc(accountDocRef, {
                lastLogin: nowStr,
                serverUpdatedAt: serverTimestamp(),
              });
            }
          } catch (err) {
            reportFirestoreError(err, OperationType.UPDATE, `accounts/${emailLower}`);
          }
        }

        const resolvedRole = isBootstrapAdmin
          ? 'Admin'
          : accountRecord?.role || 'Guru Mapel';
        const isAdminUser = isBootstrapAdmin || resolvedRole === 'Admin';

        // Sync Admin credentials bi-directionally between Cloud Firestore (/credential_sync) and this Express server
        if (isAdminUser) {
          try {
            const idToken = await firebaseUser.getIdToken();
            const cloudCredSnap = await getDocs(collection(firestoreDb, 'credential_sync'));
            const cloudMap = new Map<string, CredentialSyncRecord>();
            cloudCredSnap.docs.forEach((d) => {
              const raw = stripServerMeta(d.data() as Record<string, unknown>) as unknown as CredentialSyncRecord;
              if (raw && raw.accountId && raw.passwordHash && raw.salt) {
                cloudMap.set(normalizeEmailId(raw.accountId), raw);
              }
            });

            const syncRes = await fetch('/api/accounts/credentials-sync', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                Authorization: `Bearer ${idToken}`,
              },
              body: JSON.stringify({
                credentials: Array.from(cloudMap.values()),
              }),
            });

            if (syncRes.ok) {
              const syncJson = (await syncRes.json()) as {
                ok?: boolean;
                credentials?: CredentialSyncRecord[];
              };
              const mergedList = Array.isArray(syncJson.credentials) ? syncJson.credentials : [];
              const metaOut: Record<string, { username: string; hasPassword: boolean }> = {};
              const toPushToCloud: CredentialSyncRecord[] = [];

              mergedList.forEach((rec) => {
                const id = normalizeEmailId(rec.accountId || rec.email);
                if (!id) return;
                metaOut[id] = {
                  username: rec.username || '',
                  hasPassword: Boolean(rec.passwordHash && rec.salt),
                };
                const existingCloud = cloudMap.get(id);
                const localTime = Date.parse(rec.updatedAt || '') || 0;
                const cloudTime = existingCloud ? Date.parse(existingCloud.updatedAt || '') || 0 : -1;
                if (!existingCloud || localTime > cloudTime) {
                  toPushToCloud.push(rec);
                }
              });

              credentialsMetaRef.current = {
                ...credentialsMetaRef.current,
                ...metaOut,
              };

              if (toPushToCloud.length > 0) {
                const ops: BatchWriteOp[] = toPushToCloud.map((rec) => {
                  const id = normalizeEmailId(rec.accountId || rec.email);
                  const exists = cloudMap.has(id);
                  return {
                    type: exists ? 'update' : 'set',
                    collectionName: 'credential_sync',
                    docId: id,
                    data: exists
                      ? { ...rec, serverUpdatedAt: serverTimestamp() }
                      : {
                          ...rec,
                          serverCreatedAt: serverTimestamp(),
                          serverUpdatedAt: serverTimestamp(),
                        },
                  };
                });
                await commitBatchChunked(ops);
              }
            }
          } catch {
            // Non-fatal; fallback to metadata endpoint
          }
        }

        // Load credential metadata (username & hasPassword) from Express server
        try {
          const idToken = await firebaseUser.getIdToken();
          const metaRes = await fetch('/api/accounts/credentials-meta', {
            headers: { Authorization: `Bearer ${idToken}` },
          });
          if (metaRes.ok) {
            const metaJson = (await metaRes.json()) as {
              ok?: boolean;
              items?: Record<string, { username: string; hasPassword: boolean }>;
            };
            if (metaJson.ok && metaJson.items) {
              credentialsMetaRef.current = {
                ...credentialsMetaRef.current,
                ...metaJson.items,
              };
            }
          }
        } catch {
          // Non-fatal
        }

        const settingsDocRef = doc(firestoreDb, 'settings', 'global');
        let resolvedSettings: SystemSettings = defaultSettings;
        try {
          const settingsSnap = await getDoc(settingsDocRef);
          const wantsOverride = Boolean(sessionOverrideRef.current.applyOverride);
          if (!settingsSnap.exists()) {
            const initialSettings = normalizeSettings({
              ...defaultSettings,
              ...(wantsOverride && sessionOverrideRef.current.academicYear
                ? { academicYear: sessionOverrideRef.current.academicYear }
                : {}),
              ...(wantsOverride && sessionOverrideRef.current.semester
                ? { semester: sessionOverrideRef.current.semester }
                : {}),
            });
            resolvedSettings = initialSettings;
            await setDoc(settingsDocRef, {
              ...stripUndefined(initialSettings),
              serverCreatedAt: serverTimestamp(),
              serverUpdatedAt: serverTimestamp(),
            });
          } else {
            const savedSettings = normalizeSettings(
              stripServerMeta(settingsSnap.data() as Record<string, unknown>) as Partial<SystemSettings>
            );
            resolvedSettings = savedSettings;
            if (
              wantsOverride &&
              (sessionOverrideRef.current.academicYear || sessionOverrideRef.current.semester)
            ) {
              const nextYear = sessionOverrideRef.current.academicYear || savedSettings.academicYear;
              const nextSem = sessionOverrideRef.current.semester || savedSettings.semester;
              if (nextYear !== savedSettings.academicYear || nextSem !== savedSettings.semester) {
                resolvedSettings = {
                  ...savedSettings,
                  academicYear: nextYear,
                  semester: nextSem,
                };
                await updateDoc(settingsDocRef, {
                  academicYear: nextYear,
                  semester: nextSem,
                  serverUpdatedAt: serverTimestamp(),
                });
              }
            }
          }
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'settings/global');
        }

        const resolvedName =
          accountRecord?.name || firebaseUser.displayName || resolvedSettings.teacherName || defaultSettings.teacherName;
        const ownMeta = credentialsMetaRef.current[emailLower];

        setDb((prev) => ({
          ...prev,
          settings: resolvedSettings,
          session: {
            isLoggedIn: true,
            authMethod: 'google',
            accountId: emailLower,
            username: ownMeta?.username || accountRecord?.username || '',
            email: emailLower,
            userName: resolvedName,
            userRole: resolvedRole,
            academicYear: resolvedSettings.academicYear,
            semester: resolvedSettings.semester,
          },
        }));

        const makeInitialGate = () => {
          let settled = false;
          let resolveFn!: () => void;
          const promise = new Promise<void>((res) => {
            resolveFn = res;
          });
          const mark = () => {
            if (!settled) {
              settled = true;
              resolveFn();
            }
          };
          return { promise, mark };
        };

        const gSettings = makeInitialGate();
        const gAccounts = makeInitialGate();
        const gStudents = makeInitialGate();
        const gTeachers = makeInitialGate();
        const gSchedules = makeInitialGate();
        const gJournals = makeInitialGate();
        const gGrades = makeInitialGate();

        const unsubSettings = onSnapshot(
          settingsDocRef,
          (snapshot) => {
            if (snapshot.exists()) {
              const normalized = normalizeSettings(
                stripServerMeta(snapshot.data() as Record<string, unknown>) as Partial<SystemSettings>
              );
              setDb((prev) => ({
                ...prev,
                settings: normalized,
                session: {
                  ...prev.session,
                  academicYear: normalized.academicYear,
                  semester: normalized.semester,
                },
              }));
            }
            gSettings.mark();
          },
          (error) => {
            gSettings.mark();
            reportFirestoreError(error, OperationType.GET, 'settings/global');
          }
        );

        const unsubAccounts = onSnapshot(
          collection(firestoreDb, 'accounts'),
          (snapshot) => {
            const list = snapshot.docs.map((d) => {
              const acc = normalizeAccount(
                stripServerMeta(d.data() as Record<string, unknown>) as Partial<UserAccount>,
                dbRef.current.settings.defaultSchoolStampImage
              );
              const meta = credentialsMetaRef.current[acc.id];
              return meta
                ? {
                    ...acc,
                    username: meta.username || acc.username || '',
                    hasPassword: meta.hasPassword || Boolean(acc.hasPassword),
                  }
                : acc;
            });
            const currentAcc = list.find(
              (a) => normalizeEmailId(a.email) === emailLower
            );
            setDb((prev) => ({
              ...prev,
              accounts: list,
              session: currentAcc
                ? {
                    ...prev.session,
                    accountId: currentAcc.id,
                    username: currentAcc.username || prev.session.username || '',
                    userName: currentAcc.name,
                    userRole: isBootstrapAdmin ? 'Admin' : currentAcc.role,
                  }
                : prev.session,
            }));
            gAccounts.mark();
          },
          (error) => {
            gAccounts.mark();
            reportFirestoreError(error, OperationType.LIST, 'accounts');
          }
        );

        const unsubStudents = onSnapshot(
          collection(firestoreDb, 'students'),
          (snapshot) => {
            const list = snapshot.docs.map((d, idx) =>
              normalizeStudent(
                stripServerMeta(d.data() as Record<string, unknown>) as Partial<Student>,
                idx
              )
            );
            setDb((prev) => ({ ...prev, students: list }));
            gStudents.mark();
          },
          (error) => {
            gStudents.mark();
            reportFirestoreError(error, OperationType.LIST, 'students');
          }
        );

        const unsubTeachers = onSnapshot(
          collection(firestoreDb, 'teachers'),
          (snapshot) => {
            const list = snapshot.docs.map((d, idx) =>
              normalizeTeacher(
                stripServerMeta(d.data() as Record<string, unknown>) as Partial<Teacher>,
                idx
              )
            );
            setDb((prev) => ({ ...prev, teachers: list }));
            gTeachers.mark();
          },
          (error) => {
            gTeachers.mark();
            reportFirestoreError(error, OperationType.LIST, 'teachers');
          }
        );

        const unsubSchedules = onSnapshot(
          collection(firestoreDb, 'schedules'),
          (snapshot) => {
            const list = snapshot.docs.map((d, idx) =>
              normalizeSchedule(
                stripServerMeta(d.data() as Record<string, unknown>) as Partial<TeachingSchedule>,
                idx
              )
            );
            setDb((prev) => ({ ...prev, schedules: list }));
            gSchedules.mark();
          },
          (error) => {
            gSchedules.mark();
            reportFirestoreError(error, OperationType.LIST, 'schedules');
          }
        );

        const unsubJournals = onSnapshot(
          collection(firestoreDb, 'journals'),
          (snapshot) => {
            const list = snapshot.docs
              .map((d, idx) =>
                normalizeJournal(
                  stripServerMeta(d.data() as Record<string, unknown>) as Partial<JournalEntry>,
                  idx
                )
              )
              .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
            setDb((prev) => ({ ...prev, journals: list }));
            gJournals.mark();
          },
          (error) => {
            gJournals.mark();
            reportFirestoreError(error, OperationType.LIST, 'journals');
          }
        );

        const unsubGrades = onSnapshot(
          collection(firestoreDb, 'grades'),
          (snapshot) => {
            const list = snapshot.docs.map((d, idx) =>
              normalizeGrade(
                stripServerMeta(d.data() as Record<string, unknown>) as Partial<GradeEntry>,
                idx
              )
            );
            setDb((prev) => ({ ...prev, grades: list }));
            gGrades.mark();
          },
          (error) => {
            gGrades.mark();
            reportFirestoreError(error, OperationType.LIST, 'grades');
          }
        );

        unsubsRef.current = [
          unsubSettings,
          unsubAccounts,
          unsubStudents,
          unsubTeachers,
          unsubSchedules,
          unsubJournals,
          unsubGrades,
        ];

        await Promise.race([
          Promise.all([
            gSettings.promise,
            gAccounts.promise,
            gStudents.promise,
            gTeachers.promise,
            gSchedules.promise,
            gJournals.promise,
            gGrades.promise,
          ]),
          new Promise<void>((resolve) => window.setTimeout(resolve, 6000)),
        ]);
      } finally {
        setDataLoading(false);
        setAuthLoading(false);
      }
    });

    return () => {
      unsubscribeAuth();
      cleanupListeners();
    };
  }, [cleanupListeners, reportFirestoreError]);

  // Sync live Firestore snapshot to Express server; promote un-synced local credential edits into Firestore first
  useEffect(() => {
    if (!db.session.isLoggedIn || db.session.authMethod !== 'google' || dataLoading) {
      return;
    }
    const timer = window.setTimeout(() => {
      void (async () => {
        if (!initialServerSyncedRef.current) {
          initialServerSyncedRef.current = true;
          try {
            const token = await getActiveBearerToken();
            if (token) {
              const stateRes = await fetch('/api/portal/state', {
                headers: { Authorization: `Bearer ${token}` },
              });
              if (stateRes.ok) {
                const stateJson = (await stateRes.json()) as {
                  ok?: boolean;
                  database?: {
                    students?: Student[];
                    teachers?: Teacher[];
                    schedules?: TeachingSchedule[];
                    journals?: JournalEntry[];
                    grades?: GradeEntry[];
                  } | null;
                  meta?: { lastLocalEditAt?: number; lastFirestoreSyncAt?: number };
                };
                const lastLocal = Number(stateJson.meta?.lastLocalEditAt) || 0;
                const lastCloud = Number(stateJson.meta?.lastFirestoreSyncAt) || 0;
                if (lastLocal > lastCloud && stateJson.database) {
                  const localDb = stateJson.database;
                  const ops: BatchWriteOp[] = [];
                  const cur = dbRef.current;
                  const studentMap = new Map(cur.students.map((s) => [s.id, s]));
                  (localDb.students || []).forEach((raw, idx) => {
                    const clean = normalizeStudent(raw, idx);
                    if (!studentMap.has(clean.id)) {
                      ops.push({
                        type: 'set',
                        collectionName: 'students',
                        docId: clean.id,
                        data: { ...clean, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
                      });
                    }
                  });
                  const teacherMap = new Map(cur.teachers.map((t) => [t.id, t]));
                  (localDb.teachers || []).forEach((raw, idx) => {
                    const clean = normalizeTeacher(raw, idx);
                    if (!teacherMap.has(clean.id)) {
                      ops.push({
                        type: 'set',
                        collectionName: 'teachers',
                        docId: clean.id,
                        data: { ...clean, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
                      });
                    }
                  });
                  const schMap = new Map(cur.schedules.map((s) => [s.id, s]));
                  (localDb.schedules || []).forEach((raw, idx) => {
                    const clean = normalizeSchedule(raw, idx);
                    if (!schMap.has(clean.id)) {
                      ops.push({
                        type: 'set',
                        collectionName: 'schedules',
                        docId: clean.id,
                        data: { ...clean, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
                      });
                    }
                  });
                  const jrnMap = new Map(cur.journals.map((j) => [j.id, j]));
                  (localDb.journals || []).forEach((raw, idx) => {
                    const clean = normalizeJournal(raw, idx);
                    if (!jrnMap.has(clean.id)) {
                      ops.push({
                        type: 'set',
                        collectionName: 'journals',
                        docId: clean.id,
                        data: { ...clean, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
                      });
                    }
                  });
                  const grdMap = new Map(cur.grades.map((g) => [g.id, g]));
                  (localDb.grades || []).forEach((raw, idx) => {
                    const clean = normalizeGrade(raw, idx);
                    const existing = grdMap.get(clean.id);
                    if (!existing || clean.updatedAt > existing.updatedAt) {
                      ops.push({
                        type: existing ? 'update' : 'set',
                        collectionName: 'grades',
                        docId: clean.id,
                        data: existing
                          ? { ...clean, serverUpdatedAt: serverTimestamp() }
                          : { ...clean, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
                      });
                    }
                  });
                  if (ops.length > 0) {
                    await commitBatchChunked(ops);
                  }
                }
              }
            }
          } catch {
            // Non-fatal
          }
        }
        await pushSnapshotToServer(dbRef.current, true);
      })();
    }, 550);
    return () => window.clearTimeout(timer);
  }, [
    db.session.isLoggedIn,
    db.session.authMethod,
    db.students,
    db.teachers,
    db.schedules,
    db.accounts,
    db.journals,
    db.grades,
    db.settings,
    dataLoading,
    getActiveBearerToken,
    pushSnapshotToServer,
  ]);

  const loginWithGoogle = useCallback(
    async (academicYear: string, semester: SemesterType, applyPeriodOverride = false) => {
      sessionOverrideRef.current = {
        academicYear,
        semester,
        applyOverride: Boolean(applyPeriodOverride),
      };
      setSigningIn(true);
      try {
        await signInWithPopup(auth, googleAuthProvider);
      } finally {
        setSigningIn(false);
      }
    },
    []
  );

  const loginWithCredentials = useCallback(
    async (
      username: string,
      password: string,
      academicYear: string,
      semester: SemesterType,
      applyPeriodOverride = false
    ) => {
      setSigningIn(true);
      try {
        const res = await fetch('/api/auth/login-credentials', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, password, academicYear, semester }),
        });
        const json = (await res.json()) as {
          ok?: boolean;
          error?: string;
          token?: string;
          user?: {
            accountId: string;
            username: string;
            email: string;
            name: string;
            role: UserAccount['role'];
            nip: string;
          };
          snapshot?: {
            database?: {
              students?: Student[];
              teachers?: Teacher[];
              schedules?: TeachingSchedule[];
              accounts?: UserAccount[];
              journals?: JournalEntry[];
              grades?: GradeEntry[];
              settings?: SystemSettings;
            } | null;
          };
        };

        if (!res.ok || !json.ok || !json.token || !json.user) {
          throw new Error(json.error || 'Gagal masuk dengan username dan password.');
        }

        cleanupListeners();
        authModeRef.current = 'credentials';
        localTokenRef.current = json.token;
        setAccessDeniedMessage(null);
        setPendingUserEmail(null);

        const snapDb = json.snapshot?.database;
        const baseSettings = snapDb?.settings || createDefaultSettings();
        const mergedSettings = normalizeSettings({
          ...baseSettings,
          academicYear:
            applyPeriodOverride && academicYear
              ? academicYear
              : baseSettings.academicYear || academicYear || '2025/2026',
          semester:
            applyPeriodOverride && semester
              ? semester
              : baseSettings.semester || semester || 'Ganjil',
        });

        const accountsList = Array.isArray(snapDb?.accounts)
          ? snapDb.accounts.map((a) => normalizeAccount(a, mergedSettings.defaultSchoolStampImage))
          : [];

        const hasSelf = accountsList.some((a) => a.id === json.user!.accountId);
        const finalAccounts = hasSelf
          ? accountsList
          : [
              ...accountsList,
              normalizeAccount(
                {
                  id: json.user.accountId,
                  name: json.user.name,
                  email: json.user.email,
                  username: json.user.username,
                  hasPassword: true,
                  role: json.user.role,
                  nip: json.user.nip,
                  status: 'Aktif',
                  lastLogin: new Date().toISOString().slice(0, 16).replace('T', ' '),
                },
                mergedSettings.defaultSchoolStampImage
              ),
            ];

        const nextDb: PortalDatabase = {
          students: Array.isArray(snapDb?.students)
            ? snapDb.students.map((s, idx) => normalizeStudent(s, idx))
            : [],
          teachers: Array.isArray(snapDb?.teachers)
            ? snapDb.teachers.map((t, idx) => normalizeTeacher(t, idx))
            : [],
          schedules: Array.isArray(snapDb?.schedules)
            ? snapDb.schedules.map((s, idx) => normalizeSchedule(s, idx))
            : [],
          accounts: finalAccounts,
          journals: Array.isArray(snapDb?.journals)
            ? snapDb.journals.map((j, idx) => normalizeJournal(j, idx))
            : [],
          grades: Array.isArray(snapDb?.grades)
            ? snapDb.grades.map((g, idx) => normalizeGrade(g, idx))
            : [],
          settings: mergedSettings,
          session: {
            isLoggedIn: true,
            authMethod: 'credentials',
            accountId: json.user.accountId,
            username: json.user.username,
            email: json.user.email,
            userName: json.user.name,
            userRole: json.user.role,
            academicYear: mergedSettings.academicYear,
            semester: mergedSettings.semester,
          },
        };

        setDb(nextDb);
        setDataLoading(false);
        setAuthLoading(false);
      } finally {
        setSigningIn(false);
      }
    },
    [cleanupListeners]
  );

  const logout = useCallback(async () => {
    cleanupListeners();
    sessionOverrideRef.current = {};
    localTokenRef.current = null;
    authModeRef.current = 'none';
    initialServerSyncedRef.current = false;
    setAccessDeniedMessage(null);
    setPendingUserEmail(null);
    if (auth.currentUser) {
      await signOut(auth);
    } else {
      setDb(createEmptyDatabase());
    }
  }, [cleanupListeners]);

  // Journals
  const upsertJournal = useCallback(
    (entry: JournalEntry) => {
      const safeId = toFirestoreSafeId(entry.id);
      const cleanEntry = normalizeJournal({ ...entry, id: safeId });
      const existsLocally = dbRef.current.journals.some((j) => j.id === safeId);

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        journals: existsLocally
          ? dbRef.current.journals.map((j) => (j.id === safeId ? cleanEntry : j))
          : [cleanEntry, ...dbRef.current.journals],
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void upsertDocSafe('journals', safeId, cleanEntry as unknown as Record<string, unknown>).catch(
        (err) =>
          reportFirestoreError(
            err,
            existsLocally ? OperationType.UPDATE : OperationType.CREATE,
            `journals/${safeId}`
          )
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const importJournalsBatch = useCallback(
    (incoming: Omit<JournalEntry, 'id' | 'createdAt'>[]) => {
      const createdEntries: JournalEntry[] = [];
      const newClasses = new Set(dbRef.current.settings.masterClasses);
      const newSubjects = new Set(dbRef.current.settings.masterSubjects);

      incoming.forEach((item, idx) => {
        if (!item.topic.trim() || !item.className.trim()) return;
        newClasses.add(item.className);
        if (item.subject) newSubjects.add(item.subject);
        const id = `jrn-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`;
        const entry = normalizeJournal(
          {
            ...item,
            id,
            createdAt: new Date().toISOString(),
          },
          idx
        );
        createdEntries.push(entry);
      });

      if (createdEntries.length === 0) return 0;

      const updatedClasses = Array.from(newClasses).slice(0, 100);
      const updatedSubjects = Array.from(newSubjects).slice(0, 100);

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        journals: [...createdEntries.reverse(), ...dbRef.current.journals],
        settings: {
          ...dbRef.current.settings,
          masterClasses: updatedClasses,
          masterSubjects: updatedSubjects,
        },
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return createdEntries.length;
      }

      void (async () => {
        try {
          const ops: BatchWriteOp[] = createdEntries.map((entry) => ({
            type: 'set',
            collectionName: 'journals',
            docId: entry.id,
            data: {
              ...entry,
              serverCreatedAt: serverTimestamp(),
              serverUpdatedAt: serverTimestamp(),
            },
          }));
          const settingsSnap = await getDoc(doc(firestoreDb, 'settings', 'global'));
          if (settingsSnap.exists()) {
            ops.push({
              type: 'update',
              collectionName: 'settings',
              docId: 'global',
              data: {
                masterClasses: updatedClasses,
                masterSubjects: updatedSubjects,
                serverUpdatedAt: serverTimestamp(),
              },
            });
          }
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'journals');
        }
      })();

      return createdEntries.length;
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteJournals = useCallback(
    (ids: string[]) => {
      if (ids.length === 0) return;
      const idSet = new Set(ids);
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        journals: dbRef.current.journals.filter((j) => !idSet.has(j.id)),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = ids.map((id) => ({
        type: 'delete',
        collectionName: 'journals',
        docId: id,
      }));
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, 'journals')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const batchUpdateJournals = useCallback(
    (updatedJournals: JournalEntry[]) => {
      if (updatedJournals.length === 0) return;
      const map = new Map(updatedJournals.map((j) => [j.id, j]));
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        journals: dbRef.current.journals.map((j) => map.get(j.id) ?? j),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = updatedJournals.map((j) => ({
        type: 'update',
        collectionName: 'journals',
        docId: j.id,
        data: {
          ...j,
          serverUpdatedAt: serverTimestamp(),
        },
      }));
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.UPDATE, 'journals')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  // Schedules
  const upsertSchedule = useCallback(
    (item: TeachingSchedule) => {
      const safeId = toFirestoreSafeId(item.id);
      const cleanItem = normalizeSchedule({ ...item, id: safeId });
      const existsLocally = dbRef.current.schedules.some((s) => s.id === safeId);

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        schedules: existsLocally
          ? dbRef.current.schedules.map((s) => (s.id === safeId ? cleanItem : s))
          : [...dbRef.current.schedules, cleanItem],
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void upsertDocSafe('schedules', safeId, cleanItem as unknown as Record<string, unknown>).catch(
        (err) =>
          reportFirestoreError(
            err,
            existsLocally ? OperationType.UPDATE : OperationType.CREATE,
            `schedules/${safeId}`
          )
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const importSchedulesBatch = useCallback(
    (incoming: Omit<TeachingSchedule, 'id'>[]) => {
      const ops: BatchWriteOp[] = [];
      const created: TeachingSchedule[] = [];
      incoming.forEach((item, idx) => {
        if (!item.className.trim() || !item.subject.trim()) return;
        const id = `sch-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`;
        const record: TeachingSchedule = { ...item, id };
        created.push(record);
        ops.push({
          type: 'set',
          collectionName: 'schedules',
          docId: id,
          data: {
            ...record,
            serverCreatedAt: serverTimestamp(),
            serverUpdatedAt: serverTimestamp(),
          },
        });
      });

      if (ops.length === 0) return 0;

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        schedules: [...dbRef.current.schedules, ...created],
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return created.length;
      }

      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.WRITE, 'schedules')
      );
      return created.length;
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteSchedules = useCallback(
    (ids: string[]) => {
      if (ids.length === 0) return;
      const idSet = new Set(ids);
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        schedules: dbRef.current.schedules.filter((s) => !idSet.has(s.id)),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = ids.map((id) => ({
        type: 'delete',
        collectionName: 'schedules',
        docId: id,
      }));
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, 'schedules')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteSchedule = useCallback(
    (id: string) => {
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        schedules: dbRef.current.schedules.filter((s) => s.id !== id),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void deleteDoc(doc(firestoreDb, 'schedules', id)).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, `schedules/${id}`)
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  // Grades
  const upsertGradesBatch = useCallback(
    (entries: GradeEntry[]) => {
      if (entries.length === 0) return;
      const sanitizedEntries = entries.map((entry, idx) => {
        const safeId = toFirestoreSafeId(entry.id);
        return normalizeGrade({ ...entry, id: safeId }, idx);
      });

      const map = new Map(dbRef.current.grades.map((g) => [g.id, g]));
      sanitizedEntries.forEach((entry) => map.set(entry.id, entry));
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        grades: Array.from(map.values()),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void (async () => {
        try {
          const existingSnap = await getDocs(collection(firestoreDb, 'grades'));
          const existingIds = new Set(existingSnap.docs.map((d) => d.id));
          const ops: BatchWriteOp[] = sanitizedEntries.map((cleanEntry) => {
            const isExisting = existingIds.has(cleanEntry.id);
            return {
              type: isExisting ? 'update' : 'set',
              collectionName: 'grades',
              docId: cleanEntry.id,
              data: isExisting
                ? { ...cleanEntry, serverUpdatedAt: serverTimestamp() }
                : {
                    ...cleanEntry,
                    serverCreatedAt: serverTimestamp(),
                    serverUpdatedAt: serverTimestamp(),
                  },
            };
          });
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'grades');
        }
      })();
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteGradesForStudents = useCallback(
    (studentIds: string[], className: string, subject: string) => {
      const stSet = new Set(studentIds);
      const { academicYear, semester } = dbRef.current.session;
      const toDelete = dbRef.current.grades.filter(
        (g) =>
          stSet.has(g.studentId) &&
          g.className === className &&
          g.subject === subject &&
          g.academicYear === academicYear &&
          g.semester === semester
      );
      if (toDelete.length === 0) return;

      const deleteIdSet = new Set(toDelete.map((g) => g.id));
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        grades: dbRef.current.grades.filter((g) => !deleteIdSet.has(g.id)),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = toDelete.map((g) => ({
        type: 'delete',
        collectionName: 'grades',
        docId: g.id,
      }));
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, 'grades')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  // Students
  const upsertStudent = useCallback(
    (student: Student) => {
      const safeId = toFirestoreSafeId(student.id);
      const cleanStudent = normalizeStudent({ ...student, id: safeId });
      const existsLocally = dbRef.current.students.some((s) => s.id === safeId);
      const needsClassAdd =
        Boolean(cleanStudent.className) &&
        !dbRef.current.settings.masterClasses.includes(cleanStudent.className);
      const nextClasses = needsClassAdd
        ? [...dbRef.current.settings.masterClasses, cleanStudent.className].slice(0, 100)
        : dbRef.current.settings.masterClasses;

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        students: existsLocally
          ? dbRef.current.students.map((s) => (s.id === safeId ? cleanStudent : s))
          : [...dbRef.current.students, cleanStudent],
        settings: { ...dbRef.current.settings, masterClasses: nextClasses },
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void (async () => {
        try {
          const studentSnap = await getDoc(doc(firestoreDb, 'students', safeId));
          const ops: BatchWriteOp[] = [
            {
              type: studentSnap.exists() ? 'update' : 'set',
              collectionName: 'students',
              docId: safeId,
              data: studentSnap.exists()
                ? { ...cleanStudent, serverUpdatedAt: serverTimestamp() }
                : {
                    ...cleanStudent,
                    serverCreatedAt: serverTimestamp(),
                    serverUpdatedAt: serverTimestamp(),
                  },
            },
          ];
          if (needsClassAdd) {
            const settingsSnap = await getDoc(doc(firestoreDb, 'settings', 'global'));
            if (settingsSnap.exists()) {
              ops.push({
                type: 'update',
                collectionName: 'settings',
                docId: 'global',
                data: { masterClasses: nextClasses, serverUpdatedAt: serverTimestamp() },
              });
            }
          }
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(
            err,
            existsLocally ? OperationType.UPDATE : OperationType.CREATE,
            `students/${safeId}`
          );
        }
      })();
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const importStudentsBatch = useCallback(
    (incoming: Omit<Student, 'id'>[]) => {
      let added = 0;
      let updated = 0;
      const nextStudents = [...dbRef.current.students];
      const newClasses = new Set(dbRef.current.settings.masterClasses);
      const recordsToSave: Student[] = [];

      incoming.forEach((item, index) => {
        const cleanNis = item.nis.trim();
        const cleanName = item.name.trim();
        const cleanClass = item.className.trim();
        if (!cleanName || !cleanClass) return;

        newClasses.add(cleanClass);
        const existingIdx = cleanNis
          ? nextStudents.findIndex((s) => s.nis.toLowerCase() === cleanNis.toLowerCase())
          : -1;

        if (existingIdx >= 0) {
          const prevStudent = nextStudents[existingIdx];
          const updatedStudent = normalizeStudent(
            {
              ...prevStudent,
              ...item,
              nis: cleanNis,
              name: cleanName,
              className: cleanClass,
              photo: item.photo || prevStudent.photo,
            },
            index
          );
          nextStudents[existingIdx] = updatedStudent;
          recordsToSave.push(updatedStudent);
          updated += 1;
        } else {
          const id = `st-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 6)}`;
          const newStudent = normalizeStudent(
            {
              ...item,
              id,
              nis: cleanNis || `${Date.now().toString().slice(-4)}${index}`,
              name: cleanName,
              className: cleanClass,
            },
            index
          );
          nextStudents.push(newStudent);
          recordsToSave.push(newStudent);
          added += 1;
        }
      });

      if (recordsToSave.length === 0) return { added: 0, updated: 0 };

      const masterClasses = Array.from(newClasses).slice(0, 100);

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        students: nextStudents,
        settings: { ...dbRef.current.settings, masterClasses },
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return { added, updated };
      }

      void (async () => {
        try {
          const existingSnap = await getDocs(collection(firestoreDb, 'students'));
          const existingIds = new Set(existingSnap.docs.map((d) => d.id));
          const ops: BatchWriteOp[] = recordsToSave.map((rec) => {
            const existsInDb = existingIds.has(rec.id);
            return {
              type: existsInDb ? 'update' : 'set',
              collectionName: 'students',
              docId: rec.id,
              data: existsInDb
                ? { ...rec, serverUpdatedAt: serverTimestamp() }
                : {
                    ...rec,
                    serverCreatedAt: serverTimestamp(),
                    serverUpdatedAt: serverTimestamp(),
                  },
            };
          });
          const settingsSnap = await getDoc(doc(firestoreDb, 'settings', 'global'));
          if (settingsSnap.exists()) {
            ops.push({
              type: 'update',
              collectionName: 'settings',
              docId: 'global',
              data: { masterClasses, serverUpdatedAt: serverTimestamp() },
            });
          }
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'students');
        }
      })();

      return { added, updated };
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteStudent = useCallback(
    (studentId: string) => {
      const relatedGrades = dbRef.current.grades.filter((g) => g.studentId === studentId);
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        students: dbRef.current.students.filter((s) => s.id !== studentId),
        grades: dbRef.current.grades.filter((g) => g.studentId !== studentId),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = [
        { type: 'delete', collectionName: 'students', docId: studentId },
        ...relatedGrades.map((g) => ({
          type: 'delete' as const,
          collectionName: 'grades',
          docId: g.id,
        })),
      ];
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, `students/${studentId}`)
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteStudentsBatch = useCallback(
    (ids: string[]) => {
      if (ids.length === 0) return;
      const idSet = new Set(ids);
      const relatedGrades = dbRef.current.grades.filter((g) => idSet.has(g.studentId));
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        students: dbRef.current.students.filter((s) => !idSet.has(s.id)),
        grades: dbRef.current.grades.filter((g) => !idSet.has(g.studentId)),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = [
        ...ids.map((id) => ({
          type: 'delete' as const,
          collectionName: 'students',
          docId: id,
        })),
        ...relatedGrades.map((g) => ({
          type: 'delete' as const,
          collectionName: 'grades',
          docId: g.id,
        })),
      ];
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, 'students')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  // Teachers
  const upsertTeacher = useCallback(
    (teacher: Teacher) => {
      const safeId = toFirestoreSafeId(teacher.id);
      const cleanTeacher = normalizeTeacher({ ...teacher, id: safeId });
      const existsLocally = dbRef.current.teachers.some((t) => t.id === safeId);

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        teachers: existsLocally
          ? dbRef.current.teachers.map((t) => (t.id === safeId ? cleanTeacher : t))
          : [...dbRef.current.teachers, cleanTeacher],
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void upsertDocSafe('teachers', safeId, cleanTeacher as unknown as Record<string, unknown>).catch(
        (err) =>
          reportFirestoreError(
            err,
            existsLocally ? OperationType.UPDATE : OperationType.CREATE,
            `teachers/${safeId}`
          )
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const importTeachersBatch = useCallback(
    (incoming: Omit<Teacher, 'id'>[]) => {
      let added = 0;
      let updated = 0;
      const next = [...dbRef.current.teachers];
      const recordsToSave: Teacher[] = [];

      incoming.forEach((item, idx) => {
        if (!item.name.trim()) return;
        const existingIdx = item.nip
          ? next.findIndex((t) => t.nip.trim().toLowerCase() === item.nip.trim().toLowerCase())
          : -1;
        if (existingIdx >= 0) {
          const merged = normalizeTeacher(
            {
              ...next[existingIdx],
              ...item,
              photo: item.photo || next[existingIdx].photo,
            },
            idx
          );
          next[existingIdx] = merged;
          recordsToSave.push(merged);
          updated += 1;
        } else {
          const id = `tch-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`;
          const record = normalizeTeacher({ ...item, id }, idx);
          next.push(record);
          recordsToSave.push(record);
          added += 1;
        }
      });

      if (recordsToSave.length === 0) return { added: 0, updated: 0 };

      const nextDb: PortalDatabase = { ...dbRef.current, teachers: next };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return { added, updated };
      }

      void (async () => {
        try {
          const existingSnap = await getDocs(collection(firestoreDb, 'teachers'));
          const existingIds = new Set(existingSnap.docs.map((d) => d.id));
          const ops: BatchWriteOp[] = recordsToSave.map((rec) => {
            const existsInDb = existingIds.has(rec.id);
            return {
              type: existsInDb ? 'update' : 'set',
              collectionName: 'teachers',
              docId: rec.id,
              data: existsInDb
                ? { ...rec, serverUpdatedAt: serverTimestamp() }
                : {
                    ...rec,
                    serverCreatedAt: serverTimestamp(),
                    serverUpdatedAt: serverTimestamp(),
                  },
            };
          });
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'teachers');
        }
      })();

      return { added, updated };
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteTeacher = useCallback(
    (teacherId: string) => {
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        teachers: dbRef.current.teachers.filter((t) => t.id !== teacherId),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void deleteDoc(doc(firestoreDb, 'teachers', teacherId)).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, `teachers/${teacherId}`)
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteTeachersBatch = useCallback(
    (ids: string[]) => {
      if (ids.length === 0) return;
      const idSet = new Set(ids);
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        teachers: dbRef.current.teachers.filter((t) => !idSet.has(t.id)),
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      const ops: BatchWriteOp[] = ids.map((id) => ({
        type: 'delete',
        collectionName: 'teachers',
        docId: id,
      }));
      void commitBatchChunked(ops).catch((err) =>
        reportFirestoreError(err, OperationType.DELETE, 'teachers')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  // User Accounts (keyed by normalized lowercase email)
  const upsertAccount = useCallback(
    async (
      account: UserAccount,
      previousId?: string,
      credentials?: { username?: string; password?: string }
    ): Promise<{ ok: boolean; error?: string }> => {
      const emailLower = normalizeEmailId(account.email || account.id);
      const isBootstrap =
        normalizeEmailId(dbRef.current.session.email) === DEFAULT_ADMIN_EMAIL;
      const isAdminSession =
        isBootstrap || dbRef.current.session.userRole === 'Admin';
      const isSelfEdit =
        normalizeEmailId(dbRef.current.session.accountId || dbRef.current.session.email) ===
        emailLower;

      if (!isAdminSession && !isSelfEdit) {
        return {
          ok: false,
          error: 'Hanya Admin yang dapat mengubah data akun pengguna lain.',
        };
      }

      let resolvedUsername = (credentials?.username ?? account.username ?? '')
        .trim()
        .toLowerCase();
      let resolvedHasPassword = Boolean(account.hasPassword);
      let syncedCredentialRecord: CredentialSyncRecord | null = null;

      if (credentials !== undefined || resolvedUsername) {
        try {
          const token = await getActiveBearerToken();
          if (token) {
            const res = await fetch('/api/accounts/credentials', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                Authorization: `Bearer ${token}`,
              },
              body: JSON.stringify({
                accountId: emailLower,
                previousId,
                username: resolvedUsername,
                password: credentials?.password || '',
                email: emailLower,
                name: account.name,
                role: account.role,
                nip: account.nip,
                status: account.status,
                signatureImage: account.signatureImage,
                schoolStampImage: account.schoolStampImage,
              }),
            });
            const json = (await res.json()) as {
              ok?: boolean;
              error?: string;
              meta?: { username: string; hasPassword: boolean };
              credentialRecord?: CredentialSyncRecord;
            };
            if (!res.ok || !json.ok) {
              return {
                ok: false,
                error: json.error || 'Gagal menyimpan kredensial username/password.',
              };
            }
            if (json.meta) {
              resolvedUsername = json.meta.username;
              resolvedHasPassword = json.meta.hasPassword;
              if (previousId && previousId !== emailLower) {
                delete credentialsMetaRef.current[previousId];
              }
              credentialsMetaRef.current[emailLower] = json.meta;
            }
            if (json.credentialRecord && json.credentialRecord.passwordHash && json.credentialRecord.salt) {
              syncedCredentialRecord = json.credentialRecord;
            }
          }
        } catch {
          return {
            ok: false,
            error: 'Gagal terhubung ke server untuk menyimpan kredensial login.',
          };
        }
      }

      const cleanAccount = normalizeAccount(
        {
          ...account,
          id: emailLower,
          email: emailLower,
          username: resolvedUsername,
          hasPassword: resolvedHasPassword,
          schoolStampImage:
            account.role === 'Kepala Sekolah' ? account.schoolStampImage : '',
        },
        ''
      );

      const filtered = previousId
        ? dbRef.current.accounts.filter((a) => a.id !== previousId)
        : dbRef.current.accounts;
      const already = filtered.some((a) => a.id === emailLower);
      const nextAccounts = already
        ? filtered.map((a) => (a.id === emailLower ? cleanAccount : a))
        : [...filtered, cleanAccount];
      const isCurrent =
        normalizeEmailId(dbRef.current.session.accountId || dbRef.current.session.email) ===
        emailLower;

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        accounts: nextAccounts,
        session: isCurrent
          ? {
              ...dbRef.current.session,
              accountId: cleanAccount.id,
              username: cleanAccount.username || dbRef.current.session.username || '',
              userName: cleanAccount.name,
              userRole: isBootstrap ? 'Admin' : cleanAccount.role,
            }
          : dbRef.current.session,
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        await pushSnapshotToServer(nextDb);
        return { ok: true };
      }

      try {
        const targetRef = doc(firestoreDb, 'accounts', emailLower);
        const targetSnap = await getDoc(targetRef);
        const ops: BatchWriteOp[] = [];
        if (
          isAdminSession &&
          previousId &&
          previousId !== emailLower &&
          previousId !== DEFAULT_ADMIN_EMAIL
        ) {
          ops.push({ type: 'delete', collectionName: 'accounts', docId: previousId });
          ops.push({ type: 'delete', collectionName: 'credential_sync', docId: previousId });
        }
        if (targetSnap.exists()) {
          const existingData = targetSnap.data() as Record<string, unknown>;
          if (isAdminSession) {
            const payload: Record<string, unknown> = {
              ...cleanAccount,
              serverUpdatedAt: serverTimestamp(),
            };
            if (existingData.serverCreatedAt !== undefined) {
              payload.serverCreatedAt = existingData.serverCreatedAt;
            }
            ops.push({
              type: 'set',
              collectionName: 'accounts',
              docId: emailLower,
              data: payload,
            });
          } else {
            ops.push({
              type: 'update',
              collectionName: 'accounts',
              docId: emailLower,
              data: {
                name: cleanAccount.name,
                nip: cleanAccount.nip,
                signatureImage: cleanAccount.signatureImage,
                schoolStampImage: cleanAccount.schoolStampImage,
                serverUpdatedAt: serverTimestamp(),
              },
            });
          }
        } else {
          ops.push({
            type: 'set',
            collectionName: 'accounts',
            docId: emailLower,
            data: {
              ...cleanAccount,
              serverCreatedAt: serverTimestamp(),
              serverUpdatedAt: serverTimestamp(),
            },
          });
        }
        if (isAdminSession && syncedCredentialRecord) {
          const credSyncSnap = await getDoc(doc(firestoreDb, 'credential_sync', emailLower));
          ops.push({
            type: credSyncSnap.exists() ? 'update' : 'set',
            collectionName: 'credential_sync',
            docId: emailLower,
            data: credSyncSnap.exists()
              ? { ...syncedCredentialRecord, serverUpdatedAt: serverTimestamp() }
              : {
                  ...syncedCredentialRecord,
                  serverCreatedAt: serverTimestamp(),
                  serverUpdatedAt: serverTimestamp(),
                },
          });
        }
        await commitBatchChunked(ops);
        return { ok: true };
      } catch (err) {
        reportFirestoreError(err, OperationType.WRITE, `accounts/${emailLower}`);
        return { ok: false, error: 'Gagal menyimpan akun ke database.' };
      }
    },
    [getActiveBearerToken, pushSnapshotToServer, reportFirestoreError]
  );

  const importAccountsBatch = useCallback(
    (incoming: Omit<UserAccount, 'id' | 'lastLogin'>[]) => {
      let added = 0;
      let updated = 0;
      const next = [...dbRef.current.accounts];
      const normalizedRecords: UserAccount[] = [];

      incoming.forEach((item) => {
        const emailLower = normalizeEmailId(item.email);
        if (!item.name.trim() || !emailLower || !emailLower.includes('@')) return;
        const existingIdx = next.findIndex(
          (a) => normalizeEmailId(a.email) === emailLower
        );
        if (existingIdx >= 0) {
          const merged = normalizeAccount(
            {
              ...next[existingIdx],
              ...item,
              id: emailLower,
              email: emailLower,
              signatureImage: item.signatureImage || next[existingIdx].signatureImage,
              schoolStampImage: item.schoolStampImage || next[existingIdx].schoolStampImage,
            },
            ''
          );
          next[existingIdx] = merged;
          normalizedRecords.push(merged);
          updated += 1;
        } else {
          const created = normalizeAccount(
            {
              ...item,
              id: emailLower,
              email: emailLower,
              lastLogin: '-',
              schoolStampImage:
                item.role === 'Kepala Sekolah'
                  ? item.schoolStampImage || dbRef.current.settings.defaultSchoolStampImage
                  : '',
            },
            ''
          );
          next.push(created);
          normalizedRecords.push(created);
          added += 1;
        }
      });

      if (normalizedRecords.length === 0) return { added: 0, updated: 0 };

      const nextDb: PortalDatabase = { ...dbRef.current, accounts: next };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return { added, updated };
      }

      void (async () => {
        try {
          const existingSnap = await getDocs(collection(firestoreDb, 'accounts'));
          const existingMap = new Map(
            existingSnap.docs.map((d) => [d.id, d.data() as Record<string, unknown>])
          );
          const ops: BatchWriteOp[] = normalizedRecords.map((rec) => {
            const existingData = existingMap.get(rec.id);
            if (existingData) {
              const payload: Record<string, unknown> = {
                ...rec,
                serverUpdatedAt: serverTimestamp(),
              };
              if (existingData.serverCreatedAt !== undefined) {
                payload.serverCreatedAt = existingData.serverCreatedAt;
              }
              return {
                type: 'set',
                collectionName: 'accounts',
                docId: rec.id,
                data: payload,
              };
            }
            return {
              type: 'set',
              collectionName: 'accounts',
              docId: rec.id,
              data: {
                ...rec,
                serverCreatedAt: serverTimestamp(),
                serverUpdatedAt: serverTimestamp(),
              },
            };
          });
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'accounts');
        }
      })();

      return { added, updated };
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const deleteAccount = useCallback(
    (accountId: string) => {
      const normalized = normalizeEmailId(accountId);
      if (normalized === DEFAULT_ADMIN_EMAIL) return;
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        accounts: dbRef.current.accounts.filter((a) => a.id !== accountId),
      };
      setDb(nextDb);
      delete credentialsMetaRef.current[normalized];

      void (async () => {
        try {
          const token = await getActiveBearerToken();
          if (token) {
            await fetch(`/api/accounts/credentials/${encodeURIComponent(normalized)}`, {
              method: 'DELETE',
              headers: { Authorization: `Bearer ${token}` },
            });
          }
        } catch {
          // Non-fatal
        }
      })();

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void (async () => {
        try {
          const ops: BatchWriteOp[] = [
            { type: 'delete', collectionName: 'accounts', docId: accountId },
          ];
          const credSyncSnap = await getDoc(doc(firestoreDb, 'credential_sync', normalized));
          if (credSyncSnap.exists()) {
            ops.push({ type: 'delete', collectionName: 'credential_sync', docId: normalized });
          }
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.DELETE, `accounts/${accountId}`);
        }
      })();
    },
    [getActiveBearerToken, pushSnapshotToServer, reportFirestoreError]
  );

  const deleteAccountsBatch = useCallback(
    (ids: string[]) => {
      const idSet = new Set(ids);
      const currentEmail = normalizeEmailId(dbRef.current.session.email);
      const deletableIds = dbRef.current.accounts
        .filter(
          (a) =>
            idSet.has(a.id) &&
            normalizeEmailId(a.email) !== currentEmail &&
            normalizeEmailId(a.email) !== DEFAULT_ADMIN_EMAIL
        )
        .map((a) => a.id);

      if (deletableIds.length === 0) return;
      const delSet = new Set(deletableIds);

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        accounts: dbRef.current.accounts.filter((a) => !delSet.has(a.id)),
      };
      setDb(nextDb);

      void (async () => {
        try {
          const token = await getActiveBearerToken();
          if (token) {
            await Promise.all(
              deletableIds.map((id) =>
                fetch(`/api/accounts/credentials/${encodeURIComponent(id)}`, {
                  method: 'DELETE',
                  headers: { Authorization: `Bearer ${token}` },
                })
              )
            );
          }
        } catch {
          // Non-fatal
        }
      })();

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void (async () => {
        try {
          const ops: BatchWriteOp[] = deletableIds.map((id) => ({
            type: 'delete',
            collectionName: 'accounts',
            docId: id,
          }));
          const credSnap = await getDocs(collection(firestoreDb, 'credential_sync'));
          const existingCredIds = new Set(credSnap.docs.map((d) => d.id));
          deletableIds.forEach((id) => {
            const norm = normalizeEmailId(id);
            if (existingCredIds.has(norm)) {
              ops.push({ type: 'delete', collectionName: 'credential_sync', docId: norm });
            }
          });
          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.DELETE, 'accounts');
        }
      })();
    },
    [getActiveBearerToken, pushSnapshotToServer, reportFirestoreError]
  );

  // Settings
  const updateSettings = useCallback(
    (updater: (prev: SystemSettings) => SystemSettings) => {
      const nextSettings = normalizeSettings(updater(dbRef.current.settings));
      const nextDb: PortalDatabase = {
        ...dbRef.current,
        settings: nextSettings,
        session: {
          ...dbRef.current.session,
          academicYear: nextSettings.academicYear,
          semester: nextSettings.semester,
        },
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void upsertDocSafe('settings', 'global', nextSettings as unknown as Record<string, unknown>).catch(
        (err) => reportFirestoreError(err, OperationType.UPDATE, 'settings/global')
      );
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  // Backup & Restore
  const restoreDatabase = useCallback(
    (incoming: PortalDatabase) => {
      const mergedSettings = normalizeSettings(incoming.settings);
      const students = Array.isArray(incoming.students)
        ? incoming.students.map((s, idx) => normalizeStudent(s, idx))
        : [];
      const teachers = Array.isArray(incoming.teachers)
        ? incoming.teachers.map((t, idx) => normalizeTeacher(t, idx))
        : [];
      const schedules = Array.isArray(incoming.schedules)
        ? incoming.schedules.map((s, idx) => normalizeSchedule(s, idx))
        : [];
      const journals = Array.isArray(incoming.journals)
        ? incoming.journals.map((j, idx) => normalizeJournal(j, idx))
        : [];
      const grades = Array.isArray(incoming.grades)
        ? incoming.grades.map((g, idx) => normalizeGrade(g, idx))
        : [];
      const incomingAccounts = Array.isArray(incoming.accounts)
        ? incoming.accounts
            .filter((a) => a.email && a.email.includes('@'))
            .map((a) => normalizeAccount(a, mergedSettings.defaultSchoolStampImage))
        : [];

      const nextDb: PortalDatabase = {
        ...dbRef.current,
        students,
        teachers,
        schedules,
        journals,
        grades,
        accounts: incomingAccounts.length > 0 ? incomingAccounts : dbRef.current.accounts,
        settings: mergedSettings,
        session: {
          ...dbRef.current.session,
          academicYear: mergedSettings.academicYear,
          semester: mergedSettings.semester,
        },
      };
      setDb(nextDb);

      if (authModeRef.current === 'credentials') {
        void pushSnapshotToServer(nextDb);
        return;
      }

      void (async () => {
        try {
          const ops: BatchWriteOp[] = [];
          const colNames = ['students', 'teachers', 'schedules', 'journals', 'grades'] as const;
          for (const colName of colNames) {
            const snap = await getDocs(collection(firestoreDb, colName));
            snap.docs.forEach((d) => {
              ops.push({ type: 'delete', collectionName: colName, docId: d.id });
            });
          }

          students.forEach((s) =>
            ops.push({
              type: 'set',
              collectionName: 'students',
              docId: s.id,
              data: { ...s, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
            })
          );
          teachers.forEach((t) =>
            ops.push({
              type: 'set',
              collectionName: 'teachers',
              docId: t.id,
              data: { ...t, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
            })
          );
          schedules.forEach((s) =>
            ops.push({
              type: 'set',
              collectionName: 'schedules',
              docId: s.id,
              data: { ...s, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
            })
          );
          journals.forEach((j) =>
            ops.push({
              type: 'set',
              collectionName: 'journals',
              docId: j.id,
              data: { ...j, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
            })
          );
          grades.forEach((g) =>
            ops.push({
              type: 'set',
              collectionName: 'grades',
              docId: g.id,
              data: { ...g, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
            })
          );
          const existingAccountsSnap = await getDocs(collection(firestoreDb, 'accounts'));
          const existingAccountsMap = new Map(
            existingAccountsSnap.docs.map((d) => [d.id, d.data() as Record<string, unknown>])
          );
          incomingAccounts.forEach((a) => {
            const existingData = existingAccountsMap.get(a.id);
            if (existingData) {
              const payload: Record<string, unknown> = {
                ...a,
                serverUpdatedAt: serverTimestamp(),
              };
              if (existingData.serverCreatedAt !== undefined) {
                payload.serverCreatedAt = existingData.serverCreatedAt;
              }
              ops.push({
                type: 'set',
                collectionName: 'accounts',
                docId: a.id,
                data: payload,
              });
            } else {
              ops.push({
                type: 'set',
                collectionName: 'accounts',
                docId: a.id,
                data: { ...a, serverCreatedAt: serverTimestamp(), serverUpdatedAt: serverTimestamp() },
              });
            }
          });
          const settingsSnap = await getDoc(doc(firestoreDb, 'settings', 'global'));
          ops.push({
            type: settingsSnap.exists() ? 'update' : 'set',
            collectionName: 'settings',
            docId: 'global',
            data: settingsSnap.exists()
              ? { ...mergedSettings, serverUpdatedAt: serverTimestamp() }
              : {
                  ...mergedSettings,
                  serverCreatedAt: serverTimestamp(),
                  serverUpdatedAt: serverTimestamp(),
                },
          });

          await commitBatchChunked(ops);
        } catch (err) {
          reportFirestoreError(err, OperationType.WRITE, 'restoreDatabase');
        }
      })();
    },
    [pushSnapshotToServer, reportFirestoreError]
  );

  const resetAllData = useCallback(() => {
    const cleanSettings = createDefaultSettings();
    const nextDb: PortalDatabase = {
      ...dbRef.current,
      students: [],
      teachers: [],
      schedules: [],
      journals: [],
      grades: [],
      settings: cleanSettings,
      session: {
        ...dbRef.current.session,
        academicYear: cleanSettings.academicYear,
        semester: cleanSettings.semester,
      },
    };
    setDb(nextDb);

    if (authModeRef.current === 'credentials') {
      void pushSnapshotToServer(nextDb);
      return;
    }

    void (async () => {
      try {
        const ops: BatchWriteOp[] = [];
        const colNames = ['students', 'teachers', 'schedules', 'journals', 'grades'] as const;
        for (const colName of colNames) {
          const snap = await getDocs(collection(firestoreDb, colName));
          snap.docs.forEach((d) => {
            ops.push({ type: 'delete', collectionName: colName, docId: d.id });
          });
        }
        const settingsSnap = await getDoc(doc(firestoreDb, 'settings', 'global'));
        ops.push({
          type: settingsSnap.exists() ? 'update' : 'set',
          collectionName: 'settings',
          docId: 'global',
          data: settingsSnap.exists()
            ? { ...cleanSettings, serverUpdatedAt: serverTimestamp() }
            : {
                ...cleanSettings,
                serverCreatedAt: serverTimestamp(),
                serverUpdatedAt: serverTimestamp(),
              },
        });
        await commitBatchChunked(ops);
      } catch (err) {
        reportFirestoreError(err, OperationType.WRITE, 'resetAllData');
      }
    })();
  }, [pushSnapshotToServer, reportFirestoreError]);

  return {
    db,
    authLoading,
    signingIn,
    dataLoading,
    accessDeniedMessage,
    pendingUserEmail,
    fatalError,
    loginWithGoogle,
    loginWithCredentials,
    logout,
    upsertJournal,
    importJournalsBatch,
    deleteJournals,
    batchUpdateJournals,
    upsertSchedule,
    importSchedulesBatch,
    deleteSchedule,
    deleteSchedules,
    upsertGradesBatch,
    deleteGradesForStudents,
    upsertStudent,
    importStudentsBatch,
    deleteStudent,
    deleteStudentsBatch,
    upsertTeacher,
    importTeachersBatch,
    deleteTeacher,
    deleteTeachersBatch,
    upsertAccount,
    importAccountsBatch,
    deleteAccount,
    deleteAccountsBatch,
    updateSettings,
    restoreDatabase,
    resetAllData,
  };
}
