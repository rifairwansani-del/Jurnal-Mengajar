import {
  AttendanceStatus,
  GradeEntry,
  JournalEntry,
  KopLineKey,
  KopLineStyle,
  KopSuratSettings,
  KopTextAlign,
  PrintFontFamily,
  Student,
} from '../types';

export const ATTENDANCE_STATUS_META: Record<
  AttendanceStatus,
  { label: string; short: string; bg: string; text: string; border: string; activeBg: string }
> = {
  H: {
    label: 'Hadir',
    short: 'H',
    bg: 'bg-emerald-50',
    text: 'text-emerald-700',
    border: 'border-emerald-300',
    activeBg: 'bg-emerald-600 text-white border-emerald-600',
  },
  S: {
    label: 'Sakit',
    short: 'S',
    bg: 'bg-amber-50',
    text: 'text-amber-700',
    border: 'border-amber-300',
    activeBg: 'bg-amber-500 text-white border-amber-500',
  },
  I: {
    label: 'Izin',
    short: 'I',
    bg: 'bg-sky-50',
    text: 'text-sky-700',
    border: 'border-sky-300',
    activeBg: 'bg-sky-600 text-white border-sky-600',
  },
  A: {
    label: 'Alpha',
    short: 'A',
    bg: 'bg-rose-50',
    text: 'text-rose-700',
    border: 'border-rose-300',
    activeBg: 'bg-rose-600 text-white border-rose-600',
  },
  D: {
    label: 'Dispensasi',
    short: 'D',
    bg: 'bg-violet-50',
    text: 'text-violet-700',
    border: 'border-violet-300',
    activeBg: 'bg-violet-600 text-white border-violet-600',
  },
};

const BULAN_ID = [
  'Januari',
  'Februari',
  'Maret',
  'April',
  'Mei',
  'Juni',
  'Juli',
  'Agustus',
  'September',
  'Oktober',
  'November',
  'Desember',
];

const BULAN_SHORT = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'Mei',
  'Jun',
  'Jul',
  'Agu',
  'Sep',
  'Okt',
  'Nov',
  'Des',
];

export function getTodayISO(): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

export function formatShortDateId(isoDate: string): string {
  if (!isoDate) return '-';
  const parts = isoDate.split('-');
  if (parts.length !== 3) return isoDate;
  const year = parts[0];
  const monthIdx = Number(parts[1]) - 1;
  const day = parts[2];
  return `${day} ${BULAN_SHORT[monthIdx] ?? parts[1]} ${year}`;
}

export function formatCompactDateId(isoDate: string): string {
  if (!isoDate) return '-';
  const parts = isoDate.split('-');
  if (parts.length !== 3) return isoDate;
  return `${parts[2]}/${parts[1]}`;
}

export function formatLongDateId(isoDate: string): string {
  if (!isoDate) return '-';
  const parts = isoDate.split('-');
  if (parts.length !== 3) return isoDate;
  const year = parts[0];
  const monthIdx = Number(parts[1]) - 1;
  const day = Number(parts[2]);
  return `${day} ${BULAN_ID[monthIdx] ?? parts[1]} ${year}`;
}

export function summarizeAttendanceCounts(journals: JournalEntry[]) {
  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
  let totalRecords = 0;

  journals.forEach((j) => {
    Object.values(j.attendance).forEach((rec) => {
      if (rec && counts[rec.status] !== undefined) {
        counts[rec.status] += 1;
        totalRecords += 1;
      }
    });
  });

  const presentPercent = totalRecords > 0 ? (counts.H / totalRecords) * 100 : 0;
  return { counts, totalRecords, presentPercent };
}

export function summarizeJournalRow(journal: JournalEntry) {
  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
  let total = 0;
  Object.values(journal.attendance).forEach((rec) => {
    if (rec && counts[rec.status] !== undefined) {
      counts[rec.status] += 1;
      total += 1;
    }
  });
  return { counts, total };
}

export function computeClassSummary(
  classNames: string[],
  students: Student[],
  journals: JournalEntry[],
  grades: GradeEntry[]
) {
  return classNames.map((cls) => {
    const classStudents = students.filter((s) => s.className === cls);
    const classJournals = journals.filter((j) => j.className === cls);
    const att = summarizeAttendanceCounts(classJournals);

    const classGrades = grades.filter((g) => g.className === cls && g.score !== null);
    const avgScore =
      classGrades.length > 0
        ? classGrades.reduce((sum, g) => sum + (g.score ?? 0), 0) / classGrades.length
        : 0;

    return {
      className: cls,
      studentCount: classStudents.length,
      journalCount: classJournals.length,
      counts: att.counts,
      totalAttendance: att.totalRecords,
      presentPercent: att.presentPercent,
      gradedCount: classGrades.length,
      avgScore,
    };
  });
}

export function toFirestoreSafeId(raw: string): string {
  return raw.trim().replace(/\//g, '-');
}

export function normalizeEmailId(email: string): string {
  return email.trim().toLowerCase().replace(/\//g, '');
}

export function makeGradeId(
  academicYear: string,
  semester: string,
  className: string,
  subject: string,
  examType: string,
  studentId: string
): string {
  const safe = (v: string) => toFirestoreSafeId(v);
  return `${safe(academicYear)}__${safe(semester)}__${safe(className)}__${safe(subject)}__${safe(examType)}__${safe(studentId)}`;
}

export async function imageUrlToPngDataUrl(src: string): Promise<string | null> {
  if (!src) return null;
  if (src.startsWith('data:image/png')) return src;
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        const size = 240;
        canvas.width = size;
        canvas.height = size;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(null);
          return;
        }
        ctx.clearRect(0, 0, size, size);
        const ratio = Math.min(size / (img.width || size), size / (img.height || size));
        const w = (img.width || size) * ratio;
        const h = (img.height || size) * ratio;
        ctx.drawImage(img, (size - w) / 2, (size - h) / 2, w, h);
        resolve(canvas.toDataURL('image/png'));
      } catch {
        resolve(null);
      }
    };
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

export function compressCanvasToDataUrl(
  sourceCanvas: HTMLCanvasElement,
  maxDataUrlChars = 190000
): string {
  let out = sourceCanvas.toDataURL('image/png');
  if (out.length <= maxDataUrlChars) return out;

  const scales = [1, 0.75, 0.5, 0.35];
  const qualities = [0.85, 0.7, 0.55, 0.4];

  for (const scale of scales) {
    const w = Math.max(1, Math.round(sourceCanvas.width * scale));
    const h = Math.max(1, Math.round(sourceCanvas.height * scale));
    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    if (!ctx) continue;
    ctx.clearRect(0, 0, w, h);
    ctx.drawImage(sourceCanvas, 0, 0, w, h);

    const png = canvas.toDataURL('image/png');
    if (png.length <= maxDataUrlChars) return png;

    for (const q of qualities) {
      const webp = canvas.toDataURL('image/webp', q);
      if (webp.length <= maxDataUrlChars) return webp;
      const jpeg = canvas.toDataURL('image/jpeg', q);
      if (jpeg.length <= maxDataUrlChars) return jpeg;
    }
  }

  return '';
}

export async function compressImageFileToDataUrl(
  file: File,
  maxDimension = 480,
  maxDataUrlChars = 190000
): Promise<string> {
  const rawDataUrl = await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === 'string' ? reader.result : '');
    reader.onerror = () => reject(new Error('Gagal membaca file gambar.'));
    reader.readAsDataURL(file);
  });

  if (!rawDataUrl) return '';
  if (rawDataUrl.length <= maxDataUrlChars && !rawDataUrl.startsWith('data:image/svg+xml')) {
    return rawDataUrl;
  }

  return new Promise<string>((resolve) => {
    const img = new Image();
    img.onload = () => {
      try {
        const scale = Math.min(1, maxDimension / Math.max(img.width || 1, img.height || 1));
        const w = Math.max(1, Math.round((img.width || maxDimension) * scale));
        const h = Math.max(1, Math.round((img.height || maxDimension) * scale));
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(rawDataUrl.length <= maxDataUrlChars ? rawDataUrl : '');
          return;
        }
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        resolve(compressCanvasToDataUrl(canvas, maxDataUrlChars));
      } catch {
        resolve(rawDataUrl.length <= maxDataUrlChars ? rawDataUrl : '');
      }
    };
    img.onerror = () => resolve(rawDataUrl.length <= maxDataUrlChars ? rawDataUrl : '');
    img.src = rawDataUrl;
  });
}

export const DUMMY_AVATARS = {
  studentMale: '/assets/aistudio/Siswa_Putra_dd500b95.jpg',
  studentFemale: '/assets/aistudio/Siswi_Putri_972800f4.jpg',
  teacherMale: '/assets/aistudio/Guru_Putra_75beabf3.png',
  teacherFemale: '/assets/aistudio/Guru_Putri_05c5b30c.png',
  tuMale: '/assets/aistudio/TU_Purtra_a36e2763.png',
  tuFemale: '/assets/aistudio/TU_Purtri_05f16c00.png',
} as const;

export function isTuOrStaffPosition(position?: string): boolean {
  if (!position) return false;
  return /\b(tu|tata\s*usaha|staf|staff|admin|administrasi|operator|pustakawan|perpustakaan|laboran|tendik|kepegawaian|keuangan|bendahara|satpam|kebersihan)\b/i.test(
    position
  );
}

export function getStudentDummyPhoto(gender?: 'L' | 'P' | string): string {
  return gender === 'P' ? DUMMY_AVATARS.studentFemale : DUMMY_AVATARS.studentMale;
}

export function getTeacherDummyPhoto(gender?: 'L' | 'P' | string, position?: string): string {
  if (isTuOrStaffPosition(position)) {
    return gender === 'P' ? DUMMY_AVATARS.tuFemale : DUMMY_AVATARS.tuMale;
  }
  return gender === 'P' ? DUMMY_AVATARS.teacherFemale : DUMMY_AVATARS.teacherMale;
}

function unwrapSpreadsheetUrlFormula(raw: string): string {
  const trimmed = raw.trim();
  const formulaMatch = trimmed.match(/^=(?:IMAGE|HYPERLINK)\s*\(\s*["']([^"']+)["']/i);
  if (formulaMatch?.[1]) {
    return formulaMatch[1].trim();
  }
  return trimmed;
}

export function extractGoogleDriveFileId(rawUrl?: string): string | null {
  if (!rawUrl) return null;
  const cleaned = unwrapSpreadsheetUrlFormula(rawUrl);
  if (!cleaned || cleaned.startsWith('data:')) return null;

  const fileDMatch = cleaned.match(
    /(?:drive|docs)\.google\.com\/(?:a\/[^/]+\/)?file\/d\/([a-zA-Z0-9_-]{10,})/i
  );
  if (fileDMatch?.[1]) return fileDMatch[1];

  const lh3Match = cleaned.match(/lh3\.googleusercontent\.com\/d\/([a-zA-Z0-9_-]{10,})/i);
  if (lh3Match?.[1]) return lh3Match[1];

  const driveViewerMatch = cleaned.match(
    /drive\.google\.com\/drive-viewer\/([a-zA-Z0-9_-]{10,})/i
  );
  if (driveViewerMatch?.[1]) return driveViewerMatch[1];

  if (/(?:drive|docs)\.google\.com/i.test(cleaned)) {
    const queryIdMatch = cleaned.match(/[?&]id=([a-zA-Z0-9_-]{10,})/i);
    if (queryIdMatch?.[1]) return queryIdMatch[1];
  }

  return null;
}

export function isGoogleDrivePhotoUrl(rawUrl?: string): boolean {
  return Boolean(extractGoogleDriveFileId(rawUrl));
}

export function normalizePhotoInputUrl(rawUrl?: string): string {
  if (!rawUrl) return '';
  const cleaned = unwrapSpreadsheetUrlFormula(rawUrl);
  if (!cleaned) return '';
  if (cleaned.startsWith('data:image/') || cleaned.startsWith('/assets/')) {
    return cleaned;
  }
  const driveId = extractGoogleDriveFileId(cleaned);
  if (driveId) {
    return `https://drive.google.com/thumbnail?id=${driveId}&sz=w800`;
  }
  return cleaned;
}

export function getDriveImageCandidates(rawUrl?: string): string[] {
  if (!rawUrl) return [];
  const cleaned = unwrapSpreadsheetUrlFormula(rawUrl);
  if (!cleaned) return [];

  const driveId = extractGoogleDriveFileId(cleaned);
  if (driveId) {
    return [
      `https://drive.google.com/thumbnail?id=${driveId}&sz=w800`,
      `https://lh3.googleusercontent.com/d/${driveId}=w800`,
      `https://drive.google.com/uc?export=view&id=${driveId}`,
    ];
  }

  if (
    cleaned.startsWith('data:image/') ||
    cleaned.startsWith('http://') ||
    cleaned.startsWith('https://') ||
    cleaned.startsWith('/')
  ) {
    return [cleaned];
  }

  return [];
}

export const KOP_LINE_KEYS: KopLineKey[] = [
  'line1',
  'line2',
  'line3',
  'line4',
  'line5',
  'line6',
];

export const DEFAULT_KOP_LINE_STYLES: Record<KopLineKey, KopLineStyle> = {
  line1: { fontSize: 11, bold: false, italic: false, underline: false, align: 'center' },
  line2: { fontSize: 12, bold: true, italic: false, underline: false, align: 'center' },
  line3: { fontSize: 14, bold: true, italic: false, underline: false, align: 'center' },
  line4: { fontSize: 9, bold: false, italic: false, underline: false, align: 'center' },
  line5: { fontSize: 9, bold: false, italic: false, underline: false, align: 'center' },
  line6: { fontSize: 9, bold: false, italic: false, underline: false, align: 'center' },
};

const VALID_KOP_ALIGNMENTS: KopTextAlign[] = ['left', 'center', 'right', 'justify'];
const VALID_PRINT_FONTS: PrintFontFamily[] = [
  'Times New Roman',
  'Tahoma',
  'Arial',
  'Verdana',
  'Calibri',
  'Georgia',
];

export function getPrintFontCss(fontFamily?: PrintFontFamily): string {
  switch (fontFamily) {
    case 'Tahoma':
      return 'Tahoma, Geneva, sans-serif';
    case 'Arial':
      return 'Arial, Helvetica, sans-serif';
    case 'Verdana':
      return 'Verdana, Geneva, sans-serif';
    case 'Calibri':
      return 'Calibri, "Segoe UI", Arial, sans-serif';
    case 'Georgia':
      return 'Georgia, "Times New Roman", serif';
    case 'Times New Roman':
    default:
      return '"Times New Roman", Times, serif';
  }
}

export function resolveKopLineStyle(
  kop: Partial<KopSuratSettings> | undefined,
  key: KopLineKey
): KopLineStyle {
  const def = DEFAULT_KOP_LINE_STYLES[key];
  const raw = kop?.lineStyles?.[key];
  if (!raw) return { ...def };
  const rawSize = Number(raw.fontSize);
  return {
    fontSize: Number.isFinite(rawSize) ? Math.max(7, Math.min(28, Math.round(rawSize))) : def.fontSize,
    bold: typeof raw.bold === 'boolean' ? raw.bold : def.bold,
    italic: typeof raw.italic === 'boolean' ? raw.italic : def.italic,
    underline: typeof raw.underline === 'boolean' ? raw.underline : def.underline,
    align: VALID_KOP_ALIGNMENTS.includes(raw.align) ? raw.align : def.align,
  };
}

export const DEFAULT_KOP_LINE_SPACING = 1.25;

export function normalizeKopLineSpacing(value: unknown, fallback = DEFAULT_KOP_LINE_SPACING): number {
  const num = Number(value);
  if (!Number.isFinite(num)) return fallback;
  return Math.round(Math.max(0.85, Math.min(2.4, num)) * 100) / 100;
}

export function normalizeKopSurat(
  raw: Partial<KopSuratSettings> | undefined,
  fallback: KopSuratSettings
): KopSuratSettings {
  const merged = { ...fallback, ...(raw || {}) };
  const resolvedStyles: Record<KopLineKey, KopLineStyle> = {
    line1: resolveKopLineStyle(merged, 'line1'),
    line2: resolveKopLineStyle(merged, 'line2'),
    line3: resolveKopLineStyle(merged, 'line3'),
    line4: resolveKopLineStyle(merged, 'line4'),
    line5: resolveKopLineStyle(merged, 'line5'),
    line6: resolveKopLineStyle(merged, 'line6'),
  };
  return {
    ...merged,
    logoLeftSize: Math.max(
      28,
      Math.min(150, Number(merged.logoLeftSize ?? fallback.logoLeftSize ?? 72) || 72)
    ),
    logoRightSize: Math.max(
      28,
      Math.min(150, Number(merged.logoRightSize ?? fallback.logoRightSize ?? 72) || 72)
    ),
    fontFamily: VALID_PRINT_FONTS.includes(merged.fontFamily as PrintFontFamily)
      ? merged.fontFamily
      : fallback.fontFamily || 'Times New Roman',
    lineSpacing: normalizeKopLineSpacing(
      merged.lineSpacing,
      normalizeKopLineSpacing(fallback.lineSpacing, DEFAULT_KOP_LINE_SPACING)
    ),
    lineStyles: resolvedStyles,
  };
}

export function compareClassNames(a: string, b: string): number {
  return (a || '').trim().localeCompare((b || '').trim(), 'id', {
    numeric: true,
    sensitivity: 'base',
  });
}

function parseAbsentSortKey(absentNo?: string): { num: number; raw: string } {
  const raw = (absentNo || '').trim();
  if (!raw) return { num: Number.MAX_SAFE_INTEGER, raw: '' };
  const match = raw.match(/^\d+/);
  if (match) {
    return { num: Number(match[0]), raw };
  }
  return { num: Number.MAX_SAFE_INTEGER - 1, raw };
}

export function compareAbsentNo(a?: string, b?: string): number {
  const ka = parseAbsentSortKey(a);
  const kb = parseAbsentSortKey(b);
  if (ka.num !== kb.num) return ka.num - kb.num;
  if (ka.raw && !kb.raw) return -1;
  if (!ka.raw && kb.raw) return 1;
  return ka.raw.localeCompare(kb.raw, 'id', { numeric: true, sensitivity: 'base' });
}

export function compareStudentsByClassAndAbsent(
  a: Pick<Student, 'className' | 'absentNo' | 'name'>,
  b: Pick<Student, 'className' | 'absentNo' | 'name'>
): number {
  return (
    compareClassNames(a.className, b.className) ||
    compareAbsentNo(a.absentNo, b.absentNo) ||
    a.name.localeCompare(b.name, 'id', { sensitivity: 'base' })
  );
}



