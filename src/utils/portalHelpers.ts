import { AttendanceStatus, GradeEntry, JournalEntry, Student } from '../types';

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
