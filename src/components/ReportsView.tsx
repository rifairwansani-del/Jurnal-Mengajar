import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  CheckCircle2,
  ChevronDown,
  FileDown,
  FileSpreadsheet,
  FileText,
  FileType2,
  Move,
  Printer,
  RotateCcw,
  Save,
  Stamp,
  Undo2,
  UserCheck,
} from 'lucide-react';
import {
  AttendanceStatus,
  PaperOrientation,
  PaperSizeId,
  PortalDatabase,
  PrintFontFamily,
  ReportSectionId,
  SignatureLayoutPreset,
  SignaturePositionSettings,
  SystemSettings,
} from '../types';
import {
  PAPER_SIZE_CONFIG,
  exportPreviewDomToWordDoc,
  resolveActiveSigner,
} from '../utils/exportHelpers';
import {
  formatCompactDateId,
  formatLongDateId,
  formatShortDateId,
  getTodayISO,
  imageUrlToPngDataUrl,
  makeGradeId,
  summarizeJournalRow,
} from '../utils/portalHelpers';

type ReportCategory = 'Mengajar' | 'Presensi' | 'Akademik' | 'Kesiswaan' | 'Pimpinan';

const REPORT_TYPE_OPTIONS: {
  id: ReportSectionId;
  label: string;
  category: ReportCategory;
  description: string;
  officialTitle: string;
}[] = [
  {
    id: 'journal',
    label: 'Jurnal Mengajar',
    category: 'Mengajar',
    description: 'Catatan materi, jam & kegiatan KBM',
    officialTitle: 'LAPORAN JURNAL MENGAJAR GURU',
  },
  {
    id: 'attendance',
    label: 'Kehadiran Mendatar',
    category: 'Presensi',
    description: 'Matriks H/S/I/A/D per tanggal',
    officialTitle: 'LEGER KEHADIRAN SISWA MENDATAR',
  },
  {
    id: 'attendance_monthly',
    label: 'Rekap Presensi Bulanan',
    category: 'Presensi',
    description: 'Akumulasi & persentase hadir siswa',
    officialTitle: 'REKAPITULASI PRESENSI BULANAN SISWA',
  },
  {
    id: 'attendance_monitoring',
    label: 'Monitoring S/I/A/D',
    category: 'Presensi',
    description: 'Rekap & kronologi tidak hadir',
    officialTitle: 'LAPORAN MONITORING KETIDAKHADIRAN (S/I/A/D)',
  },
  {
    id: 'grades',
    label: 'Rekap Nilai Mapel',
    category: 'Akademik',
    description: 'Matriks skor ujian & status KKM',
    officialTitle: 'REKAPITULASI MATRIKS NILAI MATA PELAJARAN',
  },
  {
    id: 'grades_remedial',
    label: 'Ketuntasan & Remedial',
    category: 'Akademik',
    description: 'Daya serap & daftar siswa < KKM',
    officialTitle: 'LAPORAN DAYA SERAP, KETUNTASAN & REMEDIAL (< KKM)',
  },
  {
    id: 'grades_leger',
    label: 'Leger Gabungan Kelas',
    category: 'Akademik',
    description: 'Nilai lintas mapel & peringkat',
    officialTitle: 'LEGER GABUNGAN NILAI LINTAS MAPEL & PERINGKAT KELAS',
  },
  {
    id: 'students_roster',
    label: 'Daftar Peserta Didik',
    category: 'Kesiswaan',
    description: 'Daftar induk siswa & wali kelas',
    officialTitle: 'DAFTAR INDUK PESERTA DIDIK & WALI KELAS',
  },
  {
    id: 'teacher_activity',
    label: 'Aktivitas & Penilaian Guru',
    category: 'Pimpinan',
    description: 'Ringkasan & riwayat kinerja guru',
    officialTitle: 'LAPORAN AKTIVITAS & PENILAIAN KINERJA GURU',
  },
];

const CATEGORY_BADGE_CLASS: Record<ReportCategory, string> = {
  Mengajar: 'bg-indigo-50 text-indigo-700 border-indigo-200',
  Presensi: 'bg-emerald-50 text-emerald-700 border-emerald-200',
  Akademik: 'bg-amber-50 text-amber-800 border-amber-200',
  Kesiswaan: 'bg-sky-50 text-sky-700 border-sky-200',
  Pimpinan: 'bg-violet-50 text-violet-700 border-violet-200',
};

const SHORT_MONTHS_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agt', 'Sep', 'Okt', 'Nov', 'Des'];

function formatMonthLabelId(ym: string): string {
  const [y, m] = ym.split('-');
  const idx = Number(m) - 1;
  if (!y || Number.isNaN(idx) || idx < 0 || idx > 11) return ym;
  return `${SHORT_MONTHS_ID[idx]} ${y}`;
}

const FONT_OPTIONS: PrintFontFamily[] = [
  'Times New Roman',
  'Tahoma',
  'Arial',
  'Verdana',
  'Calibri',
  'Georgia',
];

const PRESET_OPTIONS: { id: SignatureLayoutPreset; label: string }[] = [
  { id: 'kiri-kanan', label: 'Kepsek Kiri · Guru Kanan' },
  { id: 'kanan-kiri', label: 'Guru Kiri · Kepsek Kanan' },
  { id: 'tengah', label: 'Sejajar Tengah' },
  { id: 'kiri', label: 'Rata Kiri Vertikal' },
  { id: 'kanan', label: 'Rata Kanan Vertikal' },
  { id: 'bebas', label: 'Posisi Bebas (Seret / Geser)' },
];

interface ReportsViewProps {
  db: PortalDatabase;
  onUpdateSettings: (updater: (prev: SystemSettings) => SystemSettings) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
}

export function ReportsView({ db, onUpdateSettings, onToast }: ReportsViewProps) {
  const { students, teachers, schedules, journals, grades, accounts, settings, session } = db;

  const classOptions = useMemo(() => {
    const set = new Set(settings.masterClasses);
    students.forEach((s) => set.add(s.className));
    return Array.from(set);
  }, [settings.masterClasses, students]);

  const savedPrefs = settings.reportPrintPreferences;

  // Filters & Print Preferences (initialized from persisted settings)
  const [selectedReportTypes, setSelectedReportTypes] = useState<ReportSectionId[]>(() =>
    savedPrefs?.selectedReportTypes && savedPrefs.selectedReportTypes.length > 0
      ? savedPrefs.selectedReportTypes
      : ['journal']
  );
  const [selectedClasses, setSelectedClasses] = useState<string[]>(() =>
    savedPrefs?.selectedClasses && savedPrefs.selectedClasses.length > 0
      ? savedPrefs.selectedClasses
      : classOptions
  );
  const [subjectFilter, setSubjectFilter] = useState<string>(
    () => savedPrefs?.subjectFilter || 'ALL'
  );
  const [startDate, setStartDate] = useState<string>(() => savedPrefs?.startDate || '');
  const [endDate, setEndDate] = useState<string>(() => savedPrefs?.endDate || '');
  const [signerAccountId, setSignerAccountId] = useState<string>(
    () => savedPrefs?.signerAccountId ?? ''
  );

  // Multi-select dropdown states
  const [reportDropdownOpen, setReportDropdownOpen] = useState(false);
  const [classDropdownOpen, setClassDropdownOpen] = useState(false);
  const reportDropdownRef = useRef<HTMLDivElement | null>(null);
  const dropdownRef = useRef<HTMLDivElement | null>(null);
  const previewRef = useRef<HTMLDivElement | null>(null);
  const sigCanvasRef = useRef<HTMLDivElement | null>(null);

  // Real-Time Preview Switches (initialized from persisted settings)
  const [showKop, setShowKop] = useState<boolean>(() => savedPrefs?.showKop ?? true);
  const [showTtdKepsek, setShowTtdKepsek] = useState<boolean>(
    () => savedPrefs?.showTtdKepsek ?? true
  );
  const [showTtdGuru, setShowTtdGuru] = useState<boolean>(() => savedPrefs?.showTtdGuru ?? true);
  const [showSignatureImage, setShowSignatureImage] = useState<boolean>(
    () => savedPrefs?.showSignatureImage ?? true
  );
  const [orientation, setOrientation] = useState<PaperOrientation>(
    () => savedPrefs?.orientation || 'landscape'
  );
  const [paperSize, setPaperSize] = useState<PaperSizeId>(() => savedPrefs?.paperSize || 'A4');
  const [fontFamily, setFontFamily] = useState<PrintFontFamily>(
    () => savedPrefs?.fontFamily || 'Times New Roman'
  );
  const [sigPos, setSigPos] = useState<SignaturePositionSettings>(
    () => settings.signaturePosition
  );
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);
  const [isExportingWord, setIsExportingWord] = useState(false);

  // Dragging state for 'bebas' signature positioning
  const [dragTarget, setDragTarget] = useState<'kepsek' | 'guru' | null>(null);
  const dragOriginRef = useRef<{
    startX: number;
    startY: number;
    initX: number;
    initY: number;
  } | null>(null);

  const activeSigner = useMemo(
    () => resolveActiveSigner(db, signerAccountId || undefined),
    [db, signerAccountId]
  );

  const updateSigPos = (patch: Partial<SignaturePositionSettings>) => {
    setSigPos((prev) => ({
      ...prev,
      ...patch,
    }));
  };

  const hasUnsavedChanges = useMemo(() => {
    const prefs = settings.reportPrintPreferences;
    const savedSig = settings.signaturePosition;
    const reportTypesChanged =
      selectedReportTypes.length !== prefs.selectedReportTypes.length ||
      selectedReportTypes.some((r, i) => r !== prefs.selectedReportTypes[i]);
    const classesChanged =
      selectedClasses.length !== prefs.selectedClasses.length ||
      selectedClasses.some((c, i) => c !== prefs.selectedClasses[i]);
    const prefsChanged =
      reportTypesChanged ||
      classesChanged ||
      subjectFilter !== prefs.subjectFilter ||
      startDate !== prefs.startDate ||
      endDate !== prefs.endDate ||
      signerAccountId !== prefs.signerAccountId ||
      showKop !== prefs.showKop ||
      showTtdKepsek !== prefs.showTtdKepsek ||
      showTtdGuru !== prefs.showTtdGuru ||
      showSignatureImage !== prefs.showSignatureImage ||
      orientation !== prefs.orientation ||
      paperSize !== prefs.paperSize ||
      fontFamily !== prefs.fontFamily;
    const sigChanged =
      sigPos.preset !== savedSig.preset ||
      sigPos.principalX !== savedSig.principalX ||
      sigPos.principalY !== savedSig.principalY ||
      sigPos.teacherX !== savedSig.teacherX ||
      sigPos.teacherY !== savedSig.teacherY ||
      sigPos.showStamp !== savedSig.showStamp ||
      sigPos.stampTarget !== savedSig.stampTarget ||
      (sigPos.principalSignatureSize ?? savedSig.signatureSize ?? 60) !==
        (savedSig.principalSignatureSize ?? savedSig.signatureSize ?? 60) ||
      (sigPos.teacherSignatureSize ?? savedSig.signatureSize ?? 60) !==
        (savedSig.teacherSignatureSize ?? savedSig.signatureSize ?? 60) ||
      (sigPos.stampSize ?? 70) !== (savedSig.stampSize ?? 70);
    return prefsChanged || sigChanged;
  }, [
    settings.reportPrintPreferences,
    settings.signaturePosition,
    selectedReportTypes,
    selectedClasses,
    subjectFilter,
    startDate,
    endDate,
    signerAccountId,
    showKop,
    showTtdKepsek,
    showTtdGuru,
    showSignatureImage,
    orientation,
    paperSize,
    fontFamily,
    sigPos,
  ]);

  const handleSavePrintSettings = () => {
    onUpdateSettings((prev) => ({
      ...prev,
      signaturePosition: sigPos,
      reportPrintPreferences: {
        ...prev.reportPrintPreferences,
        selectedReportTypes,
        selectedClasses,
        subjectFilter,
        startDate,
        endDate,
        signerAccountId,
        showKop,
        showTtdKepsek,
        showTtdGuru,
        showSignatureImage,
        orientation,
        paperSize,
        fontFamily,
      },
    }));
    onToast('success', 'Pengaturan cetak laporan resmi berhasil disimpan permanen.');
  };

  const handleRevertToSaved = () => {
    const prefs = settings.reportPrintPreferences;
    setSelectedReportTypes(
      prefs.selectedReportTypes.length > 0 ? prefs.selectedReportTypes : ['journal']
    );
    setSelectedClasses(prefs.selectedClasses.length > 0 ? prefs.selectedClasses : classOptions);
    setSubjectFilter(prefs.subjectFilter);
    setStartDate(prefs.startDate);
    setEndDate(prefs.endDate);
    setSignerAccountId(prefs.signerAccountId);
    setShowKop(prefs.showKop);
    setShowTtdKepsek(prefs.showTtdKepsek);
    setShowTtdGuru(prefs.showTtdGuru);
    setShowSignatureImage(prefs.showSignatureImage);
    setOrientation(prefs.orientation);
    setPaperSize(prefs.paperSize);
    setFontFamily(prefs.fontFamily);
    setSigPos(settings.signaturePosition);
    onToast('info', 'Pengaturan cetak dikembalikan ke versi tersimpan.');
  };

  const getPresetCoords = (preset: SignatureLayoutPreset) => {
    switch (preset) {
      case 'kanan-kiri':
        return { principalX: 62, principalY: 0, teacherX: 4, teacherY: 0 };
      case 'tengah':
        return { principalX: 20, principalY: 0, teacherX: 54, teacherY: 0 };
      case 'kiri':
        return { principalX: 4, principalY: 0, teacherX: 4, teacherY: 130 };
      case 'kanan':
        return { principalX: 62, principalY: 0, teacherX: 62, teacherY: 130 };
      case 'kiri-kanan':
        return { principalX: 4, principalY: 0, teacherX: 62, teacherY: 0 };
      case 'bebas':
      default:
        return {
          principalX: sigPos.principalX,
          principalY: sigPos.principalY,
          teacherX: sigPos.teacherX,
          teacherY: sigPos.teacherY,
        };
    }
  };

  const handlePresetChange = (preset: SignatureLayoutPreset) => {
    if (preset === 'bebas') {
      updateSigPos({ preset: 'bebas' });
      return;
    }
    const coords = getPresetCoords(preset);
    updateSigPos({ preset, ...coords });
  };

  useEffect(() => {
    if (
      selectedClasses.length === 0 &&
      classOptions.length > 0 &&
      settings.reportPrintPreferences.selectedClasses.length === 0
    ) {
      setSelectedClasses(classOptions);
    }
  }, [classOptions]);

  useEffect(() => {
    const handleOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setClassDropdownOpen(false);
      }
      if (reportDropdownRef.current && !reportDropdownRef.current.contains(e.target as Node)) {
        setReportDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutside);
    return () => document.removeEventListener('mousedown', handleOutside);
  }, []);

  // Inject @page size & orientation dynamically for window.print()
  useEffect(() => {
    const styleId = 'portal-eguru-print-page-style';
    let styleEl = document.getElementById(styleId) as HTMLStyleElement | null;
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = styleId;
      document.head.appendChild(styleEl);
    }
    const cfg = PAPER_SIZE_CONFIG[paperSize];
    styleEl.textContent = `@media print { @page { size: ${cfg.cssPageSize} ${orientation}; margin: 12mm; } }`;
  }, [orientation, paperSize]);

  // Free drag handler inside signature zone (always active even from preset layout)
  const handleSigPointerDown = (
    target: 'kepsek' | 'guru',
    e: React.PointerEvent<HTMLDivElement>
  ) => {
    e.currentTarget.setPointerCapture(e.pointerId);
    const activeCoords =
      sigPos.preset === 'bebas'
        ? {
            principalX: sigPos.principalX,
            principalY: sigPos.principalY,
            teacherX: sigPos.teacherX,
            teacherY: sigPos.teacherY,
          }
        : getPresetCoords(sigPos.preset);

    if (sigPos.preset !== 'bebas') {
      updateSigPos({ preset: 'bebas', ...activeCoords });
    }

    setDragTarget(target);
    dragOriginRef.current = {
      startX: e.clientX,
      startY: e.clientY,
      initX: target === 'kepsek' ? activeCoords.principalX : activeCoords.teacherX,
      initY: target === 'kepsek' ? activeCoords.principalY : activeCoords.teacherY,
    };
  };

  const handleSigPointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    if (!dragTarget || !dragOriginRef.current || !sigCanvasRef.current) return;
    const rect = sigCanvasRef.current.getBoundingClientRect();
    const dxPct = ((e.clientX - dragOriginRef.current.startX) / Math.max(rect.width, 1)) * 100;
    const dyPx = e.clientY - dragOriginRef.current.startY;
    const nextX = Math.max(0, Math.min(76, Math.round(dragOriginRef.current.initX + dxPct)));
    const nextY = Math.max(0, Math.min(160, Math.round(dragOriginRef.current.initY + dyPx)));

    if (dragTarget === 'kepsek') {
      updateSigPos({ preset: 'bebas', principalX: nextX, principalY: nextY });
    } else {
      updateSigPos({ preset: 'bebas', teacherX: nextX, teacherY: nextY });
    }
  };

  const handleSigPointerUp = () => {
    setDragTarget(null);
    dragOriginRef.current = null;
  };

  const applyDatePreset = (preset: 'thisMonth' | 'lastMonth' | 'reset') => {
    if (preset === 'reset') {
      setStartDate('');
      setEndDate('');
      return;
    }
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth();
    if (preset === 'thisMonth') {
      const first = new Date(year, month, 1);
      const last = new Date(year, month + 1, 0);
      setStartDate(first.toISOString().slice(0, 10));
      setEndDate(last.toISOString().slice(0, 10));
    } else {
      const first = new Date(year, month - 1, 1);
      const last = new Date(year, month, 0);
      setStartDate(first.toISOString().slice(0, 10));
      setEndDate(last.toISOString().slice(0, 10));
    }
  };

  const toggleClassSelection = (cls: string) => {
    setSelectedClasses((prev) =>
      prev.includes(cls) ? prev.filter((c) => c !== cls) : [...prev, cls]
    );
  };

  const toggleReportTypeSelection = (id: ReportSectionId) => {
    setSelectedReportTypes((prev) => {
      const nextSet = new Set(prev);
      if (nextSet.has(id)) {
        nextSet.delete(id);
      } else {
        nextSet.add(id);
      }
      return REPORT_TYPE_OPTIONS.map((o) => o.id).filter((optId) => nextSet.has(optId));
    });
  };

  const filteredJournals = useMemo(() => {
    const clsSet = new Set(selectedClasses);
    return journals
      .filter((j) => {
        if (j.academicYear !== session.academicYear || j.semester !== session.semester) {
          return false;
        }
        if (!clsSet.has(j.className)) return false;
        if (subjectFilter !== 'ALL' && j.subject !== subjectFilter) return false;
        if (startDate && j.date < startDate) return false;
        if (endDate && j.date > endDate) return false;
        return true;
      })
      .sort((a, b) =>
        a.date > b.date ? 1 : a.date < b.date ? -1 : a.className.localeCompare(b.className)
      );
  }, [
    journals,
    session.academicYear,
    session.semester,
    selectedClasses,
    subjectFilter,
    startDate,
    endDate,
  ]);

  const activeGradeSubjects = useMemo(() => {
    if (subjectFilter !== 'ALL') return [subjectFilter];
    return settings.masterSubjects;
  }, [subjectFilter, settings.masterSubjects]);

  const getAttendanceBlockForClass = (cls: string) => {
    const classStudents = students
      .filter((s) => s.className === cls)
      .sort((a, b) => a.name.localeCompare(b.name));
    const classJournals = filteredJournals.filter((j) => j.className === cls);
    return { classStudents, classJournals };
  };

  const getGradeBlockForClassAndSubject = (cls: string, subj: string) => {
    const classStudents = students
      .filter((s) => s.className === cls)
      .sort((a, b) => a.name.localeCompare(b.name));
    const rows = classStudents.map((st, idx) => {
      const scores: Record<string, number | null> = {};
      const nums: number[] = [];
      settings.masterExamTypes.forEach((ex) => {
        const id = makeGradeId(session.academicYear, session.semester, cls, subj, ex, st.id);
        const found = grades.find((g) => g.id === id);
        const val = found ? found.score : null;
        scores[ex] = val;
        if (val !== null) nums.push(val);
      });
      const avg = nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
      return {
        no: idx + 1,
        nis: st.nis,
        name: st.name,
        scores,
        avg,
        status: avg === null ? '-' : avg >= 75 ? 'Tuntas' : 'Remedial',
      };
    });
    return { classStudents, rows };
  };

  // 3. Monthly Attendance Aggregator
  const getMonthlyAttendanceForClass = (cls: string) => {
    const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
    const months = Array.from(
      new Set(classJournals.map((j) => j.date.slice(0, 7)).filter(Boolean))
    ).sort();

    const rows = classStudents.map((st, idx) => {
      const monthCounts: Record<string, { h: number; total: number }> = {};
      months.forEach((m) => {
        monthCounts[m] = { h: 0, total: 0 };
      });
      const totals: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
      classJournals.forEach((j) => {
        const ym = j.date.slice(0, 7);
        const code = j.attendance[st.id]?.status ?? 'H';
        totals[code] += 1;
        if (monthCounts[ym]) {
          monthCounts[ym].total += 1;
          if (code === 'H') monthCounts[ym].h += 1;
        }
      });
      const pct =
        classJournals.length > 0 ? (totals.H / classJournals.length) * 100 : null;
      return {
        no: idx + 1,
        nis: st.nis,
        name: st.name,
        monthCounts,
        totals,
        totalMeetings: classJournals.length,
        pct,
      };
    });

    const gradedPcts = rows.map((r) => r.pct).filter((v): v is number => v !== null);
    const classAvgPct =
      gradedPcts.length > 0
        ? gradedPcts.reduce((a, b) => a + b, 0) / gradedPcts.length
        : null;

    return { months, rows, classJournals, classAvgPct };
  };

  // 4. Monitoring S/I/A/D (Summary + Chronology)
  const monitoringAbsenceData = useMemo(() => {
    const studentLookup = new Map(students.map((s) => [s.id, s]));
    const summaryMap = new Map<
      string,
      {
        studentId: string;
        nis: string;
        name: string;
        className: string;
        S: number;
        I: number;
        A: number;
        D: number;
      }
    >();
    const chronology: {
      date: string;
      jamKe: string;
      className: string;
      subject: string;
      nis: string;
      studentName: string;
      status: Exclude<AttendanceStatus, 'H'>;
      note: string;
    }[] = [];

    filteredJournals.forEach((j) => {
      Object.entries(j.attendance).forEach(([studentId, rec]) => {
        if (!rec || rec.status === 'H') return;
        const st = studentLookup.get(studentId);
        if (!st) return;
        const status = rec.status as Exclude<AttendanceStatus, 'H'>;
        chronology.push({
          date: j.date,
          jamKe: j.jamKe,
          className: j.className,
          subject: j.subject,
          nis: st.nis,
          studentName: st.name,
          status,
          note: rec.note || '-',
        });

        const cur = summaryMap.get(studentId) ?? {
          studentId,
          nis: st.nis,
          name: st.name,
          className: st.className,
          S: 0,
          I: 0,
          A: 0,
          D: 0,
        };
        cur[status] += 1;
        summaryMap.set(studentId, cur);
      });
    });

    const summary = Array.from(summaryMap.values())
      .map((item) => {
        const total = item.S + item.I + item.A + item.D;
        const action =
          item.A >= 3 || total >= 5
            ? 'Tindak Lanjut BK'
            : total >= 3
            ? 'Perlu Pantauan'
            : 'Pembinaan Ringan';
        return { ...item, total, action };
      })
      .sort((a, b) => b.total - a.total || b.A - a.A || a.name.localeCompare(b.name));

    chronology.sort((a, b) =>
      a.date < b.date ? 1 : a.date > b.date ? -1 : a.className.localeCompare(b.className)
    );

    return { summary, chronology };
  }, [filteredJournals, students]);

  // 6. Ketuntasan & Remedial (< KKM 75)
  const getRemedialDataForClassAndSubject = (cls: string, subj: string) => {
    const { rows } = getGradeBlockForClassAndSubject(cls, subj);
    const assessed = rows.filter((r) => r.avg !== null);
    const tuntasCount = assessed.filter((r) => (r.avg ?? 0) >= 75).length;
    const remedialList = assessed
      .filter((r) => (r.avg ?? 0) < 75)
      .map((r, idx) => {
        const vals = Object.values(r.scores).filter((v): v is number => v !== null);
        const minScore = vals.length > 0 ? Math.min(...vals) : 0;
        const maxScore = vals.length > 0 ? Math.max(...vals) : 0;
        const belowItems = settings.masterExamTypes
          .filter((ex) => r.scores[ex] !== null && (r.scores[ex] as number) < 75)
          .map((ex) => `${ex} (${r.scores[ex]})`)
          .join(', ');
        const avgVal = r.avg ?? 0;
        return {
          no: idx + 1,
          nis: r.nis,
          name: r.name,
          minScore,
          maxScore,
          belowItems: belowItems || '-',
          avg: avgVal,
          gap: Math.max(0, 75 - avgVal),
          rekomendasi: avgVal < 60 ? 'Remedial Terbimbing' : 'Perbaikan & Ujian Ulang',
        };
      });

    const classAvg =
      assessed.length > 0
        ? assessed.reduce((acc, r) => acc + (r.avg ?? 0), 0) / assessed.length
        : null;
    const dayaSerap = assessed.length > 0 ? (tuntasCount / assessed.length) * 100 : null;

    return {
      totalStudents: rows.length,
      assessedCount: assessed.length,
      tuntasCount,
      remedialCount: remedialList.length,
      dayaSerap,
      classAvg,
      remedialList,
    };
  };

  // 7. Leger Gabungan Kelas (Cross-Subject & Ranking)
  const getCombinedLegerForClass = (cls: string) => {
    const classStudents = students
      .filter((s) => s.className === cls)
      .sort((a, b) => a.name.localeCompare(b.name));

    const subjectsWithData = settings.masterSubjects.filter((subj) =>
      grades.some(
        (g) =>
          g.academicYear === session.academicYear &&
          g.semester === session.semester &&
          g.className === cls &&
          g.subject === subj &&
          g.score !== null
      )
    );
    const subjects =
      subjectsWithData.length > 0 ? subjectsWithData : settings.masterSubjects.slice(0, 5);

    const rawRows = classStudents.map((st) => {
      const subjAvgs: Record<string, number | null> = {};
      const validSubjAvgs: number[] = [];
      subjects.forEach((subj) => {
        const nums: number[] = [];
        settings.masterExamTypes.forEach((ex) => {
          const id = makeGradeId(session.academicYear, session.semester, cls, subj, ex, st.id);
          const found = grades.find((g) => g.id === id);
          if (found && found.score !== null) nums.push(found.score);
        });
        const sAvg = nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
        subjAvgs[subj] = sAvg;
        if (sAvg !== null) validSubjAvgs.push(sAvg);
      });
      const totalScore =
        validSubjAvgs.length > 0 ? validSubjAvgs.reduce((a, b) => a + b, 0) : null;
      const overallAvg =
        validSubjAvgs.length > 0 ? (totalScore as number) / validSubjAvgs.length : null;
      return {
        id: st.id,
        nis: st.nis,
        name: st.name,
        subjAvgs,
        totalScore,
        overallAvg,
        status: overallAvg === null ? '-' : overallAvg >= 75 ? 'Tuntas' : 'Remedial',
        rank: null as number | null,
      };
    });

    const sorted = [...rawRows].sort((a, b) => {
      if (a.overallAvg === null && b.overallAvg === null) return a.name.localeCompare(b.name);
      if (a.overallAvg === null) return 1;
      if (b.overallAvg === null) return -1;
      if (b.overallAvg !== a.overallAvg) return b.overallAvg - a.overallAvg;
      return a.name.localeCompare(b.name);
    });

    let currentRank = 0;
    sorted.forEach((r) => {
      if (r.overallAvg !== null) {
        currentRank += 1;
        r.rank = currentRank;
      }
    });

    return { subjects, rows: sorted };
  };

  // 8. Daftar Peserta Didik (Kesiswaan)
  const studentRosterRows = useMemo(() => {
    const clsSet = new Set(selectedClasses);
    return students
      .filter((s) => clsSet.has(s.className))
      .sort(
        (a, b) =>
          a.className.localeCompare(b.className) ||
          (Number(a.absentNo) || 999) - (Number(b.absentNo) || 999) ||
          a.name.localeCompare(b.name)
      );
  }, [students, selectedClasses]);

  // 9. Aktivitas & Penilaian Guru (Pimpinan)
  const teacherActivityData = useMemo(() => {
    const periodSchedules = schedules.filter(
      (sc) => sc.academicYear === session.academicYear && sc.semester === session.semester
    );
    const periodJournals = journals.filter((j) => {
      if (j.academicYear !== session.academicYear || j.semester !== session.semester) return false;
      if (startDate && j.date < startDate) return false;
      if (endDate && j.date > endDate) return false;
      return true;
    });
    const periodGrades = grades.filter(
      (g) =>
        g.academicYear === session.academicYear &&
        g.semester === session.semester &&
        g.score !== null
    );

    const teacherSource =
      teachers.length > 0
        ? teachers
        : [
            {
              id: 'signer-fallback',
              nip: activeSigner.nip || '-',
              name: activeSigner.name,
              position: activeSigner.roleLabel,
              subject: settings.masterSubjects[0] || 'Umum',
              email: session.email,
              phone: '-',
              gender: 'L' as const,
              employmentStatus: 'PNS' as const,
              status: 'Aktif' as const,
              address: '-',
              photo: '',
            },
          ];

    const summary = teacherSource.map((t, idx) => {
      const tNorm = t.name.trim().toLowerCase();
      const tSchedules = periodSchedules.filter(
        (sc) => sc.teacherName.trim().toLowerCase() === tNorm
      );
      const classSubjPairs = new Set(tSchedules.map((sc) => `${sc.className}__${sc.subject}`));

      const matchedJournals = periodJournals.filter((j) => {
        if (classSubjPairs.size > 0) {
          return classSubjPairs.has(`${j.className}__${j.subject}`);
        }
        return t.subject && j.subject.toLowerCase() === t.subject.toLowerCase();
      });

      let totalH = 0;
      let totalRec = 0;
      matchedJournals.forEach((j) => {
        const s = summarizeJournalRow(j);
        totalH += s.counts.H;
        totalRec += s.total;
      });
      const avgAttendance = totalRec > 0 ? (totalH / totalRec) * 100 : null;

      const matchedGrades = periodGrades.filter((g) => {
        if (classSubjPairs.size > 0) {
          return classSubjPairs.has(`${g.className}__${g.subject}`);
        }
        return t.subject && g.subject.toLowerCase() === t.subject.toLowerCase();
      });
      const assessedStudentIds = new Set(matchedGrades.map((g) => g.studentId));
      const avgGrade =
        matchedGrades.length > 0
          ? matchedGrades.reduce((acc, g) => acc + (g.score ?? 0), 0) / matchedGrades.length
          : null;

      const kinerja =
        matchedJournals.length >= 4
          ? 'Sangat Aktif'
          : matchedJournals.length >= 1
          ? 'Aktif'
          : tSchedules.length > 0
          ? 'Terjadwal'
          : 'Belum Ada Jurnal';

      return {
        no: idx + 1,
        nip: t.nip || '-',
        name: t.name,
        position: t.position || 'Guru Mapel',
        subject: t.subject || '-',
        scheduleCount: tSchedules.length,
        journalCount: matchedJournals.length,
        avgAttendance,
        assessedStudents: assessedStudentIds.size,
        avgGrade,
        kinerja,
      };
    });

    const recentKbm = [...filteredJournals]
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0))
      .slice(0, 25)
      .map((j, idx) => {
        const s = summarizeJournalRow(j);
        const pct = s.total > 0 ? (s.counts.H / s.total) * 100 : 100;
        return {
          no: idx + 1,
          date: j.date,
          jamKe: j.jamKe,
          className: j.className,
          subject: j.subject,
          topic: j.topic,
          pct,
        };
      });

    return { summary, recentKbm };
  }, [
    teachers,
    schedules,
    journals,
    grades,
    filteredJournals,
    session.academicYear,
    session.semester,
    session.email,
    startDate,
    endDate,
    activeSigner.name,
    activeSigner.nip,
    activeSigner.roleLabel,
    settings.masterSubjects,
  ]);

  const reportTitle = useMemo(() => {
    if (selectedReportTypes.length === 0) return 'LAPORAN ADMINISTRASI MENGAJAR';
    if (selectedReportTypes.length === REPORT_TYPE_OPTIONS.length) {
      return 'PAKET LENGKAP LAPORAN PORTAL E-GURU';
    }
    if (selectedReportTypes.length === 1) {
      const opt = REPORT_TYPE_OPTIONS.find((o) => o.id === selectedReportTypes[0]);
      return opt ? opt.officialTitle : 'LAPORAN ADMINISTRASI MENGAJAR';
    }
    const labels = REPORT_TYPE_OPTIONS.filter((o) => selectedReportTypes.includes(o.id)).map((o) =>
      o.label.toUpperCase()
    );
    if (labels.length <= 3) return `LAPORAN TERPADU: ${labels.join(' · ')}`;
    return `LAPORAN TERPADU (${labels.length} JENIS): ${labels.slice(0, 2).join(' · ')} & Lainnya`;
  }, [selectedReportTypes]);

  const reportSlug = useMemo(() => {
    if (selectedReportTypes.length === REPORT_TYPE_OPTIONS.length) return 'paket_lengkap';
    if (selectedReportTypes.length === 0) return 'kosong';
    return selectedReportTypes.slice(0, 3).join('_');
  }, [selectedReportTypes]);

  const periodSubtitle = `Tahun Ajaran ${session.academicYear} · Semester ${session.semester} · Mata Pelajaran: ${
    subjectFilter === 'ALL' ? 'Semua Mata Pelajaran' : subjectFilter
  }`;

  const dateRangeLabel =
    startDate || endDate
      ? `${startDate ? formatShortDateId(startDate) : 'Awal'} s.d. ${
          endDate ? formatShortDateId(endDate) : 'Sekarang'
        }`
      : 'Seluruh Tanggal Pertemuan';

  // Build Workbook for XLSX / XLS
  const buildWorkbook = () => {
    const wb = XLSX.utils.book_new();
    const has = (id: ReportSectionId) => selectedReportTypes.includes(id);

    if (has('journal')) {
      const jRows: (string | number)[][] = [
        [
          'No',
          'Tanggal',
          'Jam Ke',
          'Kelas',
          'Mata Pelajaran',
          'Materi Pokok',
          'Kegiatan Pembelajaran',
          'H',
          'S',
          'I',
          'A',
          'D',
        ],
        ...filteredJournals.map((j, idx) => {
          const sum = summarizeJournalRow(j);
          return [
            idx + 1,
            formatShortDateId(j.date),
            j.jamKe,
            j.className,
            j.subject,
            j.topic,
            j.activity,
            sum.counts.H,
            sum.counts.S,
            sum.counts.I,
            sum.counts.A,
            sum.counts.D,
          ];
        }),
      ];
      const ws = XLSX.utils.aoa_to_sheet(jRows);
      XLSX.utils.book_append_sheet(wb, ws, 'Jurnal Mengajar');
    }

    if (has('attendance')) {
      selectedClasses.forEach((cls) => {
        const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
        const header = [
          'No',
          'NIS',
          'Nama Siswa',
          ...classJournals.map((j) => `${formatCompactDateId(j.date)} (${j.subject})`),
          'H',
          'S',
          'I',
          'A',
          'D',
          '% Hadir',
        ];
        const body = classStudents.map((st, idx) => {
          const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
          const cells = classJournals.map((j) => {
            const stCode = j.attendance[st.id]?.status ?? 'H';
            counts[stCode] += 1;
            return stCode;
          });
          const pct =
            classJournals.length > 0
              ? `${((counts.H / classJournals.length) * 100).toFixed(0)}%`
              : '-';
          return [
            idx + 1,
            st.nis,
            st.name,
            ...cells,
            counts.H,
            counts.S,
            counts.I,
            counts.A,
            counts.D,
            pct,
          ];
        });
        const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
        XLSX.utils.book_append_sheet(wb, ws, `Absen ${cls}`.slice(0, 31));
      });
    }

    if (has('attendance_monthly')) {
      selectedClasses.forEach((cls) => {
        const { months, rows } = getMonthlyAttendanceForClass(cls);
        const header = [
          'No',
          'NIS',
          'Nama Siswa',
          ...months.map((m) => formatMonthLabelId(m)),
          'Pertemuan',
          'H',
          'S',
          'I',
          'A',
          'D',
          '% Hadir',
        ];
        const body = rows.map((r) => [
          r.no,
          r.nis,
          r.name,
          ...months.map((m) => {
            const mc = r.monthCounts[m];
            return mc && mc.total > 0 ? `${mc.h}/${mc.total}` : '-';
          }),
          r.totalMeetings,
          r.totals.H,
          r.totals.S,
          r.totals.I,
          r.totals.A,
          r.totals.D,
          r.pct !== null ? `${r.pct.toFixed(0)}%` : '-',
        ]);
        const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
        XLSX.utils.book_append_sheet(wb, ws, `Presensi Bln ${cls}`.slice(0, 31));
      });
    }

    if (has('attendance_monitoring')) {
      const sumHeader = ['No', 'NIS', 'Nama Siswa', 'Kelas', 'S', 'I', 'A', 'D', 'Total', 'Status Pembinaan'];
      const sumBody = monitoringAbsenceData.summary.map((item, idx) => [
        idx + 1,
        item.nis,
        item.name,
        item.className,
        item.S,
        item.I,
        item.A,
        item.D,
        item.total,
        item.action,
      ]);
      XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.aoa_to_sheet([sumHeader, ...sumBody]),
        'Rekap SIAD'
      );

      const logHeader = ['No', 'Tanggal', 'Jam', 'Kelas', 'Mapel', 'NIS', 'Nama Siswa', 'Status', 'Keterangan'];
      const logBody = monitoringAbsenceData.chronology.map((c, idx) => [
        idx + 1,
        formatShortDateId(c.date),
        c.jamKe,
        c.className,
        c.subject,
        c.nis,
        c.studentName,
        c.status,
        c.note,
      ]);
      XLSX.utils.book_append_sheet(
        wb,
        XLSX.utils.aoa_to_sheet([logHeader, ...logBody]),
        'Kronologi SIAD'
      );
    }

    if (has('grades')) {
      selectedClasses.forEach((cls) => {
        activeGradeSubjects.forEach((subj) => {
          const { rows } = getGradeBlockForClassAndSubject(cls, subj);
          const hasAnyScore = rows.some((r) => r.avg !== null);
          if (!hasAnyScore && subjectFilter === 'ALL') return;
          const header = [
            'No',
            'NIS',
            'Nama Siswa',
            ...settings.masterExamTypes,
            'Rata-rata',
            'Status',
          ];
          const body = rows.map((r) => [
            r.no,
            r.nis,
            r.name,
            ...settings.masterExamTypes.map((ex) => (r.scores[ex] !== null ? r.scores[ex] : '')),
            r.avg !== null ? Number(r.avg.toFixed(1)) : '',
            r.status,
          ]);
          const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
          const safeSheetName = `Nilai ${cls} ${subj}`.replace(/[\\/?*[\]]/g, '').slice(0, 31);
          XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
        });
      });
    }

    if (has('grades_remedial')) {
      selectedClasses.forEach((cls) => {
        activeGradeSubjects.forEach((subj) => {
          const rem = getRemedialDataForClassAndSubject(cls, subj);
          if (rem.assessedCount === 0 && subjectFilter === 'ALL') return;
          const header = [
            'No',
            'NIS',
            'Nama Siswa',
            'Tertinggi',
            'Terendah',
            'Komponen < KKM',
            'Rata-rata',
            'Selisih KKM',
            'Rekomendasi',
          ];
          const body = rem.remedialList.map((r) => [
            r.no,
            r.nis,
            r.name,
            r.maxScore,
            r.minScore,
            r.belowItems,
            Number(r.avg.toFixed(1)),
            Number(r.gap.toFixed(1)),
            r.rekomendasi,
          ]);
          const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
          const safeSheetName = `Remed ${cls} ${subj}`.replace(/[\\/?*[\]]/g, '').slice(0, 31);
          XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
        });
      });
    }

    if (has('grades_leger')) {
      selectedClasses.forEach((cls) => {
        const { subjects, rows } = getCombinedLegerForClass(cls);
        const header = ['Peringkat', 'NIS', 'Nama Siswa', ...subjects, 'Jumlah', 'Rata-rata', 'Status'];
        const body = rows.map((r) => [
          r.rank ?? '-',
          r.nis,
          r.name,
          ...subjects.map((s) => (r.subjAvgs[s] !== null ? Number((r.subjAvgs[s] as number).toFixed(1)) : '')),
          r.totalScore !== null ? Number(r.totalScore.toFixed(1)) : '',
          r.overallAvg !== null ? Number(r.overallAvg.toFixed(1)) : '',
          r.status,
        ]);
        const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
        XLSX.utils.book_append_sheet(wb, ws, `Leger ${cls}`.slice(0, 31));
      });
    }

    if (has('students_roster')) {
      const header = [
        'No',
        'NIS',
        'NISN',
        'Nama Siswa',
        'L/P',
        'Kelas',
        'No Absen',
        'Tempat, Tgl Lahir',
        'No. HP',
        'Wali Kelas',
        'Guru BK',
        'Status',
      ];
      const body = studentRosterRows.map((st, idx) => [
        idx + 1,
        st.nis,
        st.nisn,
        st.name,
        st.gender,
        st.className,
        st.absentNo || '-',
        `${st.birthPlace || '-'}, ${st.birthDate || '-'}`,
        st.phone || '-',
        st.homeroomTeacher || '-',
        st.counselorTeacher || '-',
        st.status,
      ]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([header, ...body]), 'Daftar Siswa');
    }

    if (has('teacher_activity')) {
      const header = [
        'No',
        'NIP',
        'Nama Guru',
        'Jabatan / Mapel',
        'Jadwal/Mgg',
        'Jurnal Terisi',
        '% Hadir Siswa',
        'Siswa Dinilai',
        'Rata Nilai',
        'Status Kinerja',
      ];
      const body = teacherActivityData.summary.map((t) => [
        t.no,
        t.nip,
        t.name,
        `${t.position} · ${t.subject}`,
        t.scheduleCount,
        t.journalCount,
        t.avgAttendance !== null ? `${t.avgAttendance.toFixed(0)}%` : '-',
        t.assessedStudents,
        t.avgGrade !== null ? Number(t.avgGrade.toFixed(1)) : '-',
        t.kinerja,
      ]);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([header, ...body]), 'Kinerja Guru');
    }

    if (wb.SheetNames.length === 0) {
      const ws = XLSX.utils.aoa_to_sheet([['Tidak ada data untuk filter terpilih']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Laporan');
    }

    return wb;
  };

  const handleDownloadXlsx = () => {
    if (selectedReportTypes.length === 0) {
      onToast('error', 'Pilih minimal satu jenis laporan terlebih dahulu.');
      return;
    }
    const wb = buildWorkbook();
    XLSX.writeFile(wb, `Laporan_E-Guru_${reportSlug}_${getTodayISO()}.xlsx`);
    onToast('success', 'Laporan Excel (.xlsx) berhasil diunduh.');
  };

  const handleDownloadXls = () => {
    if (selectedReportTypes.length === 0) {
      onToast('error', 'Pilih minimal satu jenis laporan terlebih dahulu.');
      return;
    }
    const wb = buildWorkbook();
    XLSX.writeFile(wb, `Laporan_E-Guru_${reportSlug}_${getTodayISO()}.xls`, {
      bookType: 'biff8',
    });
    onToast('success', 'Laporan Excel (.xls) berhasil diunduh.');
  };

  const handleDownloadWord = async () => {
    if (selectedReportTypes.length === 0) {
      onToast('error', 'Pilih minimal satu jenis laporan terlebih dahulu.');
      return;
    }
    if (!previewRef.current) return;
    setIsExportingWord(true);
    try {
      await exportPreviewDomToWordDoc({
        previewElement: previewRef.current,
        title: reportTitle,
        filename: `Laporan_E-Guru_${reportSlug}_${getTodayISO()}.doc`,
        orientation,
        cssPageSize: PAPER_SIZE_CONFIG[paperSize].cssPageSize,
        fontFamily,
      });
      onToast('success', 'Dokumen Word (.doc) persis pratinjau berhasil diunduh.');
    } catch {
      onToast('error', 'Gagal mengekspor ke Word (.doc).');
    } finally {
      setIsExportingWord(false);
    }
  };

  // Vector PDF Generator matching Live Preview 1:1 (including signatures, stamp & signaturePosition)
  const handleDownloadPdf = async () => {
    if (selectedReportTypes.length === 0) {
      onToast('error', 'Pilih minimal satu jenis laporan terlebih dahulu.');
      return;
    }
    setIsGeneratingPdf(true);
    try {
      const paperCfg = PAPER_SIZE_CONFIG[paperSize];
      const doc = new jsPDF({
        orientation,
        unit: 'mm',
        format: paperCfg.jsPdfFormat,
      });

      const pdfFont =
        fontFamily === 'Times New Roman' || fontFamily === 'Georgia'
          ? 'times'
          : 'helvetica';
      doc.setFont(pdfFont, 'normal');

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const marginX = 14;
      let cursorY = 14;

      const [leftLogoPng, rightLogoPng, kepsekSigPng, guruSigPng, stampPng] =
        await Promise.all([
          showKop ? imageUrlToPngDataUrl(settings.kopSurat.logoLeft) : Promise.resolve(null),
          showKop ? imageUrlToPngDataUrl(settings.kopSurat.logoRight) : Promise.resolve(null),
          showTtdKepsek && showSignatureImage
            ? imageUrlToPngDataUrl(settings.principalSignatureImage)
            : Promise.resolve(null),
          showTtdGuru && showSignatureImage
            ? imageUrlToPngDataUrl(activeSigner.signatureImage)
            : Promise.resolve(null),
          sigPos.showStamp
            ? imageUrlToPngDataUrl(activeSigner.schoolStampImage)
            : Promise.resolve(null),
        ]);

      const drawKopSurat = () => {
        if (!showKop) return;
        const logoSize = 20;
        if (leftLogoPng) {
          doc.addImage(leftLogoPng, 'PNG', marginX, cursorY, logoSize, logoSize);
        }
        if (rightLogoPng) {
          doc.addImage(
            rightLogoPng,
            'PNG',
            pageWidth - marginX - logoSize,
            cursorY,
            logoSize,
            logoSize
          );
        }

        const centerX = pageWidth / 2;
        doc.setFont(pdfFont, 'normal');
        doc.setFontSize(10);
        if (settings.kopSurat.line1) {
          doc.text(settings.kopSurat.line1, centerX, cursorY + 4, { align: 'center' });
        }
        if (settings.kopSurat.line2) {
          doc.setFont(pdfFont, 'bold');
          doc.setFontSize(11);
          doc.text(settings.kopSurat.line2, centerX, cursorY + 9, { align: 'center' });
        }
        if (settings.kopSurat.line3) {
          doc.setFont(pdfFont, 'bold');
          doc.setFontSize(13);
          doc.text(settings.kopSurat.line3, centerX, cursorY + 14.5, { align: 'center' });
        }
        doc.setFont(pdfFont, 'normal');
        doc.setFontSize(8.5);
        const contactLines = [
          settings.kopSurat.line4,
          settings.kopSurat.line5,
          settings.kopSurat.line6,
        ].filter(Boolean);
        contactLines.forEach((line, idx) => {
          doc.text(line, centerX, cursorY + 19 + idx * 3.8, { align: 'center' });
        });

        const lineY = cursorY + 19 + Math.max(1, contactLines.length) * 3.8 + 1;
        doc.setLineWidth(0.8);
        doc.line(marginX, lineY, pageWidth - marginX, lineY);
        doc.setLineWidth(0.25);
        doc.line(marginX, lineY + 1, pageWidth - marginX, lineY + 1);
        cursorY = lineY + 7;
      };

      drawKopSurat();

      // Title
      doc.setFont(pdfFont, 'bold');
      doc.setFontSize(12);
      doc.text(reportTitle, pageWidth / 2, cursorY, { align: 'center' });
      cursorY += 5;

      doc.setFont(pdfFont, 'normal');
      doc.setFontSize(9);
      doc.text(`${periodSubtitle} · Rentang Tanggal: ${dateRangeLabel}`, pageWidth / 2, cursorY, {
        align: 'center',
      });
      cursorY += 7;

      const commonTableStyles = {
        font: pdfFont,
        fontSize: 8,
        cellPadding: 2,
        lineColor: [0, 0, 0] as [number, number, number],
        lineWidth: 0.2,
      };

      const ensureSpace = (minMm = 36) => {
        if (cursorY > pageHeight - minMm) {
          doc.addPage();
          cursorY = 16;
        }
      };

      const drawSectionTitle = (label: string) => {
        ensureSpace(36);
        doc.setFont(pdfFont, 'bold');
        doc.setFontSize(10);
        doc.text(label, marginX, cursorY);
        cursorY += 3;
      };

      const has = (id: ReportSectionId) => selectedReportTypes.includes(id);

      if (has('journal')) {
        if (selectedReportTypes.length > 1) {
          drawSectionTitle('Jurnal Mengajar (Mengajar) — Catatan Materi, Jam & Kegiatan KBM');
        }
        autoTable(doc, {
          startY: cursorY,
          margin: { left: marginX, right: marginX },
          styles: commonTableStyles,
          headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
          head: [
            [
              'No',
              'Tanggal',
              'Jam',
              'Kelas',
              'Mata Pelajaran',
              'Materi Pokok',
              'Kegiatan Pembelajaran',
              'H/S/I/A/D',
            ],
          ],
          body:
            filteredJournals.length === 0
              ? [['-', '-', '-', '-', '-', 'Belum ada data jurnal pada filter ini', '-', '-']]
              : filteredJournals.map((j, idx) => {
                  const sum = summarizeJournalRow(j);
                  return [
                    idx + 1,
                    formatShortDateId(j.date),
                    j.jamKe,
                    j.className,
                    j.subject,
                    j.topic,
                    j.activity,
                    `${sum.counts.H}/${sum.counts.S}/${sum.counts.I}/${sum.counts.A}/${sum.counts.D}`,
                  ];
                }),
        });
        cursorY =
          (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
      }

      if (has('attendance')) {
        selectedClasses.forEach((cls) => {
          const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
          drawSectionTitle(`Kehadiran Mendatar — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: { ...commonTableStyles, fontSize: 7.5 },
            headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
            head: [
              [
                'No',
                'NIS',
                'Nama Siswa',
                ...classJournals.map((j) => formatCompactDateId(j.date)),
                'H',
                'S',
                'I',
                'A',
                'D',
                '%',
              ],
            ],
            body:
              classStudents.length === 0
                ? [['-', '-', 'Belum ada siswa di kelas ini']]
                : classStudents.map((st, idx) => {
                    const counts: Record<AttendanceStatus, number> = {
                      H: 0,
                      S: 0,
                      I: 0,
                      A: 0,
                      D: 0,
                    };
                    const cells = classJournals.map((j) => {
                      const code = j.attendance[st.id]?.status ?? 'H';
                      counts[code] += 1;
                      return code;
                    });
                    const pct =
                      classJournals.length > 0
                        ? `${((counts.H / classJournals.length) * 100).toFixed(0)}%`
                        : '-';
                    return [
                      idx + 1,
                      st.nis,
                      st.name,
                      ...cells,
                      counts.H,
                      counts.S,
                      counts.I,
                      counts.A,
                      counts.D,
                      pct,
                    ];
                  }),
          });
          cursorY =
            (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
        });
      }

      if (has('attendance_monthly')) {
        selectedClasses.forEach((cls) => {
          const { months, rows } = getMonthlyAttendanceForClass(cls);
          drawSectionTitle(`Rekap Presensi Bulanan — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: commonTableStyles,
            headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
            head: [
              [
                'No',
                'NIS',
                'Nama Siswa',
                ...months.map((m) => formatMonthLabelId(m)),
                'Pertemuan',
                'H',
                'S',
                'I',
                'A',
                'D',
                '% Hadir',
              ],
            ],
            body:
              rows.length === 0
                ? [['-', '-', 'Belum ada siswa di kelas ini']]
                : rows.map((r) => [
                    r.no,
                    r.nis,
                    r.name,
                    ...months.map((m) => {
                      const mc = r.monthCounts[m];
                      return mc && mc.total > 0 ? `${mc.h}/${mc.total}` : '-';
                    }),
                    r.totalMeetings,
                    r.totals.H,
                    r.totals.S,
                    r.totals.I,
                    r.totals.A,
                    r.totals.D,
                    r.pct !== null ? `${r.pct.toFixed(0)}%` : '-',
                  ]),
          });
          cursorY =
            (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
        });
      }

      if (has('attendance_monitoring')) {
        drawSectionTitle('Monitoring Ketidakhadiran (S/I/A/D) — Rekap Siswa');
        autoTable(doc, {
          startY: cursorY,
          margin: { left: marginX, right: marginX },
          styles: commonTableStyles,
          headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
          head: [['No', 'NIS', 'Nama Siswa', 'Kelas', 'S', 'I', 'A', 'D', 'Total', 'Pembinaan']],
          body:
            monitoringAbsenceData.summary.length === 0
              ? [['-', '-', 'Seluruh siswa hadir 100% pada periode ini', '-', 0, 0, 0, 0, 0, '-']]
              : monitoringAbsenceData.summary.map((it, idx) => [
                  idx + 1,
                  it.nis,
                  it.name,
                  it.className,
                  it.S,
                  it.I,
                  it.A,
                  it.D,
                  it.total,
                  it.action,
                ]),
        });
        cursorY =
          (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 6;

        drawSectionTitle('Kronologi Ketidakhadiran (S/I/A/D)');
        autoTable(doc, {
          startY: cursorY,
          margin: { left: marginX, right: marginX },
          styles: commonTableStyles,
          headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
          head: [['No', 'Tanggal', 'Jam', 'Kelas', 'Mapel', 'Nama Siswa', 'Status', 'Keterangan']],
          body:
            monitoringAbsenceData.chronology.length === 0
              ? [['-', '-', '-', '-', '-', 'Belum ada catatan ketidakhadiran', '-', '-']]
              : monitoringAbsenceData.chronology.map((c, idx) => [
                  idx + 1,
                  formatShortDateId(c.date),
                  c.jamKe,
                  c.className,
                  c.subject,
                  c.studentName,
                  c.status,
                  c.note,
                ]),
        });
        cursorY =
          (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
      }

      if (has('grades')) {
        selectedClasses.forEach((cls) => {
          activeGradeSubjects.forEach((subj) => {
            const { rows } = getGradeBlockForClassAndSubject(cls, subj);
            const hasAnyScore = rows.some((r) => r.avg !== null);
            if (!hasAnyScore && subjectFilter === 'ALL') return;

            drawSectionTitle(`Rekap Nilai Mapel — Kelas ${cls} (${subj})`);
            autoTable(doc, {
              startY: cursorY,
              margin: { left: marginX, right: marginX },
              styles: commonTableStyles,
              headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
              head: [['No', 'NIS', 'Nama Siswa', ...settings.masterExamTypes, 'Rata-rata', 'Status']],
              body: rows.map((r) => [
                r.no,
                r.nis,
                r.name,
                ...settings.masterExamTypes.map((ex) =>
                  r.scores[ex] !== null ? String(r.scores[ex]) : '-'
                ),
                r.avg !== null ? r.avg.toFixed(1) : '-',
                r.status,
              ]),
            });
            cursorY =
              (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
          });
        });
      }

      if (has('grades_remedial')) {
        selectedClasses.forEach((cls) => {
          activeGradeSubjects.forEach((subj) => {
            const rem = getRemedialDataForClassAndSubject(cls, subj);
            if (rem.assessedCount === 0 && subjectFilter === 'ALL') return;
            const dayaLabel = rem.dayaSerap !== null ? `${rem.dayaSerap.toFixed(0)}%` : '-';
            drawSectionTitle(
              `Ketuntasan & Remedial (< KKM 75) — ${cls} (${subj}) · Daya Serap: ${dayaLabel} (${rem.tuntasCount}/${rem.assessedCount})`
            );
            autoTable(doc, {
              startY: cursorY,
              margin: { left: marginX, right: marginX },
              styles: commonTableStyles,
              headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
              head: [
                [
                  'No',
                  'NIS',
                  'Nama Siswa',
                  'Max',
                  'Min',
                  'Komponen < KKM',
                  'Rata-rata',
                  'Selisih',
                  'Rekomendasi',
                ],
              ],
              body:
                rem.remedialList.length === 0
                  ? [['-', '-', 'Seluruh siswa yang dinilai sudah tuntas (>= KKM 75)', '-', '-', '-', '-', '-', '-']]
                  : rem.remedialList.map((r) => [
                      r.no,
                      r.nis,
                      r.name,
                      r.maxScore,
                      r.minScore,
                      r.belowItems,
                      r.avg.toFixed(1),
                      `-${r.gap.toFixed(1)}`,
                      r.rekomendasi,
                    ]),
            });
            cursorY =
              (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
          });
        });
      }

      if (has('grades_leger')) {
        selectedClasses.forEach((cls) => {
          const { subjects, rows } = getCombinedLegerForClass(cls);
          drawSectionTitle(`Leger Gabungan Kelas & Peringkat — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: { ...commonTableStyles, fontSize: 7.5 },
            headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
            head: [['Rank', 'NIS', 'Nama Siswa', ...subjects, 'Jumlah', 'Rata-rata', 'Status']],
            body:
              rows.length === 0
                ? [['-', '-', `Belum ada siswa di kelas ${cls}`]]
                : rows.map((r) => [
                    r.rank ?? '-',
                    r.nis,
                    r.name,
                    ...subjects.map((s) =>
                      r.subjAvgs[s] !== null ? (r.subjAvgs[s] as number).toFixed(1) : '-'
                    ),
                    r.totalScore !== null ? r.totalScore.toFixed(1) : '-',
                    r.overallAvg !== null ? r.overallAvg.toFixed(1) : '-',
                    r.status,
                  ]),
          });
          cursorY =
            (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
        });
      }

      if (has('students_roster')) {
        drawSectionTitle(`Daftar Peserta Didik & Wali Kelas (${studentRosterRows.length} Siswa)`);
        autoTable(doc, {
          startY: cursorY,
          margin: { left: marginX, right: marginX },
          styles: { ...commonTableStyles, fontSize: 7.5 },
          headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
          head: [
            [
              'No',
              'NIS / NISN',
              'Nama Siswa',
              'L/P',
              'Kelas',
              'Abs',
              'TTL',
              'No. HP',
              'Wali Kelas',
              'Guru BK',
              'Status',
            ],
          ],
          body:
            studentRosterRows.length === 0
              ? [['-', '-', 'Belum ada data peserta didik pada kelas terpilih']]
              : studentRosterRows.map((st, idx) => [
                  idx + 1,
                  `${st.nis} / ${st.nisn || '-'}`,
                  st.name,
                  st.gender,
                  st.className,
                  st.absentNo || '-',
                  `${st.birthPlace || '-'}, ${st.birthDate || '-'}`,
                  st.phone || '-',
                  st.homeroomTeacher || '-',
                  st.counselorTeacher || '-',
                  st.status,
                ]),
        });
        cursorY =
          (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
      }

      if (has('teacher_activity')) {
        drawSectionTitle('Aktivitas & Penilaian Kinerja Guru — Ringkasan');
        autoTable(doc, {
          startY: cursorY,
          margin: { left: marginX, right: marginX },
          styles: commonTableStyles,
          headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
          head: [
            [
              'No',
              'NIP',
              'Nama Guru',
              'Jabatan / Mapel',
              'Jadwal',
              'Jurnal',
              '% Hadir',
              'Dinilai',
              'Rata Nilai',
              'Kinerja',
            ],
          ],
          body: teacherActivityData.summary.map((t) => [
            t.no,
            t.nip,
            t.name,
            `${t.position} (${t.subject})`,
            t.scheduleCount,
            t.journalCount,
            t.avgAttendance !== null ? `${t.avgAttendance.toFixed(0)}%` : '-',
            t.assessedStudents,
            t.avgGrade !== null ? t.avgGrade.toFixed(1) : '-',
            t.kinerja,
          ]),
        });
        cursorY =
          (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 6;

        drawSectionTitle('Riwayat Pelaksanaan KBM Terbaru');
        autoTable(doc, {
          startY: cursorY,
          margin: { left: marginX, right: marginX },
          styles: commonTableStyles,
          headStyles: { fillColor: [241, 245, 249], textColor: 0, fontStyle: 'bold' },
          head: [['No', 'Tanggal', 'Jam', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', '% Hadir']],
          body:
            teacherActivityData.recentKbm.length === 0
              ? [['-', '-', '-', '-', '-', 'Belum ada riwayat KBM pada periode filter ini', '-']]
              : teacherActivityData.recentKbm.map((r) => [
                  r.no,
                  formatShortDateId(r.date),
                  r.jamKe,
                  r.className,
                  r.subject,
                  r.topic,
                  `${r.pct.toFixed(0)}%`,
                ]),
        });
        cursorY =
          (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;
      }

      // Signature Block matching Preview & sigPos 1:1
      if (showTtdKepsek || showTtdGuru) {
        const kepsekSigPx = Math.max(
          24,
          Math.min(160, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
        );
        const guruSigPx = Math.max(
          24,
          Math.min(160, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
        );
        const stampPx = Math.max(28, Math.min(160, sigPos.stampSize ?? 70));
        const stampMm = Number((stampPx * 0.285).toFixed(2));

        const neededH =
          sigPos.preset === 'kiri' || sigPos.preset === 'kanan'
            ? 76
            : sigPos.preset === 'bebas'
            ? Math.max(48, Math.max(sigPos.principalY, sigPos.teacherY) * 0.26 + 38)
            : 48;
        if (cursorY > pageHeight - neededH) {
          doc.addPage();
          cursorY = 18;
        }
        const sigBaseY = cursorY + 4;
        const usableWidth = pageWidth - marginX * 2;
        const maxSigX = pageWidth - marginX - 50;

        const drawPdfSigBlock = (
          xMm: number,
          yMm: number,
          topLabel: string,
          roleTitle: string,
          personName: string,
          personNip: string,
          sigPng: string | null,
          withStamp: boolean,
          sigPx: number
        ) => {
          const sigH = Number((sigPx * 0.265).toFixed(2));
          const sigW = Number((sigH * 2.38).toFixed(2));
          doc.setFont(pdfFont, 'normal');
          doc.setFontSize(9.5);
          doc.text(topLabel, xMm, yMm);
          doc.text(roleTitle, xMm, yMm + 4.5);
          if (withStamp && stampPng) {
            doc.addImage(stampPng, 'PNG', xMm + 4.5 - stampMm / 2, yMm + 16 - stampMm / 2, stampMm, stampMm);
          }
          if (sigPng) {
            doc.addImage(sigPng, 'PNG', xMm, yMm + 14.5 - sigH / 2, sigW, sigH);
          }
          doc.setFont(pdfFont, 'bold');
          doc.text(personName, xMm, yMm + 25.5);
          doc.setFont(pdfFont, 'normal');
          doc.text(`NIP. ${personNip || '-'}`, xMm, yMm + 30);
        };

        const stampOnKepsek =
          sigPos.showStamp &&
          (sigPos.stampTarget === 'kepsek' || sigPos.stampTarget === 'keduanya');
        const stampOnGuru =
          sigPos.showStamp &&
          (sigPos.stampTarget === 'guru' || sigPos.stampTarget === 'keduanya');

        let kepsekX = marginX + 6;
        let kepsekY = sigBaseY;
        let guruX = pageWidth - marginX - 66;
        let guruY = sigBaseY;

        if (sigPos.preset === 'kanan-kiri') {
          guruX = marginX + 6;
          kepsekX = pageWidth - marginX - 66;
        } else if (sigPos.preset === 'kiri') {
          kepsekX = marginX + 6;
          guruX = marginX + 6;
          guruY = sigBaseY + 36;
        } else if (sigPos.preset === 'tengah') {
          kepsekX = pageWidth / 2 - 64;
          guruX = pageWidth / 2 + 10;
        } else if (sigPos.preset === 'kanan') {
          kepsekX = pageWidth - marginX - 66;
          guruX = pageWidth - marginX - 66;
          guruY = sigBaseY + 36;
        } else if (sigPos.preset === 'bebas') {
          kepsekX = Math.min(maxSigX, marginX + (usableWidth * sigPos.principalX) / 100);
          kepsekY = sigBaseY + sigPos.principalY * 0.26;
          guruX = Math.min(maxSigX, marginX + (usableWidth * sigPos.teacherX) / 100);
          guruY = sigBaseY + sigPos.teacherY * 0.26;
        }

        if (showTtdKepsek) {
          drawPdfSigBlock(
            kepsekX,
            kepsekY,
            'Mengetahui,',
            settings.principalTitle,
            settings.principalName,
            settings.principalNip,
            kepsekSigPng,
            stampOnKepsek,
            kepsekSigPx
          );
        }

        if (showTtdGuru) {
          drawPdfSigBlock(
            guruX,
            guruY,
            `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
            activeSigner.roleLabel,
            activeSigner.name,
            activeSigner.nip,
            guruSigPng,
            stampOnGuru,
            guruSigPx
          );
        }
      }

      doc.save(`Laporan_E-Guru_${reportSlug}_${getTodayISO()}.pdf`);
      onToast('success', 'Laporan PDF vektor 1:1 berhasil diunduh.');
    } catch {
      onToast('error', 'Terjadi kendala saat membuat PDF. Silakan coba lagi.');
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  const handlePrintDirect = () => {
    if (selectedReportTypes.length === 0) {
      onToast('error', 'Pilih minimal satu jenis laporan terlebih dahulu.');
      return;
    }
    window.print();
  };

  const includeJournal = selectedReportTypes.includes('journal');
  const includeAttendance = selectedReportTypes.includes('attendance');
  const includeAttendanceMonthly = selectedReportTypes.includes('attendance_monthly');
  const includeAttendanceMonitoring = selectedReportTypes.includes('attendance_monitoring');
  const includeGrades = selectedReportTypes.includes('grades');
  const includeGradesRemedial = selectedReportTypes.includes('grades_remedial');
  const includeGradesLeger = selectedReportTypes.includes('grades_leger');
  const includeStudentsRoster = selectedReportTypes.includes('students_roster');
  const includeTeacherActivity = selectedReportTypes.includes('teacher_activity');

  const stampOnKepsek =
    sigPos.showStamp && (sigPos.stampTarget === 'kepsek' || sigPos.stampTarget === 'keduanya');
  const stampOnGuru =
    sigPos.showStamp && (sigPos.stampTarget === 'guru' || sigPos.stampTarget === 'keduanya');

  const kepsekSigSizePx = Math.max(
    24,
    Math.min(160, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const guruSigSizePx = Math.max(
    24,
    Math.min(160, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const stampSizePx = Math.max(28, Math.min(160, sigPos.stampSize ?? 70));

  const renderSigCard = (
    topLabel: string,
    roleTitle: string,
    personName: string,
    personNip: string,
    sigImage: string,
    withStamp: boolean,
    sigPx: number
  ) => (
    <div style={{ fontSize: '11.5px', lineHeight: 1.4, textAlign: sigPos.preset === 'tengah' ? 'center' : 'left' }}>
      <p style={{ margin: 0 }}>{topLabel}</p>
      <p style={{ margin: 0 }}>{roleTitle}</p>
      <div
        style={{
          position: 'relative',
          height: '72px',
          margin: '4px 0',
          display: 'flex',
          alignItems: 'center',
          justifyContent: sigPos.preset === 'tengah' ? 'center' : 'flex-start',
          overflow: 'visible',
        }}
      >
        {withStamp && activeSigner.schoolStampImage && (
          <img
            src={activeSigner.schoolStampImage}
            alt="Stempel Sekolah"
            style={{
              position: 'absolute',
              left:
                sigPos.preset === 'tengah'
                  ? `calc(50% - ${23 + Math.round(stampSizePx / 2)}px)`
                  : `${19 - Math.round(stampSizePx / 2)}px`,
              top: `${Math.round((72 - stampSizePx) / 2)}px`,
              width: `${stampSizePx}px`,
              height: `${stampSizePx}px`,
              objectFit: 'contain',
              opacity: 0.85,
              mixBlendMode: 'multiply',
              pointerEvents: 'none',
            }}
          />
        )}
        {sigImage && (
          <img
            src={sigImage}
            alt={`TTD ${personName}`}
            style={{
              position: 'relative',
              zIndex: 2,
              height: `${sigPx}px`,
              maxWidth: `${Math.round(sigPx * 2.85)}px`,
              objectFit: 'contain',
              pointerEvents: 'none',
            }}
          />
        )}
      </div>
      <p style={{ margin: 0, fontWeight: 700, textDecoration: 'underline', textUnderlineOffset: '2px' }}>
        {personName}
      </p>
      <p style={{ margin: 0 }}>NIP. {personNip || '-'}</p>
    </div>
  );

  return (
    <div className="space-y-6">
      {/* Control Panel (Hidden on Print) */}
      <section className="no-print bg-white border border-slate-200 rounded-xl p-5 space-y-5">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 border-b border-slate-200 pb-4">
          <div>
            <h2 className="text-base font-bold text-slate-900">
              Filter & Pengaturan Cetak Laporan Resmi
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Hasil cetak, PDF, dan Word disesuaikan 1:1 dengan pratinjau dokumen, TTD pemilik akun, stempel sekolah & posisi pengesahan
            </p>
          </div>

          {/* Action & Output Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={handleSavePrintSettings}
              className={`px-3 py-2 rounded-lg text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap ${
                hasUnsavedChanges
                  ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-xs'
                  : 'border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-800'
              }`}
              title="Simpan permanen seluruh pengaturan cetak agar tidak berubah lagi"
            >
              <Save className="w-3.5 h-3.5" />
              {hasUnsavedChanges ? 'Simpan Pengaturan*' : 'Simpan Pengaturan'}
            </button>
            <button
              type="button"
              onClick={handleDownloadXlsx}
              className="px-3 py-2 rounded-lg border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Excel (.xlsx)
            </button>
            <button
              type="button"
              onClick={handleDownloadXls}
              className="px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
            >
              <FileDown className="w-3.5 h-3.5" />
              .XLS
            </button>
            <button
              type="button"
              onClick={handleDownloadWord}
              disabled={isExportingWord}
              className="px-3 py-2 rounded-lg border border-sky-200 bg-sky-50 hover:bg-sky-100 text-sky-800 text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap disabled:opacity-60"
            >
              <FileType2 className="w-3.5 h-3.5" />
              {isExportingWord ? 'Membuat Word...' : 'Word (.doc)'}
            </button>
            <button
              type="button"
              onClick={handleDownloadPdf}
              disabled={isGeneratingPdf}
              className="px-3 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap disabled:opacity-60"
            >
              <FileText className="w-3.5 h-3.5" />
              {isGeneratingPdf ? 'Membuat PDF...' : 'Unduh PDF'}
            </button>
            <button
              type="button"
              onClick={handlePrintDirect}
              className="px-3.5 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
            >
              <Printer className="w-3.5 h-3.5" />
              Cetak Langsung
            </button>
          </div>
        </div>

        {/* Row 1 Filters */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
          <div className="relative" ref={reportDropdownRef}>
            <span className="block text-xs font-semibold text-slate-700 mb-1.5">
              Jenis Laporan (Multi-Select)
            </span>
            <button
              type="button"
              onClick={() => setReportDropdownOpen((o) => !o)}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white text-left flex items-center justify-between gap-2 focus:outline-none focus:border-indigo-600"
            >
              <span className="truncate text-slate-800">
                {selectedReportTypes.length === 0
                  ? 'Tidak ada laporan dipilih'
                  : selectedReportTypes.length === REPORT_TYPE_OPTIONS.length
                  ? `Semua Laporan (${REPORT_TYPE_OPTIONS.length})`
                  : REPORT_TYPE_OPTIONS.filter((o) => selectedReportTypes.includes(o.id))
                      .map((o) => o.label)
                      .join(', ')}
              </span>
              <ChevronDown className="w-4 h-4 text-slate-400 shrink-0" />
            </button>

            {reportDropdownOpen && (
              <div className="absolute left-0 mt-1.5 z-30 w-full min-w-[300px] sm:min-w-[370px] bg-white border border-slate-200 rounded-xl shadow-xl p-2 space-y-1">
                <div className="flex items-center justify-between px-2 py-1 border-b border-slate-100 text-xs">
                  <button
                    type="button"
                    onClick={() =>
                      setSelectedReportTypes(REPORT_TYPE_OPTIONS.map((o) => o.id))
                    }
                    className="font-semibold text-indigo-600 hover:underline"
                  >
                    Pilih Semua ({REPORT_TYPE_OPTIONS.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedReportTypes([])}
                    className="text-slate-500 hover:text-slate-700"
                  >
                    Kosongkan
                  </button>
                </div>
                <div className="max-h-72 overflow-y-auto space-y-1 pt-1 pr-0.5">
                  {REPORT_TYPE_OPTIONS.map((opt) => (
                    <label
                      key={opt.id}
                      className="flex items-start gap-2.5 px-2.5 py-2 rounded-lg hover:bg-slate-50 cursor-pointer text-slate-800"
                    >
                      <input
                        type="checkbox"
                        checked={selectedReportTypes.includes(opt.id)}
                        onChange={() => toggleReportTypeSelection(opt.id)}
                        className="mt-0.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                      />
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center justify-between gap-2">
                          <span className="text-xs font-semibold text-slate-900">
                            {opt.label}
                          </span>
                          <span
                            className={`text-[10px] font-semibold px-1.5 py-0.5 rounded border shrink-0 ${CATEGORY_BADGE_CLASS[opt.category]}`}
                          >
                            {opt.category}
                          </span>
                        </div>
                        <p className="text-[11px] text-slate-500 mt-0.5 leading-snug">
                          {opt.description}
                        </p>
                      </div>
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>

          <div className="relative" ref={dropdownRef}>
            <span className="block text-xs font-semibold text-slate-700 mb-1.5">
              Pilih Kelas (Multi-Select)
            </span>
            <button
              type="button"
              onClick={() => setClassDropdownOpen((o) => !o)}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white text-left flex items-center justify-between gap-2 focus:outline-none focus:border-indigo-600"
            >
              <span className="truncate text-slate-800">
                {selectedClasses.length === 0
                  ? 'Tidak ada kelas dipilih'
                  : selectedClasses.length === classOptions.length
                  ? `Semua Kelas (${classOptions.length})`
                  : selectedClasses.join(', ')}
              </span>
              <ChevronDown className="w-4 h-4 text-slate-400 shrink-0" />
            </button>

            {classDropdownOpen && (
              <div className="absolute left-0 right-0 mt-1.5 z-30 bg-white border border-slate-200 rounded-xl shadow-lg p-2 space-y-1">
                <div className="flex items-center justify-between px-2 py-1 border-b border-slate-100 text-xs">
                  <button
                    type="button"
                    onClick={() => setSelectedClasses(classOptions)}
                    className="font-semibold text-indigo-600 hover:underline"
                  >
                    Pilih Semua
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedClasses([])}
                    className="text-slate-500 hover:text-slate-700"
                  >
                    Kosongkan
                  </button>
                </div>
                <div className="max-h-48 overflow-y-auto space-y-0.5 pt-1">
                  {classOptions.map((cls) => (
                    <label
                      key={cls}
                      className="flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer text-sm text-slate-800"
                    >
                      <input
                        type="checkbox"
                        checked={selectedClasses.includes(cls)}
                        onChange={() => toggleClassSelection(cls)}
                        className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                      />
                      <span>{cls}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>

          <div>
            <label htmlFor="rep-subject" className="block text-xs font-semibold text-slate-700 mb-1.5">
              Mata Pelajaran
            </label>
            <select
              id="rep-subject"
              value={subjectFilter}
              onChange={(e) => setSubjectFilter(e.target.value)}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              <option value="ALL">Semua Mata Pelajaran</option>
              {settings.masterSubjects.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label htmlFor="rep-paper" className="block text-xs font-semibold text-slate-700 mb-1.5">
              Ukuran Kertas
            </label>
            <select
              id="rep-paper"
              value={paperSize}
              onChange={(e) => setPaperSize(e.target.value as PaperSizeId)}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              {(Object.keys(PAPER_SIZE_CONFIG) as PaperSizeId[]).map((key) => (
                <option key={key} value={key}>
                  {PAPER_SIZE_CONFIG[key].label}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label htmlFor="rep-font" className="block text-xs font-semibold text-slate-700 mb-1.5">
              Jenis Huruf Cetak
            </label>
            <select
              id="rep-font"
              value={fontFamily}
              onChange={(e) => setFontFamily(e.target.value as PrintFontFamily)}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              {FONT_OPTIONS.map((f) => (
                <option key={f} value={f} style={{ fontFamily: f }}>
                  {f}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Row 2: Date Range + Toggles */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 pt-2 border-t border-slate-100 items-end">
          <div className="lg:col-span-6 flex flex-wrap items-end gap-2.5">
            <div>
              <label htmlFor="rep-start" className="block text-xs font-semibold text-slate-700 mb-1">
                Tanggal Mulai
              </label>
              <input
                id="rep-start"
                type="date"
                value={startDate}
                onChange={(e) => setStartDate(e.target.value)}
                className="px-3 py-1.5 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>
            <div>
              <label htmlFor="rep-end" className="block text-xs font-semibold text-slate-700 mb-1">
                Tanggal Selesai
              </label>
              <input
                id="rep-end"
                type="date"
                value={endDate}
                onChange={(e) => setEndDate(e.target.value)}
                className="px-3 py-1.5 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => applyDatePreset('thisMonth')}
                className="px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-50"
              >
                Bulan Ini
              </button>
              <button
                type="button"
                onClick={() => applyDatePreset('lastMonth')}
                className="px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-50"
              >
                Bulan Lalu
              </button>
              <button
                type="button"
                onClick={() => applyDatePreset('reset')}
                className="px-2.5 py-1.5 rounded-lg border border-slate-200 text-xs font-medium text-slate-500 hover:bg-slate-50 inline-flex items-center gap-1"
              >
                <RotateCcw className="w-3 h-3" />
                Reset
              </button>
            </div>
          </div>

          <div className="lg:col-span-6 flex flex-wrap items-center lg:justify-end gap-4 text-xs sm:text-sm">
            <label className="inline-flex items-center gap-2 cursor-pointer text-slate-700">
              <input
                type="checkbox"
                checked={showKop}
                onChange={(e) => setShowKop(e.target.checked)}
                className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
              />
              Kop Surat Dinas
            </label>
            <label className="inline-flex items-center gap-2 cursor-pointer text-slate-700">
              <input
                type="checkbox"
                checked={showTtdKepsek}
                onChange={(e) => setShowTtdKepsek(e.target.checked)}
                className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
              />
              TTD Kepala Sekolah
            </label>
            <label className="inline-flex items-center gap-2 cursor-pointer text-slate-700">
              <input
                type="checkbox"
                checked={showTtdGuru}
                onChange={(e) => setShowTtdGuru(e.target.checked)}
                className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
              />
              TTD Pemilik Akun
            </label>
            <div className="inline-flex items-center p-1 bg-slate-100 rounded-lg text-xs">
              <button
                type="button"
                onClick={() => setOrientation('portrait')}
                className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                  orientation === 'portrait'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Portrait
              </button>
              <button
                type="button"
                onClick={() => setOrientation('landscape')}
                className={`px-2.5 py-1 rounded-md font-medium transition-colors ${
                  orientation === 'landscape'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                Landscape
              </button>
            </div>
          </div>
        </div>

        {/* Row 3: Pemilik Akun Pengesah + Posisi Blok Pengesahan Bebas + Stempel */}
        <div className="pt-3 border-t border-slate-100 space-y-3">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-end">
            <div>
              <label htmlFor="rep-signer" className="block text-xs font-semibold text-slate-700 mb-1.5">
                <UserCheck className="w-3.5 h-3.5 inline mr-1 text-indigo-600" />
                Akun Penanda Tangan & Pemilik Stempel
              </label>
              <select
                id="rep-signer"
                value={signerAccountId}
                onChange={(e) => setSignerAccountId(e.target.value)}
                className="w-full px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">Otomatis Sesuai Akun Login ({activeSigner.name})</option>
                {accounts.map((acc) => (
                  <option key={acc.id} value={acc.id}>
                    {acc.name} — {acc.role}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label htmlFor="rep-sig-preset" className="block text-xs font-semibold text-slate-700 mb-1.5">
                <Move className="w-3.5 h-3.5 inline mr-1 text-indigo-600" />
                Posisi Blok Pengesahan (TTD & Stempel)
              </label>
              <select
                id="rep-sig-preset"
                value={sigPos.preset}
                onChange={(e) => handlePresetChange(e.target.value as SignatureLayoutPreset)}
                className="w-full px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                {PRESET_OPTIONS.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.label}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex flex-wrap items-center gap-3 pb-0.5">
              <div>
                <span className="block text-[11px] font-semibold text-slate-700 mb-1">
                  Gambar TTD Digital
                </span>
                <div className="inline-flex items-center p-0.5 bg-slate-100 rounded-lg text-xs border border-slate-200">
                  <button
                    type="button"
                    onClick={() => setShowSignatureImage(true)}
                    className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                      showSignatureImage
                        ? 'bg-white text-indigo-700 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Terlihat
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowSignatureImage(false)}
                    className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                      !showSignatureImage
                        ? 'bg-white text-rose-700 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Tidak
                  </button>
                </div>
              </div>

              <div>
                <span className="block text-[11px] font-semibold text-slate-700 mb-1">
                  <Stamp className="w-3 h-3 inline mr-1 text-indigo-600" />
                  Gambar Stempel Sekolah
                </span>
                <div className="inline-flex items-center gap-1.5">
                  <div className="inline-flex items-center p-0.5 bg-slate-100 rounded-lg text-xs border border-slate-200">
                    <button
                      type="button"
                      onClick={() => updateSigPos({ showStamp: true })}
                      className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                        sigPos.showStamp
                          ? 'bg-white text-indigo-700 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Terlihat
                    </button>
                    <button
                      type="button"
                      onClick={() => updateSigPos({ showStamp: false })}
                      className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                        !sigPos.showStamp
                          ? 'bg-white text-rose-700 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Tidak
                    </button>
                  </div>

                  {sigPos.showStamp && (
                    <select
                      value={sigPos.stampTarget}
                      onChange={(e) =>
                        updateSigPos({
                          stampTarget: e.target.value as 'guru' | 'kepsek' | 'keduanya',
                        })
                      }
                      aria-label="Target stempel sekolah"
                      className="px-2 py-1.5 text-xs border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                    >
                      <option value="guru">Di Blok Guru</option>
                      <option value="kepsek">Di Blok Kepsek</option>
                      <option value="keduanya">Di Keduanya</option>
                    </select>
                  )}
                </div>
              </div>
            </div>
          </div>

          <div className="bg-indigo-50/70 border border-indigo-200 rounded-xl p-3.5 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <p className="text-xs font-semibold text-indigo-900">
                Atur Posisi & Ukuran Gambar Tanda Tangan / Stempel (atau seret langsung kotak TTD di pratinjau bawah):
              </p>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() =>
                    updateSigPos({
                      principalSignatureSize: 60,
                      teacherSignatureSize: 60,
                      stampSize: 70,
                    })
                  }
                  className="px-2.5 py-1 rounded-md bg-white border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100"
                >
                  Reset Ukuran
                </button>
                <button
                  type="button"
                  onClick={() =>
                    updateSigPos({
                      preset: 'bebas',
                      principalX: 4,
                      principalY: 0,
                      teacherX: 62,
                      teacherY: 0,
                    })
                  }
                  className="px-2.5 py-1 rounded-md bg-white border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100"
                >
                  Reset Posisi
                </button>
              </div>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-1">
                  <span className="font-semibold text-slate-700">
                    Kepsek Kiri ↔ Kanan: {sigPos.principalX}%
                  </span>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({
                          preset: 'bebas',
                          principalX: Math.max(0, sigPos.principalX - 1),
                        })
                      }
                      aria-label="Geser Kepsek ke kiri"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({
                          preset: 'bebas',
                          principalX: Math.min(76, sigPos.principalX + 1),
                        })
                      }
                      aria-label="Geser Kepsek ke kanan"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      +
                    </button>
                  </div>
                </div>
                <input
                  type="range"
                  min={0}
                  max={76}
                  value={sigPos.principalX}
                  onChange={(e) =>
                    updateSigPos({ preset: 'bebas', principalX: Number(e.target.value) })
                  }
                  className="w-full accent-indigo-600"
                />
              </div>

              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-1">
                  <span className="font-semibold text-slate-700">
                    Kepsek Atas / Bawah: {sigPos.principalY}px
                  </span>
                </div>
                <input
                  type="range"
                  min={0}
                  max={160}
                  value={sigPos.principalY}
                  onChange={(e) =>
                    updateSigPos({ preset: 'bebas', principalY: Number(e.target.value) })
                  }
                  className="w-full accent-indigo-600"
                />
              </div>

              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-1">
                  <span className="font-semibold text-slate-700">
                    Guru Kiri ↔ Kanan: {sigPos.teacherX}%
                  </span>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({
                          preset: 'bebas',
                          teacherX: Math.max(0, sigPos.teacherX - 1),
                        })
                      }
                      aria-label="Geser Guru ke kiri"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({
                          preset: 'bebas',
                          teacherX: Math.min(76, sigPos.teacherX + 1),
                        })
                      }
                      aria-label="Geser Guru ke kanan"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      +
                    </button>
                  </div>
                </div>
                <input
                  type="range"
                  min={0}
                  max={76}
                  value={sigPos.teacherX}
                  onChange={(e) =>
                    updateSigPos({ preset: 'bebas', teacherX: Number(e.target.value) })
                  }
                  className="w-full accent-indigo-600"
                />
              </div>

              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-1">
                  <span className="font-semibold text-slate-700">
                    Guru Atas / Bawah: {sigPos.teacherY}px
                  </span>
                </div>
                <input
                  type="range"
                  min={0}
                  max={160}
                  value={sigPos.teacherY}
                  onChange={(e) =>
                    updateSigPos({ preset: 'bebas', teacherY: Number(e.target.value) })
                  }
                  className="w-full accent-indigo-600"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="font-semibold text-slate-700 tabular-nums">
                    Ukuran TTD Kepsek: {kepsekSigSizePx}px
                  </span>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ principalSignatureSize: Math.max(28, kepsekSigSizePx - 4) })
                      }
                      aria-label="Perkecil TTD Kepala Sekolah"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ principalSignatureSize: Math.min(130, kepsekSigSizePx + 4) })
                      }
                      aria-label="Perbesar TTD Kepala Sekolah"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      +
                    </button>
                  </div>
                </div>
                <input
                  type="range"
                  min={28}
                  max={130}
                  value={kepsekSigSizePx}
                  onChange={(e) =>
                    updateSigPos({ principalSignatureSize: Number(e.target.value) })
                  }
                  aria-label="Slider ukuran TTD Kepala Sekolah"
                  className="w-full accent-indigo-600"
                />
              </div>

              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="font-semibold text-slate-700 tabular-nums">
                    Ukuran TTD Guru / Akun: {guruSigSizePx}px
                  </span>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ teacherSignatureSize: Math.max(28, guruSigSizePx - 4) })
                      }
                      aria-label="Perkecil TTD Guru"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ teacherSignatureSize: Math.min(130, guruSigSizePx + 4) })
                      }
                      aria-label="Perbesar TTD Guru"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      +
                    </button>
                  </div>
                </div>
                <input
                  type="range"
                  min={28}
                  max={130}
                  value={guruSigSizePx}
                  onChange={(e) =>
                    updateSigPos({ teacherSignatureSize: Number(e.target.value) })
                  }
                  aria-label="Slider ukuran TTD Guru"
                  className="w-full accent-indigo-600"
                />
              </div>

              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="font-semibold text-slate-700 tabular-nums">
                    Ukuran Gambar Stempel: {stampSizePx}px
                  </span>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ stampSize: Math.max(32, stampSizePx - 4) })
                      }
                      aria-label="Perkecil gambar stempel"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ stampSize: Math.min(130, stampSizePx + 4) })
                      }
                      aria-label="Perbesar gambar stempel"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      +
                    </button>
                  </div>
                </div>
                <input
                  type="range"
                  min={32}
                  max={130}
                  value={stampSizePx}
                  onChange={(e) => updateSigPos({ stampSize: Number(e.target.value) })}
                  aria-label="Slider ukuran gambar stempel"
                  className="w-full accent-indigo-600"
                />
              </div>
            </div>
          </div>

          {/* Save Print Settings Footer Bar */}
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 pt-3 border-t border-slate-200">
            <div className="flex items-center gap-2 text-xs">
              {hasUnsavedChanges ? (
                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-amber-50 border border-amber-200 text-amber-800 font-medium">
                  Ada perubahan pengaturan cetak yang belum disimpan
                </span>
              ) : (
                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700 font-medium">
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  Pengaturan cetak tersimpan permanen di perangkat
                </span>
              )}
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {hasUnsavedChanges && (
                <button
                  type="button"
                  onClick={handleRevertToSaved}
                  className="px-3 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold inline-flex items-center gap-1.5 transition-colors"
                >
                  <Undo2 className="w-3.5 h-3.5" />
                  Kembalikan Tersimpan
                </button>
              )}
              <button
                type="button"
                onClick={handleSavePrintSettings}
                className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold inline-flex items-center gap-1.5 transition-colors"
              >
                <Save className="w-3.5 h-3.5" />
                Simpan Pengaturan Cetak
              </button>
            </div>
          </div>
        </div>
      </section>

      {/* Real-Time Document Preview (Exact 1:1 source for Print & Word) */}
      <section className="overflow-x-auto pb-4">
        <div
          ref={previewRef}
          style={{
            fontFamily: `"${fontFamily}", serif`,
            backgroundColor: '#ffffff',
            color: '#000000',
          }}
          className={`print-area mx-auto bg-white border border-slate-200 shadow-sm p-6 sm:p-10 text-black ${
            orientation === 'landscape'
              ? `${PAPER_SIZE_CONFIG[paperSize].landscapeMinW} ${PAPER_SIZE_CONFIG[paperSize].landscapeMaxW}`
              : PAPER_SIZE_CONFIG[paperSize].portraitMaxW
          }`}
        >
          {/* KOP SURAT DINAS */}
          {showKop && (
            <header style={{ borderBottom: '4px double #000', paddingBottom: '10px', marginBottom: '20px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <tbody>
                  <tr>
                    <td style={{ width: '82px', textAlign: 'center', verticalAlign: 'middle', border: 'none' }}>
                      {settings.kopSurat.logoLeft && (
                        <img
                          src={settings.kopSurat.logoLeft}
                          alt="Logo Instansi Kiri"
                          style={{ width: '72px', height: '72px', objectFit: 'contain', display: 'inline-block' }}
                        />
                      )}
                    </td>
                    <td style={{ textAlign: 'center', verticalAlign: 'middle', lineHeight: 1.25, border: 'none' }}>
                      {settings.kopSurat.line1 && (
                        <div style={{ fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.03em' }}>
                          {settings.kopSurat.line1}
                        </div>
                      )}
                      {settings.kopSurat.line2 && (
                        <div style={{ fontSize: '14px', fontWeight: 700, textTransform: 'uppercase' }}>
                          {settings.kopSurat.line2}
                        </div>
                      )}
                      {settings.kopSurat.line3 && (
                        <div style={{ fontSize: '16px', fontWeight: 700, textTransform: 'uppercase', marginTop: '2px' }}>
                          {settings.kopSurat.line3}
                        </div>
                      )}
                      {settings.kopSurat.line4 && (
                        <div style={{ fontSize: '11px', marginTop: '3px' }}>{settings.kopSurat.line4}</div>
                      )}
                      {settings.kopSurat.line5 && (
                        <div style={{ fontSize: '11px' }}>{settings.kopSurat.line5}</div>
                      )}
                      {settings.kopSurat.line6 && (
                        <div style={{ fontSize: '11px' }}>{settings.kopSurat.line6}</div>
                      )}
                    </td>
                    <td style={{ width: '82px', textAlign: 'center', verticalAlign: 'middle', border: 'none' }}>
                      {settings.kopSurat.logoRight && (
                        <img
                          src={settings.kopSurat.logoRight}
                          alt="Logo Sekolah Kanan"
                          style={{ width: '72px', height: '72px', objectFit: 'contain', display: 'inline-block' }}
                        />
                      )}
                    </td>
                  </tr>
                </tbody>
              </table>
            </header>
          )}

          {/* Judul Dokumen */}
          <div style={{ textAlign: 'center', marginBottom: '18px' }}>
            <h2
              style={{
                fontSize: '14px',
                fontWeight: 700,
                textTransform: 'uppercase',
                textDecoration: 'underline',
                textUnderlineOffset: '4px',
                margin: 0,
              }}
            >
              {reportTitle}
            </h2>
            <p style={{ fontSize: '11px', marginTop: '6px', color: '#334155' }}>
              {periodSubtitle} · Rentang Tanggal: {dateRangeLabel}
            </p>
          </div>

          {selectedReportTypes.length === 0 && (
            <div
              style={{
                border: '1px dashed #94a3b8',
                padding: '24px',
                textAlign: 'center',
                color: '#475569',
                fontSize: '12px',
                marginBottom: '20px',
              }}
            >
              Belum ada jenis laporan yang dipilih. Silakan centang minimal 1 jenis laporan di menu <strong>Jenis Laporan (Multi-Select)</strong> di atas.
            </div>
          )}

          {/* 1. Jurnal Mengajar */}
          {includeJournal && (
            <div style={{ marginBottom: '24px' }}>
              {selectedReportTypes.length > 1 && (
                <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                  A. Jurnal Kegiatan Mengajar
                </h3>
              )}
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px' }}>
                <thead>
                  <tr style={{ backgroundColor: '#f1f5f9' }}>
                    <th style={{ border: '1px solid #000', padding: '5px 6px', textAlign: 'center', width: '30px' }}>No</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px', whiteSpace: 'nowrap' }}>Tanggal</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px', textAlign: 'center', whiteSpace: 'nowrap' }}>Jam</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px', whiteSpace: 'nowrap' }}>Kelas</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px', whiteSpace: 'nowrap' }}>Mata Pelajaran</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px' }}>Materi Pokok</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px' }}>Kegiatan Pembelajaran</th>
                    <th style={{ border: '1px solid #000', padding: '5px 6px', textAlign: 'center', whiteSpace: 'nowrap' }}>H/S/I/A/D</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredJournals.length === 0 ? (
                    <tr>
                      <td colSpan={8} style={{ border: '1px solid #000', padding: '12px', textAlign: 'center', color: '#64748b' }}>
                        Tidak ada catatan jurnal mengajar untuk kriteria filter terpilih.
                      </td>
                    </tr>
                  ) : (
                    filteredJournals.map((j, idx) => {
                      const sum = summarizeJournalRow(j);
                      return (
                        <tr key={j.id}>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }} className="tabular-nums">
                            {idx + 1}
                          </td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }} className="tabular-nums">
                            {formatShortDateId(j.date)}
                          </td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }} className="tabular-nums">
                            {j.jamKe}
                          </td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap', fontWeight: 600 }}>
                            {j.className}
                          </td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>
                            {j.subject}
                          </td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px' }}>{j.topic}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px' }}>{j.activity}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', whiteSpace: 'nowrap' }} className="tabular-nums">
                            {sum.counts.H}/{sum.counts.S}/{sum.counts.I}/{sum.counts.A}/{sum.counts.D}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* 2. Leger Kehadiran Mendatar */}
          {includeAttendance && (
            <div style={{ marginBottom: '24px' }}>
              {selectedClasses.map((cls) => {
                const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
                return (
                  <div key={cls} style={{ marginBottom: '18px' }}>
                    <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                      Leger Kehadiran Mendatar — Kelas {cls}
                    </h3>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                      <thead>
                        <tr style={{ backgroundColor: '#f1f5f9' }}>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', width: '28px' }}>No</th>
                          <th style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }}>NIS</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                          {classJournals.map((j) => (
                            <th key={j.id} style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                              {formatCompactDateId(j.date)}
                            </th>
                          ))}
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>H</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>S</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>I</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>A</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>D</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', whiteSpace: 'nowrap' }}>%</th>
                        </tr>
                      </thead>
                      <tbody>
                        {classStudents.length === 0 ? (
                          <tr>
                            <td colSpan={9 + classJournals.length} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                              Belum ada data siswa di kelas {cls}.
                            </td>
                          </tr>
                        ) : (
                          classStudents.map((st, idx) => {
                            const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                            const codes = classJournals.map((j) => {
                              const code = j.attendance[st.id]?.status ?? 'H';
                              counts[code] += 1;
                              return code;
                            });
                            const pct =
                              classJournals.length > 0
                                ? `${((counts.H / classJournals.length) * 100).toFixed(0)}%`
                                : '-';
                            return (
                              <tr key={st.id}>
                                <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }} className="tabular-nums">{idx + 1}</td>
                                <td style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }} className="tabular-nums">{st.nis}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{st.name}</td>
                                {codes.map((c, cIdx) => (
                                  <td key={cIdx} style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', fontWeight: 600 }}>
                                    {c}
                                  </td>
                                ))}
                                <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{counts.H}</td>
                                <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{counts.S}</td>
                                <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{counts.I}</td>
                                <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{counts.A}</td>
                                <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{counts.D}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', fontWeight: 700 }}>{pct}</td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
          )}

          {/* 3. Rekap Matriks Nilai */}
          {includeGrades && (
            <div style={{ marginBottom: '24px' }}>
              {selectedClasses.map((cls) =>
                activeGradeSubjects.map((subj) => {
                  const { rows } = getGradeBlockForClassAndSubject(cls, subj);
                  const hasAnyScore = rows.some((r) => r.avg !== null);
                  if (!hasAnyScore && subjectFilter === 'ALL') return null;

                  return (
                    <div key={`${cls}-${subj}`} style={{ marginBottom: '18px' }}>
                      <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                        Rekap Matriks Nilai — Kelas {cls} ({subj})
                      </h3>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px' }}>
                        <thead>
                          <tr style={{ backgroundColor: '#f1f5f9' }}>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '30px' }}>No</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NIS</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                            {settings.masterExamTypes.map((ex) => (
                              <th key={ex} style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                                {ex}
                              </th>
                            ))}
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', whiteSpace: 'nowrap' }}>Rata-rata</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left', whiteSpace: 'nowrap' }}>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rows.length === 0 ? (
                            <tr>
                              <td
                                colSpan={5 + settings.masterExamTypes.length}
                                style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}
                              >
                                Belum ada siswa di kelas {cls}.
                              </td>
                            </tr>
                          ) : (
                            rows.map((r) => (
                              <tr key={r.nis + r.name}>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{r.no}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.nis}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.name}</td>
                                {settings.masterExamTypes.map((ex) => (
                                  <td key={ex} style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>
                                    {r.scores[ex] !== null ? r.scores[ex] : '-'}
                                  </td>
                                ))}
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', fontWeight: 700 }}>
                                  {r.avg !== null ? r.avg.toFixed(1) : '-'}
                                </td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.status}</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  );
                })
              )}
            </div>
          )}

          {/* 3. Rekap Presensi Bulanan (Presensi) */}
          {includeAttendanceMonthly && (
            <div style={{ marginBottom: '24px' }}>
              {selectedClasses.map((cls) => {
                const { months, rows, classAvgPct } = getMonthlyAttendanceForClass(cls);
                return (
                  <div key={`monthly-${cls}`} style={{ marginBottom: '18px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: '6px' }}>
                      <h3 style={{ fontSize: '12px', fontWeight: 700, margin: 0 }}>
                        Rekap Presensi Bulanan — Kelas {cls}
                      </h3>
                      <span style={{ fontSize: '10.5px', color: '#334155' }}>
                        Rata-rata Kehadiran Kelas: {classAvgPct !== null ? `${classAvgPct.toFixed(1)}%` : '-'}
                      </span>
                    </div>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                      <thead>
                        <tr style={{ backgroundColor: '#f1f5f9' }}>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', width: '28px' }}>No</th>
                          <th style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }}>NIS</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                          {months.map((m) => (
                            <th key={m} style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                              {formatMonthLabelId(m)}
                            </th>
                          ))}
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', whiteSpace: 'nowrap' }}>Sesi</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>H</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>S</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>I</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>A</th>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>D</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', whiteSpace: 'nowrap' }}>% Hadir</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.length === 0 ? (
                          <tr>
                            <td colSpan={10 + months.length} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                              Belum ada siswa di kelas {cls}.
                            </td>
                          </tr>
                        ) : (
                          rows.map((r) => (
                            <tr key={`m-${cls}-${r.nis}-${r.no}`}>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }} className="tabular-nums">{r.no}</td>
                              <td style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }} className="tabular-nums">{r.nis}</td>
                              <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.name}</td>
                              {months.map((m) => {
                                const mc = r.monthCounts[m];
                                return (
                                  <td key={m} style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }} className="tabular-nums">
                                    {mc && mc.total > 0 ? `${mc.h}/${mc.total}` : '-'}
                                  </td>
                                );
                              })}
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }} className="tabular-nums">{r.totalMeetings}</td>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{r.totals.H}</td>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{r.totals.S}</td>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{r.totals.I}</td>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{r.totals.A}</td>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{r.totals.D}</td>
                              <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', fontWeight: 700 }}>
                                {r.pct !== null ? `${r.pct.toFixed(0)}%` : '-'}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
          )}

          {/* 4. Monitoring S/I/A/D (Presensi) */}
          {includeAttendanceMonitoring && (
            <div style={{ marginBottom: '24px' }}>
              <div style={{ marginBottom: '14px' }}>
                <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                  Monitoring S/I/A/D — Rekap Akumulasi Ketidakhadiran Siswa
                </h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f1f5f9' }}>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '28px' }}>No</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NIS</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Kelas</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>S</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>I</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>A</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>D</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Total</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left', whiteSpace: 'nowrap' }}>Status Pembinaan</th>
                    </tr>
                  </thead>
                  <tbody>
                    {monitoringAbsenceData.summary.length === 0 ? (
                      <tr>
                        <td colSpan={10} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                          Seluruh siswa tercatat Hadir (H) 100% pada periode filter ini.
                        </td>
                      </tr>
                    ) : (
                      monitoringAbsenceData.summary.map((it, idx) => (
                        <tr key={it.studentId}>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{idx + 1}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{it.nis}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{it.name}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{it.className}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{it.S}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{it.I}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', fontWeight: 700 }}>{it.A}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{it.D}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', fontWeight: 700 }}>{it.total}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{it.action}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

              <div>
                <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                  Kronologi Ketidakhadiran (S/I/A/D)
                </h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f1f5f9' }}>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '28px' }}>No</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Tanggal</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Jam</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Kelas</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Mata Pelajaran</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Status</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Keterangan</th>
                    </tr>
                  </thead>
                  <tbody>
                    {monitoringAbsenceData.chronology.length === 0 ? (
                      <tr>
                        <td colSpan={8} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                          Belum ada catatan ketidakhadiran pada periode filter ini.
                        </td>
                      </tr>
                    ) : (
                      monitoringAbsenceData.chronology.map((c, idx) => (
                        <tr key={`${c.date}-${c.nis}-${idx}`}>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{idx + 1}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{formatShortDateId(c.date)}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{c.jamKe}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{c.className}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{c.subject}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{c.studentName}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', fontWeight: 700 }}>{c.status}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px' }}>{c.note}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* 6. Ketuntasan & Remedial (< KKM 75) */}
          {includeGradesRemedial && (
            <div style={{ marginBottom: '24px' }}>
              {selectedClasses.map((cls) =>
                activeGradeSubjects.map((subj) => {
                  const rem = getRemedialDataForClassAndSubject(cls, subj);
                  if (rem.assessedCount === 0 && subjectFilter === 'ALL') return null;

                  return (
                    <div key={`rem-${cls}-${subj}`} style={{ marginBottom: '18px' }}>
                      <div style={{ display: 'flex', flexWrap: 'wrap', justifyContent: 'space-between', alignItems: 'baseline', gap: '8px', marginBottom: '6px' }}>
                        <h3 style={{ fontSize: '12px', fontWeight: 700, margin: 0 }}>
                          Ketuntasan & Remedial (&lt; KKM 75) — Kelas {cls} ({subj})
                        </h3>
                        <span style={{ fontSize: '10.5px', color: '#334155' }}>
                          Daya Serap: <strong>{rem.dayaSerap !== null ? `${rem.dayaSerap.toFixed(0)}%` : '-'}</strong> · Tuntas: {rem.tuntasCount}/{rem.assessedCount} · Remedial: {rem.remedialCount} · Rata Kelas: {rem.classAvg !== null ? rem.classAvg.toFixed(1) : '-'}
                        </span>
                      </div>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                        <thead>
                          <tr style={{ backgroundColor: '#f1f5f9' }}>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '28px' }}>No</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NIS</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>Tertinggi</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>Terendah</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Komponen &lt; KKM</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>Rata-rata</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>Selisih</th>
                            <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Rekomendasi</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rem.assessedCount === 0 ? (
                            <tr>
                              <td colSpan={9} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                                Belum ada nilai tercatat untuk kelas {cls} ({subj}).
                              </td>
                            </tr>
                          ) : rem.remedialList.length === 0 ? (
                            <tr>
                              <td colSpan={9} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#15803d', fontWeight: 600 }}>
                                Seluruh {rem.assessedCount} siswa yang dinilai telah mencapai KKM (&ge; 75) — Daya serap 100%.
                              </td>
                            </tr>
                          ) : (
                            rem.remedialList.map((r) => (
                              <tr key={`${cls}-${subj}-rem-${r.nis}-${r.no}`}>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{r.no}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.nis}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.name}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{r.maxScore}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{r.minScore}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px' }}>{r.belowItems}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', fontWeight: 700 }}>{r.avg.toFixed(1)}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>-{r.gap.toFixed(1)}</td>
                                <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.rekomendasi}</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  );
                })
              )}
            </div>
          )}

          {/* 7. Leger Gabungan Kelas (Akademik) */}
          {includeGradesLeger && (
            <div style={{ marginBottom: '24px' }}>
              {selectedClasses.map((cls) => {
                const { subjects, rows } = getCombinedLegerForClass(cls);
                return (
                  <div key={`leger-${cls}`} style={{ marginBottom: '18px' }}>
                    <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                      Leger Gabungan Nilai Lintas Mapel & Peringkat — Kelas {cls}
                    </h3>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                      <thead>
                        <tr style={{ backgroundColor: '#f1f5f9' }}>
                          <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', width: '42px' }}>Rank</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NIS</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Siswa</th>
                          {subjects.map((subj) => (
                            <th key={subj} style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                              {subj}
                            </th>
                          ))}
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', whiteSpace: 'nowrap' }}>Jumlah</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', whiteSpace: 'nowrap' }}>Rata-rata</th>
                          <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left', whiteSpace: 'nowrap' }}>Ket.</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.length === 0 ? (
                          <tr>
                            <td colSpan={6 + subjects.length} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                              Belum ada siswa di kelas {cls}.
                            </td>
                          </tr>
                        ) : (
                          rows.map((r) => (
                            <tr key={r.id}>
                              <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', fontWeight: 700 }}>
                                {r.rank ?? '-'}
                              </td>
                              <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.nis}</td>
                              <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.name}</td>
                              {subjects.map((subj) => (
                                <td key={subj} style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>
                                  {r.subjAvgs[subj] !== null ? (r.subjAvgs[subj] as number).toFixed(1) : '-'}
                                </td>
                              ))}
                              <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>
                                {r.totalScore !== null ? r.totalScore.toFixed(1) : '-'}
                              </td>
                              <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', fontWeight: 700 }}>
                                {r.overallAvg !== null ? r.overallAvg.toFixed(1) : '-'}
                              </td>
                              <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.status}</td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
          )}

          {/* 8. Daftar Peserta Didik (Kesiswaan) */}
          {includeStudentsRoster && (
            <div style={{ marginBottom: '24px' }}>
              <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                Daftar Induk Peserta Didik & Wali Kelas ({studentRosterRows.length} Siswa)
              </h3>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10px' }}>
                <thead>
                  <tr style={{ backgroundColor: '#f1f5f9' }}>
                    <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', width: '26px' }}>No</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>NIS / NISN</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', textAlign: 'left' }}>Nama Siswa</th>
                    <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>L/P</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>Kelas</th>
                    <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Abs</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', textAlign: 'left' }}>Tempat, Tgl Lahir</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>No. HP</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', textAlign: 'left' }}>Wali Kelas</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', textAlign: 'left' }}>Guru BK</th>
                    <th style={{ border: '1px solid #000', padding: '4px 5px', textAlign: 'left' }}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {studentRosterRows.length === 0 ? (
                    <tr>
                      <td colSpan={11} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                        Belum ada data peserta didik pada kelas terpilih.
                      </td>
                    </tr>
                  ) : (
                    studentRosterRows.map((st, idx) => (
                      <tr key={st.id}>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{idx + 1}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>
                          {st.nis} / {st.nisn || '-'}
                        </td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap', fontWeight: 600 }}>{st.name}</td>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{st.gender}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>{st.className}</td>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{st.absentNo || '-'}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px' }}>
                          {st.birthPlace || '-'}, {st.birthDate || '-'}
                        </td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>{st.phone || '-'}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px' }}>{st.homeroomTeacher || '-'}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px' }}>{st.counselorTeacher || '-'}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 5px', whiteSpace: 'nowrap' }}>{st.status}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* 9. Aktivitas & Penilaian Guru (Pimpinan) */}
          {includeTeacherActivity && (
            <div style={{ marginBottom: '24px' }}>
              <div style={{ marginBottom: '14px' }}>
                <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                  Ringkasan Aktivitas & Penilaian Kinerja Guru
                </h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f1f5f9' }}>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '28px' }}>No</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NIP</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Nama Guru</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Jabatan / Mapel</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Jadwal</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Jurnal</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>% Hadir</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Dinilai</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>Rata Nilai</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Status Kinerja</th>
                    </tr>
                  </thead>
                  <tbody>
                    {teacherActivityData.summary.map((t) => (
                      <tr key={`${t.no}-${t.name}`}>
                        <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{t.no}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{t.nip}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap', fontWeight: 600 }}>{t.name}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 6px' }}>
                          {t.position} ({t.subject})
                        </td>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{t.scheduleCount}</td>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', fontWeight: 700 }}>{t.journalCount}</td>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>
                          {t.avgAttendance !== null ? `${t.avgAttendance.toFixed(0)}%` : '-'}
                        </td>
                        <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{t.assessedStudents}</td>
                        <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>
                          {t.avgGrade !== null ? t.avgGrade.toFixed(1) : '-'}
                        </td>
                        <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{t.kinerja}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div>
                <h3 style={{ fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>
                  Riwayat Kinerja & Pelaksanaan KBM Terbaru
                </h3>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '10.5px' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f1f5f9' }}>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '28px' }}>No</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Tanggal</th>
                      <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>Jam</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Kelas</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>Mata Pelajaran</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'left' }}>Materi Pokok</th>
                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right' }}>% Hadir</th>
                    </tr>
                  </thead>
                  <tbody>
                    {teacherActivityData.recentKbm.length === 0 ? (
                      <tr>
                        <td colSpan={7} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                          Belum ada riwayat pelaksanaan KBM pada periode filter ini.
                        </td>
                      </tr>
                    ) : (
                      teacherActivityData.recentKbm.map((r) => (
                        <tr key={`kbm-${r.no}-${r.date}-${r.className}`}>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center' }}>{r.no}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{formatShortDateId(r.date)}</td>
                          <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }}>{r.jamKe}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.className}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>{r.subject}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px' }}>{r.topic}</td>
                          <td style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'right', fontWeight: 700 }}>
                            {r.pct.toFixed(0)}%
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Blok Pengesahan (Tanda Tangan + Stempel + Posisi Fleksibel & Seret Bebas) */}
          {(showTtdKepsek || showTtdGuru) && (
            <div style={{ marginTop: '28px', paddingTop: '8px' }}>
              {sigPos.preset === 'bebas' ? (
                <div
                  ref={sigCanvasRef}
                  data-sig-canvas="true"
                  onPointerMove={handleSigPointerMove}
                  onPointerUp={handleSigPointerUp}
                  style={{
                    position: 'relative',
                    minHeight: `${Math.max(215, Math.max(sigPos.principalY, sigPos.teacherY) + 155)}px`,
                    width: '100%',
                  }}
                >
                  {showTtdKepsek && (
                    <div
                      data-sig-block="kepsek"
                      onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                      title="Seret ke kanan / kiri untuk mengatur posisi Kepala Sekolah"
                      style={{
                        position: 'absolute',
                        left: `${sigPos.principalX}%`,
                        top: `${sigPos.principalY}px`,
                        width: '220px',
                        cursor: dragTarget === 'kepsek' ? 'grabbing' : 'grab',
                        userSelect: 'none',
                        touchAction: 'none',
                        padding: '4px',
                        borderRadius: '6px',
                      }}
                      className="group hover:outline-1 hover:outline-dashed hover:outline-indigo-500"
                    >
                      <span
                        data-export-ignore="true"
                        className="no-print opacity-0 group-hover:opacity-100 transition-opacity absolute -top-5 left-0 px-1.5 py-0.5 rounded bg-indigo-600 text-white text-[10px] font-sans font-semibold whitespace-nowrap pointer-events-none"
                      >
                        Geser Kiri/Kanan ({sigPos.principalX}%)
                      </span>
                      {renderSigCard(
                        'Mengetahui,',
                        settings.principalTitle,
                        settings.principalName,
                        settings.principalNip,
                        showSignatureImage ? settings.principalSignatureImage : '',
                        stampOnKepsek,
                        kepsekSigSizePx
                      )}
                    </div>
                  )}
                  {showTtdGuru && (
                    <div
                      data-sig-block="guru"
                      onPointerDown={(e) => handleSigPointerDown('guru', e)}
                      title="Seret ke kanan / kiri untuk mengatur posisi Guru / Pemilik Akun"
                      style={{
                        position: 'absolute',
                        left: `${sigPos.teacherX}%`,
                        top: `${sigPos.teacherY}px`,
                        width: '220px',
                        cursor: dragTarget === 'guru' ? 'grabbing' : 'grab',
                        userSelect: 'none',
                        touchAction: 'none',
                        padding: '4px',
                        borderRadius: '6px',
                      }}
                      className="group hover:outline-1 hover:outline-dashed hover:outline-indigo-500"
                    >
                      <span
                        data-export-ignore="true"
                        className="no-print opacity-0 group-hover:opacity-100 transition-opacity absolute -top-5 left-0 px-1.5 py-0.5 rounded bg-indigo-600 text-white text-[10px] font-sans font-semibold whitespace-nowrap pointer-events-none"
                      >
                        Geser Kiri/Kanan ({sigPos.teacherX}%)
                      </span>
                      {renderSigCard(
                        `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                        activeSigner.roleLabel,
                        activeSigner.name,
                        activeSigner.nip,
                        showSignatureImage ? activeSigner.signatureImage : '',
                        stampOnGuru,
                        guruSigSizePx
                      )}
                    </div>
                  )}
                </div>
              ) : (
                <div
                  ref={sigCanvasRef}
                  onPointerMove={handleSigPointerMove}
                  onPointerUp={handleSigPointerUp}
                  style={{ width: '100%' }}
                >
                  {sigPos.preset === 'kiri' ? (
                    <div style={{ paddingLeft: '16px' }}>
                      {showTtdKepsek && (
                        <div
                          onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                          title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                          style={{ width: '220px', marginBottom: '20px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                          className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                        >
                          {renderSigCard(
                            'Mengetahui,',
                            settings.principalTitle,
                            settings.principalName,
                            settings.principalNip,
                            showSignatureImage ? settings.principalSignatureImage : '',
                            stampOnKepsek,
                            kepsekSigSizePx
                          )}
                        </div>
                      )}
                      {showTtdGuru && (
                        <div
                          onPointerDown={(e) => handleSigPointerDown('guru', e)}
                          title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                          style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                          className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                        >
                          {renderSigCard(
                            `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                            activeSigner.roleLabel,
                            activeSigner.name,
                            activeSigner.nip,
                            showSignatureImage ? activeSigner.signatureImage : '',
                            stampOnGuru,
                            guruSigSizePx
                          )}
                        </div>
                      )}
                    </div>
                  ) : sigPos.preset === 'kanan' ? (
                    <div style={{ marginLeft: 'auto', width: '240px' }}>
                      {showTtdKepsek && (
                        <div
                          onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                          title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                          style={{ marginBottom: '20px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                          className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                        >
                          {renderSigCard(
                            'Mengetahui,',
                            settings.principalTitle,
                            settings.principalName,
                            settings.principalNip,
                            showSignatureImage ? settings.principalSignatureImage : '',
                            stampOnKepsek,
                            kepsekSigSizePx
                          )}
                        </div>
                      )}
                      {showTtdGuru && (
                        <div
                          onPointerDown={(e) => handleSigPointerDown('guru', e)}
                          title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                          style={{ cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                          className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                        >
                          {renderSigCard(
                            `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                            activeSigner.roleLabel,
                            activeSigner.name,
                            activeSigner.nip,
                            showSignatureImage ? activeSigner.signatureImage : '',
                            stampOnGuru,
                            guruSigSizePx
                          )}
                        </div>
                      )}
                    </div>
                  ) : sigPos.preset === 'tengah' ? (
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center' }}>
                      <tbody>
                        <tr>
                          {showTtdKepsek && (
                            <td style={{ verticalAlign: 'top', textAlign: 'center', border: 'none' }}>
                              <div
                                onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                                style={{ display: 'inline-block', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  'Mengetahui,',
                                  settings.principalTitle,
                                  settings.principalName,
                                  settings.principalNip,
                                  showSignatureImage ? settings.principalSignatureImage : '',
                                  stampOnKepsek,
                                  kepsekSigSizePx
                                )}
                              </div>
                            </td>
                          )}
                          {showTtdGuru && (
                            <td style={{ verticalAlign: 'top', textAlign: 'center', border: 'none' }}>
                              <div
                                onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                                style={{ display: 'inline-block', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                  activeSigner.roleLabel,
                                  activeSigner.name,
                                  activeSigner.nip,
                                  showSignatureImage ? activeSigner.signatureImage : '',
                                  stampOnGuru,
                                  guruSigSizePx
                                )}
                              </div>
                            </td>
                          )}
                        </tr>
                      </tbody>
                    </table>
                  ) : sigPos.preset === 'kanan-kiri' ? (
                    <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                      <tbody>
                        <tr>
                          <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '16px', border: 'none' }}>
                            {showTtdGuru && (
                              <div
                                onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                                style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                  activeSigner.roleLabel,
                                  activeSigner.name,
                                  activeSigner.nip,
                                  showSignatureImage ? activeSigner.signatureImage : '',
                                  stampOnGuru,
                                  guruSigSizePx
                                )}
                              </div>
                            )}
                          </td>
                          <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '18%', border: 'none' }}>
                            {showTtdKepsek && (
                              <div
                                onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                                style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  'Mengetahui,',
                                  settings.principalTitle,
                                  settings.principalName,
                                  settings.principalNip,
                                  showSignatureImage ? settings.principalSignatureImage : '',
                                  stampOnKepsek,
                                  kepsekSigSizePx
                                )}
                              </div>
                            )}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                      <tbody>
                        <tr>
                          <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '16px', border: 'none' }}>
                            {showTtdKepsek && (
                              <div
                                onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                                style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  'Mengetahui,',
                                  settings.principalTitle,
                                  settings.principalName,
                                  settings.principalNip,
                                  showSignatureImage ? settings.principalSignatureImage : '',
                                  stampOnKepsek,
                                  kepsekSigSizePx
                                )}
                              </div>
                            )}
                          </td>
                          <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '18%', border: 'none' }}>
                            {showTtdGuru && (
                              <div
                                onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                title="Seret ke kanan / kiri untuk menyesuaikan posisi"
                                style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                  activeSigner.roleLabel,
                                  activeSigner.name,
                                  activeSigner.nip,
                                  showSignatureImage ? activeSigner.signatureImage : '',
                                  stampOnGuru,
                                  guruSigSizePx
                                )}
                              </div>
                            )}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      </section>
    </div>
  );
}
