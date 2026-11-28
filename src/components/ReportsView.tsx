import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  BarChart3,
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
  ReportChartId,
  ReportPrintPreferences,
  ReportSectionId,
  SemesterType,
  SignatureLayoutPreset,
  SignaturePositionSettings,
  SystemSettings,
  Teacher,
} from '../types';
import {
  PAPER_SIZE_CONFIG,
  drawPdfKopSurat,
  exportPreviewDomToWordDoc,
  resolveActiveSigner,
  resolvePrincipalStampImage,
} from '../utils/exportHelpers';
import {
  compareClassNames,
  compareStudentsByClassAndAbsent,
  DEFAULT_KOP_LINE_SPACING,
  fitContainDimensions,
  formatCompactDateId,
  formatLongDateId,
  formatShortDateId,
  formatTeacherAssignmentsSummary,
  getAllTeacherClasses,
  getPrintFontCss,
  getTeacherClassesForSubject,
  getTeacherSubjects,
  getTodayISO,
  imageUrlToPngWithMeta,
  KOP_LINE_KEYS,
  makeGradeId,
  normalizeKopLineSpacing,
  PngImageInfo,
  resolveKopLineStyle,
  summarizeJournalRow,
} from '../utils/portalHelpers';
import { buildReportCharts, REPORT_CHART_OPTIONS } from '../utils/reportCharts';

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
  {
    id: 'charts',
    label: 'Laporan Grafis & Diagram',
    category: 'Pimpinan',
    description: 'Diagram batang, garis, donat, lingkaran & batang horizontal',
    officialTitle: 'LAPORAN ANALISIS GRAFIS & DIAGRAM KBM',
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
  draftTeacherPrint?: { teacherId: string; nonce: number } | null;
  onClearDraftTeacherPrint?: () => void;
  onUpdateSettings: (updater: (prev: SystemSettings) => SystemSettings) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
}

export function ReportsView({
  db,
  draftTeacherPrint,
  onClearDraftTeacherPrint,
  onUpdateSettings,
  onToast,
}: ReportsViewProps) {
  const { students, teachers, schedules, journals, grades, accounts, settings, session } = db;

  const classOptions = useMemo(() => {
    const set = new Set(settings.masterClasses);
    students.forEach((s) => {
      if (s.className) set.add(s.className);
    });
    return Array.from(set).sort(compareClassNames);
  }, [settings.masterClasses, students]);

  const subjectOptions = useMemo(() => {
    const set = new Set(settings.masterSubjects);
    teachers.forEach((t) => {
      getTeacherSubjects(t).forEach((subj) => set.add(subj));
    });
    journals.forEach((j) => {
      if (j.subject) set.add(j.subject);
    });
    grades.forEach((g) => {
      if (g.subject) set.add(g.subject);
    });
    return Array.from(set);
  }, [settings.masterSubjects, teachers, journals, grades]);

  const resolveSubjectsFromPrefs = (
    prefs: ReportPrintPreferences | undefined,
    allSubjects: string[]
  ): string[] => {
    if (prefs?.selectedSubjects && prefs.selectedSubjects.length > 0) {
      return prefs.selectedSubjects;
    }
    if (prefs?.subjectFilter && prefs.subjectFilter !== 'ALL') {
      return [prefs.subjectFilter];
    }
    return allSubjects;
  };

  const savedPrefs = settings.reportPrintPreferences;

  // Filters & Print Preferences (initialized from persisted settings)
  const [selectedReportTypes, setSelectedReportTypes] = useState<ReportSectionId[]>(() =>
    savedPrefs?.selectedReportTypes && savedPrefs.selectedReportTypes.length > 0
      ? savedPrefs.selectedReportTypes
      : ['journal']
  );
  const [selectedChartTypes, setSelectedChartTypes] = useState<ReportChartId[]>(() =>
    savedPrefs?.selectedChartTypes && savedPrefs.selectedChartTypes.length > 0
      ? savedPrefs.selectedChartTypes
      : REPORT_CHART_OPTIONS.map((c) => c.id)
  );
  const [selectedClasses, setSelectedClasses] = useState<string[]>(() =>
    savedPrefs?.selectedClasses && savedPrefs.selectedClasses.length > 0
      ? [...savedPrefs.selectedClasses].sort(compareClassNames)
      : classOptions
  );
  const [selectedSubjects, setSelectedSubjects] = useState<string[]>(() =>
    resolveSubjectsFromPrefs(savedPrefs, subjectOptions)
  );
  const [startDate, setStartDate] = useState<string>(() => savedPrefs?.startDate || '');
  const [endDate, setEndDate] = useState<string>(() => savedPrefs?.endDate || '');
  const [reportAcademicYear, setReportAcademicYear] = useState<string>(
    () => session.academicYear || settings.academicYear
  );
  const [reportSemester, setReportSemester] = useState<SemesterType>(
    () => session.semester || settings.semester
  );
  const [signerAccountId, setSignerAccountId] = useState<string>(
    () => savedPrefs?.signerAccountId ?? ''
  );
  const [selectedTeacherId, setSelectedTeacherId] = useState<string>(
    () => savedPrefs?.selectedTeacherId ?? ''
  );

  const activeTeacher = useMemo(
    () => (selectedTeacherId ? teachers.find((t) => t.id === selectedTeacherId) ?? null : null),
    [teachers, selectedTeacherId]
  );

  const reportYearOptions = useMemo(() => {
    const set = new Set(settings.masterAcademicYears);
    if (session.academicYear) set.add(session.academicYear);
    if (settings.academicYear) set.add(settings.academicYear);
    journals.forEach((j) => {
      if (j.academicYear) set.add(j.academicYear);
    });
    grades.forEach((g) => {
      if (g.academicYear) set.add(g.academicYear);
    });
    schedules.forEach((sc) => {
      if (sc.academicYear) set.add(sc.academicYear);
    });
    return Array.from(set).filter(Boolean).sort();
  }, [settings.masterAcademicYears, settings.academicYear, session.academicYear, journals, grades, schedules]);

  useEffect(() => {
    if (session.academicYear) setReportAcademicYear(session.academicYear);
    if (session.semester) setReportSemester(session.semester);
  }, [session.academicYear, session.semester]);

  // Multi-select dropdown states
  const [reportDropdownOpen, setReportDropdownOpen] = useState(false);
  const [classDropdownOpen, setClassDropdownOpen] = useState(false);
  const [subjectDropdownOpen, setSubjectDropdownOpen] = useState(false);
  const reportDropdownRef = useRef<HTMLDivElement | null>(null);
  const dropdownRef = useRef<HTMLDivElement | null>(null);
  const subjectDropdownRef = useRef<HTMLDivElement | null>(null);
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
  const [paperSize, setPaperSize] = useState<PaperSizeId | ''>(() => savedPrefs?.paperSize || 'A4');
  const [fontFamily, setFontFamily] = useState<PrintFontFamily | ''>(
    () => savedPrefs?.fontFamily || 'Times New Roman'
  );
  const effectivePaperSize: PaperSizeId = paperSize || 'A4';
  const effectiveFontFamily: PrintFontFamily = fontFamily || 'Times New Roman';
  const [sigPos, setSigPos] = useState<SignaturePositionSettings>(
    () => settings.signaturePosition
  );
  const [kopLeftSize, setKopLeftSize] = useState<number>(() =>
    Math.max(28, Math.min(150, settings.kopSurat.logoLeftSize ?? 72))
  );
  const [kopRightSize, setKopRightSize] = useState<number>(() =>
    Math.max(28, Math.min(150, settings.kopSurat.logoRightSize ?? 72))
  );
  const [kopLineSpacing, setKopLineSpacing] = useState<number>(() =>
    normalizeKopLineSpacing(settings.kopSurat.lineSpacing)
  );
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);
  const [isExportingWord, setIsExportingWord] = useState(false);
  const [filterPanelExpanded, setFilterPanelExpanded] = useState(false);
  const [sigPosExpanded, setSigPosExpanded] = useState(false);
  const [kopLogoExpanded, setKopLogoExpanded] = useState(false);

  const prevSettingsRef = useRef(settings);
  const prevClassOptionsRef = useRef<string[]>(classOptions);
  const prevSubjectOptionsRef = useRef<string[]>(subjectOptions);

  useEffect(() => {
    const prevClassOptions = prevClassOptionsRef.current;
    if (prevClassOptions === classOptions) return;
    prevClassOptionsRef.current = classOptions;
    if (!settings.reportPrintPreferences.selectedClasses?.length) {
      const wasAllSelected =
        selectedClasses.length === prevClassOptions.length &&
        selectedClasses.every((c, i) => c === prevClassOptions[i]);
      if (wasAllSelected) {
        setSelectedClasses(classOptions);
      }
    }
  }, [classOptions, selectedClasses, settings.reportPrintPreferences.selectedClasses]);

  useEffect(() => {
    const prevSubjOptions = prevSubjectOptionsRef.current;
    if (prevSubjOptions === subjectOptions) return;
    prevSubjectOptionsRef.current = subjectOptions;
    if (
      !settings.reportPrintPreferences.selectedSubjects?.length &&
      (!settings.reportPrintPreferences.subjectFilter ||
        settings.reportPrintPreferences.subjectFilter === 'ALL')
    ) {
      const wasAllSelected =
        selectedSubjects.length === prevSubjOptions.length &&
        selectedSubjects.every((s, i) => s === prevSubjOptions[i]);
      if (wasAllSelected) {
        setSelectedSubjects(subjectOptions);
      }
    }
  }, [
    subjectOptions,
    selectedSubjects,
    settings.reportPrintPreferences.selectedSubjects,
    settings.reportPrintPreferences.subjectFilter,
  ]);

  useEffect(() => {
    const prev = prevSettingsRef.current;
    if (prev === settings) return;

    const prevPrefs = prev.reportPrintPreferences;
    const nextPrefs = settings.reportPrintPreferences;
    const savedChartsPrev =
      prevPrefs.selectedChartTypes && prevPrefs.selectedChartTypes.length > 0
        ? prevPrefs.selectedChartTypes
        : REPORT_CHART_OPTIONS.map((c) => c.id);
    const savedClassesPrev =
      prevPrefs.selectedClasses && prevPrefs.selectedClasses.length > 0
        ? [...prevPrefs.selectedClasses].sort(compareClassNames)
        : classOptions;
    const savedSubjectsPrev = resolveSubjectsFromPrefs(prevPrefs, subjectOptions);

    if (
      selectedReportTypes.length === prevPrefs.selectedReportTypes.length &&
      selectedReportTypes.every((r, i) => r === prevPrefs.selectedReportTypes[i])
    ) {
      setSelectedReportTypes(
        nextPrefs.selectedReportTypes.length > 0 ? nextPrefs.selectedReportTypes : ['journal']
      );
    }
    if (
      selectedChartTypes.length === savedChartsPrev.length &&
      selectedChartTypes.every((c, i) => c === savedChartsPrev[i])
    ) {
      setSelectedChartTypes(
        nextPrefs.selectedChartTypes && nextPrefs.selectedChartTypes.length > 0
          ? nextPrefs.selectedChartTypes
          : REPORT_CHART_OPTIONS.map((c) => c.id)
      );
    }
    if (
      selectedClasses.length === savedClassesPrev.length &&
      selectedClasses.every((c, i) => c === savedClassesPrev[i])
    ) {
      setSelectedClasses(
        nextPrefs.selectedClasses && nextPrefs.selectedClasses.length > 0
          ? [...nextPrefs.selectedClasses].sort(compareClassNames)
          : classOptions
      );
    }
    if (
      selectedSubjects.length === savedSubjectsPrev.length &&
      selectedSubjects.every((s, i) => s === savedSubjectsPrev[i])
    ) {
      setSelectedSubjects(resolveSubjectsFromPrefs(nextPrefs, subjectOptions));
    }
    if (startDate === prevPrefs.startDate) setStartDate(nextPrefs.startDate);
    if (endDate === prevPrefs.endDate) setEndDate(nextPrefs.endDate);
    if (signerAccountId === prevPrefs.signerAccountId) setSignerAccountId(nextPrefs.signerAccountId);
    if (selectedTeacherId === (prevPrefs.selectedTeacherId ?? '')) {
      setSelectedTeacherId(nextPrefs.selectedTeacherId ?? '');
    }
    if (showKop === prevPrefs.showKop) setShowKop(nextPrefs.showKop);
    if (showTtdKepsek === prevPrefs.showTtdKepsek) setShowTtdKepsek(nextPrefs.showTtdKepsek);
    if (showTtdGuru === prevPrefs.showTtdGuru) setShowTtdGuru(nextPrefs.showTtdGuru);
    if (showSignatureImage === prevPrefs.showSignatureImage) {
      setShowSignatureImage(nextPrefs.showSignatureImage);
    }
    if (orientation === prevPrefs.orientation) setOrientation(nextPrefs.orientation);
    if (paperSize === prevPrefs.paperSize) setPaperSize(nextPrefs.paperSize);
    if (fontFamily === prevPrefs.fontFamily) setFontFamily(nextPrefs.fontFamily);
    if (JSON.stringify(sigPos) === JSON.stringify(prev.signaturePosition)) {
      setSigPos(settings.signaturePosition);
    }

    const prevLeft = Math.max(28, Math.min(150, prev.kopSurat.logoLeftSize ?? 72));
    const prevRight = Math.max(28, Math.min(150, prev.kopSurat.logoRightSize ?? 72));
    const prevSpacing = normalizeKopLineSpacing(prev.kopSurat.lineSpacing);
    if (kopLeftSize === prevLeft) {
      setKopLeftSize(Math.max(28, Math.min(150, settings.kopSurat.logoLeftSize ?? 72)));
    }
    if (kopRightSize === prevRight) {
      setKopRightSize(Math.max(28, Math.min(150, settings.kopSurat.logoRightSize ?? 72)));
    }
    if (kopLineSpacing === prevSpacing) {
      setKopLineSpacing(normalizeKopLineSpacing(settings.kopSurat.lineSpacing));
    }

    prevSettingsRef.current = settings;
  }, [
    settings,
    classOptions,
    subjectOptions,
    selectedReportTypes,
    selectedChartTypes,
    selectedClasses,
    selectedSubjects,
    startDate,
    endDate,
    signerAccountId,
    selectedTeacherId,
    showKop,
    showTtdKepsek,
    showTtdGuru,
    showSignatureImage,
    orientation,
    paperSize,
    fontFamily,
    sigPos,
    kopLeftSize,
    kopRightSize,
    kopLineSpacing,
  ]);

  // Interactive drag & resize state for signature blocks, signature images, and stamp
  type SigInteraction =
    | {
        kind: 'block';
        target: 'kepsek' | 'guru';
        startClientX: number;
        startClientY: number;
        initX: number;
        initY: number;
      }
    | {
        kind: 'move-kepsek-sig' | 'move-guru-sig' | 'move-stamp';
        startClientX: number;
        startClientY: number;
        initX: number;
        initY: number;
      }
    | {
        kind: 'resize-kepsek-sig' | 'resize-guru-sig' | 'resize-stamp';
        startClientX: number;
        startClientY: number;
        initSize: number;
      };

  const [activeInteraction, setActiveInteraction] = useState<SigInteraction | null>(null);
  const [selectedMedia, setSelectedMedia] = useState<'kepsek-sig' | 'guru-sig' | 'stamp' | null>(
    null
  );

  const findAccountForTeacher = (teacher: Teacher) => {
    const tEmail = (teacher.email || '').trim().toLowerCase();
    const tNip = (teacher.nip || '').trim().toLowerCase();
    const tName = (teacher.name || '').trim().toLowerCase();
    return accounts.find((acc) => {
      if (tEmail && acc.email.trim().toLowerCase() === tEmail) return true;
      if (tNip && acc.nip.trim().toLowerCase() === tNip) return true;
      if (tName && acc.name.trim().toLowerCase() === tName) return true;
      return false;
    });
  };

  const findTeacherForAccount = (accountId: string) => {
    const acc = accounts.find((a) => a.id === accountId);
    if (!acc) return undefined;
    const aEmail = (acc.email || '').trim().toLowerCase();
    const aNip = (acc.nip || '').trim().toLowerCase();
    const aName = (acc.name || '').trim().toLowerCase();
    return teachers.find((t) => {
      if (aEmail && (t.email || '').trim().toLowerCase() === aEmail) return true;
      if (aNip && (t.nip || '').trim().toLowerCase() === aNip) return true;
      if (aName && (t.name || '').trim().toLowerCase() === aName) return true;
      return false;
    });
  };

  const applyTeacherPrintPreset = (teacherId: string, announce = true) => {
    if (!teacherId) {
      setSelectedTeacherId('');
      setSelectedClasses(classOptions);
      setSelectedSubjects(subjectOptions);
      if (signerAccountId !== '__AUTO__') {
        setSignerAccountId('');
      }
      if (announce) {
        onToast('info', 'Filter guru direset ke semua kelas dan mata pelajaran.');
      }
      return;
    }
    const teacher = teachers.find((t) => t.id === teacherId);
    if (!teacher) return;
    const tSubjects = getTeacherSubjects(teacher);
    const tClasses = getAllTeacherClasses(teacher).sort(compareClassNames);
    setSelectedTeacherId(teacher.id);
    setSelectedSubjects(tSubjects.length > 0 ? tSubjects : subjectOptions);
    setSelectedClasses(tClasses.length > 0 ? tClasses : classOptions);
    const matchedAcc = findAccountForTeacher(teacher);
    setSignerAccountId(matchedAcc ? matchedAcc.id : '');
    if (announce) {
      const detail =
        tSubjects.length > 0 || tClasses.length > 0
          ? `${tSubjects.length} mapel & ${tClasses.length} kelas`
          : 'belum ada mapel/kelas khusus';
      onToast('success', `Siap cetak untuk ${teacher.name} (${detail}).`);
    }
  };

  const lastDraftNonceRef = useRef<number | null>(null);
  useEffect(() => {
    if (!draftTeacherPrint) return;
    if (lastDraftNonceRef.current === draftTeacherPrint.nonce) return;
    lastDraftNonceRef.current = draftTeacherPrint.nonce;
    applyTeacherPrintPreset(draftTeacherPrint.teacherId, false);
    onClearDraftTeacherPrint?.();
  }, [draftTeacherPrint]);

  const activeSigner = useMemo(() => {
    if (signerAccountId && signerAccountId !== '__AUTO__') {
      return resolveActiveSigner(db, signerAccountId);
    }
    const resolved = resolveActiveSigner(db, undefined);
    if (activeTeacher) {
      const matchedAcc = findAccountForTeacher(activeTeacher);
      if (matchedAcc) {
        return resolveActiveSigner(db, matchedAcc.id);
      }
      return {
        ...resolved,
        name: activeTeacher.name || resolved.name,
        nip: activeTeacher.nip || resolved.nip,
        roleLabel:
          activeTeacher.position && activeTeacher.position !== 'Admin / Operator Sekolah'
            ? activeTeacher.position
            : resolved.roleLabel,
        signatureImage: '',
      };
    }
    return resolved;
  }, [db, signerAccountId, activeTeacher, accounts]);

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
    const savedCharts =
      prefs.selectedChartTypes && prefs.selectedChartTypes.length > 0
        ? prefs.selectedChartTypes
        : REPORT_CHART_OPTIONS.map((c) => c.id);
    const chartTypesChanged =
      selectedChartTypes.length !== savedCharts.length ||
      selectedChartTypes.some((c, i) => c !== savedCharts[i]);
    const savedClasses =
      prefs.selectedClasses && prefs.selectedClasses.length > 0
        ? [...prefs.selectedClasses].sort(compareClassNames)
        : classOptions;
    const classesChanged =
      selectedClasses.length !== savedClasses.length ||
      selectedClasses.some((c, i) => c !== savedClasses[i]);
    const savedSubjects = resolveSubjectsFromPrefs(prefs, subjectOptions);
    const subjectsChanged =
      selectedSubjects.length !== savedSubjects.length ||
      selectedSubjects.some((s, i) => s !== savedSubjects[i]);
    const prefsChanged =
      reportTypesChanged ||
      chartTypesChanged ||
      classesChanged ||
      subjectsChanged ||
      startDate !== prefs.startDate ||
      endDate !== prefs.endDate ||
      signerAccountId !== prefs.signerAccountId ||
      selectedTeacherId !== (prefs.selectedTeacherId ?? '') ||
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
      (sigPos.stampSize ?? 70) !== (savedSig.stampSize ?? 70) ||
      (sigPos.principalSigOffsetX ?? 0) !== (savedSig.principalSigOffsetX ?? 0) ||
      (sigPos.principalSigOffsetY ?? 0) !== (savedSig.principalSigOffsetY ?? 0) ||
      (sigPos.teacherSigOffsetX ?? 0) !== (savedSig.teacherSigOffsetX ?? 0) ||
      (sigPos.teacherSigOffsetY ?? 0) !== (savedSig.teacherSigOffsetY ?? 0) ||
      (sigPos.stampOffsetX ?? 0) !== (savedSig.stampOffsetX ?? 0) ||
      (sigPos.stampOffsetY ?? 0) !== (savedSig.stampOffsetY ?? 0);
    const savedKopLeft = Math.max(28, Math.min(150, settings.kopSurat.logoLeftSize ?? 72));
    const savedKopRight = Math.max(28, Math.min(150, settings.kopSurat.logoRightSize ?? 72));
    const savedKopLineSpacing = normalizeKopLineSpacing(settings.kopSurat.lineSpacing);
    const kopChanged =
      kopLeftSize !== savedKopLeft ||
      kopRightSize !== savedKopRight ||
      kopLineSpacing !== savedKopLineSpacing;
    return prefsChanged || sigChanged || kopChanged;
  }, [
    settings.reportPrintPreferences,
    settings.signaturePosition,
    settings.kopSurat.logoLeftSize,
    settings.kopSurat.logoRightSize,
    settings.kopSurat.lineSpacing,
    classOptions,
    subjectOptions,
    selectedReportTypes,
    selectedChartTypes,
    selectedClasses,
    selectedSubjects,
    startDate,
    endDate,
    signerAccountId,
    selectedTeacherId,
    showKop,
    showTtdKepsek,
    showTtdGuru,
    showSignatureImage,
    orientation,
    paperSize,
    fontFamily,
    sigPos,
    kopLeftSize,
    kopRightSize,
    kopLineSpacing,
  ]);

  const handleSavePrintSettings = () => {
    if (!paperSize) {
      onToast('error', 'Pilih Ukuran Kertas terlebih dahulu.');
      return;
    }
    if (!fontFamily) {
      onToast('error', 'Pilih Jenis Huruf Cetak terlebih dahulu.');
      return;
    }
    if (!sigPos.preset) {
      onToast('error', 'Pilih Posisi Blok Pengesahan terlebih dahulu.');
      return;
    }
    const safeLeftSize = Math.max(28, Math.min(150, kopLeftSize));
    const safeRightSize = Math.max(28, Math.min(150, kopRightSize));
    const safeLineSpacing = normalizeKopLineSpacing(kopLineSpacing);
    onUpdateSettings((prev) => ({
      ...prev,
      kopSurat: {
        ...prev.kopSurat,
        logoLeftSize: safeLeftSize,
        logoRightSize: safeRightSize,
        lineSpacing: safeLineSpacing,
      },
      signaturePosition: sigPos,
      reportPrintPreferences: {
        ...prev.reportPrintPreferences,
        selectedReportTypes,
        selectedChartTypes,
        selectedClasses,
        selectedSubjects,
        subjectFilter: selectedSubjects.length === 1 ? selectedSubjects[0] : 'ALL',
        startDate,
        endDate,
        signerAccountId: signerAccountId === '__AUTO__' ? '' : signerAccountId,
        selectedTeacherId,
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
    setSelectedChartTypes(
      prefs.selectedChartTypes && prefs.selectedChartTypes.length > 0
        ? prefs.selectedChartTypes
        : REPORT_CHART_OPTIONS.map((c) => c.id)
    );
    setSelectedClasses(prefs.selectedClasses.length > 0 ? prefs.selectedClasses : classOptions);
    setSelectedSubjects(resolveSubjectsFromPrefs(prefs, subjectOptions));
    setStartDate(prefs.startDate);
    setEndDate(prefs.endDate);
    setSignerAccountId(prefs.signerAccountId);
    setSelectedTeacherId(prefs.selectedTeacherId ?? '');
    setShowKop(prefs.showKop);
    setShowTtdKepsek(prefs.showTtdKepsek);
    setShowTtdGuru(prefs.showTtdGuru);
    setShowSignatureImage(prefs.showSignatureImage);
    setOrientation(prefs.orientation);
    setPaperSize(prefs.paperSize);
    setFontFamily(prefs.fontFamily);
    setSigPos(settings.signaturePosition);
    setKopLeftSize(Math.max(28, Math.min(150, settings.kopSurat.logoLeftSize ?? 72)));
    setKopRightSize(Math.max(28, Math.min(150, settings.kopSurat.logoRightSize ?? 72)));
    setKopLineSpacing(normalizeKopLineSpacing(settings.kopSurat.lineSpacing));
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

  const handlePresetChange = (preset: SignatureLayoutPreset | '') => {
    if (!preset) {
      updateSigPos({ preset: '' as unknown as SignatureLayoutPreset });
      return;
    }
    if (preset === 'bebas') {
      updateSigPos({ preset: 'bebas' });
      onToast('info', 'Mode posisi bebas tanda tangan diaktifkan.');
      return;
    }
    const coords = getPresetCoords(preset);
    updateSigPos({ preset, ...coords });
    onToast('info', `Tata letak tanda tangan diatur ke preset "${preset}".`);
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
      if (subjectDropdownRef.current && !subjectDropdownRef.current.contains(e.target as Node)) {
        setSubjectDropdownOpen(false);
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
    const cfg = PAPER_SIZE_CONFIG[effectivePaperSize];
    styleEl.textContent = `@media print { @page { size: ${cfg.cssPageSize} ${orientation}; margin: 12mm; } }`;
  }, [orientation, effectivePaperSize]);

  // Free drag handler inside signature zone (always active even from preset layout)
  const handleSigPointerDown = (
    target: 'kepsek' | 'guru',
    e: React.PointerEvent<HTMLDivElement>
  ) => {
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

    setActiveInteraction({
      kind: 'block',
      target,
      startClientX: e.clientX,
      startClientY: e.clientY,
      initX: target === 'kepsek' ? activeCoords.principalX : activeCoords.teacherX,
      initY: target === 'kepsek' ? activeCoords.principalY : activeCoords.teacherY,
    });
  };

  const startMediaInteraction = (
    kind:
      | 'move-kepsek-sig'
      | 'move-guru-sig'
      | 'move-stamp'
      | 'resize-kepsek-sig'
      | 'resize-guru-sig'
      | 'resize-stamp',
    mediaKey: 'kepsek-sig' | 'guru-sig' | 'stamp',
    e: React.PointerEvent<HTMLElement>
  ) => {
    e.preventDefault();
    e.stopPropagation();
    setSelectedMedia(mediaKey);

    if (kind === 'move-kepsek-sig') {
      setActiveInteraction({
        kind,
        startClientX: e.clientX,
        startClientY: e.clientY,
        initX: sigPos.principalSigOffsetX ?? 0,
        initY: sigPos.principalSigOffsetY ?? 0,
      });
    } else if (kind === 'move-guru-sig') {
      setActiveInteraction({
        kind,
        startClientX: e.clientX,
        startClientY: e.clientY,
        initX: sigPos.teacherSigOffsetX ?? 0,
        initY: sigPos.teacherSigOffsetY ?? 0,
      });
    } else if (kind === 'move-stamp') {
      setActiveInteraction({
        kind,
        startClientX: e.clientX,
        startClientY: e.clientY,
        initX: sigPos.stampOffsetX ?? 0,
        initY: sigPos.stampOffsetY ?? 0,
      });
    } else if (kind === 'resize-kepsek-sig') {
      setActiveInteraction({
        kind,
        startClientX: e.clientX,
        startClientY: e.clientY,
        initSize: sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60,
      });
    } else if (kind === 'resize-guru-sig') {
      setActiveInteraction({
        kind,
        startClientX: e.clientX,
        startClientY: e.clientY,
        initSize: sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60,
      });
    } else {
      setActiveInteraction({
        kind: 'resize-stamp',
        startClientX: e.clientX,
        startClientY: e.clientY,
        initSize: sigPos.stampSize ?? 70,
      });
    }
  };

  useEffect(() => {
    if (!activeInteraction) return;

    const handleWindowPointerMove = (e: PointerEvent) => {
      const dx = e.clientX - activeInteraction.startClientX;
      const dy = e.clientY - activeInteraction.startClientY;

      if (activeInteraction.kind === 'block') {
        const canvasWidth = sigCanvasRef.current?.getBoundingClientRect().width || 800;
        const dxPct = (dx / Math.max(canvasWidth, 1)) * 100;
        const nextX = Math.max(0, Math.min(82, Math.round(activeInteraction.initX + dxPct)));
        const nextY = Math.max(0, Math.min(240, Math.round(activeInteraction.initY + dy)));
        if (activeInteraction.target === 'kepsek') {
          updateSigPos({ preset: 'bebas', principalX: nextX, principalY: nextY });
        } else {
          updateSigPos({ preset: 'bebas', teacherX: nextX, teacherY: nextY });
        }
      } else if (activeInteraction.kind === 'move-kepsek-sig') {
        updateSigPos({
          principalSigOffsetX: Math.max(-450, Math.min(450, Math.round(activeInteraction.initX + dx))),
          principalSigOffsetY: Math.max(-160, Math.min(200, Math.round(activeInteraction.initY + dy))),
        });
      } else if (activeInteraction.kind === 'move-guru-sig') {
        updateSigPos({
          teacherSigOffsetX: Math.max(-450, Math.min(450, Math.round(activeInteraction.initX + dx))),
          teacherSigOffsetY: Math.max(-160, Math.min(200, Math.round(activeInteraction.initY + dy))),
        });
      } else if (activeInteraction.kind === 'move-stamp') {
        updateSigPos({
          stampOffsetX: Math.max(-550, Math.min(550, Math.round(activeInteraction.initX + dx))),
          stampOffsetY: Math.max(-200, Math.min(200, Math.round(activeInteraction.initY + dy))),
        });
      } else if (activeInteraction.kind === 'resize-kepsek-sig') {
        const nextSize = Math.max(
          16,
          Math.min(260, Math.round(activeInteraction.initSize + dx * 0.45 + dy * 0.85))
        );
        updateSigPos({ principalSignatureSize: nextSize });
      } else if (activeInteraction.kind === 'resize-guru-sig') {
        const nextSize = Math.max(
          16,
          Math.min(260, Math.round(activeInteraction.initSize + dx * 0.45 + dy * 0.85))
        );
        updateSigPos({ teacherSignatureSize: nextSize });
      } else if (activeInteraction.kind === 'resize-stamp') {
        const nextSize = Math.max(
          20,
          Math.min(260, Math.round(activeInteraction.initSize + (dx + dy) * 0.65))
        );
        updateSigPos({ stampSize: nextSize });
      }
    };

    const handleWindowPointerUp = () => {
      setActiveInteraction(null);
    };

    window.addEventListener('pointermove', handleWindowPointerMove);
    window.addEventListener('pointerup', handleWindowPointerUp);
    window.addEventListener('pointercancel', handleWindowPointerUp);
    return () => {
      window.removeEventListener('pointermove', handleWindowPointerMove);
      window.removeEventListener('pointerup', handleWindowPointerUp);
      window.removeEventListener('pointercancel', handleWindowPointerUp);
    };
  }, [activeInteraction]);

  const applyDatePreset = (preset: 'thisMonth' | 'lastMonth' | 'reset') => {
    if (preset === 'reset') {
      setStartDate('');
      setEndDate('');
      onToast('info', 'Filter rentang tanggal laporan direset.');
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
      onToast('info', 'Periode laporan diatur ke Bulan Ini.');
    } else {
      const first = new Date(year, month - 1, 1);
      const last = new Date(year, month, 0);
      setStartDate(first.toISOString().slice(0, 10));
      setEndDate(last.toISOString().slice(0, 10));
      onToast('info', 'Periode laporan diatur ke Bulan Lalu.');
    }
  };

  const toggleClassSelection = (cls: string) => {
    setSelectedClasses((prev) =>
      (prev.includes(cls) ? prev.filter((c) => c !== cls) : [...prev, cls]).sort(compareClassNames)
    );
  };

  const toggleSubjectSelection = (subj: string) => {
    setSelectedSubjects((prev) => {
      const nextSet = new Set(prev);
      if (nextSet.has(subj)) {
        nextSet.delete(subj);
      } else {
        nextSet.add(subj);
      }
      const ordered = subjectOptions.filter((s) => nextSet.has(s));
      const extras = Array.from(nextSet).filter((s) => !subjectOptions.includes(s));
      return [...ordered, ...extras];
    });
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

  const toggleChartTypeSelection = (id: ReportChartId) => {
    setSelectedChartTypes((prev) => {
      const nextSet = new Set(prev);
      if (nextSet.has(id)) {
        nextSet.delete(id);
      } else {
        nextSet.add(id);
      }
      return REPORT_CHART_OPTIONS.map((o) => o.id).filter((optId) => nextSet.has(optId));
    });
  };

  const filteredJournals = useMemo(() => {
    const clsSet = new Set(selectedClasses);
    const subjSet = new Set(selectedSubjects);
    return journals
      .filter((j) => {
        if (j.academicYear !== reportAcademicYear || j.semester !== reportSemester) {
          return false;
        }
        if (!clsSet.has(j.className)) return false;
        if (!subjSet.has(j.subject)) return false;
        if (activeTeacher) {
          const assignedClasses = getTeacherClassesForSubject(activeTeacher, j.subject);
          if (assignedClasses.length > 0 && !assignedClasses.includes(j.className)) {
            return false;
          }
        }
        if (startDate && j.date < startDate) return false;
        if (endDate && j.date > endDate) return false;
        return true;
      })
      .sort((a, b) =>
        a.date > b.date ? 1 : a.date < b.date ? -1 : compareClassNames(a.className, b.className)
      );
  }, [
    journals,
    reportAcademicYear,
    reportSemester,
    selectedClasses,
    selectedSubjects,
    activeTeacher,
    startDate,
    endDate,
  ]);

  const activeGradeSubjects = useMemo(() => {
    return selectedSubjects;
  }, [selectedSubjects]);

  const getSubjectsForClassSheet = (cls: string): string[] => {
    if (!activeTeacher) return activeGradeSubjects;
    return activeGradeSubjects.filter((subj) => {
      const assignedClasses = getTeacherClassesForSubject(activeTeacher, subj);
      return assignedClasses.length === 0 || assignedClasses.includes(cls);
    });
  };

  const getAttendanceBlockForClass = (cls: string) => {
    const classStudents = students
      .filter((s) => s.className === cls)
      .sort(compareStudentsByClassAndAbsent);
    const classJournals = filteredJournals.filter((j) => j.className === cls);
    return { classStudents, classJournals };
  };

  const getGradeBlockForClassAndSubject = (cls: string, subj: string) => {
    const classStudents = students
      .filter((s) => s.className === cls)
      .sort(compareStudentsByClassAndAbsent);
    const rows = classStudents.map((st, idx) => {
      const scores: Record<string, number | null> = {};
      const nums: number[] = [];
      settings.masterExamTypes.forEach((ex) => {
        const id = makeGradeId(reportAcademicYear, reportSemester, cls, subj, ex, st.id);
        const found = grades.find((g) => g.id === id);
        const val = found ? found.score : null;
        scores[ex] = val;
        if (val !== null) nums.push(val);
      });
      const avg = nums.length > 0 ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
      return {
        no: idx + 1,
        nis: st.nis,
        nisn: st.nisn || '-',
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
        nisn: st.nisn || '-',
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

  const getMonitoringAbsenceForClass = (cls: string) => ({
    summary: monitoringAbsenceData.summary.filter((it) => it.className === cls),
    chronology: monitoringAbsenceData.chronology.filter((c) => c.className === cls),
  });

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
      .sort(compareStudentsByClassAndAbsent);

    const subjectsWithData = settings.masterSubjects.filter((subj) =>
      grades.some(
        (g) =>
          g.academicYear === reportAcademicYear &&
          g.semester === reportSemester &&
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
          const id = makeGradeId(reportAcademicYear, reportSemester, cls, subj, ex, st.id);
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
      .sort(compareStudentsByClassAndAbsent);
  }, [students, selectedClasses]);

  // 9. Aktivitas & Penilaian Guru (Pimpinan)
  const teacherActivityData = useMemo(() => {
    const periodSchedules = schedules.filter(
      (sc) => sc.academicYear === reportAcademicYear && sc.semester === reportSemester
    );
    const periodJournals = journals.filter((j) => {
      if (j.academicYear !== reportAcademicYear || j.semester !== reportSemester) return false;
      if (startDate && j.date < startDate) return false;
      if (endDate && j.date > endDate) return false;
      return true;
    });
    const periodGrades = grades.filter(
      (g) =>
        g.academicYear === reportAcademicYear &&
        g.semester === reportSemester &&
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
      const tSubjects = getTeacherSubjects(t);
      const tSubjectSet = new Set(tSubjects.map((s) => s.toLowerCase()));
      const assignedPairs = new Set<string>();
      tSubjects.forEach((subj) => {
        getTeacherClassesForSubject(t, subj).forEach((clsName) => {
          assignedPairs.add(`${clsName}__${subj.toLowerCase()}`);
        });
      });
      const schedulePairs = new Set(
        tSchedules.map((sc) => `${sc.className}__${sc.subject.toLowerCase()}`)
      );

      const matchesTeacherRecord = (recClass: string, recSubject: string) => {
        const key = `${recClass}__${recSubject.toLowerCase()}`;
        if (assignedPairs.size > 0 && assignedPairs.has(key)) return true;
        if (schedulePairs.size > 0 && schedulePairs.has(key)) return true;
        if (assignedPairs.size === 0 && schedulePairs.size === 0) {
          return tSubjectSet.has(recSubject.toLowerCase());
        }
        return false;
      };

      const matchedJournals = periodJournals.filter((j) =>
        matchesTeacherRecord(j.className, j.subject)
      );

      let totalH = 0;
      let totalRec = 0;
      matchedJournals.forEach((j) => {
        const s = summarizeJournalRow(j);
        totalH += s.counts.H;
        totalRec += s.total;
      });
      const avgAttendance = totalRec > 0 ? (totalH / totalRec) * 100 : null;

      const matchedGrades = periodGrades.filter((g) =>
        matchesTeacherRecord(g.className, g.subject)
      );
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
        subject: formatTeacherAssignmentsSummary(t) || t.subject || '-',
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
    reportAcademicYear,
    reportSemester,
    session.email,
    startDate,
    endDate,
    activeSigner.name,
    activeSigner.nip,
    activeSigner.roleLabel,
    settings.masterSubjects,
  ]);

  const orderedSelectedReportTypes = useMemo(
    () => REPORT_TYPE_OPTIONS.map((o) => o.id).filter((id) => selectedReportTypes.includes(id)),
    [selectedReportTypes]
  );

  const orderedSelectedClasses = useMemo(
    () => [...selectedClasses].sort(compareClassNames),
    [selectedClasses]
  );

  const reportTitle = useMemo(() => {
    if (orderedSelectedReportTypes.length === 0) return 'LAPORAN ADMINISTRASI MENGAJAR';
    if (orderedSelectedReportTypes.length === REPORT_TYPE_OPTIONS.length) {
      return 'PAKET LENGKAP LAPORAN PORTAL E-GURU';
    }
    if (orderedSelectedReportTypes.length === 1) {
      const opt = REPORT_TYPE_OPTIONS.find((o) => o.id === orderedSelectedReportTypes[0]);
      return opt ? opt.officialTitle : 'LAPORAN ADMINISTRASI MENGAJAR';
    }
    const labels = REPORT_TYPE_OPTIONS.filter((o) => orderedSelectedReportTypes.includes(o.id)).map((o) =>
      o.label.toUpperCase()
    );
    if (labels.length <= 3) return `LAPORAN TERPADU: ${labels.join(' · ')}`;
    return `LAPORAN TERPADU (${labels.length} JENIS): ${labels.slice(0, 2).join(' · ')} & Lainnya`;
  }, [orderedSelectedReportTypes]);

  const reportSlug = useMemo(() => {
    if (orderedSelectedReportTypes.length === REPORT_TYPE_OPTIONS.length) return 'paket_lengkap';
    if (orderedSelectedReportTypes.length === 0) return 'kosong';
    return orderedSelectedReportTypes.slice(0, 3).join('_');
  }, [orderedSelectedReportTypes]);

  const subjectLabel =
    selectedSubjects.length === 0
      ? 'Belum Dipilih'
      : selectedSubjects.length === subjectOptions.length && subjectOptions.length > 0
      ? 'Semua Mata Pelajaran'
      : selectedSubjects.join(', ');

  const includeCharts = orderedSelectedReportTypes.includes('charts');

  const chartCards = useMemo(() => {
    if (!includeCharts) return [];
    return buildReportCharts({
      selectedChartTypes,
      selectedClasses: orderedSelectedClasses,
      selectedSubjects,
      subjectFilter: selectedSubjects.length === 1 ? selectedSubjects[0] : 'ALL',
      students,
      journals: filteredJournals,
      grades,
      academicYear: reportAcademicYear,
      semester: reportSemester,
      fontFamily: effectiveFontFamily,
    });
  }, [
    includeCharts,
    selectedChartTypes,
    orderedSelectedClasses,
    selectedSubjects,
    students,
    filteredJournals,
    grades,
    reportAcademicYear,
    reportSemester,
    effectiveFontFamily,
  ]);

  const dateRangeLabel =
    startDate || endDate
      ? `${startDate ? formatShortDateId(startDate) : 'Awal'} s.d. ${
          endDate ? formatShortDateId(endDate) : 'Sekarang'
        }`
      : 'Seluruh Tanggal Pertemuan';

  const reportSheets = useMemo(() => {
    const classScopedTypes = new Set<ReportSectionId>([
      'journal',
      'attendance',
      'attendance_monthly',
      'attendance_monitoring',
      'grades',
      'grades_remedial',
      'grades_leger',
      'students_roster',
    ]);

    const sheets: {
      key: string;
      reportType: ReportSectionId;
      className: string | null;
      label: string;
      officialTitle: string;
      subtitle: string;
    }[] = [];

    orderedSelectedClasses.forEach((cls) => {
      const classSubjects = getSubjectsForClassSheet(cls);
      const classSubjectLabel =
        classSubjects.length === 0
          ? subjectLabel
          : !activeTeacher && classSubjects.length === subjectOptions.length && subjectOptions.length > 0
          ? 'Semua Mata Pelajaran'
          : classSubjects.join(', ');
      const teacherSuffix = activeTeacher ? ` · Guru: ${activeTeacher.name}` : '';
      orderedSelectedReportTypes.forEach((reportType) => {
        if (!classScopedTypes.has(reportType)) return;
        const opt = REPORT_TYPE_OPTIONS.find((o) => o.id === reportType);
        if (!opt) return;
        const subtitle =
          reportType === 'students_roster' || reportType === 'grades_leger'
            ? `Kelas ${cls} · Tahun Ajaran ${reportAcademicYear} · Semester ${reportSemester}${teacherSuffix}`
            : `Kelas ${cls} · Tahun Ajaran ${reportAcademicYear} · Semester ${reportSemester} · Mata Pelajaran: ${classSubjectLabel} · Rentang Tanggal: ${dateRangeLabel}${teacherSuffix}`;
        sheets.push({
          key: `${cls}__${reportType}`,
          reportType,
          className: cls,
          label: opt.label,
          officialTitle: opt.officialTitle,
          subtitle,
        });
      });
    });

    if (orderedSelectedReportTypes.includes('teacher_activity')) {
      const opt = REPORT_TYPE_OPTIONS.find((o) => o.id === 'teacher_activity')!;
      sheets.push({
        key: 'global__teacher_activity',
        reportType: 'teacher_activity',
        className: null,
        label: opt.label,
        officialTitle: opt.officialTitle,
        subtitle: `Tahun Ajaran ${reportAcademicYear} · Semester ${reportSemester} · Rentang Tanggal: ${dateRangeLabel}`,
      });
    }

    if (orderedSelectedReportTypes.includes('charts')) {
      const opt = REPORT_TYPE_OPTIONS.find((o) => o.id === 'charts')!;
      sheets.push({
        key: 'global__charts',
        reportType: 'charts',
        className: null,
        label: opt.label,
        officialTitle: opt.officialTitle,
        subtitle: `Kelas: ${orderedSelectedClasses.length > 0 ? orderedSelectedClasses.join(', ') : 'Belum Dipilih'} · Tahun Ajaran ${reportAcademicYear} · Semester ${reportSemester} · Mata Pelajaran: ${subjectLabel}`,
      });
    }

    return sheets;
  }, [
    orderedSelectedClasses,
    orderedSelectedReportTypes,
    reportAcademicYear,
    reportSemester,
    subjectLabel,
    dateRangeLabel,
    activeTeacher,
    activeGradeSubjects,
    subjectOptions.length,
  ]);

  // Build Workbook for XLSX / XLS (grouped per class first, separate sheet per report type)
  const buildWorkbook = () => {
    const wb = XLSX.utils.book_new();
    const usedNames = new Set<string>();
    const appendUniqueSheet = (ws: XLSX.WorkSheet, rawName: string) => {
      const sanitized = rawName.replace(/[\\/?*[\]]/g, '').trim() || 'Sheet';
      let candidate = sanitized.slice(0, 31);
      let suffix = 2;
      while (usedNames.has(candidate.toLowerCase())) {
        const tag = ` (${suffix++})`;
        candidate = `${sanitized.slice(0, Math.max(1, 31 - tag.length))}${tag}`;
      }
      usedNames.add(candidate.toLowerCase());
      XLSX.utils.book_append_sheet(wb, ws, candidate);
    };

    reportSheets.forEach((sheet) => {
      const cls = sheet.className;

      if (sheet.reportType === 'journal' && cls) {
        const classJournals = filteredJournals.filter((j) => j.className === cls);
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
          ...classJournals.map((j, idx) => {
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet(jRows), `${cls} Jurnal`);
      }

      if (sheet.reportType === 'attendance' && cls) {
        const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
        const header = [
          'No',
          'NIS',
          'NISN',
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
            st.nisn || '-',
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), `${cls} Hadir`);
      }

      if (sheet.reportType === 'attendance_monthly' && cls) {
        const { months, rows } = getMonthlyAttendanceForClass(cls);
        const header = [
          'No',
          'NIS',
          'NISN',
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
          r.nisn,
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), `${cls} Presensi`);
      }

      if (sheet.reportType === 'attendance_monitoring' && cls) {
        const clsMonitoring = getMonitoringAbsenceForClass(cls);
        const sumHeader = ['No', 'NIS', 'Nama Siswa', 'Kelas', 'S', 'I', 'A', 'D', 'Total', 'Status Pembinaan'];
        const sumBody = clsMonitoring.summary.map((item, idx) => [
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([sumHeader, ...sumBody]), `${cls} Rekap SIAD`);

        const logHeader = ['No', 'Tanggal', 'Jam', 'Kelas', 'Mapel', 'NIS', 'Nama Siswa', 'Status', 'Keterangan'];
        const logBody = clsMonitoring.chronology.map((c, idx) => [
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([logHeader, ...logBody]), `${cls} Log SIAD`);
      }

      if (sheet.reportType === 'grades' && cls) {
        const classGradeSubjects = getSubjectsForClassSheet(cls);
        const visibleSubjects = classGradeSubjects.filter((subj) => {
          const { rows } = getGradeBlockForClassAndSubject(cls, subj);
          return rows.some((r) => r.avg !== null) || classGradeSubjects.length === 1;
        });
        if (visibleSubjects.length === 0) {
          appendUniqueSheet(
            XLSX.utils.aoa_to_sheet([[`Belum ada nilai tercatat untuk kelas ${cls} pada mata pelajaran terpilih.`]]),
            `${cls} Nilai`
          );
        } else {
          visibleSubjects.forEach((subj) => {
            const { rows } = getGradeBlockForClassAndSubject(cls, subj);
            const header = [
              'No',
              'NIS',
              'NISN',
              'Nama Siswa',
              ...settings.masterExamTypes,
              'Rata-rata',
              'Status',
            ];
            const body = rows.map((r) => [
              r.no,
              r.nis,
              r.nisn,
              r.name,
              ...settings.masterExamTypes.map((ex) => (r.scores[ex] !== null ? r.scores[ex] : '')),
              r.avg !== null ? Number(r.avg.toFixed(1)) : '',
              r.status,
            ]);
            appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), `${cls} Nilai ${subj}`);
          });
        }
      }

      if (sheet.reportType === 'grades_remedial' && cls) {
        const classGradeSubjects = getSubjectsForClassSheet(cls);
        const visibleSubjects = classGradeSubjects.filter((subj) => {
          const rem = getRemedialDataForClassAndSubject(cls, subj);
          return rem.assessedCount > 0 || classGradeSubjects.length === 1;
        });
        if (visibleSubjects.length === 0) {
          appendUniqueSheet(
            XLSX.utils.aoa_to_sheet([[`Belum ada data ketuntasan/remedial untuk kelas ${cls} pada mata pelajaran terpilih.`]]),
            `${cls} Remedial`
          );
        } else {
          visibleSubjects.forEach((subj) => {
            const rem = getRemedialDataForClassAndSubject(cls, subj);
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
            appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), `${cls} Remed ${subj}`);
          });
        }
      }

      if (sheet.reportType === 'grades_leger' && cls) {
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), `${cls} Leger`);
      }

      if (sheet.reportType === 'students_roster' && cls) {
        const clsStudents = studentRosterRows.filter((s) => s.className === cls);
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
        const body = clsStudents.map((st, idx) => [
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), `${cls} Siswa`);
      }

      if (sheet.reportType === 'teacher_activity') {
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
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), 'Kinerja Guru');
      }

      if (sheet.reportType === 'charts') {
        const header = ['Jenis Grafik', 'Kategori / Label', 'Nilai Utama', 'Keterangan'];
        const body =
          chartCards.length === 0
            ? [['Grafis', 'Belum ada jenis diagram dipilih', '-', '-']]
            : chartCards.flatMap((c) => c.summaryRows);
        appendUniqueSheet(XLSX.utils.aoa_to_sheet([header, ...body]), 'Ringkasan Grafis');
      }
    });

    if (wb.SheetNames.length === 0) {
      const ws = XLSX.utils.aoa_to_sheet([['Tidak ada data untuk filter terpilih']]);
      XLSX.utils.book_append_sheet(wb, ws, 'Laporan');
    }

    return wb;
  };

  const validateBeforeExport = () => {
    if (orderedSelectedReportTypes.length === 0) {
      onToast('error', 'Pilih minimal satu jenis laporan terlebih dahulu.');
      return false;
    }
    if (reportSheets.length === 0) {
      onToast('error', 'Pilih minimal satu kelas terlebih dahulu.');
      return false;
    }
    return true;
  };

  const handleDownloadXlsx = () => {
    if (!validateBeforeExport()) return;
    const wb = buildWorkbook();
    XLSX.writeFile(wb, `Laporan_E-Guru_${reportSlug}_${getTodayISO()}.xlsx`);
    onToast('success', 'Laporan Excel (.xlsx) berhasil diunduh.');
  };

  const handleDownloadXls = () => {
    if (!validateBeforeExport()) return;
    const wb = buildWorkbook();
    XLSX.writeFile(wb, `Laporan_E-Guru_${reportSlug}_${getTodayISO()}.xls`, {
      bookType: 'biff8',
    });
    onToast('success', 'Laporan Excel (.xls) berhasil diunduh.');
  };

  const handleDownloadWord = async () => {
    if (!validateBeforeExport()) return;
    if (!previewRef.current) return;
    setIsExportingWord(true);
    try {
      await exportPreviewDomToWordDoc({
        previewElement: previewRef.current,
        title: reportTitle,
        filename: `Laporan_E-Guru_${reportSlug}_${getTodayISO()}.doc`,
        orientation,
        paperSize: effectivePaperSize,
        cssPageSize: PAPER_SIZE_CONFIG[effectivePaperSize].cssPageSize,
        fontFamily: effectiveFontFamily,
      });
      onToast('success', 'Dokumen Word (.doc) persis pratinjau berhasil diunduh.');
    } catch {
      onToast('error', 'Gagal mengekspor ke Word (.doc).');
    } finally {
      setIsExportingWord(false);
    }
  };

  // Vector PDF Generator matching Live Preview 1:1 (per-class order, each report type on a separate page)
  const handleDownloadPdf = async () => {
    if (!validateBeforeExport()) return;
    setIsGeneratingPdf(true);
    try {
      const paperCfg = PAPER_SIZE_CONFIG[effectivePaperSize];
      const doc = new jsPDF({
        orientation,
        unit: 'mm',
        format: paperCfg.jsPdfFormat,
      });

      const pdfFont =
        effectiveFontFamily === 'Times New Roman' || effectiveFontFamily === 'Georgia'
          ? 'times'
          : 'helvetica';
      doc.setFont(pdfFont, 'normal');

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const marginX = 12;
      const usableWidth = pageWidth - marginX * 2;
      const centerX = pageWidth / 2;

      const [leftLogoInfo, rightLogoInfo, kepsekSigInfo, guruSigInfo, stampInfo] =
        await Promise.all([
          showKop ? imageUrlToPngWithMeta(settings.kopSurat.logoLeft) : Promise.resolve(null),
          showKop ? imageUrlToPngWithMeta(settings.kopSurat.logoRight) : Promise.resolve(null),
          showTtdKepsek && showSignatureImage
            ? imageUrlToPngWithMeta(settings.principalSignatureImage)
            : Promise.resolve(null),
          showTtdGuru && showSignatureImage
            ? imageUrlToPngWithMeta(activeSigner.signatureImage)
            : Promise.resolve(null),
          sigPos.showStamp
            ? imageUrlToPngWithMeta(resolvePrincipalStampImage(db))
            : Promise.resolve(null),
        ]);

      const commonTableStyles = {
        font: pdfFont,
        fontSize: 8,
        textColor: 0 as const,
        cellPadding: 1.4,
        lineColor: [0, 0, 0] as [number, number, number],
        lineWidth: 0.2,
      };

      const commonHeadStyles = {
        fillColor: [241, 245, 249] as [number, number, number],
        textColor: 0 as const,
        fontStyle: 'bold' as const,
        lineColor: [0, 0, 0] as [number, number, number],
        lineWidth: 0.2,
      };

      const PX_TO_MM = 0.264583;
      const kepsekSigPx = Math.max(
        16,
        Math.min(260, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
      );
      const guruSigPx = Math.max(
        16,
        Math.min(260, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
      );
      const stampPx = Math.max(20, Math.min(260, sigPos.stampSize ?? 70));
      const stampMm = Number((stampPx * PX_TO_MM).toFixed(2));
      const kepsekSigOffsetX = sigPos.principalSigOffsetX ?? 0;
      const kepsekSigOffsetY = sigPos.principalSigOffsetY ?? 0;
      const guruSigOffsetX = sigPos.teacherSigOffsetX ?? 0;
      const guruSigOffsetY = sigPos.teacherSigOffsetY ?? 0;
      const stampOffsetX = sigPos.stampOffsetX ?? 0;
      const stampOffsetY = sigPos.stampOffsetY ?? 0;

      const drawPdfSigBlock = (
        xMm: number,
        yMm: number,
        topLabel: string,
        roleTitle: string,
        personName: string,
        personNip: string,
        sigInfo: PngImageInfo | null,
        withStamp: boolean,
        sigPx: number,
        sigOffsetX: number,
        sigOffsetY: number
      ) => {
        const sigBoxH = Number((sigPx * PX_TO_MM).toFixed(2));
        const sigMaxW = Number((sigBoxH * 2.85).toFixed(2));
        doc.setFont(pdfFont, 'normal');
        doc.setFontSize(9.5);
        doc.setTextColor(0, 0, 0);
        doc.text(topLabel, xMm, yMm);
        doc.text(roleTitle, xMm, yMm + 4.5);
        if (withStamp && stampInfo) {
          const fittedStamp = fitContainDimensions(stampInfo.width, stampInfo.height, stampMm, stampMm);
          const stampCenterX = xMm + 4.8 + stampOffsetX * PX_TO_MM;
          const stampCenterY = yMm + 16 + stampOffsetY * PX_TO_MM;
          doc.addImage(
            stampInfo.dataUrl,
            'PNG',
            stampCenterX - fittedStamp.width / 2,
            stampCenterY - fittedStamp.height / 2,
            fittedStamp.width,
            fittedStamp.height
          );
        }
        if (sigInfo) {
          const fittedSig = fitContainDimensions(sigInfo.width, sigInfo.height, sigMaxW, sigBoxH);
          doc.addImage(
            sigInfo.dataUrl,
            'PNG',
            xMm + sigOffsetX * PX_TO_MM,
            yMm + 15 - fittedSig.height / 2 + sigOffsetY * PX_TO_MM,
            fittedSig.width,
            fittedSig.height
          );
        }
        doc.setFont(pdfFont, 'bold');
        doc.text(personName, xMm, yMm + 25.5);
        const nameW = doc.getTextWidth(personName);
        doc.setLineWidth(0.25);
        doc.line(xMm, yMm + 26.2, xMm + nameW, yMm + 26.2);
        doc.setFont(pdfFont, 'normal');
        doc.text(`NIP. ${personNip || '-'}`, xMm, yMm + 30);
      };

      reportSheets.forEach((sheet, sheetIdx) => {
        if (sheetIdx > 0) {
          doc.addPage();
        }
        let cursorY = 12;

        const ensureSpace = (minMm = 36) => {
          if (cursorY > pageHeight - minMm) {
            doc.addPage();
            cursorY = 14;
          }
        };

        const drawSectionTitle = (label: string) => {
          ensureSpace(36);
          doc.setFont(pdfFont, 'bold');
          doc.setFontSize(10);
          doc.setTextColor(0, 0, 0);
          doc.text(label, marginX, cursorY);
          cursorY += 2.6;
        };

        if (showKop) {
          cursorY = drawPdfKopSurat({
            doc,
            kop: {
              ...settings.kopSurat,
              logoLeftSize: kopLeftSize,
              logoRightSize: kopRightSize,
              lineSpacing: kopLineSpacing,
            },
            leftLogoPng: leftLogoInfo,
            rightLogoPng: rightLogoInfo,
            logoLeftPx: kopLeftSize,
            logoRightPx: kopRightSize,
            marginX,
            startY: cursorY,
            fallbackFont: pdfFont,
          });
        }

        // Sheet Title (matching Preview 12pt bold uppercase + underline)
        doc.setFont(pdfFont, 'bold');
        doc.setFontSize(12);
        const wrappedTitle = doc.splitTextToSize(sheet.officialTitle.toUpperCase(), usableWidth) as string[];
        wrappedTitle.forEach((line) => {
          doc.text(line, centerX, cursorY, { align: 'center' });
          const lineW = doc.getTextWidth(line);
          doc.setLineWidth(0.35);
          doc.line(centerX - lineW / 2, cursorY + 1, centerX + lineW / 2, cursorY + 1);
          cursorY += 5.5;
        });

        doc.setFont(pdfFont, 'normal');
        doc.setFontSize(9);
        doc.setTextColor(51, 65, 85);
        const wrappedSub = doc.splitTextToSize(sheet.subtitle, usableWidth) as string[];
        wrappedSub.forEach((line) => {
          doc.text(line, centerX, cursorY, { align: 'center' });
          cursorY += 4.2;
        });
        doc.setTextColor(0, 0, 0);
        cursorY += 2;

        const cls = sheet.className;

        if (sheet.reportType === 'journal' && cls) {
          const classJournals = filteredJournals.filter((j) => j.className === cls);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: commonTableStyles,
            headStyles: commonHeadStyles,
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
              classJournals.length === 0
                ? [['-', '-', '-', cls, '-', `Belum ada data jurnal untuk kelas ${cls}`, '-', '-']]
                : classJournals.map((j, idx) => {
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

        if (sheet.reportType === 'attendance' && cls) {
          const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
          drawSectionTitle(`Leger Kehadiran Mendatar — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: { ...commonTableStyles, fontSize: 7.5 },
            headStyles: commonHeadStyles,
            head: [
              [
                'No',
                'NIS',
                'NISN',
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
                ? [['-', '-', '-', 'Belum ada siswa di kelas ini']]
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
                      st.nisn || '-',
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
        }

        if (sheet.reportType === 'attendance_monthly' && cls) {
          const { months, rows } = getMonthlyAttendanceForClass(cls);
          drawSectionTitle(`Rekap Presensi Bulanan — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: commonTableStyles,
            headStyles: commonHeadStyles,
            head: [
              [
                'No',
                'NIS',
                'NISN',
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
                ? [['-', '-', '-', 'Belum ada siswa di kelas ini']]
                : rows.map((r) => [
                    r.no,
                    r.nis,
                    r.nisn,
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
        }

        if (sheet.reportType === 'attendance_monitoring' && cls) {
          const clsMonitoring = getMonitoringAbsenceForClass(cls);
          drawSectionTitle(`Monitoring Ketidakhadiran (S/I/A/D) — Rekap Siswa Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: commonTableStyles,
            headStyles: commonHeadStyles,
            head: [['No', 'NIS', 'Nama Siswa', 'Kelas', 'S', 'I', 'A', 'D', 'Total', 'Pembinaan']],
            body:
              clsMonitoring.summary.length === 0
                ? [['-', '-', `Seluruh siswa kelas ${cls} hadir 100% pada periode ini`, cls, 0, 0, 0, 0, 0, '-']]
                : clsMonitoring.summary.map((it, idx) => [
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

          drawSectionTitle(`Kronologi Ketidakhadiran (S/I/A/D) — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: commonTableStyles,
            headStyles: commonHeadStyles,
            head: [['No', 'Tanggal', 'Jam', 'Kelas', 'Mapel', 'Nama Siswa', 'Status', 'Keterangan']],
            body:
              clsMonitoring.chronology.length === 0
                ? [['-', '-', '-', cls, '-', `Belum ada catatan ketidakhadiran di kelas ${cls}`, '-', '-']]
                : clsMonitoring.chronology.map((c, idx) => [
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

        if (sheet.reportType === 'grades' && cls) {
          const classGradeSubjects = getSubjectsForClassSheet(cls);
          const visibleSubjects = classGradeSubjects.filter((subj) => {
            const { rows } = getGradeBlockForClassAndSubject(cls, subj);
            return rows.some((r) => r.avg !== null) || classGradeSubjects.length === 1;
          });
          if (visibleSubjects.length === 0) {
            doc.setFont(pdfFont, 'normal');
            doc.setFontSize(9);
            doc.text(
              `Belum ada nilai tercatat untuk kelas ${cls} pada mata pelajaran terpilih.`,
              marginX,
              cursorY + 4
            );
            cursorY += 12;
          } else {
            visibleSubjects.forEach((subj) => {
              const { rows } = getGradeBlockForClassAndSubject(cls, subj);
              drawSectionTitle(`Rekap Nilai Mapel — Kelas ${cls} (${subj})`);
              autoTable(doc, {
                startY: cursorY,
                margin: { left: marginX, right: marginX },
                styles: commonTableStyles,
                headStyles: commonHeadStyles,
                head: [['No', 'NIS', 'NISN', 'Nama Siswa', ...settings.masterExamTypes, 'Rata-rata', 'Status']],
                body:
                  rows.length === 0
                    ? [['-', '-', '-', `Belum ada siswa di kelas ${cls}`]]
                    : rows.map((r) => [
                        r.no,
                        r.nis,
                        r.nisn,
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
          }
        }

        if (sheet.reportType === 'grades_remedial' && cls) {
          const classGradeSubjects = getSubjectsForClassSheet(cls);
          const visibleSubjects = classGradeSubjects.filter((subj) => {
            const rem = getRemedialDataForClassAndSubject(cls, subj);
            return rem.assessedCount > 0 || classGradeSubjects.length === 1;
          });
          if (visibleSubjects.length === 0) {
            doc.setFont(pdfFont, 'normal');
            doc.setFontSize(9);
            doc.text(
              `Belum ada data ketuntasan & remedial untuk kelas ${cls} pada mata pelajaran terpilih.`,
              marginX,
              cursorY + 4
            );
            cursorY += 12;
          } else {
            visibleSubjects.forEach((subj) => {
              const rem = getRemedialDataForClassAndSubject(cls, subj);
              const dayaLabel = rem.dayaSerap !== null ? `${rem.dayaSerap.toFixed(0)}%` : '-';
              drawSectionTitle(
                `Ketuntasan & Remedial (< KKM 75) — ${cls} (${subj}) · Daya Serap: ${dayaLabel} (${rem.tuntasCount}/${rem.assessedCount})`
              );
              autoTable(doc, {
                startY: cursorY,
                margin: { left: marginX, right: marginX },
                styles: commonTableStyles,
                headStyles: commonHeadStyles,
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
                  rem.assessedCount === 0
                    ? [['-', '-', `Belum ada nilai tercatat untuk kelas ${cls} (${subj})`, '-', '-', '-', '-', '-', '-']]
                    : rem.remedialList.length === 0
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
          }
        }

        if (sheet.reportType === 'grades_leger' && cls) {
          const { subjects, rows } = getCombinedLegerForClass(cls);
          drawSectionTitle(`Leger Gabungan Kelas & Peringkat — Kelas ${cls}`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: { ...commonTableStyles, fontSize: 7.5 },
            headStyles: commonHeadStyles,
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
        }

        if (sheet.reportType === 'students_roster' && cls) {
          const clsStudents = studentRosterRows.filter((s) => s.className === cls);
          drawSectionTitle(`Daftar Peserta Didik & Wali Kelas — Kelas ${cls} (${clsStudents.length} Siswa)`);
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: { ...commonTableStyles, fontSize: 7.5 },
            headStyles: commonHeadStyles,
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
              clsStudents.length === 0
                ? [['-', '-', `Belum ada data peserta didik di kelas ${cls}`]]
                : clsStudents.map((st, idx) => [
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

        if (sheet.reportType === 'teacher_activity') {
          drawSectionTitle('Aktivitas & Penilaian Kinerja Guru — Ringkasan');
          autoTable(doc, {
            startY: cursorY,
            margin: { left: marginX, right: marginX },
            styles: commonTableStyles,
            headStyles: commonHeadStyles,
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
            headStyles: commonHeadStyles,
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

        if (sheet.reportType === 'charts') {
          drawSectionTitle('Laporan Grafis & Diagram Analisis KBM');
          if (chartCards.length === 0) {
            doc.setFont(pdfFont, 'normal');
            doc.setFontSize(9);
            doc.text('Belum ada jenis diagram grafik yang dipilih.', marginX, cursorY + 4);
            cursorY += 10;
          } else {
            const cols = effectivePaperSize === 'A5' && orientation === 'portrait' ? 1 : 2;
            const gapMm = 4;
            const cardW = (usableWidth - gapMm * (cols - 1)) / cols;
            const imgW = cardW - 5;

            for (let i = 0; i < chartCards.length; i += cols) {
              const rowCards = chartCards.slice(i, i + cols);
              const rowMaxCardH = Math.max(
                ...rowCards.map((c) => 12.5 + imgW * (c.heightPx / c.widthPx))
              );
              ensureSpace(rowMaxCardH + 4);

              rowCards.forEach((card, cIdx) => {
                const cardX = marginX + cIdx * (cardW + gapMm);
                const cardY = cursorY;
                const cardImgH = imgW * (card.heightPx / card.widthPx);
                doc.setDrawColor(0, 0, 0);
                doc.setLineWidth(0.25);
                doc.setFillColor(255, 255, 255);
                doc.rect(cardX, cardY, cardW, rowMaxCardH, 'FD');

                doc.setFont(pdfFont, 'bold');
                doc.setFontSize(8.5);
                doc.setTextColor(0, 0, 0);
                doc.text(card.title, cardX + 2.5, cardY + 4.5, { maxWidth: cardW - 5 });

                doc.setFont(pdfFont, 'normal');
                doc.setFontSize(7);
                doc.setTextColor(71, 85, 105);
                doc.text(`${card.chartKind} · ${card.subtitle}`, cardX + 2.5, cardY + 8.5, {
                  maxWidth: cardW - 5,
                });
                doc.setTextColor(0, 0, 0);

                if (card.pngDataUrl) {
                  doc.addImage(card.pngDataUrl, 'PNG', cardX + 2.5, cardY + 10.2, imgW, cardImgH);
                }
              });

              cursorY += rowMaxCardH + 4.5;
            }
            cursorY += 3;
          }
        }

        // Signature Block matching Preview & sigPos 1:1 on each sheet
        if (showTtdKepsek || showTtdGuru) {
          const neededH =
            sigPos.preset === 'kiri' || sigPos.preset === 'kanan'
              ? 76
              : sigPos.preset === 'bebas'
              ? Math.max(48, Math.max(sigPos.principalY, sigPos.teacherY) * PX_TO_MM + 38)
              : 48;
          if (cursorY > pageHeight - neededH) {
            doc.addPage();
            cursorY = 14;
          }
          const sigBaseY = cursorY + 4;
          const maxSigX = pageWidth - marginX - 50;

          const stampOnKepsek = sigPos.showStamp;
          const stampOnGuru = false;

          let kepsekX = marginX + 4.5;
          let kepsekY = sigBaseY;
          let guruX = marginX + usableWidth * 0.59;
          let guruY = sigBaseY;

          if (sigPos.preset === 'kanan-kiri') {
            guruX = marginX + 4.5;
            kepsekX = marginX + usableWidth * 0.59;
          } else if (sigPos.preset === 'kiri') {
            kepsekX = marginX + 4.5;
            guruX = marginX + 4.5;
            guruY = sigBaseY + 36;
          } else if (sigPos.preset === 'tengah') {
            kepsekX = pageWidth / 2 - 58;
            guruX = pageWidth / 2 + 12;
          } else if (sigPos.preset === 'kanan') {
            kepsekX = pageWidth - marginX - 63.5;
            guruX = pageWidth - marginX - 63.5;
            guruY = sigBaseY + 36;
          } else if (sigPos.preset === 'bebas') {
            kepsekX = Math.min(maxSigX, marginX + (usableWidth * sigPos.principalX) / 100);
            kepsekY = sigBaseY + sigPos.principalY * PX_TO_MM;
            guruX = Math.min(maxSigX, marginX + (usableWidth * sigPos.teacherX) / 100);
            guruY = sigBaseY + sigPos.teacherY * PX_TO_MM;
          }

          if (showTtdKepsek) {
            drawPdfSigBlock(
              kepsekX,
              kepsekY,
              'Mengetahui,',
              settings.principalTitle,
              settings.principalName,
              settings.principalNip,
              kepsekSigInfo,
              stampOnKepsek,
              kepsekSigPx,
              kepsekSigOffsetX,
              kepsekSigOffsetY
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
              guruSigInfo,
              stampOnGuru,
              guruSigPx,
              guruSigOffsetX,
              guruSigOffsetY
            );
          }
        }
      });

      doc.save(`Laporan_E-Guru_${reportSlug}_${getTodayISO()}.pdf`);
      onToast('success', 'Laporan PDF vektor 1:1 berhasil diunduh.');
    } catch {
      onToast('error', 'Terjadi kendala saat membuat PDF. Silakan coba lagi.');
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  const handlePrintDirect = () => {
    if (!validateBeforeExport()) return;
    onToast('info', 'Membuka dialog cetak laporan resmi...');
    window.print();
  };

  const activePaperCfg = PAPER_SIZE_CONFIG[effectivePaperSize];
  const activePaperWidthMm =
    orientation === 'landscape' ? activePaperCfg.heightMm : activePaperCfg.widthMm;
  const chartCols = effectivePaperSize === 'A5' && orientation === 'portrait' ? 1 : 2;

  const principalStampImage = resolvePrincipalStampImage(db);
  const stampOnKepsek = sigPos.showStamp;
  const stampOnGuru = false;

  const kepsekSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const guruSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const stampSizePx = Math.max(20, Math.min(260, sigPos.stampSize ?? 70));
  const kepsekSigOffsetX = sigPos.principalSigOffsetX ?? 0;
  const kepsekSigOffsetY = sigPos.principalSigOffsetY ?? 0;
  const guruSigOffsetX = sigPos.teacherSigOffsetX ?? 0;
  const guruSigOffsetY = sigPos.teacherSigOffsetY ?? 0;
  const stampOffsetX = sigPos.stampOffsetX ?? 0;
  const stampOffsetY = sigPos.stampOffsetY ?? 0;

  const renderSigCard = (
    sigTarget: 'kepsek' | 'guru',
    topLabel: string,
    roleTitle: string,
    personName: string,
    personNip: string,
    sigImage: string,
    withStamp: boolean,
    sigPx: number,
    sigOffsetX: number,
    sigOffsetY: number
  ) => {
    const mediaKey: 'kepsek-sig' | 'guru-sig' =
      sigTarget === 'kepsek' ? 'kepsek-sig' : 'guru-sig';
    const isSigSelected = selectedMedia === mediaKey;
    const isStampSelected = selectedMedia === 'stamp';
    const stampLeftCss =
      sigPos.preset === 'tengah'
        ? `calc(50% - ${23 + Math.round(stampSizePx / 2) - stampOffsetX}px)`
        : `${18 - Math.round(stampSizePx / 2) + stampOffsetX}px`;
    const stampTopPx = Math.round((72 - stampSizePx) / 2) + stampOffsetY;

    return (
      <div
        style={{
          fontSize: '9.5pt',
          lineHeight: 1.35,
          textAlign: sigPos.preset === 'tengah' ? 'center' : 'left',
        }}
      >
        <p style={{ margin: 0 }}>{topLabel}</p>
        <p style={{ margin: 0 }}>{roleTitle}</p>
        <div
          data-sig-media="true"
          data-stamp-src={withStamp && principalStampImage ? principalStampImage : ''}
          data-stamp-left={String(
            (sigPos.preset === 'tengah' ? 87 : 18) - Math.round(stampSizePx / 2) + stampOffsetX
          )}
          data-stamp-top={String(stampTopPx)}
          data-stamp-size={String(stampSizePx)}
          data-sig-src={sigImage || ''}
          data-sig-left={String(sigOffsetX)}
          data-sig-top={String(sigOffsetY)}
          data-sig-height={String(sigPx)}
          data-sig-max-width={String(Math.round(sigPx * 2.85))}
          data-align-center={sigPos.preset === 'tengah' ? 'true' : 'false'}
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
          {withStamp && principalStampImage && (
            <>
              <img
                src={principalStampImage}
                alt="Stempel Sekolah"
                draggable={false}
                style={{
                  position: 'absolute',
                  left: stampLeftCss,
                  top: `${stampTopPx}px`,
                  width: `${stampSizePx}px`,
                  height: `${stampSizePx}px`,
                  objectFit: 'contain',
                  opacity: 0.85,
                  mixBlendMode: 'multiply',
                  pointerEvents: 'none',
                  zIndex: 1,
                }}
              />
              <div
                data-export-ignore="true"
                onPointerDown={(e) => startMediaInteraction('move-stamp', 'stamp', e)}
                title="Seret gambar stempel untuk memindahkan posisi dengan bebas"
                style={{
                  position: 'absolute',
                  left: stampLeftCss,
                  top: `${stampTopPx}px`,
                  width: `${stampSizePx}px`,
                  height: `${stampSizePx}px`,
                  zIndex: isStampSelected ? 20 : 1,
                  touchAction: 'none',
                }}
                className={`no-print group/stamp cursor-move rounded-full transition-all ${
                  isStampSelected
                    ? 'ring-2 ring-emerald-600 bg-emerald-500/5'
                    : 'hover:ring-2 hover:ring-dashed hover:ring-emerald-500'
                }`}
              >
                <span
                  className={`absolute -top-5 left-1/2 -translate-x-1/2 px-1.5 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-sans font-semibold whitespace-nowrap pointer-events-none transition-opacity ${
                    isStampSelected ? 'opacity-100' : 'opacity-0 group-hover/stamp:opacity-100'
                  }`}
                >
                  Stempel ({stampSizePx}px)
                </span>
                <button
                  type="button"
                  onPointerDown={(e) => startMediaInteraction('resize-stamp', 'stamp', e)}
                  title="Seret sudut untuk memperbesar / memperkecil gambar stempel"
                  style={{ touchAction: 'none' }}
                  className={`absolute -bottom-1 -right-1 w-4 h-4 rounded-sm bg-emerald-600 border-2 border-white shadow cursor-nwse-resize transition-opacity ${
                    isStampSelected ? 'opacity-100' : 'opacity-0 group-hover/stamp:opacity-100'
                  }`}
                  aria-label="Ubah ukuran gambar stempel"
                />
              </div>
            </>
          )}
          {sigImage && (
            <div
              style={{
                position: 'relative',
                left: `${sigOffsetX}px`,
                top: `${sigOffsetY}px`,
                zIndex: isSigSelected ? 15 : 2,
                display: 'inline-block',
                lineHeight: 0,
              }}
            >
              <img
                src={sigImage}
                alt={`TTD ${personName}`}
                draggable={false}
                style={{
                  display: 'block',
                  height: `${sigPx}px`,
                  maxWidth: `${Math.round(sigPx * 2.85)}px`,
                  objectFit: 'contain',
                  pointerEvents: 'none',
                }}
              />
              <div
                data-export-ignore="true"
                onPointerDown={(e) =>
                  startMediaInteraction(
                    sigTarget === 'kepsek' ? 'move-kepsek-sig' : 'move-guru-sig',
                    mediaKey,
                    e
                  )
                }
                title={`Seret gambar TTD ${
                  sigTarget === 'kepsek' ? 'Kepala Sekolah' : 'Guru'
                } untuk memindahkan posisi dengan bebas`}
                style={{ touchAction: 'none' }}
                className={`no-print group/sig absolute inset-0 cursor-move rounded transition-all ${
                  isSigSelected
                    ? 'ring-2 ring-indigo-600 bg-indigo-500/5'
                    : 'hover:ring-2 hover:ring-dashed hover:ring-indigo-500'
                }`}
              >
                <span
                  className={`absolute -top-5 left-0 px-1.5 py-0.5 rounded bg-indigo-600 text-white text-[10px] font-sans font-semibold whitespace-nowrap pointer-events-none transition-opacity ${
                    isSigSelected ? 'opacity-100' : 'opacity-0 group-hover/sig:opacity-100'
                  }`}
                >
                  {sigTarget === 'kepsek' ? 'TTD Kepsek' : 'TTD Guru'} ({sigPx}px)
                </span>
                <button
                  type="button"
                  onPointerDown={(e) =>
                    startMediaInteraction(
                      sigTarget === 'kepsek' ? 'resize-kepsek-sig' : 'resize-guru-sig',
                      mediaKey,
                      e
                    )
                  }
                  title="Seret sudut kanan bawah untuk memperbesar / memperkecil gambar TTD"
                  style={{ touchAction: 'none' }}
                  className={`absolute -bottom-1.5 -right-1.5 w-4 h-4 rounded-sm bg-indigo-600 border-2 border-white shadow cursor-nwse-resize transition-opacity ${
                    isSigSelected ? 'opacity-100' : 'opacity-0 group-hover/sig:opacity-100'
                  }`}
                  aria-label="Ubah ukuran gambar tanda tangan"
                />
              </div>
            </div>
          )}
        </div>
        <p style={{ margin: 0, fontWeight: 700, textDecoration: 'underline', textUnderlineOffset: '2px' }}>
          {personName}
        </p>
        <p style={{ margin: 0 }}>NIP. {personNip || '-'}</p>
      </div>
    );
  };

  return (
    <div className="space-y-6 print:space-y-0">
      {/* Control Panel (Hidden on Print) */}
      <section className={`no-print bg-white border border-slate-200 rounded-xl p-5 ${filterPanelExpanded ? 'space-y-5' : ''}`}>
        <div className={`flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 ${filterPanelExpanded ? 'border-b border-slate-200 pb-4' : ''}`}>
          <button
            type="button"
            onClick={() => setFilterPanelExpanded((prev) => !prev)}
            aria-expanded={filterPanelExpanded}
            className="flex items-center justify-between lg:justify-start gap-3 text-left group min-w-0"
          >
            <div className="min-w-0">
              <div className="flex flex-wrap items-center gap-2">
                <h2 className="text-base font-bold text-slate-900 group-hover:text-indigo-700 transition-colors">
                  Filter &amp; Pengaturan Cetak Laporan Resmi
                </h2>
                <span className="text-[11px] font-medium text-slate-500">
                  · TA {reportAcademicYear} ({reportSemester}) · {selectedReportTypes.length} laporan · {selectedClasses.length} kelas · {selectedSubjects.length} mapel · {effectivePaperSize}
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                Hasil cetak, PDF, dan Word disesuaikan 1:1 dengan pratinjau dokumen, TTD pemilik akun, stempel sekolah &amp; posisi pengesahan
              </p>
            </div>
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 group-hover:bg-slate-100 text-xs font-semibold text-slate-700 whitespace-nowrap shrink-0 transition-colors">
              <span>{filterPanelExpanded ? 'Ciutkan' : 'Tampilkan'}</span>
              <ChevronDown
                className={`w-4 h-4 text-slate-500 transition-transform duration-150 ${
                  filterPanelExpanded ? 'rotate-180' : ''
                }`}
              />
            </span>
          </button>

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

        {/* Always-visible Quick Teacher Ready-to-Print Bar */}
        <div className="mt-4 pt-3.5 border-t border-slate-100 flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          <div className="flex flex-col sm:flex-row sm:items-center gap-2.5 flex-1 min-w-0">
            <label
              htmlFor="rep-teacher-preset"
              className="text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 whitespace-nowrap shrink-0"
            >
              <UserCheck className="w-4 h-4 text-indigo-600" />
              Siap Cetak per Guru:
            </label>
            <select
              id="rep-teacher-preset"
              value={selectedTeacherId}
              onChange={(e) => applyTeacherPrintPreset(e.target.value, true)}
              className="w-full sm:max-w-md px-3 py-1.5 text-xs sm:text-sm border border-slate-200 rounded-lg bg-slate-50 focus:bg-white focus:outline-none focus:border-indigo-600 text-slate-800 font-medium"
            >
              <option value="">--Semua Guru (Filter Manual)--</option>
              {teachers.map((t) => {
                const summaryText = formatTeacherAssignmentsSummary(t);
                return (
                  <option key={t.id} value={t.id}>
                    {t.name}
                    {summaryText && summaryText !== '-' ? ` — ${summaryText}` : ''}
                  </option>
                );
              })}
            </select>
            {activeTeacher && (
              <button
                type="button"
                onClick={() => applyTeacherPrintPreset('', true)}
                className="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-medium text-slate-600 inline-flex items-center gap-1 whitespace-nowrap self-start sm:self-auto"
                title="Kembali ke filter manual semua kelas & mapel"
              >
                <RotateCcw className="w-3 h-3" />
                Reset Guru
              </button>
            )}
          </div>
          {activeTeacher && (
            <div className="flex flex-wrap items-center gap-1.5 text-[11px] text-indigo-800 bg-indigo-50 border border-indigo-200 rounded-lg px-2.5 py-1.5">
              <span className="font-semibold">Penugasan Aktif:</span>
              <span>
                {formatTeacherAssignmentsSummary(activeTeacher) || 'Belum ada mapel/kelas ditetapkan'}
              </span>
            </div>
          )}
        </div>

        {filterPanelExpanded && (
          <div className="space-y-5">
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
                  ? '--Pilih Jenis Laporan--'
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
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedReportTypes([]);
                      setReportDropdownOpen(false);
                    }}
                    className={`w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                      selectedReportTypes.length === 0
                        ? 'bg-indigo-50 text-indigo-700 font-semibold'
                        : 'text-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    --Pilih Jenis Laporan--
                  </button>
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
                  ? '--Pilih Kelas--'
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
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedClasses([]);
                      setClassDropdownOpen(false);
                    }}
                    className={`w-full text-left px-2.5 py-1.5 rounded-lg text-sm transition-colors ${
                      selectedClasses.length === 0
                        ? 'bg-indigo-50 text-indigo-700 font-semibold'
                        : 'text-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    --Pilih Kelas--
                  </button>
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

          <div className="relative" ref={subjectDropdownRef}>
            <span className="block text-xs font-semibold text-slate-700 mb-1.5">
              Mata Pelajaran (Multi-Select)
            </span>
            <button
              type="button"
              onClick={() => setSubjectDropdownOpen((o) => !o)}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white text-left flex items-center justify-between gap-2 focus:outline-none focus:border-indigo-600"
            >
              <span className="truncate text-slate-800">
                {selectedSubjects.length === 0
                  ? '--Pilih Mapel--'
                  : selectedSubjects.length === subjectOptions.length && subjectOptions.length > 0
                  ? `Semua Mata Pelajaran (${subjectOptions.length})`
                  : selectedSubjects.join(', ')}
              </span>
              <ChevronDown className="w-4 h-4 text-slate-400 shrink-0" />
            </button>

            {subjectDropdownOpen && (
              <div className="absolute left-0 right-0 mt-1.5 z-30 bg-white border border-slate-200 rounded-xl shadow-lg p-2 space-y-1">
                <div className="flex items-center justify-between px-2 py-1 border-b border-slate-100 text-xs">
                  <button
                    type="button"
                    onClick={() => setSelectedSubjects(subjectOptions)}
                    className="font-semibold text-indigo-600 hover:underline"
                  >
                    Pilih Semua
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedSubjects([])}
                    className="text-slate-500 hover:text-slate-700"
                  >
                    Kosongkan
                  </button>
                </div>
                <div className="max-h-48 overflow-y-auto space-y-0.5 pt-1">
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedSubjects([]);
                      setSubjectDropdownOpen(false);
                    }}
                    className={`w-full text-left px-2.5 py-1.5 rounded-lg text-sm transition-colors ${
                      selectedSubjects.length === 0
                        ? 'bg-indigo-50 text-indigo-700 font-semibold'
                        : 'text-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    --Pilih Mapel--
                  </button>
                  {subjectOptions.map((subj) => (
                    <label
                      key={subj}
                      className="flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer text-sm text-slate-800"
                    >
                      <input
                        type="checkbox"
                        checked={selectedSubjects.includes(subj)}
                        onChange={() => toggleSubjectSelection(subj)}
                        className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                      />
                      <span>{subj}</span>
                    </label>
                  ))}
                </div>
              </div>
            )}
          </div>

          <div>
            <label htmlFor="rep-paper" className="block text-xs font-semibold text-slate-700 mb-1.5">
              Ukuran Kertas
            </label>
            <select
              id="rep-paper"
              value={paperSize}
              onChange={(e) => setPaperSize(e.target.value as PaperSizeId | '')}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              <option value="">--Pilih Ukuran Kertas--</option>
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
              onChange={(e) => setFontFamily(e.target.value as PrintFontFamily | '')}
              className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              <option value="">--Pilih Jenis Huruf--</option>
              {FONT_OPTIONS.map((f) => (
                <option key={f} value={f} style={{ fontFamily: f }}>
                  {f}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Row 2: Tahun Ajaran, Semester, Date Range + Toggles */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 pt-2 border-t border-slate-100 items-end">
          <div className="lg:col-span-7 flex flex-wrap items-end gap-2.5">
            <div>
              <label htmlFor="rep-year" className="block text-xs font-semibold text-slate-700 mb-1">
                Tahun Ajaran
              </label>
              <select
                id="rep-year"
                value={reportAcademicYear}
                onChange={(e) => setReportAcademicYear(e.target.value)}
                className="px-3 py-1.5 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600 font-medium text-slate-800"
              >
                {reportYearOptions.map((yr) => (
                  <option key={yr} value={yr}>
                    {yr}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label htmlFor="rep-sem" className="block text-xs font-semibold text-slate-700 mb-1">
                Semester
              </label>
              <select
                id="rep-sem"
                value={reportSemester}
                onChange={(e) => setReportSemester(e.target.value as SemesterType)}
                className="px-3 py-1.5 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600 font-medium text-slate-800"
              >
                <option value="Ganjil">Ganjil</option>
                <option value="Genap">Genap</option>
              </select>
            </div>
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

          <div className="lg:col-span-5 flex flex-wrap items-center lg:justify-end gap-4 text-xs sm:text-sm">
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

        {/* Chart Types Panel when Laporan Grafis & Diagram is selected */}
        {includeCharts && (
          <div className="rounded-xl border border-violet-200 bg-violet-50/70 p-3.5 space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <BarChart3 className="w-4 h-4 text-violet-700" />
                <span className="text-xs font-bold text-violet-950">
                  Pilihan Jenis Grafik &amp; Diagram Laporan ({selectedChartTypes.length}/{REPORT_CHART_OPTIONS.length})
                </span>
              </div>
              <div className="flex items-center gap-2 text-xs">
                <button
                  type="button"
                  onClick={() => setSelectedChartTypes(REPORT_CHART_OPTIONS.map((c) => c.id))}
                  className="font-semibold text-violet-700 hover:underline"
                >
                  Pilih Semua Diagram
                </button>
                <span className="text-violet-300">|</span>
                <button
                  type="button"
                  onClick={() => setSelectedChartTypes([])}
                  className="text-slate-500 hover:text-slate-700"
                >
                  Kosongkan
                </button>
              </div>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2.5">
              {REPORT_CHART_OPTIONS.map((opt) => {
                const checked = selectedChartTypes.includes(opt.id);
                return (
                  <label
                    key={opt.id}
                    className={`flex items-start gap-2 p-2.5 rounded-lg border cursor-pointer transition-colors ${
                      checked
                        ? 'bg-white border-violet-400 shadow-2xs'
                        : 'bg-white/60 border-slate-200 hover:bg-white'
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => toggleChartTypeSelection(opt.id)}
                      className="mt-0.5 rounded border-slate-300 text-violet-600 focus:ring-violet-600"
                    />
                    <div className="min-w-0">
                      <p className="text-xs font-semibold text-slate-900 leading-tight">{opt.label}</p>
                      <p className="text-[11px] text-slate-500 mt-0.5 leading-snug">{opt.description}</p>
                    </div>
                  </label>
                );
              })}
            </div>
          </div>
        )}

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
                onChange={(e) => {
                  const nextId = e.target.value;
                  setSignerAccountId(nextId);
                  if (nextId && nextId !== '__AUTO__') {
                    const matchedTeacher = findTeacherForAccount(nextId);
                    if (
                      matchedTeacher &&
                      (getTeacherSubjects(matchedTeacher).length > 0 ||
                        getAllTeacherClasses(matchedTeacher).length > 0)
                    ) {
                      applyTeacherPrintPreset(matchedTeacher.id, true);
                    }
                  }
                }}
                className="w-full px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Penandatangan--</option>
                <option value="__AUTO__">Otomatis Sesuai Akun Login ({activeSigner.name})</option>
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
                onChange={(e) => handlePresetChange(e.target.value as SignatureLayoutPreset | '')}
                className="w-full px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Posisi Blok--</option>
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
                  Stempel Sekolah (Khusus Kepala Sekolah)
                </span>
                <div className="inline-flex items-center gap-2">
                  <div className="inline-flex items-center p-0.5 bg-slate-100 rounded-lg text-xs border border-slate-200">
                    <button
                      type="button"
                      onClick={() => updateSigPos({ showStamp: true, stampTarget: 'kepsek' })}
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
                      onClick={() => updateSigPos({ showStamp: false, stampTarget: 'kepsek' })}
                      className={`px-2.5 py-1 rounded-md font-semibold transition-colors ${
                        !sigPos.showStamp
                          ? 'bg-white text-rose-700 shadow-xs'
                          : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      Tidak
                    </button>
                  </div>
                  <span className="text-[11px] font-medium text-indigo-700 bg-indigo-50 border border-indigo-200 px-2 py-1 rounded-md">
                    Hanya Blok Kepsek
                  </span>
                </div>
              </div>
            </div>
          </div>

          <div className="bg-indigo-50/70 border border-indigo-200 rounded-xl p-3.5 space-y-3">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <button
                type="button"
                onClick={() => setSigPosExpanded((prev) => !prev)}
                aria-expanded={sigPosExpanded}
                className="flex items-center justify-between gap-3 text-left flex-1 min-w-0 group"
              >
                <div className="min-w-0">
                  <div className="flex flex-wrap items-center gap-2">
                    <span className="text-xs font-semibold text-indigo-900 group-hover:text-indigo-700 transition-colors">
                      Posisi &amp; Ukuran Bebas Gambar TTD / Stempel
                    </span>
                    <span className="text-[11px] text-indigo-700/80 tabular-nums">
                      · TTD Kepsek {kepsekSigSizePx}px · TTD Guru {guruSigSizePx}px · Stempel {stampSizePx}px
                    </span>
                  </div>
                  <p className="text-[11px] text-indigo-700/80 mt-0.5">
                    Seret gambar TTD / Stempel langsung di pratinjau bawah untuk memindahkan posisi, atau seret titik sudut kanan bawah gambar untuk memperbesar/memperkecil.
                  </p>
                </div>
                <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg border border-indigo-200 bg-white group-hover:bg-indigo-100/70 text-xs font-semibold text-indigo-700 whitespace-nowrap shrink-0 transition-colors">
                  <span>{sigPosExpanded ? 'Ciutkan' : 'Tampilkan'}</span>
                  <ChevronDown
                    className={`w-4 h-4 text-indigo-600 transition-transform duration-150 ${
                      sigPosExpanded ? 'rotate-180' : ''
                    }`}
                  />
                </span>
              </button>
            </div>

            {sigPosExpanded && (
              <>
                <div className="flex flex-wrap items-center justify-end gap-1.5 pt-2 border-t border-indigo-200/70">
                  <button
                    type="button"
                    onClick={() => {
                      updateSigPos({
                        principalSignatureSize: 60,
                        teacherSignatureSize: 60,
                        stampSize: 70,
                      });
                      onToast('info', 'Ukuran gambar TTD & stempel dikembalikan ke standar.');
                    }}
                    className="px-2.5 py-1 rounded-md bg-white border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100"
                  >
                    Reset Ukuran
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      updateSigPos({
                        principalSigOffsetX: 0,
                        principalSigOffsetY: 0,
                        teacherSigOffsetX: 0,
                        teacherSigOffsetY: 0,
                        stampOffsetX: 0,
                        stampOffsetY: 0,
                      });
                      onToast('info', 'Posisi offset gambar TTD & stempel direset.');
                    }}
                    className="px-2.5 py-1 rounded-md bg-white border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100"
                  >
                    Reset Posisi Gambar
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      updateSigPos({
                        preset: 'bebas',
                        principalX: 4,
                        principalY: 0,
                        teacherX: 62,
                        teacherY: 0,
                      });
                      onToast('info', 'Posisi blok tanda tangan direset ke standar.');
                    }}
                    className="px-2.5 py-1 rounded-md bg-white border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100"
                  >
                    Reset Posisi Blok
                  </button>
                </div>

                {/* Quick layer focus for overlapping stamp & signatures */}
                <div className="flex flex-wrap items-center gap-2 text-xs bg-white/70 border border-indigo-100 rounded-lg px-2.5 py-1.5">
              <span className="font-semibold text-slate-600">Pilih Objek Aktif di Pratinjau:</span>
              <button
                type="button"
                onClick={() =>
                  setSelectedMedia((prev) => (prev === 'kepsek-sig' ? null : 'kepsek-sig'))
                }
                className={`px-2 py-0.5 rounded font-semibold transition-colors ${
                  selectedMedia === 'kepsek-sig'
                    ? 'bg-indigo-600 text-white'
                    : 'bg-slate-100 text-slate-700 hover:bg-indigo-100'
                }`}
              >
                TTD Kepsek
              </button>
              <button
                type="button"
                onClick={() =>
                  setSelectedMedia((prev) => (prev === 'guru-sig' ? null : 'guru-sig'))
                }
                className={`px-2 py-0.5 rounded font-semibold transition-colors ${
                  selectedMedia === 'guru-sig'
                    ? 'bg-indigo-600 text-white'
                    : 'bg-slate-100 text-slate-700 hover:bg-indigo-100'
                }`}
              >
                TTD Guru
              </button>
              <button
                type="button"
                onClick={() => setSelectedMedia((prev) => (prev === 'stamp' ? null : 'stamp'))}
                className={`px-2 py-0.5 rounded font-semibold transition-colors ${
                  selectedMedia === 'stamp'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-slate-100 text-slate-700 hover:bg-emerald-100'
                }`}
              >
                Stempel Sekolah
              </button>
              {selectedMedia && (
                <button
                  type="button"
                  onClick={() => setSelectedMedia(null)}
                  className="ml-auto text-[11px] font-medium text-slate-500 hover:text-slate-800"
                >
                  Selesai Pilih
                </button>
              )}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
              <div className="bg-white/80 rounded-lg border border-indigo-100 p-2.5 space-y-1.5">
                <div className="flex items-center justify-between gap-1">
                  <span className="font-semibold text-slate-700">
                    Blok Kepsek Kiri ↔ Kanan: {sigPos.principalX}%
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
                          principalX: Math.min(82, sigPos.principalX + 1),
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
                  max={82}
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
                    Blok Kepsek Atas / Bawah: {sigPos.principalY}px
                  </span>
                </div>
                <input
                  type="range"
                  min={0}
                  max={240}
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
                    Blok Guru Kiri ↔ Kanan: {sigPos.teacherX}%
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
                          teacherX: Math.min(82, sigPos.teacherX + 1),
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
                  max={82}
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
                    Blok Guru Atas / Bawah: {sigPos.teacherY}px
                  </span>
                </div>
                <input
                  type="range"
                  min={0}
                  max={240}
                  value={sigPos.teacherY}
                  onChange={(e) =>
                    updateSigPos({ preset: 'bebas', teacherY: Number(e.target.value) })
                  }
                  className="w-full accent-indigo-600"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-3 text-xs">
              {/* Gambar TTD Kepsek */}
              <div
                className={`bg-white/90 rounded-lg border p-2.5 space-y-2 transition-colors ${
                  selectedMedia === 'kepsek-sig' ? 'border-indigo-500 ring-1 ring-indigo-300' : 'border-indigo-100'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <button
                    type="button"
                    onClick={() => setSelectedMedia('kepsek-sig')}
                    className="font-semibold text-slate-800 hover:text-indigo-700 text-left tabular-nums"
                  >
                    TTD Kepsek ({kepsekSigSizePx}px)
                  </button>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ principalSignatureSize: Math.max(16, kepsekSigSizePx - 4) })
                      }
                      aria-label="Perkecil TTD Kepala Sekolah"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <input
                      type="number"
                      min={16}
                      max={260}
                      value={kepsekSigSizePx}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        if (!Number.isNaN(v)) {
                          updateSigPos({ principalSignatureSize: Math.max(16, Math.min(260, v)) });
                        }
                      }}
                      aria-label="Ukuran TTD Kepala Sekolah px"
                      className="w-12 px-1 py-0.5 text-center border border-slate-200 rounded text-[11px] font-semibold tabular-nums"
                    />
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ principalSignatureSize: Math.min(260, kepsekSigSizePx + 4) })
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
                  min={16}
                  max={260}
                  value={kepsekSigSizePx}
                  onChange={(e) =>
                    updateSigPos({ principalSignatureSize: Number(e.target.value) })
                  }
                  aria-label="Slider ukuran TTD Kepala Sekolah"
                  className="w-full accent-indigo-600"
                />
                <div className="grid grid-cols-2 gap-2 pt-1 border-t border-slate-100">
                  <div>
                    <span className="block text-[10px] text-slate-500 tabular-nums">
                      Geser X: {kepsekSigOffsetX}px
                    </span>
                    <input
                      type="range"
                      min={-350}
                      max={350}
                      value={kepsekSigOffsetX}
                      onChange={(e) =>
                        updateSigPos({ principalSigOffsetX: Number(e.target.value) })
                      }
                      aria-label="Geser horizontal gambar TTD Kepsek"
                      className="w-full accent-indigo-600"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] text-slate-500 tabular-nums">
                      Geser Y: {kepsekSigOffsetY}px
                    </span>
                    <input
                      type="range"
                      min={-160}
                      max={200}
                      value={kepsekSigOffsetY}
                      onChange={(e) =>
                        updateSigPos({ principalSigOffsetY: Number(e.target.value) })
                      }
                      aria-label="Geser vertikal gambar TTD Kepsek"
                      className="w-full accent-indigo-600"
                    />
                  </div>
                </div>
              </div>

              {/* Gambar TTD Guru / Akun */}
              <div
                className={`bg-white/90 rounded-lg border p-2.5 space-y-2 transition-colors ${
                  selectedMedia === 'guru-sig' ? 'border-indigo-500 ring-1 ring-indigo-300' : 'border-indigo-100'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <button
                    type="button"
                    onClick={() => setSelectedMedia('guru-sig')}
                    className="font-semibold text-slate-800 hover:text-indigo-700 text-left tabular-nums"
                  >
                    TTD Guru ({guruSigSizePx}px)
                  </button>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ teacherSignatureSize: Math.max(16, guruSigSizePx - 4) })
                      }
                      aria-label="Perkecil TTD Guru"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <input
                      type="number"
                      min={16}
                      max={260}
                      value={guruSigSizePx}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        if (!Number.isNaN(v)) {
                          updateSigPos({ teacherSignatureSize: Math.max(16, Math.min(260, v)) });
                        }
                      }}
                      aria-label="Ukuran TTD Guru px"
                      className="w-12 px-1 py-0.5 text-center border border-slate-200 rounded text-[11px] font-semibold tabular-nums"
                    />
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ teacherSignatureSize: Math.min(260, guruSigSizePx + 4) })
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
                  min={16}
                  max={260}
                  value={guruSigSizePx}
                  onChange={(e) =>
                    updateSigPos({ teacherSignatureSize: Number(e.target.value) })
                  }
                  aria-label="Slider ukuran TTD Guru"
                  className="w-full accent-indigo-600"
                />
                <div className="grid grid-cols-2 gap-2 pt-1 border-t border-slate-100">
                  <div>
                    <span className="block text-[10px] text-slate-500 tabular-nums">
                      Geser X: {guruSigOffsetX}px
                    </span>
                    <input
                      type="range"
                      min={-350}
                      max={350}
                      value={guruSigOffsetX}
                      onChange={(e) =>
                        updateSigPos({ teacherSigOffsetX: Number(e.target.value) })
                      }
                      aria-label="Geser horizontal gambar TTD Guru"
                      className="w-full accent-indigo-600"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] text-slate-500 tabular-nums">
                      Geser Y: {guruSigOffsetY}px
                    </span>
                    <input
                      type="range"
                      min={-160}
                      max={200}
                      value={guruSigOffsetY}
                      onChange={(e) =>
                        updateSigPos({ teacherSigOffsetY: Number(e.target.value) })
                      }
                      aria-label="Geser vertikal gambar TTD Guru"
                      className="w-full accent-indigo-600"
                    />
                  </div>
                </div>
              </div>

              {/* Gambar Stempel Sekolah */}
              <div
                className={`bg-white/90 rounded-lg border p-2.5 space-y-2 transition-colors ${
                  selectedMedia === 'stamp' ? 'border-emerald-500 ring-1 ring-emerald-300' : 'border-indigo-100'
                }`}
              >
                <div className="flex items-center justify-between gap-2">
                  <button
                    type="button"
                    onClick={() => setSelectedMedia('stamp')}
                    className="font-semibold text-slate-800 hover:text-emerald-700 text-left tabular-nums"
                  >
                    Stempel ({stampSizePx}px)
                  </button>
                  <div className="inline-flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ stampSize: Math.max(20, stampSizePx - 4) })
                      }
                      aria-label="Perkecil gambar stempel"
                      className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                    >
                      -
                    </button>
                    <input
                      type="number"
                      min={20}
                      max={260}
                      value={stampSizePx}
                      onChange={(e) => {
                        const v = Number(e.target.value);
                        if (!Number.isNaN(v)) {
                          updateSigPos({ stampSize: Math.max(20, Math.min(260, v)) });
                        }
                      }}
                      aria-label="Ukuran gambar stempel px"
                      className="w-12 px-1 py-0.5 text-center border border-slate-200 rounded text-[11px] font-semibold tabular-nums"
                    />
                    <button
                      type="button"
                      onClick={() =>
                        updateSigPos({ stampSize: Math.min(260, stampSizePx + 4) })
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
                  min={20}
                  max={260}
                  value={stampSizePx}
                  onChange={(e) => updateSigPos({ stampSize: Number(e.target.value) })}
                  aria-label="Slider ukuran gambar stempel"
                  className="w-full accent-emerald-600"
                />
                <div className="grid grid-cols-2 gap-2 pt-1 border-t border-slate-100">
                  <div>
                    <span className="block text-[10px] text-slate-500 tabular-nums">
                      Geser X: {stampOffsetX}px
                    </span>
                    <input
                      type="range"
                      min={-550}
                      max={550}
                      value={stampOffsetX}
                      onChange={(e) => updateSigPos({ stampOffsetX: Number(e.target.value) })}
                      aria-label="Geser horizontal gambar stempel"
                      className="w-full accent-emerald-600"
                    />
                  </div>
                  <div>
                    <span className="block text-[10px] text-slate-500 tabular-nums">
                      Geser Y: {stampOffsetY}px
                    </span>
                    <input
                      type="range"
                      min={-200}
                      max={200}
                      value={stampOffsetY}
                      onChange={(e) => updateSigPos({ stampOffsetY: Number(e.target.value) })}
                      aria-label="Geser vertikal gambar stempel"
                      className="w-full accent-emerald-600"
                    />
                  </div>
                </div>
              </div>
            </div>
              </>
            )}
          </div>

          {showKop && (
            <div className="rounded-xl border border-slate-200 bg-slate-50/70 p-3 space-y-2.5">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <button
                  type="button"
                  onClick={() => setKopLogoExpanded((prev) => !prev)}
                  aria-expanded={kopLogoExpanded}
                  className="flex items-center justify-between gap-3 text-left flex-1 min-w-0 group"
                >
                  <div className="min-w-0 flex flex-wrap items-center gap-2">
                    <span className="text-xs font-semibold text-slate-800 group-hover:text-indigo-700 transition-colors">
                      Ukuran Logo &amp; Spasi Baris Kop Surat Dinas
                    </span>
                    <span className="text-[11px] text-slate-500 tabular-nums">
                      · Kiri {kopLeftSize}px · Kanan {kopRightSize}px · Spasi {kopLineSpacing.toFixed(2)}x
                    </span>
                  </div>
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg border border-slate-200 bg-white group-hover:bg-slate-100 text-xs font-semibold text-slate-700 whitespace-nowrap shrink-0 transition-colors">
                    <span>{kopLogoExpanded ? 'Ciutkan' : 'Tampilkan'}</span>
                    <ChevronDown
                      className={`w-4 h-4 text-slate-500 transition-transform duration-150 ${
                        kopLogoExpanded ? 'rotate-180' : ''
                      }`}
                    />
                  </span>
                </button>
              </div>
              {kopLogoExpanded && (
                <>
                  <div className="flex justify-end pt-2 border-t border-slate-200/80">
                    <button
                      type="button"
                      onClick={() => {
                        setKopLeftSize(72);
                        setKopRightSize(72);
                        setKopLineSpacing(DEFAULT_KOP_LINE_SPACING);
                        onToast('info', 'Ukuran logo (72px) & spasi baris kop surat (1.25x) dikembalikan ke standar.');
                      }}
                      className="px-2.5 py-1 rounded-md border border-slate-200 bg-white hover:bg-slate-100 text-[11px] font-semibold text-slate-600 inline-flex items-center gap-1"
                    >
                      <RotateCcw className="w-3 h-3" />
                      Reset Logo &amp; Spasi
                    </button>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
                <div className="bg-white rounded-lg border border-slate-200 p-2.5 space-y-1.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-slate-700 tabular-nums">
                      Logo Kiri ({kopLeftSize}px)
                    </span>
                    <div className="inline-flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => setKopLeftSize((v) => Math.max(28, v - 4))}
                        aria-label="Perkecil logo kiri kop surat"
                        className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                      >
                        -
                      </button>
                      <input
                        type="number"
                        min={28}
                        max={150}
                        value={kopLeftSize}
                        onChange={(e) => {
                          const v = Number(e.target.value);
                          if (!Number.isNaN(v)) {
                            setKopLeftSize(Math.max(28, Math.min(150, v)));
                          }
                        }}
                        aria-label="Ukuran logo kiri kop surat px"
                        className="w-12 px-1 py-0.5 text-center border border-slate-200 rounded text-[11px] font-semibold tabular-nums"
                      />
                      <button
                        type="button"
                        onClick={() => setKopLeftSize((v) => Math.min(150, v + 4))}
                        aria-label="Perbesar logo kiri kop surat"
                        className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                      >
                        +
                      </button>
                    </div>
                  </div>
                  <input
                    type="range"
                    min={28}
                    max={150}
                    value={kopLeftSize}
                    onChange={(e) => setKopLeftSize(Number(e.target.value))}
                    aria-label="Slider ukuran logo kiri kop surat"
                    className="w-full accent-indigo-600"
                  />
                </div>

                <div className="bg-white rounded-lg border border-slate-200 p-2.5 space-y-1.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-slate-700 tabular-nums">
                      Logo Kanan ({kopRightSize}px)
                    </span>
                    <div className="inline-flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => setKopRightSize((v) => Math.max(28, v - 4))}
                        aria-label="Perkecil logo kanan kop surat"
                        className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                      >
                        -
                      </button>
                      <input
                        type="number"
                        min={28}
                        max={150}
                        value={kopRightSize}
                        onChange={(e) => {
                          const v = Number(e.target.value);
                          if (!Number.isNaN(v)) {
                            setKopRightSize(Math.max(28, Math.min(150, v)));
                          }
                        }}
                        aria-label="Ukuran logo kanan kop surat px"
                        className="w-12 px-1 py-0.5 text-center border border-slate-200 rounded text-[11px] font-semibold tabular-nums"
                      />
                      <button
                        type="button"
                        onClick={() => setKopRightSize((v) => Math.min(150, v + 4))}
                        aria-label="Perbesar logo kanan kop surat"
                        className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                      >
                        +
                      </button>
                    </div>
                  </div>
                  <input
                    type="range"
                    min={28}
                    max={150}
                    value={kopRightSize}
                    onChange={(e) => setKopRightSize(Number(e.target.value))}
                    aria-label="Slider ukuran logo kanan kop surat"
                    className="w-full accent-indigo-600"
                  />
                </div>

                <div className="bg-white rounded-lg border border-slate-200 p-2.5 space-y-1.5">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-slate-700 tabular-nums">
                      Spasi Baris Kop ({kopLineSpacing.toFixed(2)}x)
                    </span>
                    <div className="inline-flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() =>
                          setKopLineSpacing((v) => normalizeKopLineSpacing(v - 0.05))
                        }
                        aria-label="Rapatkan spasi baris kop surat"
                        className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                      >
                        -
                      </button>
                      <input
                        type="number"
                        min={0.85}
                        max={2.4}
                        step={0.05}
                        value={kopLineSpacing}
                        onChange={(e) => {
                          const v = Number(e.target.value);
                          if (!Number.isNaN(v)) {
                            setKopLineSpacing(normalizeKopLineSpacing(v));
                          }
                        }}
                        aria-label="Spasi antar baris kop surat"
                        className="w-14 px-1 py-0.5 text-center border border-slate-200 rounded text-[11px] font-semibold tabular-nums"
                      />
                      <button
                        type="button"
                        onClick={() =>
                          setKopLineSpacing((v) => normalizeKopLineSpacing(v + 0.05))
                        }
                        aria-label="Renggangkan spasi baris kop surat"
                        className="w-5 h-5 rounded border border-slate-200 bg-white hover:bg-slate-100 font-bold text-slate-700"
                      >
                        +
                      </button>
                    </div>
                  </div>
                  <input
                    type="range"
                    min={0.85}
                    max={2.4}
                    step={0.05}
                    value={kopLineSpacing}
                    onChange={(e) => setKopLineSpacing(normalizeKopLineSpacing(e.target.value))}
                    aria-label="Slider spasi baris kop surat"
                    className="w-full accent-indigo-600"
                  />
                </div>
              </div>
                </>
              )}
            </div>
          )}

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
          </div>
        )}
      </section>

      {/* Real-Time Document Preview (Exact 1:1 source for Print, PDF & Word — per class, separate sheet per report type) */}
      <section className="overflow-x-auto pb-4 print:overflow-visible print:pb-0">
        <div ref={previewRef} className="print-area space-y-6 print:space-y-0">
          {reportSheets.length === 0 ? (
            <div
              style={{
                width: `${activePaperWidthMm}mm`,
                minWidth: `${activePaperWidthMm}mm`,
                padding: '12mm',
                fontFamily: getPrintFontCss(effectiveFontFamily),
                backgroundColor: '#ffffff',
                color: '#000000',
              }}
              className="mx-auto bg-white border border-slate-200 shadow-sm text-black print:!w-full print:!min-w-0 print:!p-0 print:!border-0 print:!shadow-none"
            >
              <div
                style={{
                  border: '1px dashed #94a3b8',
                  padding: '24px',
                  textAlign: 'center',
                  color: '#475569',
                  fontSize: '10pt',
                }}
              >
                {orderedSelectedReportTypes.length === 0 ? (
                  <>
                    Belum ada jenis laporan yang dipilih. Silakan centang minimal 1 jenis laporan di menu{' '}
                    <strong>Jenis Laporan (Multi-Select)</strong> di atas.
                  </>
                ) : (
                  <>
                    Belum ada kelas yang dipilih. Silakan centang minimal 1 kelas di menu{' '}
                    <strong>Pilih Kelas (Multi-Select)</strong> di atas.
                  </>
                )}
              </div>
            </div>
          ) : (
            reportSheets.map((sheet, sheetIdx) => {
              const cls = sheet.className;

              const renderSignatureZone = () => {
                if (!showTtdKepsek && !showTtdGuru) return null;
                return (
                  <div className="print-avoid-break" style={{ marginTop: '24px', paddingTop: '8px' }}>
                    {sigPos.preset === 'bebas' ? (
                      <div
                        ref={sheetIdx === 0 ? sigCanvasRef : undefined}
                        data-sig-canvas="true"
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
                            title="Seret teks blok untuk mengatur posisi Kepala Sekolah"
                            style={{
                              position: 'absolute',
                              left: `${sigPos.principalX}%`,
                              top: `${sigPos.principalY}px`,
                              width: '220px',
                              cursor:
                                activeInteraction?.kind === 'block' &&
                                activeInteraction.target === 'kepsek'
                                  ? 'grabbing'
                                  : 'grab',
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
                              Geser Blok ({sigPos.principalX}%, {sigPos.principalY}px)
                            </span>
                            {renderSigCard(
                              'kepsek',
                              'Mengetahui,',
                              settings.principalTitle,
                              settings.principalName,
                              settings.principalNip,
                              showSignatureImage ? settings.principalSignatureImage : '',
                              stampOnKepsek,
                              kepsekSigSizePx,
                              kepsekSigOffsetX,
                              kepsekSigOffsetY
                            )}
                          </div>
                        )}
                        {showTtdGuru && (
                          <div
                            data-sig-block="guru"
                            onPointerDown={(e) => handleSigPointerDown('guru', e)}
                            title="Seret teks blok untuk mengatur posisi Guru / Pemilik Akun"
                            style={{
                              position: 'absolute',
                              left: `${sigPos.teacherX}%`,
                              top: `${sigPos.teacherY}px`,
                              width: '220px',
                              cursor:
                                activeInteraction?.kind === 'block' &&
                                activeInteraction.target === 'guru'
                                  ? 'grabbing'
                                  : 'grab',
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
                              Geser Blok ({sigPos.teacherX}%, {sigPos.teacherY}px)
                            </span>
                            {renderSigCard(
                              'guru',
                              `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                              activeSigner.roleLabel,
                              activeSigner.name,
                              activeSigner.nip,
                              showSignatureImage ? activeSigner.signatureImage : '',
                              stampOnGuru,
                              guruSigSizePx,
                              guruSigOffsetX,
                              guruSigOffsetY
                            )}
                          </div>
                        )}
                      </div>
                    ) : (
                      <div ref={sheetIdx === 0 ? sigCanvasRef : undefined} style={{ width: '100%' }}>
                        {sigPos.preset === 'kiri' ? (
                          <div style={{ paddingLeft: '16px' }}>
                            {showTtdKepsek && (
                              <div
                                onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                title="Seret blok untuk menyesuaikan posisi"
                                style={{ width: '220px', marginBottom: '20px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  'kepsek',
                                  'Mengetahui,',
                                  settings.principalTitle,
                                  settings.principalName,
                                  settings.principalNip,
                                  showSignatureImage ? settings.principalSignatureImage : '',
                                  stampOnKepsek,
                                  kepsekSigSizePx,
                                  kepsekSigOffsetX,
                                  kepsekSigOffsetY
                                )}
                              </div>
                            )}
                            {showTtdGuru && (
                              <div
                                onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                title="Seret blok untuk menyesuaikan posisi"
                                style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                              >
                                {renderSigCard(
                                  'guru',
                                  `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                  activeSigner.roleLabel,
                                  activeSigner.name,
                                  activeSigner.nip,
                                  showSignatureImage ? activeSigner.signatureImage : '',
                                  stampOnGuru,
                                  guruSigSizePx,
                                  guruSigOffsetX,
                                  guruSigOffsetY
                                )}
                              </div>
                            )}
                          </div>
                        ) : sigPos.preset === 'kanan' ? (
                          <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse' }}>
                            <tbody>
                              <tr>
                                <td style={{ border: 'none', padding: 0 }} />
                                <td style={{ width: '240px', verticalAlign: 'top', border: 'none', padding: 0 }}>
                                  {showTtdKepsek && (
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ marginBottom: '20px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'kepsek',
                                        'Mengetahui,',
                                        settings.principalTitle,
                                        settings.principalName,
                                        settings.principalNip,
                                        showSignatureImage ? settings.principalSignatureImage : '',
                                        stampOnKepsek,
                                        kepsekSigSizePx,
                                        kepsekSigOffsetX,
                                        kepsekSigOffsetY
                                      )}
                                    </div>
                                  )}
                                  {showTtdGuru && (
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'guru',
                                        `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                        activeSigner.roleLabel,
                                        activeSigner.name,
                                        activeSigner.nip,
                                        showSignatureImage ? activeSigner.signatureImage : '',
                                        stampOnGuru,
                                        guruSigSizePx,
                                        guruSigOffsetX,
                                        guruSigOffsetY
                                      )}
                                    </div>
                                  )}
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        ) : sigPos.preset === 'tengah' ? (
                          <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'center' }}>
                            <tbody>
                              <tr>
                                {showTtdKepsek && (
                                  <td style={{ verticalAlign: 'top', textAlign: 'center', border: 'none' }}>
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ display: 'inline-block', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'kepsek',
                                        'Mengetahui,',
                                        settings.principalTitle,
                                        settings.principalName,
                                        settings.principalNip,
                                        showSignatureImage ? settings.principalSignatureImage : '',
                                        stampOnKepsek,
                                        kepsekSigSizePx,
                                        kepsekSigOffsetX,
                                        kepsekSigOffsetY
                                      )}
                                    </div>
                                  </td>
                                )}
                                {showTtdGuru && (
                                  <td style={{ verticalAlign: 'top', textAlign: 'center', border: 'none' }}>
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ display: 'inline-block', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'guru',
                                        `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                        activeSigner.roleLabel,
                                        activeSigner.name,
                                        activeSigner.nip,
                                        showSignatureImage ? activeSigner.signatureImage : '',
                                        stampOnGuru,
                                        guruSigSizePx,
                                        guruSigOffsetX,
                                        guruSigOffsetY
                                      )}
                                    </div>
                                  </td>
                                )}
                              </tr>
                            </tbody>
                          </table>
                        ) : sigPos.preset === 'kanan-kiri' ? (
                          <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse' }}>
                            <tbody>
                              <tr>
                                <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '16px', border: 'none' }}>
                                  {showTtdGuru && (
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'guru',
                                        `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                        activeSigner.roleLabel,
                                        activeSigner.name,
                                        activeSigner.nip,
                                        showSignatureImage ? activeSigner.signatureImage : '',
                                        stampOnGuru,
                                        guruSigSizePx,
                                        guruSigOffsetX,
                                        guruSigOffsetY
                                      )}
                                    </div>
                                  )}
                                </td>
                                <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '18%', border: 'none' }}>
                                  {showTtdKepsek && (
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'kepsek',
                                        'Mengetahui,',
                                        settings.principalTitle,
                                        settings.principalName,
                                        settings.principalNip,
                                        showSignatureImage ? settings.principalSignatureImage : '',
                                        stampOnKepsek,
                                        kepsekSigSizePx,
                                        kepsekSigOffsetX,
                                        kepsekSigOffsetY
                                      )}
                                    </div>
                                  )}
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        ) : (
                          <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse' }}>
                            <tbody>
                              <tr>
                                <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '16px', border: 'none' }}>
                                  {showTtdKepsek && (
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('kepsek', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'kepsek',
                                        'Mengetahui,',
                                        settings.principalTitle,
                                        settings.principalName,
                                        settings.principalNip,
                                        showSignatureImage ? settings.principalSignatureImage : '',
                                        stampOnKepsek,
                                        kepsekSigSizePx,
                                        kepsekSigOffsetX,
                                        kepsekSigOffsetY
                                      )}
                                    </div>
                                  )}
                                </td>
                                <td style={{ width: '50%', verticalAlign: 'top', paddingLeft: '18%', border: 'none' }}>
                                  {showTtdGuru && (
                                    <div
                                      onPointerDown={(e) => handleSigPointerDown('guru', e)}
                                      title="Seret blok untuk menyesuaikan posisi"
                                      style={{ width: '220px', cursor: 'grab', touchAction: 'none', userSelect: 'none' }}
                                      className="rounded hover:outline-1 hover:outline-dashed hover:outline-indigo-400"
                                    >
                                      {renderSigCard(
                                        'guru',
                                        `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
                                        activeSigner.roleLabel,
                                        activeSigner.name,
                                        activeSigner.nip,
                                        showSignatureImage ? activeSigner.signatureImage : '',
                                        stampOnGuru,
                                        guruSigSizePx,
                                        guruSigOffsetX,
                                        guruSigOffsetY
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
                );
              };

              return (
                <div
                  key={sheet.key}
                  data-report-sheet="true"
                  data-page-break={sheetIdx > 0 ? 'true' : 'false'}
                  style={{
                    width: `${activePaperWidthMm}mm`,
                    minWidth: `${activePaperWidthMm}mm`,
                    padding: '12mm',
                    fontFamily: getPrintFontCss(effectiveFontFamily),
                    backgroundColor: '#ffffff',
                    color: '#000000',
                  }}
                  className={`mx-auto bg-white border border-slate-200 shadow-sm text-black print:!w-full print:!min-w-0 print:!p-0 print:!border-0 print:!shadow-none print:!m-0 ${
                    sheetIdx > 0 ? 'print-page-break' : ''
                  }`}
                >
                  {/* Screen-only page indicator */}
                  {reportSheets.length > 1 && (
                    <div
                      data-export-ignore="true"
                      className="no-print mb-3 pb-2 border-b border-dashed border-slate-200 flex items-center justify-between text-[11px] font-sans text-slate-500"
                    >
                      <span className="font-semibold text-indigo-700">
                        Lembaran {sheetIdx + 1} dari {reportSheets.length}
                        {cls ? ` · Kelas ${cls}` : ''} — {sheet.label}
                      </span>
                      <span>Lembaran terpisah saat cetak / unduh PDF &amp; Word</span>
                    </div>
                  )}

                  {/* KOP SURAT DINAS */}
                  {showKop && (
                    <div
                      data-kop-wrapper="true"
                      style={{ borderBottom: '4px double #000000', paddingBottom: '10px', marginBottom: '16px' }}
                    >
                      <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse' }}>
                        <tbody>
                          <tr>
                            <td
                              style={{
                                width: `${Math.max(76, kopLeftSize + 14)}px`,
                                textAlign: 'center',
                                verticalAlign: 'middle',
                                padding: '0 6px',
                                border: 'none',
                              }}
                            >
                              {settings.kopSurat.logoLeft && (
                                <img
                                  src={settings.kopSurat.logoLeft}
                                  alt="Logo Instansi Kiri"
                                  style={{
                                    width: `${kopLeftSize}px`,
                                    height: `${kopLeftSize}px`,
                                    objectFit: 'contain',
                                    display: 'inline-block',
                                  }}
                                />
                              )}
                            </td>
                            <td
                              style={{
                                fontFamily: getPrintFontCss(settings.kopSurat.fontFamily),
                                verticalAlign: 'middle',
                                lineHeight: `${Math.round(kopLineSpacing * 100)}%`,
                                padding: '0 4px',
                                border: 'none',
                              }}
                            >
                              {KOP_LINE_KEYS.map((key, idx) => {
                                const text = settings.kopSurat[key];
                                if (!text) return null;
                                const s = resolveKopLineStyle(settings.kopSurat, key);
                                return (
                                  <div
                                    key={key}
                                    style={{
                                      margin: 0,
                                      padding: 0,
                                      lineHeight: `${Math.round(kopLineSpacing * 100)}%`,
                                      fontSize: `${s.fontSize}pt`,
                                      fontWeight: s.bold ? 700 : 400,
                                      fontStyle: s.italic ? 'italic' : 'normal',
                                      textDecoration: s.underline ? 'underline' : 'none',
                                      textAlign: s.align,
                                      ...(s.align === 'justify' ? { textAlignLast: 'justify' as const } : {}),
                                      ...(idx < 3 ? { textTransform: 'uppercase' as const } : {}),
                                      ...(key === 'line4' ? { marginTop: '3px' } : {}),
                                    }}
                                  >
                                    {text}
                                  </div>
                                );
                              })}
                            </td>
                            <td
                              style={{
                                width: `${Math.max(76, kopRightSize + 14)}px`,
                                textAlign: 'center',
                                verticalAlign: 'middle',
                                padding: '0 6px',
                                border: 'none',
                              }}
                            >
                              {settings.kopSurat.logoRight && (
                                <img
                                  src={settings.kopSurat.logoRight}
                                  alt="Logo Sekolah Kanan"
                                  style={{
                                    width: `${kopRightSize}px`,
                                    height: `${kopRightSize}px`,
                                    objectFit: 'contain',
                                    display: 'inline-block',
                                  }}
                                />
                              )}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  )}

                  {/* Judul Dokumen per Lembaran */}
                  <div style={{ textAlign: 'center', marginBottom: '14px' }}>
                    <h2
                      style={{
                        fontSize: '12pt',
                        fontWeight: 700,
                        textTransform: 'uppercase',
                        textDecoration: 'underline',
                        textUnderlineOffset: '3px',
                        margin: 0,
                      }}
                    >
                      {sheet.officialTitle}
                    </h2>
                    <p style={{ fontSize: '9pt', marginTop: '4px', marginBottom: 0, color: '#334155' }}>
                      {sheet.subtitle}
                    </p>
                  </div>

                  {/* 1. Jurnal Mengajar per Kelas */}
                  {sheet.reportType === 'journal' && cls && (() => {
                    const classJournals = filteredJournals.filter((j) => j.className === cls);
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                          Jurnal Kegiatan Mengajar — Kelas {cls}
                        </h3>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
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
                            {classJournals.length === 0 ? (
                              <tr>
                                <td colSpan={8} style={{ border: '1px solid #000', padding: '12px', textAlign: 'center', color: '#64748b' }}>
                                  Tidak ada catatan jurnal mengajar untuk kelas {cls} pada filter terpilih.
                                </td>
                              </tr>
                            ) : (
                              classJournals.map((j, idx) => {
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
                    );
                  })()}

                  {/* 2. Leger Kehadiran Mendatar per Kelas */}
                  {sheet.reportType === 'attendance' && cls && (() => {
                    const { classStudents, classJournals } = getAttendanceBlockForClass(cls);
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                          Leger Kehadiran Mendatar — Kelas {cls}
                        </h3>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '7.5pt' }}>
                          <thead>
                            <tr style={{ backgroundColor: '#f1f5f9' }}>
                              <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', width: '28px' }}>No</th>
                              <th style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }}>NIS</th>
                              <th style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }}>NISN</th>
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
                                <td colSpan={10 + classJournals.length} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
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
                                    <td style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }} className="tabular-nums">{st.nisn || '-'}</td>
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
                  })()}

                  {/* 3. Rekap Presensi Bulanan per Kelas */}
                  {sheet.reportType === 'attendance_monthly' && cls && (() => {
                    const { months, rows, classAvgPct } = getMonthlyAttendanceForClass(cls);
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '6px' }}>
                          <tbody>
                            <tr>
                              <td style={{ border: 'none', padding: 0, textAlign: 'left', verticalAlign: 'baseline' }}>
                                <h3 style={{ fontSize: '10pt', fontWeight: 700, margin: 0 }}>
                                  Rekap Presensi Bulanan — Kelas {cls}
                                </h3>
                              </td>
                              <td style={{ border: 'none', padding: 0, textAlign: 'right', verticalAlign: 'baseline', fontSize: '8pt', color: '#334155' }}>
                                Rata-rata Kehadiran Kelas: {classAvgPct !== null ? `${classAvgPct.toFixed(1)}%` : '-'}
                              </td>
                            </tr>
                          </tbody>
                        </table>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
                          <thead>
                            <tr style={{ backgroundColor: '#f1f5f9' }}>
                              <th style={{ border: '1px solid #000', padding: '4px', textAlign: 'center', width: '28px' }}>No</th>
                              <th style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }}>NIS</th>
                              <th style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }}>NISN</th>
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
                                <td colSpan={11 + months.length} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                                  Belum ada siswa di kelas {cls}.
                                </td>
                              </tr>
                            ) : (
                              rows.map((r) => (
                                <tr key={`m-${cls}-${r.nis}-${r.no}`}>
                                  <td style={{ border: '1px solid #000', padding: '4px', textAlign: 'center' }} className="tabular-nums">{r.no}</td>
                                  <td style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }} className="tabular-nums">{r.nis}</td>
                                  <td style={{ border: '1px solid #000', padding: '4px', whiteSpace: 'nowrap' }} className="tabular-nums">{r.nisn}</td>
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
                  })()}

                  {/* 4. Monitoring S/I/A/D per Kelas */}
                  {sheet.reportType === 'attendance_monitoring' && cls && (() => {
                    const clsMonitoring = getMonitoringAbsenceForClass(cls);
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        <div style={{ marginBottom: '14px' }}>
                          <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                            Monitoring S/I/A/D — Rekap Akumulasi Ketidakhadiran Kelas {cls}
                          </h3>
                          <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
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
                              {clsMonitoring.summary.length === 0 ? (
                                <tr>
                                  <td colSpan={10} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                                    Seluruh siswa kelas {cls} tercatat Hadir (H) 100% pada periode filter ini.
                                  </td>
                                </tr>
                              ) : (
                                clsMonitoring.summary.map((it, idx) => (
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
                          <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                            Kronologi Ketidakhadiran (S/I/A/D) — Kelas {cls}
                          </h3>
                          <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
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
                              {clsMonitoring.chronology.length === 0 ? (
                                <tr>
                                  <td colSpan={8} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                                    Belum ada catatan ketidakhadiran di kelas {cls} pada periode filter ini.
                                  </td>
                                </tr>
                              ) : (
                                clsMonitoring.chronology.map((c, idx) => (
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
                    );
                  })()}

                  {/* 5. Rekap Nilai Mapel per Kelas */}
                  {sheet.reportType === 'grades' && cls && (() => {
                    const classGradeSubjects = getSubjectsForClassSheet(cls);
                    const visibleSubjects = classGradeSubjects.filter((subj) => {
                      const { rows } = getGradeBlockForClassAndSubject(cls, subj);
                      return rows.some((r) => r.avg !== null) || classGradeSubjects.length === 1;
                    });
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        {visibleSubjects.length === 0 ? (
                          <div style={{ border: '1px solid #000', padding: '12px', textAlign: 'center', fontSize: '9pt', color: '#64748b' }}>
                            Belum ada nilai tercatat untuk kelas {cls} pada mata pelajaran terpilih.
                          </div>
                        ) : (
                          visibleSubjects.map((subj) => {
                            const { rows } = getGradeBlockForClassAndSubject(cls, subj);
                            return (
                              <div key={`${cls}-${subj}`} style={{ marginBottom: '16px' }}>
                                <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                                  Rekap Matriks Nilai — Kelas {cls} ({subj})
                                </h3>
                                <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
                                  <thead>
                                    <tr style={{ backgroundColor: '#f1f5f9' }}>
                                      <th style={{ border: '1px solid #000', padding: '4px 6px', textAlign: 'center', width: '30px' }}>No</th>
                                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NIS</th>
                                      <th style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }}>NISN</th>
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
                                          colSpan={6 + settings.masterExamTypes.length}
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
                                          <td style={{ border: '1px solid #000', padding: '4px 6px', whiteSpace: 'nowrap' }} className="tabular-nums">{r.nisn}</td>
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
                    );
                  })()}

                  {/* 6. Ketuntasan & Remedial (< KKM 75) per Kelas */}
                  {sheet.reportType === 'grades_remedial' && cls && (() => {
                    const classGradeSubjects = getSubjectsForClassSheet(cls);
                    const visibleSubjects = classGradeSubjects.filter((subj) => {
                      const rem = getRemedialDataForClassAndSubject(cls, subj);
                      return rem.assessedCount > 0 || classGradeSubjects.length === 1;
                    });
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        {visibleSubjects.length === 0 ? (
                          <div style={{ border: '1px solid #000', padding: '12px', textAlign: 'center', fontSize: '9pt', color: '#64748b' }}>
                            Belum ada data ketuntasan &amp; remedial untuk kelas {cls} pada mata pelajaran terpilih.
                          </div>
                        ) : (
                          visibleSubjects.map((subj) => {
                            const rem = getRemedialDataForClassAndSubject(cls, subj);
                            return (
                              <div key={`rem-${cls}-${subj}`} style={{ marginBottom: '16px' }}>
                                <table data-layout-table="true" style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '6px' }}>
                                  <tbody>
                                    <tr>
                                      <td style={{ border: 'none', padding: 0, textAlign: 'left', verticalAlign: 'baseline' }}>
                                        <h3 style={{ fontSize: '10pt', fontWeight: 700, margin: 0 }}>
                                          Ketuntasan &amp; Remedial (&lt; KKM 75) — Kelas {cls} ({subj})
                                        </h3>
                                      </td>
                                      <td style={{ border: 'none', padding: 0, textAlign: 'right', verticalAlign: 'baseline', fontSize: '8pt', color: '#334155' }}>
                                        Daya Serap: <strong>{rem.dayaSerap !== null ? `${rem.dayaSerap.toFixed(0)}%` : '-'}</strong> · Tuntas: {rem.tuntasCount}/{rem.assessedCount} · Remedial: {rem.remedialCount} · Rata Kelas: {rem.classAvg !== null ? rem.classAvg.toFixed(1) : '-'}
                                      </td>
                                    </tr>
                                  </tbody>
                                </table>
                                <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
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
                    );
                  })()}

                  {/* 7. Leger Gabungan Kelas */}
                  {sheet.reportType === 'grades_leger' && cls && (() => {
                    const { subjects, rows } = getCombinedLegerForClass(cls);
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                          Leger Gabungan Nilai Lintas Mapel &amp; Peringkat — Kelas {cls}
                        </h3>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '7.5pt' }}>
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
                  })()}

                  {/* 8. Daftar Peserta Didik per Kelas */}
                  {sheet.reportType === 'students_roster' && cls && (() => {
                    const clsStudents = studentRosterRows.filter((s) => s.className === cls);
                    return (
                      <div style={{ marginBottom: '20px' }}>
                        <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                          Daftar Induk Peserta Didik &amp; Wali Kelas — Kelas {cls} ({clsStudents.length} Siswa)
                        </h3>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '7.5pt' }}>
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
                            {clsStudents.length === 0 ? (
                              <tr>
                                <td colSpan={11} style={{ border: '1px solid #000', padding: '10px', textAlign: 'center', color: '#64748b' }}>
                                  Belum ada data peserta didik di kelas {cls}.
                                </td>
                              </tr>
                            ) : (
                              clsStudents.map((st, idx) => (
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
                    );
                  })()}

                  {/* 9. Aktivitas & Penilaian Guru (Pimpinan) */}
                  {sheet.reportType === 'teacher_activity' && (
                    <div style={{ marginBottom: '20px' }}>
                      <div style={{ marginBottom: '14px' }}>
                        <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                          Ringkasan Aktivitas &amp; Penilaian Kinerja Guru
                        </h3>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
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
                        <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '6px' }}>
                          Riwayat Kinerja &amp; Pelaksanaan KBM Terbaru
                        </h3>
                        <table className="report-data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8pt' }}>
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

                  {/* 10. Laporan Grafis & Diagram (Pimpinan) */}
                  {sheet.reportType === 'charts' && (
                    <div style={{ marginBottom: '20px' }}>
                      <h3 style={{ fontSize: '10pt', fontWeight: 700, marginBottom: '8px' }}>
                        Laporan Grafis &amp; Diagram Analisis KBM
                      </h3>
                      {chartCards.length === 0 ? (
                        <div
                          style={{
                            border: '1px dashed #94a3b8',
                            padding: '16px',
                            textAlign: 'center',
                            fontSize: '9pt',
                            color: '#64748b',
                          }}
                        >
                          Belum ada jenis diagram grafik yang dipilih pada panel filter di atas.
                        </div>
                      ) : (
                        <div
                          data-chart-grid="true"
                          data-chart-cols={chartCols}
                          style={{
                            display: 'grid',
                            gridTemplateColumns: `repeat(${chartCols}, minmax(0, 1fr))`,
                            gap: '12px',
                          }}
                        >
                          {chartCards.map((card) => (
                            <div
                              key={card.id}
                              data-chart-card="true"
                              className="print-avoid-break"
                              style={{
                                border: '1px solid #000000',
                                backgroundColor: '#ffffff',
                                padding: '8px',
                              }}
                            >
                              <div style={{ fontSize: '8.5pt', fontWeight: 700, color: '#000000', lineHeight: 1.25 }}>
                                {card.title}
                              </div>
                              <div style={{ fontSize: '7pt', color: '#475569', marginTop: '2px', marginBottom: '6px' }}>
                                {card.chartKind} · {card.subtitle}
                              </div>
                              {card.pngDataUrl && (
                                <img
                                  src={card.pngDataUrl}
                                  alt={card.title}
                                  style={{ width: '100%', height: 'auto', display: 'block' }}
                                />
                              )}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}

                  {/* Blok Pengesahan per Lembaran */}
                  {renderSignatureZone()}
                </div>
              );
            })
          )}
        </div>
      </section>
    </div>
  );
}
