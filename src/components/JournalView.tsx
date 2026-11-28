import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  CheckCheck,
  Copy,
  Pencil,
  Plus,
  Save,
  Search,
  Trash2,
  Upload,
  ChevronDown,
} from 'lucide-react';
import {
  AttendanceRecord,
  AttendanceStatus,
  JournalEntry,
  PortalDatabase,
  Student,
} from '../types';
import {
  exportDatasetToPdf,
  exportDatasetToWordDoc,
  exportRowsToXls,
  exportRowsToXlsx,
  printDatasetExact,
} from '../utils/exportHelpers';
import {
  ATTENDANCE_STATUS_META,
  compareClassNames,
  compareStudentsByClassAndAbsent,
  formatCompactDateId,
  formatShortDateId,
  getTeacherSubjects,
  getTodayISO,
  summarizeJournalRow,
} from '../utils/portalHelpers';
import { ConfirmDialogState } from './FeedbackUI';
import { FormModal } from './FormModal';
import { StudentDetailModal } from './StudentDetailModal';
import { TableExportToolbar } from './TableExportToolbar';

interface JournalViewProps {
  db: PortalDatabase;
  draftJournal: { mode: 'edit' | 'duplicate'; entry: JournalEntry } | null;
  onClearDraft: () => void;
  onUpsertJournal: (entry: JournalEntry) => void;
  onImportJournalsBatch: (incoming: Omit<JournalEntry, 'id' | 'createdAt'>[]) => number;
  onDeleteJournals: (ids: string[]) => void;
  onBatchUpdateJournals: (entries: JournalEntry[]) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const STATUS_KEYS: AttendanceStatus[] = ['H', 'S', 'I', 'A', 'D'];

export function JournalView({
  db,
  draftJournal,
  onClearDraft,
  onUpsertJournal,
  onImportJournalsBatch,
  onDeleteJournals,
  onBatchUpdateJournals,
  onToast,
  onConfirm,
}: JournalViewProps) {
  const { students, teachers, journals, settings, session } = db;

  // Section 1: Form state
  const [editingId, setEditingId] = useState<string | null>(null);
  const [date, setDate] = useState<string>(getTodayISO());
  const [jamKe, setJamKe] = useState<string>('1-2');
  const [className, setClassName] = useState<string>('');
  const [subject, setSubject] = useState<string>('');
  const [topic, setTopic] = useState<string>('');
  const [activity, setActivity] = useState<string>('');
  const [attendanceMap, setAttendanceMap] = useState<Record<string, AttendanceRecord>>({});

  // Section 2: History filter, bulk selection & import
  const [historyClassFilter, setHistoryClassFilter] = useState<string>('');
  const [historySearch, setHistorySearch] = useState<string>('');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [showImportBox, setShowImportBox] = useState(false);
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Section 3: Horizontal Attendance Ledger
  const [ledgerClass, setLedgerClass] = useState<string>('');
  const [ledgerSubject, setLedgerSubject] = useState<string>('');
  const [ledgerEdits, setLedgerEdits] = useState<Record<string, Record<string, AttendanceStatus>>>({});
  const [ledgerDirty, setLedgerDirty] = useState(false);
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [detailStudent, setDetailStudent] = useState<Student | null>(null);
  const [historyExpanded, setHistoryExpanded] = useState(false);
  const [ledgerExpanded, setLedgerExpanded] = useState(false);

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
    return Array.from(set);
  }, [settings.masterSubjects, teachers]);

  // Apply external draft (from Dashboard Edit/Duplicate)
  useEffect(() => {
    if (!draftJournal) return;
    const { mode, entry } = draftJournal;
    setEditingId(mode === 'edit' ? entry.id : null);
    setDate(mode === 'edit' ? entry.date : getTodayISO());
    setJamKe(entry.jamKe);
    setClassName(entry.className);
    setSubject(entry.subject);
    setTopic(
      mode === 'edit'
        ? entry.topic
        : entry.topic
        ? `${entry.topic} (Salinan)`
        : ''
    );
    setActivity(entry.activity);
    setAttendanceMap(JSON.parse(JSON.stringify(entry.attendance)));
    setFormModalOpen(true);
    onClearDraft();
  }, [draftJournal, onClearDraft]);

  const studentsInSelectedClass = useMemo(
    () =>
      students
        .filter((s) => s.className === className)
        .sort(compareStudentsByClassAndAbsent),
    [students, className]
  );

  // Sync attendanceMap keys whenever className changes in new mode
  useEffect(() => {
    if (!className) {
      setAttendanceMap({});
      return;
    }
    setAttendanceMap((prev) => {
      const next: Record<string, AttendanceRecord> = {};
      studentsInSelectedClass.forEach((s) => {
        next[s.id] = prev[s.id] ?? { status: 'H', note: '' };
      });
      return next;
    });
  }, [className, studentsInSelectedClass]);

  const handleSetAllPresent = () => {
    const next: Record<string, AttendanceRecord> = {};
    studentsInSelectedClass.forEach((s) => {
      next[s.id] = { status: 'H', note: attendanceMap[s.id]?.note || '' };
    });
    setAttendanceMap(next);
    onToast('info', `Seluruh siswa kelas ${className} diatur Hadir (H).`);
  };

  const resetForm = () => {
    setEditingId(null);
    setDate(getTodayISO());
    setJamKe('1-2');
    setClassName('');
    setSubject('');
    setTopic('');
    setActivity('');
    setAttendanceMap({});
  };

  const handleSaveJournal = (e: React.FormEvent) => {
    e.preventDefault();
    if (!date) {
      onToast('error', 'Tanggal pertemuan wajib diisi.');
      return;
    }
    if (!className) {
      onToast('error', 'Silakan pilih Kelas terlebih dahulu.');
      return;
    }
    if (!subject) {
      onToast('error', 'Silakan pilih Mata Pelajaran terlebih dahulu.');
      return;
    }
    if (!topic.trim()) {
      onToast('error', 'Materi Pokok wajib diisi.');
      return;
    }

    const entry: JournalEntry = {
      id: editingId ?? `jrn-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      date,
      jamKe: jamKe.trim() || '-',
      className,
      subject,
      topic: topic.trim(),
      activity: activity.trim() || '-',
      academicYear: session.academicYear,
      semester: session.semester,
      attendance: attendanceMap,
      createdAt: new Date().toISOString(),
    };

    onUpsertJournal(entry);
    onToast(
      'success',
      editingId
        ? `Jurnal kelas ${className} (${formatShortDateId(date)}) berhasil diperbarui.`
        : `Jurnal & presensi kelas ${className} berhasil disimpan.`
    );
    resetForm();
    setFormModalOpen(false);
  };

  const handleStartEdit = (entry: JournalEntry) => {
    setEditingId(entry.id);
    setDate(entry.date);
    setJamKe(entry.jamKe);
    setClassName(entry.className);
    setSubject(entry.subject);
    setTopic(entry.topic);
    setActivity(entry.activity);
    setAttendanceMap(JSON.parse(JSON.stringify(entry.attendance)));
    setFormModalOpen(true);
  };

  const handleStartDuplicate = (entry: JournalEntry) => {
    setEditingId(null);
    setDate(getTodayISO());
    setJamKe(entry.jamKe);
    setClassName(entry.className);
    setSubject(entry.subject);
    setTopic(`${entry.topic} (Salinan)`);
    setActivity(entry.activity);
    setAttendanceMap(JSON.parse(JSON.stringify(entry.attendance)));
    setFormModalOpen(true);
    onToast('info', 'Draft jurnal disalin. Sesuaikan tanggal/materi lalu klik Simpan.');
  };

  // Active period journals
  const activeJournals = useMemo(
    () =>
      journals
        .filter(
          (j) =>
            j.academicYear === session.academicYear &&
            j.semester === session.semester
        )
        .sort((a, b) => (a.date < b.date ? 1 : -1)),
    [journals, session.academicYear, session.semester]
  );

  // Filtered History
  const filteredHistory = useMemo(() => {
    const q = historySearch.trim().toLowerCase();
    return activeJournals.filter((j) => {
      if (historyClassFilter && historyClassFilter !== 'ALL' && j.className !== historyClassFilter) {
        return false;
      }
      if (q) {
        const matchTopic = j.topic.toLowerCase().includes(q);
        const matchSubject = j.subject.toLowerCase().includes(q);
        const matchActivity = j.activity.toLowerCase().includes(q);
        if (!matchTopic && !matchSubject && !matchActivity) return false;
      }
      return true;
    });
  }, [activeJournals, historyClassFilter, historySearch]);

  const toggleSelectRow = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (
      filteredHistory.length > 0 &&
      filteredHistory.every((j) => selectedIds.has(j.id))
    ) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredHistory.map((j) => j.id)));
    }
  };

  const handleBulkDelete = () => {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    onConfirm({
      title: 'Hapus Jurnal Terpilih',
      description: `Anda yakin ingin menghapus ${count} riwayat jurnal beserta data presensinya? Tindakan ini tidak dapat dibatalkan.`,
      confirmLabel: `Hapus ${count} Jurnal`,
      danger: true,
      onConfirm: () => {
        onDeleteJournals(Array.from(selectedIds));
        setSelectedIds(new Set());
        onToast('success', `${count} jurnal berhasil dihapus.`);
      },
    });
  };

  const handleSingleDelete = (entry: JournalEntry) => {
    onConfirm({
      title: 'Hapus Jurnal Mengajar',
      description: `Hapus jurnal "${entry.topic}" kelas ${entry.className} (${formatShortDateId(entry.date)})?`,
      confirmLabel: 'Hapus Jurnal',
      danger: true,
      onConfirm: () => {
        onDeleteJournals([entry.id]);
        setSelectedIds((prev) => {
          const next = new Set(prev);
          next.delete(entry.id);
          return next;
        });
        onToast('success', 'Jurnal berhasil dihapus.');
      },
    });
  };

  const handleDownloadJournalTemplate = () => {
    exportRowsToXlsx(
      'Template Jurnal',
      ['Tanggal', 'Jam Ke', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'Kegiatan Pembelajaran'],
      [
        [
          getTodayISO(),
          '1-2',
          classOptions[0] || 'X-A',
          settings.masterSubjects[0] || 'Matematika',
          'Persamaan Linear Dua Variabel',
          'Diskusi kelompok dan latihan soal terbimbing',
        ],
      ],
      'Template_Impor_Jurnal_E-Guru.xlsx'
    );
    onToast('success', 'Template Excel (.xlsx) Jurnal Mengajar diunduh.');
  };

  const parseRawRowsToJournals = (
    rawRows: string[][]
  ): Omit<JournalEntry, 'id' | 'createdAt'>[] => {
    const res: Omit<JournalEntry, 'id' | 'createdAt'>[] = [];
    rawRows.forEach((cols, idx) => {
      if (cols.length < 4) return;
      const clean = (i: number) => String(cols[i] ?? '').trim();
      if (
        idx === 0 &&
        (clean(0).toLowerCase().includes('tanggal') ||
          clean(2).toLowerCase() === 'kelas' ||
          clean(4).toLowerCase().includes('materi'))
      ) {
        return;
      }
      const rawDate = clean(0) || getTodayISO();
      const rawJam = clean(1) || '1-2';
      const rawClass = clean(2);
      const rawSubj = clean(3) || settings.masterSubjects[0] || 'Umum';
      const rawTopic = clean(4);
      const rawActivity = clean(5) || '-';
      if (!rawClass || !rawTopic) return;

      const att: Record<string, AttendanceRecord> = {};
      students
        .filter((s) => s.className === rawClass)
        .forEach((s) => {
          att[s.id] = { status: 'H', note: '' };
        });

      res.push({
        date: rawDate,
        jamKe: rawJam,
        className: rawClass,
        subject: rawSubj,
        topic: rawTopic,
        activity: rawActivity,
        academicYear: session.academicYear,
        semester: session.semester,
        attendance: att,
      });
    });
    return res;
  };

  const handleImportPasteJournals = () => {
    const lines = pasteText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    if (lines.length === 0) {
      onToast('error', 'Tempelkan baris jurnal dari Excel/CSV terlebih dahulu.');
      return;
    }
    const matrix = lines.map((line) => {
      if (line.includes('\t')) return line.split('\t');
      if (line.includes(';')) return line.split(';');
      return line.split(',');
    });
    const parsed = parseRawRowsToJournals(matrix);
    if (parsed.length === 0) {
      onToast('error', 'Format tidak dikenali. Gunakan urutan: Tanggal, Jam Ke, Kelas, Mapel, Materi, Kegiatan.');
      return;
    }
    onConfirm({
      title: 'Simpan Data Impor Jurnal?',
      description: `Ditemukan ${parsed.length} entri jurnal mengajar siap disimpan. Apakah Anda ingin menyimpan data impor ini sekarang?`,
      confirmLabel: 'Ya, Simpan Jurnal',
      onConfirm: () => {
        const added = onImportJournalsBatch(parsed);
        setPasteText('');
        setShowImportBox(false);
        onToast('success', `${added} jurnal mengajar berhasil diimpor.`);
      },
      onCancel: () => {
        onToast('info', 'Impor jurnal mengajar dibatalkan, data belum disimpan.');
      },
    });
  };

  const handleImportJournalFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array' });
      const first = wb.SheetNames[0];
      if (!first) return;
      const rows = XLSX.utils.sheet_to_json<string[]>(wb.Sheets[first], {
        header: 1,
        defval: '',
      });
      const parsed = parseRawRowsToJournals(rows);
      if (parsed.length === 0) {
        onToast('error', 'Tidak ditemukan baris jurnal yang valid pada file.');
        return;
      }
      onConfirm({
        title: 'Simpan Data Impor Jurnal?',
        description: `File "${file.name}" berisi ${parsed.length} entri jurnal mengajar valid. Apakah Anda ingin menyimpan data impor ini?`,
        confirmLabel: 'Ya, Simpan Jurnal',
        onConfirm: () => {
          const added = onImportJournalsBatch(parsed);
          setShowImportBox(false);
          onToast('success', `Impor "${file.name}" selesai: ${added} jurnal ditambahkan.`);
        },
        onCancel: () => {
          onToast('info', `Impor file "${file.name}" dibatalkan, data belum disimpan.`);
        },
      });
    } catch {
      onToast('error', 'Gagal membaca file Excel/CSV.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const exportSourceHistory = useMemo(() => {
    if (selectedIds.size > 0) {
      const chosen = filteredHistory.filter((j) => selectedIds.has(j.id));
      if (chosen.length > 0) return chosen;
    }
    return filteredHistory;
  }, [filteredHistory, selectedIds]);

  const buildHistoryDataset = () => ({
    headers: ['No', 'Tanggal', 'Jam', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'Kegiatan', 'H', 'S', 'I', 'A', 'D'],
    rows: exportSourceHistory.map((j, idx) => {
      const s = summarizeJournalRow(j);
      return [
        idx + 1,
        formatShortDateId(j.date),
        j.jamKe,
        j.className,
        j.subject,
        j.topic,
        j.activity,
        s.counts.H,
        s.counts.S,
        s.counts.I,
        s.counts.A,
        s.counts.D,
      ];
    }),
  });

  // Section 3: Horizontal Ledger Data
  const ledgerJournals = useMemo(
    () =>
      ledgerClass
        ? activeJournals
            .filter(
              (j) =>
                j.className === ledgerClass &&
                (!ledgerSubject || ledgerSubject === 'ALL' || j.subject === ledgerSubject)
            )
            .sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : a.jamKe.localeCompare(b.jamKe)))
        : [],
    [activeJournals, ledgerClass, ledgerSubject]
  );

  const ledgerStudents = useMemo(
    () =>
      students
        .filter((s) => s.className === ledgerClass)
        .sort(compareStudentsByClassAndAbsent),
    [students, ledgerClass]
  );

  // Initialize / reset ledger edit buffer when underlying class/subject/journals change
  useEffect(() => {
    const initial: Record<string, Record<string, AttendanceStatus>> = {};
    ledgerJournals.forEach((j) => {
      initial[j.id] = {};
      ledgerStudents.forEach((s) => {
        initial[j.id][s.id] = j.attendance[s.id]?.status ?? 'H';
      });
    });
    setLedgerEdits(initial);
    setLedgerDirty(false);
  }, [ledgerJournals, ledgerStudents]);

  const handleLedgerCellChange = (journalId: string, studentId: string, rawVal: string) => {
    const upper = rawVal.trim().toUpperCase().slice(-1);
    if (!['H', 'S', 'I', 'A', 'D'].includes(upper)) return;
    setLedgerEdits((prev) => ({
      ...prev,
      [journalId]: {
        ...(prev[journalId] || {}),
        [studentId]: upper as AttendanceStatus,
      },
    }));
    setLedgerDirty(true);
  };

  const handleSaveLedgerChanges = () => {
    const updated: JournalEntry[] = ledgerJournals.map((j) => {
      const nextAtt: Record<string, AttendanceRecord> = { ...j.attendance };
      const editedForJournal = ledgerEdits[j.id] || {};
      ledgerStudents.forEach((s) => {
        const nextStatus = editedForJournal[s.id] ?? nextAtt[s.id]?.status ?? 'H';
        nextAtt[s.id] = {
          status: nextStatus,
          note: nextAtt[s.id]?.note ?? '',
        };
      });
      return { ...j, attendance: nextAtt };
    });
    onBatchUpdateJournals(updated);
    setLedgerDirty(false);
    onToast('success', `Perubahan leger kehadiran kelas ${ledgerClass} berhasil disimpan.`);
  };

  const formSummary = useMemo(() => {
    const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
    studentsInSelectedClass.forEach((s) => {
      const st = attendanceMap[s.id]?.status ?? 'H';
      counts[st] += 1;
    });
    return counts;
  }, [studentsInSelectedClass, attendanceMap]);

  return (
    <div className="space-y-6">
      {/* Compact Header Action Bar */}
      <section className="bg-white border border-slate-200 rounded-xl px-5 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h2 className="text-base font-bold text-slate-900">
            Jurnal Mengajar & Presensi Kelas
          </h2>
          <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
            {activeJournals.length} pertemuan tercatat pada TA {session.academicYear} · Semester {session.semester}
          </p>
        </div>
        <button
          type="button"
          onClick={() => {
            resetForm();
            setFormModalOpen(true);
          }}
          className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold transition-colors inline-flex items-center justify-center gap-2 whitespace-nowrap shadow-xs"
        >
          <Plus className="w-4 h-4" />
          <span>Form Jurnal & Presensi Pertemuan</span>
        </button>
      </section>

      {/* BAGIAN 1: RIWAYAT JURNAL TERSIMPAN */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div
          className={`px-5 py-4 space-y-3 ${
            historyExpanded ? 'border-b border-slate-200' : ''
          }`}
        >
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
            <button
              type="button"
              onClick={() => setHistoryExpanded((prev) => !prev)}
              aria-expanded={historyExpanded}
              className="flex items-center justify-between gap-3 text-left flex-1 min-w-0 group"
            >
              <div className="min-w-0">
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-base font-bold text-slate-900 group-hover:text-indigo-700 transition-colors">
                    Riwayat Jurnal Tersimpan
                  </h2>
                  <span className="text-xs text-slate-500 tabular-nums">
                    · {filteredHistory.length} jurnal
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-0.5">
                  Cari, filter per kelas, pilih untuk edit, duplikat, hapus, impor, ekspor, atau cetak
                </p>
              </div>
              <span className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 group-hover:bg-slate-100 text-xs font-semibold text-slate-700 whitespace-nowrap shrink-0 transition-colors">
                <span>{historyExpanded ? 'Ciutkan' : 'Tampilkan'}</span>
                <ChevronDown
                  className={`w-4 h-4 text-slate-500 transition-transform duration-150 ${
                    historyExpanded ? 'rotate-180' : ''
                  }`}
                />
              </span>
            </button>

            {historyExpanded && (
              <div className="flex flex-wrap items-end gap-2.5">
              <div>
                <label htmlFor="jrn-hist-class" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Kelas
                </label>
                <select
                  id="jrn-hist-class"
                  value={historyClassFilter}
                  onChange={(e) => setHistoryClassFilter(e.target.value)}
                  aria-label="Filter riwayat berdasarkan kelas"
                  className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                >
                  <option value="">--Pilih Kelas--</option>
                  <option value="ALL">Semua Kelas</option>
                  {classOptions.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label htmlFor="jrn-hist-search" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Cari Jurnal
                </label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    id="jrn-hist-search"
                    type="text"
                    value={historySearch}
                    onChange={(e) => setHistorySearch(e.target.value)}
                    placeholder="Cari materi / mapel..."
                    className="pl-9 pr-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 w-48 sm:w-56"
                  />
                </div>
              </div>

              <button
                type="button"
                onClick={() => setShowImportBox((prev) => !prev)}
                className="px-3 py-2 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold inline-flex items-center gap-1.5 whitespace-nowrap"
              >
                <Upload className="w-3.5 h-3.5" />
                {showImportBox ? 'Tutup Impor' : 'Impor Jurnal'}
              </button>
            </div>
            )}
          </div>

          {historyExpanded && showImportBox && (
            <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
              <p className="text-xs text-slate-600">
                Format kolom: <strong>Tanggal (YYYY-MM-DD), Jam Ke, Kelas, Mata Pelajaran, Materi Pokok, Kegiatan Pembelajaran</strong>. Kehadiran siswa otomatis di-set Hadir (H).
              </p>
              <textarea
                rows={3}
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
                placeholder="Tempel baris jurnal dari Excel/CSV di sini..."
                className="w-full px-3 py-2 text-xs font-mono-tabular border border-slate-200 bg-white rounded-lg focus:outline-none focus:border-indigo-600"
              />
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleImportJournalFile}
                    className="hidden"
                    id="journal-file-upload"
                  />
                  <label
                    htmlFor="journal-file-upload"
                    className="cursor-pointer px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5"
                  >
                    <Upload className="w-3.5 h-3.5 text-indigo-600" />
                    Unggah File Excel / CSV
                  </label>
                </div>
                <button
                  type="button"
                  onClick={handleImportPasteJournals}
                  className="px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold inline-flex items-center gap-1.5"
                >
                  <Plus className="w-3.5 h-3.5" />
                  Proses Impor Teks
                </button>
              </div>
            </div>
          )}

          {historyExpanded && (
          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedIds.size > 0
                ? `${selectedIds.size} jurnal dipilih`
                : `Menampilkan ${filteredHistory.length} jurnal`}
            </span>
            <TableExportToolbar
              selectedCount={selectedIds.size}
              onDeleteSelected={handleBulkDelete}
              onDownloadTemplate={handleDownloadJournalTemplate}
              onExportXlsx={() => {
                const { headers, rows } = buildHistoryDataset();
                exportRowsToXlsx('Riwayat Jurnal', headers, rows, `Riwayat_Jurnal_${getTodayISO()}.xlsx`);
                onToast('success', 'Riwayat jurnal diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildHistoryDataset();
                exportRowsToXls('Riwayat Jurnal', headers, rows, `Riwayat_Jurnal_${getTodayISO()}.xls`);
                onToast('success', 'Riwayat jurnal diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildHistoryDataset();
                exportDatasetToWordDoc({
                  db,
                  title: 'JURNAL KEGIATAN MENGAJAR GURU',
                  subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Riwayat_Jurnal_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Riwayat Jurnal diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildHistoryDataset();
                exportDatasetToPdf({
                  db,
                  title: 'JURNAL KEGIATAN MENGAJAR GURU',
                  subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Riwayat_Jurnal_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Riwayat Jurnal diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildHistoryDataset();
                printDatasetExact({
                  db,
                  title: 'JURNAL KEGIATAN MENGAJAR GURU',
                  subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('info', 'Menyiapkan tampilan cetak riwayat jurnal...');
              }}
            />
          </div>
          )}
        </div>

        {historyExpanded && (
        <div className="overflow-x-auto max-h-[420px] table-scroll">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="sticky top-0 z-10 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="px-4 py-3 w-10">
                  <input
                    type="checkbox"
                    checked={
                      filteredHistory.length > 0 &&
                      filteredHistory.every((j) => selectedIds.has(j.id))
                    }
                    onChange={toggleSelectAll}
                    aria-label="Pilih semua jurnal"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-4 py-3 whitespace-nowrap">Tanggal</th>
                <th className="px-4 py-3 whitespace-nowrap">Jam</th>
                <th className="px-4 py-3 whitespace-nowrap">Kelas</th>
                <th className="px-4 py-3 whitespace-nowrap">Mata Pelajaran</th>
                <th className="px-4 py-3 min-w-[240px]">Materi Pokok & Kegiatan</th>
                <th className="px-4 py-3 whitespace-nowrap text-right">Rekap Absen</th>
                <th className="px-4 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {filteredHistory.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-10 text-center text-sm text-slate-500">
                    Tidak ada riwayat jurnal yang cocok dengan filter saat ini.
                  </td>
                </tr>
              ) : (
                filteredHistory.map((j) => {
                  const sum = summarizeJournalRow(j);
                  const checked = selectedIds.has(j.id);
                  return (
                    <tr
                      key={j.id}
                      className={`hover:bg-slate-50/80 transition-colors ${
                        checked ? 'bg-indigo-50/40' : ''
                      }`}
                    >
                      <td className="px-4 py-3">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectRow(j.id)}
                          aria-label={`Pilih jurnal ${j.topic}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap tabular-nums font-medium text-slate-900">
                        {formatShortDateId(j.date)}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap tabular-nums text-slate-600">
                        {j.jamKe}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap font-semibold text-slate-900">
                        {j.className}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">{j.subject}</td>
                      <td className="px-4 py-3">
                        <p className="font-medium text-slate-900">{j.topic}</p>
                        <p className="text-xs text-slate-500 mt-0.5 line-clamp-1">{j.activity}</p>
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-right tabular-nums text-xs">
                        H:{sum.counts.H} · S:{sum.counts.S} · I:{sum.counts.I} · A:{sum.counts.A} · D:{sum.counts.D}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-right">
                        <div className="inline-flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => handleStartEdit(j)}
                            className="px-2.5 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 inline-flex items-center gap-1"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                            Edit
                          </button>
                          <button
                            type="button"
                            onClick={() => handleStartDuplicate(j)}
                            className="px-2.5 py-1.5 rounded-md border border-indigo-200 text-xs font-medium text-indigo-700 bg-indigo-50/60 hover:bg-indigo-100 inline-flex items-center gap-1"
                          >
                            <Copy className="w-3.5 h-3.5" />
                            Duplikat
                          </button>
                          <button
                            type="button"
                            onClick={() => handleSingleDelete(j)}
                            className="px-2.5 py-1.5 rounded-md border border-rose-200 text-xs font-medium text-rose-700 hover:bg-rose-50 inline-flex items-center gap-1"
                            title="Hapus jurnal ini"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            Hapus
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
        )}
      </section>

      {/* BAGIAN 3: LEGER KEHADIRAN PER KELAS (MENDATAR & EDITABLE INLINE) */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div
          className={`px-5 py-4 space-y-3 ${
            ledgerExpanded ? 'border-b border-slate-200' : ''
          }`}
        >
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
            <button
              type="button"
              onClick={() => setLedgerExpanded((prev) => !prev)}
              aria-expanded={ledgerExpanded}
              className="flex items-center justify-between gap-3 text-left flex-1 min-w-0 group"
            >
              <div className="min-w-0">
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-base font-bold text-slate-900 group-hover:text-indigo-700 transition-colors">
                    Leger Kehadiran per Kelas
                  </h2>
                  <span className="text-xs text-slate-500 tabular-nums">
                    · {ledgerClass || '-'} · {ledgerStudents.length} siswa · {ledgerJournals.length} pertemuan
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-0.5">
                  Ketik langsung huruf status (H / S / I / A / D) di sel tanggal lalu klik Simpan Perubahan
                </p>
              </div>
              <span className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 group-hover:bg-slate-100 text-xs font-semibold text-slate-700 whitespace-nowrap shrink-0 transition-colors">
                <span>{ledgerExpanded ? 'Ciutkan' : 'Tampilkan'}</span>
                <ChevronDown
                  className={`w-4 h-4 text-slate-500 transition-transform duration-150 ${
                    ledgerExpanded ? 'rotate-180' : ''
                  }`}
                />
              </span>
            </button>

            {ledgerExpanded && (
            <div className="flex flex-wrap items-end gap-2.5">
              <div>
                <label htmlFor="jrn-ledger-class" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Kelas
                </label>
                <select
                  id="jrn-ledger-class"
                  value={ledgerClass}
                  onChange={(e) => setLedgerClass(e.target.value)}
                  aria-label="Pilih kelas untuk leger kehadiran"
                  className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                >
                  <option value="">--Pilih Kelas--</option>
                  {classOptions.map((c) => (
                    <option key={c} value={c}>
                      Kelas {c}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label htmlFor="jrn-ledger-subject" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Mata Pelajaran
                </label>
                <select
                  id="jrn-ledger-subject"
                  value={ledgerSubject}
                  onChange={(e) => setLedgerSubject(e.target.value)}
                  aria-label="Filter mata pelajaran leger"
                  className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                >
                  <option value="">--Pilih Mapel--</option>
                  <option value="ALL">Semua Mata Pelajaran</option>
                  {subjectOptions.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>

              <button
                type="button"
                onClick={handleSaveLedgerChanges}
                disabled={!ledgerDirty || ledgerJournals.length === 0}
                className={`px-4 py-2 rounded-lg text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap ${
                  ledgerDirty && ledgerJournals.length > 0
                    ? 'bg-indigo-600 hover:bg-indigo-700 text-white'
                    : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                }`}
              >
                <Save className="w-4 h-4" />
                Simpan Perubahan
              </button>
            </div>
            )}
          </div>

          {ledgerExpanded && (
          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {ledgerStudents.length} siswa · {ledgerJournals.length} pertemuan KBM ({!ledgerSubject || ledgerSubject === 'ALL' ? 'Semua Mapel' : ledgerSubject})
            </span>
            <TableExportToolbar
              onExportXlsx={() => {
                const headers = ['No', 'NIS', 'NISN', 'Nama Siswa', ...ledgerJournals.map((j) => `${formatCompactDateId(j.date)} (${j.subject})`), 'H', 'S', 'I', 'A', 'D', '% Hadir'];
                const rows = ledgerStudents.map((st, idx) => {
                  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                  const vals = ledgerJournals.map((j) => {
                    const v = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                    counts[v] += 1;
                    return v;
                  });
                  const pct = ledgerJournals.length > 0 ? `${((counts.H / ledgerJournals.length) * 100).toFixed(0)}%` : '0%';
                  return [idx + 1, st.nis, st.nisn || '-', st.name, ...vals, counts.H, counts.S, counts.I, counts.A, counts.D, pct];
                });
                exportRowsToXlsx(`Leger ${ledgerClass}`, headers, rows, `Leger_Kehadiran_${ledgerClass}_${getTodayISO()}.xlsx`);
                onToast('success', `Leger kehadiran ${ledgerClass} diekspor ke XLSX.`);
              }}
              onExportXls={() => {
                const headers = ['No', 'NIS', 'NISN', 'Nama Siswa', ...ledgerJournals.map((j) => `${formatCompactDateId(j.date)} (${j.subject})`), 'H', 'S', 'I', 'A', 'D', '% Hadir'];
                const rows = ledgerStudents.map((st, idx) => {
                  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                  const vals = ledgerJournals.map((j) => {
                    const v = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                    counts[v] += 1;
                    return v;
                  });
                  const pct = ledgerJournals.length > 0 ? `${((counts.H / ledgerJournals.length) * 100).toFixed(0)}%` : '0%';
                  return [idx + 1, st.nis, st.nisn || '-', st.name, ...vals, counts.H, counts.S, counts.I, counts.A, counts.D, pct];
                });
                exportRowsToXls(`Leger ${ledgerClass}`, headers, rows, `Leger_Kehadiran_${ledgerClass}_${getTodayISO()}.xls`);
                onToast('success', `Leger kehadiran ${ledgerClass} diekspor ke XLS.`);
              }}
              onExportWord={(cfg) => {
                const headers = ['No', 'NIS', 'NISN', 'Nama Siswa', ...ledgerJournals.map((j) => formatCompactDateId(j.date)), 'H', 'S', 'I', 'A', 'D', '%'];
                const rows = ledgerStudents.map((st, idx) => {
                  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                  const vals = ledgerJournals.map((j) => {
                    const v = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                    counts[v] += 1;
                    return v;
                  });
                  const pct = ledgerJournals.length > 0 ? `${((counts.H / ledgerJournals.length) * 100).toFixed(0)}%` : '0%';
                  return [idx + 1, st.nis, st.nisn || '-', st.name, ...vals, counts.H, counts.S, counts.I, counts.A, counts.D, pct];
                });
                exportDatasetToWordDoc({
                  db,
                  title: `LEGER KEHADIRAN SISWA KELAS ${ledgerClass}`,
                  subtitle: `${!ledgerSubject || ledgerSubject === 'ALL' ? 'Semua Mata Pelajaran' : ledgerSubject} · TA ${session.academicYear}`,
                  headers,
                  rows,
                  filename: `Leger_Kehadiran_${ledgerClass}_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Leger Kehadiran diunduh.');
              }}
              onExportPdf={(cfg) => {
                const headers = ['No', 'NIS', 'NISN', 'Nama Siswa', ...ledgerJournals.map((j) => formatCompactDateId(j.date)), 'H', 'S', 'I', 'A', 'D', '%'];
                const rows = ledgerStudents.map((st, idx) => {
                  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                  const vals = ledgerJournals.map((j) => {
                    const v = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                    counts[v] += 1;
                    return v;
                  });
                  const pct = ledgerJournals.length > 0 ? `${((counts.H / ledgerJournals.length) * 100).toFixed(0)}%` : '0%';
                  return [idx + 1, st.nis, st.nisn || '-', st.name, ...vals, counts.H, counts.S, counts.I, counts.A, counts.D, pct];
                });
                exportDatasetToPdf({
                  db,
                  title: `LEGER KEHADIRAN SISWA KELAS ${ledgerClass}`,
                  subtitle: `${!ledgerSubject || ledgerSubject === 'ALL' ? 'Semua Mata Pelajaran' : ledgerSubject} · TA ${session.academicYear}`,
                  headers,
                  rows,
                  filename: `Leger_Kehadiran_${ledgerClass}_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Leger Kehadiran diunduh.');
              }}
              onPrint={(cfg) => {
                const headers = ['No', 'NIS', 'NISN', 'Nama Siswa', ...ledgerJournals.map((j) => formatCompactDateId(j.date)), 'H', 'S', 'I', 'A', 'D', '%'];
                const rows = ledgerStudents.map((st, idx) => {
                  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                  const vals = ledgerJournals.map((j) => {
                    const v = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                    counts[v] += 1;
                    return v;
                  });
                  const pct = ledgerJournals.length > 0 ? `${((counts.H / ledgerJournals.length) * 100).toFixed(0)}%` : '0%';
                  return [idx + 1, st.nis, st.nisn || '-', st.name, ...vals, counts.H, counts.S, counts.I, counts.A, counts.D, pct];
                });
                printDatasetExact({
                  db,
                  title: `LEGER KEHADIRAN SISWA KELAS ${ledgerClass}`,
                  subtitle: `${!ledgerSubject || ledgerSubject === 'ALL' ? 'Semua Mata Pelajaran' : ledgerSubject} · TA ${session.academicYear}`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('info', `Menyiapkan tampilan cetak leger kehadiran ${ledgerClass}...`);
              }}
            />
          </div>
          )}
        </div>

        {ledgerExpanded && (
        <div className="overflow-x-auto max-h-[420px] table-scroll">
          <table className="w-full text-left border-collapse text-xs sm:text-sm">
            <thead className="sticky top-0 z-20 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="sticky left-0 z-30 bg-slate-100 px-3 py-3 w-10 text-center border-r border-slate-200">
                  No
                </th>
                <th className="sticky left-10 z-30 bg-slate-100 px-3 py-3 min-w-[170px] border-r border-slate-200">
                  Nama Siswa
                </th>
                <th className="px-2.5 py-3 whitespace-nowrap">NIS</th>
                <th className="px-2.5 py-3 whitespace-nowrap">NISN</th>
                {ledgerJournals.map((j) => (
                  <th
                    key={j.id}
                    className="px-2 py-2.5 text-center whitespace-nowrap min-w-[68px]"
                    title={`${formatShortDateId(j.date)} · ${j.subject} · ${j.topic}`}
                  >
                    <div className="tabular-nums font-bold text-slate-900">
                      {formatCompactDateId(j.date)}
                    </div>
                    <div className="text-[10px] font-normal text-slate-500 truncate max-w-[64px] mx-auto">
                      {j.subject}
                    </div>
                  </th>
                ))}
                <th className="px-2.5 py-3 text-center bg-emerald-50 text-emerald-800 border-l border-slate-200">H</th>
                <th className="px-2.5 py-3 text-center bg-amber-50 text-amber-800">S</th>
                <th className="px-2.5 py-3 text-center bg-sky-50 text-sky-800">I</th>
                <th className="px-2.5 py-3 text-center bg-rose-50 text-rose-800">A</th>
                <th className="px-2.5 py-3 text-center bg-violet-50 text-violet-800">D</th>
                <th className="px-3 py-3 text-right whitespace-nowrap bg-slate-100">% Hadir</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {!ledgerClass ? (
                <tr>
                  <td
                    colSpan={10 + ledgerJournals.length}
                    className="px-4 py-10 text-center text-sm text-slate-500"
                  >
                    Silakan pilih kelas terlebih dahulu untuk melihat leger kehadiran.
                  </td>
                </tr>
              ) : ledgerStudents.length === 0 ? (
                <tr>
                  <td
                    colSpan={10 + ledgerJournals.length}
                    className="px-4 py-10 text-center text-sm text-slate-500"
                  >
                    Belum ada siswa di kelas {ledgerClass}.
                  </td>
                </tr>
              ) : ledgerJournals.length === 0 ? (
                <tr>
                  <td
                    colSpan={10}
                    className="px-4 py-10 text-center text-sm text-slate-500"
                  >
                    Belum ada pertemuan KBM tercatat untuk kelas {ledgerClass}.
                  </td>
                </tr>
              ) : (
                ledgerStudents.map((st, idx) => {
                  const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
                  ledgerJournals.forEach((j) => {
                    const val = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                    counts[val] += 1;
                  });
                  const pct = ledgerJournals.length > 0 ? (counts.H / ledgerJournals.length) * 100 : 0;

                  return (
                    <tr key={st.id} className="hover:bg-slate-50/70">
                      <td className="sticky left-0 z-10 bg-white px-3 py-2 text-center text-xs text-slate-500 tabular-nums border-r border-slate-200">
                        {idx + 1}
                      </td>
                      <td className="sticky left-10 z-10 bg-white px-3 py-2 font-medium text-slate-900 whitespace-nowrap border-r border-slate-200">
                        <button
                          type="button"
                          onClick={() => setDetailStudent(st)}
                          title="Klik untuk melihat detail informasi siswa"
                          className="text-left font-medium text-slate-900 hover:text-indigo-600 hover:underline cursor-pointer"
                        >
                          {st.name}
                        </button>
                      </td>
                      <td className="px-2.5 py-2 text-xs font-mono-tabular text-slate-600 whitespace-nowrap">
                        {st.nis}
                      </td>
                      <td className="px-2.5 py-2 text-xs font-mono-tabular text-slate-600 whitespace-nowrap">
                        {st.nisn || '-'}
                      </td>
                      {ledgerJournals.map((j) => {
                        const val = ledgerEdits[j.id]?.[st.id] ?? j.attendance[st.id]?.status ?? 'H';
                        const meta = ATTENDANCE_STATUS_META[val];
                        return (
                          <td key={j.id} className="px-1.5 py-1.5 text-center">
                            <input
                              type="text"
                              maxLength={2}
                              value={val}
                              onChange={(e) => handleLedgerCellChange(j.id, st.id, e.target.value)}
                              onFocus={(e) => e.target.select()}
                              aria-label={`Status ${st.name} tanggal ${j.date}`}
                              className={`w-9 h-8 text-center text-xs font-bold uppercase rounded border focus:outline-none focus:ring-2 focus:ring-indigo-600/30 ${meta.bg} ${meta.text} ${meta.border}`}
                            />
                          </td>
                        );
                      })}
                      <td className="px-2.5 py-2 text-center tabular-nums font-semibold text-emerald-700 border-l border-slate-200">
                        {counts.H}
                      </td>
                      <td className="px-2.5 py-2 text-center tabular-nums text-amber-700">{counts.S}</td>
                      <td className="px-2.5 py-2 text-center tabular-nums text-sky-700">{counts.I}</td>
                      <td className="px-2.5 py-2 text-center tabular-nums text-rose-700">{counts.A}</td>
                      <td className="px-2.5 py-2 text-center tabular-nums text-violet-700">{counts.D}</td>
                      <td className="px-3 py-2 text-right tabular-nums font-bold text-slate-900 whitespace-nowrap">
                        {pct.toFixed(0)}%
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
        )}
      </section>

      <FormModal
        open={formModalOpen}
        title={editingId ? 'Edit Jurnal & Presensi Pertemuan' : 'Form Jurnal & Presensi Pertemuan'}
        subtitle="Catat materi pembelajaran dan status kehadiran siswa dalam satu langkah"
        maxWidthClass="max-w-5xl"
        onClose={() => {
          setFormModalOpen(false);
          resetForm();
        }}
      >
        <form onSubmit={handleSaveJournal} className="p-5 space-y-5">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label htmlFor="j-date" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Tanggal Pertemuan
              </label>
              <input
                id="j-date"
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="j-jam" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Jam Ke-
              </label>
              <input
                id="j-jam"
                type="text"
                value={jamKe}
                onChange={(e) => setJamKe(e.target.value)}
                placeholder="Contoh: 1-2"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="j-class" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Kelas
              </label>
              <select
                id="j-class"
                value={className}
                onChange={(e) => setClassName(e.target.value)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
              >
                <option value="">--Pilih Kelas--</option>
                {classOptions.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label htmlFor="j-subject" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Mata Pelajaran <span className="text-rose-600">*</span>
              </label>
              <select
                id="j-subject"
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
              >
                <option value="">--Pilih Mapel--</option>
                {subjectOptions.map((subj) => (
                  <option key={subj} value={subj}>
                    {subj}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div>
              <label htmlFor="j-topic" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Materi Pokok <span className="text-rose-600">*</span>
              </label>
              <input
                id="j-topic"
                type="text"
                value={topic}
                onChange={(e) => setTopic(e.target.value)}
                placeholder="Tuliskan topik / materi pokok pertemuan"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="j-activity" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Kegiatan Pembelajaran
              </label>
              <textarea
                id="j-activity"
                rows={2}
                value={activity}
                onChange={(e) => setActivity(e.target.value)}
                placeholder="Ringkasan metode, aktivitas diskusi, atau tugas yang diberikan"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
              />
            </div>
          </div>

          {className && (
            <div className="border border-slate-200 rounded-xl overflow-hidden">
              <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
                <div className="text-xs text-slate-700">
                  <span className="font-bold text-slate-900">Daftar Presensi Kelas {className}</span>
                  <span className="mx-2 text-slate-400">·</span>
                  <span className="tabular-nums">
                    {studentsInSelectedClass.length} siswa (H:{formSummary.H} · S:{formSummary.S} · I:{formSummary.I} · A:{formSummary.A} · D:{formSummary.D})
                  </span>
                </div>
                <button
                  type="button"
                  onClick={handleSetAllPresent}
                  className="px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
                >
                  <CheckCheck className="w-3.5 h-3.5" />
                  Set Semua Hadir (H)
                </button>
              </div>

              {studentsInSelectedClass.length === 0 ? (
                <div className="p-6 text-center text-sm text-slate-500">
                  Belum ada siswa terdaftar di kelas {className}. Tambahkan siswa pada menu Data Siswa terlebih dahulu.
                </div>
              ) : (
                <div className="overflow-x-auto max-h-[320px] table-scroll">
                  <table className="w-full text-left border-collapse text-sm">
                    <thead className="sticky top-0 z-10 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
                      <tr>
                        <th className="px-3 py-2.5 w-10 text-center">No</th>
                        <th className="px-3 py-2.5 whitespace-nowrap">NIS</th>
                        <th className="px-3 py-2.5 whitespace-nowrap">NISN</th>
                        <th className="px-3 py-2.5 min-w-[180px]">Nama Siswa</th>
                        <th className="px-3 py-2.5 whitespace-nowrap">Status Kehadiran</th>
                        <th className="px-3 py-2.5 min-w-[200px]">Keterangan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {studentsInSelectedClass.map((st, idx) => {
                        const rec = attendanceMap[st.id] ?? { status: 'H', note: '' };
                        return (
                          <tr key={st.id} className="hover:bg-slate-50/70">
                            <td className="px-3 py-2 text-center text-xs text-slate-500 tabular-nums">
                              {idx + 1}
                            </td>
                            <td className="px-3 py-2 text-xs font-mono-tabular text-slate-600 whitespace-nowrap">
                              {st.nis}
                            </td>
                            <td className="px-3 py-2 text-xs font-mono-tabular text-slate-600 whitespace-nowrap">
                              {st.nisn || '-'}
                            </td>
                            <td className="px-3 py-2 font-medium text-slate-900 whitespace-nowrap">
                              <button
                                type="button"
                                onClick={() => setDetailStudent(st)}
                                title="Klik untuk melihat detail informasi siswa"
                                className="text-left font-medium text-slate-900 hover:text-indigo-600 hover:underline cursor-pointer"
                              >
                                {st.name}
                              </button>
                            </td>
                            <td className="px-3 py-2 whitespace-nowrap">
                              <div className="inline-flex items-center gap-1.5" role="radiogroup" aria-label={`Status ${st.name}`}>
                                {STATUS_KEYS.map((key) => {
                                  const meta = ATTENDANCE_STATUS_META[key];
                                  const checked = rec.status === key;
                                  return (
                                    <label
                                      key={key}
                                      className={`cursor-pointer px-2.5 py-1 rounded-md border text-xs font-semibold transition-colors select-none ${
                                        checked
                                          ? meta.activeBg
                                          : `${meta.bg} ${meta.text} ${meta.border} hover:opacity-90`
                                      }`}
                                      title={meta.label}
                                    >
                                      <input
                                        type="radio"
                                        name={`att-${st.id}`}
                                        value={key}
                                        checked={checked}
                                        onChange={() =>
                                          setAttendanceMap((prev) => ({
                                            ...prev,
                                            [st.id]: { ...rec, status: key },
                                          }))
                                        }
                                        className="sr-only"
                                      />
                                      {key}
                                    </label>
                                  );
                                })}
                              </div>
                            </td>
                            <td className="px-3 py-2">
                              <input
                                type="text"
                                value={rec.note}
                                onChange={(e) =>
                                  setAttendanceMap((prev) => ({
                                    ...prev,
                                    [st.id]: { ...rec, note: e.target.value },
                                  }))
                                }
                                placeholder="Opsional: alasan izin/sakit..."
                                className="w-full px-2.5 py-1.5 text-xs border border-slate-200 rounded-md focus:outline-none focus:border-indigo-600"
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          <div className="flex flex-wrap items-center justify-end gap-2.5 pt-2 border-t border-slate-100">
            <button
              type="button"
              onClick={() => {
                resetForm();
                setFormModalOpen(false);
              }}
              className="px-4 py-2 rounded-lg border border-slate-200 text-sm font-medium text-slate-700 hover:bg-slate-50 transition-colors"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={resetForm}
              className="px-4 py-2 rounded-lg border border-slate-200 text-sm font-medium text-slate-700 hover:bg-slate-50 transition-colors"
            >
              Reset Form
            </button>
            <button
              type="submit"
              className="px-5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
            >
              <Save className="w-4 h-4" />
              {editingId ? 'Simpan Perubahan Jurnal' : 'Simpan Jurnal & Absensi'}
            </button>
          </div>
        </form>
      </FormModal>

      <StudentDetailModal
        student={detailStudent}
        db={db}
        onClose={() => setDetailStudent(null)}
      />
    </div>
  );
}
