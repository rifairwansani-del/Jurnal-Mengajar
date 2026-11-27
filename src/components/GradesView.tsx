import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { ChevronDown, Plus, Save, Sparkles, Trash2, Upload } from 'lucide-react';
import { GradeEntry, GradeRowMarker, PortalDatabase, Student } from '../types';
import {
  exportDatasetToPdf,
  exportDatasetToWordDoc,
  exportRowsToXls,
  exportRowsToXlsx,
  printDatasetExact,
} from '../utils/exportHelpers';
import {
  compareClassNames,
  compareStudentsByClassAndAbsent,
  getTodayISO,
  makeGradeId,
} from '../utils/portalHelpers';
import { ConfirmDialogState } from './FeedbackUI';
import { StudentDetailModal } from './StudentDetailModal';
import { TableExportToolbar } from './TableExportToolbar';

interface GradesViewProps {
  db: PortalDatabase;
  onUpsertGradesBatch: (entries: GradeEntry[]) => void;
  onDeleteGradesForStudents: (studentIds: string[], className: string, subject: string) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const MARKER_CONFIG: {
  id: GradeRowMarker;
  label: string;
  dotClass: string;
  rowBgClass: string;
}[] = [
  { id: 'neutral', label: 'Netral', dotClass: 'bg-slate-400', rowBgClass: 'bg-white' },
  { id: 'green', label: 'Tuntas', dotClass: 'bg-emerald-500', rowBgClass: 'bg-emerald-50/70' },
  { id: 'yellow', label: 'Remedial', dotClass: 'bg-amber-400', rowBgClass: 'bg-amber-50/70' },
  { id: 'red', label: 'Perhatian', dotClass: 'bg-rose-500', rowBgClass: 'bg-rose-50/70' },
  { id: 'blue', label: 'Pengayaan', dotClass: 'bg-sky-500', rowBgClass: 'bg-sky-50/70' },
];

export function GradesView({
  db,
  onUpsertGradesBatch,
  onDeleteGradesForStudents,
  onToast,
  onConfirm,
}: GradesViewProps) {
  const { students, grades, settings, session } = db;

  const classOptions = useMemo(() => {
    const set = new Set(settings.masterClasses);
    students.forEach((s) => {
      if (s.className) set.add(s.className);
    });
    return Array.from(set).sort(compareClassNames);
  }, [settings.masterClasses, students]);

  // Section 1 state
  const [selectedClass, setSelectedClass] = useState<string>('');
  const [selectedSubject, setSelectedSubject] = useState<string>('');
  const [selectedExamType, setSelectedExamType] = useState<string>('');
  const [selectedExamStudentIds, setSelectedExamStudentIds] = useState<Set<string>>(new Set());
  const [showImportBox, setShowImportBox] = useState(false);
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [draftScores, setDraftScores] = useState<
    Record<string, { score: string; note: string; marker: GradeRowMarker }>
  >({});

  // Section 2 state (Matrix)
  const [matrixClass, setMatrixClass] = useState<string>('');
  const [matrixSubject, setMatrixSubject] = useState<string>('');
  const [matrixDraft, setMatrixDraft] = useState<Record<string, string>>({});
  const [matrixDirty, setMatrixDirty] = useState(false);
  const [selectedMatrixIds, setSelectedMatrixIds] = useState<Set<string>>(new Set());
  const [detailStudent, setDetailStudent] = useState<Student | null>(null);
  const [matrixExpanded, setMatrixExpanded] = useState(false);

  const hasExamSelection = Boolean(selectedClass && selectedSubject && selectedExamType);
  const hasMatrixSelection = Boolean(matrixClass && matrixSubject);

  const studentsInClass = useMemo(
    () =>
      selectedClass
        ? students
            .filter((s) => s.className === selectedClass)
            .sort(compareStudentsByClassAndAbsent)
        : [],
    [students, selectedClass]
  );

  // Hydrate Section 1 draft from stored grades
  useEffect(() => {
    const next: Record<string, { score: string; note: string; marker: GradeRowMarker }> = {};
    studentsInClass.forEach((st) => {
      const id = makeGradeId(
        session.academicYear,
        session.semester,
        selectedClass,
        selectedSubject,
        selectedExamType,
        st.id
      );
      const found = grades.find((g) => g.id === id);
      next[st.id] = {
        score: found && found.score !== null ? String(found.score) : '',
        note: found?.note ?? '',
        marker: found?.marker ?? 'neutral',
      };
    });
    setDraftScores(next);
  }, [
    studentsInClass,
    grades,
    session.academicYear,
    session.semester,
    selectedClass,
    selectedSubject,
    selectedExamType,
  ]);

  const liveAverage = useMemo(() => {
    const vals: number[] = [];
    studentsInClass.forEach((st) => {
      const raw = draftScores[st.id]?.score;
      if (raw !== undefined && raw.trim() !== '') {
        const num = Number(raw);
        if (!Number.isNaN(num)) vals.push(num);
      }
    });
    if (vals.length === 0) return null;
    return vals.reduce((a, b) => a + b, 0) / vals.length;
  }, [studentsInClass, draftScores]);

  const handleScoreChange = (studentId: string, val: string) => {
    if (val === '') {
      setDraftScores((prev) => ({
        ...prev,
        [studentId]: { ...(prev[studentId] || { note: '', marker: 'neutral' }), score: '' },
      }));
      return;
    }
    const num = Number(val);
    if (Number.isNaN(num)) return;
    const clamped = Math.max(0, Math.min(100, num));
    setDraftScores((prev) => ({
      ...prev,
      [studentId]: {
        ...(prev[studentId] || { note: '', marker: 'neutral' }),
        score: String(clamped),
      },
    }));
  };

  const handleAutoKkm75 = () => {
    let adjusted = 0;
    setDraftScores((prev) => {
      const next = { ...prev };
      studentsInClass.forEach((st) => {
        const current = next[st.id];
        if (!current) return;
        const num = current.score.trim() === '' ? null : Number(current.score);
        if (num !== null && !Number.isNaN(num) && num < 75) {
          next[st.id] = {
            ...current,
            score: '75',
            marker: current.marker === 'red' || current.marker === 'yellow' ? 'green' : current.marker,
          };
          adjusted += 1;
        }
      });
      return next;
    });
    if (adjusted > 0) {
      onToast('info', `${adjusted} nilai di bawah KKM 75 otomatis disesuaikan menjadi 75.`);
    } else {
      onToast('info', 'Tidak ada nilai angka di bawah 75 pada daftar ini.');
    }
  };

  const handleSaveExamGrades = () => {
    if (!selectedClass || !selectedSubject || !selectedExamType) {
      onToast('error', 'Pilih Kelas, Mata Pelajaran, dan Jenis Ujian terlebih dahulu.');
      return;
    }

    const now = new Date().toISOString();
    const payload: GradeEntry[] = studentsInClass.map((st) => {
      const row = draftScores[st.id] ?? { score: '', note: '', marker: 'neutral' };
      const parsed = row.score.trim() === '' ? null : Math.max(0, Math.min(100, Number(row.score)));
      return {
        id: makeGradeId(
          session.academicYear,
          session.semester,
          selectedClass,
          selectedSubject,
          selectedExamType,
          st.id
        ),
        className: selectedClass,
        subject: selectedSubject,
        examType: selectedExamType,
        studentId: st.id,
        score: parsed !== null && !Number.isNaN(parsed) ? parsed : null,
        note: row.note.trim(),
        marker: row.marker,
        academicYear: session.academicYear,
        semester: session.semester,
        updatedAt: now,
      };
    });

    onUpsertGradesBatch(payload);
    onToast(
      'success',
      `Nilai ${selectedExamType} (${selectedSubject} · ${selectedClass}) berhasil disimpan.`
    );
  };

  // Section 2: Matrix Students & Values
  const matrixStudents = useMemo(
    () =>
      matrixClass
        ? students
            .filter((s) => s.className === matrixClass)
            .sort(compareStudentsByClassAndAbsent)
        : [],
    [students, matrixClass]
  );

  useEffect(() => {
    const initial: Record<string, string> = {};
    matrixStudents.forEach((st) => {
      settings.masterExamTypes.forEach((ex) => {
        const id = makeGradeId(
          session.academicYear,
          session.semester,
          matrixClass,
          matrixSubject,
          ex,
          st.id
        );
        const found = grades.find((g) => g.id === id);
        initial[`${st.id}__${ex}`] = found && found.score !== null ? String(found.score) : '';
      });
    });
    setMatrixDraft(initial);
    setMatrixDirty(false);
  }, [
    matrixStudents,
    settings.masterExamTypes,
    grades,
    session.academicYear,
    session.semester,
    matrixClass,
    matrixSubject,
  ]);

  const handleMatrixCellChange = (studentId: string, examType: string, rawVal: string) => {
    const key = `${studentId}__${examType}`;
    if (rawVal === '') {
      setMatrixDraft((prev) => ({ ...prev, [key]: '' }));
      setMatrixDirty(true);
      return;
    }
    const num = Number(rawVal);
    if (Number.isNaN(num)) return;
    const clamped = Math.max(0, Math.min(100, num));
    setMatrixDraft((prev) => ({ ...prev, [key]: String(clamped) }));
    setMatrixDirty(true);
  };

  const handleSaveMatrix = () => {
    if (!hasMatrixSelection) {
      onToast('error', 'Pilih Kelas dan Mata Pelajaran terlebih dahulu.');
      return;
    }
    const now = new Date().toISOString();
    const payload: GradeEntry[] = [];
    matrixStudents.forEach((st) => {
      settings.masterExamTypes.forEach((ex) => {
        const id = makeGradeId(
          session.academicYear,
          session.semester,
          matrixClass,
          matrixSubject,
          ex,
          st.id
        );
        const existing = grades.find((g) => g.id === id);
        const raw = matrixDraft[`${st.id}__${ex}`] ?? '';
        const parsed = raw.trim() === '' ? null : Math.max(0, Math.min(100, Number(raw)));
        payload.push({
          id,
          className: matrixClass,
          subject: matrixSubject,
          examType: ex,
          studentId: st.id,
          score: parsed !== null && !Number.isNaN(parsed) ? parsed : null,
          note: existing?.note ?? '',
          marker: existing?.marker ?? 'neutral',
          academicYear: session.academicYear,
          semester: session.semester,
          updatedAt: now,
        });
      });
    });
    onUpsertGradesBatch(payload);
    setMatrixDirty(false);
    onToast('success', `Matriks nilai kelas ${matrixClass} (${matrixSubject}) berhasil diperbarui.`);
  };

  // Section 1 import / template / selection / export
  const handleDownloadGradeTemplate = () => {
    const rows = studentsInClass.map((st) => [
      st.nis,
      st.name,
      draftScores[st.id]?.score ?? '',
      draftScores[st.id]?.note ?? '',
    ]);
    exportRowsToXlsx(
      `Template ${selectedExamType}`,
      ['NIS', 'Nama Siswa', 'Nilai (0-100)', 'Catatan'],
      rows.length > 0 ? rows : [['25001', 'Contoh Siswa', 85, 'Baik']],
      `Template_Nilai_${selectedClass}_${selectedExamType}.xlsx`
    );
    onToast('success', 'Template Excel (.xlsx) Nilai berhasil diunduh.');
  };

  const applyImportedGradeRows = (rawRows: string[][], sourceLabel?: string) => {
    const now = new Date().toISOString();
    const updates: GradeEntry[] = [];
    rawRows.forEach((cols, idx) => {
      if (cols.length < 2) return;
      const c0 = String(cols[0] ?? '').trim();
      const c1 = String(cols[1] ?? '').trim();
      const c2 = String(cols[2] ?? '').trim();
      const c3 = String(cols[3] ?? '').trim();
      if (idx === 0 && (c0.toLowerCase() === 'nis' || c1.toLowerCase().includes('nama'))) {
        return;
      }
      const student =
        studentsInClass.find((s) => s.nis.toLowerCase() === c0.toLowerCase()) ||
        studentsInClass.find((s) => s.name.toLowerCase() === c0.toLowerCase() || s.name.toLowerCase() === c1.toLowerCase());
      if (!student) return;

      // Support either 2/3 cols (NIS | Nilai | Catatan) or 4 cols (NIS | Nama | Nilai | Catatan)
      const rawScoreStr = cols.length >= 3 && Number.isNaN(Number(c1)) ? c2 : c1;
      const rawNoteStr = cols.length >= 4 ? c3 : cols.length === 3 && Number.isNaN(Number(c1)) ? '' : c2;
      const num = rawScoreStr === '' ? null : Math.max(0, Math.min(100, Number(rawScoreStr)));
      if (num !== null && Number.isNaN(num)) return;

      const marker: GradeRowMarker =
        num === null ? 'neutral' : num >= 75 ? 'green' : 'yellow';

      updates.push({
        id: makeGradeId(
          session.academicYear,
          session.semester,
          selectedClass,
          selectedSubject,
          selectedExamType,
          student.id
        ),
        className: selectedClass,
        subject: selectedSubject,
        examType: selectedExamType,
        studentId: student.id,
        score: num,
        note: rawNoteStr,
        marker,
        academicYear: session.academicYear,
        semester: session.semester,
        updatedAt: now,
      });
    });

    if (updates.length === 0) {
      onToast('error', 'Tidak ada baris nilai yang cocok dengan NIS/Nama siswa di kelas ini.');
      return;
    }
    onConfirm({
      title: 'Simpan Data Impor Nilai?',
      description: `Ditemukan ${updates.length} nilai siswa${sourceLabel ? ` dari ${sourceLabel}` : ''} untuk ${selectedExamType} (${selectedSubject} · Kelas ${selectedClass}). Apakah Anda ingin menyimpan nilai impor ini sekarang?`,
      confirmLabel: 'Ya, Simpan Nilai',
      onConfirm: () => {
        onUpsertGradesBatch(updates);
        setShowImportBox(false);
        setPasteText('');
        onToast('success', `${updates.length} nilai siswa berhasil diimpor & disimpan.`);
      },
      onCancel: () => {
        onToast('info', 'Impor nilai siswa dibatalkan, data belum disimpan.');
      },
    });
  };

  const handleImportPasteGrades = () => {
    const lines = pasteText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    if (lines.length === 0) {
      onToast('error', 'Tempelkan baris nilai dari Excel/CSV terlebih dahulu.');
      return;
    }
    const matrix = lines.map((line) => {
      if (line.includes('\t')) return line.split('\t');
      if (line.includes(';')) return line.split(';');
      return line.split(',');
    });
    applyImportedGradeRows(matrix);
  };

  const handleImportGradeFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
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
      applyImportedGradeRows(rows, `file "${file.name}"`);
    } catch {
      onToast('error', 'Gagal membaca file Excel/CSV.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const toggleSelectExamStudent = (id: string) => {
    setSelectedExamStudentIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAllExamStudents = () => {
    if (
      studentsInClass.length > 0 &&
      studentsInClass.every((s) => selectedExamStudentIds.has(s.id))
    ) {
      setSelectedExamStudentIds(new Set());
    } else {
      setSelectedExamStudentIds(new Set(studentsInClass.map((s) => s.id)));
    }
  };

  const handleClearExamGrades = (studentIds: string[], label: string) => {
    if (studentIds.length === 0) return;
    onConfirm({
      title: 'Kosongkan / Hapus Nilai Ujian',
      description: `Hapus nilai ${selectedExamType} (${selectedSubject}) untuk ${label}?`,
      confirmLabel: 'Hapus Nilai',
      danger: true,
      onConfirm: () => {
        const now = new Date().toISOString();
        const payload: GradeEntry[] = studentIds.map((studentId) => ({
          id: makeGradeId(
            session.academicYear,
            session.semester,
            selectedClass,
            selectedSubject,
            selectedExamType,
            studentId
          ),
          className: selectedClass,
          subject: selectedSubject,
          examType: selectedExamType,
          studentId,
          score: null,
          note: '',
          marker: 'neutral',
          academicYear: session.academicYear,
          semester: session.semester,
          updatedAt: now,
        }));
        onUpsertGradesBatch(payload);
        setSelectedExamStudentIds((prev) => {
          const next = new Set(prev);
          studentIds.forEach((id) => next.delete(id));
          return next;
        });
        onToast('success', `Nilai ${selectedExamType} untuk ${label} berhasil dikosongkan.`);
      },
    });
  };

  const exportExamStudents = useMemo(() => {
    if (selectedExamStudentIds.size > 0) {
      const chosen = studentsInClass.filter((s) => selectedExamStudentIds.has(s.id));
      if (chosen.length > 0) return chosen;
    }
    return studentsInClass;
  }, [studentsInClass, selectedExamStudentIds]);

  const buildExamDataset = () => ({
    headers: ['No', 'NIS', 'Nama Siswa', `Nilai ${selectedExamType}`, 'Status KKM (75)', 'Catatan'],
    rows: exportExamStudents.map((st, idx) => {
      const row = draftScores[st.id] ?? { score: '', note: '', marker: 'neutral' };
      const numScore = row.score.trim() === '' ? null : Number(row.score);
      const statusText =
        numScore === null ? 'Belum diisi' : numScore >= 75 ? 'Tuntas' : 'Remedial';
      return [
        idx + 1,
        st.nis,
        st.name,
        numScore !== null ? numScore : '-',
        statusText,
        row.note || '-',
      ];
    }),
  });

  // Section 2 selection & export
  const toggleSelectMatrix = (id: string) => {
    setSelectedMatrixIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAllMatrix = () => {
    if (
      matrixStudents.length > 0 &&
      matrixStudents.every((s) => selectedMatrixIds.has(s.id))
    ) {
      setSelectedMatrixIds(new Set());
    } else {
      setSelectedMatrixIds(new Set(matrixStudents.map((s) => s.id)));
    }
  };

  const handleDeleteMatrixSelected = () => {
    if (selectedMatrixIds.size === 0) return;
    const count = selectedMatrixIds.size;
    onConfirm({
      title: 'Hapus Semua Nilai Siswa Terpilih',
      description: `Hapus seluruh nilai mata pelajaran ${matrixSubject} untuk ${count} siswa terpilih di kelas ${matrixClass}?`,
      confirmLabel: `Hapus Nilai (${count})`,
      danger: true,
      onConfirm: () => {
        onDeleteGradesForStudents(Array.from(selectedMatrixIds), matrixClass, matrixSubject);
        setSelectedMatrixIds(new Set());
        onToast('success', `Nilai ${matrixSubject} untuk ${count} siswa berhasil dihapus.`);
      },
    });
  };

  const exportMatrixStudents = useMemo(() => {
    if (selectedMatrixIds.size > 0) {
      const chosen = matrixStudents.filter((s) => selectedMatrixIds.has(s.id));
      if (chosen.length > 0) return chosen;
    }
    return matrixStudents;
  }, [matrixStudents, selectedMatrixIds]);

  const buildMatrixDataset = () => ({
    headers: ['No', 'NIS', 'Nama Siswa', ...settings.masterExamTypes, 'Rata-rata', 'Status'],
    rows: exportMatrixStudents.map((st, idx) => {
      const numericVals: number[] = [];
      const examCells = settings.masterExamTypes.map((ex) => {
        const raw = matrixDraft[`${st.id}__${ex}`] ?? '';
        if (raw.trim() !== '') {
          const num = Number(raw);
          if (!Number.isNaN(num)) {
            numericVals.push(num);
            return num;
          }
        }
        return '-';
      });
      const avg =
        numericVals.length > 0
          ? numericVals.reduce((a, b) => a + b, 0) / numericVals.length
          : null;
      const statusLabel =
        avg === null ? 'Belum ada nilai' : avg >= 75 ? 'Tuntas' : 'Remedial';
      return [
        idx + 1,
        st.nis,
        st.name,
        ...examCells,
        avg !== null ? Number(avg.toFixed(1)) : '-',
        statusLabel,
      ];
    }),
  });

  return (
    <div className="space-y-8">
      {/* BAGIAN 1: INPUT NILAI PER JENIS UJIAN */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex flex-col xl:flex-row xl:items-center xl:justify-between gap-4">
          <div>
            <h2 className="text-base font-bold text-slate-900">1. Input Nilai per Jenis Ujian</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Masukkan nilai 0–100, catatan siswa, serta penanda warna status belajar
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            <select
              value={selectedClass}
              onChange={(e) => setSelectedClass(e.target.value)}
              aria-label="Pilih kelas"
              className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              <option value="">--Pilih Kelas--</option>
              {classOptions.map((c) => (
                <option key={c} value={c}>
                  Kelas {c}
                </option>
              ))}
            </select>

            <select
              value={selectedSubject}
              onChange={(e) => setSelectedSubject(e.target.value)}
              aria-label="Pilih mata pelajaran"
              className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              <option value="">--Pilih Mapel--</option>
              {settings.masterSubjects.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>

            <select
              value={selectedExamType}
              onChange={(e) => setSelectedExamType(e.target.value)}
              aria-label="Pilih jenis ujian"
              className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
            >
              <option value="">--Pilih Ujian--</option>
              {settings.masterExamTypes.map((ex) => (
                <option key={ex} value={ex}>
                  {ex}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Live toolbar */}
        <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-4 text-xs sm:text-sm">
              <span className="text-slate-600">
                Rata-rata Live:{' '}
                <strong className="text-indigo-700 font-mono-tabular text-base">
                  {liveAverage !== null ? liveAverage.toFixed(1) : '-'}
                </strong>
              </span>
              <span className="text-slate-400 hidden sm:inline">·</span>
              <span className="text-xs text-slate-500">KKM Sekolah: 75</span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  if (!hasExamSelection) {
                    onToast('error', 'Pilih Kelas, Mata Pelajaran, dan Jenis Ujian terlebih dahulu.');
                    return;
                  }
                  setShowImportBox((prev) => !prev);
                }}
                disabled={!hasExamSelection}
                className="px-3 py-2 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap disabled:opacity-50"
              >
                <Upload className="w-3.5 h-3.5" />
                {showImportBox ? 'Tutup Impor' : 'Impor Nilai'}
              </button>
              <button
                type="button"
                onClick={handleAutoKkm75}
                disabled={!hasExamSelection || studentsInClass.length === 0}
                className="px-3 py-2 rounded-lg border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap disabled:opacity-50"
              >
                <Sparkles className="w-3.5 h-3.5" />
                Auto KKM 75
              </button>
              <button
                type="button"
                onClick={handleSaveExamGrades}
                disabled={!hasExamSelection || studentsInClass.length === 0}
                className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap disabled:opacity-50"
              >
                <Save className="w-4 h-4" />
                Simpan Nilai
              </button>
            </div>
          </div>

          {showImportBox && (
            <div className="p-3.5 rounded-xl bg-white border border-slate-200 space-y-2.5">
              <p className="text-xs text-slate-600">
                Tempel data dari Excel dengan format <strong>NIS, Nama Siswa (opsional), Nilai (0–100), Catatan</strong> atau unggah file Excel/CSV.
              </p>
              <textarea
                rows={3}
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
                placeholder="Contoh: 25001&#9;88&#9;Sangat baik"
                className="w-full px-3 py-2 text-xs font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={handleImportGradeFile}
                    className="hidden"
                    id="grade-file-upload"
                  />
                  <label
                    htmlFor="grade-file-upload"
                    className="cursor-pointer px-3 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5"
                  >
                    <Upload className="w-3.5 h-3.5 text-indigo-600" />
                    Unggah Excel / CSV
                  </label>
                </div>
                <button
                  type="button"
                  onClick={handleImportPasteGrades}
                  className="px-3.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold inline-flex items-center gap-1.5"
                >
                  <Plus className="w-3.5 h-3.5" />
                  Terapkan Impor Nilai
                </button>
              </div>
            </div>
          )}

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-200/80">
            <span className="text-xs text-slate-500">
              {selectedExamStudentIds.size > 0
                ? `${selectedExamStudentIds.size} siswa dipilih`
                : `${studentsInClass.length} siswa · ${selectedExamType}`}
            </span>
            <TableExportToolbar
              selectedCount={selectedExamStudentIds.size}
              defaultOrientation="portrait"
              onDeleteSelected={() =>
                handleClearExamGrades(
                  Array.from(selectedExamStudentIds),
                  `${selectedExamStudentIds.size} siswa terpilih`
                )
              }
              onDownloadTemplate={handleDownloadGradeTemplate}
              onExportXlsx={() => {
                const { headers, rows } = buildExamDataset();
                exportRowsToXlsx(`Nilai ${selectedExamType}`, headers, rows, `Nilai_${selectedClass}_${selectedSubject}_${selectedExamType}.xlsx`);
                onToast('success', 'Nilai ujian diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildExamDataset();
                exportRowsToXls(`Nilai ${selectedExamType}`, headers, rows, `Nilai_${selectedClass}_${selectedSubject}_${selectedExamType}.xls`);
                onToast('success', 'Nilai ujian diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildExamDataset();
                exportDatasetToWordDoc({
                  db,
                  title: `DAFTAR NILAI ${selectedExamType.toUpperCase()} KELAS ${selectedClass}`,
                  subtitle: `${selectedSubject} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Nilai_${selectedClass}_${selectedExamType}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Nilai diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildExamDataset();
                exportDatasetToPdf({
                  db,
                  title: `DAFTAR NILAI ${selectedExamType.toUpperCase()} KELAS ${selectedClass}`,
                  subtitle: `${selectedSubject} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Nilai_${selectedClass}_${selectedExamType}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Nilai diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildExamDataset();
                printDatasetExact({
                  db,
                  title: `DAFTAR NILAI ${selectedExamType.toUpperCase()} KELAS ${selectedClass}`,
                  subtitle: `${selectedSubject} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('info', `Menyiapkan tampilan cetak daftar nilai ${selectedClass}...`);
              }}
            />
          </div>
        </div>

        <div className="overflow-x-auto max-h-[420px] table-scroll">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="sticky top-0 z-10 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 w-10 text-center">
                  <input
                    type="checkbox"
                    checked={
                      studentsInClass.length > 0 &&
                      studentsInClass.every((s) => selectedExamStudentIds.has(s.id))
                    }
                    onChange={toggleSelectAllExamStudents}
                    aria-label="Pilih semua siswa"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-3 py-3 w-10 text-center">No</th>
                <th className="px-3 py-3 whitespace-nowrap">NIS</th>
                <th className="px-3 py-3 min-w-[180px]">Nama Siswa</th>
                <th className="px-3 py-3 w-28 whitespace-nowrap">Nilai (0–100)</th>
                <th className="px-3 py-3 whitespace-nowrap">Status KKM</th>
                <th className="px-3 py-3 whitespace-nowrap">Penanda Baris</th>
                <th className="px-3 py-3 min-w-[180px]">Catatan</th>
                <th className="px-3 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {!hasExamSelection ? (
                <tr>
                  <td colSpan={9} className="px-4 py-10 text-center text-sm text-slate-500">
                    Silakan pilih Kelas, Mata Pelajaran, dan Jenis Ujian terlebih dahulu.
                  </td>
                </tr>
              ) : studentsInClass.length === 0 ? (
                <tr>
                  <td colSpan={9} className="px-4 py-10 text-center text-sm text-slate-500">
                    Belum ada siswa di kelas {selectedClass}.
                  </td>
                </tr>
              ) : (
                studentsInClass.map((st, idx) => {
                  const row = draftScores[st.id] ?? { score: '', note: '', marker: 'neutral' };
                  const markerMeta =
                    MARKER_CONFIG.find((m) => m.id === row.marker) ?? MARKER_CONFIG[0];
                  const numScore = row.score.trim() === '' ? null : Number(row.score);
                  const statusText =
                    numScore === null
                      ? 'Belum diisi'
                      : numScore >= 75
                      ? 'Tuntas'
                      : 'Remedial';
                  const statusColor =
                    numScore === null
                      ? 'text-slate-400'
                      : numScore >= 75
                      ? 'text-emerald-700 font-semibold'
                      : 'text-rose-700 font-semibold';
                  const checked = selectedExamStudentIds.has(st.id);

                  return (
                    <tr key={st.id} className={`${markerMeta.rowBgClass} transition-colors`}>
                      <td className="px-3 py-2.5 text-center">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectExamStudent(st.id)}
                          aria-label={`Pilih ${st.name}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-3 py-2.5 text-center text-xs text-slate-500 tabular-nums">
                        {idx + 1}
                      </td>
                      <td className="px-3 py-2.5 text-xs font-mono-tabular text-slate-600 whitespace-nowrap">
                        {st.nis}
                      </td>
                      <td className="px-3 py-2.5 font-medium text-slate-900 whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => setDetailStudent(st)}
                          title="Klik untuk melihat detail informasi siswa"
                          className="text-left font-medium text-slate-900 hover:text-indigo-600 hover:underline cursor-pointer"
                        >
                          {st.name}
                        </button>
                      </td>
                      <td className="px-3 py-2.5">
                        <input
                          type="number"
                          min={0}
                          max={100}
                          value={row.score}
                          onChange={(e) => handleScoreChange(st.id, e.target.value)}
                          placeholder="-"
                          aria-label={`Nilai ${st.name}`}
                          className="w-24 px-2.5 py-1.5 text-sm font-mono-tabular text-right border border-slate-200 bg-white rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
                        />
                      </td>
                      <td className={`px-3 py-2.5 text-xs whitespace-nowrap ${statusColor}`}>
                        {statusText}
                      </td>
                      <td className="px-3 py-2.5 whitespace-nowrap">
                        <div className="inline-flex items-center gap-2" role="radiogroup" aria-label={`Penanda baris ${st.name}`}>
                          {MARKER_CONFIG.map((m) => {
                            const active = row.marker === m.id;
                            return (
                              <button
                                key={m.id}
                                type="button"
                                onClick={() =>
                                  setDraftScores((prev) => ({
                                    ...prev,
                                    [st.id]: { ...row, marker: m.id },
                                  }))
                                }
                                title={m.label}
                                aria-label={m.label}
                                className={`w-5 h-5 rounded-full ${m.dotClass} transition-transform ${
                                  active
                                    ? 'ring-2 ring-offset-2 ring-slate-900 scale-110'
                                    : 'opacity-50 hover:opacity-100'
                                }`}
                              />
                            );
                          })}
                        </div>
                      </td>
                      <td className="px-3 py-2.5">
                        <input
                          type="text"
                          value={row.note}
                          onChange={(e) =>
                            setDraftScores((prev) => ({
                              ...prev,
                              [st.id]: { ...row, note: e.target.value },
                            }))
                          }
                          placeholder="Catatan opsional..."
                          className="w-full px-2.5 py-1.5 text-xs border border-slate-200 bg-white rounded-md focus:outline-none focus:border-indigo-600"
                        />
                      </td>
                      <td className="px-3 py-2.5 text-right whitespace-nowrap">
                        <button
                          type="button"
                          onClick={() => handleClearExamGrades([st.id], st.name)}
                          className="px-2 py-1 rounded border border-rose-200 text-xs text-rose-700 hover:bg-rose-50 inline-flex items-center gap-1"
                          title="Kosongkan nilai siswa ini"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                          Hapus
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </section>

      {/* BAGIAN 2: RIWAYAT MATRIKS NILAI PER KELAS */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div
          className={`px-5 py-4 space-y-3 ${
            matrixExpanded ? 'border-b border-slate-200' : ''
          }`}
        >
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
            <button
              type="button"
              onClick={() => setMatrixExpanded((prev) => !prev)}
              aria-expanded={matrixExpanded}
              className="flex items-center justify-between gap-3 text-left flex-1 min-w-0 group"
            >
              <div className="min-w-0">
                <div className="flex flex-wrap items-center gap-2">
                  <h2 className="text-base font-bold text-slate-900 group-hover:text-indigo-700 transition-colors">
                    2. Riwayat Matriks Nilai per Kelas
                  </h2>
                  <span className="text-xs text-slate-500">
                    · {matrixClass || '-'} · {matrixSubject || '-'}
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-0.5">
                  Leger nilai mendatar seluruh jenis ujian — edit nilai langsung di sel tabel dan simpan
                </p>
              </div>
              <span className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 group-hover:bg-slate-100 text-xs font-semibold text-slate-700 whitespace-nowrap shrink-0 transition-colors">
                <span>{matrixExpanded ? 'Ciutkan' : 'Tampilkan'}</span>
                <ChevronDown
                  className={`w-4 h-4 text-slate-500 transition-transform duration-150 ${
                    matrixExpanded ? 'rotate-180' : ''
                  }`}
                />
              </span>
            </button>

            {matrixExpanded && (
            <div className="flex flex-wrap items-center gap-2.5">
              <select
                value={matrixClass}
                onChange={(e) => setMatrixClass(e.target.value)}
                aria-label="Pilih kelas matriks nilai"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Kelas--</option>
                {classOptions.map((c) => (
                  <option key={c} value={c}>
                    Kelas {c}
                  </option>
                ))}
              </select>

              <select
                value={matrixSubject}
                onChange={(e) => setMatrixSubject(e.target.value)}
                aria-label="Pilih mata pelajaran matriks nilai"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Mapel--</option>
                {settings.masterSubjects.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>

              <button
                type="button"
                onClick={handleSaveMatrix}
                disabled={!matrixDirty || !hasMatrixSelection || matrixStudents.length === 0}
                className={`px-4 py-2 rounded-lg text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap ${
                  matrixDirty && hasMatrixSelection && matrixStudents.length > 0
                    ? 'bg-indigo-600 hover:bg-indigo-700 text-white'
                    : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                }`}
              >
                <Save className="w-4 h-4" />
                Simpan Matriks
              </button>
            </div>
            )}
          </div>

          {matrixExpanded && (
          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedMatrixIds.size > 0
                ? `${selectedMatrixIds.size} siswa dipilih`
                : `Leger Nilai ${matrixClass} · ${matrixSubject}`}
            </span>
            <TableExportToolbar
              selectedCount={selectedMatrixIds.size}
              onDeleteSelected={handleDeleteMatrixSelected}
              onExportXlsx={() => {
                const { headers, rows } = buildMatrixDataset();
                exportRowsToXlsx(`Matriks ${matrixClass}`, headers, rows, `Matriks_Nilai_${matrixClass}_${matrixSubject}.xlsx`);
                onToast('success', 'Matriks nilai diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildMatrixDataset();
                exportRowsToXls(`Matriks ${matrixClass}`, headers, rows, `Matriks_Nilai_${matrixClass}_${matrixSubject}.xls`);
                onToast('success', 'Matriks nilai diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildMatrixDataset();
                exportDatasetToWordDoc({
                  db,
                  title: `MATRIKS NILAI SISWA KELAS ${matrixClass}`,
                  subtitle: `${matrixSubject} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Matriks_Nilai_${matrixClass}_${matrixSubject}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Matriks Nilai diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildMatrixDataset();
                exportDatasetToPdf({
                  db,
                  title: `MATRIKS NILAI SISWA KELAS ${matrixClass}`,
                  subtitle: `${matrixSubject} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Matriks_Nilai_${matrixClass}_${matrixSubject}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Matriks Nilai diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildMatrixDataset();
                printDatasetExact({
                  db,
                  title: `MATRIKS NILAI SISWA KELAS ${matrixClass}`,
                  subtitle: `${matrixSubject} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('info', `Menyiapkan tampilan cetak matriks nilai ${matrixClass}...`);
              }}
            />
          </div>
          )}
        </div>

        {matrixExpanded && (
        <div className="overflow-x-auto max-h-[460px] table-scroll">
          <table className="w-full text-left border-collapse text-xs sm:text-sm">
            <thead className="sticky top-0 z-20 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="sticky left-0 z-30 bg-slate-100 px-3 py-3 w-10 text-center border-r border-slate-200">
                  <input
                    type="checkbox"
                    checked={
                      matrixStudents.length > 0 &&
                      matrixStudents.every((s) => selectedMatrixIds.has(s.id))
                    }
                    onChange={toggleSelectAllMatrix}
                    aria-label="Pilih semua siswa matriks"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-3 py-3 w-10 text-center">No</th>
                <th className="sticky left-10 z-30 bg-slate-100 px-3 py-3 min-w-[175px] border-r border-slate-200">
                  Nama Siswa
                </th>
                {settings.masterExamTypes.map((ex) => (
                  <th key={ex} className="px-2.5 py-3 text-center whitespace-nowrap min-w-[82px]">
                    {ex}
                  </th>
                ))}
                <th className="px-3 py-3 text-right whitespace-nowrap border-l border-slate-200 bg-slate-100">
                  Rata-rata
                </th>
                <th className="px-3 py-3 text-left whitespace-nowrap bg-slate-100">
                  Status (KKM 75)
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {!hasMatrixSelection ? (
                <tr>
                  <td
                    colSpan={5 + settings.masterExamTypes.length}
                    className="px-4 py-10 text-center text-sm text-slate-500"
                  >
                    Silakan pilih Kelas dan Mata Pelajaran terlebih dahulu untuk melihat matriks nilai.
                  </td>
                </tr>
              ) : matrixStudents.length === 0 ? (
                <tr>
                  <td
                    colSpan={5 + settings.masterExamTypes.length}
                    className="px-4 py-10 text-center text-sm text-slate-500"
                  >
                    Belum ada siswa di kelas {matrixClass}.
                  </td>
                </tr>
              ) : (
                matrixStudents.map((st, idx) => {
                  const numericVals: number[] = [];
                  settings.masterExamTypes.forEach((ex) => {
                    const raw = matrixDraft[`${st.id}__${ex}`] ?? '';
                    if (raw.trim() !== '') {
                      const num = Number(raw);
                      if (!Number.isNaN(num)) numericVals.push(num);
                    }
                  });
                  const avg =
                    numericVals.length > 0
                      ? numericVals.reduce((a, b) => a + b, 0) / numericVals.length
                      : null;
                  const statusLabel =
                    avg === null ? 'Belum ada nilai' : avg >= 75 ? 'Tuntas' : 'Remedial';
                  const statusClass =
                    avg === null
                      ? 'text-slate-400'
                      : avg >= 75
                      ? 'text-emerald-700 font-semibold'
                      : 'text-rose-700 font-semibold';
                  const checked = selectedMatrixIds.has(st.id);

                  return (
                    <tr key={st.id} className={`hover:bg-slate-50/70 ${checked ? 'bg-indigo-50/40' : ''}`}>
                      <td className="sticky left-0 z-10 bg-white px-3 py-2 text-center border-r border-slate-200">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectMatrix(st.id)}
                          aria-label={`Pilih ${st.name}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-3 py-2 text-center text-xs text-slate-500 tabular-nums">
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
                      {settings.masterExamTypes.map((ex) => (
                        <td key={ex} className="px-1.5 py-1.5 text-center">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={matrixDraft[`${st.id}__${ex}`] ?? ''}
                            onChange={(e) => handleMatrixCellChange(st.id, ex, e.target.value)}
                            placeholder="-"
                            aria-label={`Nilai ${ex} untuk ${st.name}`}
                            className="w-16 px-2 py-1.5 text-center text-xs font-mono-tabular border border-slate-200 rounded bg-white focus:outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600"
                          />
                        </td>
                      ))}
                      <td className="px-3 py-2 text-right font-mono-tabular font-bold text-slate-900 border-l border-slate-200 whitespace-nowrap">
                        {avg !== null ? avg.toFixed(1) : '-'}
                      </td>
                      <td className={`px-3 py-2 text-xs whitespace-nowrap ${statusClass}`}>
                        {statusLabel}
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

      <StudentDetailModal
        student={detailStudent}
        db={db}
        onClose={() => setDetailStudent(null)}
      />
    </div>
  );
}
