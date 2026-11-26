import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  CalendarClock,
  ClipboardPlus,
  Pencil,
  Plus,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import { DayOfWeek, PortalDatabase, TeachingSchedule } from '../types';
import {
  exportDatasetToPdf,
  exportDatasetToWordDoc,
  exportRowsToXls,
  exportRowsToXlsx,
  printDatasetExact,
} from '../utils/exportHelpers';
import { getTodayISO } from '../utils/portalHelpers';
import { ConfirmDialogState } from './FeedbackUI';
import { FormModal } from './FormModal';
import { TableExportToolbar } from './TableExportToolbar';

interface ScheduleViewProps {
  db: PortalDatabase;
  onUpsertSchedule: (item: TeachingSchedule) => void;
  onImportSchedulesBatch: (incoming: Omit<TeachingSchedule, 'id'>[]) => number;
  onDeleteSchedule: (id: string) => void;
  onDeleteSchedules: (ids: string[]) => void;
  onCreateJournalFromSchedule: (item: TeachingSchedule) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const DAYS_OF_WEEK: DayOfWeek[] = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const SCHEDULE_HEADERS = ['Hari', 'Jam Ke', 'Waktu', 'Kelas', 'Mata Pelajaran', 'Guru Pengampu', 'Ruang', 'Keterangan'];

export function ScheduleView({
  db,
  onUpsertSchedule,
  onImportSchedulesBatch,
  onDeleteSchedule,
  onDeleteSchedules,
  onCreateJournalFromSchedule,
  onToast,
  onConfirm,
}: ScheduleViewProps) {
  const { schedules, teachers, students, settings, session } = db;

  const classOptions = useMemo(() => {
    const set = new Set(settings.masterClasses);
    students.forEach((s) => set.add(s.className));
    return Array.from(set);
  }, [settings.masterClasses, students]);

  const teacherOptions = useMemo(() => {
    const set = new Set<string>();
    if (settings.teacherName) set.add(settings.teacherName);
    teachers.forEach((t) => {
      if (t.name) set.add(t.name);
    });
    return Array.from(set);
  }, [teachers, settings.teacherName]);

  // Form state
  const [editingId, setEditingId] = useState<string | null>(null);
  const [day, setDay] = useState<DayOfWeek>('Senin');
  const [jamKe, setJamKe] = useState('1-2');
  const [timeRange, setTimeRange] = useState('07:00 - 08:30');
  const [className, setClassName] = useState(classOptions[0] || '');
  const [subject, setSubject] = useState(settings.masterSubjects[0] || '');
  const [teacherName, setTeacherName] = useState(teacherOptions[0] || '');
  const [room, setRoom] = useState('');
  const [notes, setNotes] = useState('');

  // Import state
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Filter & Selection state
  const [filterDay, setFilterDay] = useState<string>('ALL');
  const [filterClass, setFilterClass] = useState<string>('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);

  const resetForm = () => {
    setEditingId(null);
    setDay('Senin');
    setJamKe('1-2');
    setTimeRange('07:00 - 08:30');
    setClassName(classOptions[0] || '');
    setSubject(settings.masterSubjects[0] || '');
    setTeacherName(teacherOptions[0] || '');
    setRoom('');
    setNotes('');
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!className || !subject || !teacherName.trim() || !jamKe.trim()) {
      onToast('error', 'Hari, Jam Ke, Kelas, Mata Pelajaran, dan Guru Pengampu wajib diisi.');
      return;
    }
    onUpsertSchedule({
      id: editingId ?? `sch-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      day,
      jamKe: jamKe.trim(),
      timeRange: timeRange.trim() || '-',
      className,
      subject,
      teacherName: teacherName.trim(),
      room: room.trim() || `Ruang ${className}`,
      academicYear: session.academicYear,
      semester: session.semester,
      notes: notes.trim(),
    });
    onToast(
      'success',
      editingId ? 'Jadwal mengajar berhasil diperbarui.' : 'Jadwal mengajar baru ditambahkan.'
    );
    resetForm();
    setFormModalOpen(false);
  };

  const handleEdit = (item: TeachingSchedule) => {
    setEditingId(item.id);
    setDay(item.day);
    setJamKe(item.jamKe);
    setTimeRange(item.timeRange);
    setClassName(item.className);
    setSubject(item.subject);
    setTeacherName(item.teacherName);
    setRoom(item.room);
    setNotes(item.notes);
    setFormModalOpen(true);
  };

  const handleDelete = (item: TeachingSchedule) => {
    onConfirm({
      title: 'Hapus Jadwal Mengajar',
      description: `Hapus jadwal ${item.day} (${item.jamKe}) kelas ${item.className} — ${item.subject}?`,
      confirmLabel: 'Hapus Jadwal',
      danger: true,
      onConfirm: () => {
        onDeleteSchedule(item.id);
        setSelectedIds((prev) => {
          const next = new Set(prev);
          next.delete(item.id);
          return next;
        });
        onToast('success', 'Jadwal berhasil dihapus.');
      },
    });
  };

  const activeSchedules = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return schedules
      .filter((s) => {
        if (s.academicYear !== session.academicYear || s.semester !== session.semester) {
          return false;
        }
        if (filterDay !== 'ALL' && s.day !== filterDay) return false;
        if (filterClass !== 'ALL' && s.className !== filterClass) return false;
        if (q) {
          return (
            s.subject.toLowerCase().includes(q) ||
            s.teacherName.toLowerCase().includes(q) ||
            s.className.toLowerCase().includes(q) ||
            s.room.toLowerCase().includes(q)
          );
        }
        return true;
      })
      .sort((a, b) => {
        const dDiff = DAYS_OF_WEEK.indexOf(a.day) - DAYS_OF_WEEK.indexOf(b.day);
        if (dDiff !== 0) return dDiff;
        return a.jamKe.localeCompare(b.jamKe);
      });
  }, [schedules, session.academicYear, session.semester, filterDay, filterClass, searchQuery]);

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
      activeSchedules.length > 0 &&
      activeSchedules.every((s) => selectedIds.has(s.id))
    ) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(activeSchedules.map((s) => s.id)));
    }
  };

  const handleBulkDelete = () => {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    onConfirm({
      title: 'Hapus Jadwal Terpilih',
      description: `Hapus ${count} jadwal mengajar yang dipilih?`,
      confirmLabel: `Hapus ${count} Jadwal`,
      danger: true,
      onConfirm: () => {
        onDeleteSchedules(Array.from(selectedIds));
        setSelectedIds(new Set());
        onToast('success', `${count} jadwal mengajar berhasil dihapus.`);
      },
    });
  };

  const exportSource = useMemo(() => {
    if (selectedIds.size > 0) {
      const chosen = activeSchedules.filter((s) => selectedIds.has(s.id));
      if (chosen.length > 0) return chosen;
    }
    return activeSchedules;
  }, [activeSchedules, selectedIds]);

  const buildScheduleMatrix = () => {
    const headers = ['No', ...SCHEDULE_HEADERS];
    const rows = exportSource.map((s, idx) => [
      idx + 1,
      s.day,
      s.jamKe,
      s.timeRange,
      s.className,
      s.subject,
      s.teacherName,
      s.room,
      s.notes || '-',
    ]);
    return { headers, rows };
  };

  const handleDownloadTemplate = () => {
    exportRowsToXlsx(
      'Template Jadwal',
      SCHEDULE_HEADERS,
      [
        [
          'Senin',
          '1-2',
          '07:00 - 08:30',
          classOptions[0] || 'X-A',
          settings.masterSubjects[0] || 'Matematika',
          teacherOptions[0] || settings.teacherName,
          'Ruang X-A',
          'Pertemuan reguler',
        ],
      ],
      'Template_Jadwal_Mengajar_E-Guru.xlsx'
    );
    onToast('success', 'Template Excel (.xlsx) Jadwal Mengajar berhasil diunduh.');
  };

  const parseRowsToSchedules = (rawRows: string[][]): Omit<TeachingSchedule, 'id'>[] => {
    const result: Omit<TeachingSchedule, 'id'>[] = [];
    rawRows.forEach((cols, index) => {
      if (cols.length < 4) return;
      const clean = (i: number) => String(cols[i] ?? '').trim();
      if (index === 0 && clean(0).toLowerCase() === 'hari') return;
      const rawDay = clean(0);
      const validDay: DayOfWeek = DAYS_OF_WEEK.includes(rawDay as DayOfWeek)
        ? (rawDay as DayOfWeek)
        : 'Senin';
      const cls = clean(3);
      const subj = clean(4);
      if (!cls || !subj) return;
      result.push({
        day: validDay,
        jamKe: clean(1) || '1-2',
        timeRange: clean(2) || '-',
        className: cls,
        subject: subj,
        teacherName: clean(5) || settings.teacherName,
        room: clean(6) || `Ruang ${cls}`,
        notes: clean(7),
        academicYear: session.academicYear,
        semester: session.semester,
      });
    });
    return result;
  };

  const handleImportPaste = () => {
    const lines = pasteText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) {
      onToast('error', 'Tempelkan baris jadwal terlebih dahulu.');
      return;
    }
    const matrix = lines.map((l) => (l.includes('\t') ? l.split('\t') : l.includes(';') ? l.split(';') : l.split(',')));
    const parsed = parseRowsToSchedules(matrix);
    if (parsed.length === 0) {
      onToast('error', 'Format jadwal tidak valid.');
      return;
    }
    const count = onImportSchedulesBatch(parsed);
    setPasteText('');
    setImportModalOpen(false);
    onToast('success', `${count} jadwal mengajar berhasil diimpor.`);
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' });
      const parsed = parseRowsToSchedules(rows);
      if (parsed.length === 0) {
        onToast('error', 'Tidak ditemukan baris jadwal yang valid.');
        return;
      }
      const count = onImportSchedulesBatch(parsed);
      setImportModalOpen(false);
      onToast('success', `${count} jadwal berhasil diimpor dari "${file.name}".`);
    } catch {
      onToast('error', 'Gagal membaca file Excel/CSV.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="space-y-6">
      {/* Compact Top Action Bar */}
      <section className="bg-white border border-slate-200 rounded-xl px-5 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <CalendarClock className="w-5 h-5 text-indigo-600 shrink-0" />
          <div>
            <h2 className="text-base font-bold text-slate-900">
              Jadwal Guru Mengajar ({activeSchedules.length})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              TA {session.academicYear} · Semester {session.semester}
            </p>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => setImportModalOpen(true)}
            className="px-3.5 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
          >
            <Upload className="w-4 h-4 text-indigo-600" />
            <span>Impor Jadwal</span>
          </button>
          <button
            type="button"
            onClick={() => {
              resetForm();
              setFormModalOpen(true);
            }}
            className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shadow-xs"
          >
            <Plus className="w-4 h-4" />
            <span>Tambah Jadwal Guru Mengajar</span>
          </button>
        </div>
      </section>

      {/* Tabel Daftar Jadwal Mengajar */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 space-y-3">
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Daftar Jadwal Mengajar ({activeSchedules.length})
              </h2>
              <p className="text-xs text-slate-500">
                Pilih baris untuk hapus massal, ekspor XLSX/XLS/Word/PDF, atau cetak jadwal resmi
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <select
                value={filterDay}
                onChange={(e) => setFilterDay(e.target.value)}
                aria-label="Filter hari"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="ALL">Semua Hari</option>
                {DAYS_OF_WEEK.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>

              <select
                value={filterClass}
                onChange={(e) => setFilterClass(e.target.value)}
                aria-label="Filter kelas"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="ALL">Semua Kelas</option>
                {classOptions.map((c) => (
                  <option key={c} value={c}>
                    Kelas {c}
                  </option>
                ))}
              </select>

              <div className="relative">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Cari mapel / guru..."
                  className="pl-9 pr-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 w-44 sm:w-52"
                />
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedIds.size > 0
                ? `${selectedIds.size} jadwal dipilih`
                : 'Centang baris untuk ekspor selektif / hapus massal'}
            </span>
            <TableExportToolbar
              selectedCount={selectedIds.size}
              onDeleteSelected={handleBulkDelete}
              onDownloadTemplate={handleDownloadTemplate}
              onExportXlsx={() => {
                const { headers, rows } = buildScheduleMatrix();
                exportRowsToXlsx('Jadwal Mengajar', headers, rows, `Jadwal_Mengajar_${getTodayISO()}.xlsx`);
                onToast('success', 'Jadwal mengajar berhasil diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildScheduleMatrix();
                exportRowsToXls('Jadwal Mengajar', headers, rows, `Jadwal_Mengajar_${getTodayISO()}.xls`);
                onToast('success', 'Jadwal mengajar berhasil diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildScheduleMatrix();
                exportDatasetToWordDoc({
                  db,
                  title: 'JADWAL GURU MENGAJAR',
                  subtitle: `${settings.schoolName} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Jadwal_Mengajar_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Jadwal Mengajar diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildScheduleMatrix();
                exportDatasetToPdf({
                  db,
                  title: 'JADWAL GURU MENGAJAR',
                  subtitle: `${settings.schoolName} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Jadwal_Mengajar_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Jadwal Mengajar diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildScheduleMatrix();
                printDatasetExact({
                  db,
                  title: 'JADWAL GURU MENGAJAR',
                  subtitle: `${settings.schoolName} · TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
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
                      activeSchedules.length > 0 &&
                      activeSchedules.every((s) => selectedIds.has(s.id))
                    }
                    onChange={toggleSelectAll}
                    aria-label="Pilih semua jadwal"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-3 py-3 w-10 text-center">No</th>
                <th className="px-4 py-3 whitespace-nowrap">Hari</th>
                <th className="px-4 py-3 whitespace-nowrap">Jam / Waktu</th>
                <th className="px-4 py-3 whitespace-nowrap">Kelas</th>
                <th className="px-4 py-3 whitespace-nowrap">Mata Pelajaran</th>
                <th className="px-4 py-3 min-w-[180px]">Guru Pengampu</th>
                <th className="px-4 py-3 whitespace-nowrap">Ruang</th>
                <th className="px-4 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {activeSchedules.length === 0 ? (
                <tr>
                  <td colSpan={9} className="px-4 py-10 text-center text-sm text-slate-500">
                    Belum ada jadwal mengajar yang sesuai filter.
                  </td>
                </tr>
              ) : (
                activeSchedules.map((item, idx) => {
                  const checked = selectedIds.has(item.id);
                  return (
                    <tr
                      key={item.id}
                      className={`hover:bg-slate-50/80 transition-colors ${checked ? 'bg-indigo-50/40' : ''}`}
                    >
                      <td className="px-3 py-3 text-center">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectRow(item.id)}
                          aria-label={`Pilih jadwal ${item.className} ${item.subject}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-3 py-3 text-center text-xs text-slate-500 tabular-nums">
                        {idx + 1}
                      </td>
                      <td className="px-4 py-3 font-semibold text-slate-900 whitespace-nowrap">
                        {item.day}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                        <p className="font-mono-tabular text-xs font-semibold text-slate-900">
                          Jam {item.jamKe}
                        </p>
                        <p className="text-xs text-slate-500 font-mono-tabular">{item.timeRange}</p>
                      </td>
                      <td className="px-4 py-3 font-semibold text-indigo-700 whitespace-nowrap">
                        {item.className}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-slate-800">{item.subject}</td>
                      <td className="px-4 py-3 whitespace-nowrap">{item.teacherName}</td>
                      <td className="px-4 py-3 whitespace-nowrap text-xs text-slate-600">
                        {item.room}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-right">
                        <div className="inline-flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => onCreateJournalFromSchedule(item)}
                            className="px-2.5 py-1.5 rounded-md bg-indigo-50 border border-indigo-200 text-xs font-semibold text-indigo-700 hover:bg-indigo-100 inline-flex items-center gap-1"
                            title="Isi Jurnal & Presensi dari jadwal ini"
                          >
                            <ClipboardPlus className="w-3.5 h-3.5" />
                            Isi Jurnal
                          </button>
                          <button
                            type="button"
                            onClick={() => handleEdit(item)}
                            className="px-2.5 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 inline-flex items-center gap-1"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                            Edit
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDelete(item)}
                            className="px-2.5 py-1.5 rounded-md border border-rose-200 text-xs font-medium text-rose-700 bg-rose-50/60 hover:bg-rose-100 inline-flex items-center gap-1"
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
      </section>

      <FormModal
        open={formModalOpen}
        title={editingId ? 'Edit Jadwal Guru Mengajar' : 'Tambah Jadwal Guru Mengajar'}
        subtitle={`Atur jadwal mingguan KBM untuk Tahun Ajaran ${session.academicYear} · Semester ${session.semester}`}
        maxWidthClass="max-w-4xl"
        onClose={() => {
          setFormModalOpen(false);
          resetForm();
        }}
      >
        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label htmlFor="sch-day" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Hari
              </label>
              <select
                id="sch-day"
                value={day}
                onChange={(e) => setDay(e.target.value as DayOfWeek)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                {DAYS_OF_WEEK.map((d) => (
                  <option key={d} value={d}>
                    {d}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label htmlFor="sch-jam" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Jam Ke-
              </label>
              <input
                id="sch-jam"
                type="text"
                value={jamKe}
                onChange={(e) => setJamKe(e.target.value)}
                placeholder="Contoh: 1-2"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="sch-time" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Rentang Waktu
              </label>
              <input
                id="sch-time"
                type="text"
                value={timeRange}
                onChange={(e) => setTimeRange(e.target.value)}
                placeholder="07:00 - 08:30"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="sch-class" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Kelas
              </label>
              <select
                id="sch-class"
                value={className}
                onChange={(e) => setClassName(e.target.value)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                {classOptions.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label htmlFor="sch-subject" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Mata Pelajaran
              </label>
              <select
                id="sch-subject"
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                {settings.masterSubjects.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label htmlFor="sch-teacher" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Guru Pengampu
              </label>
              <input
                id="sch-teacher"
                list="schedule-teacher-list"
                value={teacherName}
                onChange={(e) => setTeacherName(e.target.value)}
                placeholder="Nama guru pengampu"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
              <datalist id="schedule-teacher-list">
                {teacherOptions.map((t) => (
                  <option key={t} value={t} />
                ))}
              </datalist>
            </div>

            <div>
              <label htmlFor="sch-room" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Ruang / Lab
              </label>
              <input
                id="sch-room"
                type="text"
                value={room}
                onChange={(e) => setRoom(e.target.value)}
                placeholder="Contoh: Ruang X-A / Lab 1"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="sch-notes" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Catatan
              </label>
              <input
                id="sch-notes"
                type="text"
                value={notes}
                onChange={(e) => setNotes(e.target.value)}
                placeholder="Opsional..."
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>
          </div>

          <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button
              type="button"
              onClick={() => {
                resetForm();
                setFormModalOpen(false);
              }}
              className="px-4 py-2 rounded-lg border border-slate-200 text-sm font-medium text-slate-700 hover:bg-slate-50"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={resetForm}
              className="px-4 py-2 rounded-lg border border-slate-200 text-sm font-medium text-slate-700 hover:bg-slate-50"
            >
              Reset
            </button>
            <button
              type="submit"
              className="px-5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
            >
              <Plus className="w-4 h-4" />
              {editingId ? 'Simpan Perubahan' : 'Tambahkan Jadwal'}
            </button>
          </div>
        </form>
      </FormModal>

      <FormModal
        open={importModalOpen}
        title="Impor Jadwal Mengajar (Excel / CSV / Tempel)"
        subtitle="Urutan kolom: Hari, Jam Ke, Waktu, Kelas, Mata Pelajaran, Guru Pengampu, Ruang, Keterangan"
        maxWidthClass="max-w-2xl"
        onClose={() => setImportModalOpen(false)}
      >
        <div className="p-5 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <button
              type="button"
              onClick={handleDownloadTemplate}
              className="px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-xs font-semibold text-slate-700"
            >
              Unduh Template Excel (.xlsx)
            </button>
            <div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleImportFile}
                className="hidden"
                id="schedule-file-upload"
              />
              <label
                htmlFor="schedule-file-upload"
                className="cursor-pointer px-3 py-1.5 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-xs font-semibold text-indigo-700 inline-flex items-center gap-1.5"
              >
                <Upload className="w-3.5 h-3.5" />
                Unggah File Excel / CSV
              </label>
            </div>
          </div>

          <textarea
            rows={4}
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            placeholder="Tempel baris jadwal dari Excel (Tab / CSV)..."
            className="w-full px-3 py-2 text-xs font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
          />

          <div className="flex justify-end gap-2 pt-2 border-t border-slate-100">
            <button
              type="button"
              onClick={() => setImportModalOpen(false)}
              className="px-4 py-2 rounded-lg border border-slate-200 text-sm font-medium text-slate-700 hover:bg-slate-50"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={handleImportPaste}
              className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold inline-flex items-center gap-1.5"
            >
              <Plus className="w-4 h-4" />
              Proses Impor
            </button>
          </div>
        </div>
      </FormModal>
    </div>
  );
}
