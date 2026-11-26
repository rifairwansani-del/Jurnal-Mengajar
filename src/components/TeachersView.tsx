import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ImagePlus,
  Pencil,
  Plus,
  Search,
  Trash2,
  Upload,
  UserCheck,
} from 'lucide-react';
import { PortalDatabase, Teacher, TeacherEmploymentStatus } from '../types';
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

interface TeachersViewProps {
  db: PortalDatabase;
  onUpsertTeacher: (teacher: Teacher) => void;
  onImportTeachersBatch: (incoming: Omit<Teacher, 'id'>[]) => { added: number; updated: number };
  onDeleteTeacher: (id: string) => void;
  onDeleteTeachersBatch: (ids: string[]) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const EMPLOYMENT_TYPES: TeacherEmploymentStatus[] = [
  'PNS',
  'PPPK',
  'Tetap Yayasan',
  'Honorer',
];

const TEACHER_IMPORT_HEADERS = [
  'NIP',
  'Nama Lengkap',
  'L/P',
  'Jabatan',
  'Mata Pelajaran',
  'Kepegawaian',
  'Status',
  'Email',
  'No HP',
  'Alamat',
];

export function TeachersView({
  db,
  onUpsertTeacher,
  onImportTeachersBatch,
  onDeleteTeacher,
  onDeleteTeachersBatch,
  onToast,
  onConfirm,
}: TeachersViewProps) {
  const { teachers, settings } = db;

  const [editingId, setEditingId] = useState<string | null>(null);
  const [nip, setNip] = useState('');
  const [name, setName] = useState('');
  const [position, setPosition] = useState('Guru Mata Pelajaran');
  const [subject, setSubject] = useState(settings.masterSubjects[0] || '');
  const [email, setEmail] = useState('');
  const [phone, setPhone] = useState('');
  const [gender, setGender] = useState<'L' | 'P'>('L');
  const [employmentStatus, setEmploymentStatus] = useState<TeacherEmploymentStatus>('PNS');
  const [status, setStatus] = useState<'Aktif' | 'Cuti' | 'Nonaktif'>('Aktif');
  const [address, setAddress] = useState('');
  const [photo, setPhoto] = useState('');

  // Import state
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Filters & Multi-Select
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);

  const resetForm = () => {
    setEditingId(null);
    setNip('');
    setName('');
    setPosition('Guru Mata Pelajaran');
    setSubject(settings.masterSubjects[0] || '');
    setEmail('');
    setPhone('');
    setGender('L');
    setEmploymentStatus('PNS');
    setStatus('Aktif');
    setAddress('');
    setPhoto('');
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran foto maksimal 2 MB.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') setPhoto(reader.result);
    };
    reader.readAsDataURL(file);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim()) {
      onToast('error', 'Nama Guru wajib diisi.');
      return;
    }
    onUpsertTeacher({
      id: editingId ?? `tch-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      nip: nip.trim(),
      name: name.trim(),
      position: position.trim() || 'Guru Mata Pelajaran',
      subject: subject.trim() || '-',
      email: email.trim(),
      phone: phone.trim(),
      gender,
      employmentStatus,
      status,
      address: address.trim(),
      photo,
    });
    onToast(
      'success',
      editingId ? 'Data guru berhasil diperbarui.' : 'Data guru baru berhasil ditambahkan.'
    );
    resetForm();
    setFormModalOpen(false);
  };

  const handleEdit = (t: Teacher) => {
    setEditingId(t.id);
    setNip(t.nip);
    setName(t.name);
    setPosition(t.position);
    setSubject(t.subject);
    setEmail(t.email);
    setPhone(t.phone);
    setGender(t.gender);
    setEmploymentStatus(t.employmentStatus);
    setStatus(t.status);
    setAddress(t.address);
    setPhoto(t.photo);
    setFormModalOpen(true);
  };

  const handleDelete = (t: Teacher) => {
    onConfirm({
      title: 'Hapus Data Guru',
      description: `Anda yakin ingin menghapus data guru "${t.name}"?`,
      confirmLabel: 'Hapus Guru',
      danger: true,
      onConfirm: () => {
        onDeleteTeacher(t.id);
        setSelectedIds((prev) => {
          const next = new Set(prev);
          next.delete(t.id);
          return next;
        });
        onToast('success', `Data guru ${t.name} dihapus.`);
      },
    });
  };

  const filteredTeachers = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return teachers
      .filter((t) => {
        if (statusFilter !== 'ALL' && t.status !== statusFilter) return false;
        if (q) {
          return (
            t.name.toLowerCase().includes(q) ||
            t.nip.toLowerCase().includes(q) ||
            t.subject.toLowerCase().includes(q) ||
            t.position.toLowerCase().includes(q)
          );
        }
        return true;
      })
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [teachers, statusFilter, searchQuery]);

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
      filteredTeachers.length > 0 &&
      filteredTeachers.every((t) => selectedIds.has(t.id))
    ) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredTeachers.map((t) => t.id)));
    }
  };

  const handleBulkDelete = () => {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    onConfirm({
      title: 'Hapus Data Guru Terpilih',
      description: `Hapus ${count} data guru yang dipilih?`,
      confirmLabel: `Hapus ${count} Guru`,
      danger: true,
      onConfirm: () => {
        onDeleteTeachersBatch(Array.from(selectedIds));
        setSelectedIds(new Set());
        onToast('success', `${count} data guru berhasil dihapus.`);
      },
    });
  };

  const exportSourceTeachers = useMemo(() => {
    if (selectedIds.size > 0) {
      const chosen = filteredTeachers.filter((t) => selectedIds.has(t.id));
      if (chosen.length > 0) return chosen;
    }
    return filteredTeachers;
  }, [filteredTeachers, selectedIds]);

  const buildTeacherMatrix = () => {
    const headers = ['No', ...TEACHER_IMPORT_HEADERS];
    const rows = exportSourceTeachers.map((t, idx) => [
      idx + 1,
      t.nip || '-',
      t.name,
      t.gender,
      t.position,
      t.subject,
      t.employmentStatus,
      t.status,
      t.email || '-',
      t.phone || '-',
      t.address || '-',
    ]);
    return { headers, rows };
  };

  const handleDownloadTemplate = () => {
    exportRowsToXlsx(
      'Template Guru',
      TEACHER_IMPORT_HEADERS,
      [
        [
          '19870615 201402 1 002',
          'Drs. Bambang Sutrisno, M.Pd.',
          'L',
          'Guru Mata Pelajaran',
          settings.masterSubjects[0] || 'Matematika',
          'PNS',
          'Aktif',
          'bambang@sekolah.sch.id',
          '081234567800',
          'Jl. Kartini No. 12, Kudus',
        ],
      ],
      'Template_Impor_Guru_E-Guru.xlsx'
    );
    onToast('success', 'Template Excel (.xlsx) Data Guru berhasil diunduh.');
  };

  const parseRowsToTeachers = (rawRows: string[][]): Omit<Teacher, 'id'>[] => {
    const result: Omit<Teacher, 'id'>[] = [];
    rawRows.forEach((cols, index) => {
      if (cols.length < 2) return;
      const clean = (i: number) => String(cols[i] ?? '').trim();
      if (index === 0 && (clean(0).toLowerCase() === 'nip' || clean(1).toLowerCase().includes('nama'))) {
        return;
      }
      const tName = clean(1) || clean(0);
      if (!tName) return;
      const rawEmp = clean(5);
      const validEmp: TeacherEmploymentStatus = EMPLOYMENT_TYPES.includes(
        rawEmp as TeacherEmploymentStatus
      )
        ? (rawEmp as TeacherEmploymentStatus)
        : 'PNS';
      const rawStatus = clean(6);
      const validStatus: 'Aktif' | 'Cuti' | 'Nonaktif' =
        rawStatus === 'Cuti' || rawStatus === 'Nonaktif' ? rawStatus : 'Aktif';

      result.push({
        nip: clean(0),
        name: tName,
        gender: clean(2).toUpperCase().startsWith('P') ? 'P' : 'L',
        position: clean(3) || 'Guru Mata Pelajaran',
        subject: clean(4) || '-',
        employmentStatus: validEmp,
        status: validStatus,
        email: clean(7),
        phone: clean(8),
        address: clean(9),
        photo: '',
      });
    });
    return result;
  };

  const handleImportPaste = () => {
    const lines = pasteText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) {
      onToast('error', 'Tempelkan baris data guru terlebih dahulu.');
      return;
    }
    const matrix = lines.map((l) => (l.includes('\t') ? l.split('\t') : l.includes(';') ? l.split(';') : l.split(',')));
    const parsed = parseRowsToTeachers(matrix);
    if (parsed.length === 0) {
      onToast('error', 'Format data guru tidak valid.');
      return;
    }
    const stats = onImportTeachersBatch(parsed);
    setPasteText('');
    setImportModalOpen(false);
    onToast('success', `Impor selesai: ${stats.added} ditambah, ${stats.updated} diperbarui.`);
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' });
      const parsed = parseRowsToTeachers(rows);
      if (parsed.length === 0) {
        onToast('error', 'Tidak ditemukan baris guru valid di file.');
        return;
      }
      const stats = onImportTeachersBatch(parsed);
      setImportModalOpen(false);
      onToast('success', `Impor "${file.name}": ${stats.added} ditambah, ${stats.updated} diperbarui.`);
    } catch {
      onToast('error', 'Gagal membaca file Excel/CSV.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="space-y-6">
      {/* Compact Header Action Bar */}
      <section className="bg-white border border-slate-200 rounded-xl px-5 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <UserCheck className="w-5 h-5 text-indigo-600 shrink-0" />
          <div>
            <h2 className="text-base font-bold text-slate-900">
              Manajemen Data Guru ({teachers.length})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Direktori pendidik, wali kelas, dan guru BK untuk referensi jadwal serta data siswa
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
            <span>Impor Massal Data Guru</span>
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
            <span>Tambah Data Guru</span>
          </button>
        </div>
      </section>

      {/* Tabel Data Guru */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Daftar Guru Terdaftar ({filteredTeachers.length})
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Pilih guru untuk hapus massal, ekspor XLSX/XLS/Word/PDF, atau cetak resmi
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <select
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value)}
                aria-label="Filter status guru"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="ALL">Semua Status</option>
                <option value="Aktif">Aktif</option>
                <option value="Cuti">Cuti</option>
                <option value="Nonaktif">Nonaktif</option>
              </select>

              <div className="relative">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Cari nama / NIP / mapel..."
                  className="pl-9 pr-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 w-48 sm:w-56"
                />
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedIds.size > 0
                ? `${selectedIds.size} guru dipilih`
                : 'Centang baris untuk ekspor selektif / hapus massal'}
            </span>
            <TableExportToolbar
              selectedCount={selectedIds.size}
              onDeleteSelected={handleBulkDelete}
              onDownloadTemplate={handleDownloadTemplate}
              onExportXlsx={() => {
                const { headers, rows } = buildTeacherMatrix();
                exportRowsToXlsx('Data Guru', headers, rows, `Data_Guru_E-Guru_${getTodayISO()}.xlsx`);
                onToast('success', 'Data guru berhasil diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildTeacherMatrix();
                exportRowsToXls('Data Guru', headers, rows, `Data_Guru_E-Guru_${getTodayISO()}.xls`);
                onToast('success', 'Data guru berhasil diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildTeacherMatrix();
                exportDatasetToWordDoc({
                  db,
                  title: 'DAFTAR PENDIDIK & TENAGA KEPENDIDIKAN',
                  subtitle: `${settings.schoolName} · Total ${exportSourceTeachers.length} Guru`,
                  headers,
                  rows,
                  filename: `Data_Guru_E-Guru_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Data Guru diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildTeacherMatrix();
                exportDatasetToPdf({
                  db,
                  title: 'DAFTAR PENDIDIK & TENAGA KEPENDIDIKAN',
                  subtitle: `${settings.schoolName} · Total ${exportSourceTeachers.length} Guru`,
                  headers,
                  rows,
                  filename: `Data_Guru_E-Guru_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Data Guru diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildTeacherMatrix();
                printDatasetExact({
                  db,
                  title: 'DAFTAR PENDIDIK & TENAGA KEPENDIDIKAN',
                  subtitle: `${settings.schoolName} · Total ${exportSourceTeachers.length} Guru`,
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
                      filteredTeachers.length > 0 &&
                      filteredTeachers.every((t) => selectedIds.has(t.id))
                    }
                    onChange={toggleSelectAll}
                    aria-label="Pilih semua guru"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-3 py-3 w-10 text-center">No</th>
                <th className="px-3 py-3 whitespace-nowrap">Foto</th>
                <th className="px-3 py-3 whitespace-nowrap">NIP</th>
                <th className="px-3 py-3 min-w-[200px]">Nama Lengkap</th>
                <th className="px-3 py-3 whitespace-nowrap">L/P</th>
                <th className="px-3 py-3 whitespace-nowrap">Jabatan</th>
                <th className="px-3 py-3 whitespace-nowrap">Mata Pelajaran</th>
                <th className="px-3 py-3 whitespace-nowrap">Kepegawaian</th>
                <th className="px-3 py-3 whitespace-nowrap">Status</th>
                <th className="px-3 py-3 whitespace-nowrap">Email & No HP</th>
                <th className="px-3 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {filteredTeachers.length === 0 ? (
                <tr>
                  <td colSpan={12} className="px-4 py-10 text-center text-sm text-slate-500">
                    Belum ada data guru yang sesuai filter.
                  </td>
                </tr>
              ) : (
                filteredTeachers.map((t, idx) => {
                  const checked = selectedIds.has(t.id);
                  return (
                    <tr
                      key={t.id}
                      className={`hover:bg-slate-50/80 transition-colors ${checked ? 'bg-indigo-50/40' : ''}`}
                    >
                      <td className="px-3 py-2.5 text-center">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectRow(t.id)}
                          aria-label={`Pilih ${t.name}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-3 py-2.5 text-center text-xs text-slate-500 tabular-nums">
                        {idx + 1}
                      </td>
                      <td className="px-3 py-2.5">
                        {t.photo ? (
                          <img
                            src={t.photo}
                            alt={t.name}
                            className="w-8 h-8 rounded-full object-cover border border-slate-200"
                          />
                        ) : (
                          <div className="w-8 h-8 rounded-full bg-indigo-100 text-indigo-700 text-xs font-bold flex items-center justify-center">
                            {t.name.slice(0, 2).toUpperCase()}
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2.5 font-mono-tabular text-xs text-slate-600 whitespace-nowrap">
                        {t.nip || '-'}
                      </td>
                      <td className="px-3 py-2.5 font-medium text-slate-900 whitespace-nowrap">
                        {t.name}
                      </td>
                      <td className="px-3 py-2.5 text-xs font-semibold">{t.gender}</td>
                      <td className="px-3 py-2.5 text-xs whitespace-nowrap">{t.position}</td>
                      <td className="px-3 py-2.5 text-xs whitespace-nowrap">{t.subject}</td>
                      <td className="px-3 py-2.5 text-xs whitespace-nowrap">{t.employmentStatus}</td>
                      <td className="px-3 py-2.5 text-xs font-semibold whitespace-nowrap">
                        <span
                          className={
                            t.status === 'Aktif'
                              ? 'text-emerald-700'
                              : t.status === 'Cuti'
                              ? 'text-amber-700'
                              : 'text-rose-700'
                          }
                        >
                          {t.status}
                        </span>
                      </td>
                      <td className="px-3 py-2.5 text-xs whitespace-nowrap">
                        <p className="text-slate-700">{t.email || '-'}</p>
                        <p className="text-slate-500 font-mono-tabular">{t.phone || '-'}</p>
                      </td>
                      <td className="px-3 py-2.5 whitespace-nowrap text-right">
                        <div className="inline-flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => handleEdit(t)}
                            className="px-2.5 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 inline-flex items-center gap-1"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                            Edit
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDelete(t)}
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
        title={editingId ? 'Edit Data Guru' : 'Tambah Data Guru'}
        subtitle="Direktori pendidik, wali kelas, dan guru BK untuk referensi jadwal serta data siswa"
        maxWidthClass="max-w-4xl"
        onClose={() => {
          setFormModalOpen(false);
          resetForm();
        }}
      >
        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label htmlFor="tch-nip" className="block text-xs font-semibold text-slate-700 mb-1.5">
                NIP / NUPTK
              </label>
              <input
                id="tch-nip"
                type="text"
                value={nip}
                onChange={(e) => setNip(e.target.value)}
                placeholder="19890418 201502 1 003"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div className="sm:col-span-2">
              <label htmlFor="tch-name" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Nama Lengkap & Gelar <span className="text-rose-600">*</span>
              </label>
              <input
                id="tch-name"
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Contoh: Dra. Siti Aminah, M.Pd."
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="tch-gender" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Jenis Kelamin
              </label>
              <select
                id="tch-gender"
                value={gender}
                onChange={(e) => setGender(e.target.value as 'L' | 'P')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="L">Laki-laki (L)</option>
                <option value="P">Perempuan (P)</option>
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label htmlFor="tch-pos" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Jabatan / Tugas Tambahan
              </label>
              <input
                id="tch-pos"
                type="text"
                value={position}
                onChange={(e) => setPosition(e.target.value)}
                placeholder="Guru Mapel / Wali Kelas / Guru BK"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="tch-subj" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Mata Pelajaran Diampu
              </label>
              <input
                id="tch-subj"
                list="teacher-subject-list"
                value={subject}
                onChange={(e) => setSubject(e.target.value)}
                placeholder="Mata pelajaran"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
              <datalist id="teacher-subject-list">
                {settings.masterSubjects.map((s) => (
                  <option key={s} value={s} />
                ))}
                <option value="Bimbingan Konseling" />
              </datalist>
            </div>

            <div>
              <label htmlFor="tch-emp" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Status Kepegawaian
              </label>
              <select
                id="tch-emp"
                value={employmentStatus}
                onChange={(e) => setEmploymentStatus(e.target.value as TeacherEmploymentStatus)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                {EMPLOYMENT_TYPES.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label htmlFor="tch-status" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Status Aktif
              </label>
              <select
                id="tch-status"
                value={status}
                onChange={(e) => setStatus(e.target.value as 'Aktif' | 'Cuti' | 'Nonaktif')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="Aktif">Aktif</option>
                <option value="Cuti">Cuti</option>
                <option value="Nonaktif">Nonaktif</option>
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 items-end">
            <div>
              <label htmlFor="tch-email" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Email
              </label>
              <input
                id="tch-email"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="guru@sekolah.sch.id"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="tch-phone" className="block text-xs font-semibold text-slate-700 mb-1.5">
                No HP / WhatsApp
              </label>
              <input
                id="tch-phone"
                type="text"
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                placeholder="0812..."
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="tch-addr" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Alamat
              </label>
              <input
                id="tch-addr"
                type="text"
                value={address}
                onChange={(e) => setAddress(e.target.value)}
                placeholder="Alamat domisili"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div className="flex items-center gap-2">
              <label className="cursor-pointer flex-1 px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 inline-flex items-center justify-center gap-1.5">
                <ImagePlus className="w-4 h-4 text-indigo-600" />
                {photo ? 'Ganti Foto' : 'Unggah Foto'}
                <input type="file" accept="image/*" className="hidden" onChange={handlePhotoUpload} />
              </label>
              {photo && (
                <button
                  type="button"
                  onClick={() => setPhoto('')}
                  className="px-2.5 py-2 rounded-lg border border-rose-200 text-xs text-rose-600 hover:bg-rose-50"
                >
                  Hapus
                </button>
              )}
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
              <UserCheck className="w-4 h-4" />
              {editingId ? 'Simpan Perubahan' : 'Tambahkan Guru'}
            </button>
          </div>
        </form>
      </FormModal>

      <FormModal
        open={importModalOpen}
        title="Impor Massal Data Guru (Excel / CSV / Tempel)"
        subtitle="Urutan kolom: NIP, Nama Lengkap, L/P, Jabatan, Mata Pelajaran, Kepegawaian, Status, Email, No HP, Alamat"
        maxWidthClass="max-w-2xl"
        onClose={() => setImportModalOpen(false)}
      >
        <div className="p-5 space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <button
              type="button"
              onClick={handleDownloadTemplate}
              className="px-3.5 py-2 rounded-lg border border-slate-200 hover:bg-slate-50 text-xs font-semibold text-slate-700"
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
                id="teacher-file-upload"
              />
              <label
                htmlFor="teacher-file-upload"
                className="cursor-pointer px-3.5 py-2 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-xs font-semibold text-indigo-700 inline-flex items-center gap-1.5"
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
            placeholder="Tempel baris data guru dari Excel (Tab / CSV)..."
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
