import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ImagePlus,
  Pencil,
  Plus,
  Search,
  Trash2,
  Upload,
  UserPlus,
  Users,
} from 'lucide-react';
import { PortalDatabase, Student, StudentStatus } from '../types';
import {
  exportDatasetToPdf,
  exportDatasetToWordDoc,
  exportRowsToXls,
  exportRowsToXlsx,
  printDatasetExact,
} from '../utils/exportHelpers';
import { formatShortDateId, getTodayISO } from '../utils/portalHelpers';
import { ConfirmDialogState } from './FeedbackUI';
import { FormModal } from './FormModal';
import { StudentDetailModal } from './StudentDetailModal';
import { TableExportToolbar } from './TableExportToolbar';

interface StudentsViewProps {
  db: PortalDatabase;
  onUpsertStudent: (student: Student) => void;
  onImportStudentsBatch: (incoming: Omit<Student, 'id'>[]) => { added: number; updated: number };
  onDeleteStudent: (id: string) => void;
  onDeleteStudentsBatch: (ids: string[]) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const STUDENT_STATUS_OPTIONS: StudentStatus[] = ['Aktif', 'Pindah', 'Lulus', 'Nonaktif'];

const EXCEL_HEADERS_16 = [
  'NIS',
  'NISN',
  'Nama',
  'Kelas',
  'No Absen',
  'Tempat',
  'Tanggal Lahir',
  'Email',
  'Foto',
  'ALAMAT',
  'NO_HP',
  'JENIS_KELAMIN',
  'Status',
  'Wali Kelas',
  'Guru BK',
  'Tahun Masuk',
];

export function StudentsView({
  db,
  onUpsertStudent,
  onImportStudentsBatch,
  onDeleteStudent,
  onDeleteStudentsBatch,
  onToast,
  onConfirm,
}: StudentsViewProps) {
  const { students, teachers, settings } = db;

  const classOptions = useMemo(() => {
    const set = new Set(settings.masterClasses);
    students.forEach((s) => {
      if (s.className) set.add(s.className);
    });
    return Array.from(set);
  }, [settings.masterClasses, students]);

  const teacherNames = useMemo(() => {
    const set = new Set<string>();
    if (settings.teacherName) set.add(settings.teacherName);
    teachers.forEach((t) => {
      if (t.name) set.add(t.name);
    });
    return Array.from(set);
  }, [teachers, settings.teacherName]);

  // Add/Edit Student Form (16 fields)
  const [editingId, setEditingId] = useState<string | null>(null);
  const [nis, setNis] = useState('');
  const [nisn, setNisn] = useState('');
  const [name, setName] = useState('');
  const [className, setClassName] = useState(classOptions[0] || '');
  const [absentNo, setAbsentNo] = useState('');
  const [birthPlace, setBirthPlace] = useState('');
  const [birthDate, setBirthDate] = useState('');
  const [email, setEmail] = useState('');
  const [photo, setPhoto] = useState('');
  const [address, setAddress] = useState('');
  const [phone, setPhone] = useState('');
  const [gender, setGender] = useState<'L' | 'P'>('L');
  const [status, setStatus] = useState<StudentStatus>('Aktif');
  const [homeroomTeacher, setHomeroomTeacher] = useState('');
  const [counselorTeacher, setCounselorTeacher] = useState('');
  const [entryYear, setEntryYear] = useState('2025');

  // Import Box
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Table Filters & Multi-Select
  const [searchQuery, setSearchQuery] = useState('');
  const [classFilter, setClassFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [detailStudent, setDetailStudent] = useState<Student | null>(null);

  const resetForm = () => {
    setEditingId(null);
    setNis('');
    setNisn('');
    setName('');
    setClassName(classOptions[0] || '');
    setAbsentNo('');
    setBirthPlace('');
    setBirthDate('');
    setEmail('');
    setPhoto('');
    setAddress('');
    setPhone('');
    setGender('L');
    setStatus('Aktif');
    setHomeroomTeacher('');
    setCounselorTeacher('');
    setEntryYear('2025');
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran file foto siswa maksimal 2 MB.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        setPhoto(reader.result);
      }
    };
    reader.readAsDataURL(file);
  };

  const handleSubmitStudent = (e: React.FormEvent) => {
    e.preventDefault();
    if (!nis.trim() || !name.trim() || !className.trim()) {
      onToast('error', 'NIS, Nama Lengkap, dan Kelas wajib diisi.');
      return;
    }

    onUpsertStudent({
      id: editingId ?? `st-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      nis: nis.trim(),
      nisn: nisn.trim(),
      name: name.trim(),
      className: className.trim(),
      absentNo: absentNo.trim(),
      birthPlace: birthPlace.trim(),
      birthDate: birthDate.trim(),
      email: email.trim(),
      photo: photo.trim(),
      address: address.trim(),
      phone: phone.trim(),
      gender,
      status,
      homeroomTeacher: homeroomTeacher.trim(),
      counselorTeacher: counselorTeacher.trim(),
      entryYear: entryYear.trim(),
    });
    onToast('success', editingId ? 'Data siswa berhasil diperbarui.' : 'Siswa baru berhasil ditambahkan.');
    resetForm();
    setFormModalOpen(false);
  };

  const handleEditClick = (st: Student) => {
    setEditingId(st.id);
    setNis(st.nis);
    setNisn(st.nisn || '');
    setName(st.name);
    setClassName(st.className);
    setAbsentNo(st.absentNo || '');
    setBirthPlace(st.birthPlace || '');
    setBirthDate(st.birthDate || '');
    setEmail(st.email || '');
    setPhoto(st.photo || '');
    setAddress(st.address || '');
    setPhone(st.phone || '');
    setGender(st.gender);
    setStatus(st.status || 'Aktif');
    setHomeroomTeacher(st.homeroomTeacher || '');
    setCounselorTeacher(st.counselorTeacher || '');
    setEntryYear(st.entryYear || '');
    setFormModalOpen(true);
  };

  const handleDeleteClick = (st: Student) => {
    onConfirm({
      title: 'Hapus Data Siswa',
      description: `Hapus siswa "${st.name}" (NIS ${st.nis}) dari kelas ${st.className}? Nilai terkait siswa ini juga akan dihapus.`,
      confirmLabel: 'Hapus Siswa',
      danger: true,
      onConfirm: () => {
        onDeleteStudent(st.id);
        onToast('success', `Siswa ${st.name} dihapus.`);
      },
    });
  };

  // Download Excel Template (16 columns)
  const handleDownloadTemplate = () => {
    const rows = [
      EXCEL_HEADERS_16,
      [
        '25101',
        '0091234599',
        'Ahmad Fauzan',
        classOptions[0] || 'X-A',
        '01',
        'Kudus',
        '2009-05-14',
        'ahmad.25101@siswa.sch.id',
        '',
        'Jl. Pramuka No. 10, Kudus',
        '085712349901',
        'L',
        'Aktif',
        teacherNames[0] || 'Dra. Siti Aminah, M.Pd.',
        'Ratna Kusumawati, S.Psi.',
        '2025',
      ],
      [
        '25102',
        '0091234598',
        'Siti Nurhaliza',
        classOptions[0] || 'X-A',
        '02',
        'Jepara',
        '2009-08-21',
        'siti.25102@siswa.sch.id',
        '',
        'Desa Rendeng RT 02/01, Kudus',
        '085712349902',
        'P',
        'Aktif',
        teacherNames[0] || 'Dra. Siti Aminah, M.Pd.',
        'Ratna Kusumawati, S.Psi.',
        '2025',
      ],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 10 },
      { wch: 14 },
      { wch: 26 },
      { wch: 12 },
      { wch: 10 },
      { wch: 14 },
      { wch: 14 },
      { wch: 24 },
      { wch: 14 },
      { wch: 28 },
      { wch: 15 },
      { wch: 14 },
      { wch: 10 },
      { wch: 24 },
      { wch: 24 },
      { wch: 12 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template Siswa');
    XLSX.writeFile(wb, 'Template_Impor_Siswa_Lengkap_E-Guru.xlsx');
    onToast('success', 'Template Excel (.xlsx) 16 kolom berhasil diunduh.');
  };

  const filteredStudents = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return students
      .filter((s) => {
        if (classFilter !== 'ALL' && s.className !== classFilter) return false;
        if (statusFilter !== 'ALL' && (s.status || 'Aktif') !== statusFilter) return false;
        if (q) {
          return (
            s.name.toLowerCase().includes(q) ||
            s.nis.toLowerCase().includes(q) ||
            (s.nisn || '').toLowerCase().includes(q) ||
            s.className.toLowerCase().includes(q) ||
            (s.email || '').toLowerCase().includes(q) ||
            (s.homeroomTeacher || '').toLowerCase().includes(q)
          );
        }
        return true;
      })
      .sort(
        (a, b) =>
          a.className.localeCompare(b.className) ||
          (a.absentNo || '').localeCompare(b.absentNo || '') ||
          a.name.localeCompare(b.name)
      );
  }, [students, classFilter, statusFilter, searchQuery]);

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
      filteredStudents.length > 0 &&
      filteredStudents.every((s) => selectedIds.has(s.id))
    ) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredStudents.map((s) => s.id)));
    }
  };

  const handleBulkDelete = () => {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    onConfirm({
      title: 'Hapus Siswa Terpilih',
      description: `Hapus ${count} data siswa yang dipilih beserta data nilai terkait?`,
      confirmLabel: `Hapus ${count} Siswa`,
      danger: true,
      onConfirm: () => {
        onDeleteStudentsBatch(Array.from(selectedIds));
        setSelectedIds(new Set());
        onToast('success', `${count} siswa berhasil dihapus.`);
      },
    });
  };

  const exportSourceStudents = useMemo(() => {
    if (selectedIds.size > 0) {
      const chosen = filteredStudents.filter((s) => selectedIds.has(s.id));
      if (chosen.length > 0) return chosen;
    }
    return filteredStudents;
  }, [filteredStudents, selectedIds]);

  const buildStudentRows16 = () =>
    exportSourceStudents.map((s) => [
      s.nis,
      s.nisn || '',
      s.name,
      s.className,
      s.absentNo || '',
      s.birthPlace || '',
      s.birthDate || '',
      s.email || '',
      s.photo && s.photo.startsWith('http') ? s.photo : '',
      s.address || '',
      s.phone || '',
      s.gender,
      s.status || 'Aktif',
      s.homeroomTeacher || '',
      s.counselorTeacher || '',
      s.entryYear || '',
    ]);

  const buildStudentPrintMatrix = () => ({
    headers: ['No', 'NIS', 'NISN', 'Nama Siswa', 'Kelas', 'No Absen', 'L/P', 'Status', 'No HP', 'Wali Kelas', 'Guru BK', 'Thn Masuk'],
    rows: exportSourceStudents.map((s, idx) => [
      idx + 1,
      s.nis,
      s.nisn || '-',
      s.name,
      s.className,
      s.absentNo || '-',
      s.gender,
      s.status || 'Aktif',
      s.phone || '-',
      s.homeroomTeacher || '-',
      s.counselorTeacher || '-',
      s.entryYear || '-',
    ]),
  });

  // Export Students to Excel (16 columns)
  const handleExportXlsx = () => {
    if (exportSourceStudents.length === 0) {
      onToast('error', 'Tidak ada data siswa untuk diekspor.');
      return;
    }
    exportRowsToXlsx(
      'Data Siswa',
      EXCEL_HEADERS_16,
      buildStudentRows16(),
      `Data_Siswa_Lengkap_E-Guru_${getTodayISO()}.xlsx`
    );
    onToast('success', `${exportSourceStudents.length} data siswa berhasil diekspor ke XLSX.`);
  };

  const handleExportXls = () => {
    if (exportSourceStudents.length === 0) {
      onToast('error', 'Tidak ada data siswa untuk diekspor.');
      return;
    }
    exportRowsToXls(
      'Data Siswa',
      EXCEL_HEADERS_16,
      buildStudentRows16(),
      `Data_Siswa_Lengkap_E-Guru_${getTodayISO()}.xls`
    );
    onToast('success', `${exportSourceStudents.length} data siswa berhasil diekspor ke XLS.`);
  };

  // Parse pasted text or Excel rows (supports both 16-column and compact 4-column formats)
  const parseRawRowsToStudents = (rawRows: string[][]): Omit<Student, 'id'>[] => {
    const result: Omit<Student, 'id'>[] = [];
    rawRows.forEach((cols, index) => {
      if (cols.length < 2) return;
      const clean = (idx: number) => String(cols[idx] ?? '').trim();

      if (
        index === 0 &&
        (clean(0).toLowerCase() === 'nis' ||
          clean(1).toLowerCase() === 'nisn' ||
          clean(2).toLowerCase() === 'nama')
      ) {
        return;
      }

      // Compact 4-column fallback: NIS | Nama | L/P | Kelas
      if (cols.length <= 5) {
        const c0 = clean(0);
        const c1 = clean(1);
        const c2 = clean(2).toUpperCase();
        const c3 = clean(3);
        if (!c1 && !c0) return;
        result.push({
          nis: c0,
          nisn: '',
          name: c1 || c0,
          className: c3 || classOptions[0] || 'X-A',
          absentNo: '',
          birthPlace: '',
          birthDate: '',
          email: '',
          photo: '',
          address: '',
          phone: '',
          gender: c2.startsWith('P') ? 'P' : 'L',
          status: 'Aktif',
          homeroomTeacher: '',
          counselorTeacher: '',
          entryYear: '2025',
        });
        return;
      }

      // Full 16-column format
      const sNis = clean(0);
      const sNisn = clean(1);
      const sName = clean(2);
      const sClass = clean(3) || classOptions[0] || 'X-A';
      if (!sName && !sNis) return;

      const rawGender = clean(11).toUpperCase();
      const rawStatus = clean(12);
      const validStatus: StudentStatus = STUDENT_STATUS_OPTIONS.includes(rawStatus as StudentStatus)
        ? (rawStatus as StudentStatus)
        : 'Aktif';

      result.push({
        nis: sNis,
        nisn: sNisn,
        name: sName || sNis,
        className: sClass,
        absentNo: clean(4),
        birthPlace: clean(5),
        birthDate: clean(6),
        email: clean(7),
        photo: clean(8),
        address: clean(9),
        phone: clean(10),
        gender: rawGender.startsWith('P') ? 'P' : 'L',
        status: validStatus,
        homeroomTeacher: clean(13),
        counselorTeacher: clean(14),
        entryYear: clean(15),
      });
    });
    return result;
  };

  const handleImportPaste = () => {
    const lines = pasteText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    if (lines.length === 0) {
      onToast('error', 'Tempelkan baris data siswa dari Excel/CSV terlebih dahulu.');
      return;
    }

    const matrix = lines.map((line) => {
      if (line.includes('\t')) return line.split('\t');
      if (line.includes(';')) return line.split(';');
      return line.split(',');
    });

    const parsed = parseRawRowsToStudents(matrix);
    if (parsed.length === 0) {
      onToast('error', 'Format data tidak dikenali. Gunakan template 16 kolom atau format 4 kolom.');
      return;
    }

    const stats = onImportStudentsBatch(parsed);
    setPasteText('');
    setImportModalOpen(false);
    onToast(
      'success',
      `Impor selesai: ${stats.added} siswa baru ditambahkan, ${stats.updated} data diperbarui.`
    );
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array' });
      const firstSheet = wb.SheetNames[0];
      if (!firstSheet) {
        onToast('error', 'File Excel/CSV tidak memiliki lembar kerja.');
        return;
      }
      const ws = wb.Sheets[firstSheet];
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' });
      const parsed = parseRawRowsToStudents(rows);
      if (parsed.length === 0) {
        onToast('error', 'Tidak ditemukan baris siswa yang valid di dalam file.');
        return;
      }
      const stats = onImportStudentsBatch(parsed);
      setImportModalOpen(false);
      onToast(
        'success',
        `Impor file "${file.name}" selesai: ${stats.added} ditambahkan, ${stats.updated} diperbarui.`
      );
    } catch {
      onToast('error', 'Gagal membaca file. Pastikan format .xlsx, .xls, atau .csv valid.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="space-y-6">
      {/* Compact Header Action Bar */}
      <section className="bg-white border border-slate-200 rounded-xl px-5 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <Users className="w-5 h-5 text-indigo-600 shrink-0" />
          <div>
            <h2 className="text-base font-bold text-slate-900">
              Manajemen Data Siswa ({students.length})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Klik nama siswa pada tabel untuk membuka detail informasi lengkap & riwayat belajar
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
            <span>Impor Massal & Template</span>
          </button>
          <button
            type="button"
            onClick={() => {
              resetForm();
              setFormModalOpen(true);
            }}
            className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shadow-xs"
          >
            <UserPlus className="w-4 h-4" />
            <span>Tambah Siswa Baru</span>
          </button>
        </div>
      </section>

      {/* Tabel Data Siswa Lengkap (max-h-[420px], sticky header) */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 space-y-3">
          <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Daftar Siswa Terdaftar ({filteredStudents.length})
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Pilih siswa untuk hapus massal, ekspor Excel/Word/PDF, atau cetak resmi
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2.5">
              <select
                value={classFilter}
                onChange={(e) => setClassFilter(e.target.value)}
                aria-label="Filter data siswa menurut kelas"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="ALL">Semua Kelas</option>
                {classOptions.map((c) => (
                  <option key={c} value={c}>
                    Kelas {c}
                  </option>
                ))}
              </select>

              <select
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value)}
                aria-label="Filter status siswa"
                className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="ALL">Semua Status</option>
                {STUDENT_STATUS_OPTIONS.map((stOpt) => (
                  <option key={stOpt} value={stOpt}>
                    {stOpt}
                  </option>
                ))}
              </select>

              <div className="relative">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Cari NIS / NISN / Nama..."
                  className="pl-9 pr-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 w-48 sm:w-56"
                />
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedIds.size > 0
                ? `${selectedIds.size} siswa dipilih`
                : 'Centang baris untuk hapus/ekspor selektif, atau ekspor semua hasil filter'}
            </span>
            <TableExportToolbar
              selectedCount={selectedIds.size}
              onDeleteSelected={handleBulkDelete}
              onDownloadTemplate={handleDownloadTemplate}
              onExportXlsx={handleExportXlsx}
              onExportXls={handleExportXls}
              onExportWord={(cfg) => {
                const { headers, rows } = buildStudentPrintMatrix();
                exportDatasetToWordDoc({
                  db,
                  title: 'DAFTAR INDUK DATA SISWA',
                  subtitle: `${settings.schoolName} · Total ${exportSourceStudents.length} Siswa`,
                  headers,
                  rows,
                  filename: `Data_Siswa_E-Guru_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Data Siswa diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildStudentPrintMatrix();
                exportDatasetToPdf({
                  db,
                  title: 'DAFTAR INDUK DATA SISWA',
                  subtitle: `${settings.schoolName} · Total ${exportSourceStudents.length} Siswa`,
                  headers,
                  rows,
                  filename: `Data_Siswa_E-Guru_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Data Siswa diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildStudentPrintMatrix();
                printDatasetExact({
                  db,
                  title: 'DAFTAR INDUK DATA SISWA',
                  subtitle: `${settings.schoolName} · Total ${exportSourceStudents.length} Siswa`,
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
          <table className="w-full text-left border-collapse text-xs sm:text-sm">
            <thead className="sticky top-0 z-20 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="sticky left-0 z-30 bg-slate-100 px-3 py-3 w-10 text-center border-r border-slate-200">
                  <input
                    type="checkbox"
                    checked={
                      filteredStudents.length > 0 &&
                      filteredStudents.every((s) => selectedIds.has(s.id))
                    }
                    onChange={toggleSelectAll}
                    aria-label="Pilih semua siswa"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-3 py-3 w-10 text-center">No</th>
                <th className="px-3 py-3 whitespace-nowrap">Foto</th>
                <th className="px-3 py-3 whitespace-nowrap">NIS</th>
                <th className="px-3 py-3 whitespace-nowrap">NISN</th>
                <th className="sticky left-10 z-30 bg-slate-100 px-3 py-3 min-w-[180px] border-r border-slate-200">
                  Nama
                </th>
                <th className="px-3 py-3 whitespace-nowrap">Kelas</th>
                <th className="px-3 py-3 whitespace-nowrap text-center">No Absen</th>
                <th className="px-3 py-3 whitespace-nowrap">Tempat</th>
                <th className="px-3 py-3 whitespace-nowrap">Tanggal Lahir</th>
                <th className="px-3 py-3 whitespace-nowrap">Email</th>
                <th className="px-3 py-3 min-w-[200px]">ALAMAT</th>
                <th className="px-3 py-3 whitespace-nowrap">NO_HP</th>
                <th className="px-3 py-3 whitespace-nowrap text-center">JENIS_KELAMIN</th>
                <th className="px-3 py-3 whitespace-nowrap">Status</th>
                <th className="px-3 py-3 whitespace-nowrap">Wali Kelas</th>
                <th className="px-3 py-3 whitespace-nowrap">Guru BK</th>
                <th className="px-3 py-3 whitespace-nowrap text-center">Tahun Masuk</th>
                <th className="px-3 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {filteredStudents.length === 0 ? (
                <tr>
                  <td colSpan={19} className="px-4 py-10 text-center text-sm text-slate-500">
                    Belum ada data siswa yang sesuai filter. Tambahkan siswa baru atau impor dari Excel.
                  </td>
                </tr>
              ) : (
                filteredStudents.map((st, idx) => {
                  const checked = selectedIds.has(st.id);
                  return (
                  <tr key={st.id} className={`hover:bg-slate-50/80 transition-colors ${checked ? 'bg-indigo-50/40' : ''}`}>
                    <td className="sticky left-0 z-10 bg-white px-3 py-2.5 text-center border-r border-slate-200">
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => toggleSelectRow(st.id)}
                        aria-label={`Pilih ${st.name}`}
                        className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                      />
                    </td>
                    <td className="px-3 py-2.5 text-center text-xs text-slate-500 tabular-nums">
                      {idx + 1}
                    </td>
                    <td className="px-3 py-2.5">
                      {st.photo ? (
                        <img
                          src={st.photo}
                          alt={st.name}
                          className="w-8 h-8 rounded-full object-cover border border-slate-200"
                        />
                      ) : (
                        <div className="w-8 h-8 rounded-full bg-slate-200 text-slate-700 text-xs font-bold flex items-center justify-center">
                          {st.name.slice(0, 2).toUpperCase()}
                        </div>
                      )}
                    </td>
                    <td className="px-3 py-2.5 font-mono-tabular text-xs text-slate-700 whitespace-nowrap">
                      {st.nis}
                    </td>
                    <td className="px-3 py-2.5 font-mono-tabular text-xs text-slate-600 whitespace-nowrap">
                      {st.nisn || '-'}
                    </td>
                    <td className="sticky left-10 z-10 bg-white px-3 py-2.5 font-medium text-slate-900 whitespace-nowrap border-r border-slate-200">
                      <button
                        type="button"
                        onClick={() => setDetailStudent(st)}
                        title="Klik untuk melihat detail informasi siswa"
                        className="text-left font-medium text-slate-900 hover:text-indigo-600 hover:underline cursor-pointer"
                      >
                        {st.name}
                      </button>
                    </td>
                    <td className="px-3 py-2.5 font-semibold text-indigo-700 whitespace-nowrap">
                      {st.className}
                    </td>
                    <td className="px-3 py-2.5 text-center font-mono-tabular text-xs">
                      {st.absentNo || '-'}
                    </td>
                    <td className="px-3 py-2.5 text-xs whitespace-nowrap">{st.birthPlace || '-'}</td>
                    <td className="px-3 py-2.5 text-xs font-mono-tabular whitespace-nowrap">
                      {st.birthDate ? formatShortDateId(st.birthDate) : '-'}
                    </td>
                    <td className="px-3 py-2.5 text-xs whitespace-nowrap">{st.email || '-'}</td>
                    <td className="px-3 py-2.5 text-xs">{st.address || '-'}</td>
                    <td className="px-3 py-2.5 text-xs font-mono-tabular whitespace-nowrap">
                      {st.phone || '-'}
                    </td>
                    <td className="px-3 py-2.5 text-center text-xs font-semibold">{st.gender}</td>
                    <td className="px-3 py-2.5 text-xs font-semibold whitespace-nowrap">
                      <span
                        className={
                          (st.status || 'Aktif') === 'Aktif'
                            ? 'text-emerald-700'
                            : st.status === 'Lulus'
                            ? 'text-indigo-700'
                            : st.status === 'Pindah'
                            ? 'text-amber-700'
                            : 'text-rose-700'
                        }
                      >
                        {st.status || 'Aktif'}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-xs whitespace-nowrap">
                      {st.homeroomTeacher || '-'}
                    </td>
                    <td className="px-3 py-2.5 text-xs whitespace-nowrap">
                      {st.counselorTeacher || '-'}
                    </td>
                    <td className="px-3 py-2.5 text-center text-xs font-mono-tabular">
                      {st.entryYear || '-'}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap text-right">
                      <div className="inline-flex items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => handleEditClick(st)}
                          className="px-2.5 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 inline-flex items-center gap-1"
                        >
                          <Pencil className="w-3.5 h-3.5" />
                          Edit
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeleteClick(st)}
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
        title={editingId ? 'Edit Data Siswa Lengkap' : 'Tambah Siswa Baru'}
        subtitle="Lengkapi identitas NIS, NISN, biodata kelahiran, kontak, wali kelas, guru BK, dan foto siswa"
        maxWidthClass="max-w-5xl"
        onClose={() => {
          setFormModalOpen(false);
          resetForm();
        }}
      >
        <form onSubmit={handleSubmitStudent} className="p-5 space-y-5">
          {/* Row 1: NIS, NISN, Nama, Kelas */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            <div>
              <label htmlFor="st-nis" className="block text-xs font-semibold text-slate-700 mb-1.5">
                NIS <span className="text-rose-600">*</span>
              </label>
              <input
                id="st-nis"
                type="text"
                value={nis}
                onChange={(e) => setNis(e.target.value)}
                placeholder="25001"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-nisn" className="block text-xs font-semibold text-slate-700 mb-1.5">
                NISN
              </label>
              <input
                id="st-nisn"
                type="text"
                value={nisn}
                onChange={(e) => setNisn(e.target.value)}
                placeholder="0091234501"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div className="lg:col-span-2">
              <label htmlFor="st-name" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Nama Lengkap <span className="text-rose-600">*</span>
              </label>
              <input
                id="st-name"
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Nama lengkap sesuai ijazah/rapor"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-class" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Kelas <span className="text-rose-600">*</span>
              </label>
              <select
                id="st-class"
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

          {/* Row 2: No Absen, Jenis Kelamin, Tempat Lahir, Tanggal Lahir, Tahun Masuk */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            <div>
              <label htmlFor="st-absen" className="block text-xs font-semibold text-slate-700 mb-1.5">
                No Absen
              </label>
              <input
                id="st-absen"
                type="text"
                value={absentNo}
                onChange={(e) => setAbsentNo(e.target.value)}
                placeholder="01"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-gender" className="block text-xs font-semibold text-slate-700 mb-1.5">
                JENIS_KELAMIN
              </label>
              <select
                id="st-gender"
                value={gender}
                onChange={(e) => setGender(e.target.value as 'L' | 'P')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="L">Laki-laki (L)</option>
                <option value="P">Perempuan (P)</option>
              </select>
            </div>

            <div>
              <label htmlFor="st-bplace" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Tempat Lahir
              </label>
              <input
                id="st-bplace"
                type="text"
                value={birthPlace}
                onChange={(e) => setBirthPlace(e.target.value)}
                placeholder="Contoh: Kudus"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-bdate" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Tanggal Lahir
              </label>
              <input
                id="st-bdate"
                type="date"
                value={birthDate}
                onChange={(e) => setBirthDate(e.target.value)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-entry" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Tahun Masuk
              </label>
              <input
                id="st-entry"
                type="text"
                value={entryYear}
                onChange={(e) => setEntryYear(e.target.value)}
                placeholder="2025"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>
          </div>

          {/* Row 3: Email, NO_HP, ALAMAT, Status */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            <div>
              <label htmlFor="st-email" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Email
              </label>
              <input
                id="st-email"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="siswa@sekolah.sch.id"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-phone" className="block text-xs font-semibold text-slate-700 mb-1.5">
                NO_HP
              </label>
              <input
                id="st-phone"
                type="text"
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                placeholder="0857..."
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div className="lg:col-span-2">
              <label htmlFor="st-address" className="block text-xs font-semibold text-slate-700 mb-1.5">
                ALAMAT
              </label>
              <input
                id="st-address"
                type="text"
                value={address}
                onChange={(e) => setAddress(e.target.value)}
                placeholder="Alamat tempat tinggal siswa"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="st-status" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Status
              </label>
              <select
                id="st-status"
                value={status}
                onChange={(e) => setStatus(e.target.value as StudentStatus)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                {STUDENT_STATUS_OPTIONS.map((opt) => (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                ))}
              </select>
            </div>
          </div>

          {/* Row 4: Wali Kelas, Guru BK, Foto */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 items-end">
            <div className="lg:col-span-2">
              <label htmlFor="st-wali" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Wali Kelas
              </label>
              <input
                id="st-wali"
                list="student-teacher-options"
                value={homeroomTeacher}
                onChange={(e) => setHomeroomTeacher(e.target.value)}
                placeholder="Pilih atau ketik nama Wali Kelas"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div className="lg:col-span-2">
              <label htmlFor="st-bk" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Guru BK
              </label>
              <input
                id="st-bk"
                list="student-teacher-options"
                value={counselorTeacher}
                onChange={(e) => setCounselorTeacher(e.target.value)}
                placeholder="Pilih atau ketik nama Guru BK"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <datalist id="student-teacher-options">
              {teacherNames.map((nameItem) => (
                <option key={nameItem} value={nameItem} />
              ))}
            </datalist>

            <div className="flex items-center gap-2">
              <label className="cursor-pointer flex-1 px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 inline-flex items-center justify-center gap-1.5">
                <ImagePlus className="w-4 h-4 text-indigo-600" />
                {photo ? 'Ganti Foto' : 'Foto Siswa'}
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
              Reset Form
            </button>
            <button
              type="submit"
              className="px-5 py-2.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
            >
              <UserPlus className="w-4 h-4" />
              {editingId ? 'Simpan Perubahan Siswa' : 'Tambahkan Siswa'}
            </button>
          </div>
        </form>
      </FormModal>

      <FormModal
        open={importModalOpen}
        title="Impor Massal & Template Data Siswa"
        subtitle="Mendukung format 16 kolom lengkap maupun format ringkas 4 kolom"
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
              Unduh Template 16 Kolom (.xlsx)
            </button>
            <div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleImportFile}
                className="hidden"
                id="student-file-upload"
              />
              <label
                htmlFor="student-file-upload"
                className="cursor-pointer px-3.5 py-2 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-xs font-semibold text-indigo-700 inline-flex items-center gap-2 transition-colors whitespace-nowrap"
              >
                <Upload className="w-3.5 h-3.5" />
                Unggah File CSV / Excel
              </label>
            </div>
          </div>

          <textarea
            id="paste-students"
            rows={5}
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            placeholder={'Tempel baris data dari Excel / Spreadsheet di sini (Tab-separated atau CSV)...'}
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
              className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
            >
              <Plus className="w-4 h-4" />
              Proses Impor Teks
            </button>
          </div>
        </div>
      </FormModal>

      <StudentDetailModal
        student={detailStudent}
        db={db}
        onClose={() => setDetailStudent(null)}
        onEdit={handleEditClick}
      />
    </div>
  );
}
