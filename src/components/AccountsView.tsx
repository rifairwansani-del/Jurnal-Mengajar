import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Eye,
  EyeOff,
  KeyRound,
  Lock,
  Pencil,
  PenLine,
  Plus,
  Power,
  Search,
  ShieldCheck,
  Stamp,
  Trash2,
  Upload,
  User,
  UserPlus,
} from 'lucide-react';
import { DEFAULT_ADMIN_EMAIL } from '../data/seedData';
import { AccountRole, PortalDatabase, UserAccount } from '../types';
import {
  exportDatasetToPdf,
  exportDatasetToWordDoc,
  exportRowsToXls,
  exportRowsToXlsx,
  printDatasetExact,
} from '../utils/exportHelpers';
import { compressImageFileToDataUrl, getTodayISO, normalizeEmailId } from '../utils/portalHelpers';
import { ConfirmDialogState } from './FeedbackUI';
import { FormModal } from './FormModal';
import { SignaturePadModal } from './SignaturePadModal';
import { TableExportToolbar } from './TableExportToolbar';

interface AccountsViewProps {
  db: PortalDatabase;
  onUpsertAccount: (
    account: UserAccount,
    previousId?: string,
    credentials?: { username?: string; password?: string }
  ) => Promise<{ ok: boolean; error?: string }>;
  onImportAccountsBatch: (
    incoming: Omit<UserAccount, 'id' | 'lastLogin'>[]
  ) => { added: number; updated: number };
  onDeleteAccount: (id: string) => void;
  onDeleteAccountsBatch: (ids: string[]) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const ROLE_OPTIONS: AccountRole[] = [
  'Admin',
  'Guru Mapel',
  'Wali Kelas',
  'Guru BK',
  'Kepala Sekolah',
];

const ACCOUNT_HEADERS = ['Nama Pengguna', 'Email Google', 'Peran', 'NIP', 'Status'];

export function AccountsView({
  db,
  onUpsertAccount,
  onImportAccountsBatch,
  onDeleteAccount,
  onDeleteAccountsBatch,
  onToast,
  onConfirm,
}: AccountsViewProps) {
  const { accounts, teachers, settings, session } = db;
  const currentEmailLower = normalizeEmailId(session.email || session.accountId || '');
  const isAdminSession =
    session.userRole === 'Admin' || currentEmailLower === DEFAULT_ADMIN_EMAIL;

  const [editingId, setEditingId] = useState<string | null>(null);
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [role, setRole] = useState<AccountRole | ''>('');
  const [nip, setNip] = useState('');
  const [status, setStatus] = useState<'Aktif' | 'Nonaktif' | ''>('');
  const [signatureImage, setSignatureImage] = useState('');
  const [schoolStampImage, setSchoolStampImage] = useState('');
  const [sigModalOpen, setSigModalOpen] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [roleFilter, setRoleFilter] = useState('');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);

  const editingAccount = editingId ? accounts.find((a) => a.id === editingId) : undefined;
  const isSelfLimitedEdit = Boolean(editingId && !isAdminSession);

  const resetForm = () => {
    setEditingId(null);
    setName('');
    setEmail('');
    setUsername('');
    setPassword('');
    setShowPassword(false);
    setRole('');
    setNip('');
    setStatus('');
    setSignatureImage('');
    setSchoolStampImage('');
  };

  const handleLinkFromTeacher = (teacherId: string) => {
    const t = teachers.find((item) => item.id === teacherId);
    if (!t) return;
    setName(t.name);
    setEmail(t.email || '');
    setNip(t.nip || '');
    const posLower = (t.position || '').trim().toLowerCase();
    if (posLower === 'kepala sekolah') setRole('Kepala Sekolah');
    else if (posLower === 'wali kelas') setRole('Wali Kelas');
    else if (posLower === 'guru bk' || posLower === 'bimbingan konseling') setRole('Guru BK');
    else if (posLower.includes('admin') || posLower.includes('operator')) setRole('Admin');
    else setRole('Guru Mapel');
    setStatus(t.status === 'Nonaktif' ? 'Nonaktif' : 'Aktif');
    if (!username && t.email && t.email.includes('@')) {
      const prefix = t.email.split('@')[0].toLowerCase().replace(/[^a-z0-9._-]/g, '');
      if (prefix.length >= 3) setUsername(prefix);
    }
  };

  const handleStampUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran stempel maksimal 2 MB.');
      return;
    }
    const compressed = await compressImageFileToDataUrl(file, 360, 200000);
    setSchoolStampImage(compressed);
    onToast('success', 'Stempel sekolah berhasil diunggah. Klik Simpan Akun untuk menerapkan.');
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim()) {
      onToast('error', 'Nama lengkap pengguna wajib diisi.');
      return;
    }

    const cleanUsername = username.trim().toLowerCase();
    const rawEmailLower = normalizeEmailId(email);

    if (!rawEmailLower && !cleanUsername) {
      onToast('error', 'Isi minimal Username Login atau Email Google.');
      return;
    }

    if (cleanUsername && !/^[a-z0-9._-]{3,40}$/.test(cleanUsername)) {
      onToast(
        'error',
        'Username harus 3–40 karakter (huruf kecil, angka, titik, garis bawah, atau tanda hubung).'
      );
      return;
    }

    const emailLower =
      rawEmailLower || (cleanUsername ? `${cleanUsername}@akun.portal` : '');

    if (!emailLower || !/^[^@]+@[^@]+\.[^@]+$/.test(emailLower)) {
      onToast('error', 'Alamat Email tidak valid.');
      return;
    }

    const existing = editingId ? accounts.find((a) => a.id === editingId) : undefined;

    if (password && password.length < 6) {
      onToast('error', 'Password minimal 6 karakter.');
      return;
    }

    if (!rawEmailLower && !existing?.hasPassword && !password) {
      onToast('error', 'Password minimal 6 karakter wajib diisi jika akun dibuat tanpa Email Google.');
      return;
    }

    if (cleanUsername && !existing?.hasPassword && !password) {
      onToast('error', 'Isi password (min. 6 karakter) untuk mengaktifkan login dengan username.');
      return;
    }

    const duplicateEmail = accounts.find(
      (a) => normalizeEmailId(a.email) === emailLower && a.id !== editingId
    );
    if (duplicateEmail) {
      onToast('error', `Email "${emailLower}" sudah digunakan oleh akun lain.`);
      return;
    }

    if (cleanUsername) {
      const duplicateUser = accounts.find(
        (a) => (a.username || '').trim().toLowerCase() === cleanUsername && a.id !== editingId
      );
      if (duplicateUser) {
        onToast('error', `Username "${cleanUsername}" sudah digunakan oleh akun lain.`);
        return;
      }
    }

    const credentialsPayload =
      cleanUsername || password
        ? {
            username: cleanUsername || existing?.username || '',
            password: password || undefined,
          }
        : undefined;

    if (!isSelfLimitedEdit && (!role || !status)) {
      onToast('error', 'Peran / Hak Akses dan Status Akun wajib dipilih.');
      return;
    }

    setSubmitting(true);
    try {
      const resolvedRole: AccountRole =
        isSelfLimitedEdit && existing ? existing.role : (role as AccountRole);
      const resolvedStatus: 'Aktif' | 'Nonaktif' =
        isSelfLimitedEdit && existing ? existing.status : (status as 'Aktif' | 'Nonaktif');
      const resolvedStamp =
        resolvedRole === 'Kepala Sekolah' ? schoolStampImage : '';
      const result = await onUpsertAccount(
        {
          id: emailLower,
          name: name.trim(),
          email: emailLower,
          username: cleanUsername || existing?.username || '',
          hasPassword: Boolean(password || existing?.hasPassword),
          role: resolvedRole,
          nip: nip.trim(),
          status: resolvedStatus,
          lastLogin: existing?.lastLogin ?? '-',
          signatureImage,
          schoolStampImage: resolvedStamp,
        },
        editingId && editingId !== emailLower ? editingId : undefined,
        credentialsPayload
      );

      if (!result.ok) {
        onToast('error', result.error || 'Gagal menyimpan data akun.');
        return;
      }

      onToast(
        'success',
        editingId ? 'Akun pengguna & kredensial berhasil diperbarui.' : 'Akun pengguna baru berhasil didaftarkan.'
      );
      resetForm();
      setFormModalOpen(false);
    } finally {
      setSubmitting(false);
    }
  };

  const handleEdit = (acc: UserAccount) => {
    const isSelf =
      normalizeEmailId(acc.email) === currentEmailLower ||
      (Boolean(session.accountId) && acc.id === session.accountId);
    if (!isAdminSession && !isSelf) {
      onToast('error', 'Hanya Admin yang dapat mengedit akun pengguna lain.');
      return;
    }
    setEditingId(acc.id);
    setName(acc.name);
    setEmail(acc.email);
    setUsername(acc.username || '');
    setPassword('');
    setShowPassword(false);
    setRole(acc.role);
    setNip(acc.nip);
    setStatus(acc.status);
    setSignatureImage(acc.signatureImage || '');
    setSchoolStampImage(
      acc.role === 'Kepala Sekolah'
        ? acc.schoolStampImage || settings.defaultSchoolStampImage || ''
        : ''
    );
    setFormModalOpen(true);
  };

  const handleToggleStatus = async (acc: UserAccount) => {
    if (!isAdminSession) {
      onToast('error', 'Hanya Admin yang dapat mengubah status akun.');
      return;
    }
    const emailLower = normalizeEmailId(acc.email);
    if (emailLower === DEFAULT_ADMIN_EMAIL) {
      onToast('error', 'Admin utama tidak dapat dinonaktifkan.');
      return;
    }
    if (emailLower === currentEmailLower && acc.status === 'Aktif') {
      onToast('error', 'Tidak dapat menonaktifkan akun yang sedang login saat ini.');
      return;
    }
    const nextStatus = acc.status === 'Aktif' ? 'Nonaktif' : 'Aktif';
    const result = await onUpsertAccount({ ...acc, status: nextStatus });
    if (!result.ok) {
      onToast('error', result.error || 'Gagal mengubah status akun.');
      return;
    }
    onToast('info', `Status akun ${acc.name} diubah menjadi ${nextStatus}.`);
  };

  const handleDelete = (acc: UserAccount) => {
    if (!isAdminSession) {
      onToast('error', 'Hanya Admin yang dapat menghapus akun.');
      return;
    }
    const emailLower = normalizeEmailId(acc.email);
    if (emailLower === DEFAULT_ADMIN_EMAIL) {
      onToast('error', 'Akun Admin utama tidak dapat dihapus.');
      return;
    }
    if (emailLower === currentEmailLower) {
      onToast('error', 'Tidak dapat menghapus akun yang sedang aktif login.');
      return;
    }
    onConfirm({
      title: 'Hapus Akun Pengguna',
      description: `Hapus akses login "${acc.name}" (${acc.username ? `@${acc.username}` : acc.email}) dari sistem?`,
      confirmLabel: 'Hapus Akun',
      danger: true,
      onConfirm: () => {
        onDeleteAccount(acc.id);
        setSelectedIds((prev) => {
          const next = new Set(prev);
          next.delete(acc.id);
          return next;
        });
        onToast('success', `Akun ${acc.name} berhasil dihapus.`);
      },
    });
  };

  const filteredAccounts = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return accounts.filter((a) => {
      if (roleFilter && roleFilter !== 'ALL' && a.role !== roleFilter) return false;
      if (q) {
        return (
          a.name.toLowerCase().includes(q) ||
          a.email.toLowerCase().includes(q) ||
          (a.username || '').toLowerCase().includes(q) ||
          a.nip.toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [accounts, roleFilter, searchQuery]);

  const toggleSelectRow = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (filteredAccounts.length > 0 && filteredAccounts.every((a) => selectedIds.has(a.id))) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(filteredAccounts.map((a) => a.id)));
    }
  };

  const handleBulkDelete = () => {
    if (selectedIds.size === 0) return;
    const count = selectedIds.size;
    onConfirm({
      title: 'Hapus Akun Terpilih',
      description: `Hapus ${count} akun yang dipilih? (Akun yang sedang aktif login dan Admin utama akan tetap dipertahankan).`,
      confirmLabel: `Hapus ${count} Akun`,
      danger: true,
      onConfirm: () => {
        onDeleteAccountsBatch(Array.from(selectedIds));
        setSelectedIds(new Set());
        onToast('success', 'Akun terpilih berhasil diproses.');
      },
    });
  };

  const exportSourceAccounts = useMemo(() => {
    if (selectedIds.size > 0) {
      const selected = filteredAccounts.filter((a) => selectedIds.has(a.id));
      if (selected.length > 0) return selected;
    }
    return filteredAccounts;
  }, [filteredAccounts, selectedIds]);

  const buildAccountMatrix = () => {
    const headers = ['No', 'Nama Pengguna', 'NIP', 'Username', 'Email Google', 'Peran', 'Status', 'TTD Digital', 'Stempel', 'Login Terakhir'];
    const rows = exportSourceAccounts.map((a, idx) => [
      idx + 1,
      a.name,
      a.nip || '-',
      a.username || '-',
      a.email,
      a.role,
      a.status,
      a.signatureImage ? 'Tersedia' : 'Belum Ada',
      a.role === 'Kepala Sekolah' && a.schoolStampImage ? 'Tersedia' : '-',
      a.lastLogin || '-',
    ]);
    return { headers, rows };
  };

  const handleDownloadTemplate = () => {
    exportRowsToXlsx(
      'Template Akun',
      ACCOUNT_HEADERS,
      [
        ['Budi Santoso, S.Pd.', 'budi@sma1kudus.sch.id', 'Guru Mapel', '19900101 201801 1 001', 'Aktif'],
      ],
      'Template_Impor_Akun_E-Guru.xlsx'
    );
    onToast('success', 'Template Excel akun berhasil diunduh.');
  };

  const parseRowsToAccounts = (rawRows: string[][]): Omit<UserAccount, 'id' | 'lastLogin'>[] => {
    const result: Omit<UserAccount, 'id' | 'lastLogin'>[] = [];
    rawRows.forEach((cols, index) => {
      if (cols.length < 2) return;
      const clean = (i: number) => String(cols[i] ?? '').trim();
      if (index === 0 && clean(0).toLowerCase().includes('nama')) return;
      const aName = clean(0);
      const aEmail = normalizeEmailId(clean(1));
      if (!aName || !aEmail || !aEmail.includes('@')) return;
      const rawRole = clean(2);
      const validRole: AccountRole = ROLE_OPTIONS.includes(rawRole as AccountRole)
        ? (rawRole as AccountRole)
        : 'Guru Mapel';
      result.push({
        name: aName,
        email: aEmail,
        role: validRole,
        nip: clean(3),
        status: clean(4) === 'Nonaktif' ? 'Nonaktif' : 'Aktif',
        signatureImage: '',
        schoolStampImage:
          validRole === 'Kepala Sekolah' ? settings.defaultSchoolStampImage || '' : '',
      });
    });
    return result;
  };

  const handleImportPaste = () => {
    const lines = pasteText.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    if (lines.length === 0) {
      onToast('error', 'Tempelkan baris data akun terlebih dahulu.');
      return;
    }
    const matrix = lines.map((l) => (l.includes('\t') ? l.split('\t') : l.includes(';') ? l.split(';') : l.split(',')));
    const parsed = parseRowsToAccounts(matrix);
    if (parsed.length === 0) {
      onToast('error', 'Format data akun tidak valid.');
      return;
    }
    onConfirm({
      title: 'Simpan Data Impor Akun?',
      description: `Ditemukan ${parsed.length} akun pengguna siap disimpan ke database. Apakah Anda ingin menyimpan data impor ini sekarang?`,
      confirmLabel: 'Ya, Simpan Akun',
      onConfirm: () => {
        const stats = onImportAccountsBatch(parsed);
        setPasteText('');
        setImportModalOpen(false);
        onToast('success', `Impor selesai: ${stats.added} akun ditambahkan, ${stats.updated} diperbarui.`);
      },
      onCancel: () => {
        onToast('info', 'Impor akun pengguna dibatalkan, data belum disimpan.');
      },
    });
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' });
      const parsed = parseRowsToAccounts(rows);
      if (parsed.length === 0) {
        onToast('error', 'Tidak ditemukan baris akun valid di dalam file.');
        return;
      }
      onConfirm({
        title: 'Simpan Data Impor Akun?',
        description: `File "${file.name}" berisi ${parsed.length} akun pengguna valid. Apakah Anda ingin menyimpan data impor ini ke database?`,
        confirmLabel: 'Ya, Simpan Akun',
        onConfirm: () => {
          const stats = onImportAccountsBatch(parsed);
          setImportModalOpen(false);
          onToast('success', `Impor "${file.name}": ${stats.added} ditambah, ${stats.updated} diperbarui.`);
        },
        onCancel: () => {
          onToast('info', `Impor file "${file.name}" dibatalkan, data belum disimpan.`);
        },
      });
    } catch {
      onToast('error', 'Gagal membaca file impor akun.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  return (
    <div className="space-y-6">
      {/* Compact Header Action Bar */}
      <section className="bg-white border border-slate-200 rounded-xl px-5 py-4 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div className="flex items-center gap-2.5">
          <ShieldCheck className="w-5 h-5 text-indigo-600 shrink-0" />
          <div>
            <h2 className="text-base font-bold text-slate-900">
              Manajemen Akun Pengguna ({accounts.length})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Kelola akun login Username & Password atau Email Google beserta tanda tangan digital dan stempel sekolah
            </p>
          </div>
        </div>
        {isAdminSession && (
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setImportModalOpen(true)}
              className="px-3.5 py-2 rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
            >
              <Upload className="w-4 h-4 text-indigo-600" />
              <span>Impor Akun Massal</span>
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
              <span>Tambah Akun Pengguna Baru</span>
            </button>
          </div>
        )}
      </section>

      {/* Tabel Daftar Akun */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 space-y-3">
          <div className="flex flex-col lg:flex-row lg:items-end lg:justify-between gap-3">
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Daftar Akun Terdaftar ({filteredAccounts.length})
              </h2>
              <p className="text-xs text-slate-500">
                Pilih baris untuk cetak/ekspor selektif atau edit TTD & stempel masing-masing akun
              </p>
            </div>

            <div className="flex flex-wrap items-end gap-2.5">
              <div>
                <label htmlFor="acc-filter-role" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Peran Akun
                </label>
                <select
                  id="acc-filter-role"
                  value={roleFilter}
                  onChange={(e) => setRoleFilter(e.target.value)}
                  aria-label="Filter peran akun"
                  className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                >
                  <option value="">--Pilih Peran--</option>
                  <option value="ALL">Semua Peran</option>
                  {ROLE_OPTIONS.map((r) => (
                    <option key={r} value={r}>
                      {r}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label htmlFor="acc-filter-search" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Cari Akun
                </label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    id="acc-filter-search"
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Cari nama / username / email..."
                    className="pl-9 pr-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 w-48 sm:w-60"
                  />
                </div>
              </div>
            </div>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedIds.size > 0
                ? `${selectedIds.size} akun dipilih`
                : 'Ekspor seluruh baris terfilter atau centang baris tertentu'}
            </span>
            <TableExportToolbar
              selectedCount={selectedIds.size}
              defaultOrientation="portrait"
              onDeleteSelected={isAdminSession ? handleBulkDelete : undefined}
              onDownloadTemplate={isAdminSession ? handleDownloadTemplate : undefined}
              onExportXlsx={() => {
                const { headers, rows } = buildAccountMatrix();
                exportRowsToXlsx('Data Akun', headers, rows, `Daftar_Akun_E-Guru_${getTodayISO()}.xlsx`);
                onToast('success', 'Data akun berhasil diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildAccountMatrix();
                exportRowsToXls('Data Akun', headers, rows, `Daftar_Akun_E-Guru_${getTodayISO()}.xls`);
                onToast('success', 'Data akun berhasil diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildAccountMatrix();
                exportDatasetToWordDoc({
                  db,
                  title: 'DAFTAR AKUN PENGGUNA PORTAL E-GURU',
                  subtitle: `${settings.schoolName} · Total ${exportSourceAccounts.length} Akun`,
                  headers,
                  rows,
                  filename: `Daftar_Akun_E-Guru_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) Daftar Akun diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildAccountMatrix();
                exportDatasetToPdf({
                  db,
                  title: 'DAFTAR AKUN PENGGUNA PORTAL E-GURU',
                  subtitle: `${settings.schoolName} · Total ${exportSourceAccounts.length} Akun`,
                  headers,
                  rows,
                  filename: `Daftar_Akun_E-Guru_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF Daftar Akun diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildAccountMatrix();
                printDatasetExact({
                  db,
                  title: 'DAFTAR AKUN PENGGUNA PORTAL E-GURU',
                  subtitle: `${settings.schoolName} · Total ${exportSourceAccounts.length} Akun`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('info', 'Menyiapkan tampilan cetak daftar akun...');
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
                      filteredAccounts.length > 0 &&
                      filteredAccounts.every((a) => selectedIds.has(a.id))
                    }
                    onChange={toggleSelectAll}
                    aria-label="Pilih semua akun"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-3 py-3 w-10 text-center">No</th>
                <th className="px-3 py-3 min-w-[180px]">Nama Pengguna</th>
                <th className="px-3 py-3 whitespace-nowrap">Username & Email</th>
                <th className="px-3 py-3 whitespace-nowrap">Peran</th>
                <th className="px-3 py-3 whitespace-nowrap">TTD & Stempel</th>
                <th className="px-3 py-3 whitespace-nowrap">Login Terakhir</th>
                <th className="px-3 py-3 whitespace-nowrap">Status</th>
                <th className="px-3 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {filteredAccounts.length === 0 ? (
                <tr>
                  <td colSpan={9} className="px-4 py-10 text-center text-sm text-slate-500">
                    Tidak ada akun yang cocok dengan filter.
                  </td>
                </tr>
              ) : (
                filteredAccounts.map((acc, idx) => {
                  const isCurrentUser =
                    normalizeEmailId(acc.email) === currentEmailLower ||
                    (Boolean(session.accountId) && acc.id === session.accountId);
                  const canEditRow = isAdminSession || isCurrentUser;
                  const checked = selectedIds.has(acc.id);
                  const isSyntheticEmail = acc.email.endsWith('@akun.portal');
                  return (
                    <tr
                      key={acc.id}
                      className={`hover:bg-slate-50/80 transition-colors ${checked ? 'bg-indigo-50/40' : ''}`}
                    >
                      <td className="px-3 py-3 text-center">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectRow(acc.id)}
                          aria-label={`Pilih ${acc.name}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-3 py-3 text-center text-xs text-slate-500 tabular-nums">
                        {idx + 1}
                      </td>
                      <td className="px-3 py-3 whitespace-nowrap">
                        <p className="font-medium text-slate-900">
                          {acc.name}{' '}
                          {isCurrentUser && (
                            <span className="text-xs font-semibold text-indigo-600">(Aktif)</span>
                          )}
                        </p>
                        {acc.nip && (
                          <p className="text-xs text-slate-500 font-mono-tabular">NIP {acc.nip}</p>
                        )}
                      </td>
                      <td className="px-3 py-3 text-xs whitespace-nowrap">
                        <div className="space-y-1">
                          {acc.username && (
                            <div className="inline-flex items-center gap-1.5 mr-2">
                              <span className="font-mono-tabular font-semibold text-indigo-700 bg-indigo-50 border border-indigo-200 px-1.5 py-0.5 rounded">
                                @{acc.username}
                              </span>
                              {acc.hasPassword && (
                                <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded inline-flex items-center gap-0.5">
                                  <KeyRound className="w-2.5 h-2.5" />
                                  Password Aktif
                                </span>
                              )}
                            </div>
                          )}
                          <p className="text-slate-600">
                            {isSyntheticEmail ? (
                              <span className="text-[11px] text-slate-400">
                                Login Username & Password
                              </span>
                            ) : (
                              acc.email
                            )}
                          </p>
                        </div>
                      </td>
                      <td className="px-3 py-3 text-xs font-semibold text-slate-800 whitespace-nowrap">
                        {acc.role}
                      </td>
                      <td className="px-3 py-3 whitespace-nowrap">
                        <div className="flex items-center gap-2">
                          {acc.signatureImage ? (
                            <img
                              src={acc.signatureImage}
                              alt={`TTD ${acc.name}`}
                              className="h-7 max-w-[68px] object-contain border border-slate-200 rounded bg-white px-1"
                            />
                          ) : (
                            <span className="text-[11px] text-slate-400">Tanpa TTD</span>
                          )}
                          {acc.role === 'Kepala Sekolah' && acc.schoolStampImage && (
                            <img
                              src={acc.schoolStampImage}
                              alt="Stempel"
                              className="h-7 w-7 object-contain rounded-full border border-slate-200 bg-white"
                            />
                          )}
                        </div>
                      </td>
                      <td className="px-3 py-3 whitespace-nowrap text-xs font-mono-tabular text-slate-500">
                        {acc.lastLogin || '-'}
                      </td>
                      <td className="px-3 py-3 text-xs font-semibold whitespace-nowrap">
                        <span
                          className={
                            acc.status === 'Aktif' ? 'text-emerald-700' : 'text-rose-700'
                          }
                        >
                          {acc.status}
                        </span>
                      </td>
                      <td className="px-3 py-3 whitespace-nowrap text-right">
                        <div className="inline-flex items-center gap-1.5">
                          {isAdminSession && (
                            <button
                              type="button"
                              onClick={() => {
                                void handleToggleStatus(acc);
                              }}
                              className="px-2 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 inline-flex items-center gap-1"
                              title={acc.status === 'Aktif' ? 'Nonaktifkan Akun' : 'Aktifkan Akun'}
                            >
                              <Power className="w-3.5 h-3.5" />
                            </button>
                          )}
                          {canEditRow ? (
                            <button
                              type="button"
                              onClick={() => handleEdit(acc)}
                              className="px-2 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 inline-flex items-center gap-1"
                            >
                              <Pencil className="w-3.5 h-3.5" />
                              Edit
                            </button>
                          ) : (
                            <span className="text-[11px] text-slate-400 px-1">Hanya Baca</span>
                          )}
                          {isAdminSession && (
                            <button
                              type="button"
                              onClick={() => handleDelete(acc)}
                              className="px-2 py-1.5 rounded-md border border-rose-200 text-xs font-medium text-rose-700 bg-rose-50/60 hover:bg-rose-100 inline-flex items-center gap-1"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
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
        title={editingId ? 'Edit Akun, Kredensial & TTD' : 'Tambah Akun Pengguna Baru'}
        subtitle="Isi Username & Password dan/atau Email Google untuk akses masuk, serta TTD & stempel saat cetak"
        maxWidthClass="max-w-4xl"
        onClose={() => {
          setFormModalOpen(false);
          resetForm();
        }}
      >
        <form onSubmit={handleSubmit} className="p-5 space-y-4">
          {teachers.length > 0 && !editingId && (
            <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 flex flex-col sm:flex-row sm:items-center gap-2.5">
              <span className="text-xs font-semibold text-slate-700 whitespace-nowrap">
                Isi Cepat dari Data Guru:
              </span>
              <select
                defaultValue=""
                onChange={(e) => handleLinkFromTeacher(e.target.value)}
                className="flex-1 px-3 py-1.5 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Guru / Tendik--</option>
                {teachers.map((t) => (
                  <option key={t.id} value={t.id}>
                    {t.name} ({t.position})
                  </option>
                ))}
              </select>
            </div>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label htmlFor="acc-name" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Nama Pengguna <span className="text-rose-600">*</span>
              </label>
              <input
                id="acc-name"
                type="text"
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Nama lengkap"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="acc-email" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Email Google Login <span className="text-slate-400 font-normal">(Opsional jika pakai Username)</span>
              </label>
              <input
                id="acc-email"
                type="email"
                value={email}
                disabled={isSelfLimitedEdit}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="email@sekolah.sch.id"
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 disabled:bg-slate-100 disabled:text-slate-500"
              />
            </div>
          </div>

          {/* Kredensial Login Username & Password */}
          <div className="rounded-xl border border-indigo-200/80 bg-indigo-50/40 p-4 space-y-3">
            <div className="flex items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <KeyRound className="w-4 h-4 text-indigo-600 shrink-0" />
                <span className="text-xs font-bold text-slate-800">
                  Kredensial Login Username & Password
                </span>
              </div>
              {editingAccount?.hasPassword && (
                <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md">
                  Password Sudah Terpasang
                </span>
              )}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div>
                <label htmlFor="acc-username" className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Username Login
                </label>
                <div className="relative">
                  <User className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    id="acc-username"
                    type="text"
                    autoComplete="off"
                    value={username}
                    disabled={isSelfLimitedEdit}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="contoh: budi.santoso"
                    className="w-full pl-8 pr-3 py-2 text-sm font-mono-tabular bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 disabled:bg-slate-100 disabled:text-slate-500"
                  />
                </div>
                <p className="mt-1 text-[11px] text-slate-500">
                  3–40 karakter huruf kecil, angka, titik (.), garis bawah (_), atau (-).
                </p>
              </div>

              <div>
                <label htmlFor="acc-password" className="block text-xs font-semibold text-slate-700 mb-1.5">
                  {editingAccount?.hasPassword ? 'Password Baru (Opsional)' : 'Password Login'}
                </label>
                <div className="relative">
                  <Lock className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    id="acc-password"
                    type={showPassword ? 'text' : 'password'}
                    autoComplete="new-password"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder={
                      editingAccount?.hasPassword
                        ? 'Kosongkan jika tidak ingin mengubah password'
                        : 'Minimal 6 karakter'
                    }
                    className="w-full pl-8 pr-9 py-2 text-sm bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword((prev) => !prev)}
                    aria-label={showPassword ? 'Sembunyikan password' : 'Tampilkan password'}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-0.5"
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
                <p className="mt-1 text-[11px] text-slate-500">
                  Password di-hash secara aman di server dan tidak pernah disimpan sebagai teks biasa.
                </p>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div>
              <label htmlFor="acc-role" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Peran / Hak Akses
              </label>
              <select
                id="acc-role"
                value={role}
                disabled={isSelfLimitedEdit}
                onChange={(e) => setRole(e.target.value as AccountRole | '')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600 disabled:bg-slate-100 disabled:text-slate-500"
              >
                <option value="">--Pilih Peran--</option>
                {ROLE_OPTIONS.map((r) => (
                  <option key={r} value={r}>
                    {r}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label htmlFor="acc-nip" className="block text-xs font-semibold text-slate-700 mb-1.5">
                NIP
              </label>
              <input
                id="acc-nip"
                type="text"
                value={nip}
                onChange={(e) => setNip(e.target.value)}
                placeholder="Opsional"
                className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
              />
            </div>

            <div>
              <label htmlFor="acc-status" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Status Akun
              </label>
              <select
                id="acc-status"
                value={status}
                disabled={isSelfLimitedEdit}
                onChange={(e) => setStatus(e.target.value as 'Aktif' | 'Nonaktif' | '')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600 disabled:bg-slate-100 disabled:text-slate-500"
              >
                <option value="">--Pilih Status--</option>
                <option value="Aktif">Aktif</option>
                <option value="Nonaktif">Nonaktif</option>
              </select>
            </div>
          </div>

          {/* TTD Digital & Stempel Sekolah Milik Akun */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-1">
            <div className="border border-slate-200 rounded-xl p-3.5 space-y-2 bg-slate-50/60">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-slate-700">
                  Tanda Tangan Digital Pemilik Akun
                </span>
                {signatureImage && (
                  <button
                    type="button"
                    onClick={() => {
                      setSignatureImage('');
                      onToast('info', 'Tanda tangan dihapus dari pratinjau formulir akun.');
                    }}
                    className="text-xs text-rose-600 hover:underline"
                  >
                    Hapus
                  </button>
                )}
              </div>
              <div className="h-20 bg-white rounded-lg border border-dashed border-slate-300 flex items-center justify-center p-2 overflow-hidden">
                {signatureImage ? (
                  <img src={signatureImage} alt="TTD Akun" className="max-h-full max-w-full object-contain" />
                ) : (
                  <span className="text-xs text-slate-400">Belum ada tanda tangan</span>
                )}
              </div>
              <button
                type="button"
                onClick={() => setSigModalOpen(true)}
                className="w-full px-3 py-1.5 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold inline-flex items-center justify-center gap-1.5"
              >
                <PenLine className="w-3.5 h-3.5" />
                Gambar / Unggah TTD Digital
              </button>
            </div>

            {role === 'Kepala Sekolah' ? (
              <div className="border border-slate-200 rounded-xl p-3.5 space-y-2 bg-slate-50/60">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-700">
                    Stempel Sekolah (Kepala Sekolah)
                  </span>
                  <div className="flex items-center gap-2">
                    {settings.defaultSchoolStampImage && (
                      <button
                        type="button"
                        onClick={() => {
                          setSchoolStampImage(settings.defaultSchoolStampImage);
                          onToast('info', 'Stempel default sekolah diterapkan ke formulir akun.');
                        }}
                        className="text-xs text-indigo-600 hover:underline"
                      >
                        Pakai Default
                      </button>
                    )}
                    {schoolStampImage && (
                      <button
                        type="button"
                        onClick={() => {
                          setSchoolStampImage('');
                          onToast('info', 'Stempel dihapus dari pratinjau formulir akun.');
                        }}
                        className="text-xs text-rose-600 hover:underline"
                      >
                        Hapus
                      </button>
                    )}
                  </div>
                </div>
                <div className="h-20 bg-white rounded-lg border border-dashed border-slate-300 flex items-center justify-center p-2 overflow-hidden">
                  {schoolStampImage ? (
                    <img src={schoolStampImage} alt="Stempel Akun" className="max-h-full max-w-full object-contain" />
                  ) : (
                    <span className="text-xs text-slate-400">Belum ada stempel</span>
                  )}
                </div>
                <label className="cursor-pointer w-full px-3 py-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 text-slate-700 text-xs font-semibold inline-flex items-center justify-center gap-1.5">
                  <Stamp className="w-3.5 h-3.5 text-indigo-600" />
                  Unggah Stempel Kepala Sekolah
                  <input type="file" accept="image/*" className="hidden" onChange={handleStampUpload} />
                </label>
              </div>
            ) : (
              <div className="border border-slate-200 rounded-xl p-3.5 flex flex-col items-center justify-center text-center bg-slate-50/60 gap-1.5">
                <Stamp className="w-5 h-5 text-slate-400" />
                <p className="text-xs font-semibold text-slate-700">
                  Stempel Khusus Kepala Sekolah
                </p>
                <p className="text-[11px] text-slate-500 max-w-xs">
                  Stempel sekolah hanya berlaku untuk peran Kepala Sekolah dan diatur pada menu Pengaturan atau akun Kepala Sekolah.
                </p>
              </div>
            )}
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
              disabled={submitting}
              className="px-5 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 disabled:opacity-60 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
            >
              <UserPlus className="w-4 h-4" />
              {submitting
                ? 'Menyimpan...'
                : editingId
                  ? 'Simpan Perubahan Akun'
                  : 'Buat Akun'}
            </button>
          </div>
        </form>
      </FormModal>

      <FormModal
        open={importModalOpen}
        title="Impor Akun Massal (Excel / CSV / Tempel)"
        subtitle="Kolom: Nama Pengguna, Email Google, Peran, NIP, Status"
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
                id="account-file-upload"
              />
              <label
                htmlFor="account-file-upload"
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
            placeholder="Tempel baris data akun dari Excel (Tab / CSV)..."
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

      <SignaturePadModal
        open={sigModalOpen}
        title={`Tanda Tangan Digital${name ? ` — ${name}` : ''}`}
        initialImage={signatureImage}
        onClose={() => setSigModalOpen(false)}
        onSave={(dataUrl) => {
          setSignatureImage(dataUrl);
          onToast(
            dataUrl ? 'success' : 'info',
            dataUrl ? 'Tanda tangan digital disimpan ke formulir akun.' : 'Tanda tangan digital dikosongkan.'
          );
        }}
      />
    </div>
  );
}
