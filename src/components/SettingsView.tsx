import React, { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Check,
  Download,
  FileSpreadsheet,
  Globe,
  ImagePlus,
  Palette,
  Pencil,
  PenLine,
  Plus,
  RotateCcw,
  Save,
  Stamp,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import { PortalDatabase, SemesterType, Student, SystemSettings, ThemeId } from '../types';
import { exportRowsToXlsx } from '../utils/exportHelpers';
import { formatShortDateId, getTodayISO } from '../utils/portalHelpers';
import { PORTAL_THEMES } from '../utils/themeConfig';
import { ConfirmDialogState } from './FeedbackUI';
import { SignaturePadModal } from './SignaturePadModal';
import { StudentDetailModal } from './StudentDetailModal';

interface SettingsViewProps {
  db: PortalDatabase;
  onUpdateSettings: (updater: (prev: SystemSettings) => SystemSettings) => void;
  onRestoreDatabase: (incoming: PortalDatabase) => void;
  onResetAllData: () => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

type MasterKey =
  | 'masterClasses'
  | 'masterSubjects'
  | 'masterExamTypes'
  | 'masterAcademicYears';

const HEADER_STOPWORDS = new Set([
  'no',
  'nomor',
  'item',
  'nama',
  'kelas',
  'nama kelas',
  'daftar kelas',
  'mapel',
  'mata pelajaran',
  'nama mapel',
  'jenis ujian',
  'ujian',
  'tahun ajaran',
  'periode',
]);

function MasterListCard({
  title,
  subtitle,
  columnLabel,
  samplePlaceholder,
  sampleRows,
  exportFilePrefix,
  items,
  activeHighlightItem,
  onAdd,
  onBatchAdd,
  onRename,
  onDelete,
  onBatchDelete,
  onInspectItem,
  onToast,
}: {
  title: string;
  subtitle: string;
  columnLabel: string;
  samplePlaceholder: string;
  sampleRows: string[];
  exportFilePrefix: string;
  items: string[];
  activeHighlightItem?: string;
  onAdd: (value: string) => void;
  onBatchAdd: (values: string[], sourceLabel?: string) => number;
  onRename: (index: number, nextValue: string) => void;
  onDelete: (index: number) => void;
  onBatchDelete: (indices: number[]) => void;
  onInspectItem?: (item: string) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
}) {
  const [newValue, setNewValue] = useState('');
  const [editIdx, setEditIdx] = useState<number | null>(null);
  const [editVal, setEditVal] = useState('');
  const [importOpen, setImportOpen] = useState(false);
  const [importText, setImportText] = useState('');
  const [selectMode, setSelectMode] = useState(false);
  const [selectedIndices, setSelectedIndices] = useState<Set<number>>(new Set());
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = newValue.trim();
    if (!trimmed) return;
    onAdd(trimmed);
    setNewValue('');
  };

  const commitEdit = (idx: number) => {
    const trimmed = editVal.trim();
    if (trimmed) {
      onRename(idx, trimmed);
    }
    setEditIdx(null);
  };

  const handleDownloadTemplate = () => {
    exportRowsToXlsx(
      columnLabel,
      [columnLabel],
      sampleRows.map((val) => [val]),
      `Template_${exportFilePrefix}_E-Guru.xlsx`
    );
    onToast('success', `Template Excel ${title} berhasil diunduh.`);
  };

  const handleExportItems = () => {
    if (items.length === 0) {
      onToast('info', `Belum ada item pada ${title} untuk diekspor.`);
      return;
    }
    exportRowsToXlsx(
      columnLabel,
      ['No', columnLabel],
      items.map((item, idx) => [idx + 1, item]),
      `Ekspor_${exportFilePrefix}_${getTodayISO()}.xlsx`
    );
    onToast('success', `${items.length} item ${title} berhasil diekspor ke Excel.`);
  };

  const handlePasteImport = () => {
    const lines = importText
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    if (lines.length === 0) {
      onToast('info', 'Masukkan minimal 1 baris item untuk diimpor.');
      return;
    }
    const addedCount = onBatchAdd(lines);
    if (addedCount > 0) {
      setImportText('');
      setImportOpen(false);
    }
  };

  const handleExcelFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const ext = file.name.split('.').pop()?.toLowerCase() || '';
      let parsedValues: string[] = [];

      if (ext === 'csv' || ext === 'txt') {
        const text = await file.text();
        parsedValues = text
          .split(/\r?\n/)
          .map((line) => {
            const cells = line.split(/[;,\t]/).map((c) => c.trim()).filter(Boolean);
            if (cells.length === 0) return '';
            if (cells.length >= 2 && /^\d+$/.test(cells[0])) return cells[1];
            return cells[0];
          })
          .filter(Boolean);
      } else {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<(string | number | undefined)[]>(sheet, {
          header: 1,
          defval: '',
        });
        parsedValues = rows
          .map((row) => {
            const cleanCells = row
              .map((cell) => String(cell ?? '').trim())
              .filter(Boolean);
            if (cleanCells.length === 0) return '';
            if (cleanCells.length >= 2 && /^\d+$/.test(cleanCells[0])) {
              return cleanCells[1];
            }
            return cleanCells[0];
          })
          .filter(Boolean);
      }

      if (parsedValues.length > 0 && HEADER_STOPWORDS.has(parsedValues[0].toLowerCase())) {
        parsedValues = parsedValues.slice(1);
      }

      if (parsedValues.length === 0) {
        onToast('error', 'File tidak berisi data item yang valid.');
        return;
      }

      const addedCount = onBatchAdd(parsedValues, 'File Excel');
      if (addedCount > 0) {
        setImportText('');
        setImportOpen(false);
      }
    } catch {
      onToast('error', 'Gagal membaca file Excel/CSV.');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const toggleSelectMode = (checked: boolean) => {
    setSelectMode(checked);
    if (!checked) {
      setSelectedIndices(new Set());
    }
  };

  const toggleItemIndex = (idx: number) => {
    setSelectedIndices((prev) => {
      const next = new Set(prev);
      if (next.has(idx)) next.delete(idx);
      else next.add(idx);
      return next;
    });
  };

  const toggleSelectAllItems = () => {
    if (items.length > 0 && selectedIndices.size === items.length) {
      setSelectedIndices(new Set());
    } else {
      setSelectedIndices(new Set(items.map((_, i) => i)));
    }
  };

  const handleDeleteSelected = () => {
    if (selectedIndices.size === 0) return;
    onBatchDelete(Array.from(selectedIndices));
    setSelectedIndices(new Set());
    setSelectMode(false);
  };

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-4 shadow-xs flex flex-col gap-3">
      {/* Header Row */}
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <h3 className="text-[15px] font-bold text-slate-900 leading-snug">{title}</h3>
          <p className="text-[11px] text-slate-500 mt-0.5 leading-snug">{subtitle}</p>
        </div>
        <label className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-600 cursor-pointer shrink-0 select-none pt-0.5">
          <input
            type="checkbox"
            checked={selectMode}
            onChange={(e) => toggleSelectMode(e.target.checked)}
            className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500"
          />
          <span>Pilih</span>
        </label>
      </div>

      {/* Action Buttons Row: Template | Impor/Tutup | Ekspor */}
      <div className="flex items-center gap-1.5 flex-wrap">
        <button
          type="button"
          onClick={handleDownloadTemplate}
          className="px-2.5 py-1.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-[11px] font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shadow-2xs"
        >
          <Download className="w-3.5 h-3.5 text-indigo-500" />
          <span>Template</span>
        </button>

        <button
          type="button"
          onClick={() => setImportOpen((v) => !v)}
          className="px-3 py-1.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-[11px] font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shadow-2xs"
        >
          <Upload className="w-3.5 h-3.5" />
          <span>{importOpen ? 'Tutup' : 'Impor'}</span>
        </button>

        <button
          type="button"
          onClick={handleExportItems}
          className="px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-[11px] font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap shadow-2xs"
        >
          <FileSpreadsheet className="w-3.5 h-3.5" />
          <span>Ekspor</span>
        </button>
      </div>

      {/* Expandable Inline Bulk Import Panel */}
      {importOpen && (
        <div className="rounded-2xl border border-indigo-200 bg-indigo-50/25 p-3 space-y-2.5">
          <div className="flex items-center justify-between gap-2">
            <span className="text-[11px] font-bold text-slate-800">1 baris = 1 item</span>
            <label className="cursor-pointer px-2.5 py-1 rounded-full border border-indigo-200 bg-white hover:bg-indigo-50 text-indigo-600 text-[11px] font-semibold inline-flex items-center gap-1 shadow-2xs transition-colors whitespace-nowrap">
              <Upload className="w-3 h-3" />
              <span>File Excel</span>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv,.txt"
                onChange={handleExcelFileUpload}
                className="hidden"
              />
            </label>
          </div>

          <textarea
            rows={3}
            value={importText}
            onChange={(e) => setImportText(e.target.value)}
            placeholder={samplePlaceholder}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs text-slate-800 placeholder:text-slate-400 focus:outline-none focus:border-indigo-500 resize-y"
          />

          <div className="flex items-center justify-end gap-2">
            <button
              type="button"
              onClick={() => {
                setImportText('');
                setImportOpen(false);
              }}
              className="px-3 py-1.5 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-semibold transition-colors"
            >
              Batal
            </button>
            <button
              type="button"
              onClick={handlePasteImport}
              disabled={!importText.trim()}
              className="px-3.5 py-1.5 rounded-full bg-indigo-600 hover:bg-indigo-700 disabled:bg-indigo-300 disabled:cursor-not-allowed text-white text-[11px] font-semibold transition-colors"
            >
              Impor
            </button>
          </div>
        </div>
      )}

      {/* Quick Single-Item Add Bar */}
      <form onSubmit={handleAdd} className="flex items-center gap-2">
        <input
          type="text"
          value={newValue}
          onChange={(e) => setNewValue(e.target.value)}
          placeholder={`Contoh: ${samplePlaceholder}`}
          className="flex-1 min-w-0 px-3.5 py-2 text-xs border border-slate-200 rounded-full bg-white focus:outline-none focus:border-indigo-600 placeholder:text-slate-400"
        />
        <button
          type="submit"
          className="px-3.5 py-2 rounded-full bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold transition-colors inline-flex items-center gap-1 whitespace-nowrap shrink-0 shadow-xs"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>Tambah</span>
        </button>
      </form>

      {/* Bulk Select Controls Bar */}
      {selectMode && items.length > 0 && (
        <div className="flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-xl bg-slate-50 border border-slate-200 text-[11px]">
          <button
            type="button"
            onClick={toggleSelectAllItems}
            className="font-semibold text-indigo-600 hover:underline"
          >
            {selectedIndices.size === items.length ? 'Batal Semua' : `Pilih Semua (${items.length})`}
          </button>
          <button
            type="button"
            onClick={handleDeleteSelected}
            disabled={selectedIndices.size === 0}
            className="px-2.5 py-1 rounded-lg bg-rose-600 hover:bg-rose-700 disabled:bg-rose-300 text-white font-semibold inline-flex items-center gap-1 transition-colors"
          >
            <Trash2 className="w-3 h-3" />
            <span>Hapus ({selectedIndices.size})</span>
          </button>
        </div>
      )}

      {/* Scrollable Items List */}
      <div className="divide-y divide-slate-100 max-h-56 overflow-y-auto table-scroll border-t border-slate-100">
        {items.length === 0 ? (
          <p className="px-3 py-5 text-xs text-slate-400 text-center">Belum ada item.</p>
        ) : (
          items.map((item, idx) => (
            <div
              key={`${item}-${idx}`}
              className="px-2 py-2 flex items-center justify-between gap-2 hover:bg-slate-50 rounded-lg transition-colors"
            >
              {editIdx === idx ? (
                <div className="flex-1 flex items-center gap-1.5">
                  <input
                    type="text"
                    value={editVal}
                    onChange={(e) => setEditVal(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        commitEdit(idx);
                      } else if (e.key === 'Escape') {
                        setEditIdx(null);
                      }
                    }}
                    autoFocus
                    className="flex-1 min-w-0 px-2.5 py-1 text-xs border border-indigo-400 rounded-lg focus:outline-none"
                  />
                  <button
                    type="button"
                    onClick={() => commitEdit(idx)}
                    className="p-1 rounded text-emerald-700 hover:bg-emerald-50"
                    aria-label="Simpan"
                  >
                    <Check className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => setEditIdx(null)}
                    className="p-1 rounded text-slate-500 hover:bg-slate-100"
                    aria-label="Batal"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                </div>
              ) : (
                <>
                  <div className="flex items-center gap-2 min-w-0 flex-1">
                    {selectMode && (
                      <input
                        type="checkbox"
                        checked={selectedIndices.has(idx)}
                        onChange={() => toggleItemIndex(idx)}
                        className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 shrink-0"
                        aria-label={`Pilih ${item}`}
                      />
                    )}
                    {onInspectItem ? (
                      <button
                        type="button"
                        onClick={() => onInspectItem(item)}
                        title={`Klik untuk membuka detail ${item}`}
                        className="text-left text-xs sm:text-sm font-medium text-slate-800 hover:text-indigo-600 hover:underline truncate"
                      >
                        {item}
                      </button>
                    ) : (
                      <span className="text-xs sm:text-sm text-slate-800 truncate">
                        {item}
                        {activeHighlightItem === item && (
                          <span className="text-[11px] font-semibold text-indigo-600 ml-1.5">
                            · Aktif
                          </span>
                        )}
                      </span>
                    )}
                  </div>
                  <div className="flex items-center gap-0.5 shrink-0">
                    <button
                      type="button"
                      onClick={() => {
                        setEditIdx(idx);
                        setEditVal(item);
                      }}
                      className="p-1.5 rounded text-slate-500 hover:text-slate-800 hover:bg-slate-100"
                      aria-label={`Edit ${item}`}
                    >
                      <Pencil className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => onDelete(idx)}
                      className="p-1.5 rounded text-rose-600 hover:bg-rose-50"
                      aria-label={`Hapus ${item}`}
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </>
              )}
            </div>
          ))
        )}
      </div>
    </div>
  );
}

export function SettingsView({
  db,
  onUpdateSettings,
  onRestoreDatabase,
  onResetAllData,
  onToast,
  onConfirm,
}: SettingsViewProps) {
  const { settings } = db;
  const restoreInputRef = useRef<HTMLInputElement | null>(null);

  // Local editable state for Period & Officials
  const [academicYear, setAcademicYear] = useState(settings.academicYear);
  const [semester, setSemester] = useState<SemesterType>(settings.semester);
  const [signatureLocation, setSignatureLocation] = useState(settings.signatureLocation);
  const [principalTitle, setPrincipalTitle] = useState(settings.principalTitle);
  const [principalName, setPrincipalName] = useState(settings.principalName);
  const [principalNip, setPrincipalNip] = useState(settings.principalNip);
  const [principalSignatureImage, setPrincipalSignatureImage] = useState(
    settings.principalSignatureImage
  );
  const [defaultSchoolStampImage, setDefaultSchoolStampImage] = useState(
    settings.defaultSchoolStampImage
  );
  const [sigModalOpen, setSigModalOpen] = useState(false);

  // Identity & Theme
  const [themeId, setThemeId] = useState<ThemeId>(settings.themeId);
  const [identity, setIdentity] = useState(settings.identity);

  // Kop Surat state
  const [kop, setKop] = useState(settings.kopSurat);

  // Teacher Profile & Account
  const [teacherName, setTeacherName] = useState(settings.teacherName);
  const [teacherNip, setTeacherNip] = useState(settings.teacherNip);
  const [schoolName, setSchoolName] = useState(settings.schoolName);
  const [email, setEmail] = useState(settings.email);
  const [password, setPassword] = useState(settings.password);

  // Interactive Detail Modals for Master Kelas, Master Mapel & Detail Siswa
  const [detailClass, setDetailClass] = useState<string | null>(null);
  const [detailSubject, setDetailSubject] = useState<string | null>(null);
  const [detailStudent, setDetailStudent] = useState<Student | null>(null);

  const rombelStudents = useMemo(() => {
    if (!detailClass) return [];
    return db.students
      .filter((s) => s.className === detailClass)
      .sort(
        (a, b) =>
          (a.absentNo || '').localeCompare(b.absentNo || '', 'id', { numeric: true }) ||
          a.name.localeCompare(b.name)
      );
  }, [db.students, detailClass]);

  const subjectJournals = useMemo(() => {
    if (!detailSubject) return [];
    return db.journals
      .filter(
        (j) =>
          j.subject === detailSubject &&
          j.academicYear === db.session.academicYear &&
          j.semester === db.session.semester
      )
      .sort((a, b) => b.date.localeCompare(a.date));
  }, [db.journals, db.session.academicYear, db.session.semester, detailSubject]);

  const handleSaveIdentityAndTheme = (e: React.FormEvent) => {
    e.preventDefault();
    onUpdateSettings((prev) => ({
      ...prev,
      themeId,
      identity: {
        ...identity,
        webTitle: identity.webTitle.trim() || 'Portal E-Guru v10.2',
        sidebarTitle: identity.sidebarTitle.trim() || 'Portal E-Guru v10.2',
        loginTitle: identity.loginTitle.trim() || 'Selamat Datang',
      },
    }));
    onToast('success', 'Identitas Web, Sidebar, Halaman Login & Tema Portal berhasil disimpan.');
  };

  const handleQuickThemeSelect = (nextTheme: ThemeId) => {
    setThemeId(nextTheme);
    onUpdateSettings((prev) => ({
      ...prev,
      themeId: nextTheme,
    }));
    onToast('success', `Tema tampilan diubah ke "${PORTAL_THEMES[nextTheme].name}".`);
  };

  const handleSidebarLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran logo maksimal 2 MB.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const result = typeof reader.result === 'string' ? reader.result : '';
      setIdentity((prev) => ({ ...prev, sidebarLogo: result }));
    };
    reader.readAsDataURL(file);
  };

  const handleStampUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran stempel maksimal 2 MB.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const result = typeof reader.result === 'string' ? reader.result : '';
      setDefaultSchoolStampImage(result);
    };
    reader.readAsDataURL(file);
  };

  const handleSavePeriodAndOfficials = (e: React.FormEvent) => {
    e.preventDefault();
    onUpdateSettings((prev) => ({
      ...prev,
      academicYear,
      semester,
      signatureLocation: signatureLocation.trim(),
      principalTitle: principalTitle.trim(),
      principalName: principalName.trim(),
      principalNip: principalNip.trim(),
      principalSignatureImage,
      defaultSchoolStampImage,
    }));
    onToast('success', 'Pengaturan periode aktif, pejabat, TTD Kepsek & stempel berhasil disimpan.');
  };

  const handleSaveKopSurat = (e: React.FormEvent) => {
    e.preventDefault();
    onUpdateSettings((prev) => ({
      ...prev,
      kopSurat: { ...kop },
    }));
    onToast('success', 'Pengaturan Kop Surat Dinas berhasil disimpan.');
  };

  const handleLogoUpload = (
    side: 'logoLeft' | 'logoRight',
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran file logo maksimal 2 MB.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const result = typeof reader.result === 'string' ? reader.result : '';
      setKop((prev) => ({ ...prev, [side]: result }));
    };
    reader.readAsDataURL(file);
  };

  const handleSaveProfileAccount = (e: React.FormEvent) => {
    e.preventDefault();
    if (!teacherName.trim() || !email.trim() || !password.trim()) {
      onToast('error', 'Nama Guru, Email, dan Password wajib diisi.');
      return;
    }
    onUpdateSettings((prev) => ({
      ...prev,
      teacherName: teacherName.trim(),
      teacherNip: teacherNip.trim(),
      schoolName: schoolName.trim(),
      email: email.trim(),
      password,
    }));
    onToast('success', 'Profil guru & akun login lokal berhasil diperbarui.');
  };

  // Master Data Handlers
  const addMasterItem = (key: MasterKey, value: string) => {
    onUpdateSettings((prev) => {
      if (prev[key].some((v) => v.toLowerCase() === value.toLowerCase())) {
        onToast('info', `"${value}" sudah ada dalam daftar.`);
        return prev;
      }
      onToast('success', `"${value}" ditambahkan.`);
      return { ...prev, [key]: [...prev[key], value] };
    });
  };

  const batchAddMasterItems = (
    key: MasterKey,
    values: string[],
    sourceLabel?: string
  ): number => {
    const existingLower = new Set(db.settings[key].map((v) => v.toLowerCase()));
    const toAdd: string[] = [];
    const seenLower = new Set<string>();

    for (const raw of values) {
      const cleaned = raw.trim();
      if (!cleaned) continue;
      const lower = cleaned.toLowerCase();
      if (existingLower.has(lower) || seenLower.has(lower)) continue;
      seenLower.add(lower);
      toAdd.push(cleaned);
    }

    if (toAdd.length === 0) {
      onToast('info', 'Semua item dalam daftar impor sudah ada.');
      return 0;
    }

    onUpdateSettings((prev) => {
      const currLower = new Set(prev[key].map((v) => v.toLowerCase()));
      const filtered = toAdd.filter((v) => !currLower.has(v.toLowerCase()));
      return { ...prev, [key]: [...prev[key], ...filtered] };
    });
    onToast(
      'success',
      `${toAdd.length} item baru berhasil diimpor${sourceLabel ? ` dari ${sourceLabel}` : ''}.`
    );
    return toAdd.length;
  };

  const renameMasterItem = (key: MasterKey, index: number, nextValue: string) => {
    onUpdateSettings((prev) => {
      const nextList = prev[key].map((item, i) => (i === index ? nextValue : item));
      return { ...prev, [key]: nextList };
    });
    onToast('success', 'Item master data diperbarui.');
  };

  const deleteMasterItem = (key: MasterKey, index: number) => {
    onUpdateSettings((prev) => ({
      ...prev,
      [key]: prev[key].filter((_, i) => i !== index),
    }));
    onToast('info', 'Item dihapus dari master data.');
  };

  const batchDeleteMasterItems = (key: MasterKey, indices: number[]) => {
    const idxSet = new Set(indices);
    onUpdateSettings((prev) => ({
      ...prev,
      [key]: prev[key].filter((_, i) => !idxSet.has(i)),
    }));
    onToast('info', `${indices.length} item terpilih berhasil dihapus.`);
  };

  // Backup JSON
  const handleExportJson = () => {
    const blob = new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Backup_Portal_E-Guru_${getTodayISO()}.json`;
    a.click();
    URL.revokeObjectURL(url);
    onToast('success', 'File backup JSON berhasil diunduh.');
  };

  // Restore JSON
  const handleImportJson = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text) as PortalDatabase;
      if (!parsed || !parsed.settings || !Array.isArray(parsed.students)) {
        onToast('error', 'Format file backup JSON tidak valid.');
        return;
      }
      onRestoreDatabase(parsed);
      setAcademicYear(parsed.settings.academicYear);
      setSemester(parsed.settings.semester);
      setSignatureLocation(parsed.settings.signatureLocation);
      setPrincipalTitle(parsed.settings.principalTitle);
      setPrincipalName(parsed.settings.principalName);
      setPrincipalNip(parsed.settings.principalNip);
      setPrincipalSignatureImage(parsed.settings.principalSignatureImage || '');
      setDefaultSchoolStampImage(parsed.settings.defaultSchoolStampImage || '');
      if (parsed.settings.themeId) setThemeId(parsed.settings.themeId);
      if (parsed.settings.identity) setIdentity(parsed.settings.identity);
      setKop(parsed.settings.kopSurat);
      setTeacherName(parsed.settings.teacherName);
      setTeacherNip(parsed.settings.teacherNip);
      setSchoolName(parsed.settings.schoolName);
      setEmail(parsed.settings.email);
      setPassword(parsed.settings.password);
      onToast('success', 'Seluruh data berhasil dipulihkan dari file backup JSON.');
    } catch {
      onToast('error', 'Gagal membaca file JSON backup.');
    } finally {
      if (restoreInputRef.current) restoreInputRef.current.value = '';
    }
  };

  const handleResetLocal = () => {
    onConfirm({
      title: 'Reset Seluruh Data Lokal',
      description:
        'Tindakan ini akan menghapus seluruh data siswa, jurnal mengajar, nilai, dan pengaturan yang tersimpan di browser Anda dan mengembalikan ke sesi awal.',
      confirmLabel: 'Ya, Reset Semua Data',
      danger: true,
      onConfirm: () => {
        onResetAllData();
        onToast('info', 'Seluruh data lokal telah direset.');
      },
    });
  };

  return (
    <div className="space-y-8">
      {/* Row 0: Identitas Web, Sidebar, Halaman Login & Pilihan Tema Portal */}
      <div className="grid grid-cols-1 xl:grid-cols-12 gap-6 items-stretch">
        <section className="xl:col-span-7 bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col">
          <div className="px-5 py-4 border-b border-slate-200 flex items-center gap-2.5">
            <Globe className="w-5 h-5 text-indigo-600 shrink-0" />
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Identitas Web, Sidebar & Halaman Login
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Sesuaikan judul tab browser, teks & logo sidebar, serta sapaan pada halaman Login
              </p>
            </div>
          </div>

          <form onSubmit={handleSaveIdentityAndTheme} className="p-5 space-y-4 flex-1 flex flex-col justify-between">
            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="id-web-title" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Judul Tab Browser (Web Title)
                  </label>
                  <input
                    id="id-web-title"
                    type="text"
                    value={identity.webTitle}
                    onChange={(e) => setIdentity((p) => ({ ...p, webTitle: e.target.value }))}
                    placeholder="Portal E-Guru v10.2"
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="id-web-sub" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Subjudul Aplikasi
                  </label>
                  <input
                    id="id-web-sub"
                    type="text"
                    value={identity.webSubtitle}
                    onChange={(e) => setIdentity((p) => ({ ...p, webSubtitle: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="id-side-title" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Nama Brand Sidebar
                  </label>
                  <input
                    id="id-side-title"
                    type="text"
                    value={identity.sidebarTitle}
                    onChange={(e) => setIdentity((p) => ({ ...p, sidebarTitle: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="id-side-sub" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Subjudul Sidebar
                  </label>
                  <input
                    id="id-side-sub"
                    type="text"
                    value={identity.sidebarSubtitle}
                    onChange={(e) => setIdentity((p) => ({ ...p, sidebarSubtitle: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="id-login-title" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Judul Halaman Login
                  </label>
                  <input
                    id="id-login-title"
                    type="text"
                    value={identity.loginTitle}
                    onChange={(e) => setIdentity((p) => ({ ...p, loginTitle: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="id-login-sub" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Subjudul Halaman Login
                  </label>
                  <input
                    id="id-login-sub"
                    type="text"
                    value={identity.loginSubtitle}
                    onChange={(e) => setIdentity((p) => ({ ...p, loginSubtitle: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 items-center">
                <div>
                  <label htmlFor="id-login-foot" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Catatan Kaki Login
                  </label>
                  <input
                    id="id-login-foot"
                    type="text"
                    value={identity.loginFooterNote}
                    onChange={(e) => setIdentity((p) => ({ ...p, loginFooterNote: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>

                <div>
                  <span className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Logo Sidebar & Login
                  </span>
                  <div className="flex items-center gap-2">
                    <div className="w-10 h-10 rounded-lg border border-slate-200 bg-slate-50 flex items-center justify-center overflow-hidden shrink-0">
                      {identity.sidebarLogo ? (
                        <img src={identity.sidebarLogo} alt="Logo" className="max-w-full max-h-full object-contain" />
                      ) : (
                        <span className="text-[10px] text-slate-400">Ikon</span>
                      )}
                    </div>
                    <label className="cursor-pointer px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5">
                      <ImagePlus className="w-3.5 h-3.5 text-indigo-600" />
                      Unggah Logo
                      <input type="file" accept="image/*" className="hidden" onChange={handleSidebarLogoUpload} />
                    </label>
                    {identity.sidebarLogo && (
                      <button
                        type="button"
                        onClick={() => setIdentity((p) => ({ ...p, sidebarLogo: '' }))}
                        className="px-2.5 py-2 rounded-lg text-xs text-rose-600 hover:bg-rose-50"
                      >
                        Reset
                      </button>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                type="submit"
                className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
              >
                <Save className="w-4 h-4" />
                Simpan Identitas Portal
              </button>
            </div>
          </form>
        </section>

        {/* Pilihan Tema Tampilan Portal */}
        <section className="xl:col-span-5 bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col">
          <div className="px-5 py-4 border-b border-slate-200 flex items-center gap-2.5">
            <Palette className="w-5 h-5 text-indigo-600 shrink-0" />
            <div>
              <h2 className="text-base font-bold text-slate-900">Pilihan Tema Tampilan Portal</h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Klik tema untuk menerapkan tampilan warna sidebar, aksen, dan mode malam seketika
              </p>
            </div>
          </div>

          <div className="p-5 space-y-2.5 flex-1">
            {(Object.keys(PORTAL_THEMES) as ThemeId[]).map((id) => {
              const t = PORTAL_THEMES[id];
              const selected = themeId === id;
              return (
                <button
                  key={id}
                  type="button"
                  onClick={() => handleQuickThemeSelect(id)}
                  className={`w-full text-left p-3 rounded-xl border transition-all flex items-center justify-between gap-3 ${
                    selected
                      ? 'border-indigo-600 bg-indigo-50/60 ring-2 ring-indigo-600/20'
                      : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50'
                  }`}
                >
                  <div className="flex items-center gap-3 min-w-0">
                    <div className="flex items-center -space-x-1.5 shrink-0">
                      <span
                        className="w-6 h-6 rounded-full border-2 border-white shadow-xs"
                        style={{ backgroundColor: t.swatchSidebar }}
                      />
                      <span
                        className="w-6 h-6 rounded-full border-2 border-white shadow-xs"
                        style={{ backgroundColor: t.swatchPrimary }}
                      />
                    </div>
                    <div className="min-w-0">
                      <p className="text-xs sm:text-sm font-bold text-slate-900 truncate">
                        {t.name}
                      </p>
                      <p className="text-[11px] text-slate-500 truncate">{t.description}</p>
                    </div>
                  </div>
                  {selected && (
                    <span className="w-6 h-6 rounded-full bg-indigo-600 text-white flex items-center justify-center shrink-0">
                      <Check className="w-3.5 h-3.5" />
                    </span>
                  )}
                </button>
              );
            })}
          </div>
        </section>
      </div>

      {/* Row 1: Periode & Pejabat + Profil Guru & Akun */}
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 items-stretch">
        {/* Pengaturan Periode Aktif & Pejabat */}
        <section className="bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col">
          <div className="px-5 py-4 border-b border-slate-200">
            <h2 className="text-base font-bold text-slate-900">Periode Aktif & Pejabat Pengesah</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Menentukan tahun ajaran aktif serta identitas Kepala Sekolah pada dokumen cetak
            </p>
          </div>
          <form onSubmit={handleSavePeriodAndOfficials} className="p-5 space-y-4 flex-1 flex flex-col justify-between">
            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="set-year" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Tahun Ajaran Aktif
                  </label>
                  <select
                    id="set-year"
                    value={academicYear}
                    onChange={(e) => setAcademicYear(e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                  >
                    {settings.masterAcademicYears.map((yr) => (
                      <option key={yr} value={yr}>
                        {yr}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label htmlFor="set-sem" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Semester Aktif
                  </label>
                  <select
                    id="set-sem"
                    value={semester}
                    onChange={(e) => setSemester(e.target.value as SemesterType)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                  >
                    <option value="Ganjil">Ganjil</option>
                    <option value="Genap">Genap</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="set-loc" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Lokasi Tanda Tangan
                  </label>
                  <input
                    id="set-loc"
                    type="text"
                    value={signatureLocation}
                    onChange={(e) => setSignatureLocation(e.target.value)}
                    placeholder="Contoh: Kudus"
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="set-ptitle" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Jabatan Pejabat
                  </label>
                  <input
                    id="set-ptitle"
                    type="text"
                    value={principalTitle}
                    onChange={(e) => setPrincipalTitle(e.target.value)}
                    placeholder="Kepala Sekolah"
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="set-pname" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Nama Kepala Sekolah
                  </label>
                  <input
                    id="set-pname"
                    type="text"
                    value={principalName}
                    onChange={(e) => setPrincipalName(e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="set-pnip" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    NIP Kepala Sekolah
                  </label>
                  <input
                    id="set-pnip"
                    type="text"
                    value={principalNip}
                    onChange={(e) => setPrincipalNip(e.target.value)}
                    className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>

              {/* TTD Kepala Sekolah & Stempel Sekolah Default */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                <div className="border border-slate-200 rounded-xl p-3 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-700">TTD Kepala Sekolah</span>
                    {principalSignatureImage && (
                      <button
                        type="button"
                        onClick={() => setPrincipalSignatureImage('')}
                        className="text-[11px] text-rose-600 hover:underline"
                      >
                        Hapus
                      </button>
                    )}
                  </div>
                  <div className="h-16 rounded-lg bg-slate-50 border border-dashed border-slate-300 flex items-center justify-center p-1.5 overflow-hidden">
                    {principalSignatureImage ? (
                      <img
                        src={principalSignatureImage}
                        alt="TTD Kepala Sekolah"
                        className="max-h-full max-w-full object-contain"
                      />
                    ) : (
                      <span className="text-[11px] text-slate-400">Belum ada TTD</span>
                    )}
                  </div>
                  <button
                    type="button"
                    onClick={() => setSigModalOpen(true)}
                    className="w-full px-2.5 py-1.5 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs font-semibold inline-flex items-center justify-center gap-1.5"
                  >
                    <PenLine className="w-3.5 h-3.5" />
                    Gambar / Unggah TTD
                  </button>
                </div>

                <div className="border border-slate-200 rounded-xl p-3 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-700">Stempel Sekolah Default</span>
                    {defaultSchoolStampImage && (
                      <button
                        type="button"
                        onClick={() => setDefaultSchoolStampImage('')}
                        className="text-[11px] text-rose-600 hover:underline"
                      >
                        Hapus
                      </button>
                    )}
                  </div>
                  <div className="h-16 rounded-lg bg-slate-50 border border-dashed border-slate-300 flex items-center justify-center p-1.5 overflow-hidden">
                    {defaultSchoolStampImage ? (
                      <img
                        src={defaultSchoolStampImage}
                        alt="Stempel Sekolah"
                        className="max-h-full max-w-full object-contain"
                      />
                    ) : (
                      <span className="text-[11px] text-slate-400">Belum ada Stempel</span>
                    )}
                  </div>
                  <label className="cursor-pointer w-full px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-semibold inline-flex items-center justify-center gap-1.5">
                    <Stamp className="w-3.5 h-3.5 text-indigo-600" />
                    Unggah Stempel PNG
                    <input type="file" accept="image/*" className="hidden" onChange={handleStampUpload} />
                  </label>
                </div>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                type="submit"
                className="px-4 py-2 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
              >
                <Save className="w-4 h-4" />
                Simpan Periode, Pejabat & Pengesahan
              </button>
            </div>
          </form>
        </section>

        {/* Profil Guru Mapel & Akun Login */}
        <section className="bg-white border border-slate-200 rounded-xl overflow-hidden flex flex-col">
          <div className="px-5 py-4 border-b border-slate-200">
            <h2 className="text-base font-bold text-slate-900">Profil Guru Mapel & Akun Login</h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Identitas pengajar dan kredensial masuk lokal perangkat
            </p>
          </div>
          <form onSubmit={handleSaveProfileAccount} className="p-5 space-y-4 flex-1 flex flex-col justify-between">
            <div className="space-y-4">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="prof-name" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Nama Lengkap Guru
                  </label>
                  <input
                    id="prof-name"
                    type="text"
                    value={teacherName}
                    onChange={(e) => setTeacherName(e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="prof-nip" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    NIP Guru
                  </label>
                  <input
                    id="prof-nip"
                    type="text"
                    value={teacherNip}
                    onChange={(e) => setTeacherNip(e.target.value)}
                    className="w-full px-3 py-2 text-sm font-mono-tabular border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>

              <div>
                <label htmlFor="prof-school" className="block text-xs font-semibold text-slate-700 mb-1.5">
                  Unit Sekolah
                </label>
                <input
                  id="prof-school"
                  type="text"
                  value={schoolName}
                  onChange={(e) => setSchoolName(e.target.value)}
                  className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label htmlFor="prof-email" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Email Login
                  </label>
                  <input
                    id="prof-email"
                    type="email"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
                <div>
                  <label htmlFor="prof-pass" className="block text-xs font-semibold text-slate-700 mb-1.5">
                    Password Login
                  </label>
                  <input
                    id="prof-pass"
                    type="text"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                type="submit"
                className="px-4 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-sm font-semibold transition-colors inline-flex items-center gap-2"
              >
                <Save className="w-4 h-4" />
                Simpan Profil & Akun
              </button>
            </div>
          </form>
        </section>
      </div>

      {/* Pengaturan Kop Surat Standar Dinas + Live Preview */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200">
          <h2 className="text-base font-bold text-slate-900">Pengaturan Kop Surat Standar Dinas</h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Atur 6 baris teks kop surat dan unggah logo instansi kiri/kanan beserta pratinjau langsung bergaris ganda
          </p>
        </div>

        <form onSubmit={handleSaveKopSurat} className="p-5 space-y-6">
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div className="lg:col-span-7 space-y-3">
              {[
                { key: 'line1' as const, label: 'Baris 1 (Pemerintah / Yayasan)' },
                { key: 'line2' as const, label: 'Baris 2 (Dinas / Instansi Induk)' },
                { key: 'line3' as const, label: 'Baris 3 (Nama Satuan Pendidikan)' },
                { key: 'line4' as const, label: 'Baris 4 (Alamat Jalan & Kode Pos)' },
                { key: 'line5' as const, label: 'Baris 5 (Telepon & Laman)' },
                { key: 'line6' as const, label: 'Baris 6 (Email & NPSN)' },
              ].map((field) => (
                <div key={field.key}>
                  <label className="block text-xs font-semibold text-slate-700 mb-1">
                    {field.label}
                  </label>
                  <input
                    type="text"
                    value={kop[field.key]}
                    onChange={(e) => setKop((prev) => ({ ...prev, [field.key]: e.target.value }))}
                    className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                </div>
              ))}
            </div>

            <div className="lg:col-span-5 space-y-4">
              {/* Upload Logo Kiri & Kanan */}
              <div className="grid grid-cols-2 gap-3">
                <div className="border border-slate-200 rounded-xl p-3 text-center space-y-2">
                  <p className="text-xs font-semibold text-slate-700">Logo Kiri</p>
                  <div className="h-24 bg-slate-50 rounded-lg border border-dashed border-slate-300 flex items-center justify-center overflow-hidden p-2">
                    {kop.logoLeft ? (
                      <img src={kop.logoLeft} alt="Logo Kiri" className="max-h-full max-w-full object-contain" />
                    ) : (
                      <span className="text-xs text-slate-400">Kosong</span>
                    )}
                  </div>
                  <div className="flex items-center justify-center gap-1.5">
                    <label className="cursor-pointer px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-xs font-medium text-slate-700 inline-flex items-center gap-1">
                      <ImagePlus className="w-3.5 h-3.5" />
                      Unggah
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => handleLogoUpload('logoLeft', e)}
                      />
                    </label>
                    {kop.logoLeft && (
                      <button
                        type="button"
                        onClick={() => setKop((prev) => ({ ...prev, logoLeft: '' }))}
                        className="px-2 py-1.5 rounded-md text-xs text-rose-600 hover:bg-rose-50"
                      >
                        Hapus
                      </button>
                    )}
                  </div>
                </div>

                <div className="border border-slate-200 rounded-xl p-3 text-center space-y-2">
                  <p className="text-xs font-semibold text-slate-700">Logo Kanan</p>
                  <div className="h-24 bg-slate-50 rounded-lg border border-dashed border-slate-300 flex items-center justify-center overflow-hidden p-2">
                    {kop.logoRight ? (
                      <img src={kop.logoRight} alt="Logo Kanan" className="max-h-full max-w-full object-contain" />
                    ) : (
                      <span className="text-xs text-slate-400">Kosong</span>
                    )}
                  </div>
                  <div className="flex items-center justify-center gap-1.5">
                    <label className="cursor-pointer px-2.5 py-1.5 rounded-md bg-slate-100 hover:bg-slate-200 text-xs font-medium text-slate-700 inline-flex items-center gap-1">
                      <ImagePlus className="w-3.5 h-3.5" />
                      Unggah
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => handleLogoUpload('logoRight', e)}
                      />
                    </label>
                    {kop.logoRight && (
                      <button
                        type="button"
                        onClick={() => setKop((prev) => ({ ...prev, logoRight: '' }))}
                        className="px-2 py-1.5 rounded-md text-xs text-rose-600 hover:bg-rose-50"
                      >
                        Hapus
                      </button>
                    )}
                  </div>
                </div>
              </div>

              <div className="pt-2">
                <button
                  type="submit"
                  className="w-full py-2.5 px-4 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-semibold transition-colors inline-flex items-center justify-center gap-2"
                >
                  <Save className="w-4 h-4" />
                  Simpan Pengaturan Kop Surat
                </button>
              </div>
            </div>
          </div>

          {/* Live Preview Kop Surat Bergaris Ganda */}
          <div className="border border-slate-200 rounded-xl p-5 bg-slate-50">
            <p className="text-xs font-semibold text-slate-500 mb-3">
              Pratinjau Langsung (Live Preview) Kop Surat Resmi
            </p>
            <div className="bg-white p-5 border border-slate-200 text-black font-serif">
              <div className="pb-3 border-b-4 border-double border-black flex items-center justify-between gap-4">
                <div className="w-16 h-16 shrink-0 flex items-center justify-center">
                  {kop.logoLeft && (
                    <img src={kop.logoLeft} alt="Preview Logo Kiri" className="max-w-full max-h-full object-contain" />
                  )}
                </div>
                <div className="flex-1 text-center leading-tight">
                  {kop.line1 && <p className="text-xs sm:text-sm uppercase">{kop.line1}</p>}
                  {kop.line2 && <p className="text-sm sm:text-base font-bold uppercase">{kop.line2}</p>}
                  {kop.line3 && <p className="text-base sm:text-lg font-bold uppercase">{kop.line3}</p>}
                  {kop.line4 && <p className="text-[11px] sm:text-xs mt-1">{kop.line4}</p>}
                  {kop.line5 && <p className="text-[11px] sm:text-xs">{kop.line5}</p>}
                  {kop.line6 && <p className="text-[11px] sm:text-xs">{kop.line6}</p>}
                </div>
                <div className="w-16 h-16 shrink-0 flex items-center justify-center">
                  {kop.logoRight && (
                    <img src={kop.logoRight} alt="Preview Logo Kanan" className="max-w-full max-h-full object-contain" />
                  )}
                </div>
              </div>
            </div>
          </div>
        </form>
      </section>

      {/* 4 Kartu Master Data */}
      <section className="space-y-3">
        <div>
          <h2 className="text-base font-bold text-slate-900">Master Data Referensi</h2>
          <p className="text-xs text-slate-500">
            Kelola daftar kelas, mata pelajaran, jenis ujian, dan tahun ajaran beserta fitur template, impor massal/Excel, dan ekspor
          </p>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 2xl:grid-cols-4 gap-5 items-start">
          <MasterListCard
            title="Master Daftar Kelas"
            subtitle="Klik nama kelas untuk buka detail rombel"
            columnLabel="Nama Kelas"
            samplePlaceholder="XI IPA 2"
            sampleRows={['X-MIPA 1', 'X-MIPA 2', 'XI IPA 1', 'XI IPA 2']}
            exportFilePrefix="Master_Daftar_Kelas"
            items={settings.masterClasses}
            onAdd={(v) => addMasterItem('masterClasses', v)}
            onBatchAdd={(vals, src) => batchAddMasterItems('masterClasses', vals, src)}
            onRename={(i, v) => renameMasterItem('masterClasses', i, v)}
            onDelete={(i) => deleteMasterItem('masterClasses', i)}
            onBatchDelete={(idxs) => batchDeleteMasterItems('masterClasses', idxs)}
            onInspectItem={(cls) => setDetailClass(cls)}
            onToast={onToast}
          />
          <MasterListCard
            title="Master Mata Pelajaran"
            subtitle="Klik nama mapel untuk buka rekap KBM"
            columnLabel="Mata Pelajaran"
            samplePlaceholder="Matematika Lanjut"
            sampleRows={['Matematika Wajib', 'Matematika Lanjut', 'Fisika', 'Kimia']}
            exportFilePrefix="Master_Mata_Pelajaran"
            items={settings.masterSubjects}
            onAdd={(v) => addMasterItem('masterSubjects', v)}
            onBatchAdd={(vals, src) => batchAddMasterItems('masterSubjects', vals, src)}
            onRename={(i, v) => renameMasterItem('masterSubjects', i, v)}
            onDelete={(i) => deleteMasterItem('masterSubjects', i)}
            onBatchDelete={(idxs) => batchDeleteMasterItems('masterSubjects', idxs)}
            onInspectItem={(subj) => setDetailSubject(subj)}
            onToast={onToast}
          />
          <MasterListCard
            title="Master Jenis Ujian"
            subtitle="Kolom penilaian pada Leger Matriks Nilai"
            columnLabel="Jenis Ujian"
            samplePlaceholder="Remedial 1"
            sampleRows={['UH 1', 'UH 2', 'PTS', 'PAS', 'Remedial 1']}
            exportFilePrefix="Master_Jenis_Ujian"
            items={settings.masterExamTypes}
            onAdd={(v) => addMasterItem('masterExamTypes', v)}
            onBatchAdd={(vals, src) => batchAddMasterItems('masterExamTypes', vals, src)}
            onRename={(i, v) => renameMasterItem('masterExamTypes', i, v)}
            onDelete={(i) => deleteMasterItem('masterExamTypes', i)}
            onBatchDelete={(idxs) => batchDeleteMasterItems('masterExamTypes', idxs)}
            onToast={onToast}
          />
          <MasterListCard
            title="Master Tahun Ajaran"
            subtitle="Opsi periode pada halaman login & sistem"
            columnLabel="Tahun Ajaran"
            samplePlaceholder="2027/2028"
            sampleRows={['2025/2026', '2026/2027', '2027/2028']}
            exportFilePrefix="Master_Tahun_Ajaran"
            items={settings.masterAcademicYears}
            activeHighlightItem={settings.academicYear}
            onAdd={(v) => addMasterItem('masterAcademicYears', v)}
            onBatchAdd={(vals, src) => batchAddMasterItems('masterAcademicYears', vals, src)}
            onRename={(i, v) => renameMasterItem('masterAcademicYears', i, v)}
            onDelete={(i) => deleteMasterItem('masterAcademicYears', i)}
            onBatchDelete={(idxs) => batchDeleteMasterItems('masterAcademicYears', idxs)}
            onToast={onToast}
          />
        </div>
      </section>

      {/* Utilitas Manajemen Data Lokal */}
      <section className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
        <div>
          <h2 className="text-base font-bold text-slate-900">Utilitas Manajemen Data Lokal (LocalStorage)</h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Cadangkan seluruh data ke file JSON sebelum berganti perangkat atau bersihkan penyimpanan lokal browser
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          <button
            type="button"
            onClick={handleExportJson}
            className="px-4 py-2 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-2 whitespace-nowrap"
          >
            <Download className="w-4 h-4" />
            Ekspor Backup JSON
          </button>

          <input
            ref={restoreInputRef}
            type="file"
            accept=".json"
            onChange={handleImportJson}
            className="hidden"
            id="restore-json-input"
          />
          <label
            htmlFor="restore-json-input"
            className="cursor-pointer px-4 py-2 rounded-lg border border-indigo-200 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-2 whitespace-nowrap"
          >
            <Upload className="w-4 h-4" />
            Impor Restore JSON
          </label>

          <button
            type="button"
            onClick={handleResetLocal}
            className="px-4 py-2 rounded-lg border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-700 text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-2 whitespace-nowrap"
          >
            <RotateCcw className="w-4 h-4" />
            Reset Seluruh Data Lokal
          </button>
        </div>
      </section>

      <SignaturePadModal
        open={sigModalOpen}
        title="Tanda Tangan Digital Kepala Sekolah"
        initialImage={principalSignatureImage}
        onClose={() => setSigModalOpen(false)}
        onSave={(dataUrl) => setPrincipalSignatureImage(dataUrl)}
      />

      {/* Modal Detail Rombel (Master Daftar Kelas) */}
      {detailClass && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-[1px] flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-3xl w-full shadow-xl overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between gap-3">
              <div>
                <h3 className="text-base font-bold text-slate-900">
                  Detail Rombel: {detailClass}
                </h3>
                <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
                  {rombelStudents.length} siswa terdaftar · L:{' '}
                  {rombelStudents.filter((s) => s.gender === 'L').length} · P:{' '}
                  {rombelStudents.filter((s) => s.gender === 'P').length} · Aktif:{' '}
                  {rombelStudents.filter((s) => s.status === 'Aktif').length}
                </p>
              </div>
              <button
                type="button"
                onClick={() => setDetailClass(null)}
                className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100"
                aria-label="Tutup"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="max-h-[420px] overflow-y-auto table-scroll">
              <table className="w-full text-left border-collapse text-xs sm:text-sm">
                <thead className="sticky top-0 bg-slate-50 border-b border-slate-200 z-10">
                  <tr>
                    <th className="px-4 py-2.5 font-semibold text-slate-700 w-14">Absen</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">NIS / NISN</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Nama Siswa</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700 w-12">L/P</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Status</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Wali Kelas</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {rombelStudents.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="px-4 py-8 text-center text-xs text-slate-400">
                        Belum ada siswa yang terdaftar pada rombel {detailClass}.
                      </td>
                    </tr>
                  ) : (
                    rombelStudents.map((st, idx) => (
                      <tr key={st.id} className="hover:bg-slate-50">
                        <td className="px-4 py-2 font-mono-tabular text-slate-600">
                          {st.absentNo || idx + 1}
                        </td>
                        <td className="px-4 py-2 font-mono-tabular text-slate-600">
                          {st.nis}
                          {st.nisn ? ` / ${st.nisn}` : ''}
                        </td>
                        <td className="px-4 py-2 font-medium text-slate-900">
                          <button
                            type="button"
                            onClick={() => setDetailStudent(st)}
                            title="Klik untuk melihat detail informasi siswa"
                            className="text-left font-medium text-slate-900 hover:text-indigo-600 hover:underline cursor-pointer"
                          >
                            {st.name}
                          </button>
                        </td>
                        <td className="px-4 py-2 text-slate-700">{st.gender}</td>
                        <td className="px-4 py-2 text-slate-700">{st.status}</td>
                        <td className="px-4 py-2 text-slate-600">{st.homeroomTeacher || '-'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            <div className="px-5 py-3 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setDetailClass(null)}
                className="px-4 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Rekap KBM (Master Mata Pelajaran) */}
      {detailSubject && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-[1px] flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-2xl max-w-3xl w-full shadow-xl overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between gap-3">
              <div>
                <h3 className="text-base font-bold text-slate-900">
                  Rekap KBM: {detailSubject}
                </h3>
                <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
                  {subjectJournals.length} pertemuan tercatat · Periode {db.session.academicYear} ({db.session.semester})
                </p>
              </div>
              <button
                type="button"
                onClick={() => setDetailSubject(null)}
                className="p-1.5 rounded-lg text-slate-500 hover:bg-slate-100"
                aria-label="Tutup"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="max-h-[420px] overflow-y-auto table-scroll">
              <table className="w-full text-left border-collapse text-xs sm:text-sm">
                <thead className="sticky top-0 bg-slate-50 border-b border-slate-200 z-10">
                  <tr>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Tanggal</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Jam Ke</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Kelas</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Materi Pokok</th>
                    <th className="px-4 py-2.5 font-semibold text-slate-700">Presensi (H/S/I/A/D)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {subjectJournals.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="px-4 py-8 text-center text-xs text-slate-400">
                        Belum ada jurnal mengajar untuk mata pelajaran {detailSubject} pada periode aktif.
                      </td>
                    </tr>
                  ) : (
                    subjectJournals.map((j) => {
                      const vals = Object.values(j.attendance || {});
                      const h = vals.filter((a) => a.status === 'H').length;
                      const s = vals.filter((a) => a.status === 'S').length;
                      const i = vals.filter((a) => a.status === 'I').length;
                      const a = vals.filter((a) => a.status === 'A').length;
                      const d = vals.filter((a) => a.status === 'D').length;
                      return (
                        <tr key={j.id} className="hover:bg-slate-50">
                          <td className="px-4 py-2 font-mono-tabular text-slate-700 whitespace-nowrap">
                            {formatShortDateId(j.date)}
                          </td>
                          <td className="px-4 py-2 font-mono-tabular text-slate-600">{j.jamKe}</td>
                          <td className="px-4 py-2 font-medium text-slate-900">{j.className}</td>
                          <td className="px-4 py-2 text-slate-700">{j.topic}</td>
                          <td className="px-4 py-2 font-mono-tabular text-slate-600 whitespace-nowrap">
                            {h} / {s} / {i} / {a} / {d}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
            <div className="px-5 py-3 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setDetailSubject(null)}
                className="px-4 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      <StudentDetailModal
        student={detailStudent}
        db={db}
        onClose={() => setDetailStudent(null)}
      />
    </div>
  );
}
