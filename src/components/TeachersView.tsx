import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  BookOpen,
  Check,
  ChevronDown,
  Copy,
  ImagePlus,
  Pencil,
  Plus,
  Printer,
  RefreshCw,
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
import {
  compareClassNames,
  compressImageFileToDataUrl,
  formatTeacherAssignmentsSummary,
  getAllTeacherClasses,
  getTeacherClassesForSubject,
  getTeacherDummyPhoto,
  getTeacherSubjects,
  getTodayISO,
  isNonTeachingSubjectLabel,
  normalizePhotoInputUrl,
  normalizeTeacherSubjectClasses,
} from '../utils/portalHelpers';
import { ConfirmDialogState } from './FeedbackUI';
import { FormModal } from './FormModal';
import { PersonAvatar } from './PersonAvatar';
import { TableExportToolbar } from './TableExportToolbar';

interface TeachersViewProps {
  db: PortalDatabase;
  onUpsertTeacher: (teacher: Teacher) => void;
  onImportTeachersBatch: (incoming: Omit<Teacher, 'id'>[]) => { added: number; updated: number };
  onDeleteTeacher: (id: string) => void;
  onDeleteTeachersBatch: (ids: string[]) => void;
  onSyncFromAccounts?: () => { added: number; updated?: number };
  onOpenPrintForTeacher?: (teacher: Teacher) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

const EMPLOYMENT_TYPES: TeacherEmploymentStatus[] = [
  'PNS',
  'PPPK',
  'Tetap Yayasan',
  'Honorer',
];

const STANDARD_TEACHER_POSITIONS: string[] = [
  'Guru Mata Pelajaran',
  'Wali Kelas',
  'Guru BK',
  'Kepala Sekolah',
  'Wakil Kepala Sekolah',
  'Wakasek Kurikulum',
  'Wakasek Kesiswaan',
  'Wakasek Sarpras',
  'Wakasek Humas',
  'Pembina OSIS / Ekstrakurikuler',
  'Kepala Perpustakaan',
  'Kepala Laboratorium',
  'Tata Usaha (TU)',
  'Admin / Operator Sekolah',
];

const STANDARD_SUBJECT_OPTIONS: string[] = [
  'Pendidikan Agama',
  'Pendidikan Pancasila / PPKn',
  'Bahasa Indonesia',
  'Bahasa Inggris',
  'Matematika',
  'Matematika Wajib',
  'Matematika Peminatan',
  'Sejarah',
  'Fisika',
  'Kimia',
  'Biologi',
  'Geografi',
  'Ekonomi',
  'Sosiologi',
  'Informatika',
  'Seni Budaya',
  'PJOK',
  'Prakarya & Kewirausahaan',
  'Bahasa Daerah',
  'IPA',
  'IPS',
  'Bimbingan Konseling',
  'Umum / Non-Mapel',
  'Semua Mata Pelajaran',
];

const NON_TEACHING_POSITIONS = new Set([
  'Kepala Sekolah',
  'Tata Usaha (TU)',
  'Admin / Operator Sekolah',
]);

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
  onSyncFromAccounts,
  onOpenPrintForTeacher,
  onToast,
  onConfirm,
}: TeachersViewProps) {
  const { teachers, students, accounts, settings } = db;

  const handleSyncFromAccounts = () => {
    if (!onSyncFromAccounts) return;
    const { added, updated = 0 } = onSyncFromAccounts();
    if (added > 0 && updated > 0) {
      onToast(
        'success',
        `Sinkronisasi selesai: ${added} guru ditambahkan dan ${updated} data guru diperbarui dari Manajemen Akun.`
      );
    } else if (added > 0) {
      onToast('success', `${added} data guru berhasil disinkronkan dari Manajemen Akun.`);
    } else if (updated > 0) {
      onToast('success', `${updated} data guru & mata pelajaran berhasil disinkronkan.`);
    } else {
      onToast('info', 'Semua akun, jabatan, dan mata pelajaran sudah tersinkronisasi.');
    }
  };

  const positionOptions = useMemo(() => {
    const seen = new Set<string>();
    const list: string[] = [];
    const pushUnique = (val: string) => {
      const trimmed = val.trim();
      const lower = trimmed.toLowerCase();
      if (!trimmed || seen.has(lower)) return;
      seen.add(lower);
      list.push(trimmed);
    };
    STANDARD_TEACHER_POSITIONS.forEach(pushUnique);
    teachers.forEach((t) => pushUnique(t.position));
    return list;
  }, [teachers]);

  const subjectOptions = useMemo(() => {
    const seen = new Set<string>();
    const list: string[] = [];
    const pushUnique = (val: string) => {
      const trimmed = val.trim();
      const lower = trimmed.toLowerCase();
      if (
        !trimmed ||
        trimmed === '-' ||
        trimmed === 'Semua Mata Pelajaran' ||
        seen.has(lower)
      ) {
        return;
      }
      seen.add(lower);
      list.push(trimmed);
    };
    settings.masterSubjects.forEach(pushUnique);
    teachers.forEach((t) => {
      getTeacherSubjects(t).forEach(pushUnique);
      if (t.subject === 'Umum / Non-Mapel') pushUnique(t.subject);
    });
    STANDARD_SUBJECT_OPTIONS.forEach(pushUnique);
    return list;
  }, [settings.masterSubjects, teachers]);

  const classOptions = useMemo(() => {
    const set = new Set<string>(settings.masterClasses);
    students.forEach((s) => {
      if (s.className) set.add(s.className);
    });
    teachers.forEach((t) => {
      getAllTeacherClasses(t).forEach((c) => set.add(c));
    });
    return Array.from(set).filter(Boolean).sort(compareClassNames);
  }, [settings.masterClasses, students, teachers]);

  const [editingId, setEditingId] = useState<string | null>(null);
  const [nip, setNip] = useState('');
  const [name, setName] = useState('');
  const [position, setPosition] = useState('');
  const [customPositionMode, setCustomPositionMode] = useState(false);
  const [selectedSubjects, setSelectedSubjects] = useState<string[]>([]);
  const [subjectClassesMap, setSubjectClassesMap] = useState<Record<string, string[]>>({});
  const [subjectDropdownOpen, setSubjectDropdownOpen] = useState(false);
  const [customSubjectInput, setCustomSubjectInput] = useState('');
  const [customClassInputs, setCustomClassInputs] = useState<Record<string, string>>({});
  const subjectDropdownRef = useRef<HTMLDivElement | null>(null);
  const [email, setEmail] = useState('');
  const [phone, setPhone] = useState('');
  const [gender, setGender] = useState<'L' | 'P' | ''>('');
  const [employmentStatus, setEmploymentStatus] = useState<TeacherEmploymentStatus | ''>('');
  const [status, setStatus] = useState<'Aktif' | 'Cuti' | 'Nonaktif' | ''>('');
  const [address, setAddress] = useState('');
  const [photo, setPhoto] = useState('');

  // Import state
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Filters & Multi-Select
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [formModalOpen, setFormModalOpen] = useState(false);
  const [importModalOpen, setImportModalOpen] = useState(false);

  useEffect(() => {
    if (!subjectDropdownOpen) return;
    const handleOutside = (e: MouseEvent) => {
      if (
        subjectDropdownRef.current &&
        !subjectDropdownRef.current.contains(e.target as Node)
      ) {
        setSubjectDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutside);
    return () => document.removeEventListener('mousedown', handleOutside);
  }, [subjectDropdownOpen]);

  const formSubjectOptions = useMemo(() => {
    const seen = new Set(subjectOptions.map((s) => s.toLowerCase()));
    const extras = selectedSubjects.filter((s) => !seen.has(s.toLowerCase()));
    return [...subjectOptions, ...extras];
  }, [subjectOptions, selectedSubjects]);

  const formClassOptions = useMemo(() => {
    const set = new Set<string>(classOptions);
    Object.values(subjectClassesMap).forEach((arr) => {
      arr.forEach((cls) => {
        if (cls.trim()) set.add(cls.trim());
      });
    });
    return Array.from(set).sort(compareClassNames);
  }, [classOptions, subjectClassesMap]);

  const classLevelGroups = useMemo(() => {
    const groups = new Map<string, string[]>();
    formClassOptions.forEach((cls) => {
      const prefix = cls.split(/[-\s]/)[0]?.trim();
      if (!prefix) return;
      const arr = groups.get(prefix) || [];
      arr.push(cls);
      groups.set(prefix, arr);
    });
    return Array.from(groups.entries()).filter(([, list]) => list.length >= 2);
  }, [formClassOptions]);

  const resetForm = () => {
    setEditingId(null);
    setNip('');
    setName('');
    setPosition('');
    setCustomPositionMode(false);
    setSelectedSubjects([]);
    setSubjectClassesMap({});
    setSubjectDropdownOpen(false);
    setCustomSubjectInput('');
    setCustomClassInputs({});
    setEmail('');
    setPhone('');
    setGender('');
    setEmploymentStatus('');
    setStatus('');
    setAddress('');
    setPhoto('');
  };

  const handlePhotoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      onToast('error', 'Ukuran foto maksimal 2 MB.');
      return;
    }
    const compressed = await compressImageFileToDataUrl(file, 420, 200000);
    setPhoto(compressed);
    onToast('success', 'Foto guru berhasil diunggah ke pratinjau. Klik Simpan untuk menyimpan.');
  };

  const handlePositionSelectChange = (val: string) => {
    if (val === '__CUSTOM__') {
      setCustomPositionMode(true);
      setPosition('');
      return;
    }
    setCustomPositionMode(false);
    setPosition(val);
    if (val === 'Guru BK' && selectedSubjects.length === 0) {
      setSelectedSubjects(['Bimbingan Konseling']);
    } else if (NON_TEACHING_POSITIONS.has(val) && selectedSubjects.length === 0) {
      setSelectedSubjects(['Umum / Non-Mapel']);
    }
  };

  const toggleSubjectSelection = (subj: string) => {
    setSelectedSubjects((prev) => {
      const lower = subj.toLowerCase();
      const exists = prev.some((s) => s.toLowerCase() === lower);
      if (exists) {
        return prev.filter((s) => s.toLowerCase() !== lower);
      }
      if (subj === 'Umum / Non-Mapel') {
        return ['Umum / Non-Mapel'];
      }
      return [...prev.filter((s) => s !== 'Umum / Non-Mapel'), subj];
    });
  };

  const handleAddCustomSubject = () => {
    const rawItems = customSubjectInput
      .split(/[,;]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    if (rawItems.length === 0) return;
    setSelectedSubjects((prev) => {
      const next = prev.filter((s) => s !== 'Umum / Non-Mapel');
      const seen = new Set(next.map((s) => s.toLowerCase()));
      rawItems.forEach((item) => {
        if (!isNonTeachingSubjectLabel(item) && !seen.has(item.toLowerCase())) {
          seen.add(item.toLowerCase());
          next.push(item);
        }
      });
      return next;
    });
    setCustomSubjectInput('');
  };

  const toggleClassForSubject = (subj: string, cls: string) => {
    setSubjectClassesMap((prev) => {
      const current = prev[subj] || [];
      const exists = current.includes(cls);
      const nextList = exists
        ? current.filter((c) => c !== cls)
        : [...current, cls].sort(compareClassNames);
      return { ...prev, [subj]: nextList };
    });
  };

  const setClassesForSubject = (subj: string, classes: string[]) => {
    const unique = Array.from(new Set(classes.map((c) => c.trim()).filter(Boolean))).sort(
      compareClassNames
    );
    setSubjectClassesMap((prev) => ({ ...prev, [subj]: unique }));
  };

  const toggleClassGroupForSubject = (subj: string, groupClasses: string[]) => {
    const current = subjectClassesMap[subj] || [];
    const allInGroupSelected = groupClasses.every((c) => current.includes(c));
    if (allInGroupSelected) {
      setClassesForSubject(
        subj,
        current.filter((c) => !groupClasses.includes(c))
      );
    } else {
      setClassesForSubject(subj, [...current, ...groupClasses]);
    }
  };

  const handleAddCustomClassForSubject = (subj: string) => {
    const raw = (customClassInputs[subj] || '')
      .split(/[,;]+/)
      .map((c) => c.trim())
      .filter(Boolean);
    if (raw.length === 0) return;
    const current = subjectClassesMap[subj] || [];
    setClassesForSubject(subj, [...current, ...raw]);
    setCustomClassInputs((prev) => ({ ...prev, [subj]: '' }));
  };

  const handleCopyClassesToAllSubjects = (sourceSubj: string) => {
    const sourceClasses = subjectClassesMap[sourceSubj] || [];
    setSubjectClassesMap((prev) => {
      const next = { ...prev };
      teachingSelectedSubjects.forEach((subj) => {
        next[subj] = [...sourceClasses];
      });
      return next;
    });
    onToast(
      'info',
      `Daftar ${sourceClasses.length} kelas dari "${sourceSubj}" disalin ke semua mata pelajaran dipilih.`
    );
  };

  const teachingSelectedSubjects = useMemo(
    () => selectedSubjects.filter((s) => !isNonTeachingSubjectLabel(s)),
    [selectedSubjects]
  );

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim() || !gender || !position.trim() || !employmentStatus || !status) {
      onToast(
        'error',
        'Nama Guru, Jenis Kelamin, Jabatan, Status Kepegawaian, dan Status Aktif wajib diisi.'
      );
      return;
    }
    const cleanPos = position.trim();
    const cleanedTeachingSubjects = teachingSelectedSubjects.map((s) => s.trim()).filter(Boolean);
    const fallbackLabel =
      selectedSubjects.includes('Umum / Non-Mapel')
        ? 'Umum / Non-Mapel'
        : cleanPos === 'Guru BK'
        ? 'Bimbingan Konseling'
        : NON_TEACHING_POSITIONS.has(cleanPos)
        ? 'Umum / Non-Mapel'
        : '-';

    const finalSubjects =
      cleanedTeachingSubjects.length > 0
        ? cleanedTeachingSubjects
        : fallbackLabel === 'Bimbingan Konseling'
        ? ['Bimbingan Konseling']
        : [];

    const cleanedSubjectClasses = normalizeTeacherSubjectClasses(
      finalSubjects,
      subjectClassesMap
    );

    const resolvedSubject =
      finalSubjects.length > 0 ? finalSubjects.join(', ') : fallbackLabel;

    const newSubjectsCount = finalSubjects.filter(
      (subj) =>
        !isNonTeachingSubjectLabel(subj) &&
        !settings.masterSubjects.some((ms) => ms.trim().toLowerCase() === subj.toLowerCase())
    ).length;

    const totalAssignedClasses = Object.values(cleanedSubjectClasses).reduce(
      (acc, list) => acc + list.length,
      0
    );

    onUpsertTeacher({
      id: editingId ?? `tch-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      nip: nip.trim(),
      name: name.trim(),
      position: cleanPos,
      subject: resolvedSubject,
      subjects: finalSubjects,
      subjectClasses: cleanedSubjectClasses,
      email: email.trim(),
      phone: phone.trim(),
      gender,
      employmentStatus,
      status,
      address: address.trim(),
      photo: normalizePhotoInputUrl(photo),
    });

    onToast(
      'success',
      editingId
        ? totalAssignedClasses > 0
          ? `Data guru diperbarui (${finalSubjects.length} mapel siap cetak).`
          : newSubjectsCount > 0
          ? `Data guru diperbarui & ${newSubjectsCount} mapel disinkronkan ke Master Mapel.`
          : 'Data guru berhasil diperbarui & tersinkronisasi.'
        : totalAssignedClasses > 0
        ? `Data guru ditambahkan (${finalSubjects.length} mapel & kelas diajar siap cetak).`
        : newSubjectsCount > 0
        ? `Data guru ditambahkan & ${newSubjectsCount} mapel disinkronkan ke Master Mapel.`
        : 'Data guru baru berhasil ditambahkan.'
    );
    resetForm();
    setFormModalOpen(false);
  };

  const handleEdit = (t: Teacher) => {
    setEditingId(t.id);
    setNip(t.nip);
    setName(t.name);
    setPosition(t.position || '');
    setCustomPositionMode(false);
    const parsedSubjects = getTeacherSubjects(t);
    if (parsedSubjects.length > 0) {
      setSelectedSubjects(parsedSubjects);
      setSubjectClassesMap(normalizeTeacherSubjectClasses(parsedSubjects, t.subjectClasses));
    } else if (t.subject === 'Umum / Non-Mapel') {
      setSelectedSubjects(['Umum / Non-Mapel']);
      setSubjectClassesMap({});
    } else {
      setSelectedSubjects([]);
      setSubjectClassesMap({});
    }
    setSubjectDropdownOpen(false);
    setCustomSubjectInput('');
    setCustomClassInputs({});
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
        if (statusFilter && statusFilter !== 'ALL' && t.status !== statusFilter) return false;
        if (q) {
          const assignmentSummary = formatTeacherAssignmentsSummary(t).toLowerCase();
          return (
            t.name.toLowerCase().includes(q) ||
            t.nip.toLowerCase().includes(q) ||
            t.subject.toLowerCase().includes(q) ||
            assignmentSummary.includes(q) ||
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
    const headers = ['No', ...TEACHER_IMPORT_HEADERS, 'Kelas Diajar'];
    const rows = exportSourceTeachers.map((t, idx) => {
      const subjects = getTeacherSubjects(t);
      const classSummary =
        subjects.length > 0
          ? subjects
              .map((s) => {
                const cls = getTeacherClassesForSubject(t, s);
                return cls.length > 0 ? `${s}: ${cls.join(', ')}` : '';
              })
              .filter(Boolean)
              .join(' | ') || '-'
          : '-';
      return [
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
        classSummary,
      ];
    });
    return { headers, rows };
  };

  const handleDownloadTemplate = () => {
    exportRowsToXlsx(
      'Template Guru',
      [...TEACHER_IMPORT_HEADERS, 'Link Foto', 'Kelas Diajar'],
      [
        [
          '19870615 201402 1 002',
          'Drs. Bambang Sutrisno, M.Pd.',
          'L',
          'Guru Mata Pelajaran',
          'Matematika, Fisika',
          'PNS',
          'Aktif',
          'bambang@sekolah.sch.id',
          '081234567800',
          'Jl. Kartini No. 12, Kudus',
          '',
          'Matematika: X-A, X-B | Fisika: XI-IPA 1',
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

      const rawSubjectStr = clean(4) || '-';
      const parsedSubjects = getTeacherSubjects({ subject: rawSubjectStr });
      const rawClassesCol = clean(11);
      const parsedMap: Record<string, string[]> = {};
      if (rawClassesCol && rawClassesCol !== '-') {
        rawClassesCol.split(/[|;]+/).forEach((seg) => {
          const colonIdx = seg.indexOf(':');
          if (colonIdx > 0) {
            const sName = seg.slice(0, colonIdx).trim();
            const cList = seg
              .slice(colonIdx + 1)
              .split(',')
              .map((c) => c.trim())
              .filter(Boolean);
            if (sName) parsedMap[sName] = cList;
          } else if (parsedSubjects.length > 0) {
            const cList = seg
              .split(',')
              .map((c) => c.trim())
              .filter(Boolean);
            parsedSubjects.forEach((s) => {
              parsedMap[s] = cList;
            });
          }
        });
      }

      const cleanSubjects = parsedSubjects.length > 0 ? parsedSubjects : [];
      const cleanMap = normalizeTeacherSubjectClasses(cleanSubjects, parsedMap);

      result.push({
        nip: clean(0),
        name: tName,
        gender: clean(2).toUpperCase().startsWith('P') ? 'P' : 'L',
        position: clean(3) || 'Guru Mata Pelajaran',
        subject: cleanSubjects.length > 0 ? cleanSubjects.join(', ') : rawSubjectStr,
        subjects: cleanSubjects,
        subjectClasses: cleanMap,
        employmentStatus: validEmp,
        status: validStatus,
        email: clean(7),
        phone: clean(8),
        address: clean(9),
        photo: normalizePhotoInputUrl(clean(10)),
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
    onConfirm({
      title: 'Simpan Data Impor Guru?',
      description: `Ditemukan ${parsed.length} data guru siap disimpan ke database. Apakah Anda ingin menyimpan data impor ini sekarang?`,
      confirmLabel: 'Ya, Simpan Data',
      onConfirm: () => {
        const stats = onImportTeachersBatch(parsed);
        setPasteText('');
        setImportModalOpen(false);
        onToast('success', `Impor selesai: ${stats.added} ditambah, ${stats.updated} diperbarui.`);
      },
      onCancel: () => {
        onToast('info', 'Impor data guru dibatalkan, data belum disimpan.');
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
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: '' }).map(
        (row, rIdx) => {
          const nextRow = [...row];
          const photoCellAddr = XLSX.utils.encode_cell({ r: rIdx, c: 10 });
          const photoCell = ws[photoCellAddr] as XLSX.CellObject | undefined;
          const rawVal = String(nextRow[10] ?? '').trim();
          if (photoCell?.l?.Target && !/^https?:\/\//i.test(rawVal)) {
            nextRow[10] = photoCell.l.Target;
          }
          return nextRow;
        }
      );
      const parsed = parseRowsToTeachers(rows);
      if (parsed.length === 0) {
        onToast('error', 'Tidak ditemukan baris guru valid di file.');
        return;
      }
      onConfirm({
        title: 'Simpan Data Impor Guru?',
        description: `File "${file.name}" berisi ${parsed.length} data guru valid. Apakah Anda ingin menyimpan data impor ini ke database?`,
        confirmLabel: 'Ya, Simpan Data',
        onConfirm: () => {
          const stats = onImportTeachersBatch(parsed);
          setImportModalOpen(false);
          onToast('success', `Impor "${file.name}": ${stats.added} ditambah, ${stats.updated} diperbarui.`);
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
          {onSyncFromAccounts && accounts.length > 0 && (
            <button
              type="button"
              onClick={handleSyncFromAccounts}
              className="px-3.5 py-2 rounded-lg border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs sm:text-sm font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
              title="Sinkronkan data akun, jabatan, dan mata pelajaran antara Manajemen Akun, Data Guru, dan Master Mapel"
            >
              <RefreshCw className="w-4 h-4" />
              <span>Sinkron dari Akun</span>
            </button>
          )}
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
          <div className="flex flex-col sm:flex-row sm:items-end sm:justify-between gap-3">
            <div>
              <h2 className="text-base font-bold text-slate-900">
                Daftar Guru Terdaftar ({filteredTeachers.length})
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Pilih guru untuk hapus massal, ekspor XLSX/XLS/Word/PDF, atau cetak resmi
              </p>
            </div>

            <div className="flex flex-wrap items-end gap-2.5">
              <div>
                <label htmlFor="tch-filter-status" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Status Guru
                </label>
                <select
                  id="tch-filter-status"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                  aria-label="Filter status guru"
                  className="px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                >
                  <option value="">--Pilih Status--</option>
                  <option value="ALL">Semua Status</option>
                  <option value="Aktif">Aktif</option>
                  <option value="Cuti">Cuti</option>
                  <option value="Nonaktif">Nonaktif</option>
                </select>
              </div>

              <div>
                <label htmlFor="tch-filter-search" className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Cari Guru
                </label>
                <div className="relative">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    id="tch-filter-search"
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Cari nama / NIP / mapel..."
                    className="pl-9 pr-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600 w-48 sm:w-56"
                  />
                </div>
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
                onToast('info', 'Menyiapkan tampilan cetak data guru...');
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
                    <p>Belum ada data guru yang sesuai filter.</p>
                    {teachers.length === 0 && onSyncFromAccounts && accounts.length > 0 && (
                      <button
                        type="button"
                        onClick={handleSyncFromAccounts}
                        className="mt-3 px-3.5 py-2 rounded-lg border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-semibold inline-flex items-center gap-1.5"
                      >
                        <RefreshCw className="w-3.5 h-3.5" />
                        <span>Sinkronkan {accounts.length} Akun dari Manajemen Akun</span>
                      </button>
                    )}
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
                        <PersonAvatar
                          src={t.photo}
                          fallbackSrc={getTeacherDummyPhoto(t.gender, t.position)}
                          alt={t.name}
                          className="w-8 h-8 rounded-full object-cover object-top border border-slate-200 bg-slate-100"
                        />
                      </td>
                      <td className="px-3 py-2.5 font-mono-tabular text-xs text-slate-600 whitespace-nowrap">
                        {t.nip || '-'}
                      </td>
                      <td className="px-3 py-2.5 font-medium text-slate-900 whitespace-nowrap">
                        {t.name}
                      </td>
                      <td className="px-3 py-2.5 text-xs font-semibold">{t.gender}</td>
                      <td className="px-3 py-2.5 text-xs whitespace-nowrap">{t.position}</td>
                      <td className="px-3 py-2.5 text-xs min-w-[190px]">
                        {(() => {
                          const subjs = getTeacherSubjects(t);
                          if (subjs.length === 0) {
                            return <span className="text-slate-600">{t.subject || '-'}</span>;
                          }
                          return (
                            <div className="space-y-1">
                              {subjs.map((subj) => {
                                const clsList = getTeacherClassesForSubject(t, subj);
                                return (
                                  <div key={subj} className="flex flex-wrap items-center gap-1">
                                    <span className="font-semibold text-slate-800">{subj}</span>
                                    {clsList.length > 0 && (
                                      <span className="text-[11px] text-indigo-700 bg-indigo-50 border border-indigo-200 rounded px-1.5 py-0.5">
                                        {clsList.join(', ')}
                                      </span>
                                    )}
                                  </div>
                                );
                              })}
                            </div>
                          );
                        })()}
                      </td>
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
                          {onOpenPrintForTeacher && (
                            <button
                              type="button"
                              onClick={() => onOpenPrintForTeacher(t)}
                              title="Buka Pusat Laporan Resmi siap cetak sesuai mapel & kelas guru ini"
                              className="px-2.5 py-1.5 rounded-md border border-indigo-200 bg-indigo-50 text-xs font-semibold text-indigo-700 hover:bg-indigo-100 inline-flex items-center gap-1"
                            >
                              <Printer className="w-3.5 h-3.5" />
                              Siap Cetak
                            </button>
                          )}
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
                onChange={(e) => setGender(e.target.value as 'L' | 'P' | '')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Jenis Kelamin--</option>
                <option value="L">Laki-laki (L)</option>
                <option value="P">Perempuan (P)</option>
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div>
              <label htmlFor="tch-pos" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Jabatan / Tugas Tambahan <span className="text-rose-600">*</span>
              </label>
              {customPositionMode ? (
                <div className="flex items-center gap-1.5">
                  <input
                    id="tch-pos"
                    type="text"
                    value={position}
                    onChange={(e) => setPosition(e.target.value)}
                    placeholder="Ketik jabatan lainnya..."
                    autoFocus
                    className="flex-1 min-w-0 px-3 py-2 text-sm border border-indigo-300 rounded-lg focus:outline-none focus:border-indigo-600"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      setCustomPositionMode(false);
                      if (!positionOptions.includes(position.trim())) {
                        setPosition('');
                      }
                    }}
                    className="px-2.5 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-medium text-slate-600 shrink-0"
                    title="Kembali pilih dari daftar jabatan"
                  >
                    Daftar
                  </button>
                </div>
              ) : (
                <select
                  id="tch-pos"
                  value={position}
                  onChange={(e) => handlePositionSelectChange(e.target.value)}
                  className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
                >
                  <option value="">--Pilih Jabatan / Tugas--</option>
                  {positionOptions.map((pos) => (
                    <option key={pos} value={pos}>
                      {pos}
                    </option>
                  ))}
                  <option value="__CUSTOM__">+ Ketik Jabatan Lainnya...</option>
                </select>
              )}
            </div>

            <div className="relative" ref={subjectDropdownRef}>
              <span className="block text-xs font-semibold text-slate-700 mb-1.5">
                Mata Pelajaran Diampu (Multi-Select)
              </span>
              <button
                id="tch-subj"
                type="button"
                onClick={() => setSubjectDropdownOpen((o) => !o)}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white text-left flex items-center justify-between gap-2 focus:outline-none focus:border-indigo-600"
              >
                <span className="truncate text-slate-800">
                  {selectedSubjects.length === 0
                    ? '--Pilih Mapel--'
                    : selectedSubjects.join(', ')}
                </span>
                <ChevronDown className="w-4 h-4 text-slate-400 shrink-0" />
              </button>

              {subjectDropdownOpen && (
                <div className="absolute left-0 mt-1.5 z-30 w-full min-w-[270px] bg-white border border-slate-200 rounded-xl shadow-xl p-2 space-y-1.5">
                  <div className="flex items-center justify-between px-2 py-1 border-b border-slate-100 text-xs">
                    <button
                      type="button"
                      onClick={() =>
                        setSelectedSubjects(
                          formSubjectOptions.filter((s) => !isNonTeachingSubjectLabel(s))
                        )
                      }
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

                  <div className="max-h-52 overflow-y-auto space-y-0.5 pr-0.5">
                    <button
                      type="button"
                      onClick={() => {
                        setSelectedSubjects([]);
                        setSubjectDropdownOpen(false);
                      }}
                      className={`w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                        selectedSubjects.length === 0
                          ? 'bg-indigo-50 text-indigo-700 font-semibold'
                          : 'text-slate-600 hover:bg-slate-50'
                      }`}
                    >
                      --Pilih Mapel--
                    </button>
                    {formSubjectOptions.map((subj) => {
                      const isChecked = selectedSubjects.some(
                        (s) => s.toLowerCase() === subj.toLowerCase()
                      );
                      return (
                        <label
                          key={subj}
                          className="flex items-center gap-2.5 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 cursor-pointer text-xs sm:text-sm text-slate-800"
                        >
                          <input
                            type="checkbox"
                            checked={isChecked}
                            onChange={() => toggleSubjectSelection(subj)}
                            className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                          />
                          <span className="truncate">{subj}</span>
                        </label>
                      );
                    })}
                  </div>

                  <div className="pt-1.5 border-t border-slate-100 flex items-center gap-1.5">
                    <input
                      type="text"
                      value={customSubjectInput}
                      onChange={(e) => setCustomSubjectInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          handleAddCustomSubject();
                        }
                      }}
                      placeholder="+ Ketik mapel baru..."
                      className="flex-1 min-w-0 px-2.5 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                    />
                    <button
                      type="button"
                      onClick={handleAddCustomSubject}
                      className="px-2.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shrink-0"
                    >
                      Tambah
                    </button>
                  </div>
                </div>
              )}
            </div>

            <div>
              <label htmlFor="tch-emp" className="block text-xs font-semibold text-slate-700 mb-1.5">
                Status Kepegawaian
              </label>
              <select
                id="tch-emp"
                value={employmentStatus}
                onChange={(e) => setEmploymentStatus(e.target.value as TeacherEmploymentStatus | '')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Kepegawaian--</option>
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
                onChange={(e) => setStatus(e.target.value as 'Aktif' | 'Cuti' | 'Nonaktif' | '')}
                className="w-full px-3 py-2 text-sm border border-slate-200 rounded-lg bg-white focus:outline-none focus:border-indigo-600"
              >
                <option value="">--Pilih Status--</option>
                <option value="Aktif">Aktif</option>
                <option value="Cuti">Cuti</option>
                <option value="Nonaktif">Nonaktif</option>
              </select>
            </div>
          </div>

          {teachingSelectedSubjects.length > 0 && (
            <div className="rounded-xl border border-indigo-200 bg-indigo-50/40 p-3.5 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  <BookOpen className="w-4 h-4 text-indigo-600 shrink-0" />
                  <div>
                    <p className="text-xs font-bold text-slate-900">
                      Kelas yang Diajar per Mata Pelajaran ({teachingSelectedSubjects.length} Mapel)
                    </p>
                    <p className="text-[11px] text-slate-500">
                      Pilih kelas untuk setiap mata pelajaran yang diampu agar laporan langsung siap cetak
                    </p>
                  </div>
                </div>
              </div>

              <div className="space-y-2.5">
                {teachingSelectedSubjects.map((subj) => {
                  const selectedCls = subjectClassesMap[subj] || [];
                  return (
                    <div
                      key={subj}
                      className="rounded-lg border border-slate-200 bg-white p-3 space-y-2.5"
                    >
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="text-xs font-bold text-slate-900">{subj}</span>
                          <span className="text-[11px] font-semibold px-2 py-0.5 rounded border border-indigo-200 bg-indigo-50 text-indigo-700">
                            {selectedCls.length > 0
                              ? `${selectedCls.length} Kelas: ${selectedCls.join(', ')}`
                              : 'Belum pilih kelas'}
                          </span>
                        </div>

                        <div className="flex flex-wrap items-center gap-2 text-[11px]">
                          {classLevelGroups.map(([prefix, groupCls]) => {
                            const allSelected = groupCls.every((c) => selectedCls.includes(c));
                            return (
                              <button
                                key={prefix}
                                type="button"
                                onClick={() => toggleClassGroupForSubject(subj, groupCls)}
                                className={`px-2 py-0.5 rounded border font-semibold transition-colors ${
                                  allSelected
                                    ? 'bg-indigo-600 text-white border-indigo-600'
                                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                                }`}
                              >
                                Kelas {prefix}
                              </button>
                            );
                          })}
                          <button
                            type="button"
                            onClick={() => setClassesForSubject(subj, formClassOptions)}
                            className="font-semibold text-indigo-600 hover:underline"
                          >
                            Pilih Semua
                          </button>
                          <button
                            type="button"
                            onClick={() => setClassesForSubject(subj, [])}
                            className="text-slate-500 hover:text-slate-700"
                          >
                            Kosongkan
                          </button>
                          {teachingSelectedSubjects.length > 1 && selectedCls.length > 0 && (
                            <button
                              type="button"
                              onClick={() => handleCopyClassesToAllSubjects(subj)}
                              className="inline-flex items-center gap-1 font-semibold text-emerald-700 hover:underline"
                              title="Salin kelas ini ke semua mata pelajaran lain yang dipilih"
                            >
                              <Copy className="w-3 h-3" />
                              Samakan ke Semua Mapel
                            </button>
                          )}
                        </div>
                      </div>

                      {formClassOptions.length > 0 ? (
                        <div className="grid grid-cols-3 sm:grid-cols-6 lg:grid-cols-8 gap-1.5 max-h-36 overflow-y-auto pr-0.5">
                          {formClassOptions.map((cls) => {
                            const checked = selectedCls.includes(cls);
                            return (
                              <label
                                key={`${subj}-${cls}`}
                                className={`flex items-center gap-1.5 px-2 py-1.5 rounded-lg border text-xs cursor-pointer transition-colors select-none ${
                                  checked
                                    ? 'bg-indigo-50 border-indigo-400 text-indigo-900 font-semibold'
                                    : 'bg-slate-50/70 border-slate-200 text-slate-700 hover:bg-slate-100'
                                }`}
                              >
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => toggleClassForSubject(subj, cls)}
                                  className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                                />
                                <span className="truncate">{cls}</span>
                              </label>
                            );
                          })}
                        </div>
                      ) : (
                        <p className="text-[11px] text-slate-500">
                          Belum ada daftar kelas master. Tambahkan kelas di bawah atau melalui Data Siswa / Pengaturan.
                        </p>
                      )}

                      <div className="flex items-center gap-1.5 pt-1">
                        <input
                          type="text"
                          value={customClassInputs[subj] || ''}
                          onChange={(e) =>
                            setCustomClassInputs((prev) => ({ ...prev, [subj]: e.target.value }))
                          }
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              handleAddCustomClassForSubject(subj);
                            }
                          }}
                          placeholder={`+ Tambah kelas lain untuk ${subj} (contoh: X-A, X-B)...`}
                          className="flex-1 min-w-0 px-2.5 py-1.5 text-xs border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                        />
                        <button
                          type="button"
                          onClick={() => handleAddCustomClassForSubject(subj)}
                          className="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 shrink-0 inline-flex items-center gap-1"
                        >
                          <Check className="w-3.5 h-3.5 text-indigo-600" />
                          Tambah Kelas
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
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
          </div>

          <div>
            <label htmlFor="tch-photo-url" className="block text-xs font-semibold text-slate-700 mb-1.5">
              Foto Pendidik / Staf TU (Link Google Drive / Unggah File)
            </label>
            <div className="flex items-center gap-3">
              <PersonAvatar
                src={photo}
                fallbackSrc={getTeacherDummyPhoto(gender === 'P' ? 'P' : 'L', position)}
                alt={name || 'Preview foto guru/TU'}
                className="w-12 h-12 rounded-xl object-cover object-top border border-slate-200 bg-slate-100 shrink-0"
              />
              <div className="flex-1 flex flex-wrap items-center gap-2 min-w-0">
                <input
                  id="tch-photo-url"
                  type="text"
                  value={photo.startsWith('data:') ? '' : photo}
                  onChange={(e) => setPhoto(e.target.value)}
                  placeholder={
                    photo.startsWith('data:')
                      ? 'Foto file lokal terunggah (tempel link Drive untuk mengganti)'
                      : 'Tempel link Google Drive / URL foto...'
                  }
                  className="flex-1 min-w-[200px] px-3 py-2 text-xs sm:text-sm border border-slate-200 rounded-lg focus:outline-none focus:border-indigo-600"
                />
                <label className="cursor-pointer px-3 py-2 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-xs font-semibold text-slate-700 inline-flex items-center gap-1.5 whitespace-nowrap">
                  <ImagePlus className="w-4 h-4 text-indigo-600" />
                  <span>{photo ? 'Ganti File' : 'Unggah Foto'}</span>
                  <input type="file" accept="image/*" className="hidden" onChange={handlePhotoUpload} />
                </label>
                {photo && (
                  <button
                    type="button"
                    onClick={() => {
                      setPhoto('');
                      onToast('info', 'Foto guru dihapus dari pratinjau.');
                    }}
                    className="px-2.5 py-2 rounded-lg border border-rose-200 text-xs text-rose-600 hover:bg-rose-50 whitespace-nowrap"
                  >
                    Hapus
                  </button>
                )}
              </div>
            </div>
            <p className="text-[11px] text-slate-500 mt-1">
              {photo
                ? 'Menampilkan foto dari Google Drive / unggahan (otomatis memakai foto dummy jika gagal dimuat).'
                : 'Belum ada foto — otomatis menampilkan foto dummy Guru atau Staf TU sesuai jabatan dan jenis kelamin.'}
            </p>
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
        subtitle="Urutan kolom: NIP, Nama Lengkap, L/P, Jabatan, Mata Pelajaran, Kepegawaian, Status, Email, No HP, Alamat, Link Foto (opsional)"
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
