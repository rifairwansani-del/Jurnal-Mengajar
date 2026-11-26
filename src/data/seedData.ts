import {
  AttendanceRecord,
  AttendanceStatus,
  GradeEntry,
  GradeRowMarker,
  JournalEntry,
  PortalDatabase,
  Student,
  SystemSettings,
  Teacher,
  TeachingSchedule,
  UserAccount,
} from '../types';

const DEMO_LOGO_LEFT = `data:image/svg+xml;utf8,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120">
    <path d="M60 8 L104 24 V60 C104 88 84 104 60 114 C36 104 16 88 16 60 V24 Z" fill="#1e293b" stroke="#4f46e5" stroke-width="5"/>
    <path d="M60 18 L94 31 V59 C94 81 78 94 60 103 C42 94 26 81 26 59 V31 Z" fill="#ffffff"/>
    <circle cx="60" cy="48" r="14" fill="#4f46e5"/>
    <path d="M36 74 Q60 88 84 74" fill="none" stroke="#1e293b" stroke-width="5" stroke-linecap="round"/>
    <text x="60" y="94" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#1e293b" text-anchor="middle">PEMPROV</text>
  </svg>`
)}`;

const DEMO_LOGO_RIGHT = `data:image/svg+xml;utf8,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120">
    <circle cx="60" cy="60" r="52" fill="#4f46e5" stroke="#1e293b" stroke-width="4"/>
    <circle cx="60" cy="60" r="42" fill="#ffffff"/>
    <path d="M40 44 H80 V78 H40 Z" fill="none" stroke="#1e293b" stroke-width="4"/>
    <path d="M60 44 V78" stroke="#1e293b" stroke-width="3"/>
    <path d="M46 54 H55 M46 62 H55 M65 54 H74 M65 62 H74" stroke="#4f46e5" stroke-width="3" stroke-linecap="round"/>
    <text x="60" y="34" font-family="Arial, sans-serif" font-size="10" font-weight="bold" fill="#1e293b" text-anchor="middle">SMAN 1</text>
    <text x="60" y="94" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#1e293b" text-anchor="middle">UNGGUL</text>
  </svg>`
)}`;

export const DEMO_SCHOOL_STAMP = `data:image/svg+xml;utf8,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 160 160">
    <circle cx="80" cy="80" r="72" fill="none" stroke="#4338ca" stroke-width="4" opacity="0.82"/>
    <circle cx="80" cy="80" r="64" fill="none" stroke="#4338ca" stroke-width="1.5" opacity="0.82"/>
    <circle cx="80" cy="80" r="42" fill="none" stroke="#4338ca" stroke-width="2" opacity="0.82"/>
    <text x="80" y="34" font-family="Arial, sans-serif" font-size="10" font-weight="bold" fill="#4338ca" text-anchor="middle" opacity="0.85">DINAS PENDIDIKAN</text>
    <text x="80" y="76" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#4338ca" text-anchor="middle" opacity="0.85">SMA NEGERI 1</text>
    <text x="80" y="92" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#4338ca" text-anchor="middle" opacity="0.85">KUDUS</text>
    <text x="80" y="136" font-family="Arial, sans-serif" font-size="9" font-weight="bold" fill="#4338ca" text-anchor="middle" opacity="0.85">KABUPATEN KUDUS</text>
  </svg>`
)}`;

function makeSignatureSvg(pathD: string, color = '#0f172a'): string {
  return `data:image/svg+xml;utf8,${encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 220 90">
      <path d="${pathD}" fill="none" stroke="${color}" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>`
  )}`;
}

const DEMO_SIG_GURU = makeSignatureSvg(
  'M26 62 C38 24, 54 18, 52 52 C50 72, 68 32, 82 38 C94 44, 78 66, 98 54 C118 42, 134 28, 146 48 C156 64, 178 44, 198 46 M32 72 Q115 62 190 68'
);

const DEMO_SIG_KEPSEK = makeSignatureSvg(
  'M22 56 C44 16, 62 20, 56 54 C52 70, 76 22, 92 42 C104 58, 122 26, 136 44 C150 62, 168 30, 188 48 M28 74 Q108 66 196 72'
);

const DEMO_SIG_AMINAH = makeSignatureSvg(
  'M24 58 Q48 14 64 56 T108 48 Q134 26 152 54 T192 44 M30 70 H182'
);

const DEMO_SIG_RATNA = makeSignatureSvg(
  'M28 52 C42 22, 68 24, 62 56 C58 72, 94 28, 114 48 C132 66, 152 24, 184 48 M26 70 Q105 76 186 64'
);

export function createDefaultSettings(): SystemSettings {
  return {
    academicYear: '2025/2026',
    semester: 'Ganjil',
    signatureLocation: 'Kudus',
    principalTitle: 'Kepala Sekolah',
    principalName: 'Drs. H. Sudirman, M.Pd.',
    principalNip: '19680512 199403 1 008',
    principalSignatureImage: DEMO_SIG_KEPSEK,
    defaultSchoolStampImage: DEMO_SCHOOL_STAMP,
    teacherName: 'Rifa Irwansani, S.Pd., M.Pd.',
    teacherNip: '19890418 201502 1 003',
    schoolName: 'SMA Negeri 1 Kudus',
    email: 'guru@sma1kudus.sch.id',
    password: 'guru123',
    themeId: 'slate-indigo',
    identity: {
      webTitle: 'Portal E-Guru v10.2 — Sistem Administrasi Mengajar Terpadu',
      webSubtitle: 'Sistem Jurnal Mengajar, Presensi & Penilaian Guru Terpadu',
      sidebarTitle: 'Portal E-Guru v10.2',
      sidebarSubtitle: 'Administrasi Mengajar Lokal',
      sidebarLogo: '',
      loginTitle: 'Selamat Datang',
      loginSubtitle: 'Portal E-Guru v10.2 · Masuk untuk melanjutkan',
      loginFooterNote: 'Penyimpanan Lokal Browser · Mandiri & Siap Cetak',
    },
    signaturePosition: {
      preset: 'kiri-kanan',
      principalX: 6,
      principalY: 0,
      teacherX: 58,
      teacherY: 0,
      showStamp: true,
      stampTarget: 'guru',
      principalSignatureSize: 60,
      teacherSignatureSize: 60,
      stampSize: 70,
      principalSigOffsetX: 0,
      principalSigOffsetY: 0,
      teacherSigOffsetX: 0,
      teacherSigOffsetY: 0,
      stampOffsetX: 0,
      stampOffsetY: 0,
    },
    reportPrintPreferences: {
      reportType: 'journal',
      selectedReportTypes: ['journal'],
      selectedClasses: ['X-A', 'X-B', 'XI-IPA 1', 'XI-IPA 2'],
      subjectFilter: 'ALL',
      startDate: '',
      endDate: '',
      signerAccountId: '',
      showKop: true,
      showTtdKepsek: true,
      showTtdGuru: true,
      showSignatureImage: true,
      orientation: 'landscape',
      paperSize: 'A4',
      fontFamily: 'Times New Roman',
    },
    kopSurat: {
      line1: 'PEMERINTAH PROVINSI JAWA TENGAH',
      line2: 'DINAS PENDIDIKAN DAN KEBUDAYAAN',
      line3: 'SMA NEGERI 1 KUDUS',
      line4: 'Jl. Pramuka No. 41 Mlati Lor, Kec. Kota, Kab. Kudus 59319',
      line5: 'Telepon (0291) 431368 · Laman: https://sma1kudus.sch.id',
      line6: 'Email: info@sma1kudus.sch.id · NPSN: 20317497',
      logoLeft: DEMO_LOGO_LEFT,
      logoRight: DEMO_LOGO_RIGHT,
    },
    masterClasses: ['X-A', 'X-B', 'XI-IPA 1', 'XI-IPA 2'],
    masterSubjects: ['Matematika Wajib', 'Matematika Peminatan', 'Informatika'],
    masterExamTypes: ['Tugas 1', 'UH 1', 'UTS', 'Tugas 2', 'UH 2', 'UAS'],
    masterAcademicYears: ['2024/2025', '2025/2026', '2026/2027'],
  };
}

export function createEmptyDatabase(): PortalDatabase {
  const settings = createDefaultSettings();
  return {
    students: [],
    teachers: [],
    schedules: [],
    accounts: [
      {
        id: 'acc-default',
        name: settings.teacherName,
        email: settings.email,
        password: settings.password,
        role: 'Admin',
        nip: settings.teacherNip,
        status: 'Aktif',
        lastLogin: '',
        signatureImage: DEMO_SIG_GURU,
        schoolStampImage: DEMO_SCHOOL_STAMP,
      },
    ],
    journals: [],
    grades: [],
    settings,
    session: {
      isLoggedIn: false,
      accountId: 'acc-default',
      email: settings.email,
      userName: settings.teacherName,
      userRole: 'Admin',
      academicYear: settings.academicYear,
      semester: settings.semester,
    },
  };
}

export function createDemoDatabase(): PortalDatabase {
  const settings = createDefaultSettings();

  const teachers: Teacher[] = [
    {
      id: 'tch-01',
      nip: '19890418 201502 1 003',
      name: 'Rifa Irwansani, S.Pd., M.Pd.',
      position: 'Guru Mapel & Admin Kurikulum',
      subject: 'Matematika Wajib',
      email: 'guru@sma1kudus.sch.id',
      phone: '081234567801',
      gender: 'L',
      employmentStatus: 'PNS',
      status: 'Aktif',
      address: 'Jl. Diponegoro No. 18, Kudus',
      photo: '',
    },
    {
      id: 'tch-02',
      nip: '19850711 201001 2 014',
      name: 'Dra. Siti Aminah, M.Pd.',
      position: 'Wali Kelas X-A',
      subject: 'Matematika Wajib',
      email: 'aminah@sma1kudus.sch.id',
      phone: '081234567802',
      gender: 'P',
      employmentStatus: 'PNS',
      status: 'Aktif',
      address: 'Perum Muria Indah Blok C-12, Kudus',
      photo: '',
    },
    {
      id: 'tch-03',
      nip: '19910224 201903 1 009',
      name: 'Budi Hartono, S.Kom., M.Cs.',
      position: 'Wali Kelas XI-IPA 1',
      subject: 'Informatika',
      email: 'budi.hartono@sma1kudus.sch.id',
      phone: '081234567803',
      gender: 'L',
      employmentStatus: 'PPPK',
      status: 'Aktif',
      address: 'Jl. Jend. Sudirman No. 89, Kudus',
      photo: '',
    },
    {
      id: 'tch-04',
      nip: '19881103 201402 2 005',
      name: 'Ratna Kusumawati, S.Psi.',
      position: 'Guru Bimbingan Konseling (BK)',
      subject: 'Bimbingan Konseling',
      email: 'ratna.bk@sma1kudus.sch.id',
      phone: '081234567804',
      gender: 'P',
      employmentStatus: 'PNS',
      status: 'Aktif',
      address: 'Jl. Sunan Kudus No. 45, Kudus',
      photo: '',
    },
    {
      id: 'tch-05',
      nip: '19930615 202203 2 011',
      name: 'Dewi Anggraini, S.Pd.',
      position: 'Wali Kelas X-B & XI-IPA 2',
      subject: 'Matematika Peminatan',
      email: 'dewi.anggraini@sma1kudus.sch.id',
      phone: '081234567805',
      gender: 'P',
      employmentStatus: 'PPPK',
      status: 'Aktif',
      address: 'Desa Bae RT 02/03, Kudus',
      photo: '',
    },
  ];

  const accounts: UserAccount[] = [
    {
      id: 'acc-01',
      name: 'Rifa Irwansani, S.Pd., M.Pd.',
      email: 'guru@sma1kudus.sch.id',
      password: 'guru123',
      role: 'Admin',
      nip: '19890418 201502 1 003',
      status: 'Aktif',
      lastLogin: '2026-11-25 07:15',
      signatureImage: DEMO_SIG_GURU,
      schoolStampImage: DEMO_SCHOOL_STAMP,
    },
    {
      id: 'acc-02',
      name: 'Dra. Siti Aminah, M.Pd.',
      email: 'aminah@sma1kudus.sch.id',
      password: 'guru123',
      role: 'Wali Kelas',
      nip: '19850711 201001 2 014',
      status: 'Aktif',
      lastLogin: '2026-11-24 13:40',
      signatureImage: DEMO_SIG_AMINAH,
      schoolStampImage: DEMO_SCHOOL_STAMP,
    },
    {
      id: 'acc-03',
      name: 'Ratna Kusumawati, S.Psi.',
      email: 'ratna.bk@sma1kudus.sch.id',
      password: 'guru123',
      role: 'Guru BK',
      nip: '19881103 201402 2 005',
      status: 'Aktif',
      lastLogin: '2026-11-23 09:20',
      signatureImage: DEMO_SIG_RATNA,
      schoolStampImage: DEMO_SCHOOL_STAMP,
    },
    {
      id: 'acc-04',
      name: 'Drs. H. Sudirman, M.Pd.',
      email: 'kepsek@sma1kudus.sch.id',
      password: 'guru123',
      role: 'Kepala Sekolah',
      nip: '19680512 199403 1 008',
      status: 'Aktif',
      lastLogin: '2026-11-20 08:05',
      signatureImage: DEMO_SIG_KEPSEK,
      schoolStampImage: DEMO_SCHOOL_STAMP,
    },
  ];

  const schedules: TeachingSchedule[] = [
    {
      id: 'sch-01',
      day: 'Senin',
      jamKe: '1-2',
      timeRange: '07:00 - 08:30',
      className: 'X-A',
      subject: 'Matematika Wajib',
      teacherName: 'Rifa Irwansani, S.Pd., M.Pd.',
      room: 'Ruang X-A',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      notes: 'Kelas reguler pagi',
    },
    {
      id: 'sch-02',
      day: 'Senin',
      jamKe: '3-4',
      timeRange: '08:30 - 10:00',
      className: 'X-B',
      subject: 'Matematika Wajib',
      teacherName: 'Rifa Irwansani, S.Pd., M.Pd.',
      room: 'Ruang X-B',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      notes: '',
    },
    {
      id: 'sch-03',
      day: 'Selasa',
      jamKe: '1-2',
      timeRange: '07:00 - 08:30',
      className: 'XI-IPA 1',
      subject: 'Matematika Peminatan',
      teacherName: 'Dewi Anggraini, S.Pd.',
      room: 'Ruang XI-IPA 1',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      notes: '',
    },
    {
      id: 'sch-04',
      day: 'Rabu',
      jamKe: '3-4',
      timeRange: '08:30 - 10:00',
      className: 'XI-IPA 1',
      subject: 'Informatika',
      teacherName: 'Budi Hartono, S.Kom., M.Cs.',
      room: 'Lab Komputer 1',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      notes: 'Praktikum Algoritma',
    },
    {
      id: 'sch-05',
      day: 'Kamis',
      jamKe: '3-4',
      timeRange: '08:30 - 10:00',
      className: 'X-A',
      subject: 'Matematika Wajib',
      teacherName: 'Rifa Irwansani, S.Pd., M.Pd.',
      room: 'Ruang X-A',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      notes: '',
    },
    {
      id: 'sch-06',
      day: 'Jumat',
      jamKe: '5-6',
      timeRange: '10:15 - 11:45',
      className: 'XI-IPA 2',
      subject: 'Informatika',
      teacherName: 'Budi Hartono, S.Kom., M.Cs.',
      room: 'Lab Komputer 2',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      notes: 'Praktikum Struktur Data',
    },
  ];

  const rawStudents: Array<{
    id: string;
    nis: string;
    nisn: string;
    name: string;
    gender: 'L' | 'P';
    className: string;
    absentNo: string;
    birthPlace: string;
    birthDate: string;
    address: string;
    phone: string;
    homeroomTeacher: string;
    entryYear: string;
  }> = [
    { id: 'st-01', nis: '25001', nisn: '0091234501', name: 'Aditya Pratama Putra', gender: 'L', className: 'X-A', absentNo: '01', birthPlace: 'Kudus', birthDate: '2009-03-14', address: 'Jl. Sosrokartono No. 12, Kudus', phone: '085712340001', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-02', nis: '25002', nisn: '0091234502', name: 'Aisyah Nurul Hidayah', gender: 'P', className: 'X-A', absentNo: '02', birthPlace: 'Jepara', birthDate: '2009-05-22', address: 'Desa Kaliputu RT 03/01, Kudus', phone: '085712340002', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-03', nis: '25003', nisn: '0091234503', name: 'Bagus Setiawan', gender: 'L', className: 'X-A', absentNo: '03', birthPlace: 'Kudus', birthDate: '2009-01-19', address: 'Desa Rendeng RT 04/02, Kudus', phone: '085712340003', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-04', nis: '25004', nisn: '0091234504', name: 'Citra Dewi Lestari', gender: 'P', className: 'X-A', absentNo: '04', birthPlace: 'Pati', birthDate: '2009-07-08', address: 'Perum Megawon Indah No. 7, Kudus', phone: '085712340004', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-05', nis: '25005', nisn: '0091234505', name: 'Dimas Wahyu Nugroho', gender: 'L', className: 'X-A', absentNo: '05', birthPlace: 'Kudus', birthDate: '2009-09-11', address: 'Desa Getas Pejaten No. 24, Kudus', phone: '085712340005', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-06', nis: '25006', nisn: '0091234506', name: 'Eka Putri Ramadhani', gender: 'P', className: 'X-A', absentNo: '06', birthPlace: 'Semarang', birthDate: '2009-08-29', address: 'Jl. Kudus-Pati KM 4, Kudus', phone: '085712340006', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-07', nis: '25007', nisn: '0091234507', name: 'Fajar Maulana Malik', gender: 'L', className: 'X-A', absentNo: '07', birthPlace: 'Demak', birthDate: '2009-04-03', address: 'Desa Jati Wetan RT 01/04, Kudus', phone: '085712340007', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-08', nis: '25008', nisn: '0091234508', name: 'Gita Saraswati', gender: 'P', className: 'X-A', absentNo: '08', birthPlace: 'Kudus', birthDate: '2009-11-17', address: 'Jl. Pemuda No. 55, Kudus', phone: '085712340008', homeroomTeacher: 'Dra. Siti Aminah, M.Pd.', entryYear: '2025' },
    { id: 'st-09', nis: '25009', nisn: '0091234509', name: 'Hendra Wijaya Kusuma', gender: 'L', className: 'X-B', absentNo: '01', birthPlace: 'Kudus', birthDate: '2009-02-10', address: 'Desa Barongan No. 19, Kudus', phone: '085712340009', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-10', nis: '25010', nisn: '0091234510', name: 'Intan Permata Sari', gender: 'P', className: 'X-B', absentNo: '02', birthPlace: 'Kudus', birthDate: '2009-06-15', address: 'Desa Panjunan RT 02/02, Kudus', phone: '085712340010', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-11', nis: '25011', nisn: '0091234511', name: 'Joko Susilo', gender: 'L', className: 'X-B', absentNo: '03', birthPlace: 'Pati', birthDate: '2009-10-05', address: 'Desa Gondangmanis RT 05/01, Kudus', phone: '085712340011', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-12', nis: '25012', nisn: '0091234512', name: 'Kartika Anggraini', gender: 'P', className: 'X-B', absentNo: '04', birthPlace: 'Kudus', birthDate: '2009-04-21', address: 'Jl. Ahmad Yani No. 30, Kudus', phone: '085712340012', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-13', nis: '25013', nisn: '0091234513', name: 'Lukman Hakim', gender: 'L', className: 'X-B', absentNo: '05', birthPlace: 'Jepara', birthDate: '2009-12-01', address: 'Desa Purwosari RT 03/03, Kudus', phone: '085712340013', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-14', nis: '25014', nisn: '0091234514', name: 'Maya Puspita Dewi', gender: 'P', className: 'X-B', absentNo: '06', birthPlace: 'Kudus', birthDate: '2009-03-27', address: 'Desa Mlati Kidul No. 14, Kudus', phone: '085712340014', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-15', nis: '25015', nisn: '0091234515', name: 'Nanda Rizky Pratama', gender: 'L', className: 'X-B', absentNo: '07', birthPlace: 'Kudus', birthDate: '2009-07-30', address: 'Desa Karangbener RT 01/02, Kudus', phone: '085712340015', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-16', nis: '25016', nisn: '0091234516', name: 'Olivia Maharani', gender: 'P', className: 'X-B', absentNo: '08', birthPlace: 'Semarang', birthDate: '2009-09-19', address: 'Perum Gerbang Harapan No. 9, Kudus', phone: '085712340016', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2025' },
    { id: 'st-17', nis: '24001', nisn: '0081234501', name: 'Pandu Aji Santoso', gender: 'L', className: 'XI-IPA 1', absentNo: '01', birthPlace: 'Kudus', birthDate: '2008-01-12', address: 'Jl. Sunan Muria No. 21, Kudus', phone: '085712340017', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-18', nis: '24002', nisn: '0081234502', name: 'Qoriatul Khasanah', gender: 'P', className: 'XI-IPA 1', absentNo: '02', birthPlace: 'Demak', birthDate: '2008-05-04', address: 'Desa Loram Kulon RT 02/04, Kudus', phone: '085712340018', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-19', nis: '24003', nisn: '0081234503', name: 'Raka Aditya Mahendra', gender: 'L', className: 'XI-IPA 1', absentNo: '03', birthPlace: 'Kudus', birthDate: '2008-06-25', address: 'Jl. HM Subchan ZE No. 8, Kudus', phone: '085712340019', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-20', nis: '24004', nisn: '0081234504', name: 'Salsabila Azzahra', gender: 'P', className: 'XI-IPA 1', absentNo: '04', birthPlace: 'Kudus', birthDate: '2008-08-18', address: 'Desa Demaan RT 04/01, Kudus', phone: '085712340020', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-21', nis: '24005', nisn: '0081234505', name: 'Taufik Hidayatullah', gender: 'L', className: 'XI-IPA 1', absentNo: '05', birthPlace: 'Pati', birthDate: '2008-10-14', address: 'Desa Tumpangkrasak No. 11, Kudus', phone: '085712340021', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-22', nis: '24006', nisn: '0081234506', name: 'Umi Kultsum', gender: 'P', className: 'XI-IPA 1', absentNo: '06', birthPlace: 'Kudus', birthDate: '2008-03-09', address: 'Desa Pasuruhan Lor RT 02/03, Kudus', phone: '085712340022', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-23', nis: '24007', nisn: '0081234507', name: 'Vino Bastian Putra', gender: 'L', className: 'XI-IPA 1', absentNo: '07', birthPlace: 'Kudus', birthDate: '2008-11-28', address: 'Jl. KHR Asnawi No. 63, Kudus', phone: '085712340023', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-24', nis: '24008', nisn: '0081234508', name: 'Wulan Dari', gender: 'P', className: 'XI-IPA 1', absentNo: '08', birthPlace: 'Jepara', birthDate: '2008-07-02', address: 'Desa Kaliwungu RT 01/05, Kudus', phone: '085712340024', homeroomTeacher: 'Budi Hartono, S.Kom., M.Cs.', entryYear: '2024' },
    { id: 'st-25', nis: '24009', nisn: '0081234509', name: 'Xavier Ardiansyah', gender: 'L', className: 'XI-IPA 2', absentNo: '01', birthPlace: 'Semarang', birthDate: '2008-02-19', address: 'Perum Salam Residence A-3, Kudus', phone: '085712340025', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-26', nis: '24010', nisn: '0081234510', name: 'Yulia Rahmawati', gender: 'P', className: 'XI-IPA 2', absentNo: '02', birthPlace: 'Kudus', birthDate: '2008-04-16', address: 'Desa Dersalam RT 03/02, Kudus', phone: '085712340026', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-27', nis: '24011', nisn: '0081234511', name: 'Zaki Mubarak', gender: 'L', className: 'XI-IPA 2', absentNo: '03', birthPlace: 'Kudus', birthDate: '2008-09-23', address: 'Desa Ngembalrejo RT 04/01, Kudus', phone: '085712340027', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-28', nis: '24012', nisn: '0081234512', name: 'Anisa Fitriani', gender: 'P', className: 'XI-IPA 2', absentNo: '04', birthPlace: 'Pati', birthDate: '2008-12-07', address: 'Desa Honggosoco RT 02/02, Kudus', phone: '085712340028', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-29', nis: '24013', nisn: '0081234513', name: 'Bayu Segara', gender: 'L', className: 'XI-IPA 2', absentNo: '05', birthPlace: 'Kudus', birthDate: '2008-06-11', address: 'Desa Panjang RT 01/04, Kudus', phone: '085712340029', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-30', nis: '24014', nisn: '0081234514', name: 'Clarissa Amelia', gender: 'P', className: 'XI-IPA 2', absentNo: '06', birthPlace: 'Kudus', birthDate: '2008-08-30', address: 'Jl. Mejobo No. 72, Kudus', phone: '085712340030', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-31', nis: '24015', nisn: '0081234515', name: 'Daffa Alfarizi', gender: 'L', className: 'XI-IPA 2', absentNo: '07', birthPlace: 'Demak', birthDate: '2008-05-26', address: 'Desa Tanjungkarang No. 16, Kudus', phone: '085712340031', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
    { id: 'st-32', nis: '24016', nisn: '0081234516', name: 'Elsa Novita Sari', gender: 'P', className: 'XI-IPA 2', absentNo: '08', birthPlace: 'Kudus', birthDate: '2008-11-03', address: 'Desa Jepangpakis RT 03/03, Kudus', phone: '085712340032', homeroomTeacher: 'Dewi Anggraini, S.Pd.', entryYear: '2024' },
  ];

  const students: Student[] = rawStudents.map((s) => ({
    ...s,
    email: `${s.name.split(' ')[0].toLowerCase()}.${s.nis}@siswa.sma1kudus.sch.id`,
    photo: '',
    status: 'Aktif',
    counselorTeacher: 'Ratna Kusumawati, S.Psi.',
  }));

  const makeAttendance = (
    className: string,
    overrides: Record<string, { status: AttendanceStatus; note: string }>
  ): Record<string, AttendanceRecord> => {
    const map: Record<string, AttendanceRecord> = {};
    students
      .filter((s) => s.className === className)
      .forEach((s) => {
        map[s.id] = overrides[s.id] ?? { status: 'H', note: '' };
      });
    return map;
  };

  const journals: JournalEntry[] = [
    {
      id: 'jrn-01',
      date: '2026-11-02',
      jamKe: '1-2',
      className: 'X-A',
      subject: 'Matematika Wajib',
      topic: 'Eksponen & Sifat-Sifat Bilangan Berpangkat',
      activity: 'Penyampaian konsep eksponen, latihan terbimbing 8 soal, dan diskusi kelompok kecil.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-A', {
        'st-03': { status: 'S', note: 'Demam, surat orang tua' },
      }),
      createdAt: '2026-11-02T07:30:00Z',
    },
    {
      id: 'jrn-02',
      date: '2026-11-05',
      jamKe: '3-4',
      className: 'X-A',
      subject: 'Matematika Wajib',
      topic: 'Persamaan Eksponen Sederhana',
      activity: 'Pembahasan PR, kuis cepat 15 menit, dan latihan soal berbasis masalah kontekstual.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-A', {
        'st-05': { status: 'I', note: 'Izin keperluan keluarga' },
        'st-07': { status: 'D', note: 'Dispensasi lomba OSN' },
      }),
      createdAt: '2026-11-05T09:15:00Z',
    },
    {
      id: 'jrn-03',
      date: '2026-11-09',
      jamKe: '1-2',
      className: 'X-A',
      subject: 'Matematika Wajib',
      topic: 'Logaritma dan Sifat Operasi Dasar',
      activity: 'Eksplorasi hubungan eksponen-logaritma serta pengerjaan LKPD kelompok.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-A', {
        'st-03': { status: 'H', note: '' },
        'st-08': { status: 'S', note: 'Sakit gigi' },
      }),
      createdAt: '2026-11-09T07:35:00Z',
    },
    {
      id: 'jrn-04',
      date: '2026-11-12',
      jamKe: '3-4',
      className: 'X-A',
      subject: 'Matematika Wajib',
      topic: 'Barisan dan Deret Aritmetika',
      activity: 'Penurunan rumus suku ke-n dan jumlah n suku pertama melalui studi kasus menabung.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-A', {
        'st-05': { status: 'A', note: 'Tanpa keterangan' },
      }),
      createdAt: '2026-11-12T09:10:00Z',
    },
    {
      id: 'jrn-05',
      date: '2026-11-03',
      jamKe: '5-6',
      className: 'X-B',
      subject: 'Matematika Wajib',
      topic: 'Eksponen & Sifat-Sifat Bilangan Berpangkat',
      activity: 'Penjelasan materi inti, demonstrasi soal di papan tulis, dan refleksi akhir.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-B', {
        'st-11': { status: 'I', note: 'Izin mengantar keluarga' },
        'st-14': { status: 'S', note: 'Sakit flu' },
      }),
      createdAt: '2026-11-03T10:45:00Z',
    },
    {
      id: 'jrn-06',
      date: '2026-11-06',
      jamKe: '1-2',
      className: 'X-B',
      subject: 'Matematika Wajib',
      topic: 'Persamaan & Pertidaksamaan Eksponen',
      activity: 'Latihan berpasangan dan presentasi hasil penyelesaian soal di depan kelas.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-B', {
        'st-13': { status: 'A', note: 'Belum ada kabar' },
      }),
      createdAt: '2026-11-06T07:30:00Z',
    },
    {
      id: 'jrn-07',
      date: '2026-11-10',
      jamKe: '5-6',
      className: 'X-B',
      subject: 'Matematika Wajib',
      topic: 'Logaritma dan Penerapannya',
      activity: 'Diskusi soal aplikasi pertumbuhan penduduk dan skala pH.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('X-B', {
        'st-10': { status: 'D', note: 'Delegasi Paskibraka kab.' },
      }),
      createdAt: '2026-11-10T10:50:00Z',
    },
    {
      id: 'jrn-08',
      date: '2026-11-04',
      jamKe: '1-2',
      className: 'XI-IPA 1',
      subject: 'Matematika Peminatan',
      topic: 'Persamaan Trigonometri Dasar',
      activity: 'Menentukan himpunan penyelesaian sin, cos, dan tan pada interval 0°–360°.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('XI-IPA 1', {
        'st-21': { status: 'D', note: 'Pertandingan basket Popda' },
      }),
      createdAt: '2026-11-04T07:30:00Z',
    },
    {
      id: 'jrn-09',
      date: '2026-11-11',
      jamKe: '1-2',
      className: 'XI-IPA 1',
      subject: 'Matematika Peminatan',
      topic: 'Rumus Jumlah & Selisih Dua Sudut',
      activity: 'Pembuktian identitas trigonometri dan latihan soal HOTS.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('XI-IPA 1', {
        'st-18': { status: 'S', note: 'Surat dokter 1 hari' },
      }),
      createdAt: '2026-11-11T07:35:00Z',
    },
    {
      id: 'jrn-10',
      date: '2026-11-18',
      jamKe: '3-4',
      className: 'XI-IPA 1',
      subject: 'Informatika',
      topic: 'Algoritma Pencarian & Pengurutan Data',
      activity: 'Simulasi Bubble Sort dan Binary Search menggunakan Python di laboratorium komputer.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('XI-IPA 1', {}),
      createdAt: '2026-11-18T09:15:00Z',
    },
    {
      id: 'jrn-11',
      date: '2026-11-04',
      jamKe: '5-6',
      className: 'XI-IPA 2',
      subject: 'Matematika Peminatan',
      topic: 'Persamaan Trigonometri Bentuk Kuadrat',
      activity: 'Pemfaktoran fungsi trigonometri dan pendampingan individu.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('XI-IPA 2', {
        'st-29': { status: 'I', note: 'Izin upacara adat keluarga' },
        'st-31': { status: 'A', note: 'Alpa tanpa keterangan' },
      }),
      createdAt: '2026-11-04T11:00:00Z',
    },
    {
      id: 'jrn-12',
      date: '2026-11-19',
      jamKe: '5-6',
      className: 'XI-IPA 2',
      subject: 'Informatika',
      topic: 'Struktur Data Array & Rekursi',
      activity: 'Praktikum implementasi rekursi faktorial dan Fibonacci dengan analisis kompleksitas.',
      academicYear: '2025/2026',
      semester: 'Ganjil',
      attendance: makeAttendance('XI-IPA 2', {
        'st-26': { status: 'S', note: 'Izin sakit dari UKS' },
      }),
      createdAt: '2026-11-19T11:05:00Z',
    },
  ];

  const grades: GradeEntry[] = [];
  const pushClassGrades = (
    className: string,
    subject: string,
    scoresByStudentIndex: Record<string, number[]>
  ) => {
    const examTypes = ['Tugas 1', 'UH 1', 'UTS', 'Tugas 2'];
    const classStudents = students.filter((s) => s.className === className);
    classStudents.forEach((student, idx) => {
      const arr = scoresByStudentIndex[String(idx)] ?? [80, 78, 82, 84];
      examTypes.forEach((examType, eIdx) => {
        const score = arr[eIdx] ?? null;
        let marker: GradeRowMarker = 'neutral';
        if (score !== null) {
          if (score >= 90) marker = 'blue';
          else if (score >= 75) marker = 'green';
          else if (score >= 65) marker = 'yellow';
          else marker = 'red';
        }
        grades.push({
          id: `2025/2026__Ganjil__${className}__${subject}__${examType}__${student.id}`,
          className,
          subject,
          examType,
          studentId: student.id,
          score,
          note: score !== null && score < 75 ? 'Perlu remedial' : '',
          marker,
          academicYear: '2025/2026',
          semester: 'Ganjil',
          updatedAt: '2026-11-20T08:00:00Z',
        });
      });
    });
  };

  pushClassGrades('X-A', 'Matematika Wajib', {
    '0': [88, 84, 86, 90],
    '1': [92, 90, 94, 95],
    '2': [76, 72, 70, 78],
    '3': [85, 82, 80, 86],
    '4': [68, 65, 70, 74],
    '5': [90, 88, 91, 92],
    '6': [82, 80, 78, 84],
    '7': [78, 76, 80, 82],
  });

  pushClassGrades('X-B', 'Matematika Wajib', {
    '0': [80, 78, 76, 82],
    '1': [86, 88, 85, 89],
    '2': [70, 68, 72, 75],
    '3': [91, 92, 89, 94],
    '4': [64, 60, 68, 72],
    '5': [84, 82, 86, 88],
    '6': [79, 75, 77, 81],
    '7': [88, 85, 87, 90],
  });

  pushClassGrades('XI-IPA 1', 'Matematika Peminatan', {
    '0': [90, 88, 92, 91],
    '1': [84, 80, 82, 86],
    '2': [95, 93, 96, 94],
    '3': [88, 86, 85, 89],
    '4': [74, 72, 76, 78],
    '5': [82, 84, 80, 85],
    '6': [77, 79, 75, 80],
    '7': [91, 89, 90, 93],
  });

  pushClassGrades('XI-IPA 2', 'Matematika Peminatan', {
    '0': [82, 80, 78, 84],
    '1': [87, 85, 88, 86],
    '2': [72, 70, 74, 76],
    '3': [90, 92, 91, 94],
    '4': [66, 68, 71, 74],
    '5': [85, 83, 84, 87],
    '6': [79, 77, 80, 82],
    '7': [83, 86, 85, 88],
  });

  return {
    students,
    teachers,
    schedules,
    accounts,
    journals,
    grades,
    settings,
    session: {
      isLoggedIn: true,
      accountId: 'acc-01',
      email: settings.email,
      userName: settings.teacherName,
      userRole: 'Admin',
      academicYear: settings.academicYear,
      semester: settings.semester,
    },
  };
}
