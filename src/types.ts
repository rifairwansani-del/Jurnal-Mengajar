export type MenuId =
  | 'dashboard'
  | 'journal'
  | 'schedule'
  | 'grades'
  | 'students'
  | 'teachers'
  | 'reports'
  | 'accounts'
  | 'settings';

export type ThemeId =
  | 'slate-indigo'
  | 'royal-blue'
  | 'emerald-forest'
  | 'crimson-charcoal'
  | 'midnight-dark';

export type AttendanceStatus = 'H' | 'S' | 'I' | 'A' | 'D';

export type GradeRowMarker = 'neutral' | 'green' | 'yellow' | 'red' | 'blue';

export type SemesterType = 'Ganjil' | 'Genap';

export type StudentStatus = 'Aktif' | 'Pindah' | 'Lulus' | 'Nonaktif';

export type TeacherEmploymentStatus = 'PNS' | 'PPPK' | 'Tetap Yayasan' | 'Honorer';

export type AccountRole =
  | 'Admin'
  | 'Guru Mapel'
  | 'Wali Kelas'
  | 'Guru BK'
  | 'Kepala Sekolah';

export type DayOfWeek =
  | 'Senin'
  | 'Selasa'
  | 'Rabu'
  | 'Kamis'
  | 'Jumat'
  | 'Sabtu';

export type PaperOrientation = 'portrait' | 'landscape';

export type PaperSizeId = 'A4' | 'F4' | 'Letter' | 'Legal' | 'A5';

export type PrintFontFamily =
  | 'Times New Roman'
  | 'Tahoma'
  | 'Arial'
  | 'Verdana'
  | 'Calibri'
  | 'Georgia';

export type ReportSectionId =
  | 'journal'
  | 'attendance'
  | 'attendance_monthly'
  | 'attendance_monitoring'
  | 'grades'
  | 'grades_remedial'
  | 'grades_leger'
  | 'students_roster'
  | 'teacher_activity'
  | 'charts';
export type ReportType = ReportSectionId | 'package';

export type ReportChartId =
  | 'bar_class'
  | 'line_trend'
  | 'doughnut_attendance'
  | 'hbar_subject'
  | 'pie_kkm';

export interface ReportPrintPreferences {
  reportType?: ReportType;
  selectedReportTypes: ReportSectionId[];
  selectedChartTypes?: ReportChartId[];
  selectedClasses: string[];
  selectedSubjects?: string[];
  selectedTeacherId?: string;
  subjectFilter: string;
  startDate: string;
  endDate: string;
  signerAccountId: string;
  showKop: boolean;
  showTtdKepsek: boolean;
  showTtdGuru: boolean;
  showSignatureImage: boolean;
  orientation: PaperOrientation;
  paperSize: PaperSizeId;
  fontFamily: PrintFontFamily;
}

export type SignatureLayoutPreset =
  | 'kiri-kanan'
  | 'kanan-kiri'
  | 'kiri'
  | 'tengah'
  | 'kanan'
  | 'bebas';

export interface SignaturePositionSettings {
  preset: SignatureLayoutPreset;
  principalX: number; // 0-82 (% from left)
  principalY: number; // 0-240 (px top offset)
  teacherX: number;   // 0-82 (% from left)
  teacherY: number;   // 0-240 (px top offset)
  showStamp: boolean;
  stampTarget: 'guru' | 'kepsek' | 'keduanya';
  principalSignatureSize: number; // px height for Kepsek signature (16-260)
  teacherSignatureSize: number;   // px height for Guru signature (16-260)
  stampSize: number;              // px square size for school stamp (20-260)
  principalSigOffsetX: number;    // px free horizontal offset for Kepsek signature image
  principalSigOffsetY: number;    // px free vertical offset for Kepsek signature image
  teacherSigOffsetX: number;      // px free horizontal offset for Guru signature image
  teacherSigOffsetY: number;      // px free vertical offset for Guru signature image
  stampOffsetX: number;           // px free horizontal offset for school stamp image
  stampOffsetY: number;           // px free vertical offset for school stamp image
  signatureSize?: number;
}

export interface PortalIdentitySettings {
  webTitle: string;
  webSubtitle: string;
  sidebarTitle: string;
  sidebarSubtitle: string;
  sidebarLogo: string;
  loginTitle: string;
  loginSubtitle: string;
  loginFooterNote: string;
}

export interface Student {
  id: string;
  nis: string;
  nisn: string;
  name: string;
  className: string;
  absentNo: string;
  birthPlace: string;
  birthDate: string;
  email: string;
  photo: string;
  address: string;
  phone: string;
  gender: 'L' | 'P';
  status: StudentStatus;
  homeroomTeacher: string;
  counselorTeacher: string;
  entryYear: string;
}

export interface Teacher {
  id: string;
  nip: string;
  name: string;
  position: string;
  subject: string;
  subjects?: string[];
  subjectClasses?: Record<string, string[]>;
  email: string;
  phone: string;
  gender: 'L' | 'P';
  employmentStatus: TeacherEmploymentStatus;
  status: 'Aktif' | 'Cuti' | 'Nonaktif';
  address: string;
  photo: string;
}

export interface TeachingSchedule {
  id: string;
  day: DayOfWeek;
  jamKe: string;
  timeRange: string;
  className: string;
  subject: string;
  teacherName: string;
  room: string;
  academicYear: string;
  semester: SemesterType;
  notes: string;
}

export interface UserAccount {
  id: string;
  name: string;
  email: string;
  username?: string;
  hasPassword?: boolean;
  role: AccountRole;
  nip: string;
  status: 'Aktif' | 'Nonaktif';
  lastLogin: string;
  signatureImage: string;
  schoolStampImage: string;
}

export interface AttendanceRecord {
  status: AttendanceStatus;
  note: string;
}

export interface JournalEntry {
  id: string;
  date: string; // YYYY-MM-DD
  jamKe: string;
  className: string;
  subject: string;
  topic: string;
  activity: string;
  academicYear: string;
  semester: SemesterType;
  attendance: Record<string, AttendanceRecord>;
  createdAt: string;
}

export interface GradeEntry {
  id: string; // sanitized `${academicYear}__${semester}__${className}__${subject}__${examType}__${studentId}`
  className: string;
  subject: string;
  examType: string;
  studentId: string;
  score: number | null;
  note: string;
  marker: GradeRowMarker;
  academicYear: string;
  semester: SemesterType;
  updatedAt: string;
}

export type KopTextAlign = 'left' | 'center' | 'right' | 'justify';

export type KopLineKey = 'line1' | 'line2' | 'line3' | 'line4' | 'line5' | 'line6';

export interface KopLineStyle {
  fontSize: number; // pt (7-28)
  bold: boolean;
  italic: boolean;
  underline: boolean;
  align: KopTextAlign;
}

export interface KopSuratSettings {
  line1: string;
  line2: string;
  line3: string;
  line4: string;
  line5: string;
  line6: string;
  logoLeft: string;
  logoRight: string;
  logoLeftSize?: number;  // px width/height (28-150, default 72)
  logoRightSize?: number; // px width/height (28-150, default 72)
  fontFamily?: PrintFontFamily;
  lineSpacing?: number;   // CSS line-height multiplier (0.85-2.4, default 1.25)
  lineStyles?: Partial<Record<KopLineKey, KopLineStyle>>;
}

export interface SystemSettings {
  academicYear: string;
  semester: SemesterType;
  signatureLocation: string;
  principalTitle: string;
  principalName: string;
  principalNip: string;
  principalSignatureImage: string;
  defaultSchoolStampImage: string;
  teacherName: string;
  teacherNip: string;
  schoolName: string;
  email: string;
  themeId: ThemeId;
  identity: PortalIdentitySettings;
  signaturePosition: SignaturePositionSettings;
  reportPrintPreferences: ReportPrintPreferences;
  kopSurat: KopSuratSettings;
  masterClasses: string[];
  masterSubjects: string[];
  masterExamTypes: string[];
  masterAcademicYears: string[];
}

export interface SessionState {
  isLoggedIn: boolean;
  authMethod?: 'google' | 'credentials';
  accountId?: string;
  username?: string;
  email: string;
  userName?: string;
  userRole?: AccountRole;
  academicYear: string;
  semester: SemesterType;
}

export interface PortalDatabase {
  students: Student[];
  teachers: Teacher[];
  schedules: TeachingSchedule[];
  accounts: UserAccount[];
  journals: JournalEntry[];
  grades: GradeEntry[];
  settings: SystemSettings;
  session: SessionState;
}
