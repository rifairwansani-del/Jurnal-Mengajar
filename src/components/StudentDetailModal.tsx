import { useEffect, useMemo, useState } from 'react';
import {
  Calendar,
  ClipboardCheck,
  FileText,
  MapPin,
  Pencil,
  Phone,
  User,
  Users,
  X,
} from 'lucide-react';
import { AttendanceStatus, PortalDatabase, Student } from '../types';
import {
  ATTENDANCE_STATUS_META,
  formatLongDateId,
  formatShortDateId,
  getStudentDummyPhoto,
  isGoogleDrivePhotoUrl,
} from '../utils/portalHelpers';
import { PersonAvatar } from './PersonAvatar';

interface StudentDetailModalProps {
  student: Student | null;
  db: PortalDatabase;
  onClose: () => void;
  onEdit?: (student: Student) => void;
}

type DetailTab = 'attendance' | 'grades';

export function StudentDetailModal({
  student,
  db,
  onClose,
  onEdit,
}: StudentDetailModalProps) {
  const [activeTab, setActiveTab] = useState<DetailTab>('attendance');

  useEffect(() => {
    if (!student) return;
    setActiveTab('attendance');
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [student, onClose]);

  const stats = useMemo(() => {
    if (!student) {
      return {
        totalMeetings: 0,
        H: 0,
        S: 0,
        I: 0,
        A: 0,
        D: 0,
        attPct: 0,
        meetingHistory: [] as {
          id: string;
          date: string;
          jamKe: string;
          subject: string;
          topic: string;
          status: AttendanceStatus;
          note: string;
        }[],
        studentGrades: [] as {
          subject: string;
          examType: string;
          score: number;
          note: string;
        }[],
        avgScore: null as number | null,
      };
    }

    const activeJournals = db.journals
      .filter(
        (j) =>
          j.className === student.className &&
          j.academicYear === db.session.academicYear &&
          j.semester === db.session.semester
      )
      .sort((a, b) => b.date.localeCompare(a.date));

    let H = 0;
    let S = 0;
    let I = 0;
    let A = 0;
    let D = 0;

    const meetingHistory = activeJournals.map((j) => {
      const rec = j.attendance[student.id];
      const status: AttendanceStatus = rec?.status || 'H';
      if (status === 'H') H += 1;
      else if (status === 'S') S += 1;
      else if (status === 'I') I += 1;
      else if (status === 'A') A += 1;
      else if (status === 'D') D += 1;

      return {
        id: j.id,
        date: j.date,
        jamKe: j.jamKe,
        subject: j.subject,
        topic: j.topic,
        status,
        note: rec?.note || '',
      };
    });

    const totalMeetings = activeJournals.length;
    const attPct = totalMeetings > 0 ? (H / totalMeetings) * 100 : 0;

    const studentGrades = db.grades
      .filter(
        (g) =>
          g.studentId === student.id &&
          g.academicYear === db.session.academicYear &&
          g.semester === db.session.semester &&
          g.score !== null
      )
      .map((g) => ({
        subject: g.subject,
        examType: g.examType,
        score: g.score as number,
        note: g.note || '',
      }))
      .sort(
        (a, b) =>
          a.subject.localeCompare(b.subject) || a.examType.localeCompare(b.examType)
      );

    const avgScore =
      studentGrades.length > 0
        ? studentGrades.reduce((sum, g) => sum + g.score, 0) / studentGrades.length
        : null;

    return {
      totalMeetings,
      H,
      S,
      I,
      A,
      D,
      attPct,
      meetingHistory,
      studentGrades,
      avgScore,
    };
  }, [db.grades, db.journals, db.session.academicYear, db.session.semester, student]);

  if (!student) return null;

  const ttlDisplay =
    student.birthPlace && student.birthDate
      ? `${student.birthPlace}, ${formatLongDateId(student.birthDate)}`
      : student.birthPlace || (student.birthDate ? formatLongDateId(student.birthDate) : '-');

  return (
    <div
      className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-[2px] flex items-center justify-center p-4"
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      aria-labelledby="student-detail-title"
    >
      <div
        className="bg-white border border-slate-200 rounded-2xl max-w-3xl w-full shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Top Dark Navy Header */}
        <div className="bg-slate-900 px-6 py-5 text-white flex items-start justify-between gap-4">
          <div className="flex items-center gap-4 min-w-0">
            <PersonAvatar
              src={student.photo}
              fallbackSrc={getStudentDummyPhoto(student.gender)}
              alt={student.name}
              className="w-16 h-16 rounded-2xl object-cover object-top border-2 border-white/20 bg-white/10 shrink-0 shadow-md"
            />

            <div className="min-w-0">
              <h3
                id="student-detail-title"
                className="text-lg sm:text-xl font-bold text-white leading-tight truncate"
              >
                {student.name}
              </h3>

              <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-slate-200 mt-1 font-mono-tabular">
                <span className="font-sans font-semibold text-indigo-200 underline underline-offset-4 decoration-indigo-400/60">
                  Kelas {student.className}
                </span>
                <span aria-hidden="true" className="text-slate-400">
                  ·
                </span>
                <span>Absen {student.absentNo || '-'}</span>
                <span aria-hidden="true" className="text-slate-400">
                  ·
                </span>
                <span>NIS {student.nis || '-'}</span>
                <span aria-hidden="true" className="text-slate-400">
                  ·
                </span>
                <span>NISN {student.nisn || '-'}</span>
              </div>

              <div className="flex flex-wrap items-center gap-x-1.5 gap-y-0.5 text-xs text-slate-300 mt-1.5">
                <User className="w-3.5 h-3.5 text-indigo-300 shrink-0" />
                <span>{student.status || 'Aktif'}</span>
                <span aria-hidden="true">·</span>
                <span>{student.gender === 'P' ? 'Perempuan' : 'Laki-laki'}</span>
                <span aria-hidden="true">·</span>
                <span className="tabular-nums">Masuk {student.entryYear || '-'}</span>
                {isGoogleDrivePhotoUrl(student.photo) ? (
                  <>
                    <span aria-hidden="true">·</span>
                    <span className="text-indigo-200">Foto Google Drive</span>
                  </>
                ) : !student.photo ? (
                  <>
                    <span aria-hidden="true">·</span>
                    <span className="text-slate-400">Foto Dummy</span>
                  </>
                ) : null}
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            {onEdit && (
              <button
                type="button"
                onClick={() => {
                  onClose();
                  onEdit(student);
                }}
                className="px-3.5 py-2 rounded-xl bg-white/10 hover:bg-white/20 border border-white/15 text-white text-xs font-semibold transition-colors inline-flex items-center gap-1.5 whitespace-nowrap"
              >
                <Pencil className="w-3.5 h-3.5" />
                <span>Edit Biodata</span>
              </button>
            )}
            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-xl text-slate-300 hover:text-white hover:bg-white/10 transition-colors"
              aria-label="Tutup detail siswa"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Biodata Grid */}
        <div className="bg-white">
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 px-6 py-4 border-b border-slate-100 text-xs sm:text-sm">
            <div className="flex items-start gap-2.5 min-w-0">
              <Calendar className="w-4 h-4 text-indigo-500 mt-0.5 shrink-0" />
              <div className="min-w-0">
                <p className="text-[11px] font-medium text-slate-500">Tempat, Tanggal Lahir</p>
                <p className="font-semibold text-slate-900 mt-0.5">{ttlDisplay}</p>
              </div>
            </div>

            <div className="flex items-start gap-2.5 min-w-0">
              <Phone className="w-4 h-4 text-indigo-500 mt-0.5 shrink-0" />
              <div className="min-w-0">
                <p className="text-[11px] font-medium text-slate-500">No. HP & Email Siswa</p>
                <p className="font-semibold font-mono-tabular text-slate-900 mt-0.5">
                  {student.phone || '-'}
                </p>
                <p className="text-xs text-slate-500 break-all mt-0.5">{student.email || '-'}</p>
              </div>
            </div>

            <div className="flex items-start gap-2.5 min-w-0">
              <Users className="w-4 h-4 text-indigo-500 mt-0.5 shrink-0" />
              <div className="min-w-0">
                <p className="text-[11px] font-medium text-slate-500">Wali Kelas & Guru BK</p>
                <p className="font-semibold text-slate-900 mt-0.5">
                  WK: {student.homeroomTeacher || '-'}
                </p>
                <p className="text-xs text-slate-500 mt-0.5">
                  BK: {student.counselorTeacher || '-'}
                </p>
              </div>
            </div>
          </div>

          <div className="px-6 py-3.5 border-b border-slate-200 flex items-start gap-2.5 text-xs sm:text-sm">
            <MapPin className="w-4 h-4 text-indigo-500 mt-0.5 shrink-0" />
            <div className="min-w-0">
              <p className="text-[11px] font-medium text-slate-500">Alamat Tempat Tinggal</p>
              <p className="font-semibold text-slate-900 mt-0.5 leading-relaxed">
                {student.address || '-'}
              </p>
            </div>
          </div>

          {/* 4-Column Summary Strip */}
          <div className="grid grid-cols-2 sm:grid-cols-4 border-b border-slate-200 divide-y sm:divide-y-0 sm:divide-x divide-slate-200 tabular-nums">
            <div className="px-5 py-3.5">
              <p className="text-[11px] font-medium text-slate-500">Pertemuan Kelas</p>
              <p className="text-base sm:text-lg font-bold text-slate-900 mt-0.5">
                {stats.totalMeetings} Sesi
              </p>
            </div>
            <div className="px-5 py-3.5">
              <p className="text-[11px] font-medium text-slate-500">Persentase Kehadiran</p>
              <p className="text-base sm:text-lg font-bold text-emerald-700 mt-0.5">
                {stats.totalMeetings > 0 ? `${stats.attPct.toFixed(0)}%` : '0%'}
              </p>
            </div>
            <div className="px-5 py-3.5">
              <p className="text-[11px] font-medium text-slate-500">Rincian Presensi</p>
              <p className="text-sm font-bold font-mono-tabular text-slate-900 mt-1">
                H:{stats.H} S:{stats.S} I:{stats.I} A:{stats.A} D:{stats.D}
              </p>
            </div>
            <div className="px-5 py-3.5">
              <p className="text-[11px] font-medium text-slate-500">Rata-Rata Nilai (KKM 75)</p>
              <p className="text-base sm:text-lg font-bold text-indigo-700 mt-0.5 font-mono-tabular">
                {stats.avgScore !== null ? stats.avgScore.toFixed(1) : '—'}
              </p>
            </div>
          </div>

          {/* Tab Navigation */}
          <div className="px-6 border-b border-slate-200 flex items-center gap-6 bg-white">
            <button
              type="button"
              onClick={() => setActiveTab('attendance')}
              className={`py-3 text-xs sm:text-sm font-semibold inline-flex items-center gap-1.5 border-b-2 transition-colors whitespace-nowrap ${
                activeTab === 'attendance'
                  ? 'border-indigo-600 text-indigo-600'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              <ClipboardCheck className="w-4 h-4" />
              <span>Aktivitas & Kehadiran di Kelas ({stats.meetingHistory.length})</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab('grades')}
              className={`py-3 text-xs sm:text-sm font-semibold inline-flex items-center gap-1.5 border-b-2 transition-colors whitespace-nowrap ${
                activeTab === 'grades'
                  ? 'border-indigo-600 text-indigo-600'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              }`}
            >
              <FileText className="w-4 h-4" />
              <span>Riwayat Nilai & Evaluasi ({stats.studentGrades.length})</span>
            </button>
          </div>
        </div>

        {/* Scrollable Tab Content */}
        <div className="flex-1 overflow-y-auto table-scroll bg-slate-50/40 min-h-[190px] max-h-[280px]">
          {activeTab === 'attendance' ? (
            stats.meetingHistory.length === 0 ? (
              <div className="h-44 flex items-center justify-center px-6 text-center">
                <p className="text-xs sm:text-sm text-slate-500">
                  Belum ada catatan pertemuan atau aktivitas KBM di kelas {student.className}.
                </p>
              </div>
            ) : (
              <table className="w-full text-left border-collapse text-xs sm:text-sm">
                <thead className="sticky top-0 bg-slate-100/95 border-b border-slate-200 z-10">
                  <tr>
                    <th className="px-5 py-2.5 font-semibold text-slate-700 whitespace-nowrap">Tanggal</th>
                    <th className="px-3 py-2.5 font-semibold text-slate-700 whitespace-nowrap">Jam</th>
                    <th className="px-3 py-2.5 font-semibold text-slate-700">Mata Pelajaran</th>
                    <th className="px-3 py-2.5 font-semibold text-slate-700">Materi Pokok</th>
                    <th className="px-5 py-2.5 font-semibold text-slate-700 text-right whitespace-nowrap">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {stats.meetingHistory.map((m) => {
                    const meta = ATTENDANCE_STATUS_META[m.status];
                    return (
                      <tr key={m.id} className="hover:bg-slate-50">
                        <td className="px-5 py-2 font-mono-tabular text-slate-700 whitespace-nowrap">
                          {formatShortDateId(m.date)}
                        </td>
                        <td className="px-3 py-2 font-mono-tabular text-slate-500 whitespace-nowrap">
                          {m.jamKe}
                        </td>
                        <td className="px-3 py-2 font-medium text-slate-900">{m.subject}</td>
                        <td className="px-3 py-2 text-slate-600">
                          {m.topic}
                          {m.note && (
                            <span className="block text-[11px] text-slate-400 mt-0.5">
                              Catatan: {m.note}
                            </span>
                          )}
                        </td>
                        <td className={`px-5 py-2 text-right font-semibold whitespace-nowrap ${meta.text}`}>
                          {meta.label} ({m.status})
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )
          ) : stats.studentGrades.length === 0 ? (
            <div className="h-44 flex items-center justify-center px-6 text-center">
              <p className="text-xs sm:text-sm text-slate-500">
                Belum ada riwayat nilai atau evaluasi untuk {student.name} pada periode aktif.
              </p>
            </div>
          ) : (
            <table className="w-full text-left border-collapse text-xs sm:text-sm">
              <thead className="sticky top-0 bg-slate-100/95 border-b border-slate-200 z-10">
                <tr>
                  <th className="px-5 py-2.5 font-semibold text-slate-700">Mata Pelajaran</th>
                  <th className="px-3 py-2.5 font-semibold text-slate-700">Jenis Ujian</th>
                  <th className="px-3 py-2.5 font-semibold text-slate-700 text-right">Nilai</th>
                  <th className="px-3 py-2.5 font-semibold text-slate-700">Keterangan</th>
                  <th className="px-5 py-2.5 font-semibold text-slate-700">Catatan</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {stats.studentGrades.map((g, idx) => (
                  <tr key={`${g.subject}-${g.examType}-${idx}`} className="hover:bg-slate-50">
                    <td className="px-5 py-2 font-medium text-slate-900">{g.subject}</td>
                    <td className="px-3 py-2 text-slate-700">{g.examType}</td>
                    <td className="px-3 py-2 text-right font-mono-tabular font-bold text-slate-900">
                      {g.score}
                    </td>
                    <td
                      className={`px-3 py-2 font-semibold whitespace-nowrap ${
                        g.score >= 75 ? 'text-emerald-700' : 'text-rose-700'
                      }`}
                    >
                      {g.score >= 75 ? 'Tuntas' : 'Remedial'}
                    </td>
                    <td className="px-5 py-2 text-slate-500">{g.note || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}
