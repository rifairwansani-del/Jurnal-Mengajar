import { useCallback, useEffect, useState } from 'react';
import { createDemoDatabase, createEmptyDatabase } from '../data/seedData';
import {
  GradeEntry,
  JournalEntry,
  PortalDatabase,
  ReportPrintPreferences,
  ReportSectionId,
  SemesterType,
  SignaturePositionSettings,
  Student,
  SystemSettings,
  Teacher,
  TeachingSchedule,
  UserAccount,
} from '../types';

const STORAGE_KEY = 'portal_eguru_v10_2_db';

const VALID_REPORT_SECTION_IDS: ReportSectionId[] = [
  'journal',
  'attendance',
  'attendance_monthly',
  'attendance_monitoring',
  'grades',
  'grades_remedial',
  'grades_leger',
  'students_roster',
  'teacher_activity',
];

function normalizeReportPrefs(
  raw: Partial<ReportPrintPreferences> | undefined,
  fallback: ReportPrintPreferences
): ReportPrintPreferences {
  const merged = { ...fallback, ...(raw || {}) };
  let selectedReportTypes: ReportSectionId[];
  if (Array.isArray(raw?.selectedReportTypes)) {
    selectedReportTypes = raw.selectedReportTypes.filter((id): id is ReportSectionId =>
      VALID_REPORT_SECTION_IDS.includes(id as ReportSectionId)
    );
  } else if (raw?.reportType === 'package') {
    selectedReportTypes = ['journal', 'attendance', 'grades'];
  } else if (
    raw?.reportType &&
    VALID_REPORT_SECTION_IDS.includes(raw.reportType as ReportSectionId)
  ) {
    selectedReportTypes = [raw.reportType as ReportSectionId];
  } else {
    selectedReportTypes = fallback.selectedReportTypes;
  }

  return {
    ...merged,
    selectedReportTypes,
    showSignatureImage: raw?.showSignatureImage ?? fallback.showSignatureImage,
  };
}

function normalizeSignaturePosition(
  raw: Partial<SignaturePositionSettings> | undefined,
  fallback: SignaturePositionSettings
): SignaturePositionSettings {
  const legacySig = raw?.signatureSize;
  return {
    ...fallback,
    ...(raw || {}),
    principalSignatureSize:
      raw?.principalSignatureSize ?? legacySig ?? fallback.principalSignatureSize,
    teacherSignatureSize:
      raw?.teacherSignatureSize ?? legacySig ?? fallback.teacherSignatureSize,
    stampSize: raw?.stampSize ?? fallback.stampSize,
  };
}

function normalizeStudent(raw: Partial<Student>, fallbackIdx = 0): Student {
  return {
    id: raw.id || `st-${Date.now()}-${fallbackIdx}`,
    nis: raw.nis || '',
    nisn: raw.nisn || '',
    name: raw.name || '',
    className: raw.className || '',
    absentNo: raw.absentNo || '',
    birthPlace: raw.birthPlace || '',
    birthDate: raw.birthDate || '',
    email: raw.email || '',
    photo: raw.photo || '',
    address: raw.address || '',
    phone: raw.phone || '',
    gender: raw.gender === 'P' ? 'P' : 'L',
    status: raw.status || 'Aktif',
    homeroomTeacher: raw.homeroomTeacher || '',
    counselorTeacher: raw.counselorTeacher || '',
    entryYear: raw.entryYear || '',
  };
}

function normalizeAccount(
  raw: Partial<UserAccount>,
  fallbackStamp: string,
  idx = 0
): UserAccount {
  return {
    id: raw.id || `acc-${Date.now()}-${idx}`,
    name: raw.name || '',
    email: raw.email || '',
    password: raw.password || 'guru123',
    role: raw.role || 'Guru Mapel',
    nip: raw.nip || '',
    status: raw.status || 'Aktif',
    lastLogin: raw.lastLogin || '-',
    signatureImage: raw.signatureImage || '',
    schoolStampImage: raw.schoolStampImage || fallbackStamp,
  };
}

function loadInitialDatabase(): PortalDatabase {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return createEmptyDatabase();
    }
    const parsed = JSON.parse(raw) as Partial<PortalDatabase>;
    const fallback = createEmptyDatabase();

    const mergedSettings: SystemSettings = parsed.settings
      ? {
          ...fallback.settings,
          ...parsed.settings,
          themeId: parsed.settings.themeId || fallback.settings.themeId,
          identity: {
            ...fallback.settings.identity,
            ...(parsed.settings.identity || {}),
          },
          signaturePosition: normalizeSignaturePosition(
            parsed.settings.signaturePosition,
            fallback.settings.signaturePosition
          ),
          reportPrintPreferences: normalizeReportPrefs(
            parsed.settings.reportPrintPreferences,
            fallback.settings.reportPrintPreferences
          ),
          kopSurat: {
            ...fallback.settings.kopSurat,
            ...(parsed.settings.kopSurat || {}),
          },
          masterClasses: Array.isArray(parsed.settings.masterClasses)
            ? parsed.settings.masterClasses
            : fallback.settings.masterClasses,
          masterSubjects: Array.isArray(parsed.settings.masterSubjects)
            ? parsed.settings.masterSubjects
            : fallback.settings.masterSubjects,
          masterExamTypes: Array.isArray(parsed.settings.masterExamTypes)
            ? parsed.settings.masterExamTypes
            : fallback.settings.masterExamTypes,
          masterAcademicYears: Array.isArray(parsed.settings.masterAcademicYears)
            ? parsed.settings.masterAcademicYears
            : fallback.settings.masterAcademicYears,
        }
      : fallback.settings;

    return {
      students: Array.isArray(parsed.students)
        ? parsed.students.map((s, idx) => normalizeStudent(s, idx))
        : fallback.students,
      teachers: Array.isArray(parsed.teachers) ? parsed.teachers : fallback.teachers,
      schedules: Array.isArray(parsed.schedules) ? parsed.schedules : fallback.schedules,
      accounts:
        Array.isArray(parsed.accounts) && parsed.accounts.length > 0
          ? parsed.accounts.map((a, idx) =>
              normalizeAccount(a, mergedSettings.defaultSchoolStampImage, idx)
            )
          : fallback.accounts,
      journals: Array.isArray(parsed.journals) ? parsed.journals : fallback.journals,
      grades: Array.isArray(parsed.grades) ? parsed.grades : fallback.grades,
      settings: mergedSettings,
      session: parsed.session
        ? { ...fallback.session, ...parsed.session }
        : fallback.session,
    };
  } catch {
    return createEmptyDatabase();
  }
}

export function usePortalStore() {
  const [db, setDb] = useState<PortalDatabase>(loadInitialDatabase);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(db));
    } catch {
      // Ignore storage quota errors gracefully
    }
  }, [db]);

  const loginWithCredentials = useCallback(
    (email: string, password: string, academicYear: string, semester: SemesterType) => {
      const normalizedEmail = email.trim().toLowerCase();

      const matchedAccount = db.accounts.find(
        (acc) => acc.email.trim().toLowerCase() === normalizedEmail
      );

      if (matchedAccount) {
        if (matchedAccount.status !== 'Aktif') {
          return { ok: false, message: 'Akun ini sedang dinonaktifkan oleh Admin.' };
        }
        if (matchedAccount.password !== password) {
          return { ok: false, message: 'Password tidak sesuai dengan akun terdaftar.' };
        }
        const nowStr = new Date().toISOString().slice(0, 16).replace('T', ' ');
        setDb((prev) => ({
          ...prev,
          accounts: prev.accounts.map((a) =>
            a.id === matchedAccount.id ? { ...a, lastLogin: nowStr } : a
          ),
          settings: {
            ...prev.settings,
            academicYear,
            semester,
          },
          session: {
            isLoggedIn: true,
            accountId: matchedAccount.id,
            email: matchedAccount.email,
            userName: matchedAccount.name,
            userRole: matchedAccount.role,
            academicYear,
            semester,
          },
        }));
        return { ok: true, message: `Selamat datang, ${matchedAccount.name}.` };
      }

      const targetEmail = db.settings.email.trim().toLowerCase();
      if (normalizedEmail !== targetEmail || password !== db.settings.password) {
        return { ok: false, message: 'Email atau password tidak sesuai dengan akun lokal.' };
      }

      setDb((prev) => ({
        ...prev,
        settings: {
          ...prev.settings,
          academicYear,
          semester,
        },
        session: {
          isLoggedIn: true,
          email: email.trim(),
          userName: prev.settings.teacherName,
          userRole: 'Admin',
          academicYear,
          semester,
        },
      }));
      return { ok: true, message: 'Berhasil masuk ke Portal E-Guru.' };
    },
    [db.accounts, db.settings.email, db.settings.password]
  );

  const loginWithDemoSeed = useCallback((academicYear?: string, semester?: SemesterType) => {
    const demo = createDemoDatabase();
    if (academicYear) {
      demo.settings.academicYear = academicYear;
      demo.session.academicYear = academicYear;
    }
    if (semester) {
      demo.settings.semester = semester;
      demo.session.semester = semester;
    }
    setDb(demo);
  }, []);

  const loginEmpty = useCallback((email: string, academicYear: string, semester: SemesterType) => {
    const empty = createEmptyDatabase();
    if (email.trim()) {
      empty.settings.email = email.trim();
      empty.session.email = email.trim();
      if (empty.accounts[0]) {
        empty.accounts[0].email = email.trim();
      }
    }
    empty.settings.academicYear = academicYear;
    empty.settings.semester = semester;
    empty.session.academicYear = academicYear;
    empty.session.semester = semester;
    empty.session.isLoggedIn = true;
    setDb(empty);
  }, []);

  const logout = useCallback(() => {
    setDb((prev) => ({
      ...prev,
      session: {
        ...prev.session,
        isLoggedIn: false,
      },
    }));
  }, []);

  // Journals
  const upsertJournal = useCallback((entry: JournalEntry) => {
    setDb((prev) => {
      const exists = prev.journals.some((j) => j.id === entry.id);
      const journals = exists
        ? prev.journals.map((j) => (j.id === entry.id ? entry : j))
        : [entry, ...prev.journals];
      return { ...prev, journals };
    });
  }, []);

  const importJournalsBatch = useCallback((incoming: Omit<JournalEntry, 'id' | 'createdAt'>[]) => {
    let added = 0;
    setDb((prev) => {
      const next = [...prev.journals];
      const newClasses = new Set(prev.settings.masterClasses);
      const newSubjects = new Set(prev.settings.masterSubjects);
      incoming.forEach((item, idx) => {
        if (!item.topic.trim() || !item.className.trim()) return;
        newClasses.add(item.className);
        if (item.subject) newSubjects.add(item.subject);
        next.unshift({
          ...item,
          id: `jrn-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 5)}`,
          createdAt: new Date().toISOString(),
        });
        added += 1;
      });
      return {
        ...prev,
        journals: next,
        settings: {
          ...prev.settings,
          masterClasses: Array.from(newClasses),
          masterSubjects: Array.from(newSubjects),
        },
      };
    });
    return added;
  }, []);

  const deleteJournals = useCallback((ids: string[]) => {
    const idSet = new Set(ids);
    setDb((prev) => ({
      ...prev,
      journals: prev.journals.filter((j) => !idSet.has(j.id)),
    }));
  }, []);

  const batchUpdateJournals = useCallback((updatedJournals: JournalEntry[]) => {
    const map = new Map(updatedJournals.map((j) => [j.id, j]));
    setDb((prev) => ({
      ...prev,
      journals: prev.journals.map((j) => map.get(j.id) ?? j),
    }));
  }, []);

  // Schedules
  const upsertSchedule = useCallback((item: TeachingSchedule) => {
    setDb((prev) => {
      const exists = prev.schedules.some((s) => s.id === item.id);
      const schedules = exists
        ? prev.schedules.map((s) => (s.id === item.id ? item : s))
        : [...prev.schedules, item];
      return { ...prev, schedules };
    });
  }, []);

  const importSchedulesBatch = useCallback((incoming: Omit<TeachingSchedule, 'id'>[]) => {
    let added = 0;
    setDb((prev) => {
      const next = [...prev.schedules];
      incoming.forEach((item, idx) => {
        if (!item.className.trim() || !item.subject.trim()) return;
        next.push({
          ...item,
          id: `sch-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 5)}`,
        });
        added += 1;
      });
      return { ...prev, schedules: next };
    });
    return added;
  }, []);

  const deleteSchedules = useCallback((ids: string[]) => {
    const idSet = new Set(ids);
    setDb((prev) => ({
      ...prev,
      schedules: prev.schedules.filter((s) => !idSet.has(s.id)),
    }));
  }, []);

  const deleteSchedule = useCallback((id: string) => {
    setDb((prev) => ({
      ...prev,
      schedules: prev.schedules.filter((s) => s.id !== id),
    }));
  }, []);

  // Grades
  const upsertGradesBatch = useCallback((entries: GradeEntry[]) => {
    setDb((prev) => {
      const map = new Map(prev.grades.map((g) => [g.id, g]));
      entries.forEach((entry) => {
        map.set(entry.id, entry);
      });
      return {
        ...prev,
        grades: Array.from(map.values()),
      };
    });
  }, []);

  const deleteGradesForStudents = useCallback(
    (studentIds: string[], className: string, subject: string) => {
      const stSet = new Set(studentIds);
      setDb((prev) => ({
        ...prev,
        grades: prev.grades.filter(
          (g) =>
            !(
              stSet.has(g.studentId) &&
              g.className === className &&
              g.subject === subject &&
              g.academicYear === prev.session.academicYear &&
              g.semester === prev.session.semester
            )
        ),
      }));
    },
    []
  );

  // Students
  const upsertStudent = useCallback((student: Student) => {
    setDb((prev) => {
      const exists = prev.students.some((s) => s.id === student.id);
      const students = exists
        ? prev.students.map((s) => (s.id === student.id ? student : s))
        : [...prev.students, student];
      const masterClasses = prev.settings.masterClasses.includes(student.className)
        ? prev.settings.masterClasses
        : [...prev.settings.masterClasses, student.className];
      return {
        ...prev,
        students,
        settings: { ...prev.settings, masterClasses },
      };
    });
  }, []);

  const importStudentsBatch = useCallback((incoming: Omit<Student, 'id'>[]) => {
    let added = 0;
    let updated = 0;
    setDb((prev) => {
      const nextStudents = [...prev.students];
      const newClasses = new Set(prev.settings.masterClasses);

      incoming.forEach((item, index) => {
        const cleanNis = item.nis.trim();
        const cleanName = item.name.trim();
        const cleanClass = item.className.trim();
        if (!cleanName || !cleanClass) return;

        newClasses.add(cleanClass);
        const existingIdx = cleanNis
          ? nextStudents.findIndex((s) => s.nis.toLowerCase() === cleanNis.toLowerCase())
          : -1;

        if (existingIdx >= 0) {
          const prevStudent = nextStudents[existingIdx];
          nextStudents[existingIdx] = {
            ...prevStudent,
            ...item,
            nis: cleanNis,
            name: cleanName,
            className: cleanClass,
            photo: item.photo || prevStudent.photo,
          };
          updated += 1;
        } else {
          nextStudents.push({
            ...item,
            id: `st-${Date.now()}-${index}-${Math.random().toString(36).slice(2, 6)}`,
            nis: cleanNis || `${Date.now().toString().slice(-4)}${index}`,
            name: cleanName,
            className: cleanClass,
          });
          added += 1;
        }
      });

      return {
        ...prev,
        students: nextStudents,
        settings: {
          ...prev.settings,
          masterClasses: Array.from(newClasses),
        },
      };
    });
    return { added, updated };
  }, []);

  const deleteStudent = useCallback((studentId: string) => {
    setDb((prev) => ({
      ...prev,
      students: prev.students.filter((s) => s.id !== studentId),
      grades: prev.grades.filter((g) => g.studentId !== studentId),
    }));
  }, []);

  const deleteStudentsBatch = useCallback((ids: string[]) => {
    const idSet = new Set(ids);
    setDb((prev) => ({
      ...prev,
      students: prev.students.filter((s) => !idSet.has(s.id)),
      grades: prev.grades.filter((g) => !idSet.has(g.studentId)),
    }));
  }, []);

  // Teachers
  const upsertTeacher = useCallback((teacher: Teacher) => {
    setDb((prev) => {
      const exists = prev.teachers.some((t) => t.id === teacher.id);
      const teachers = exists
        ? prev.teachers.map((t) => (t.id === teacher.id ? teacher : t))
        : [...prev.teachers, teacher];
      return { ...prev, teachers };
    });
  }, []);

  const importTeachersBatch = useCallback((incoming: Omit<Teacher, 'id'>[]) => {
    let added = 0;
    let updated = 0;
    setDb((prev) => {
      const next = [...prev.teachers];
      incoming.forEach((item, idx) => {
        if (!item.name.trim()) return;
        const existingIdx = item.nip
          ? next.findIndex((t) => t.nip.trim().toLowerCase() === item.nip.trim().toLowerCase())
          : -1;
        if (existingIdx >= 0) {
          next[existingIdx] = { ...next[existingIdx], ...item, photo: item.photo || next[existingIdx].photo };
          updated += 1;
        } else {
          next.push({
            ...item,
            id: `tch-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 5)}`,
          });
          added += 1;
        }
      });
      return { ...prev, teachers: next };
    });
    return { added, updated };
  }, []);

  const deleteTeacher = useCallback((teacherId: string) => {
    setDb((prev) => ({
      ...prev,
      teachers: prev.teachers.filter((t) => t.id !== teacherId),
    }));
  }, []);

  const deleteTeachersBatch = useCallback((ids: string[]) => {
    const idSet = new Set(ids);
    setDb((prev) => ({
      ...prev,
      teachers: prev.teachers.filter((t) => !idSet.has(t.id)),
    }));
  }, []);

  // User Accounts
  const upsertAccount = useCallback((account: UserAccount) => {
    setDb((prev) => {
      const exists = prev.accounts.some((a) => a.id === account.id);
      const accounts = exists
        ? prev.accounts.map((a) => (a.id === account.id ? account : a))
        : [...prev.accounts, account];
      const isCurrent = prev.session.email.toLowerCase() === account.email.toLowerCase();
      return {
        ...prev,
        accounts,
        session: isCurrent
          ? {
              ...prev.session,
              accountId: account.id,
              userName: account.name,
              userRole: account.role,
            }
          : prev.session,
      };
    });
  }, []);

  const importAccountsBatch = useCallback((incoming: Omit<UserAccount, 'id' | 'lastLogin'>[]) => {
    let added = 0;
    let updated = 0;
    setDb((prev) => {
      const next = [...prev.accounts];
      incoming.forEach((item, idx) => {
        if (!item.name.trim() || !item.email.trim()) return;
        const existingIdx = next.findIndex(
          (a) => a.email.trim().toLowerCase() === item.email.trim().toLowerCase()
        );
        if (existingIdx >= 0) {
          next[existingIdx] = {
            ...next[existingIdx],
            ...item,
            signatureImage: item.signatureImage || next[existingIdx].signatureImage,
            schoolStampImage: item.schoolStampImage || next[existingIdx].schoolStampImage,
          };
          updated += 1;
        } else {
          next.push({
            ...item,
            id: `acc-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 5)}`,
            lastLogin: '-',
            schoolStampImage: item.schoolStampImage || prev.settings.defaultSchoolStampImage,
          });
          added += 1;
        }
      });
      return { ...prev, accounts: next };
    });
    return { added, updated };
  }, []);

  const deleteAccount = useCallback((accountId: string) => {
    setDb((prev) => ({
      ...prev,
      accounts: prev.accounts.filter((a) => a.id !== accountId),
    }));
  }, []);

  const deleteAccountsBatch = useCallback((ids: string[]) => {
    const idSet = new Set(ids);
    setDb((prev) => ({
      ...prev,
      accounts: prev.accounts.filter(
        (a) => !idSet.has(a.id) || a.email.toLowerCase() === prev.session.email.toLowerCase()
      ),
    }));
  }, []);

  // Settings
  const updateSettings = useCallback((updater: (prev: SystemSettings) => SystemSettings) => {
    setDb((prev) => {
      const nextSettings = updater(prev.settings);
      return {
        ...prev,
        settings: nextSettings,
        session: {
          ...prev.session,
          academicYear: nextSettings.academicYear,
          semester: nextSettings.semester,
        },
      };
    });
  }, []);

  // Backup & Restore
  const restoreDatabase = useCallback((incoming: PortalDatabase) => {
    const fallback = createEmptyDatabase();
    const mergedSettings: SystemSettings = {
      ...fallback.settings,
      ...(incoming.settings || {}),
      identity: {
        ...fallback.settings.identity,
        ...(incoming.settings?.identity || {}),
      },
      signaturePosition: normalizeSignaturePosition(
        incoming.settings?.signaturePosition,
        fallback.settings.signaturePosition
      ),
      reportPrintPreferences: normalizeReportPrefs(
        incoming.settings?.reportPrintPreferences,
        fallback.settings.reportPrintPreferences
      ),
      kopSurat: {
        ...fallback.settings.kopSurat,
        ...(incoming.settings?.kopSurat || {}),
      },
    };
    setDb({
      ...fallback,
      ...incoming,
      settings: mergedSettings,
      students: Array.isArray(incoming.students)
        ? incoming.students.map((s, idx) => normalizeStudent(s, idx))
        : [],
      teachers: Array.isArray(incoming.teachers) ? incoming.teachers : [],
      schedules: Array.isArray(incoming.schedules) ? incoming.schedules : [],
      accounts:
        Array.isArray(incoming.accounts) && incoming.accounts.length > 0
          ? incoming.accounts.map((a, idx) =>
              normalizeAccount(a, mergedSettings.defaultSchoolStampImage, idx)
            )
          : fallback.accounts,
      session: {
        isLoggedIn: true,
        accountId: mergedSettings.email,
        email: mergedSettings.email,
        userName: mergedSettings.teacherName,
        userRole: 'Admin',
        academicYear: mergedSettings.academicYear,
        semester: mergedSettings.semester,
      },
    });
  }, []);

  const resetAllData = useCallback(() => {
    const empty = createEmptyDatabase();
    localStorage.removeItem(STORAGE_KEY);
    setDb(empty);
  }, []);

  return {
    db,
    loginWithCredentials,
    loginWithDemoSeed,
    loginEmpty,
    logout,
    upsertJournal,
    importJournalsBatch,
    deleteJournals,
    batchUpdateJournals,
    upsertSchedule,
    importSchedulesBatch,
    deleteSchedule,
    deleteSchedules,
    upsertGradesBatch,
    deleteGradesForStudents,
    upsertStudent,
    importStudentsBatch,
    deleteStudent,
    deleteStudentsBatch,
    upsertTeacher,
    importTeachersBatch,
    deleteTeacher,
    deleteTeachersBatch,
    upsertAccount,
    importAccountsBatch,
    deleteAccount,
    deleteAccountsBatch,
    updateSettings,
    restoreDatabase,
    resetAllData,
  };
}
