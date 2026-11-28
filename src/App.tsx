import { useCallback, useEffect, useRef, useState } from 'react';
import { Loader2 } from 'lucide-react';
import { AccountsView } from './components/AccountsView';
import { DashboardView } from './components/DashboardView';
import { ConfirmDialogState, ConfirmModal, ToastContainer, ToastItem } from './components/FeedbackUI';
import { GradesView } from './components/GradesView';
import { JournalView } from './components/JournalView';
import { Layout } from './components/Layout';
import { LoginView } from './components/LoginView';
import { ReportsView } from './components/ReportsView';
import { ScheduleView } from './components/ScheduleView';
import { SettingsView } from './components/SettingsView';
import { StudentsView } from './components/StudentsView';
import { TeachersView } from './components/TeachersView';
import { usePortalStore } from './store/usePortalStore';
import { JournalEntry, MenuId, Teacher, TeachingSchedule } from './types';
import { getTodayISO } from './utils/portalHelpers';

export default function App() {
  const {
    db,
    authLoading,
    signingIn,
    dataLoading,
    accessDeniedMessage,
    pendingUserEmail,
    fatalError,
    loginWithGoogle,
    loginWithCredentials,
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
    syncTeachersFromAccounts,
    upsertAccount,
    importAccountsBatch,
    deleteAccount,
    deleteAccountsBatch,
    updateSettings,
    setActivePeriod,
    restoreDatabase,
    resetAllData,
  } = usePortalStore();

  if (fatalError) {
    throw fatalError;
  }

  useEffect(() => {
    document.title = db.settings.identity?.webTitle || 'Portal E-Guru v10.2';
  }, [db.settings.identity?.webTitle]);

  const [activeMenu, setActiveMenu] = useState<MenuId>('dashboard');
  const [draftJournal, setDraftJournal] = useState<{
    mode: 'edit' | 'duplicate';
    entry: JournalEntry;
  } | null>(null);
  const [draftTeacherPrint, setDraftTeacherPrint] = useState<{
    teacherId: string;
    nonce: number;
  } | null>(null);

  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [confirmState, setConfirmState] = useState<ConfirmDialogState | null>(null);

  const showToast = useCallback(
    (type: 'success' | 'error' | 'info', message: string) => {
      const id = `${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      setToasts((prev) => [...prev, { id, type, message }]);
      setTimeout(() => {
        setToasts((prev) => prev.filter((t) => t.id !== id));
      }, 3800);
    },
    []
  );

  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  const showConfirm = useCallback((cfg: Omit<ConfirmDialogState, 'open'>) => {
    setConfirmState({ ...cfg, open: true });
  }, []);

  const pendingLoginToastRef = useRef(false);

  useEffect(() => {
    if (pendingLoginToastRef.current && db.session.isLoggedIn && !dataLoading) {
      pendingLoginToastRef.current = false;
      const roleLabel = db.session.userRole ? ` (${db.session.userRole})` : '';
      showToast(
        'success',
        `Berhasil masuk sebagai ${db.session.userName || db.session.email}${roleLabel}.`
      );
    }
  }, [
    db.session.isLoggedIn,
    db.session.userName,
    db.session.email,
    db.session.userRole,
    dataLoading,
    showToast,
  ]);

  useEffect(() => {
    if (pendingLoginToastRef.current && accessDeniedMessage) {
      pendingLoginToastRef.current = false;
      showToast('error', accessDeniedMessage);
    }
  }, [accessDeniedMessage, showToast]);

  const handleEditJournalFromDashboard = useCallback(
    (entry: JournalEntry) => {
      setDraftJournal({ mode: 'edit', entry });
      setActiveMenu('journal');
      showToast('info', `Mode edit dibuka untuk jurnal ${entry.className} (${entry.subject}).`);
    },
    [showToast]
  );

  const handleDuplicateJournalFromDashboard = useCallback(
    (entry: JournalEntry) => {
      setDraftJournal({ mode: 'duplicate', entry });
      setActiveMenu('journal');
      showToast('info', `Salinan jurnal ${entry.className} (${entry.subject}) disiapkan.`);
    },
    [showToast]
  );

  const handleCreateJournalFromSchedule = useCallback(
    (item: TeachingSchedule) => {
      const attendance: JournalEntry['attendance'] = {};
      db.students
        .filter((s) => s.className === item.className)
        .forEach((s) => {
          attendance[s.id] = { status: 'H', note: '' };
        });

      const entry: JournalEntry = {
        id: `draft-${Date.now()}`,
        date: getTodayISO(),
        jamKe: item.jamKe,
        className: item.className,
        subject: item.subject,
        topic: '',
        activity: '',
        academicYear: db.session.academicYear,
        semester: db.session.semester,
        attendance,
        createdAt: new Date().toISOString(),
      };
      setDraftJournal({ mode: 'duplicate', entry });
      setActiveMenu('journal');
      showToast('info', `Form jurnal disiapkan untuk ${item.className} (${item.subject}).`);
    },
    [db.students, db.session.academicYear, db.session.semester, showToast]
  );

  const clearDraftJournal = useCallback(() => {
    setDraftJournal(null);
  }, []);

  const handleOpenPrintForTeacher = useCallback(
    (teacher: Teacher) => {
      setDraftTeacherPrint({ teacherId: teacher.id, nonce: Date.now() });
      setActiveMenu('reports');
      showToast('info', `Pusat Laporan dibuka & disiapkan untuk ${teacher.name}.`);
    },
    [showToast]
  );

  const clearDraftTeacherPrint = useCallback(() => {
    setDraftTeacherPrint(null);
  }, []);

  return (
    <>
      {!db.session.isLoggedIn ? (
        <LoginView
          settings={db.settings}
          authLoading={authLoading || dataLoading}
          signingIn={signingIn}
          accessDeniedMessage={accessDeniedMessage}
          pendingUserEmail={pendingUserEmail}
          onGoogleLogin={async (academicYear, semester, applyPeriodOverride) => {
            pendingLoginToastRef.current = true;
            await loginWithGoogle(academicYear, semester, applyPeriodOverride);
            setActiveMenu('dashboard');
          }}
          onCredentialsLogin={async (username, password, academicYear, semester, applyPeriodOverride) => {
            pendingLoginToastRef.current = true;
            await loginWithCredentials(username, password, academicYear, semester, applyPeriodOverride);
            setActiveMenu('dashboard');
          }}
          onSwitchAccount={async () => {
            await logout();
            showToast('info', 'Silakan pilih akun Google lain.');
          }}
        />
      ) : dataLoading ? (
        <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center gap-3 text-slate-600">
          <Loader2 className="w-7 h-7 animate-spin text-indigo-600" />
          <p className="text-sm font-medium">Menyinkronkan data dari Cloud Firestore...</p>
        </div>
      ) : (
        <Layout
          activeMenu={activeMenu}
          onSelectMenu={setActiveMenu}
          settings={db.settings}
          session={db.session}
          onChangePeriod={(academicYear, semester) => {
            setActivePeriod(academicYear, semester);
            showToast('success', `Periode aktif diubah ke TA ${academicYear} (${semester}).`);
          }}
          onLogout={() => {
            void logout().then(() => {
              showToast('info', 'Anda telah keluar dari sesi Portal E-Guru.');
            });
          }}
        >
          {activeMenu === 'dashboard' && (
            <DashboardView
              db={db}
              onNavigate={setActiveMenu}
              onEditJournal={handleEditJournalFromDashboard}
              onDuplicateJournal={handleDuplicateJournalFromDashboard}
              onDeleteJournals={deleteJournals}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'journal' && (
            <JournalView
              db={db}
              draftJournal={draftJournal}
              onClearDraft={clearDraftJournal}
              onUpsertJournal={upsertJournal}
              onImportJournalsBatch={importJournalsBatch}
              onDeleteJournals={deleteJournals}
              onBatchUpdateJournals={batchUpdateJournals}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'schedule' && (
            <ScheduleView
              db={db}
              onUpsertSchedule={upsertSchedule}
              onImportSchedulesBatch={importSchedulesBatch}
              onDeleteSchedule={deleteSchedule}
              onDeleteSchedules={deleteSchedules}
              onCreateJournalFromSchedule={handleCreateJournalFromSchedule}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'grades' && (
            <GradesView
              db={db}
              onUpsertGradesBatch={upsertGradesBatch}
              onDeleteGradesForStudents={deleteGradesForStudents}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'students' && (
            <StudentsView
              db={db}
              onUpsertStudent={upsertStudent}
              onImportStudentsBatch={importStudentsBatch}
              onDeleteStudent={deleteStudent}
              onDeleteStudentsBatch={deleteStudentsBatch}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'teachers' && (
            <TeachersView
              db={db}
              onUpsertTeacher={upsertTeacher}
              onImportTeachersBatch={importTeachersBatch}
              onDeleteTeacher={deleteTeacher}
              onDeleteTeachersBatch={deleteTeachersBatch}
              onSyncFromAccounts={syncTeachersFromAccounts}
              onOpenPrintForTeacher={handleOpenPrintForTeacher}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'reports' && (
            <ReportsView
              db={db}
              draftTeacherPrint={draftTeacherPrint}
              onClearDraftTeacherPrint={clearDraftTeacherPrint}
              onUpdateSettings={updateSettings}
              onToast={showToast}
            />
          )}

          {activeMenu === 'accounts' && (
            <AccountsView
              db={db}
              onUpsertAccount={upsertAccount}
              onImportAccountsBatch={importAccountsBatch}
              onDeleteAccount={deleteAccount}
              onDeleteAccountsBatch={deleteAccountsBatch}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'settings' && (
            <SettingsView
              db={db}
              onUpdateSettings={updateSettings}
              onRestoreDatabase={restoreDatabase}
              onResetAllData={resetAllData}
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}
        </Layout>
      )}

      <ToastContainer toasts={toasts} onDismiss={dismissToast} />
      <ConfirmModal state={confirmState} onClose={() => setConfirmState(null)} />
    </>
  );
}
