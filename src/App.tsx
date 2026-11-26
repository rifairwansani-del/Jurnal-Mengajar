import { useCallback, useEffect, useState } from 'react';
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
import { JournalEntry, MenuId, TeachingSchedule } from './types';
import { getTodayISO } from './utils/portalHelpers';

export default function App() {
  const {
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
  } = usePortalStore();

  useEffect(() => {
    document.title = db.settings.identity?.webTitle || 'Portal E-Guru v10.2';
  }, [db.settings.identity?.webTitle]);

  const [activeMenu, setActiveMenu] = useState<MenuId>('dashboard');
  const [draftJournal, setDraftJournal] = useState<{
    mode: 'edit' | 'duplicate';
    entry: JournalEntry;
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

  const handleEditJournalFromDashboard = useCallback((entry: JournalEntry) => {
    setDraftJournal({ mode: 'edit', entry });
    setActiveMenu('journal');
  }, []);

  const handleDuplicateJournalFromDashboard = useCallback((entry: JournalEntry) => {
    setDraftJournal({ mode: 'duplicate', entry });
    setActiveMenu('journal');
  }, []);

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

  return (
    <>
      {!db.session.isLoggedIn ? (
        <LoginView
          settings={db.settings}
          onLogin={(email, password, academicYear, semester) => {
            const res = loginWithCredentials(email, password, academicYear, semester);
            if (res.ok) {
              setActiveMenu('dashboard');
              showToast('success', res.message);
            }
            return res;
          }}
          onDemoSeed={(academicYear, semester) => {
            loginWithDemoSeed(academicYear, semester);
            setActiveMenu('dashboard');
            showToast('success', 'Data demo lengkap berhasil dimuat ke penyimpanan lokal.');
          }}
          onStartEmpty={(email, academicYear, semester) => {
            loginEmpty(email, academicYear, semester);
            setActiveMenu('dashboard');
            showToast('info', 'Ruang kerja baru dimulai dengan data kosong.');
          }}
        />
      ) : (
        <Layout
          activeMenu={activeMenu}
          onSelectMenu={setActiveMenu}
          settings={db.settings}
          session={db.session}
          onLogout={() => {
            logout();
            showToast('info', 'Anda telah keluar dari sesi lokal.');
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
              onToast={showToast}
              onConfirm={showConfirm}
            />
          )}

          {activeMenu === 'reports' && (
            <ReportsView
              db={db}
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
