import React, { useState } from 'react';
import {
  LayoutDashboard,
  ClipboardList,
  CalendarClock,
  GraduationCap,
  Users,
  UserCheck,
  Printer,
  ShieldCheck,
  Settings,
  Menu,
  X,
  LogOut,
  BookOpenCheck,
} from 'lucide-react';
import { MenuId, SessionState, SystemSettings } from '../types';
import { PORTAL_THEMES } from '../utils/themeConfig';

interface LayoutProps {
  activeMenu: MenuId;
  onSelectMenu: (menu: MenuId) => void;
  settings: SystemSettings;
  session: SessionState;
  onLogout: () => void;
  children: React.ReactNode;
}

export const NAV_ITEMS: { id: MenuId; label: string; icon: React.ComponentType<{ className?: string }> }[] = [
  { id: 'dashboard', label: 'Dashboard Analitik', icon: LayoutDashboard },
  { id: 'journal', label: 'Jurnal & Absensi', icon: ClipboardList },
  { id: 'schedule', label: 'Jadwal Guru Mengajar', icon: CalendarClock },
  { id: 'grades', label: 'Input & Matriks Nilai', icon: GraduationCap },
  { id: 'students', label: 'Data Siswa', icon: Users },
  { id: 'teachers', label: 'Data Guru', icon: UserCheck },
  { id: 'reports', label: 'Cetak Laporan Resmi', icon: Printer },
  { id: 'accounts', label: 'Manajemen Akun', icon: ShieldCheck },
  { id: 'settings', label: 'Pengaturan Sistem', icon: Settings },
];

export function Layout({
  activeMenu,
  onSelectMenu,
  settings,
  session,
  onLogout,
  children,
}: LayoutProps) {
  const [drawerOpen, setDrawerOpen] = useState(false);

  const theme = PORTAL_THEMES[settings.themeId] || PORTAL_THEMES['slate-indigo'];
  const currentNav = NAV_ITEMS.find((item) => item.id === activeMenu);

  const handleNavClick = (id: MenuId) => {
    onSelectMenu(id);
    setDrawerOpen(false);
  };

  return (
    <div
      className={`min-h-screen flex ${theme.canvasBg} ${
        theme.isDark ? 'theme-dark text-slate-100' : 'text-slate-900'
      }`}
    >
      {/* Mobile backdrop */}
      {drawerOpen && (
        <div
          className="no-print fixed inset-0 z-40 bg-slate-950/60 backdrop-blur-xs lg:hidden"
          onClick={() => setDrawerOpen(false)}
          aria-hidden="true"
        />
      )}

      {/* Fixed Sidebar (Desktop) & Slide-Drawer (Mobile) */}
      <aside
        className={`no-print fixed inset-y-0 left-0 z-50 w-64 ${theme.sidebarBg} text-slate-100 flex flex-col transition-transform duration-200 ease-out lg:translate-x-0 ${
          drawerOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        <div className={`h-16 px-4 flex items-center justify-between border-b ${theme.sidebarBorder} shrink-0`}>
          <div className="flex items-center gap-2.5 min-w-0">
            {settings.identity?.sidebarLogo ? (
              <img
                src={settings.identity.sidebarLogo}
                alt="Logo Portal"
                className="w-9 h-9 rounded-lg object-contain bg-white/95 p-0.5 shrink-0"
              />
            ) : (
              <div className={`w-9 h-9 rounded-lg ${theme.accentBg} flex items-center justify-center shrink-0`}>
                <BookOpenCheck className="w-4 h-4 text-white" />
              </div>
            )}
            <div className="min-w-0">
              <span className="block text-sm font-bold tracking-tight text-white truncate">
                {settings.identity?.sidebarTitle || 'Portal E-Guru v10.2'}
              </span>
              {settings.identity?.sidebarSubtitle && (
                <span className="block text-[11px] text-slate-400 truncate">
                  {settings.identity.sidebarSubtitle}
                </span>
              )}
            </div>
          </div>
          <button
            type="button"
            onClick={() => setDrawerOpen(false)}
            className="lg:hidden p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-white/10"
            aria-label="Tutup menu navigasi"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto" aria-label="Navigasi utama">
          {NAV_ITEMS.map((item) => {
            const Icon = item.icon;
            const isActive = activeMenu === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => handleNavClick(item.id)}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-lg text-sm font-medium transition-colors whitespace-nowrap ${
                  isActive
                    ? theme.sidebarActiveBg
                    : 'text-slate-300 hover:text-white hover:bg-white/10'
                }`}
              >
                <Icon className={`w-4 h-4 shrink-0 ${isActive ? 'text-white' : 'text-slate-400'}`} />
                <span className="truncate">{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className={`p-4 border-t ${theme.sidebarBorder} space-y-3 shrink-0`}>
          <div className="px-2">
            <p className="text-xs font-semibold text-white truncate">
              {session.userName || settings.teacherName}
            </p>
            <p className="text-xs text-slate-400 truncate mt-0.5">
              {session.userRole || 'Admin'} · {settings.schoolName}
            </p>
          </div>
          <button
            type="button"
            onClick={onLogout}
            className="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-lg text-xs font-medium text-slate-300 bg-white/10 hover:bg-white/20 hover:text-white transition-colors whitespace-nowrap"
          >
            <LogOut className="w-3.5 h-3.5" />
            Keluar Sesi
          </button>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 lg:pl-64 print:pl-0 flex flex-col min-w-0">
        {/* Top Bar */}
        <header className="no-print sticky top-0 z-30 h-14 bg-white/95 backdrop-blur-xs border-b border-slate-200 px-4 sm:px-6 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => setDrawerOpen(true)}
              className="lg:hidden p-2 -ml-2 rounded-lg text-slate-600 hover:text-slate-900 hover:bg-slate-100"
              aria-label="Buka menu navigasi"
            >
              <Menu className="w-5 h-5" />
            </button>
            <h1 className="text-sm sm:text-base font-bold text-slate-900 truncate">
              {currentNav?.label}
            </h1>
          </div>

          <div className="flex items-center gap-4 text-xs text-slate-500 shrink-0">
            <span className="hidden sm:inline whitespace-nowrap">
              TA {session.academicYear} · Semester {session.semester}
            </span>
            <button
              type="button"
              onClick={onLogout}
              className="lg:hidden px-2.5 py-1.5 rounded-lg border border-slate-200 text-slate-700 hover:bg-slate-50 font-medium whitespace-nowrap shrink-0"
            >
              Keluar
            </button>
          </div>
        </header>

        <main className="flex-1 p-4 sm:p-6 lg:p-8 print:p-0 print:max-w-none max-w-[1440px] w-full mx-auto">
          {children}
        </main>
      </div>
    </div>
  );
}
