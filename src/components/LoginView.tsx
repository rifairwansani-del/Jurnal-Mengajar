import React, { useEffect, useState } from 'react';
import {
  AlertCircle,
  ArrowRight,
  Calendar,
  Crown,
  Eye,
  EyeOff,
  KeyRound,
  Layers,
  Loader2,
  Lock,
  LogOut,
  ShieldCheck,
  User,
} from 'lucide-react';
import { SemesterType, SystemSettings } from '../types';

interface LoginViewProps {
  settings: SystemSettings;
  authLoading: boolean;
  signingIn: boolean;
  accessDeniedMessage: string | null;
  pendingUserEmail: string | null;
  onGoogleLogin: (
    academicYear: string,
    semester: SemesterType,
    applyPeriodOverride?: boolean
  ) => Promise<void>;
  onCredentialsLogin: (
    username: string,
    password: string,
    academicYear: string,
    semester: SemesterType,
    applyPeriodOverride?: boolean
  ) => Promise<void>;
  onSwitchAccount: () => Promise<void>;
}

export function LoginView({
  settings,
  authLoading,
  signingIn,
  accessDeniedMessage,
  pendingUserEmail,
  onGoogleLogin,
  onCredentialsLogin,
  onSwitchAccount,
}: LoginViewProps) {
  const [loginMode, setLoginMode] = useState<'credentials' | 'google'>('credentials');
  const [academicYear, setAcademicYear] = useState(settings.academicYear || '2025/2026');
  const [semester, setSemester] = useState<SemesterType | ''>(settings.semester || 'Ganjil');
  const [periodChangedByUser, setPeriodChangedByUser] = useState(false);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  useEffect(() => {
    if (periodChangedByUser) return;
    if (settings.academicYear) setAcademicYear(settings.academicYear);
    if (settings.semester) setSemester(settings.semester);
  }, [settings.academicYear, settings.semester, periodChangedByUser]);

  const handleGoogleClick = async () => {
    setErrorMsg('');
    if (!academicYear || !semester) {
      setErrorMsg('Pilih Tahun Ajaran dan Semester terlebih dahulu.');
      return;
    }
    try {
      await onGoogleLogin(academicYear, semester, periodChangedByUser);
    } catch (err) {
      setErrorMsg(
        err instanceof Error ? err.message : 'Gagal masuk dengan akun Google.'
      );
    }
  };

  const handleCredentialsSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg('');
    if (!academicYear || !semester) {
      setErrorMsg('Pilih Tahun Ajaran dan Semester terlebih dahulu.');
      return;
    }
    const cleanUser = username.trim();
    if (!cleanUser || !password) {
      setErrorMsg('Username dan password wajib diisi.');
      return;
    }
    try {
      await onCredentialsLogin(cleanUser, password, academicYear, semester, periodChangedByUser);
    } catch (err) {
      setErrorMsg(
        err instanceof Error ? err.message : 'Username atau password tidak sesuai.'
      );
    }
  };

  return (
    <div className="relative min-h-screen overflow-hidden bg-gradient-to-b from-blue-500 via-sky-300 to-blue-200 flex items-center justify-center p-4 sm:p-6">
      {/* Soft 3D Study-Room Background Atmosphere */}
      <div className="pointer-events-none absolute inset-0 overflow-hidden" aria-hidden="true">
        <div className="absolute -top-24 right-[-10%] w-[460px] h-[520px] rounded-full bg-white/35 blur-3xl" />
        <div className="absolute top-12 left-[-8%] w-[360px] h-[360px] rounded-full bg-amber-200/35 blur-3xl" />

        <div className="hidden sm:flex absolute top-8 left-[22%] w-16 h-16 rounded-full bg-white/90 border-4 border-blue-700 shadow-lg items-center justify-center">
          <div className="w-1 h-4 bg-blue-900 rounded-full origin-bottom -translate-y-1.5 rotate-12" />
          <div className="w-1 h-3 bg-blue-600 rounded-full origin-bottom -translate-y-1 -rotate-45 absolute" />
          <div className="w-2 h-2 rounded-full bg-blue-800 absolute" />
        </div>

        <div className="hidden md:block absolute top-10 right-[24%] w-32 h-12 rounded-full bg-white/85 shadow-md">
          <div className="absolute -top-5 left-5 w-14 h-14 rounded-full bg-white/90" />
          <div className="absolute -top-7 left-12 w-16 h-16 rounded-full bg-white/95" />
        </div>
        <div className="hidden lg:block absolute top-52 right-[12%] w-24 h-10 rounded-full bg-white/80 shadow-md">
          <div className="absolute -top-4 left-4 w-11 h-11 rounded-full bg-white/90" />
          <div className="absolute -top-5 left-10 w-12 h-12 rounded-full bg-white/90" />
        </div>

        <div className="absolute bottom-0 inset-x-0 h-28 bg-gradient-to-b from-blue-300/80 to-indigo-400/90 border-t border-white/50" />
        <div className="absolute bottom-14 left-1/2 -translate-x-1/2 w-[500px] max-w-[92vw] h-10 rounded-full bg-blue-700/25 blur-md" />

        <div className="hidden lg:block absolute bottom-20 left-10 space-y-1 opacity-90">
          <div className="w-40 h-6 rounded-r-lg bg-blue-600 border-l-8 border-white shadow-md" />
          <div className="w-44 h-7 rounded-r-lg bg-indigo-700 border-l-8 border-sky-100 shadow-md" />
          <div className="w-48 h-7 rounded-r-lg bg-slate-800 border-l-8 border-blue-200 shadow-md" />
        </div>

        <div className="hidden lg:block absolute bottom-20 right-10 space-y-1 opacity-90">
          <div className="w-36 h-6 rounded-l-lg bg-sky-500 border-r-8 border-white shadow-md ml-auto" />
          <div className="w-44 h-7 rounded-l-lg bg-blue-700 border-r-8 border-blue-100 shadow-md ml-auto" />
          <div className="w-48 h-7 rounded-l-lg bg-indigo-900 border-r-8 border-slate-100 shadow-md ml-auto" />
        </div>
      </div>

      {/* Main Claymorphic Login Card */}
      <div className="relative z-10 w-full max-w-[440px]">
        <div className="mx-auto w-[86%] h-4 bg-blue-700/35 rounded-b-full blur-[2px] -mb-2 translate-y-full" />

        <div className="bg-[#faf9f6] rounded-[2.25rem] p-6 sm:p-8 border-2 border-white shadow-[0_26px_60px_-15px_rgba(30,58,138,0.42),inset_0_-8px_18px_rgba(148,163,184,0.16),inset_0_4px_12px_rgba(255,255,255,0.95)]">
          <div className="text-center">
            <div className="inline-flex items-center justify-center relative mb-2">
              {settings.identity?.sidebarLogo ? (
                <span className="w-16 h-16 rounded-2xl bg-white p-1.5 shadow-[0_6px_14px_rgba(37,99,235,0.25),inset_0_2px_4px_rgba(255,255,255,0.9)] flex items-center justify-center border border-slate-100">
                  <img
                    src={settings.identity.sidebarLogo}
                    alt="Logo Portal"
                    className="w-full h-full object-contain"
                  />
                </span>
              ) : (
                <span className="w-14 h-14 rounded-2xl bg-gradient-to-b from-amber-300 to-amber-500 shadow-[0_6px_14px_rgba(245,158,11,0.38),inset_0_2px_4px_rgba(255,255,255,0.75)] flex items-center justify-center">
                  <Crown className="w-7 h-7 text-amber-950" strokeWidth={2.2} />
                </span>
              )}
            </div>

            <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight text-blue-600 drop-shadow-[0_2px_0_rgba(255,255,255,1)]">
              {settings.identity?.loginTitle || 'Selamat Datang'}
            </h1>
            <p className="mt-1 text-xs sm:text-sm font-medium text-slate-600">
              {settings.identity?.loginSubtitle || 'Portal E-Guru v10.2 · Masuk Akun Portal'}
            </p>
          </div>

          <div className="mt-5 space-y-4">
            {/* Tahun Ajaran & Semester Compact Tactile Pills */}
            <div className="grid grid-cols-2 gap-2.5">
              <div className="flex items-center gap-2 bg-white rounded-full p-1.5 pr-3 border border-slate-100 shadow-[0_4px_14px_rgba(15,23,42,0.06)]">
                <span className="w-8 h-8 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0">
                  <Calendar className="w-3.5 h-3.5" />
                </span>
                <select
                  id="login-year"
                  aria-label="Tahun Ajaran"
                  value={academicYear}
                  onChange={(e) => {
                    setAcademicYear(e.target.value);
                    setPeriodChangedByUser(true);
                  }}
                  className="w-full bg-transparent text-xs font-semibold text-slate-700 focus:outline-none cursor-pointer"
                >
                  <option value="">--Pilih Tahun Ajaran--</option>
                  {settings.masterAcademicYears.map((yr) => (
                    <option key={yr} value={yr}>
                      {yr}
                    </option>
                  ))}
                </select>
              </div>

              <div className="flex items-center gap-2 bg-white rounded-full p-1.5 pr-3 border border-slate-100 shadow-[0_4px_14px_rgba(15,23,42,0.06)]">
                <span className="w-8 h-8 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0">
                  <Layers className="w-3.5 h-3.5" />
                </span>
                <select
                  id="login-semester"
                  aria-label="Semester"
                  value={semester}
                  onChange={(e) => {
                    setSemester(e.target.value as SemesterType | '');
                    setPeriodChangedByUser(true);
                  }}
                  className="w-full bg-transparent text-xs font-semibold text-slate-700 focus:outline-none cursor-pointer"
                >
                  <option value="">--Pilih Semester--</option>
                  <option value="Ganjil">Ganjil</option>
                  <option value="Genap">Genap</option>
                </select>
              </div>
            </div>

            {/* 2-Tab Mode Switcher */}
            <div
              role="tablist"
              aria-label="Metode Login"
              className="grid grid-cols-2 gap-1.5 p-1 rounded-2xl bg-slate-200/75 border border-white/80 shadow-inner"
            >
              <button
                type="button"
                role="tab"
                aria-selected={loginMode === 'credentials'}
                onClick={() => {
                  setLoginMode('credentials');
                  setErrorMsg('');
                }}
                className={`py-2 px-3 rounded-xl text-xs font-bold transition-all inline-flex items-center justify-center gap-1.5 ${
                  loginMode === 'credentials'
                    ? 'bg-white text-blue-700 shadow-sm'
                    : 'text-slate-600 hover:text-slate-800'
                }`}
              >
                <KeyRound className="w-3.5 h-3.5" />
                <span>Username & Password</span>
              </button>
              <button
                type="button"
                role="tab"
                aria-selected={loginMode === 'google'}
                onClick={() => {
                  setLoginMode('google');
                  setErrorMsg('');
                }}
                className={`py-2 px-3 rounded-xl text-xs font-bold transition-all inline-flex items-center justify-center gap-1.5 ${
                  loginMode === 'google'
                    ? 'bg-white text-blue-700 shadow-sm'
                    : 'text-slate-600 hover:text-slate-800'
                }`}
              >
                <ShieldCheck className="w-3.5 h-3.5" />
                <span>Akun Google</span>
              </button>
            </div>

            {accessDeniedMessage && (
              <div className="rounded-2xl bg-amber-50 border border-amber-200 p-4 space-y-2.5">
                <div className="flex items-start gap-2.5">
                  <AlertCircle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                  <div className="text-xs text-amber-900 space-y-1">
                    <p className="font-bold">Akses Belum Diaktifkan</p>
                    {pendingUserEmail && (
                      <p className="font-mono-tabular text-[11px] text-amber-700">
                        Akun: {pendingUserEmail}
                      </p>
                    )}
                    <p className="leading-relaxed">{accessDeniedMessage}</p>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={onSwitchAccount}
                  className="w-full py-2 px-3 rounded-xl bg-white border border-amber-300 hover:bg-amber-100/60 text-amber-900 text-xs font-semibold inline-flex items-center justify-center gap-1.5 transition-colors"
                >
                  <LogOut className="w-3.5 h-3.5" />
                  Ganti Akun Google Lain
                </button>
              </div>
            )}

            {errorMsg && (
              <div className="px-4 py-2.5 rounded-2xl bg-rose-50 border border-rose-200 text-xs font-medium text-rose-700 text-center">
                {errorMsg}
              </div>
            )}

            {loginMode === 'credentials' ? (
              <form onSubmit={handleCredentialsSubmit} className="space-y-3">
                <div className="relative">
                  <User className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                  <input
                    id="login-username"
                    type="text"
                    autoComplete="username"
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                    placeholder="Username (contoh: budi.santoso)"
                    aria-label="Username"
                    className="w-full pl-10 pr-4 py-2.5 rounded-xl bg-white border border-slate-200 text-sm text-slate-800 placeholder:text-slate-400 shadow-xs focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div className="relative">
                  <Lock className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
                  <input
                    id="login-password"
                    type={showPassword ? 'text' : 'password'}
                    autoComplete="current-password"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    placeholder="Password akun"
                    aria-label="Password"
                    className="w-full pl-10 pr-10 py-2.5 rounded-xl bg-white border border-slate-200 text-sm text-slate-800 placeholder:text-slate-400 shadow-xs focus:outline-none focus:border-blue-500"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword((prev) => !prev)}
                    aria-label={showPassword ? 'Sembunyikan password' : 'Tampilkan password'}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-0.5"
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>

                <button
                  type="submit"
                  disabled={authLoading || signingIn}
                  className="w-full py-3.5 px-6 rounded-full bg-gradient-to-b from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 disabled:opacity-60 disabled:cursor-not-allowed text-white text-sm font-bold tracking-wide shadow-[0_6px_0_#1d4ed8,0_12px_22px_rgba(37,99,235,0.35)] active:translate-y-1 active:shadow-[0_2px_0_#1d4ed8] transition-all flex items-center justify-center gap-2.5"
                >
                  {authLoading || signingIn ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>{authLoading ? 'Memeriksa Sesi...' : 'Memverifikasi Akun...'}</span>
                    </>
                  ) : (
                    <>
                      <KeyRound className="w-4 h-4" />
                      <span>Masuk dengan Username</span>
                      <ArrowRight className="w-4 h-4" />
                    </>
                  )}
                </button>
              </form>
            ) : (
              <button
                type="button"
                disabled={authLoading || signingIn}
                onClick={handleGoogleClick}
                className="w-full py-3.5 px-6 rounded-full bg-gradient-to-b from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 disabled:opacity-60 disabled:cursor-not-allowed text-white text-sm font-bold tracking-wide shadow-[0_6px_0_#1d4ed8,0_12px_22px_rgba(37,99,235,0.35)] active:translate-y-1 active:shadow-[0_2px_0_#1d4ed8] transition-all flex items-center justify-center gap-2.5"
              >
                {authLoading || signingIn ? (
                  <>
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>{authLoading ? 'Memeriksa Sesi...' : 'Menghubungkan ke Google...'}</span>
                  </>
                ) : (
                  <>
                    <ShieldCheck className="w-4 h-4" />
                    <span>Masuk dengan Google</span>
                    <ArrowRight className="w-4 h-4" />
                  </>
                )}
              </button>
            )}
          </div>

          <div className="mt-6 pt-4 border-t border-slate-200/80">
            <div className="flex items-center gap-3 text-[11px] font-medium text-slate-400">
              <span className="h-px flex-1 bg-slate-200" />
              <span>{settings.identity?.loginFooterNote || 'Portal E-Guru Cloud'}</span>
              <span className="h-px flex-1 bg-slate-200" />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
