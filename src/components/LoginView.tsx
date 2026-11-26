import React, { useState } from 'react';
import {
  ArrowRight,
  Calendar,
  Crown,
  Database,
  Eye,
  EyeOff,
  FilePlus2,
  Layers,
  Lock,
  User,
} from 'lucide-react';
import { SemesterType, SystemSettings } from '../types';

interface LoginViewProps {
  settings: SystemSettings;
  onLogin: (
    email: string,
    password: string,
    academicYear: string,
    semester: SemesterType
  ) => { ok: boolean; message: string };
  onDemoSeed: (academicYear: string, semester: SemesterType) => void;
  onStartEmpty: (email: string, academicYear: string, semester: SemesterType) => void;
}

export function LoginView({
  settings,
  onLogin,
  onDemoSeed,
  onStartEmpty,
}: LoginViewProps) {
  const [email, setEmail] = useState(settings.email || 'guru@sma1kudus.sch.id');
  const [password, setPassword] = useState(settings.password || 'guru123');
  const [showPassword, setShowPassword] = useState(false);
  const [academicYear, setAcademicYear] = useState(settings.academicYear || '2025/2026');
  const [semester, setSemester] = useState<SemesterType>(settings.semester || 'Ganjil');
  const [errorMsg, setErrorMsg] = useState('');

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg('');
    if (!email.trim() || !password.trim()) {
      setErrorMsg('Email Guru dan Password wajib diisi.');
      return;
    }
    const res = onLogin(email, password, academicYear, semester);
    if (!res.ok) {
      setErrorMsg(res.message);
    }
  };

  return (
    <div className="relative min-h-screen overflow-hidden bg-gradient-to-b from-blue-500 via-sky-300 to-blue-200 flex items-center justify-center p-4 sm:p-6">
      {/* Soft 3D Study-Room Background Atmosphere (No people, no animals) */}
      <div className="pointer-events-none absolute inset-0 overflow-hidden" aria-hidden="true">
        {/* Sunlit window glow */}
        <div className="absolute -top-24 right-[-10%] w-[460px] h-[520px] rounded-full bg-white/35 blur-3xl" />
        <div className="absolute top-12 left-[-8%] w-[360px] h-[360px] rounded-full bg-amber-200/35 blur-3xl" />

        {/* Stylized wall clock */}
        <div className="hidden sm:flex absolute top-8 left-[22%] w-16 h-16 rounded-full bg-white/90 border-4 border-blue-700 shadow-lg items-center justify-center">
          <div className="w-1 h-4 bg-blue-900 rounded-full origin-bottom -translate-y-1.5 rotate-12" />
          <div className="w-1 h-3 bg-blue-600 rounded-full origin-bottom -translate-y-1 -rotate-45 absolute" />
          <div className="w-2 h-2 rounded-full bg-blue-800 absolute" />
        </div>

        {/* Soft puffy clouds */}
        <div className="hidden md:block absolute top-10 right-[24%] w-32 h-12 rounded-full bg-white/85 shadow-md">
          <div className="absolute -top-5 left-5 w-14 h-14 rounded-full bg-white/90" />
          <div className="absolute -top-7 left-12 w-16 h-16 rounded-full bg-white/95" />
        </div>
        <div className="hidden lg:block absolute top-52 right-[12%] w-24 h-10 rounded-full bg-white/80 shadow-md">
          <div className="absolute -top-4 left-4 w-11 h-11 rounded-full bg-white/90" />
          <div className="absolute -top-5 left-10 w-12 h-12 rounded-full bg-white/90" />
        </div>

        {/* Desk surface & pedestal */}
        <div className="absolute bottom-0 inset-x-0 h-28 bg-gradient-to-b from-blue-300/80 to-indigo-400/90 border-t border-white/50" />
        <div className="absolute bottom-14 left-1/2 -translate-x-1/2 w-[500px] max-w-[92vw] h-10 rounded-full bg-blue-700/25 blur-md" />

        {/* Left decorative stacked books */}
        <div className="hidden lg:block absolute bottom-20 left-10 space-y-1 opacity-90">
          <div className="w-40 h-6 rounded-r-lg bg-blue-600 border-l-8 border-white shadow-md" />
          <div className="w-44 h-7 rounded-r-lg bg-indigo-700 border-l-8 border-sky-100 shadow-md" />
          <div className="w-48 h-7 rounded-r-lg bg-slate-800 border-l-8 border-blue-200 shadow-md" />
        </div>

        {/* Right decorative stacked books */}
        <div className="hidden lg:block absolute bottom-20 right-10 space-y-1 opacity-90">
          <div className="w-36 h-6 rounded-l-lg bg-sky-500 border-r-8 border-white shadow-md ml-auto" />
          <div className="w-44 h-7 rounded-l-lg bg-blue-700 border-r-8 border-blue-100 shadow-md ml-auto" />
          <div className="w-48 h-7 rounded-l-lg bg-indigo-900 border-r-8 border-slate-100 shadow-md ml-auto" />
        </div>
      </div>

      {/* Main Claymorphic Login Card */}
      <div className="relative z-10 w-full max-w-[440px]">
        {/* Pedestal base */}
        <div className="mx-auto w-[86%] h-4 bg-blue-700/35 rounded-b-full blur-[2px] -mb-2 translate-y-full" />

        <div className="bg-[#faf9f6] rounded-[2.25rem] p-6 sm:p-8 border-2 border-white shadow-[0_26px_60px_-15px_rgba(30,58,138,0.42),inset_0_-8px_18px_rgba(148,163,184,0.16),inset_0_4px_12px_rgba(255,255,255,0.95)]">
          {/* Header with Logo / Crown Emblem */}
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
              {settings.identity?.loginSubtitle || 'Portal E-Guru v10.2 · Masuk untuk melanjutkan'}
            </p>
          </div>

          <form onSubmit={handleSubmit} className="mt-6 space-y-3.5">
            {/* Email / Username Tactile Pill Input */}
            <div>
              <label htmlFor="login-email" className="sr-only">
                Email Guru
              </label>
              <div className="flex items-center gap-3 bg-white rounded-full p-1.5 pr-4 border border-slate-100 shadow-[0_6px_18px_rgba(15,23,42,0.08),inset_0_2px_4px_rgba(255,255,255,1)] focus-within:ring-2 focus-within:ring-blue-500/40 transition-shadow">
                <span className="w-10 h-10 rounded-full bg-gradient-to-b from-blue-500 to-blue-600 text-white shadow-[0_4px_10px_rgba(37,99,235,0.4)] flex items-center justify-center shrink-0">
                  <User className="w-4 h-4" />
                </span>
                <input
                  id="login-email"
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="Email Guru"
                  className="w-full bg-transparent text-sm font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none"
                />
              </div>
            </div>

            {/* Password Tactile Pill Input */}
            <div>
              <label htmlFor="login-password" className="sr-only">
                Password
              </label>
              <div className="flex items-center gap-3 bg-white rounded-full p-1.5 pr-3 border border-slate-100 shadow-[0_6px_18px_rgba(15,23,42,0.08),inset_0_2px_4px_rgba(255,255,255,1)] focus-within:ring-2 focus-within:ring-blue-500/40 transition-shadow">
                <span className="w-10 h-10 rounded-full bg-gradient-to-b from-blue-500 to-blue-600 text-white shadow-[0_4px_10px_rgba(37,99,235,0.4)] flex items-center justify-center shrink-0">
                  <Lock className="w-4 h-4" />
                </span>
                <input
                  id="login-password"
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Password"
                  className="w-full bg-transparent text-sm font-medium text-slate-800 placeholder:text-slate-400 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword((prev) => !prev)}
                  className="p-1.5 rounded-full text-slate-400 hover:text-blue-600 transition-colors shrink-0"
                  aria-label={showPassword ? 'Sembunyikan password' : 'Tampilkan password'}
                >
                  {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
            </div>

            {/* Tahun Ajaran & Semester Compact Tactile Pills */}
            <div className="grid grid-cols-2 gap-2.5 pt-0.5">
              <div className="flex items-center gap-2 bg-white rounded-full p-1.5 pr-3 border border-slate-100 shadow-[0_4px_14px_rgba(15,23,42,0.06)]">
                <span className="w-8 h-8 rounded-full bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0">
                  <Calendar className="w-3.5 h-3.5" />
                </span>
                <select
                  id="login-year"
                  aria-label="Tahun Ajaran"
                  value={academicYear}
                  onChange={(e) => setAcademicYear(e.target.value)}
                  className="w-full bg-transparent text-xs font-semibold text-slate-700 focus:outline-none cursor-pointer"
                >
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
                  onChange={(e) => setSemester(e.target.value as SemesterType)}
                  className="w-full bg-transparent text-xs font-semibold text-slate-700 focus:outline-none cursor-pointer"
                >
                  <option value="Ganjil">Ganjil</option>
                  <option value="Genap">Genap</option>
                </select>
              </div>
            </div>

            {errorMsg && (
              <div className="px-4 py-2.5 rounded-2xl bg-rose-50 border border-rose-200 text-xs font-medium text-rose-700 text-center">
                {errorMsg}
              </div>
            )}

            {/* Primary Tactile 3D Login Button */}
            <div className="pt-1">
              <button
                type="submit"
                className="w-full py-3 px-6 rounded-full bg-gradient-to-b from-blue-500 to-blue-600 hover:from-blue-600 hover:to-blue-700 text-white text-sm font-bold tracking-wide shadow-[0_6px_0_#1d4ed8,0_12px_22px_rgba(37,99,235,0.35)] active:translate-y-1 active:shadow-[0_2px_0_#1d4ed8] transition-all flex items-center justify-center gap-2"
              >
                <span>Masuk</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </form>

          {/* Local Access Options (Without Social Media) */}
          <div className="mt-6 pt-4 border-t border-slate-200/80 space-y-2.5">
            <div className="flex items-center gap-3 text-[11px] font-medium text-slate-400">
              <span className="h-px flex-1 bg-slate-200" />
              <span>{settings.identity?.loginFooterNote || 'Akses Penyimpanan Lokal'}</span>
              <span className="h-px flex-1 bg-slate-200" />
            </div>

            <button
              type="button"
              onClick={() => onDemoSeed(academicYear, semester)}
              className="w-full py-2.5 px-4 rounded-full bg-slate-900 hover:bg-slate-800 text-white text-xs sm:text-sm font-semibold shadow-[0_4px_12px_rgba(15,23,42,0.18)] transition-colors flex items-center justify-center gap-2"
            >
              <Database className="w-4 h-4 text-blue-400" />
              Masuk dengan Data Demo (Auto-Seed)
            </button>

            <button
              type="button"
              onClick={() => onStartEmpty(email, academicYear, semester)}
              className="w-full py-2.5 px-4 rounded-full bg-white hover:bg-slate-50 text-slate-700 text-xs sm:text-sm font-semibold border border-slate-200 shadow-xs transition-colors flex items-center justify-center gap-2"
            >
              <FilePlus2 className="w-4 h-4 text-slate-500" />
              Masuk Kosong / Mulai Baru
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
