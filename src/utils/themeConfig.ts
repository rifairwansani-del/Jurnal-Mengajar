import { ThemeId } from '../types';

export interface PortalThemeDefinition {
  id: ThemeId;
  name: string;
  description: string;
  swatchPrimary: string;
  swatchSidebar: string;
  sidebarBg: string;
  sidebarBorder: string;
  sidebarActiveBg: string;
  accentBg: string;
  accentHoverBg: string;
  accentText: string;
  accentLightBg: string;
  bannerGradient: string;
  canvasBg: string;
  surfaceBg: string;
  surfaceBorder: string;
  textPrimary: string;
  textMuted: string;
  isDark: boolean;
}

export const PORTAL_THEMES: Record<ThemeId, PortalThemeDefinition> = {
  'slate-indigo': {
    id: 'slate-indigo',
    name: 'Slate & Indigo (Default)',
    description: 'Gelap profesional pada sidebar dengan aksen indigo modern',
    swatchPrimary: '#4f46e5',
    swatchSidebar: '#0f172a',
    sidebarBg: 'bg-slate-900',
    sidebarBorder: 'border-slate-800',
    sidebarActiveBg: 'bg-indigo-600 text-white shadow-sm',
    accentBg: 'bg-indigo-600',
    accentHoverBg: 'hover:bg-indigo-700',
    accentText: 'text-indigo-600',
    accentLightBg: 'bg-indigo-50',
    bannerGradient: 'from-indigo-600 via-indigo-700 to-violet-700',
    canvasBg: 'bg-slate-50',
    surfaceBg: 'bg-white',
    surfaceBorder: 'border-slate-200',
    textPrimary: 'text-slate-900',
    textMuted: 'text-slate-500',
    isDark: false,
  },
  'royal-blue': {
    id: 'royal-blue',
    name: 'Royal Blue Akademik',
    description: 'Nuansa biru tua akademik dengan kontras cerah dan tegas',
    swatchPrimary: '#2563eb',
    swatchSidebar: '#172554',
    sidebarBg: 'bg-blue-950',
    sidebarBorder: 'border-blue-900',
    sidebarActiveBg: 'bg-blue-600 text-white shadow-sm',
    accentBg: 'bg-blue-600',
    accentHoverBg: 'hover:bg-blue-700',
    accentText: 'text-blue-600',
    accentLightBg: 'bg-blue-50',
    bannerGradient: 'from-blue-600 via-blue-700 to-sky-700',
    canvasBg: 'bg-sky-50/40',
    surfaceBg: 'bg-white',
    surfaceBorder: 'border-slate-200',
    textPrimary: 'text-slate-900',
    textMuted: 'text-slate-500',
    isDark: false,
  },
  'emerald-forest': {
    id: 'emerald-forest',
    name: 'Emerald Pendidikan',
    description: 'Sidebar hijau zamrud tenang dengan aksen segar dan bersih',
    swatchPrimary: '#059669',
    swatchSidebar: '#022c22',
    sidebarBg: 'bg-emerald-950',
    sidebarBorder: 'border-emerald-900',
    sidebarActiveBg: 'bg-emerald-600 text-white shadow-sm',
    accentBg: 'bg-emerald-600',
    accentHoverBg: 'hover:bg-emerald-700',
    accentText: 'text-emerald-600',
    accentLightBg: 'bg-emerald-50',
    bannerGradient: 'from-emerald-600 via-teal-700 to-cyan-700',
    canvasBg: 'bg-slate-50',
    surfaceBg: 'bg-white',
    surfaceBorder: 'border-slate-200',
    textPrimary: 'text-slate-900',
    textMuted: 'text-slate-500',
    isDark: false,
  },
  'crimson-charcoal': {
    id: 'crimson-charcoal',
    name: 'Charcoal & Crimson',
    description: 'Sidebar arang elegan dengan aksen merah marun berkarakter',
    swatchPrimary: '#e11d48',
    swatchSidebar: '#18181b',
    sidebarBg: 'bg-zinc-900',
    sidebarBorder: 'border-zinc-800',
    sidebarActiveBg: 'bg-rose-600 text-white shadow-sm',
    accentBg: 'bg-rose-600',
    accentHoverBg: 'hover:bg-rose-700',
    accentText: 'text-rose-600',
    accentLightBg: 'bg-rose-50',
    bannerGradient: 'from-rose-600 via-rose-700 to-slate-800',
    canvasBg: 'bg-zinc-50',
    surfaceBg: 'bg-white',
    surfaceBorder: 'border-zinc-200',
    textPrimary: 'text-zinc-900',
    textMuted: 'text-zinc-500',
    isDark: false,
  },
  'midnight-dark': {
    id: 'midnight-dark',
    name: 'Midnight Gelap',
    description: 'Mode gelap menyeluruh nyaman untuk mata saat bekerja malam',
    swatchPrimary: '#6366f1',
    swatchSidebar: '#020617',
    sidebarBg: 'bg-slate-950',
    sidebarBorder: 'border-slate-800',
    sidebarActiveBg: 'bg-indigo-600 text-white shadow-sm',
    accentBg: 'bg-indigo-600',
    accentHoverBg: 'hover:bg-indigo-500',
    accentText: 'text-indigo-400',
    accentLightBg: 'bg-indigo-950/60',
    bannerGradient: 'from-indigo-700 via-slate-900 to-violet-900',
    canvasBg: 'bg-slate-950',
    surfaceBg: 'bg-slate-900',
    surfaceBorder: 'border-slate-800',
    textPrimary: 'text-slate-100',
    textMuted: 'text-slate-400',
    isDark: true,
  },
};
