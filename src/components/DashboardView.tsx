import React, { useEffect, useMemo, useRef, useState } from 'react';
import Chart from 'chart.js/auto';
import {
  BookOpen,
  CalendarCheck,
  Award,
  Users,
  FileSpreadsheet,
  Plus,
  Printer,
  Copy,
  Pencil,
  Trash2,
} from 'lucide-react';
import { JournalEntry, MenuId, PortalDatabase } from '../types';
import {
  exportDatasetToPdf,
  exportDatasetToWordDoc,
  exportRowsToXls,
  exportRowsToXlsx,
  printDatasetExact,
} from '../utils/exportHelpers';
import {
  computeClassSummary,
  formatShortDateId,
  getTodayISO,
  summarizeAttendanceCounts,
  summarizeJournalRow,
} from '../utils/portalHelpers';
import { PORTAL_THEMES } from '../utils/themeConfig';
import { ConfirmDialogState } from './FeedbackUI';
import { TableExportToolbar } from './TableExportToolbar';

interface DashboardViewProps {
  db: PortalDatabase;
  onNavigate: (menu: MenuId) => void;
  onEditJournal: (journal: JournalEntry) => void;
  onDuplicateJournal: (journal: JournalEntry) => void;
  onDeleteJournals: (ids: string[]) => void;
  onToast: (type: 'success' | 'error' | 'info', message: string) => void;
  onConfirm: (cfg: Omit<ConfirmDialogState, 'open'>) => void;
}

export function DashboardView({
  db,
  onNavigate,
  onEditJournal,
  onDuplicateJournal,
  onDeleteJournals,
  onToast,
  onConfirm,
}: DashboardViewProps) {
  const { students, journals, grades, settings, session } = db;
  const theme = PORTAL_THEMES[settings.themeId] || PORTAL_THEMES['slate-indigo'];
  const [selectedLatestIds, setSelectedLatestIds] = useState<Set<string>>(new Set());

  const activeJournals = useMemo(
    () =>
      journals
        .filter(
          (j) =>
            j.academicYear === session.academicYear &&
            j.semester === session.semester
        )
        .sort((a, b) => (a.date < b.date ? 1 : -1)),
    [journals, session.academicYear, session.semester]
  );

  const activeGrades = useMemo(
    () =>
      grades.filter(
        (g) =>
          g.academicYear === session.academicYear &&
          g.semester === session.semester &&
          g.score !== null
      ),
    [grades, session.academicYear, session.semester]
  );

  const attendanceOverview = useMemo(
    () => summarizeAttendanceCounts(activeJournals),
    [activeJournals]
  );

  const averageScore = useMemo(() => {
    if (activeGrades.length === 0) return 0;
    const sum = activeGrades.reduce((acc, g) => acc + (g.score ?? 0), 0);
    return sum / activeGrades.length;
  }, [activeGrades]);

  const uniqueExamsCount = useMemo(() => {
    const keys = new Set(
      activeGrades.map((g) => `${g.className}__${g.subject}__${g.examType}`)
    );
    return keys.size;
  }, [activeGrades]);

  const classList = useMemo(() => {
    const set = new Set(settings.masterClasses);
    students.forEach((s) => set.add(s.className));
    return Array.from(set);
  }, [settings.masterClasses, students]);

  const classSummaries = useMemo(
    () => computeClassSummary(classList, students, activeJournals, activeGrades),
    [classList, students, activeJournals, activeGrades]
  );

  // Chart refs
  const doughnutRef = useRef<HTMLCanvasElement | null>(null);
  const attBarRef = useRef<HTMLCanvasElement | null>(null);
  const gradeBarRef = useRef<HTMLCanvasElement | null>(null);

  useEffect(() => {
    const canvas = doughnutRef.current;
    if (!canvas) return;

    const { counts } = attendanceOverview;
    const hasData = attendanceOverview.totalRecords > 0;

    const chart = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: hasData
          ? ['Hadir (H)', 'Sakit (S)', 'Izin (I)', 'Alpha (A)', 'Dispensasi (D)']
          : ['Belum Ada Data'],
        datasets: [
          {
            data: hasData
              ? [counts.H, counts.S, counts.I, counts.A, counts.D]
              : [1],
            backgroundColor: hasData
              ? ['#10b981', '#f59e0b', '#0ea5e9', '#f43f5e', '#8b5cf6']
              : ['#e2e8f0'],
            borderWidth: 0,
            hoverOffset: 4,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '66%',
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              boxWidth: 12,
              padding: 14,
              font: { family: 'Plus Jakarta Sans', size: 12 },
            },
          },
          tooltip: {
            enabled: hasData,
          },
        },
      },
    });

    return () => chart.destroy();
  }, [attendanceOverview]);

  useEffect(() => {
    const canvas = attBarRef.current;
    if (!canvas) return;

    const chart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: classSummaries.map((c) => c.className),
        datasets: [
          {
            label: 'Kehadiran (%)',
            data: classSummaries.map((c) => Number(c.presentPercent.toFixed(1))),
            backgroundColor: '#4f46e5',
            borderRadius: 6,
            maxBarThickness: 40,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            ticks: { callback: (v) => `${v}%`, font: { size: 11 } },
            grid: { color: '#f1f5f9' },
          },
          x: {
            grid: { display: false },
            ticks: { font: { size: 11 } },
          },
        },
        plugins: {
          legend: { display: false },
        },
      },
    });

    return () => chart.destroy();
  }, [classSummaries]);

  useEffect(() => {
    const canvas = gradeBarRef.current;
    if (!canvas) return;

    const chart = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: classSummaries.map((c) => c.className),
        datasets: [
          {
            label: 'Rata-rata Nilai',
            data: classSummaries.map((c) => Number(c.avgScore.toFixed(1))),
            backgroundColor: '#0f172a',
            borderRadius: 6,
            maxBarThickness: 40,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: true,
            max: 100,
            ticks: { font: { size: 11 } },
            grid: { color: '#f1f5f9' },
          },
          x: {
            grid: { display: false },
            ticks: { font: { size: 11 } },
          },
        },
        plugins: {
          legend: { display: false },
        },
      },
    });

    return () => chart.destroy();
  }, [classSummaries]);

  const latestJournals = activeJournals.slice(0, 8);

  const toggleSelectLatest = (id: string) => {
    setSelectedLatestIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const toggleSelectAllLatest = () => {
    if (
      latestJournals.length > 0 &&
      latestJournals.every((j) => selectedLatestIds.has(j.id))
    ) {
      setSelectedLatestIds(new Set());
    } else {
      setSelectedLatestIds(new Set(latestJournals.map((j) => j.id)));
    }
  };

  const handleBulkDeleteLatest = () => {
    if (selectedLatestIds.size === 0) return;
    const count = selectedLatestIds.size;
    onConfirm({
      title: 'Hapus Jurnal Terpilih',
      description: `Hapus ${count} catatan jurnal terpilih beserta data presensinya?`,
      confirmLabel: `Hapus (${count})`,
      danger: true,
      onConfirm: () => {
        onDeleteJournals(Array.from(selectedLatestIds));
        setSelectedLatestIds(new Set());
        onToast('success', `${count} jurnal berhasil dihapus.`);
      },
    });
  };

  const handleSingleDeleteLatest = (entry: JournalEntry) => {
    onConfirm({
      title: 'Hapus Jurnal Mengajar',
      description: `Hapus jurnal "${entry.topic}" kelas ${entry.className}?`,
      confirmLabel: 'Hapus Jurnal',
      danger: true,
      onConfirm: () => {
        onDeleteJournals([entry.id]);
        onToast('success', 'Jurnal berhasil dihapus.');
      },
    });
  };

  const buildClassSummaryDataset = () => ({
    headers: ['Kelas', 'Siswa', 'Jurnal', 'H', 'S', 'I', 'A', 'D', '% Hadir', 'Rata-rata Nilai'],
    rows: classSummaries.map((r) => [
      r.className,
      r.studentCount,
      r.journalCount,
      r.counts.H,
      r.counts.S,
      r.counts.I,
      r.counts.A,
      r.counts.D,
      `${r.presentPercent.toFixed(1)}%`,
      r.gradedCount > 0 ? Number(r.avgScore.toFixed(1)) : '-',
    ]),
  });

  const exportLatestSource =
    selectedLatestIds.size > 0
      ? latestJournals.filter((j) => selectedLatestIds.has(j.id))
      : latestJournals;

  const buildLatestDataset = () => ({
    headers: ['No', 'Tanggal', 'Jam', 'Kelas', 'Mata Pelajaran', 'Materi Pokok', 'H', 'S', 'I', 'A', 'D'],
    rows: exportLatestSource.map((j, idx) => {
      const s = summarizeJournalRow(j);
      return [
        idx + 1,
        formatShortDateId(j.date),
        j.jamKe,
        j.className,
        j.subject,
        j.topic,
        s.counts.H,
        s.counts.S,
        s.counts.I,
        s.counts.A,
        s.counts.D,
      ];
    }),
  });

  return (
    <div className="space-y-6">
      {/* Banner Gradient Dinamis Sesuai Tema */}
      <section className={`rounded-xl bg-gradient-to-r ${theme.bannerGradient} text-white p-6 sm:p-8 shadow-sm`}>
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-6">
          <div className="space-y-2 max-w-2xl">
            <p className="text-xs font-medium text-indigo-100">
              {settings.schoolName} · Tahun Ajaran {session.academicYear} · Semester {session.semester}
            </p>
            <h2 className="text-2xl sm:text-3xl font-bold tracking-tight text-balance">
              Selamat datang, {settings.teacherName}
            </h2>
            <p className="text-sm text-indigo-100 leading-relaxed">
              Pantau rekap kegiatan belajar mengajar, statistik kehadiran kelas, dan capaian nilai evaluasi siswa secara langsung dari data lokal Anda.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            <button
              type="button"
              onClick={() => onNavigate('journal')}
              className="px-4 py-2.5 rounded-lg bg-white text-indigo-700 hover:bg-indigo-50 text-xs sm:text-sm font-semibold transition-colors flex items-center gap-2 whitespace-nowrap shrink-0"
            >
              <Plus className="w-4 h-4" />
              Catat Jurnal Baru
            </button>
            <button
              type="button"
              onClick={() => onNavigate('grades')}
              className="px-4 py-2.5 rounded-lg bg-indigo-500/40 hover:bg-indigo-500/60 border border-white/25 text-white text-xs sm:text-sm font-semibold transition-colors flex items-center gap-2 whitespace-nowrap shrink-0"
            >
              <FileSpreadsheet className="w-4 h-4" />
              Input Nilai
            </button>
            <button
              type="button"
              onClick={() => onNavigate('reports')}
              className="px-4 py-2.5 rounded-lg bg-slate-900/40 hover:bg-slate-900/60 border border-white/20 text-white text-xs sm:text-sm font-semibold transition-colors flex items-center gap-2 whitespace-nowrap shrink-0"
            >
              <Printer className="w-4 h-4" />
              Cetak Laporan
            </button>
          </div>
        </div>
      </section>

      {/* 5 Kartu Statistik Utama */}
      <section className="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-5 gap-4">
        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
          <div className="flex items-center justify-between text-slate-500">
            <span className="text-xs font-medium">Total Jurnal</span>
            <BookOpen className="w-4 h-4 text-indigo-600" />
          </div>
          <p className="mt-3 text-2xl font-bold text-slate-900 tabular-nums">
            {activeJournals.length}
          </p>
          <p className="mt-1 text-xs text-slate-500">Pertemuan KBM tercatat</p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
          <div className="flex items-center justify-between text-slate-500">
            <span className="text-xs font-medium">Kehadiran (H)</span>
            <CalendarCheck className="w-4 h-4 text-emerald-600" />
          </div>
          <p className="mt-3 text-2xl font-bold text-slate-900 tabular-nums">
            {attendanceOverview.presentPercent.toFixed(1)}%
          </p>
          <p className="mt-1 text-xs text-slate-500 tabular-nums">
            {attendanceOverview.counts.H} dari {attendanceOverview.totalRecords} entri
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
          <div className="flex items-center justify-between text-slate-500">
            <span className="text-xs font-medium">Rata-rata Nilai</span>
            <Award className="w-4 h-4 text-indigo-600" />
          </div>
          <p className="mt-3 text-2xl font-bold text-slate-900 tabular-nums">
            {averageScore.toFixed(1)}
          </p>
          <p className="mt-1 text-xs text-slate-500 tabular-nums">
            {activeGrades.length} nilai tersimpan
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
          <div className="flex items-center justify-between text-slate-500">
            <span className="text-xs font-medium">Total Siswa</span>
            <Users className="w-4 h-4 text-slate-700" />
          </div>
          <p className="mt-3 text-2xl font-bold text-slate-900 tabular-nums">
            {students.length}
          </p>
          <p className="mt-1 text-xs text-slate-500 tabular-nums">
            Tersebar di {classList.length} kelas
          </p>
        </div>

        <div className="col-span-2 sm:col-span-1 bg-white border border-slate-200 rounded-xl p-4 sm:p-5">
          <div className="flex items-center justify-between text-slate-500">
            <span className="text-xs font-medium">Ujian Dinilai</span>
            <FileSpreadsheet className="w-4 h-4 text-violet-600" />
          </div>
          <p className="mt-3 text-2xl font-bold text-slate-900 tabular-nums">
            {uniqueExamsCount}
          </p>
          <p className="mt-1 text-xs text-slate-500">Sesi evaluasi terdata</p>
        </div>
      </section>

      {/* 3 Grafik Interaktif (Chart.js) */}
      <section className="grid grid-cols-1 lg:grid-cols-3 gap-5">
        <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col">
          <h3 className="text-sm font-bold text-slate-900">Komposisi Kehadiran</h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Distribusi status H, S, I, A, dan D pada periode aktif
          </p>
          <div className="mt-4 h-64 relative">
            <canvas ref={doughnutRef} />
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col">
          <h3 className="text-sm font-bold text-slate-900">Persentase Kehadiran per Kelas</h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Rasio kehadiran (H) dibandingkan total presensi per kelas
          </p>
          <div className="mt-4 h-64 relative">
            <canvas ref={attBarRef} />
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-xl p-5 flex flex-col">
          <h3 className="text-sm font-bold text-slate-900">Rata-rata Nilai per Kelas</h3>
          <p className="text-xs text-slate-500 mt-0.5">
            Capaian rerata skor seluruh jenis ujian per kelas
          </p>
          <div className="mt-4 h-64 relative">
            <canvas ref={gradeBarRef} />
          </div>
        </div>
      </section>

      {/* Tabel Rekap Kehadiran & Nilai per Kelas */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          <div>
            <h3 className="text-sm font-bold text-slate-900">Rekap Kehadiran & Nilai per Kelas</h3>
            <p className="text-xs text-slate-500">Ringkasan siswa, pertemuan KBM, presensi, dan nilai rata-rata</p>
          </div>
          <TableExportToolbar
            onExportXlsx={() => {
              const { headers, rows } = buildClassSummaryDataset();
              exportRowsToXlsx('Rekap Kelas', headers, rows, `Rekap_Kelas_${getTodayISO()}.xlsx`);
              onToast('success', 'Rekap kelas diekspor ke XLSX.');
            }}
            onExportXls={() => {
              const { headers, rows } = buildClassSummaryDataset();
              exportRowsToXls('Rekap Kelas', headers, rows, `Rekap_Kelas_${getTodayISO()}.xls`);
              onToast('success', 'Rekap kelas diekspor ke XLS.');
            }}
            onExportWord={(cfg) => {
              const { headers, rows } = buildClassSummaryDataset();
              exportDatasetToWordDoc({
                db,
                title: 'REKAP KEHADIRAN & NILAI PER KELAS',
                subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                headers,
                rows,
                filename: `Rekap_Kelas_${getTodayISO()}.doc`,
                orientation: cfg.orientation,
                paperSize: cfg.paperSize,
              });
              onToast('success', 'Dokumen Word (.doc) Rekap Kelas diunduh.');
            }}
            onExportPdf={(cfg) => {
              const { headers, rows } = buildClassSummaryDataset();
              exportDatasetToPdf({
                db,
                title: 'REKAP KEHADIRAN & NILAI PER KELAS',
                subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                headers,
                rows,
                filename: `Rekap_Kelas_${getTodayISO()}.pdf`,
                orientation: cfg.orientation,
                paperSize: cfg.paperSize,
              });
              onToast('success', 'Dokumen PDF Rekap Kelas diunduh.');
            }}
            onPrint={(cfg) => {
              const { headers, rows } = buildClassSummaryDataset();
              printDatasetExact({
                db,
                title: 'REKAP KEHADIRAN & NILAI PER KELAS',
                subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                headers,
                rows,
                orientation: cfg.orientation,
                paperSize: cfg.paperSize,
              });
            }}
          />
        </div>
        <div className="overflow-x-auto max-h-[380px] table-scroll">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="sticky top-0 z-10 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="px-4 py-3 whitespace-nowrap">Kelas</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">Siswa</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">Jurnal</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">H</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">S</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">I</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">A</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">D</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">% Hadir</th>
                <th className="px-4 py-3 text-right whitespace-nowrap">Rata-rata Nilai</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {classSummaries.length === 0 ? (
                <tr>
                  <td colSpan={10} className="px-4 py-8 text-center text-sm text-slate-500">
                    Belum ada kelas terdaftar. Tambahkan kelas di Pengaturan atau impor data siswa.
                  </td>
                </tr>
              ) : (
                classSummaries.map((row) => (
                  <tr key={row.className} className="hover:bg-slate-50/80 transition-colors">
                    <td className="px-4 py-3 font-semibold text-slate-900 whitespace-nowrap">
                      {row.className}
                    </td>
                    <td className="px-4 py-3 text-right tabular-nums">{row.studentCount}</td>
                    <td className="px-4 py-3 text-right tabular-nums">{row.journalCount}</td>
                    <td className="px-4 py-3 text-right tabular-nums text-emerald-700 font-medium">{row.counts.H}</td>
                    <td className="px-4 py-3 text-right tabular-nums text-amber-700">{row.counts.S}</td>
                    <td className="px-4 py-3 text-right tabular-nums text-sky-700">{row.counts.I}</td>
                    <td className="px-4 py-3 text-right tabular-nums text-rose-700">{row.counts.A}</td>
                    <td className="px-4 py-3 text-right tabular-nums text-violet-700">{row.counts.D}</td>
                    <td className="px-4 py-3 text-right tabular-nums font-semibold text-slate-900">
                      {row.presentPercent.toFixed(1)}%
                    </td>
                    <td className="px-4 py-3 text-right tabular-nums font-semibold text-indigo-700">
                      {row.gradedCount > 0 ? row.avgScore.toFixed(1) : '-'}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      {/* Tabel Kegiatan Mengajar Terbaru */}
      <section className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-5 py-4 border-b border-slate-200 space-y-3">
          <div className="flex items-center justify-between gap-4">
            <div>
              <h3 className="text-sm font-bold text-slate-900">Kegiatan Mengajar Terbaru</h3>
              <p className="text-xs text-slate-500">Riwayat pertemuan KBM terakhir dengan akses cepat Edit, Duplikat, Hapus & Ekspor</p>
            </div>
            <button
              type="button"
              onClick={() => onNavigate('journal')}
              className="text-xs font-semibold text-indigo-600 hover:text-indigo-700 whitespace-nowrap"
            >
              Lihat Semua Jurnal
            </button>
          </div>

          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-100">
            <span className="text-xs text-slate-500">
              {selectedLatestIds.size > 0
                ? `${selectedLatestIds.size} jurnal dipilih`
                : `Menampilkan ${latestJournals.length} pertemuan terbaru`}
            </span>
            <TableExportToolbar
              selectedCount={selectedLatestIds.size}
              onDeleteSelected={handleBulkDeleteLatest}
              onExportXlsx={() => {
                const { headers, rows } = buildLatestDataset();
                exportRowsToXlsx('Jurnal Terbaru', headers, rows, `Jurnal_Terbaru_${getTodayISO()}.xlsx`);
                onToast('success', 'Jurnal terbaru diekspor ke XLSX.');
              }}
              onExportXls={() => {
                const { headers, rows } = buildLatestDataset();
                exportRowsToXls('Jurnal Terbaru', headers, rows, `Jurnal_Terbaru_${getTodayISO()}.xls`);
                onToast('success', 'Jurnal terbaru diekspor ke XLS.');
              }}
              onExportWord={(cfg) => {
                const { headers, rows } = buildLatestDataset();
                exportDatasetToWordDoc({
                  db,
                  title: 'KEGIATAN MENGAJAR TERBARU',
                  subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Jurnal_Terbaru_${getTodayISO()}.doc`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen Word (.doc) diunduh.');
              }}
              onExportPdf={(cfg) => {
                const { headers, rows } = buildLatestDataset();
                exportDatasetToPdf({
                  db,
                  title: 'KEGIATAN MENGAJAR TERBARU',
                  subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  filename: `Jurnal_Terbaru_${getTodayISO()}.pdf`,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
                onToast('success', 'Dokumen PDF diunduh.');
              }}
              onPrint={(cfg) => {
                const { headers, rows } = buildLatestDataset();
                printDatasetExact({
                  db,
                  title: 'KEGIATAN MENGAJAR TERBARU',
                  subtitle: `TA ${session.academicYear} · Semester ${session.semester}`,
                  headers,
                  rows,
                  orientation: cfg.orientation,
                  paperSize: cfg.paperSize,
                });
              }}
            />
          </div>
        </div>

        <div className="overflow-x-auto max-h-[420px] table-scroll">
          <table className="w-full text-left border-collapse text-sm">
            <thead className="sticky top-0 z-10 bg-slate-100 text-slate-700 text-xs font-semibold border-b border-slate-200">
              <tr>
                <th className="px-4 py-3 w-10 text-center">
                  <input
                    type="checkbox"
                    checked={
                      latestJournals.length > 0 &&
                      latestJournals.every((j) => selectedLatestIds.has(j.id))
                    }
                    onChange={toggleSelectAllLatest}
                    aria-label="Pilih semua jurnal terbaru"
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                  />
                </th>
                <th className="px-4 py-3 whitespace-nowrap">Tanggal</th>
                <th className="px-4 py-3 whitespace-nowrap">Jam</th>
                <th className="px-4 py-3 whitespace-nowrap">Kelas</th>
                <th className="px-4 py-3 whitespace-nowrap">Mata Pelajaran</th>
                <th className="px-4 py-3 min-w-[220px]">Materi Pokok</th>
                <th className="px-4 py-3 whitespace-nowrap text-right">Presensi</th>
                <th className="px-4 py-3 whitespace-nowrap text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-slate-700">
              {latestJournals.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-4 py-10 text-center text-sm text-slate-500">
                    Belum ada catatan jurnal mengajar pada periode ini. Klik &ldquo;Catat Jurnal Baru&rdquo; untuk memulai.
                  </td>
                </tr>
              ) : (
                latestJournals.map((j) => {
                  const sum = summarizeJournalRow(j);
                  const checked = selectedLatestIds.has(j.id);
                  return (
                    <tr key={j.id} className={`hover:bg-slate-50/80 transition-colors ${checked ? 'bg-indigo-50/40' : ''}`}>
                      <td className="px-4 py-3 text-center">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleSelectLatest(j.id)}
                          aria-label={`Pilih jurnal ${j.topic}`}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-600"
                        />
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap tabular-nums text-slate-900 font-medium">
                        {formatShortDateId(j.date)}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap tabular-nums text-slate-600">
                        {j.jamKe}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap font-semibold text-slate-900">
                        {j.className}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-slate-700">
                        {j.subject}
                      </td>
                      <td className="px-4 py-3 text-slate-700">{j.topic}</td>
                      <td className="px-4 py-3 whitespace-nowrap text-right tabular-nums text-xs text-slate-600">
                        H:{sum.counts.H} · S:{sum.counts.S} · I:{sum.counts.I} · A:{sum.counts.A} · D:{sum.counts.D}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap text-right">
                        <div className="inline-flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => onEditJournal(j)}
                            className="px-2.5 py-1.5 rounded-md border border-slate-200 text-xs font-medium text-slate-700 hover:bg-slate-100 transition-colors inline-flex items-center gap-1"
                            title="Edit jurnal ini"
                          >
                            <Pencil className="w-3.5 h-3.5" />
                            Edit
                          </button>
                          <button
                            type="button"
                            onClick={() => onDuplicateJournal(j)}
                            className="px-2.5 py-1.5 rounded-md border border-indigo-200 text-xs font-medium text-indigo-700 bg-indigo-50/60 hover:bg-indigo-100 transition-colors inline-flex items-center gap-1"
                            title="Duplikat jurnal ini"
                          >
                            <Copy className="w-3.5 h-3.5" />
                            Duplikat
                          </button>
                          <button
                            type="button"
                            onClick={() => handleSingleDeleteLatest(j)}
                            className="px-2.5 py-1.5 rounded-md border border-rose-200 text-xs font-medium text-rose-700 hover:bg-rose-50 transition-colors inline-flex items-center gap-1"
                            title="Hapus jurnal ini"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            Hapus
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}
