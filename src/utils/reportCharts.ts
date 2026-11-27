import Chart, { type ChartConfiguration, type ChartType, type Plugin } from 'chart.js/auto';
import {
  AttendanceStatus,
  GradeEntry,
  JournalEntry,
  PrintFontFamily,
  ReportChartId,
  Student,
} from '../types';
import {
  compareClassNames,
  formatCompactDateId,
  summarizeJournalRow,
} from './portalHelpers';

export interface ReportChartOptionMeta {
  id: ReportChartId;
  label: string;
  chartKind: string;
  description: string;
}

export const REPORT_CHART_OPTIONS: ReportChartOptionMeta[] = [
  {
    id: 'bar_class',
    label: 'Diagram Batang',
    chartKind: 'Batang Vertikal',
    description: 'Perbandingan % Kehadiran & Rata-rata Nilai per Kelas',
  },
  {
    id: 'line_trend',
    label: 'Diagram Garis',
    chartKind: 'Garis Tren',
    description: 'Tren % Kehadiran & Jumlah Pertemuan KBM per Tanggal',
  },
  {
    id: 'doughnut_attendance',
    label: 'Diagram Donat',
    chartKind: 'Donat / Cincin',
    description: 'Proporsi Status Presensi Siswa (H / S / I / A / D)',
  },
  {
    id: 'hbar_subject',
    label: 'Diagram Batang Horizontal',
    chartKind: 'Batang Mendatar',
    description: 'Peringkat Rata-rata Nilai per Mata Pelajaran',
  },
  {
    id: 'pie_kkm',
    label: 'Diagram Lingkaran',
    chartKind: 'Lingkaran (Pie)',
    description: 'Distribusi Predikat & Ketuntasan KKM (>= 75 vs < 75)',
  },
];

export interface GeneratedReportChartCard {
  id: ReportChartId;
  title: string;
  subtitle: string;
  chartKind: string;
  pngDataUrl: string;
  widthPx: number;
  heightPx: number;
  summaryRows: (string | number)[][];
}

function getCanvasFontFamily(fontFamily: PrintFontFamily): string {
  switch (fontFamily) {
    case 'Tahoma':
      return 'Tahoma, sans-serif';
    case 'Arial':
      return 'Arial, Helvetica, sans-serif';
    case 'Verdana':
      return 'Verdana, sans-serif';
    case 'Calibri':
      return 'Calibri, Arial, sans-serif';
    case 'Georgia':
      return 'Georgia, serif';
    case 'Times New Roman':
    default:
      return '"Times New Roman", Times, serif';
  }
}

function createWhiteBgPlugin<T extends ChartType = ChartType>(): Plugin<T> {
  return {
    id: 'reportWhiteBg',
    beforeDraw: (chart) => {
      const { ctx, width, height } = chart;
      ctx.save();
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, width, height);
      ctx.restore();
    },
  };
}

function createDataLabelPlugin<T extends ChartType = ChartType>(
  mode: 'bar' | 'hbar' | 'line' | 'pie',
  fontFamilyCss: string,
  suffix = ''
): Plugin<T> {
  return {
    id: 'reportDataLabels',
    afterDatasetsDraw: (chart) => {
      const { ctx } = chart;
      ctx.save();
      ctx.font = `bold 19px ${fontFamilyCss}`;
      ctx.textBaseline = 'middle';

      chart.data.datasets.forEach((dataset, dsIndex) => {
        const meta = chart.getDatasetMeta(dsIndex);
        if (meta.hidden) return;
        const rawData = dataset.data as (number | null)[];
        const total =
          mode === 'pie'
            ? rawData.reduce<number>((sum, val) => sum + (typeof val === 'number' ? val : 0), 0)
            : 0;

        meta.data.forEach((element, idx) => {
          const val = rawData[idx];
          if (val === null || val === undefined || Number.isNaN(val)) return;
          if (mode === 'pie' && val <= 0) return;

          const pos = element.tooltipPosition(false);
          const px = pos.x;
          const py = pos.y;
          if (px === null || py === null || Number.isNaN(px) || Number.isNaN(py)) return;
          const formatted =
            Number.isInteger(val) ? String(val) : Number(val).toFixed(1);

          if (mode === 'pie') {
            const pct = total > 0 ? Math.round((val / total) * 100) : 0;
            if (pct < 4) return;
            const label = `${val} (${pct}%)`;
            ctx.textAlign = 'center';
            ctx.fillStyle = '#ffffff';
            ctx.strokeStyle = 'rgba(15, 23, 42, 0.72)';
            ctx.lineWidth = 4;
            ctx.strokeText(label, px, py);
            ctx.fillText(label, px, py);
          } else if (mode === 'hbar') {
            ctx.textAlign = 'left';
            ctx.fillStyle = '#0f172a';
            ctx.fillText(`${formatted}${suffix}`, px + 8, py);
          } else if (mode === 'line') {
            ctx.textAlign = 'center';
            ctx.fillStyle = dsIndex === 0 ? '#1d4ed8' : '#047857';
            const yOffset = dsIndex === 0 ? -16 : 18;
            ctx.fillText(
              dsIndex === 0 ? `${formatted}%` : String(formatted),
              px,
              Math.max(24, py + yOffset)
            );
          } else {
            ctx.textAlign = 'center';
            ctx.fillStyle = '#0f172a';
            ctx.fillText(`${formatted}${suffix}`, px, Math.max(22, py - 12));
          }
        });
      });
      ctx.restore();
    },
  };
}

function renderOffscreenChartPng<T extends ChartType>(
  config: ChartConfiguration<T>,
  widthPx = 960,
  heightPx = 500
): string {
  if (typeof document === 'undefined') return '';
  const canvas = document.createElement('canvas');
  canvas.width = widthPx;
  canvas.height = heightPx;

  const chart = new Chart(canvas, {
    ...config,
    options: {
      ...config.options,
      responsive: false,
      animation: false,
      devicePixelRatio: 1,
    },
  } as ChartConfiguration<T>);
  const url = canvas.toDataURL('image/png');
  chart.destroy();
  return url;
}

export function buildReportCharts(params: {
  selectedChartTypes: ReportChartId[];
  selectedClasses: string[];
  subjectFilter: string;
  students: Student[];
  journals: JournalEntry[];
  grades: GradeEntry[];
  academicYear: string;
  semester: string;
  fontFamily: PrintFontFamily;
}): GeneratedReportChartCard[] {
  const {
    selectedChartTypes,
    selectedClasses,
    subjectFilter,
    students,
    journals,
    grades,
    academicYear,
    semester,
    fontFamily,
  } = params;

  if (selectedChartTypes.length === 0) return [];

  const fontCss = getCanvasFontFamily(fontFamily);
  const sortedClasses = [...selectedClasses].sort(compareClassNames);
  const clsSet = new Set(sortedClasses);

  const periodGrades = grades.filter(
    (g) =>
      g.academicYear === academicYear &&
      g.semester === semester &&
      g.score !== null &&
      clsSet.has(g.className) &&
      (subjectFilter === 'ALL' || g.subject === subjectFilter)
  );

  const widthPx = 960;
  const heightPx = 500;

  const cards: GeneratedReportChartCard[] = [];

  for (const chartId of selectedChartTypes) {
    if (chartId === 'bar_class') {
      const labels = sortedClasses.length > 0 ? sortedClasses : ['Belum Ada Kelas'];
      const attendancePcts: number[] = [];
      const avgScores: number[] = [];
      const summaryRows: (string | number)[][] = [];

      labels.forEach((cls) => {
        const clsJournals = journals.filter((j) => j.className === cls);
        let hCount = 0;
        let recCount = 0;
        clsJournals.forEach((j) => {
          const sum = summarizeJournalRow(j);
          hCount += sum.counts.H;
          recCount += sum.total;
        });
        const attPct = recCount > 0 ? Number(((hCount / recCount) * 100).toFixed(1)) : 0;

        const clsGrades = periodGrades.filter((g) => g.className === cls);
        const avg =
          clsGrades.length > 0
            ? Number(
                (
                  clsGrades.reduce((acc, g) => acc + (g.score ?? 0), 0) / clsGrades.length
                ).toFixed(1)
              )
            : 0;

        attendancePcts.push(attPct);
        avgScores.push(avg);
        summaryRows.push(['Diagram Batang - Per Kelas', cls, `${attPct}% Hadir`, `${avg} Nilai`]);
      });

      const pngDataUrl = renderOffscreenChartPng(
        {
          type: 'bar',
          data: {
            labels,
            datasets: [
              {
                label: '% Kehadiran Siswa',
                data: attendancePcts,
                backgroundColor: '#10b981',
                borderColor: '#047857',
                borderWidth: 1.5,
                borderRadius: 6,
              },
              {
                label: 'Rata-rata Nilai',
                data: avgScores,
                backgroundColor: '#4f46e5',
                borderColor: '#3730a3',
                borderWidth: 1.5,
                borderRadius: 6,
              },
            ],
          },
          options: {
            layout: { padding: { top: 28, right: 20, bottom: 12, left: 16 } },
            scales: {
              y: {
                beginAtZero: true,
                max: 112,
                ticks: { stepSize: 20, font: { family: fontCss, size: 18 }, color: '#1e293b' },
                grid: { color: '#e2e8f0' },
              },
              x: {
                ticks: { font: { family: fontCss, size: 18, weight: 'bold' }, color: '#0f172a' },
                grid: { display: false },
              },
            },
            plugins: {
              legend: {
                position: 'bottom',
                labels: { font: { family: fontCss, size: 19, weight: 'bold' }, color: '#0f172a', padding: 18 },
              },
            },
          },
          plugins: [createWhiteBgPlugin<'bar'>(), createDataLabelPlugin<'bar'>('bar', fontCss)],
        },
        widthPx,
        heightPx
      );

      cards.push({
        id: 'bar_class',
        title: 'Perbandingan Kehadiran (%) & Rata-rata Nilai per Kelas',
        subtitle: `Kelas terpilih: ${sortedClasses.length} kelas`,
        chartKind: 'Diagram Batang',
        pngDataUrl,
        widthPx,
        heightPx,
        summaryRows,
      });
    } else if (chartId === 'line_trend') {
      const dateMap = new Map<string, { h: number; total: number; meetings: number }>();
      journals.forEach((j) => {
        if (!j.date) return;
        const cur = dateMap.get(j.date) || { h: 0, total: 0, meetings: 0 };
        const sum = summarizeJournalRow(j);
        cur.h += sum.counts.H;
        cur.total += sum.total;
        cur.meetings += 1;
        dateMap.set(j.date, cur);
      });

      const sortedDates = Array.from(dateMap.keys()).sort().slice(-12);
      const labels =
        sortedDates.length > 0
          ? sortedDates.map((d) => formatCompactDateId(d))
          : ['Belum Ada KBM'];
      const pctSeries =
        sortedDates.length > 0
          ? sortedDates.map((d) => {
              const item = dateMap.get(d)!;
              return item.total > 0 ? Number(((item.h / item.total) * 100).toFixed(1)) : 100;
            })
          : [0];
      const meetingSeries =
        sortedDates.length > 0
          ? sortedDates.map((d) => dateMap.get(d)!.meetings)
          : [0];

      const summaryRows: (string | number)[][] =
        sortedDates.length > 0
          ? sortedDates.map((d, idx) => [
              'Diagram Garis - Tren KBM',
              d,
              `${pctSeries[idx]}% Hadir`,
              `${meetingSeries[idx]} Pertemuan`,
            ])
          : [['Diagram Garis - Tren KBM', 'Belum ada data', '0%', '0']];

      const pngDataUrl = renderOffscreenChartPng(
        {
          type: 'line',
          data: {
            labels,
            datasets: [
              {
                label: '% Kehadiran',
                data: pctSeries,
                borderColor: '#2563eb',
                backgroundColor: 'rgba(37, 99, 235, 0.14)',
                borderWidth: 3.5,
                pointRadius: 6,
                pointBackgroundColor: '#1d4ed8',
                tension: 0.28,
                fill: true,
                yAxisID: 'y',
              },
              {
                label: 'Jumlah Pertemuan',
                data: meetingSeries,
                borderColor: '#059669',
                backgroundColor: '#059669',
                borderDash: [7, 5],
                borderWidth: 3,
                pointRadius: 5,
                tension: 0.2,
                yAxisID: 'y1',
              },
            ],
          },
          options: {
            layout: { padding: { top: 28, right: 20, bottom: 12, left: 16 } },
            scales: {
              y: {
                type: 'linear',
                position: 'left',
                beginAtZero: true,
                max: 115,
                ticks: { font: { family: fontCss, size: 17 }, color: '#1d4ed8' },
                grid: { color: '#e2e8f0' },
              },
              y1: {
                type: 'linear',
                position: 'right',
                beginAtZero: true,
                suggestedMax: Math.max(5, ...meetingSeries) + 2,
                ticks: { stepSize: 1, font: { family: fontCss, size: 17 }, color: '#047857' },
                grid: { drawOnChartArea: false },
              },
              x: {
                ticks: { font: { family: fontCss, size: 16, weight: 'bold' }, color: '#0f172a' },
                grid: { display: false },
              },
            },
            plugins: {
              legend: {
                position: 'bottom',
                labels: { font: { family: fontCss, size: 19, weight: 'bold' }, color: '#0f172a', padding: 18 },
              },
            },
          },
          plugins: [createWhiteBgPlugin<'line'>(), createDataLabelPlugin<'line'>('line', fontCss)],
        },
        widthPx,
        heightPx
      );

      cards.push({
        id: 'line_trend',
        title: 'Tren Persentase Kehadiran & Pertemuan KBM',
        subtitle:
          sortedDates.length > 0
            ? `${sortedDates.length} tanggal pertemuan terakhir pada filter aktif`
            : 'Belum ada jurnal pada rentang filter ini',
        chartKind: 'Diagram Garis',
        pngDataUrl,
        widthPx,
        heightPx,
        summaryRows,
      });
    } else if (chartId === 'doughnut_attendance') {
      const counts: Record<AttendanceStatus, number> = { H: 0, S: 0, I: 0, A: 0, D: 0 };
      journals.forEach((j) => {
        const s = summarizeJournalRow(j);
        counts.H += s.counts.H;
        counts.S += s.counts.S;
        counts.I += s.counts.I;
        counts.A += s.counts.A;
        counts.D += s.counts.D;
      });
      const totalRec = counts.H + counts.S + counts.I + counts.A + counts.D;
      const hasData = totalRec > 0;

      const summaryRows: (string | number)[][] = [
        ['Diagram Donat - Presensi', 'Hadir (H)', counts.H, hasData ? `${((counts.H / totalRec) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Donat - Presensi', 'Sakit (S)', counts.S, hasData ? `${((counts.S / totalRec) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Donat - Presensi', 'Izin (I)', counts.I, hasData ? `${((counts.I / totalRec) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Donat - Presensi', 'Alpha (A)', counts.A, hasData ? `${((counts.A / totalRec) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Donat - Presensi', 'Dispensasi (D)', counts.D, hasData ? `${((counts.D / totalRec) * 100).toFixed(1)}%` : '0%'],
      ];

      const pngDataUrl = renderOffscreenChartPng(
        {
          type: 'doughnut',
          data: {
            labels: hasData
              ? ['Hadir (H)', 'Sakit (S)', 'Izin (I)', 'Alpha (A)', 'Dispensasi (D)']
              : ['Belum Ada Data Presensi'],
            datasets: [
              {
                data: hasData ? [counts.H, counts.S, counts.I, counts.A, counts.D] : [1],
                backgroundColor: hasData
                  ? ['#10b981', '#f59e0b', '#0ea5e9', '#f43f5e', '#8b5cf6']
                  : ['#cbd5e1'],
                borderColor: '#ffffff',
                borderWidth: 3,
              },
            ],
          },
          options: {
            cutout: '54%',
            layout: { padding: { top: 18, right: 20, bottom: 12, left: 20 } },
            plugins: {
              legend: {
                position: 'right',
                labels: {
                  font: { family: fontCss, size: 19, weight: 'bold' },
                  color: '#0f172a',
                  padding: 16,
                  boxWidth: 22,
                },
              },
            },
          },
          plugins: hasData
            ? [createWhiteBgPlugin<'doughnut'>(), createDataLabelPlugin<'doughnut'>('pie', fontCss)]
            : [createWhiteBgPlugin<'doughnut'>()],
        },
        widthPx,
        heightPx
      );

      cards.push({
        id: 'doughnut_attendance',
        title: 'Komposisi Status Presensi Siswa (H/S/I/A/D)',
        subtitle: `Total ${totalRec.toLocaleString('id-ID')} entri presensi pada ${journals.length} pertemuan`,
        chartKind: 'Diagram Donat',
        pngDataUrl,
        widthPx,
        heightPx,
        summaryRows,
      });
    } else if (chartId === 'hbar_subject') {
      const subjMap = new Map<string, { sum: number; count: number }>();
      periodGrades.forEach((g) => {
        if (g.score === null) return;
        const cur = subjMap.get(g.subject) || { sum: 0, count: 0 };
        cur.sum += g.score;
        cur.count += 1;
        subjMap.set(g.subject, cur);
      });

      const entries = Array.from(subjMap.entries())
        .map(([subj, val]) => ({
          subj,
          avg: Number((val.sum / val.count).toFixed(1)),
          count: val.count,
        }))
        .sort((a, b) => b.avg - a.avg)
        .slice(0, 10);

      const labels = entries.length > 0 ? entries.map((e) => e.subj) : ['Belum Ada Nilai'];
      const avgs = entries.length > 0 ? entries.map((e) => e.avg) : [0];
      const colors = avgs.map((v) =>
        v >= 85 ? '#059669' : v >= 75 ? '#2563eb' : v > 0 ? '#f59e0b' : '#cbd5e1'
      );

      const summaryRows: (string | number)[][] =
        entries.length > 0
          ? entries.map((e) => [
              'Batang Horizontal - Mapel',
              e.subj,
              `Rata-rata ${e.avg}`,
              `${e.count} Skor`,
            ])
          : [['Batang Horizontal - Mapel', 'Belum ada nilai', 0, 0]];

      const pngDataUrl = renderOffscreenChartPng(
        {
          type: 'bar',
          data: {
            labels,
            datasets: [
              {
                label: 'Rata-rata Nilai Mapel',
                data: avgs,
                backgroundColor: colors,
                borderRadius: 6,
                barThickness: entries.length <= 4 ? 42 : 28,
              },
            ],
          },
          options: {
            indexAxis: 'y',
            layout: { padding: { top: 18, right: 64, bottom: 12, left: 16 } },
            scales: {
              x: {
                beginAtZero: true,
                max: 105,
                ticks: { stepSize: 20, font: { family: fontCss, size: 17 }, color: '#1e293b' },
                grid: { color: '#e2e8f0' },
              },
              y: {
                ticks: { font: { family: fontCss, size: 17, weight: 'bold' }, color: '#0f172a' },
                grid: { display: false },
              },
            },
            plugins: {
              legend: {
                position: 'bottom',
                labels: { font: { family: fontCss, size: 19, weight: 'bold' }, color: '#0f172a', padding: 14 },
              },
            },
          },
          plugins: [createWhiteBgPlugin<'bar'>(), createDataLabelPlugin<'bar'>('hbar', fontCss)],
        },
        widthPx,
        heightPx
      );

      cards.push({
        id: 'hbar_subject',
        title: 'Capaian Rata-rata Nilai per Mata Pelajaran',
        subtitle: `Semester ${semester} ${academicYear} · Standar KKM 75`,
        chartKind: 'Diagram Batang Horizontal',
        pngDataUrl,
        widthPx,
        heightPx,
        summaryRows,
      });
    } else if (chartId === 'pie_kkm') {
      // Aggregate average score per student across filtered grades
      const activeStudents = students.filter((s) => clsSet.has(s.className));
      const studentScores = new Map<string, number[]>();
      periodGrades.forEach((g) => {
        if (g.score === null) return;
        const arr = studentScores.get(g.studentId) || [];
        arr.push(g.score);
        studentScores.set(g.studentId, arr);
      });

      let sangatBaik = 0; // >= 88
      let tuntas = 0; // 75 - 87.99
      let perluBimbingan = 0; // 60 - 74.99
      let remedial = 0; // < 60

      activeStudents.forEach((st) => {
        const arr = studentScores.get(st.id);
        if (!arr || arr.length === 0) return;
        const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
        if (avg >= 88) sangatBaik += 1;
        else if (avg >= 75) tuntas += 1;
        else if (avg >= 60) perluBimbingan += 1;
        else remedial += 1;
      });

      const totalAssessed = sangatBaik + tuntas + perluBimbingan + remedial;
      const hasData = totalAssessed > 0;

      const summaryRows: (string | number)[][] = [
        ['Diagram Lingkaran - KKM', 'Sangat Baik (>= 88)', sangatBaik, hasData ? `${((sangatBaik / totalAssessed) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Lingkaran - KKM', 'Tuntas KKM (75 - 87)', tuntas, hasData ? `${((tuntas / totalAssessed) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Lingkaran - KKM', 'Mendekati KKM (60 - 74)', perluBimbingan, hasData ? `${((perluBimbingan / totalAssessed) * 100).toFixed(1)}%` : '0%'],
        ['Diagram Lingkaran - KKM', 'Remedial (< 60)', remedial, hasData ? `${((remedial / totalAssessed) * 100).toFixed(1)}%` : '0%'],
      ];

      const pngDataUrl = renderOffscreenChartPng(
        {
          type: 'pie',
          data: {
            labels: hasData
              ? [
                  'Sangat Baik (>=88)',
                  'Tuntas KKM (75-87)',
                  'Perlu Penguatan (60-74)',
                  'Remedial (<60)',
                ]
              : ['Belum Ada Data Nilai Siswa'],
            datasets: [
              {
                data: hasData ? [sangatBaik, tuntas, perluBimbingan, remedial] : [1],
                backgroundColor: hasData
                  ? ['#059669', '#3b82f6', '#f59e0b', '#ef4444']
                  : ['#cbd5e1'],
                borderColor: '#ffffff',
                borderWidth: 3,
              },
            ],
          },
          options: {
            layout: { padding: { top: 18, right: 20, bottom: 12, left: 20 } },
            plugins: {
              legend: {
                position: 'right',
                labels: {
                  font: { family: fontCss, size: 18, weight: 'bold' },
                  color: '#0f172a',
                  padding: 16,
                  boxWidth: 22,
                },
              },
            },
          },
          plugins: hasData
            ? [createWhiteBgPlugin<'pie'>(), createDataLabelPlugin<'pie'>('pie', fontCss)]
            : [createWhiteBgPlugin<'pie'>()],
        },
        widthPx,
        heightPx
      );

      cards.push({
        id: 'pie_kkm',
        title: 'Distribusi Predikat & Ketuntasan KKM Siswa',
        subtitle: `${totalAssessed} siswa ternilai · Tuntas (>=75): ${sangatBaik + tuntas} siswa`,
        chartKind: 'Diagram Lingkaran',
        pngDataUrl,
        widthPx,
        heightPx,
        summaryRows,
      });
    }
  }

  return cards;
}
