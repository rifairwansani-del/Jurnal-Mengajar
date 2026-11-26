import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  PaperOrientation,
  PaperSizeId,
  PortalDatabase,
  PrintFontFamily,
  SignaturePositionSettings,
  UserAccount,
} from '../types';
import {
  formatLongDateId,
  getTodayISO,
  imageUrlToPngDataUrl,
} from './portalHelpers';

export const PAPER_SIZE_CONFIG: Record<
  PaperSizeId,
  {
    label: string;
    shortLabel: string;
    jsPdfFormat: string | [number, number];
    cssPageSize: string;
    portraitMaxW: string;
    landscapeMinW: string;
    landscapeMaxW: string;
  }
> = {
  A4: {
    label: 'A4 (210 × 297 mm)',
    shortLabel: 'A4',
    jsPdfFormat: 'a4',
    cssPageSize: 'A4',
    portraitMaxW: 'max-w-[794px]',
    landscapeMinW: 'min-w-[920px]',
    landscapeMaxW: 'max-w-[1123px]',
  },
  F4: {
    label: 'F4 / Folio (215 × 330 mm)',
    shortLabel: 'F4',
    jsPdfFormat: [215, 330],
    cssPageSize: '215mm 330mm',
    portraitMaxW: 'max-w-[812px]',
    landscapeMinW: 'min-w-[1000px]',
    landscapeMaxW: 'max-w-[1247px]',
  },
  Letter: {
    label: 'Letter (216 × 279 mm)',
    shortLabel: 'Letter',
    jsPdfFormat: 'letter',
    cssPageSize: 'letter',
    portraitMaxW: 'max-w-[816px]',
    landscapeMinW: 'min-w-[900px]',
    landscapeMaxW: 'max-w-[1056px]',
  },
  Legal: {
    label: 'Legal (216 × 356 mm)',
    shortLabel: 'Legal',
    jsPdfFormat: 'legal',
    cssPageSize: 'legal',
    portraitMaxW: 'max-w-[816px]',
    landscapeMinW: 'min-w-[1040px]',
    landscapeMaxW: 'max-w-[1344px]',
  },
  A5: {
    label: 'A5 (148 × 210 mm)',
    shortLabel: 'A5',
    jsPdfFormat: 'a5',
    cssPageSize: 'A5',
    portraitMaxW: 'max-w-[559px]',
    landscapeMinW: 'min-w-[680px]',
    landscapeMaxW: 'max-w-[794px]',
  },
};

export interface ActiveSignerInfo {
  name: string;
  nip: string;
  roleLabel: string;
  signatureImage: string;
  schoolStampImage: string;
}

export function resolveActiveSigner(
  db: PortalDatabase,
  overrideAccountId?: string
): ActiveSignerInfo {
  const { accounts, settings, session } = db;
  let account: UserAccount | undefined;
  if (overrideAccountId) {
    account = accounts.find((a) => a.id === overrideAccountId);
  }
  if (!account && session.accountId) {
    account = accounts.find((a) => a.id === session.accountId);
  }
  if (!account && session.email) {
    account = accounts.find(
      (a) => a.email.trim().toLowerCase() === session.email.trim().toLowerCase()
    );
  }

  return {
    name: account?.name || session.userName || settings.teacherName,
    nip: account?.nip || settings.teacherNip,
    roleLabel: account?.role && account.role !== 'Admin' ? account.role : 'Guru Mata Pelajaran',
    signatureImage: account?.signatureImage || '',
    schoolStampImage: account?.schoolStampImage || settings.defaultSchoolStampImage || '',
  };
}

export function exportRowsToXlsx(
  sheetName: string,
  headers: string[],
  rows: (string | number)[][],
  filename: string
) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`);
}

export function exportRowsToXls(
  sheetName: string,
  headers: string[],
  rows: (string | number)[][],
  filename: string
) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, filename.endsWith('.xls') ? filename : `${filename}.xls`, {
    bookType: 'biff8',
  });
}

function escapeHtml(value: string | number): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildWordFreeSigTableHtml(
  items: { xPct: number; yPx: number; html: string }[]
): string {
  const valid = items.filter((it) => it.html.trim() !== '');
  if (valid.length === 0) return '';
  if (valid.length === 1) {
    const only = valid[0];
    const leftPct = Math.max(0, Math.min(72, only.xPct));
    return `
      <table style="width:100%; margin-top:26px; border-collapse:collapse;">
        <tr>
          ${leftPct > 0 ? `<td style="width:${leftPct}%; border:none; padding:0;"></td>` : ''}
          <td style="vertical-align:top; border:none; padding:${only.yPx}px 0 0 0;">${only.html}</td>
        </tr>
      </table>
    `;
  }

  const sorted = [...valid].sort((a, b) => a.xPct - b.xPct);
  const first = sorted[0];
  const second = sorted[1];
  const verticalOverlap = Math.abs(second.xPct - first.xPct) < 22 && Math.abs(second.yPx - first.yPx) >= 50;

  if (verticalOverlap) {
    const byY = [...valid].sort((a, b) => a.yPx - b.yPx);
    return byY
      .map(
        (it, idx) => `
      <table style="width:100%; margin-top:${idx === 0 ? 26 : 16}px; border-collapse:collapse;">
        <tr>
          ${it.xPct > 0 ? `<td style="width:${Math.min(72, it.xPct)}%; border:none; padding:0;"></td>` : ''}
          <td style="vertical-align:top; border:none; padding:0;">${it.html}</td>
        </tr>
      </table>`
      )
      .join('');
  }

  const leftSpacer = Math.max(0, Math.min(68, first.xPct));
  const secondX = Math.max(leftSpacer + 24, Math.min(76, second.xPct));
  const firstWidth = Math.max(24, Math.min(34, secondX - leftSpacer));
  const gapSpacer = Math.max(0, secondX - leftSpacer - firstWidth);

  return `
    <table style="width:100%; margin-top:26px; border-collapse:collapse;">
      <tr>
        ${leftSpacer > 0 ? `<td style="width:${leftSpacer}%; border:none; padding:0;"></td>` : ''}
        <td style="width:${firstWidth}%; vertical-align:top; border:none; padding:${first.yPx}px 0 0 0;">${first.html}</td>
        ${gapSpacer > 0 ? `<td style="width:${gapSpacer}%; border:none; padding:0;"></td>` : ''}
        <td style="vertical-align:top; border:none; padding:${second.yPx}px 0 0 0;">${second.html}</td>
      </tr>
    </table>
  `;
}

async function buildOfficialPrintHtml(options: {
  db: PortalDatabase;
  title: string;
  subtitle?: string;
  headers: string[];
  rows: (string | number)[][];
  orientation?: PaperOrientation;
  paperSize?: PaperSizeId;
  cssPageSize?: string;
  fontFamily?: PrintFontFamily;
  showKop?: boolean;
  showTtdKepsek?: boolean;
  showTtdGuru?: boolean;
  overrideAccountId?: string;
  forWord?: boolean;
}): Promise<string> {
  const {
    db,
    title,
    subtitle = '',
    headers,
    rows,
    orientation = 'landscape',
    paperSize = 'A4',
    cssPageSize,
    fontFamily = 'Times New Roman',
    showKop = true,
    showTtdKepsek = true,
    showTtdGuru = true,
    overrideAccountId,
    forWord = false,
  } = options;

  const resolvedCssPageSize = cssPageSize || PAPER_SIZE_CONFIG[paperSize]?.cssPageSize || 'A4';
  const { settings } = db;
  const signer = resolveActiveSigner(db, overrideAccountId);
  const sigPos: SignaturePositionSettings = settings.signaturePosition;

  const showSigImage = settings.reportPrintPreferences?.showSignatureImage ?? true;

  const [leftLogoPng, rightLogoPng, kepsekSigPng, guruSigPng, stampPng] =
    await Promise.all([
      showKop ? imageUrlToPngDataUrl(settings.kopSurat.logoLeft) : Promise.resolve(null),
      showKop ? imageUrlToPngDataUrl(settings.kopSurat.logoRight) : Promise.resolve(null),
      showTtdKepsek && showSigImage
        ? imageUrlToPngDataUrl(settings.principalSignatureImage)
        : Promise.resolve(null),
      showTtdGuru && showSigImage
        ? imageUrlToPngDataUrl(signer.signatureImage)
        : Promise.resolve(null),
      sigPos.showStamp ? imageUrlToPngDataUrl(signer.schoolStampImage) : Promise.resolve(null),
    ]);

  const kopHtml = showKop
    ? `
      <table style="width:100%; border-collapse:collapse; border-bottom:4px double #000; margin-bottom:16px; padding-bottom:8px;">
        <tr>
          <td style="width:80px; text-align:center; vertical-align:middle; padding-bottom:8px;">
            ${leftLogoPng ? `<img src="${leftLogoPng}" style="width:68px; height:68px; object-fit:contain;" />` : ''}
          </td>
          <td style="text-align:center; vertical-align:middle; line-height:1.25; padding-bottom:8px;">
            ${settings.kopSurat.line1 ? `<div style="font-size:11pt; text-transform:uppercase;">${escapeHtml(settings.kopSurat.line1)}</div>` : ''}
            ${settings.kopSurat.line2 ? `<div style="font-size:12pt; font-weight:bold; text-transform:uppercase;">${escapeHtml(settings.kopSurat.line2)}</div>` : ''}
            ${settings.kopSurat.line3 ? `<div style="font-size:14pt; font-weight:bold; text-transform:uppercase;">${escapeHtml(settings.kopSurat.line3)}</div>` : ''}
            ${settings.kopSurat.line4 ? `<div style="font-size:9pt; margin-top:2px;">${escapeHtml(settings.kopSurat.line4)}</div>` : ''}
            ${settings.kopSurat.line5 ? `<div style="font-size:9pt;">${escapeHtml(settings.kopSurat.line5)}</div>` : ''}
            ${settings.kopSurat.line6 ? `<div style="font-size:9pt;">${escapeHtml(settings.kopSurat.line6)}</div>` : ''}
          </td>
          <td style="width:80px; text-align:center; vertical-align:middle; padding-bottom:8px;">
            ${rightLogoPng ? `<img src="${rightLogoPng}" style="width:68px; height:68px; object-fit:contain;" />` : ''}
          </td>
        </tr>
      </table>
    `
    : '';

  const kepsekSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const guruSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const stampSizePx = Math.max(20, Math.min(260, sigPos.stampSize ?? 70));
  const kepsekSigOffsetX = sigPos.principalSigOffsetX ?? 0;
  const kepsekSigOffsetY = sigPos.principalSigOffsetY ?? 0;
  const guruSigOffsetX = sigPos.teacherSigOffsetX ?? 0;
  const guruSigOffsetY = sigPos.teacherSigOffsetY ?? 0;
  const stampOffsetX = sigPos.stampOffsetX ?? 0;
  const stampOffsetY = sigPos.stampOffsetY ?? 0;
  const stampLeftPx = 17 - Math.round(stampSizePx / 2) + stampOffsetX;
  const stampTopPx = Math.round((68 - stampSizePx) / 2) + stampOffsetY;

  const renderSigBox = (
    topLabel: string,
    roleTitle: string,
    personName: string,
    personNip: string,
    sigUrl: string | null,
    withStamp: boolean,
    sigPx: number,
    sigOffsetX: number,
    sigOffsetY: number
  ) => `
    <div style="font-size:10pt; line-height:1.35;">
      <div>${escapeHtml(topLabel)}</div>
      <div>${escapeHtml(roleTitle)}</div>
      <div style="position:relative; height:68px; margin:4px 0; overflow:visible;">
        ${
          withStamp && stampPng
            ? `<img src="${stampPng}" style="position:absolute; left:${stampLeftPx}px; top:${stampTopPx}px; width:${stampSizePx}px; height:${stampSizePx}px; object-fit:contain; opacity:0.82;" />`
            : ''
        }
        ${
          sigUrl
            ? `<img src="${sigUrl}" style="position:relative; z-index:2; left:${sigOffsetX}px; top:${sigOffsetY}px; height:${sigPx}px; max-width:${Math.round(sigPx * 2.85)}px; object-fit:contain; margin-top:${Math.round((68 - sigPx) / 2)}px;" />`
            : ''
        }
      </div>
      <div style="font-weight:bold; text-decoration:underline;">${escapeHtml(personName)}</div>
      <div>NIP. ${escapeHtml(personNip || '-')}</div>
    </div>
  `;

  const stampOnKepsek =
    sigPos.showStamp && (sigPos.stampTarget === 'kepsek' || sigPos.stampTarget === 'keduanya');
  const stampOnGuru =
    sigPos.showStamp && (sigPos.stampTarget === 'guru' || sigPos.stampTarget === 'keduanya');

  const kepsekBox = showTtdKepsek
    ? renderSigBox(
        'Mengetahui,',
        settings.principalTitle,
        settings.principalName,
        settings.principalNip,
        kepsekSigPng,
        stampOnKepsek,
        kepsekSigSizePx,
        kepsekSigOffsetX,
        kepsekSigOffsetY
      )
    : '';

  const guruBox = showTtdGuru
    ? renderSigBox(
        `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
        signer.roleLabel,
        signer.name,
        signer.nip,
        guruSigPng,
        stampOnGuru,
        guruSigSizePx,
        guruSigOffsetX,
        guruSigOffsetY
      )
    : '';

  let sigBlockHtml = '';
  if (showTtdKepsek || showTtdGuru) {
    if (sigPos.preset === 'kanan-kiri') {
      sigBlockHtml = `
        <table style="width:100%; margin-top:26px; border-collapse:collapse;">
          <tr>
            <td style="width:50%; vertical-align:top; padding-left:18px; border:none;">${guruBox}</td>
            <td style="width:50%; vertical-align:top; padding-left:18%; border:none;">${kepsekBox}</td>
          </tr>
        </table>
      `;
    } else if (sigPos.preset === 'kiri') {
      sigBlockHtml = `<div style="margin-top:26px; padding-left:18px;">${kepsekBox}<div style="margin-top:18px;">${guruBox}</div></div>`;
    } else if (sigPos.preset === 'tengah') {
      sigBlockHtml = `
        <table style="width:100%; margin-top:26px; border-collapse:collapse; text-align:center;">
          <tr>
            ${showTtdKepsek ? `<td style="vertical-align:top; text-align:center; border:none;">${kepsekBox}</td>` : ''}
            ${showTtdGuru ? `<td style="vertical-align:top; text-align:center; border:none;">${guruBox}</td>` : ''}
          </tr>
        </table>
      `;
    } else if (sigPos.preset === 'kanan') {
      sigBlockHtml = `<div style="margin-top:26px; margin-left:auto; width:260px;">${kepsekBox}<div style="margin-top:18px;">${guruBox}</div></div>`;
    } else if (sigPos.preset === 'bebas') {
      if (forWord) {
        sigBlockHtml = buildWordFreeSigTableHtml([
          showTtdKepsek
            ? { xPct: sigPos.principalX, yPx: sigPos.principalY, html: kepsekBox }
            : { xPct: 0, yPx: 0, html: '' },
          showTtdGuru
            ? { xPct: sigPos.teacherX, yPx: sigPos.teacherY, html: guruBox }
            : { xPct: 0, yPx: 0, html: '' },
        ]);
      } else {
        const minH = Math.max(185, Math.max(sigPos.principalY, sigPos.teacherY) + 150);
        sigBlockHtml = `
          <div style="position:relative; min-height:${minH}px; margin-top:24px;">
            ${
              showTtdKepsek
                ? `<div style="position:absolute; left:${sigPos.principalX}%; top:${sigPos.principalY}px; width:220px;">${kepsekBox}</div>`
                : ''
            }
            ${
              showTtdGuru
                ? `<div style="position:absolute; left:${sigPos.teacherX}%; top:${sigPos.teacherY}px; width:220px;">${guruBox}</div>`
                : ''
            }
          </div>
        `;
      }
    } else {
      sigBlockHtml = `
        <table style="width:100%; margin-top:26px; border-collapse:collapse;">
          <tr>
            <td style="width:50%; vertical-align:top; padding-left:18px; border:none;">${kepsekBox}</td>
            <td style="width:50%; vertical-align:top; padding-left:18%; border:none;">${guruBox}</td>
          </tr>
        </table>
      `;
    }
  }

  const thCells = headers
    .map(
      (h) =>
        `<th style="border:1px solid #000; background:#f1f5f9; padding:5px 7px; font-size:9pt; text-align:left; font-weight:bold;">${escapeHtml(
          h
        )}</th>`
    )
    .join('');

  const bodyRows =
    rows.length === 0
      ? `<tr><td colspan="${headers.length}" style="border:1px solid #000; padding:12px; text-align:center; font-size:9pt;">Tidak ada data</td></tr>`
      : rows
          .map(
            (r) =>
              `<tr>${r
                .map(
                  (cell) =>
                    `<td style="border:1px solid #000; padding:4px 7px; font-size:9pt; vertical-align:top;">${escapeHtml(
                      cell
                    )}</td>`
                )
                .join('')}</tr>`
          )
          .join('');

  return `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(title)}</title>
<style>
  @page { size: ${resolvedCssPageSize} ${orientation}; margin: 12mm; }
  body { font-family: "${fontFamily}", serif; color: #000; margin: 0; padding: 14px; }
  table { border-collapse: collapse; }
</style>
</head>
<body>
  ${kopHtml}
  <div style="text-align:center; margin-bottom:14px;">
    <div style="font-size:12pt; font-weight:bold; text-transform:uppercase; text-decoration:underline;">${escapeHtml(
      title
    )}</div>
    ${subtitle ? `<div style="font-size:9.5pt; color:#334155; margin-top:4px;">${escapeHtml(subtitle)}</div>` : ''}
  </div>
  <table style="width:100%; border-collapse:collapse;">
    <thead><tr>${thCells}</tr></thead>
    <tbody>${bodyRows}</tbody>
  </table>
  ${sigBlockHtml}
</body>
</html>`;
}

export async function exportDatasetToWordDoc(options: {
  db: PortalDatabase;
  title: string;
  subtitle?: string;
  headers: string[];
  rows: (string | number)[][];
  filename: string;
  orientation?: PaperOrientation;
  paperSize?: PaperSizeId;
  cssPageSize?: string;
  fontFamily?: PrintFontFamily;
  showKop?: boolean;
  showTtdKepsek?: boolean;
  showTtdGuru?: boolean;
  overrideAccountId?: string;
}) {
  const html = await buildOfficialPrintHtml({ ...options, forWord: true });
  const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = options.filename.endsWith('.doc')
    ? options.filename
    : `${options.filename}.doc`;
  a.click();
  URL.revokeObjectURL(url);
}

export async function printHtmlExactViaIframe(htmlContent: string) {
  const iframe = document.createElement('iframe');
  iframe.style.position = 'fixed';
  iframe.style.right = '0';
  iframe.style.bottom = '0';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = '0';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow?.document;
  if (!doc) {
    document.body.removeChild(iframe);
    return;
  }
  doc.open();
  doc.write(htmlContent);
  doc.close();

  setTimeout(() => {
    iframe.contentWindow?.focus();
    iframe.contentWindow?.print();
    setTimeout(() => {
      if (document.body.contains(iframe)) {
        document.body.removeChild(iframe);
      }
    }, 1000);
  }, 250);
}

export async function printDatasetExact(options: {
  db: PortalDatabase;
  title: string;
  subtitle?: string;
  headers: string[];
  rows: (string | number)[][];
  orientation?: PaperOrientation;
  paperSize?: PaperSizeId;
  cssPageSize?: string;
  fontFamily?: PrintFontFamily;
}) {
  const html = await buildOfficialPrintHtml({
    ...options,
    orientation: options.orientation || 'landscape',
    paperSize: options.paperSize || 'A4',
    forWord: false,
  });
  await printHtmlExactViaIframe(html);
}

export async function exportDatasetToPdf(options: {
  db: PortalDatabase;
  title: string;
  subtitle?: string;
  headers: string[];
  rows: (string | number)[][];
  filename: string;
  orientation?: PaperOrientation;
  paperSize?: PaperSizeId;
  fontFamily?: PrintFontFamily;
  overrideAccountId?: string;
}) {
  const {
    db,
    title,
    subtitle = '',
    headers,
    rows,
    filename,
    orientation = 'landscape',
    paperSize = 'A4',
    fontFamily = 'Times New Roman',
    overrideAccountId,
  } = options;

  const { settings } = db;
  const signer = resolveActiveSigner(db, overrideAccountId);
  const sigPos = settings.signaturePosition;
  const paperCfg = PAPER_SIZE_CONFIG[paperSize] || PAPER_SIZE_CONFIG.A4;

  const doc = new jsPDF({
    orientation,
    unit: 'mm',
    format: paperCfg.jsPdfFormat,
  });

  const pdfFont =
    fontFamily === 'Times New Roman' || fontFamily === 'Georgia'
      ? 'times'
      : 'helvetica';
  doc.setFont(pdfFont, 'normal');

  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  let cursorY = 14;

  const showSigImage = settings.reportPrintPreferences?.showSignatureImage ?? true;

  const [leftLogoPng, rightLogoPng, kepsekSigPng, guruSigPng, stampPng] =
    await Promise.all([
      imageUrlToPngDataUrl(settings.kopSurat.logoLeft),
      imageUrlToPngDataUrl(settings.kopSurat.logoRight),
      showSigImage ? imageUrlToPngDataUrl(settings.principalSignatureImage) : Promise.resolve(null),
      showSigImage ? imageUrlToPngDataUrl(signer.signatureImage) : Promise.resolve(null),
      sigPos.showStamp ? imageUrlToPngDataUrl(signer.schoolStampImage) : Promise.resolve(null),
    ]);

  // Kop Surat
  const logoSize = 19;
  if (leftLogoPng) doc.addImage(leftLogoPng, 'PNG', marginX, cursorY, logoSize, logoSize);
  if (rightLogoPng) {
    doc.addImage(rightLogoPng, 'PNG', pageWidth - marginX - logoSize, cursorY, logoSize, logoSize);
  }
  const centerX = pageWidth / 2;
  doc.setFontSize(10);
  if (settings.kopSurat.line1) doc.text(settings.kopSurat.line1, centerX, cursorY + 4, { align: 'center' });
  if (settings.kopSurat.line2) {
    doc.setFont(pdfFont, 'bold');
    doc.setFontSize(11);
    doc.text(settings.kopSurat.line2, centerX, cursorY + 9, { align: 'center' });
  }
  if (settings.kopSurat.line3) {
    doc.setFont(pdfFont, 'bold');
    doc.setFontSize(13);
    doc.text(settings.kopSurat.line3, centerX, cursorY + 14.5, { align: 'center' });
  }
  doc.setFont(pdfFont, 'normal');
  doc.setFontSize(8.5);
  const contactLines = [settings.kopSurat.line4, settings.kopSurat.line5, settings.kopSurat.line6].filter(Boolean);
  contactLines.forEach((line, idx) => {
    doc.text(line, centerX, cursorY + 19 + idx * 3.8, { align: 'center' });
  });
  const lineY = cursorY + 19 + Math.max(1, contactLines.length) * 3.8 + 1;
  doc.setLineWidth(0.8);
  doc.line(marginX, lineY, pageWidth - marginX, lineY);
  doc.setLineWidth(0.25);
  doc.line(marginX, lineY + 1, pageWidth - marginX, lineY + 1);
  cursorY = lineY + 7;

  // Title
  doc.setFont(pdfFont, 'bold');
  doc.setFontSize(11.5);
  doc.text(title.toUpperCase(), centerX, cursorY, { align: 'center' });
  cursorY += 5;
  if (subtitle) {
    doc.setFont(pdfFont, 'normal');
    doc.setFontSize(8.5);
    doc.text(subtitle, centerX, cursorY, { align: 'center' });
    cursorY += 6;
  }

  autoTable(doc, {
    startY: cursorY,
    margin: { left: marginX, right: marginX },
    styles: {
      font: pdfFont,
      fontSize: 8,
      cellPadding: 2,
      lineColor: [148, 163, 184],
      lineWidth: 0.2,
    },
    headStyles: { fillColor: [15, 23, 42], textColor: 255, fontStyle: 'bold' },
    head: [headers],
    body: rows.length === 0 ? [headers.map(() => '-')] : rows,
  });

  cursorY = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;

  const kepsekSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const guruSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const stampSizePx = Math.max(20, Math.min(260, sigPos.stampSize ?? 70));
  const stampMm = Number((stampSizePx * 0.285).toFixed(2));
  const kepsekSigOffsetX = sigPos.principalSigOffsetX ?? 0;
  const kepsekSigOffsetY = sigPos.principalSigOffsetY ?? 0;
  const guruSigOffsetX = sigPos.teacherSigOffsetX ?? 0;
  const guruSigOffsetY = sigPos.teacherSigOffsetY ?? 0;
  const stampOffsetX = sigPos.stampOffsetX ?? 0;
  const stampOffsetY = sigPos.stampOffsetY ?? 0;

  if (cursorY > pageHeight - 48) {
    doc.addPage();
    cursorY = 18;
  }

  const drawPdfSigBlock = (
    xMm: number,
    yMm: number,
    topLabel: string,
    roleTitle: string,
    personName: string,
    personNip: string,
    sigUrl: string | null,
    withStamp: boolean,
    sigPx: number,
    sigOffsetX: number,
    sigOffsetY: number
  ) => {
    const sigH = Number((sigPx * 0.265).toFixed(2));
    const sigW = Number((sigH * 2.38).toFixed(2));
    doc.setFont(pdfFont, 'normal');
    doc.setFontSize(9);
    doc.text(topLabel, xMm, yMm);
    doc.text(roleTitle, xMm, yMm + 4.5);
    if (withStamp && stampPng) {
      const stampCenterX = xMm + 4 + stampOffsetX * 0.265;
      const stampCenterY = yMm + 15.5 + stampOffsetY * 0.265;
      doc.addImage(
        stampPng,
        'PNG',
        stampCenterX - stampMm / 2,
        stampCenterY - stampMm / 2,
        stampMm,
        stampMm
      );
    }
    if (sigUrl) {
      doc.addImage(
        sigUrl,
        'PNG',
        xMm + sigOffsetX * 0.265,
        yMm + 14.5 - sigH / 2 + sigOffsetY * 0.265,
        sigW,
        sigH
      );
    }
    doc.setFont(pdfFont, 'bold');
    doc.text(personName, xMm, yMm + 25);
    doc.setFont(pdfFont, 'normal');
    doc.text(`NIP. ${personNip || '-'}`, xMm, yMm + 29.5);
  };

  const stampOnKepsek =
    sigPos.showStamp && (sigPos.stampTarget === 'kepsek' || sigPos.stampTarget === 'keduanya');
  const stampOnGuru =
    sigPos.showStamp && (sigPos.stampTarget === 'guru' || sigPos.stampTarget === 'keduanya');

  const usableWidth = pageWidth - marginX * 2;
  let kepsekX = marginX + 6;
  let kepsekY = cursorY;
  let guruX = pageWidth - marginX - 65;
  let guruY = cursorY;

  const maxSigX = pageWidth - marginX - 50;
  if (sigPos.preset === 'kanan-kiri') {
    guruX = marginX + 6;
    kepsekX = pageWidth - marginX - 65;
  } else if (sigPos.preset === 'kiri') {
    kepsekX = marginX + 6;
    guruX = marginX + 6;
    guruY = cursorY + 35;
  } else if (sigPos.preset === 'tengah') {
    kepsekX = pageWidth / 2 - 58;
    guruX = pageWidth / 2 + 12;
  } else if (sigPos.preset === 'kanan') {
    kepsekX = pageWidth - marginX - 65;
    guruX = pageWidth - marginX - 65;
    guruY = cursorY + 35;
  } else if (sigPos.preset === 'bebas') {
    kepsekX = Math.min(maxSigX, marginX + (usableWidth * sigPos.principalX) / 100);
    kepsekY = cursorY + sigPos.principalY * 0.26;
    guruX = Math.min(maxSigX, marginX + (usableWidth * sigPos.teacherX) / 100);
    guruY = cursorY + sigPos.teacherY * 0.26;
  }

  drawPdfSigBlock(
    kepsekX,
    kepsekY,
    'Mengetahui,',
    settings.principalTitle,
    settings.principalName,
    settings.principalNip,
    kepsekSigPng,
    stampOnKepsek,
    kepsekSigSizePx,
    kepsekSigOffsetX,
    kepsekSigOffsetY
  );
  drawPdfSigBlock(
    guruX,
    guruY,
    `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
    signer.roleLabel,
    signer.name,
    signer.nip,
    guruSigPng,
    stampOnGuru,
    guruSigSizePx,
    guruSigOffsetX,
    guruSigOffsetY
  );

  doc.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
}

export async function exportPreviewDomToWordDoc(options: {
  previewElement: HTMLElement;
  title: string;
  filename: string;
  orientation: PaperOrientation;
  cssPageSize: string;
  fontFamily: PrintFontFamily;
}) {
  const clone = options.previewElement.cloneNode(true) as HTMLElement;

  // Remove interactive UI hints (e.g. drag handles) before exporting
  clone.querySelectorAll('[data-export-ignore="true"]').forEach((el) => el.remove());

  // Convert free-positioned signature canvas into an MS Word-compatible percentage table
  const sigCanvas = clone.querySelector('[data-sig-canvas="true"]') as HTMLElement | null;
  if (sigCanvas) {
    const blocks = Array.from(sigCanvas.querySelectorAll('[data-sig-block]')) as HTMLElement[];
    const parsedBlocks = blocks.map((b) => {
      const leftStr = b.style.left || '0%';
      const topStr = b.style.top || '0px';
      const xPct = parseFloat(leftStr) || 0;
      const yPx = parseFloat(topStr) || 0;
      const inner = b.firstElementChild ? b.innerHTML : '';
      return { xPct, yPx, html: inner };
    });
    const wrapper = document.createElement('div');
    wrapper.innerHTML = buildWordFreeSigTableHtml(parsedBlocks);
    sigCanvas.replaceWith(wrapper);
  }

  // Convert any SVG/WebP images inside clone to PNG DataURLs so MS Word displays them reliably
  const imgs = Array.from(clone.querySelectorAll('img'));
  await Promise.all(
    imgs.map(async (img) => {
      const src = img.getAttribute('src') || '';
      if (!src) return;
      const png = await imageUrlToPngDataUrl(src);
      if (png) {
        img.setAttribute('src', png);
      }
    })
  );

  // Convert Tailwind border/padding basics for MS Word fidelity
  clone.style.border = 'none';
  clone.style.boxShadow = 'none';
  clone.style.padding = '0';
  clone.style.margin = '0';
  clone.style.maxWidth = 'none';
  clone.style.minWidth = '0';

  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(options.title)}</title>
<style>
  @page { size: ${options.cssPageSize} ${options.orientation}; margin: 12mm; }
  body { font-family: "${options.fontFamily}", serif; color: #000000; margin: 0; padding: 0; }
  table { border-collapse: collapse; width: 100%; }
  th, td { border: 1px solid #000000; padding: 4px 6px; font-size: 9pt; }
  th { background-color: #f1f5f9; font-weight: bold; }
</style>
</head>
<body>${clone.outerHTML}</body>
</html>`;

  const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = options.filename.endsWith('.doc')
    ? options.filename
    : `${options.filename}.doc`;
  a.click();
  URL.revokeObjectURL(url);
}
