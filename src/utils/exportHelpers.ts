import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import {
  KopSuratSettings,
  PaperOrientation,
  PaperSizeId,
  PortalDatabase,
  PrintFontFamily,
  SignaturePositionSettings,
  UserAccount,
} from '../types';
import {
  compositeSignatureMediaToPng,
  fitContainDimensions,
  formatLongDateId,
  getPrintFontCss,
  getTodayISO,
  imageUrlToPngWithMeta,
  KOP_LINE_KEYS,
  normalizeKopLineSpacing,
  PngImageInfo,
  resolveKopLineStyle,
} from './portalHelpers';

export const PAPER_SIZE_CONFIG: Record<
  PaperSizeId,
  {
    label: string;
    shortLabel: string;
    jsPdfFormat: string | [number, number];
    cssPageSize: string;
    widthMm: number;
    heightMm: number;
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
    widthMm: 210,
    heightMm: 297,
    portraitMaxW: 'max-w-[794px]',
    landscapeMinW: 'min-w-[920px]',
    landscapeMaxW: 'max-w-[1123px]',
  },
  F4: {
    label: 'F4 / Folio (215 × 330 mm)',
    shortLabel: 'F4',
    jsPdfFormat: [215, 330],
    cssPageSize: '215mm 330mm',
    widthMm: 215,
    heightMm: 330,
    portraitMaxW: 'max-w-[812px]',
    landscapeMinW: 'min-w-[1000px]',
    landscapeMaxW: 'max-w-[1247px]',
  },
  Letter: {
    label: 'Letter (216 × 279 mm)',
    shortLabel: 'Letter',
    jsPdfFormat: 'letter',
    cssPageSize: 'letter',
    widthMm: 216,
    heightMm: 279,
    portraitMaxW: 'max-w-[816px]',
    landscapeMinW: 'min-w-[900px]',
    landscapeMaxW: 'max-w-[1056px]',
  },
  Legal: {
    label: 'Legal (216 × 356 mm)',
    shortLabel: 'Legal',
    jsPdfFormat: 'legal',
    cssPageSize: 'legal',
    widthMm: 216,
    heightMm: 356,
    portraitMaxW: 'max-w-[816px]',
    landscapeMinW: 'min-w-[1040px]',
    landscapeMaxW: 'max-w-[1344px]',
  },
  A5: {
    label: 'A5 (148 × 210 mm)',
    shortLabel: 'A5',
    jsPdfFormat: 'a5',
    cssPageSize: 'A5',
    widthMm: 148,
    heightMm: 210,
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

export function resolvePrincipalStampImage(db: PortalDatabase): string {
  const kepsekAcc = db.accounts.find(
    (a) => a.role === 'Kepala Sekolah' && a.status === 'Aktif' && Boolean(a.schoolStampImage)
  );
  return kepsekAcc?.schoolStampImage || db.settings.defaultSchoolStampImage || '';
}

export function mapPrintFontToPdf(fontFamily?: PrintFontFamily): 'times' | 'helvetica' {
  return !fontFamily || fontFamily === 'Times New Roman' || fontFamily === 'Georgia'
    ? 'times'
    : 'helvetica';
}

export type PdfLogoInput = string | PngImageInfo | null;

function resolvePdfLogoInfo(input: PdfLogoInput): { dataUrl: string; width: number; height: number } | null {
  if (!input) return null;
  if (typeof input === 'string') return { dataUrl: input, width: 1, height: 1 };
  return input;
}

export function drawPdfKopSurat(params: {
  doc: jsPDF;
  kop: KopSuratSettings;
  leftLogoPng: PdfLogoInput;
  rightLogoPng: PdfLogoInput;
  logoLeftPx?: number;
  logoRightPx?: number;
  marginX: number;
  startY: number;
  fallbackFont?: 'times' | 'helvetica';
}): number {
  const {
    doc,
    kop,
    marginX,
    startY,
    fallbackFont = 'times',
  } = params;
  const leftLogo = resolvePdfLogoInfo(params.leftLogoPng);
  const rightLogo = resolvePdfLogoInfo(params.rightLogoPng);
  const pageWidth = doc.internal.pageSize.getWidth();
  const logoLeftPx = Math.max(28, Math.min(150, params.logoLeftPx ?? kop.logoLeftSize ?? 72));
  const logoRightPx = Math.max(28, Math.min(150, params.logoRightPx ?? kop.logoRightSize ?? 72));
  const PX_TO_MM = 0.264583;
  const logoLeftBoxMm = Math.round(logoLeftPx * PX_TO_MM * 10) / 10;
  const logoRightBoxMm = Math.round(logoRightPx * PX_TO_MM * 10) / 10;
  const logoLeftColMm = Math.max(19, logoLeftBoxMm + 2.5);
  const logoRightColMm = Math.max(19, logoRightBoxMm + 2.5);

  const textLeftX = marginX + (leftLogo ? logoLeftColMm + 2 : 0);
  const textRightX = pageWidth - marginX - (rightLogo ? logoRightColMm + 2 : 0);
  const textMaxWidth = Math.max(40, textRightX - textLeftX);
  const textCenterX = (textLeftX + textRightX) / 2;
  const kopPdfFont = kop.fontFamily ? mapPrintFontToPdf(kop.fontFamily) : fallbackFont;
  const lineSpacing = normalizeKopLineSpacing(kop.lineSpacing);

  const preparedLines = KOP_LINE_KEYS.map((key, idx) => {
    const raw = kop[key];
    if (!raw) return null;
    const style = resolveKopLineStyle(kop, key);
    const jsStyle =
      style.bold && style.italic
        ? 'bolditalic'
        : style.bold
        ? 'bold'
        : style.italic
        ? 'italic'
        : 'normal';
    doc.setFont(kopPdfFont, jsStyle);
    doc.setFontSize(style.fontSize);
    const text = idx < 3 ? raw.toUpperCase() : raw;
    const wrapped = doc.splitTextToSize(text, textMaxWidth) as string[];
    const lineHeightMm = Math.max(2.4, style.fontSize * 0.352778 * lineSpacing);
    const extraTopMm = key === 'line4' ? 0.8 : 0;
    return { style, jsStyle, wrapped, lineHeightMm, extraTopMm };
  }).filter((item): item is NonNullable<typeof item> => Boolean(item));

  const textHeightMm = preparedLines.reduce(
    (sum, item) => sum + item.extraTopMm + item.wrapped.length * item.lineHeightMm,
    0
  );
  const kopContentHeightMm = Math.max(
    textHeightMm,
    leftLogo ? logoLeftBoxMm : 0,
    rightLogo ? logoRightBoxMm : 0
  );

  if (leftLogo) {
    const fitted = fitContainDimensions(leftLogo.width, leftLogo.height, logoLeftBoxMm, logoLeftBoxMm);
    const x = marginX + (logoLeftColMm - fitted.width) / 2;
    const y = startY + Math.max(0, (kopContentHeightMm - fitted.height) / 2);
    doc.addImage(leftLogo.dataUrl, 'PNG', x, y, fitted.width, fitted.height);
  }
  if (rightLogo) {
    const fitted = fitContainDimensions(rightLogo.width, rightLogo.height, logoRightBoxMm, logoRightBoxMm);
    const x = pageWidth - marginX - logoRightColMm + (logoRightColMm - fitted.width) / 2;
    const y = startY + Math.max(0, (kopContentHeightMm - fitted.height) / 2);
    doc.addImage(rightLogo.dataUrl, 'PNG', x, y, fitted.width, fitted.height);
  }

  let boxTopMm = startY + Math.max(0, (kopContentHeightMm - textHeightMm) / 2);
  doc.setDrawColor(0, 0, 0);

  preparedLines.forEach(({ style, jsStyle, wrapped, lineHeightMm, extraTopMm }) => {
    boxTopMm += extraTopMm;
    doc.setFont(kopPdfFont, jsStyle);
    doc.setFontSize(style.fontSize);
    const fontMm = style.fontSize * 0.352778;
    const capMm = fontMm * 0.72;
    const baselineOffsetMm = Math.max(capMm * 0.95, (lineHeightMm + capMm) / 2);
    const underlineOffsetMm = Math.max(0.5, fontMm * 0.16);

    wrapped.forEach((wLine) => {
      const baselineY = boxTopMm + baselineOffsetMm;
      if (style.align === 'left') {
        doc.text(wLine, textLeftX, baselineY, { align: 'left' });
      } else if (style.align === 'right') {
        doc.text(wLine, textRightX, baselineY, { align: 'right' });
      } else if (style.align === 'justify') {
        doc.text(wLine, textLeftX, baselineY, { align: 'justify', maxWidth: textMaxWidth });
      } else {
        doc.text(wLine, textCenterX, baselineY, { align: 'center' });
      }

      if (style.underline) {
        const textWidth = Math.min(textMaxWidth, doc.getTextWidth(wLine));
        let uStartX = textCenterX - textWidth / 2;
        let uEndX = textCenterX + textWidth / 2;
        if (style.align === 'left') {
          uStartX = textLeftX;
          uEndX = textLeftX + textWidth;
        } else if (style.align === 'right') {
          uStartX = textRightX - textWidth;
          uEndX = textRightX;
        } else if (style.align === 'justify') {
          uStartX = textLeftX;
          uEndX = wLine.includes(' ') ? textRightX : textLeftX + textWidth;
        }
        doc.setLineWidth(Math.max(0.2, style.fontSize * 0.025));
        doc.line(uStartX, baselineY + underlineOffsetMm, uEndX, baselineY + underlineOffsetMm);
      }

      boxTopMm += lineHeightMm;
    });
  });

  // 2.65mm (~10px) paddingBottom above double border matching Live Preview 1:1
  const lineY = startY + kopContentHeightMm + 2.65;
  doc.setLineWidth(0.75);
  doc.line(marginX, lineY, pageWidth - marginX, lineY);
  doc.setLineWidth(0.3);
  doc.line(marginX, lineY + 1.0, pageWidth - marginX, lineY + 1.0);
  // 1.0mm bottom rule + ~4.2mm (16px) marginBottom + ~4.0mm 12pt title ascent = 9.2mm
  return lineY + 9.2;
}

export function resolveActiveSigner(
  db: PortalDatabase,
  overrideAccountId?: string
): ActiveSignerInfo {
  const { accounts, teachers, settings, session } = db;
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

  const signerEmailLower = (account?.email || session.email || '').trim().toLowerCase();
  const signerNipLower = (account?.nip || '').trim().toLowerCase();
  const signerNameLower = (account?.name || session.userName || '').trim().toLowerCase();
  const matchedTeacher = teachers.find((t) => {
    if (signerEmailLower && (t.email || '').trim().toLowerCase() === signerEmailLower) return true;
    if (signerNipLower && (t.nip || '').trim().toLowerCase() === signerNipLower) return true;
    if (signerNameLower && (t.name || '').trim().toLowerCase() === signerNameLower) return true;
    return false;
  });

  const roleLabel =
    account?.role && account.role !== 'Admin'
      ? account.role
      : matchedTeacher?.position && matchedTeacher.position !== 'Admin / Operator Sekolah'
      ? matchedTeacher.position
      : 'Guru Mata Pelajaran';

  return {
    name: account?.name || session.userName || settings.teacherName,
    nip: account?.nip || matchedTeacher?.nip || settings.teacherNip,
    roleLabel,
    signatureImage: account?.signatureImage || '',
    schoolStampImage: resolvePrincipalStampImage(db),
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
      <table data-layout-table="true" style="width:100%; margin-top:26px; border-collapse:collapse;">
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
      <table data-layout-table="true" style="width:100%; margin-top:${idx === 0 ? 26 : 16}px; border-collapse:collapse;">
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
    <table data-layout-table="true" style="width:100%; margin-top:26px; border-collapse:collapse;">
      <tr>
        ${leftSpacer > 0 ? `<td style="width:${leftSpacer}%; border:none; padding:0;"></td>` : ''}
        <td style="width:${firstWidth}%; vertical-align:top; border:none; padding:${first.yPx}px 0 0 0;">${first.html}</td>
        ${gapSpacer > 0 ? `<td style="width:${gapSpacer}%; border:none; padding:0;"></td>` : ''}
        <td style="vertical-align:top; border:none; padding:${second.yPx}px 0 0 0;">${second.html}</td>
      </tr>
    </table>
  `;
}

function getWordPageSizeCss(paperSize: PaperSizeId, orientation: PaperOrientation, fallbackCss?: string): string {
  const cfg = PAPER_SIZE_CONFIG[paperSize];
  if (!cfg) {
    return `${fallbackCss || 'A4'} ${orientation}`;
  }
  const w = orientation === 'landscape' ? cfg.heightMm : cfg.widthMm;
  const h = orientation === 'landscape' ? cfg.widthMm : cfg.heightMm;
  return `${w}mm ${h}mm`;
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
  const wordPageSize = getWordPageSizeCss(paperSize, orientation, resolvedCssPageSize);
  const { settings } = db;
  const signer = resolveActiveSigner(db, overrideAccountId);
  const sigPos: SignaturePositionSettings = settings.signaturePosition;

  const showSigImage = settings.reportPrintPreferences?.showSignatureImage ?? true;
  const principalStampImage = resolvePrincipalStampImage(db);

  const [leftLogoInfo, rightLogoInfo, kepsekSigInfo, guruSigInfo, stampInfo] =
    await Promise.all([
      showKop ? imageUrlToPngWithMeta(settings.kopSurat.logoLeft) : Promise.resolve(null),
      showKop ? imageUrlToPngWithMeta(settings.kopSurat.logoRight) : Promise.resolve(null),
      showTtdKepsek && showSigImage
        ? imageUrlToPngWithMeta(settings.principalSignatureImage)
        : Promise.resolve(null),
      showTtdGuru && showSigImage
        ? imageUrlToPngWithMeta(signer.signatureImage)
        : Promise.resolve(null),
      sigPos.showStamp ? imageUrlToPngWithMeta(principalStampImage) : Promise.resolve(null),
    ]);

  const kopLeftPx = Math.max(28, Math.min(150, settings.kopSurat.logoLeftSize ?? 72));
  const kopRightPx = Math.max(28, Math.min(150, settings.kopSurat.logoRightSize ?? 72));
  const kopLeftColPx = Math.max(76, kopLeftPx + 14);
  const kopRightColPx = Math.max(76, kopRightPx + 14);
  const kopFontCss = getPrintFontCss(settings.kopSurat.fontFamily);
  const bodyFontCss = getPrintFontCss(fontFamily);
  const kopLineSpacing = normalizeKopLineSpacing(settings.kopSurat.lineSpacing);
  const kopLineHeightPct = Math.round(kopLineSpacing * 100);

  const leftLogoTag = leftLogoInfo
    ? (() => {
        const d = fitContainDimensions(leftLogoInfo.width, leftLogoInfo.height, kopLeftPx, kopLeftPx);
        const w = Math.round(d.width);
        const h = Math.round(d.height);
        return `<img src="${leftLogoInfo.dataUrl}" width="${w}" height="${h}" style="width:${w}px; height:${h}px; display:inline-block;" />`;
      })()
    : '';

  const rightLogoTag = rightLogoInfo
    ? (() => {
        const d = fitContainDimensions(rightLogoInfo.width, rightLogoInfo.height, kopRightPx, kopRightPx);
        const w = Math.round(d.width);
        const h = Math.round(d.height);
        return `<img src="${rightLogoInfo.dataUrl}" width="${w}" height="${h}" style="width:${w}px; height:${h}px; display:inline-block;" />`;
      })()
    : '';

  const kopLinesHtml = KOP_LINE_KEYS.map((key, idx) => {
    const text = settings.kopSurat[key];
    if (!text) return '';
    const s = resolveKopLineStyle(settings.kopSurat, key);
    const styles = [
      'margin:0',
      'padding:0',
      `line-height:${kopLineHeightPct}%`,
      `font-size:${s.fontSize}pt`,
      `font-weight:${s.bold ? 'bold' : 'normal'}`,
      `font-style:${s.italic ? 'italic' : 'normal'}`,
      `text-decoration:${s.underline ? 'underline' : 'none'}`,
      `text-align:${s.align}`,
      s.align === 'justify' ? 'text-align-last:justify' : '',
      idx < 3 ? 'text-transform:uppercase' : '',
      key === 'line4' ? 'margin-top:3px' : '',
    ]
      .filter(Boolean)
      .join('; ');
    return `<div style="${styles};">${escapeHtml(text)}</div>`;
  }).join('');

  const kopHtml = showKop
    ? `
      <div style="border-bottom:4px double #000000; mso-border-bottom-alt:thick-thin-medium-gap windowtext 3pt; padding-bottom:10px; margin-bottom:16px;">
        <table data-layout-table="true" style="width:100%; border-collapse:collapse;">
          <tr>
            <td style="width:${kopLeftColPx}px; text-align:center; vertical-align:middle; padding:0 6px; border:none;">
              ${leftLogoTag}
            </td>
            <td style="font-family:${kopFontCss}; vertical-align:middle; line-height:${kopLineHeightPct}%; padding:0 4px; border:none;">
              ${kopLinesHtml}
            </td>
            <td style="width:${kopRightColPx}px; text-align:center; vertical-align:middle; padding:0 6px; border:none;">
              ${rightLogoTag}
            </td>
          </tr>
        </table>
      </div>
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
  const stampLeftPx =
    (sigPos.preset === 'tengah' ? 87 : 18) - Math.round(stampSizePx / 2) + stampOffsetX;
  const stampTopPx = Math.round((72 - stampSizePx) / 2) + stampOffsetY;

  const buildSigMediaHtml = async (
    sigInfo: PngImageInfo | null,
    withStamp: boolean,
    sigPx: number,
    sigOffsetX: number,
    sigOffsetY: number
  ): Promise<string> => {
    const maxSigW = Math.round(sigPx * 2.85);
    if (forWord) {
      const comp = await compositeSignatureMediaToPng({
        stampSrc: withStamp && stampInfo ? stampInfo.dataUrl : null,
        stampLeftPx,
        stampTopPx,
        stampSizePx,
        sigSrc: sigInfo ? sigInfo.dataUrl : null,
        sigLeftPx: sigOffsetX,
        sigTopPx: sigOffsetY,
        sigHeightPx: sigPx,
        sigMaxWidthPx: maxSigW,
        alignCenter: sigPos.preset === 'tengah',
        boxWidthPx: 220,
        boxHeightPx: 72,
      });
      if (!comp) return `<div style="height:72px; margin:4px 0;">&nbsp;</div>`;
      return `<div style="margin:4px 0;"><img src="${comp.dataUrl}" width="${comp.width}" height="${comp.height}" style="width:${comp.width}px; height:${comp.height}px; display:block;" /></div>`;
    }

    const stampTag =
      withStamp && stampInfo
        ? (() => {
            const d = fitContainDimensions(stampInfo.width, stampInfo.height, stampSizePx, stampSizePx);
            const w = Math.round(d.width);
            const h = Math.round(d.height);
            return `<img src="${stampInfo.dataUrl}" width="${w}" height="${h}" style="position:absolute; left:${stampLeftPx}px; top:${stampTopPx}px; width:${w}px; height:${h}px; object-fit:contain; opacity:0.85; mix-blend-mode:multiply;" />`;
          })()
        : '';

    const sigTag = sigInfo
      ? (() => {
          const d = fitContainDimensions(sigInfo.width, sigInfo.height, maxSigW, sigPx);
          const w = Math.round(d.width);
          const h = Math.round(d.height);
          return `<img src="${sigInfo.dataUrl}" width="${w}" height="${h}" style="position:relative; z-index:2; left:${sigOffsetX}px; top:${sigOffsetY}px; width:${w}px; height:${h}px; object-fit:contain; margin-top:${Math.round((72 - h) / 2)}px;" />`;
        })()
      : '';

    return `
      <div style="position:relative; height:72px; margin:4px 0; overflow:visible;">
        ${stampTag}
        ${sigTag}
      </div>
    `;
  };

  const renderSigBox = async (
    topLabel: string,
    roleTitle: string,
    personName: string,
    personNip: string,
    sigInfo: PngImageInfo | null,
    withStamp: boolean,
    sigPx: number,
    sigOffsetX: number,
    sigOffsetY: number
  ) => {
    const mediaHtml = await buildSigMediaHtml(sigInfo, withStamp, sigPx, sigOffsetX, sigOffsetY);
    return `
      <div style="font-size:9.5pt; line-height:1.35; text-align:${sigPos.preset === 'tengah' ? 'center' : 'left'};">
        <div style="margin:0;">${escapeHtml(topLabel)}</div>
        <div style="margin:0;">${escapeHtml(roleTitle)}</div>
        ${mediaHtml}
        <div style="margin:0; font-weight:bold; text-decoration:underline;">${escapeHtml(personName)}</div>
        <div style="margin:0;">NIP. ${escapeHtml(personNip || '-')}</div>
      </div>
    `;
  };

  const stampOnKepsek = sigPos.showStamp;
  const stampOnGuru = false;

  const [kepsekBox, guruBox] = await Promise.all([
    showTtdKepsek
      ? renderSigBox(
          'Mengetahui,',
          settings.principalTitle,
          settings.principalName,
          settings.principalNip,
          kepsekSigInfo,
          stampOnKepsek,
          kepsekSigSizePx,
          kepsekSigOffsetX,
          kepsekSigOffsetY
        )
      : Promise.resolve(''),
    showTtdGuru
      ? renderSigBox(
          `${settings.signatureLocation}, ${formatLongDateId(getTodayISO())}`,
          signer.roleLabel,
          signer.name,
          signer.nip,
          guruSigInfo,
          stampOnGuru,
          guruSigSizePx,
          guruSigOffsetX,
          guruSigOffsetY
        )
      : Promise.resolve(''),
  ]);

  let sigBlockHtml = '';
  if (showTtdKepsek || showTtdGuru) {
    if (sigPos.preset === 'kanan-kiri') {
      sigBlockHtml = `
        <table data-layout-table="true" style="width:100%; margin-top:26px; border-collapse:collapse;">
          <tr>
            <td style="width:50%; vertical-align:top; padding-left:16px; border:none;">${guruBox}</td>
            <td style="width:50%; vertical-align:top; padding-left:18%; border:none;">${kepsekBox}</td>
          </tr>
        </table>
      `;
    } else if (sigPos.preset === 'kiri') {
      sigBlockHtml = `<div style="margin-top:26px; padding-left:16px;">${kepsekBox}<div style="margin-top:18px;">${guruBox}</div></div>`;
    } else if (sigPos.preset === 'tengah') {
      sigBlockHtml = `
        <table data-layout-table="true" style="width:100%; margin-top:26px; border-collapse:collapse; text-align:center;">
          <tr>
            ${showTtdKepsek ? `<td style="vertical-align:top; text-align:center; border:none;">${kepsekBox}</td>` : ''}
            ${showTtdGuru ? `<td style="vertical-align:top; text-align:center; border:none;">${guruBox}</td>` : ''}
          </tr>
        </table>
      `;
    } else if (sigPos.preset === 'kanan') {
      sigBlockHtml = `
        <table data-layout-table="true" style="width:100%; margin-top:26px; border-collapse:collapse;">
          <tr>
            <td style="border:none; padding:0;"></td>
            <td style="width:240px; vertical-align:top; border:none; padding:0;">
              ${kepsekBox}
              <div style="margin-top:18px;">${guruBox}</div>
            </td>
          </tr>
        </table>
      `;
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
        <table data-layout-table="true" style="width:100%; margin-top:26px; border-collapse:collapse;">
          <tr>
            <td style="width:50%; vertical-align:top; padding-left:16px; border:none;">${kepsekBox}</td>
            <td style="width:50%; vertical-align:top; padding-left:18%; border:none;">${guruBox}</td>
          </tr>
        </table>
      `;
    }
  }

  const thCells = headers
    .map(
      (h) =>
        `<th style="border:1px solid #000000; background:#f1f5f9; padding:5px 7px; font-size:8.5pt; text-align:left; font-weight:bold;">${escapeHtml(
          h
        )}</th>`
    )
    .join('');

  const bodyRows =
    rows.length === 0
      ? `<tr><td colspan="${headers.length}" style="border:1px solid #000000; padding:12px; text-align:center; font-size:8.5pt;">Tidak ada data</td></tr>`
      : rows
          .map(
            (r) =>
              `<tr>${r
                .map(
                  (cell) =>
                    `<td style="border:1px solid #000000; padding:4px 7px; font-size:8.5pt; vertical-align:top;">${escapeHtml(
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
  @page { size: ${forWord ? wordPageSize : `${resolvedCssPageSize} ${orientation}`}; mso-page-orientation: ${orientation}; margin: 12mm; }
  * { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body { font-family: ${bodyFontCss}; color: #000000; margin: 0; padding: 0; background: #ffffff; }
  h1, h2, h3, h4, p { margin: 0; padding: 0; }
  table { border-collapse: collapse; }
  table[data-layout-table="true"] td, table[data-layout-table="true"] th { border: none !important; background: transparent !important; }
  table.report-data-table { width: 100%; border-collapse: collapse; }
  table.report-data-table thead { display: table-header-group; }
  table.report-data-table tr { break-inside: avoid; page-break-inside: avoid; }
  table.report-data-table th, table.report-data-table td { border: 1px solid #000000; padding: 4px 6px; font-size: 8.5pt; }
  table.report-data-table th { background-color: #f1f5f9; font-weight: bold; }
</style>
</head>
<body>
  ${kopHtml}
  <div style="text-align:center; margin-bottom:14px;">
    <div style="font-size:12pt; font-weight:bold; text-transform:uppercase; text-decoration:underline;">${escapeHtml(
      title
    )}</div>
    ${subtitle ? `<div style="font-size:9pt; color:#334155; margin-top:4px;">${escapeHtml(subtitle)}</div>` : ''}
  </div>
  <table class="report-data-table" style="width:100%; border-collapse:collapse;">
    <thead><tr>${thCells}</tr></thead>
    <tbody>${bodyRows}</tbody>
  </table>
  <div style="break-inside:avoid; page-break-inside:avoid;">
    ${sigBlockHtml}
  </div>
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

  const win = iframe.contentWindow;
  const doc = win?.document;
  if (!doc || !win) {
    document.body.removeChild(iframe);
    return;
  }
  doc.open();
  doc.write(htmlContent);
  doc.close();

  const imgs = Array.from(doc.querySelectorAll('img'));
  await Promise.all(
    imgs.map((img) => {
      if (img.complete) return Promise.resolve();
      return new Promise<void>((res) => {
        img.onload = () => res();
        img.onerror = () => res();
        setTimeout(res, 1500);
      });
    })
  );

  setTimeout(() => {
    win.focus();
    win.print();
    setTimeout(() => {
      if (document.body.contains(iframe)) {
        document.body.removeChild(iframe);
      }
    }, 1500);
  }, 150);
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
  const marginX = 12;
  let cursorY = 12;

  const showSigImage = settings.reportPrintPreferences?.showSignatureImage ?? true;
  const principalStampImage = resolvePrincipalStampImage(db);

  const [leftLogoInfo, rightLogoInfo, kepsekSigInfo, guruSigInfo, stampInfo] =
    await Promise.all([
      imageUrlToPngWithMeta(settings.kopSurat.logoLeft),
      imageUrlToPngWithMeta(settings.kopSurat.logoRight),
      showSigImage ? imageUrlToPngWithMeta(settings.principalSignatureImage) : Promise.resolve(null),
      showSigImage ? imageUrlToPngWithMeta(signer.signatureImage) : Promise.resolve(null),
      sigPos.showStamp ? imageUrlToPngWithMeta(principalStampImage) : Promise.resolve(null),
    ]);

  // Kop Surat
  cursorY = drawPdfKopSurat({
    doc,
    kop: settings.kopSurat,
    leftLogoPng: leftLogoInfo,
    rightLogoPng: rightLogoInfo,
    marginX,
    startY: cursorY,
    fallbackFont: pdfFont,
  });

  const centerX = pageWidth / 2;
  const upperTitle = title.toUpperCase();
  // Title
  doc.setFont(pdfFont, 'bold');
  doc.setFontSize(12);
  doc.text(upperTitle, centerX, cursorY, { align: 'center' });
  const titleW = doc.getTextWidth(upperTitle);
  doc.setLineWidth(0.35);
  doc.line(centerX - titleW / 2, cursorY + 1, centerX + titleW / 2, cursorY + 1);
  cursorY += 5.5;
  if (subtitle) {
    doc.setFont(pdfFont, 'normal');
    doc.setFontSize(9);
    doc.text(subtitle, centerX, cursorY, { align: 'center' });
    cursorY += 6;
  }

  autoTable(doc, {
    startY: cursorY,
    margin: { left: marginX, right: marginX },
    styles: {
      font: pdfFont,
      fontSize: 8,
      textColor: 0,
      cellPadding: 1.5,
      lineColor: [0, 0, 0],
      lineWidth: 0.2,
    },
    headStyles: {
      fillColor: [241, 245, 249],
      textColor: 0,
      fontStyle: 'bold',
      lineColor: [0, 0, 0],
      lineWidth: 0.2,
    },
    head: [headers],
    body: rows.length === 0 ? [headers.map(() => '-')] : rows,
  });

  cursorY = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 8;

  const PX_TO_MM = 0.264583;
  const kepsekSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.principalSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const guruSigSizePx = Math.max(
    16,
    Math.min(260, sigPos.teacherSignatureSize ?? sigPos.signatureSize ?? 60)
  );
  const stampSizePx = Math.max(20, Math.min(260, sigPos.stampSize ?? 70));
  const stampMm = Number((stampSizePx * PX_TO_MM).toFixed(2));
  const kepsekSigOffsetX = sigPos.principalSigOffsetX ?? 0;
  const kepsekSigOffsetY = sigPos.principalSigOffsetY ?? 0;
  const guruSigOffsetX = sigPos.teacherSigOffsetX ?? 0;
  const guruSigOffsetY = sigPos.teacherSigOffsetY ?? 0;
  const stampOffsetX = sigPos.stampOffsetX ?? 0;
  const stampOffsetY = sigPos.stampOffsetY ?? 0;

  if (cursorY > pageHeight - 48) {
    doc.addPage();
    cursorY = 14;
  }

  const drawPdfSigBlock = (
    xMm: number,
    yMm: number,
    topLabel: string,
    roleTitle: string,
    personName: string,
    personNip: string,
    sigInfo: PngImageInfo | null,
    withStamp: boolean,
    sigPx: number,
    sigOffsetX: number,
    sigOffsetY: number
  ) => {
    const sigBoxH = Number((sigPx * PX_TO_MM).toFixed(2));
    const sigMaxW = Number((sigBoxH * 2.85).toFixed(2));
    doc.setFont(pdfFont, 'normal');
    doc.setFontSize(9.5);
    doc.text(topLabel, xMm, yMm);
    doc.text(roleTitle, xMm, yMm + 4.5);
    if (withStamp && stampInfo) {
      const fittedStamp = fitContainDimensions(stampInfo.width, stampInfo.height, stampMm, stampMm);
      const stampCenterX = xMm + 4.8 + stampOffsetX * PX_TO_MM;
      const stampCenterY = yMm + 16 + stampOffsetY * PX_TO_MM;
      doc.addImage(
        stampInfo.dataUrl,
        'PNG',
        stampCenterX - fittedStamp.width / 2,
        stampCenterY - fittedStamp.height / 2,
        fittedStamp.width,
        fittedStamp.height
      );
    }
    if (sigInfo) {
      const fittedSig = fitContainDimensions(sigInfo.width, sigInfo.height, sigMaxW, sigBoxH);
      doc.addImage(
        sigInfo.dataUrl,
        'PNG',
        xMm + sigOffsetX * PX_TO_MM,
        yMm + 15 - fittedSig.height / 2 + sigOffsetY * PX_TO_MM,
        fittedSig.width,
        fittedSig.height
      );
    }
    doc.setFont(pdfFont, 'bold');
    doc.text(personName, xMm, yMm + 25.5);
    const nameW = doc.getTextWidth(personName);
    doc.setLineWidth(0.25);
    doc.line(xMm, yMm + 26.2, xMm + nameW, yMm + 26.2);
    doc.setFont(pdfFont, 'normal');
    doc.text(`NIP. ${personNip || '-'}`, xMm, yMm + 30);
  };

  const stampOnKepsek = sigPos.showStamp;
  const stampOnGuru = false;

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
    kepsekSigInfo,
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
    guruSigInfo,
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
  paperSize?: PaperSizeId;
  cssPageSize: string;
  fontFamily: PrintFontFamily;
}) {
  const clone = options.previewElement.cloneNode(true) as HTMLElement;

  // Remove interactive UI hints (e.g. drag handles) before exporting
  clone.querySelectorAll('[data-export-ignore="true"]').forEach((el) => el.remove());

  // Ensure every Kop Surat double border & spacing render identically in MS Word
  clone.querySelectorAll('[data-kop-wrapper="true"]').forEach((el) => {
    const kopWrapper = el as HTMLElement;
    kopWrapper.style.borderBottom = '4px double #000000';
    kopWrapper.style.setProperty('mso-border-bottom-alt', 'thick-thin-medium-gap windowtext 3pt');
    kopWrapper.style.paddingBottom = '10px';
    kopWrapper.style.marginBottom = '16px';
  });

  // Composite Stamp + Signature overlays into single crisp PNGs so MS Word matches Preview 1:1
  const sigMediaBoxes = Array.from(clone.querySelectorAll('[data-sig-media="true"]')) as HTMLElement[];
  await Promise.all(
    sigMediaBoxes.map(async (box) => {
      const stampSrc = box.getAttribute('data-stamp-src') || '';
      const sigSrc = box.getAttribute('data-sig-src') || '';
      if (!stampSrc && !sigSrc) {
        box.innerHTML = '&nbsp;';
        return;
      }
      const alignCenter = box.getAttribute('data-align-center') === 'true';
      const stampLeftPx = Number(box.getAttribute('data-stamp-left') || '0');
      const stampTopPx = Number(box.getAttribute('data-stamp-top') || '0');
      const stampSizePx = Number(box.getAttribute('data-stamp-size') || '70');
      const sigLeftPx = Number(box.getAttribute('data-sig-left') || '0');
      const sigTopPx = Number(box.getAttribute('data-sig-top') || '0');
      const sigHeightPx = Number(box.getAttribute('data-sig-height') || '60');
      const sigMaxWidthPx = Number(box.getAttribute('data-sig-max-width') || String(Math.round(sigHeightPx * 2.85)));

      const comp = await compositeSignatureMediaToPng({
        stampSrc: stampSrc || null,
        stampLeftPx,
        stampTopPx,
        stampSizePx,
        sigSrc: sigSrc || null,
        sigLeftPx,
        sigTopPx,
        sigHeightPx,
        sigMaxWidthPx,
        alignCenter,
        boxWidthPx: 220,
        boxHeightPx: 72,
      });

      if (comp) {
        box.style.height = 'auto';
        box.style.display = 'block';
        box.style.position = 'static';
        box.innerHTML = `<img src="${comp.dataUrl}" width="${comp.width}" height="${comp.height}" style="width:${comp.width}px; height:${comp.height}px; display:block;${alignCenter ? ' margin:0 auto;' : ''}" />`;
      }
    })
  );

  // Convert all free-positioned signature canvases into MS Word-compatible percentage tables
  clone.querySelectorAll('[data-sig-canvas="true"]').forEach((el) => {
    const sigCanvas = el as HTMLElement;
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
  });

  // Convert chart grid into an MS Word-compatible table layout
  const chartGrids = Array.from(clone.querySelectorAll('[data-chart-grid="true"]')) as HTMLElement[];
  chartGrids.forEach((grid) => {
    const cols = Number(grid.getAttribute('data-chart-cols') || '2');
    const cards = Array.from(grid.querySelectorAll('[data-chart-card="true"]')) as HTMLElement[];
    if (cards.length === 0) return;
    const rowsHtml: string[] = [];
    for (let i = 0; i < cards.length; i += cols) {
      const cells: string[] = [];
      for (let c = 0; c < cols; c++) {
        const card = cards[i + c];
        const widthPct = Math.floor(100 / cols);
        if (card) {
          cells.push(
            `<td style="width:${widthPct}%; vertical-align:top; border:1px solid #000000; padding:8px; background:#ffffff;">${card.innerHTML}</td>`
          );
        } else {
          cells.push(`<td style="width:${widthPct}%; border:none; padding:8px;"></td>`);
        }
      }
      rowsHtml.push(`<tr>${cells.join('')}</tr>`);
    }
    const wrapper = document.createElement('div');
    wrapper.innerHTML = `<table data-layout-table="true" style="width:100%; border-collapse:separate; border-spacing:8px;">${rowsHtml.join('')}</table>`;
    grid.replaceWith(wrapper);
  });

  // Convert all remaining images (Kop logos, charts) to PNG with explicit width/height HTML attributes for Word
  const imgs = Array.from(clone.querySelectorAll('img'));
  await Promise.all(
    imgs.map(async (img) => {
      if (img.hasAttribute('width') && img.hasAttribute('height') && (img.getAttribute('src') || '').startsWith('data:image/png')) {
        return;
      }
      const src = img.getAttribute('src') || '';
      if (!src) return;
      const info = await imageUrlToPngWithMeta(src);
      if (!info) return;
      img.setAttribute('src', info.dataUrl);

      const boxW = parseFloat(img.style.width || '') || 0;
      const boxH = parseFloat(img.style.height || '') || 0;
      if (boxW > 0 && boxH > 0) {
        const fitted = fitContainDimensions(info.width, info.height, boxW, boxH);
        const w = Math.round(fitted.width);
        const h = Math.round(fitted.height);
        img.setAttribute('width', String(w));
        img.setAttribute('height', String(h));
        img.style.width = `${w}px`;
        img.style.height = `${h}px`;
      } else {
        // Chart or full-width image inside Word cell
        const paperCfg = options.paperSize ? PAPER_SIZE_CONFIG[options.paperSize] : PAPER_SIZE_CONFIG.A4;
        const usableMm = (options.orientation === 'landscape' ? paperCfg.heightMm : paperCfg.widthMm) - 24;
        const usablePx = Math.round(usableMm * 3.7795);
        const isInChartCard = Boolean(img.closest('[data-chart-card="true"], td'));
        const targetW = isInChartCard ? Math.round((usablePx - 44) / (options.paperSize === 'A5' && options.orientation === 'portrait' ? 1 : 2)) : Math.min(info.width, usablePx);
        const fitted = fitContainDimensions(info.width, info.height, targetW, targetW * 2);
        const w = Math.round(fitted.width);
        const h = Math.round(fitted.height);
        img.setAttribute('width', String(w));
        img.setAttribute('height', String(h));
        img.style.width = `${w}px`;
        img.style.height = `${h}px`;
      }
    })
  );

  // Strip card frame styles on outer container and each report sheet, and insert Word page breaks between sheets
  clone.querySelectorAll('[data-report-sheet="true"]').forEach((el) => {
    const sheetEl = el as HTMLElement;
    sheetEl.style.border = 'none';
    sheetEl.style.boxShadow = 'none';
    sheetEl.style.padding = '0';
    sheetEl.style.margin = '0';
    sheetEl.style.width = '100%';
    sheetEl.style.maxWidth = 'none';
    sheetEl.style.minWidth = '0';
  });
  clone.querySelectorAll('[data-page-break="true"]').forEach((el) => {
    const br = document.createElement('br');
    br.setAttribute('clear', 'all');
    br.style.setProperty('page-break-before', 'always');
    br.style.setProperty('break-before', 'page');
    el.parentNode?.insertBefore(br, el);
  });

  clone.style.border = 'none';
  clone.style.boxShadow = 'none';
  clone.style.padding = '0';
  clone.style.margin = '0';
  clone.style.width = '100%';
  clone.style.maxWidth = 'none';
  clone.style.minWidth = '0';

  const wordPageSize = options.paperSize
    ? getWordPageSizeCss(options.paperSize, options.orientation, options.cssPageSize)
    : `${options.cssPageSize} ${options.orientation}`;
  const bodyFontCss = getPrintFontCss(options.fontFamily);

  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8" />
<title>${escapeHtml(options.title)}</title>
<style>
  @page { size: ${wordPageSize}; mso-page-orientation: ${options.orientation}; margin: 12mm; }
  body { font-family: ${bodyFontCss}; color: #000000; margin: 0; padding: 0; background: #ffffff; }
  h1, h2, h3, h4, p { margin: 0; padding: 0; }
  table { border-collapse: collapse; }
  table[data-layout-table="true"] td, table[data-layout-table="true"] th { border: none !important; background: transparent !important; }
  table.report-data-table { width: 100%; border-collapse: collapse; }
  table.report-data-table thead { display: table-header-group; }
  table.report-data-table tr { break-inside: avoid; page-break-inside: avoid; }
  table.report-data-table th, table.report-data-table td { border: 1px solid #000000; padding: 4px 6px; font-size: 8pt; }
  table.report-data-table th { background-color: #f1f5f9; font-weight: bold; }
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
