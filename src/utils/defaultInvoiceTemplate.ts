import ExcelJS from 'exceljs';
import { EditableCellInstruction } from '../types/invoice';

/**
 * Default editable cell instructions for the single-page Georgian Invoice template.
 * Matches the Main Page layout:
 * - 6 Header/Client fields: კომპანიის სახელი (A8), საკადასტრო (A9), მისამართი (A10),
 *   ინვოისის ნომერი (C8), თარიღი (C9), დოლარის კურსი (C10)
 * - Up to 8 Product rows starting at A17, B17, C17 (through A24, B24, C24)
 * - Grand total in Lari (₾) at colored cell D36
 */
export const DEFAULT_EDITABLE_INSTRUCTIONS: EditableCellInstruction[] = [
  {
    id: 'main-companyName',
    address: 'A8',
    label: 'კომპანიის სახელი',
    group: 'client',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'კომპანიის სახელი',
    affixPosition: 'start',
    mainFieldKey: 'companyName',
  },
  {
    id: 'main-cadastralCode',
    address: 'A9',
    label: 'საკადასტრო',
    group: 'client',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'საკადასტრო',
    affixPosition: 'start',
    mainFieldKey: 'cadastralCode',
  },
  {
    id: 'main-address',
    address: 'A10',
    label: 'მისამართი',
    group: 'client',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'მისამართი',
    affixPosition: 'start',
    mainFieldKey: 'address',
  },
  {
    id: 'main-invoiceNumber',
    address: 'C8',
    label: 'ინვოისის ნომერი',
    group: 'header',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'ინვოისის ნომერი',
    affixPosition: 'start',
    mainFieldKey: 'invoiceNumber',
  },
  {
    id: 'main-date',
    address: 'C9',
    label: 'თარიღი',
    group: 'header',
    inputType: 'date',
    placeholder: '',
    usePrefix: true,
    prefixText: 'თარიღი',
    affixPosition: 'start',
    mainFieldKey: 'date',
  },
  {
    id: 'main-usdRate',
    address: 'C10',
    label: 'დოლარის კურსი',
    group: 'header',
    inputType: 'number',
    placeholder: '',
    usePrefix: true,
    prefixText: 'დოლარის კურსი',
    affixPosition: 'start',
    mainFieldKey: 'usdRate',
  },
  {
    id: 'item1-name',
    address: 'A17',
    label: 'პროდუქტი #1 — დასახელება',
    group: 'items',
    inputType: 'text',
  },
  {
    id: 'item1-qty',
    address: 'B17',
    label: 'პროდუქტი #1 — რაოდენობა',
    group: 'items',
    inputType: 'number',
  },
  {
    id: 'item1-price',
    address: 'C17',
    label: 'პროდუქტი #1 — ერთეულის ფასი',
    group: 'items',
    inputType: 'currency',
    currencySymbol: '₾',
  },
  {
    id: 'grand-total-d36',
    address: 'D36',
    label: 'სულ ჯამი (ლარში)',
    group: 'other',
    inputType: 'formula',
    formulaExpression: '=SUM(D17:D24)',
    formulaResultFormat: 'currency',
    currencySymbol: '₾',
  },
];

/**
 * Generates crisp PNG base64 images via offscreen canvas so the default Excel (.xlsx)
 * file contains real embedded images (company emblem & official stamp) that survive
 * round-trip Excel export and PDF rendering.
 */
function createCompanyLogoPngBase64(): string {
  if (typeof document === 'undefined') return '';
  const canvas = document.createElement('canvas');
  canvas.width = 240;
  canvas.height = 70;
  const ctx = canvas.getContext('2d');
  if (!ctx) return '';

  ctx.fillStyle = '#0F172A';
  ctx.beginPath();
  ctx.roundRect(4, 4, 58, 58, 12);
  ctx.fill();

  ctx.strokeStyle = '#38BDF8';
  ctx.lineWidth = 3.5;
  ctx.beginPath();
  ctx.moveTo(18, 46);
  ctx.lineTo(33, 18);
  ctx.lineTo(48, 46);
  ctx.stroke();

  ctx.fillStyle = '#F8FAFC';
  ctx.beginPath();
  ctx.arc(33, 38, 4, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = '#0F172A';
  ctx.font = 'bold 19px sans-serif';
  ctx.fillText('ALPHA TECH', 74, 31);

  ctx.fillStyle = '#475569';
  ctx.font = '600 11px monospace';
  ctx.fillText('ID: 405123987 • LLC', 74, 50);

  return canvas.toDataURL('image/png');
}

function createCompanyStampPngBase64(): string {
  if (typeof document === 'undefined') return '';
  const canvas = document.createElement('canvas');
  canvas.width = 180;
  canvas.height = 90;
  const ctx = canvas.getContext('2d');
  if (!ctx) return '';

  ctx.save();
  ctx.translate(90, 45);
  ctx.rotate(-0.06);

  ctx.strokeStyle = 'rgba(29, 78, 216, 0.8)';
  ctx.lineWidth = 2.2;
  ctx.strokeRect(-78, -30, 156, 60);

  ctx.lineWidth = 1;
  ctx.strokeRect(-72, -24, 144, 48);

  ctx.fillStyle = 'rgba(29, 78, 216, 0.88)';
  ctx.textAlign = 'center';
  ctx.font = 'bold 10px sans-serif';
  ctx.fillText('ALPHA TECH LLC', 0, -8);

  ctx.font = 'bold 11px monospace';
  ctx.fillText('ID: 405123987', 0, 6);

  ctx.font = '600 8.5px sans-serif';
  ctx.fillText('AUTHORIZED STAMP • GE', 0, 19);

  ctx.restore();
  return canvas.toDataURL('image/png');
}

/**
 * Builds the default single-page Georgian Invoice `.xlsx` file as an ArrayBuffer.
 * Structured with:
 * - Header fields: კომპანიის სახელი (A8), საკადასტრო (A9), მისამართი (A10),
 *   ინვოისის ნომერი (C8), თარიღი (C9), დოლარის კურსი (C10)
 * - Table header at row 16 (A16:D16)
 * - 8 Product rows at rows 17..24 (A17:C17 to A24:C24, with row totals in D17:D24)
 * - Colored Grand Total cell at D36
 */
export async function createDefaultInvoiceTemplateBuffer(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'InvoExcel';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('ინვოისი', {
    pageSetup: {
      paperSize: 9, // A4
      orientation: 'portrait',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
    },
    views: [{ showGridLines: true }],
  });

  // 4 well-proportioned columns A, B, C, D for single-page A4 layout
  sheet.columns = [
    { key: 'A', width: 36 }, // პროდუქტი / დასახელება
    { key: 'B', width: 15 }, // რაოდენობა
    { key: 'C', width: 18 }, // ერთეულის ფასი
    { key: 'D', width: 18 }, // ჯამი (₾)
  ];

  for (let r = 1; r <= 36; r++) {
    if (r === 2) sheet.getRow(r).height = 26;
    else if (r >= 8 && r <= 10) sheet.getRow(r).height = 22;
    else if (r === 16) sheet.getRow(r).height = 25;
    else if (r >= 17 && r <= 24) sheet.getRow(r).height = 22;
    else if (r === 36) sheet.getRow(r).height = 26;
    else sheet.getRow(r).height = 16;
  }

  // Top Header & Title (Rows 2 to 5)
  sheet.mergeCells('A2:B4');
  sheet.mergeCells('C2:D3');
  const titleCell = sheet.getCell('C2');
  titleCell.value = 'ინვოისი / INVOICE';
  titleCell.font = {
    name: 'Noto Sans Georgian',
    size: 15,
    bold: true,
    color: { argb: 'FF0F172A' },
  };
  titleCell.alignment = { horizontal: 'right', vertical: 'middle' };

  sheet.mergeCells('A5:D5');
  const sellerTag = sheet.getCell('A5');
  sellerTag.value =
    'მიმწოდებელი: შპს ალფა ტექნოლოგიები (ს/კ: 405123987) • სს თიბისი ბანკი: GE29TB7788990011223344';
  sellerTag.font = {
    name: 'Noto Sans Georgian',
    size: 9,
    bold: true,
    color: { argb: 'FF334155' },
  };
  sellerTag.alignment = { horizontal: 'left', vertical: 'middle' };

  // Section Caption (Row 7)
  sheet.mergeCells('A7:B7');
  const buyerHdr = sheet.getCell('A7');
  buyerHdr.value = 'მყიდველის მონაცემები';
  buyerHdr.font = {
    name: 'Noto Sans Georgian',
    size: 9,
    bold: true,
    color: { argb: 'FF475569' },
  };

  sheet.mergeCells('C7:D7');
  const metaHdr = sheet.getCell('C7');
  metaHdr.value = 'ინვოისის რეკვიზიტები';
  metaHdr.font = {
    name: 'Noto Sans Georgian',
    size: 9,
    bold: true,
    color: { argb: 'FF475569' },
  };

  // 6 Main Editable Header Cells (Rows 8, 9, 10)
  const leftFields: [number, string][] = [
    [8, 'კომპანიის სახელი'],
    [9, 'საკადასტრო'],
    [10, 'მისამართი'],
  ];

  for (const [r, val] of leftFields) {
    sheet.mergeCells(`A${r}:B${r}`);
    const cell = sheet.getCell(`A${r}`);
    cell.value = val;
    cell.font = {
      name: 'Noto Sans Georgian',
      size: 10,
      bold: r === 8,
      color: { argb: 'FF0F172A' },
    };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF8FAFC' },
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    };
  }

  const rightFields: [number, string][] = [
    [8, 'ინვოისის ნომერი'],
    [9, 'თარიღი'],
    [10, 'დოლარის კურსი 2.6'],
  ];

  for (const [r, val] of rightFields) {
    sheet.mergeCells(`C${r}:D${r}`);
    const cell = sheet.getCell(`C${r}`);
    cell.value = val;
    cell.font = {
      name: 'Noto Sans Georgian',
      size: 10,
      bold: r === 8,
      color: { argb: 'FF0F172A' },
    };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF1F5F9' },
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    };
  }

  // Table Header at Row 16 (so Product #1 starts at A17, B17, C17)
  const headers: [string, string, 'left' | 'center' | 'right'][] = [
    ['A16', 'პროდუქტი / დასახელება', 'left'],
    ['B16', 'რაოდენობა', 'center'],
    ['C16', 'ერთეულის ფასი', 'right'],
    ['D16', 'ჯამი (₾)', 'right'],
  ];

  for (const [addr, text, align] of headers) {
    const cell = sheet.getCell(addr);
    cell.value = text;
    cell.font = {
      name: 'Noto Sans Georgian',
      size: 10,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' },
    };
    cell.alignment = { horizontal: align, vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF0F172A' } },
      bottom: { style: 'thin', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF334155' } },
      right: { style: 'thin', color: { argb: 'FF334155' } },
    };
  }

  // 8 Product Rows: Rows 17 to 24 (A17:C17 through A24:C24, with D17:D24 row totals)
  for (let r = 17; r <= 24; r++) {
    const isFirst = r === 17;
    const rowBg = r % 2 === 1 ? 'FFFFFFFF' : 'FFF8FAFC';

    const cA = sheet.getCell(`A${r}`);
    cA.value = isFirst ? 'ვებ-პლატფორმის არქიტექტურა და დიზაინი' : '';
    cA.alignment = { horizontal: 'left', vertical: 'middle' };
    cA.font = {
      name: 'Noto Sans Georgian',
      size: 10,
      color: { argb: 'FF0F172A' },
    };

    const cB = sheet.getCell(`B${r}`);
    cB.value = isFirst ? 1 : '';
    cB.alignment = { horizontal: 'center', vertical: 'middle' };
    cB.font = {
      name: 'IBM Plex Mono',
      size: 10,
      bold: true,
      color: { argb: 'FF0F172A' },
    };

    const cC = sheet.getCell(`C${r}`);
    cC.value = isFirst ? 166 : '';
    cC.alignment = { horizontal: 'right', vertical: 'middle' };
    cC.font = {
      name: 'IBM Plex Mono',
      size: 10,
      color: { argb: 'FF0F172A' },
    };
    cC.numFmt = '#,##0.## "₾"';

    const cD = sheet.getCell(`D${r}`);
    if (isFirst) {
      cD.value = { formula: `B${r}*C${r}`, result: 166 };
    } else {
      cD.value = '';
    }
    cD.alignment = { horizontal: 'right', vertical: 'middle' };
    cD.font = {
      name: 'IBM Plex Mono',
      size: 10,
      bold: true,
      color: { argb: 'FF0F172A' },
    };
    cD.numFmt = '#,##0.## "₾"';

    for (const colLetter of ['A', 'B', 'C', 'D']) {
      const cell = sheet.getCell(`${colLetter}${r}`);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: rowBg },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      };
    }
  }

  // Terms & Stamp Block in Rows 27 to 32
  sheet.mergeCells('A27:B27');
  const termsHdr = sheet.getCell('A27');
  termsHdr.value = 'გადახდის პირობები და საბანკო რეკვიზიტები:';
  termsHdr.font = {
    name: 'Noto Sans Georgian',
    size: 9,
    bold: true,
    color: { argb: 'FF0F172A' },
  };

  sheet.mergeCells('A28:B28');
  sheet.getCell('A28').value = 'ბანკი: სს თიბისი ბანკი • კოდი: TBCBGE22';
  sheet.getCell('A28').font = {
    name: 'Noto Sans Georgian',
    size: 9,
    color: { argb: 'FF334155' },
  };

  sheet.mergeCells('A29:B29');
  sheet.getCell('A29').value = 'IBAN: GE29TB7788990011223344';
  sheet.getCell('A29').font = {
    name: 'IBM Plex Mono',
    size: 9,
    bold: true,
    color: { argb: 'FF0F172A' },
  };

  sheet.mergeCells('C27:D27');
  const sigCell = sheet.getCell('C27');
  sigCell.value = 'უფლებამოსილი პირის ხელმოწერა და ბეჭედი:';
  sigCell.font = {
    name: 'Noto Sans Georgian',
    size: 9,
    bold: true,
    color: { argb: 'FF475569' },
  };
  sigCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 36: Grand Total Row (D36 is colored and holds the total in Lari ₾)
  sheet.mergeCells('A36:C36');
  const totalLabel = sheet.getCell('A36');
  totalLabel.value = 'სულ გადასახდელი ჯამი (ლარში):';
  totalLabel.alignment = { horizontal: 'right', vertical: 'middle' };
  totalLabel.font = {
    name: 'Noto Sans Georgian',
    size: 11,
    bold: true,
    color: { argb: 'FF0F172A' },
  };
  totalLabel.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF1F5F9' },
  };
  totalLabel.border = {
    top: { style: 'thin', color: { argb: 'FF94A3B8' } },
    bottom: { style: 'double', color: { argb: 'FF0F172A' } },
    left: { style: 'thin', color: { argb: 'FF94A3B8' } },
    right: { style: 'thin', color: { argb: 'FF94A3B8' } },
  };

  const d36Cell = sheet.getCell('D36');
  d36Cell.value = { formula: 'SUM(D17:D24)', result: 166 };
  d36Cell.alignment = { horizontal: 'right', vertical: 'middle' };
  d36Cell.numFmt = '#,##0.## "₾"';
  d36Cell.font = {
    name: 'IBM Plex Mono',
    size: 11,
    bold: true,
    color: { argb: 'FF0F172A' },
  };
  d36Cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFDBEAFE' }, // Colored D36 cell
  };
  d36Cell.border = {
    top: { style: 'thin', color: { argb: 'FF0284C7' } },
    bottom: { style: 'double', color: { argb: 'FF0F172A' } },
    left: { style: 'thin', color: { argb: 'FF0284C7' } },
    right: { style: 'thin', color: { argb: 'FF0284C7' } },
  };

  // Embed real PNG images (Logo in A2:B4 and Stamp in C28:D32)
  const logoDataUrl = createCompanyLogoPngBase64();
  if (logoDataUrl) {
    const logoId = workbook.addImage({
      base64: logoDataUrl,
      extension: 'png',
    });
    sheet.addImage(logoId, {
      tl: { col: 0.08, row: 1.1 },
      ext: { width: 210, height: 60 },
    });
  }

  const stampDataUrl = createCompanyStampPngBase64();
  if (stampDataUrl) {
    const stampId = workbook.addImage({
      base64: stampDataUrl,
      extension: 'png',
    });
    sheet.addImage(stampId, {
      tl: { col: 2.35, row: 27.2 },
      ext: { width: 165, height: 78 },
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer as ArrayBuffer;
}
