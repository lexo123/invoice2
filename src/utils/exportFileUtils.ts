import { toPng } from 'html-to-image';
import { jsPDF } from 'jspdf';

/**
 * Triggers browser download of the generated `.xlsx` ArrayBuffer.
 */
export function downloadExcelBuffer(buffer: ArrayBuffer, fileNameBase: string): void {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileNameBase.endsWith('.xlsx') ? fileNameBase : `${fileNameBase}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/**
 * Generates an exact 1:1 single-page A4 PDF from the rendered Excel sheet container
 * preserving cell layout, merged cells, borders, colors, embedded images, and Georgian fonts.
 */
export async function exportSheetElementToPdf(
  sheetElement: HTMLElement,
  fileNameBase: string
): Promise<void> {
  const rect = sheetElement.getBoundingClientRect();
  const widthPx = rect.width || sheetElement.scrollWidth || 760;
  const heightPx = rect.height || sheetElement.scrollHeight || 980;

  const dataUrl = await toPng(sheetElement, {
    quality: 1.0,
    pixelRatio: 2.5,
    backgroundColor: '#FFFFFF',
    width: widthPx,
    height: heightPx,
    style: {
      transform: 'scale(1)',
      transformOrigin: 'top left',
    },
  });

  // A4 dimensions in mm: 210 x 297
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
    compress: true,
  });

  const pageWidthMm = 210;
  const pageHeightMm = 297;
  const marginMm = 10;

  const availWidthMm = pageWidthMm - marginMm * 2;
  const availHeightMm = pageHeightMm - marginMm * 2;

  const aspect = widthPx / heightPx;
  let renderWidthMm = availWidthMm;
  let renderHeightMm = renderWidthMm / aspect;

  // Ensure it always fits onto a single A4 page ("ძირითადად ერთგვერდიანი ინვოისებისთვის")
  if (renderHeightMm > availHeightMm) {
    renderHeightMm = availHeightMm;
    renderWidthMm = renderHeightMm * aspect;
  }

  const offsetX = (pageWidthMm - renderWidthMm) / 2;
  const offsetY = marginMm;

  pdf.addImage(dataUrl, 'PNG', offsetX, offsetY, renderWidthMm, renderHeightMm, undefined, 'FAST');
  const fullFileName = fileNameBase.endsWith('.pdf') ? fileNameBase : `${fileNameBase}.pdf`;
  pdf.save(fullFileName);
}
