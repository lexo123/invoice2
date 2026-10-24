import ExcelJS from 'exceljs';
import {
  CellBorderSide,
  CurrencySymbol,
  EmbeddedSheetImage,
  GeneratedInvoicePackage,
  ParsedSheetCell,
  ParsedSheetModel,
} from '../types/invoice';
import {
  extractThemePaletteFromXml,
  resolveExcelColor,
} from './excelColorUtils';
import {
  applyAffixToDisplayValue,
  colNumberToLetter,
  expandRange,
  formatCurrencyCellValue,
  formatExcelCellValue,
  FormulaCellData,
  parseCellAddress,
  recalculateSheetCells,
} from './formulaEngine';

function extractRawCellValueAndFormula(cell: ExcelJS.Cell): {
  value: string | number;
  formula?: string;
} {
  const raw = cell.value;
  if (raw === null || raw === undefined) {
    return { value: '' };
  }
  if (typeof raw === 'number' || typeof raw === 'string') {
    return { value: raw };
  }
  if (typeof raw === 'boolean') {
    return { value: raw ? 'TRUE' : 'FALSE' };
  }
  if (raw instanceof Date) {
    return { value: raw.toLocaleDateString('ka-GE') };
  }
  if (typeof raw === 'object') {
    if ('formula' in raw && typeof raw.formula === 'string') {
      const res =
        'result' in raw && (typeof raw.result === 'number' || typeof raw.result === 'string')
          ? raw.result
          : 0;
      return { value: res, formula: raw.formula };
    }
    if ('sharedFormula' in raw && 'result' in raw) {
      const res =
        typeof (raw as any).result === 'number' || typeof (raw as any).result === 'string'
          ? (raw as any).result
          : '';
      return { value: res };
    }
    if ('richText' in raw && Array.isArray(raw.richText)) {
      return { value: raw.richText.map((t) => t.text || '').join('') };
    }
    if ('text' in raw && typeof raw.text === 'string') {
      return { value: raw.text };
    }
  }
  return { value: String(raw) };
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  const chunks: string[] = [];
  for (let i = 0; i < bytes.byteLength; i += chunkSize) {
    const sub = bytes.subarray(i, i + chunkSize);
    chunks.push(String.fromCharCode.apply(null, Array.from(sub)));
  }
  const binary = chunks.join('');
  return typeof btoa !== 'undefined' ? btoa(binary) : '';
}

function resolveCellFillColor(
  fill: ExcelJS.Fill | undefined,
  themePalette: string[]
): string {
  if (!fill || typeof fill !== 'object') return '';
  if (fill.type === 'pattern') {
    if (fill.pattern === 'none' || fill.pattern === 'gray125') return '';
    return (
      resolveExcelColor(fill.fgColor, themePalette, '') ||
      resolveExcelColor(fill.bgColor, themePalette, '')
    );
  }
  if (fill.type === 'gradient' && Array.isArray((fill as any).stops)) {
    const firstStop = (fill as any).stops[0];
    if (firstStop?.color) {
      return resolveExcelColor(firstStop.color, themePalette, '');
    }
  }
  return '';
}

/**
 * Parses an `.xlsx` ArrayBuffer into a complete visual sheet model for 1:1 rendering and PDF generation.
 * Uses `{ includeEmpty: true }` so colored or bordered cells without text (e.g. D36) are never skipped.
 */
export async function parseExcelBufferToSheetModel(buffer: ArrayBuffer): Promise<ParsedSheetModel> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('Excel ფაილში სამუშაო გვერდი (Worksheet) ვერ მოიძებნა.');
  }

  // Extract workbook theme palette from theme1.xml so theme colors & tints match Excel 100%
  const rawThemeXml =
    (workbook as any)?._themes?.theme1 ||
    (workbook.model as any)?.themes?.theme1 ||
    (Array.isArray((workbook.model as any)?.themes) ? (workbook.model as any).themes[0] : undefined);
  const themePalette = extractThemePaletteFromXml(rawThemeXml);

  // Determine bounding box of used cells, including empty cells that have background fill or borders!
  let maxRow = 1;
  let maxCol = 1;

  sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const rowFillColor = resolveCellFillColor(row.fill, themePalette);
    if (rowFillColor && rowNumber > maxRow) {
      maxRow = rowNumber;
    }
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const hasVal = cell.value !== null && cell.value !== undefined && cell.value !== '';
      const hasFill = Boolean(resolveCellFillColor(cell.fill, themePalette));
      const hasBorder = Boolean(
        cell.border?.top?.style ||
          cell.border?.bottom?.style ||
          cell.border?.left?.style ||
          cell.border?.right?.style
      );
      if (hasVal || hasFill || hasBorder) {
        if (rowNumber > maxRow) maxRow = rowNumber;
        if (colNumber > maxCol) maxCol = colNumber;
      }
    });
  });

  // Also include merged cell bounds and image bounds
  const mergesList: string[] =
    sheet.model && Array.isArray(sheet.model.merges) ? sheet.model.merges : [];
  for (const mergeRange of mergesList) {
    const parts = mergeRange.split(':');
    if (parts.length === 2) {
      const end = parseCellAddress(parts[1]);
      if (end) {
        if (end.row > maxRow) maxRow = end.row;
        if (end.col > maxCol) maxCol = end.col;
      }
    }
  }

  // Extract conditional formatting fills if any exist on the sheet
  const conditionalFillMap: Record<string, string> = {};
  try {
    const cfList: any[] = (sheet as any).conditionalFormattings ?? [];
    for (const cf of cfList) {
      const refStr: string = cf.ref || '';
      const rules: any[] = cf.rules ?? [];
      for (const rule of rules) {
        const cfColor = resolveCellFillColor(rule?.style?.fill, themePalette);
        if (cfColor && refStr) {
          const ranges = refStr.split(/\s+/);
          for (const rng of ranges) {
            const addrs = expandRange(rng);
            for (const a of addrs) {
              conditionalFillMap[a] = cfColor;
              const coord = parseCellAddress(a);
              if (coord) {
                if (coord.row > maxRow) maxRow = coord.row;
                if (coord.col > maxCol) maxCol = coord.col;
              }
            }
          }
        }
      }
    }
  } catch {
    // Ignore conditional formatting parse errors
  }

  try {
    const wsImages = sheet.getImages ? sheet.getImages() : [];
    for (const imgRef of wsImages) {
      const rAny = imgRef.range as any;
      const brC = Math.ceil(rAny?.br?.col ?? rAny?.br?.nativeCol ?? (rAny?.tl?.col ?? 0) + 2);
      const brR = Math.ceil(rAny?.br?.row ?? rAny?.br?.nativeRow ?? (rAny?.tl?.row ?? 0) + 3);
      if (brC > maxCol) maxCol = brC;
      if (brR > maxRow) maxRow = brR;
    }
  } catch {
    // ignore
  }

  maxRow = Math.min(Math.max(maxRow, 15), 100);
  maxCol = Math.min(Math.max(maxCol, 6), 26);

  // Build merge map from worksheet.model.merges (array of strings like "A2:C4")
  const mergeSlaveToMaster: Record<string, string> = {};
  const mergeMasterBounds: Record<
    string,
    {
      minRow: number;
      maxRow: number;
      minCol: number;
      maxCol: number;
      rowSpan: number;
      colSpan: number;
      allAddresses: string[];
    }
  > = {};

  for (const mergeRange of mergesList) {
    const parts = mergeRange.split(':');
    if (parts.length !== 2) continue;
    const start = parseCellAddress(parts[0]);
    const end = parseCellAddress(parts[1]);
    if (!start || !end) continue;

    const minCol = Math.min(start.col, end.col);
    const maxC = Math.max(start.col, end.col);
    const minRow = Math.min(start.row, end.row);
    const maxR = Math.max(start.row, end.row);

    const masterAddr = `${colNumberToLetter(minCol)}${minRow}`;
    const allAddrs = expandRange(mergeRange);

    mergeMasterBounds[masterAddr] = {
      minRow,
      maxRow: maxR,
      minCol,
      maxCol: maxC,
      rowSpan: maxR - minRow + 1,
      colSpan: maxC - minCol + 1,
      allAddresses: allAddrs,
    };

    for (const addr of allAddrs) {
      if (addr !== masterAddr) {
        mergeSlaveToMaster[addr] = masterAddr;
      }
    }
  }

  // Extract column widths in pixels (1 Excel width unit ~= 8.2px)
  const defaultColW = sheet.properties?.defaultColWidth || 11;
  const colWidthsPx: number[] = [0];
  for (let c = 1; c <= maxCol; c++) {
    const colObj = sheet.getColumn(c);
    const w = colObj?.width ?? defaultColW;
    colWidthsPx[c] = Math.max(Math.round(w * 8.2), 46);
  }

  // Extract row heights in pixels (1 pt ~= 1.35px)
  const defaultRowH = sheet.properties?.defaultRowHeight || 19;
  const rowHeightsPx: number[] = [0];
  for (let r = 1; r <= maxRow; r++) {
    const rowObj = sheet.getRow(r);
    const h = rowObj?.height ?? defaultRowH;
    rowHeightsPx[r] = Math.max(Math.round(h * 1.38), 22);
  }

  // Collect initial values & formulas
  const rawFormulaMap: Record<string, FormulaCellData> = {};
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const addr = `${colNumberToLetter(c)}${r}`;
      const cell = sheet.getCell(addr);
      rawFormulaMap[addr] = extractRawCellValueAndFormula(cell);
    }
  }
  const recalculatedMap = recalculateSheetCells(rawFormulaMap, {});

  const parseBorderSide = (side?: Partial<ExcelJS.Border>): CellBorderSide | undefined => {
    if (!side || !side.style) return undefined;
    return {
      style: side.style,
      color: resolveExcelColor(side.color, themePalette, '#CBD5E1'),
    };
  };

  const cells: Record<string, ParsedSheetCell> = {};

  for (let r = 1; r <= maxRow; r++) {
    const rowObj = sheet.getRow(r);
    const rowFillColor = resolveCellFillColor(rowObj.fill, themePalette);

    for (let c = 1; c <= maxCol; c++) {
      const addr = `${colNumberToLetter(c)}${r}`;
      const cell = sheet.getCell(addr);
      const colObj = sheet.getColumn(c);
      const colFillColor = resolveCellFillColor(colObj?.fill, themePalette);
      const recalc = recalculatedMap[addr] ?? { value: '' };

      const masterAddr = mergeSlaveToMaster[addr];
      const bounds = mergeMasterBounds[addr];
      const rowSpan = bounds ? bounds.rowSpan : 1;
      const colSpan = bounds ? bounds.colSpan : 1;

      // Font & colors
      const font = cell.font;
      const fontFamily = font?.name || 'Noto Sans Georgian';
      const fontSize = font?.size ? Math.round(font.size * 1.25) : 12;
      const bold = Boolean(font?.bold);
      const italic = Boolean(font?.italic);
      const textColor = resolveExcelColor(font?.color, themePalette, '#0F172A');

      // Fill background: check cell fill, merged slave fills, conditional formatting, row fill, col fill
      let bgColor =
        resolveCellFillColor(cell.fill, themePalette) ||
        conditionalFillMap[addr] ||
        rowFillColor ||
        colFillColor ||
        '';

      if (!bgColor && bounds) {
        for (const memberAddr of bounds.allAddresses) {
          const memberFill =
            resolveCellFillColor(sheet.getCell(memberAddr).fill, themePalette) ||
            conditionalFillMap[memberAddr];
          if (memberFill) {
            bgColor = memberFill;
            break;
          }
        }
      }

      // Alignment: Product cells (A17:A24) are always left-aligned whether they contain text or numbers
      const isProductNameCell = c === 1 && r >= 17 && r <= 24;
      const alignHorizontal = isProductNameCell
        ? 'left'
        : cell.alignment?.horizontal === 'center' ||
          cell.alignment?.horizontal === 'right' ||
          cell.alignment?.horizontal === 'left'
        ? cell.alignment.horizontal
        : typeof recalc.value === 'number'
        ? 'right'
        : 'left';

      const alignVertical =
        cell.alignment?.vertical === 'top' ||
        cell.alignment?.vertical === 'bottom' ||
        cell.alignment?.vertical === 'middle'
          ? cell.alignment.vertical
          : 'middle';

      // Borders (for merged master cells, also check bottom-right edge cells of the merge range!)
      const topBorder = parseBorderSide(cell.border?.top);
      const leftBorder = parseBorderSide(cell.border?.left);
      let rightBorder = parseBorderSide(cell.border?.right);
      let bottomBorder = parseBorderSide(cell.border?.bottom);

      if (bounds) {
        const rightEdgeCell = sheet.getCell(`${colNumberToLetter(bounds.maxCol)}${bounds.minRow}`);
        const bottomEdgeCell = sheet.getCell(`${colNumberToLetter(bounds.minCol)}${bounds.maxRow}`);
        if (!rightBorder) rightBorder = parseBorderSide(rightEdgeCell.border?.right);
        if (!bottomBorder) bottomBorder = parseBorderSide(bottomEdgeCell.border?.bottom);
      }

      const numFmt = cell.numFmt;
      const formattedValue = formatExcelCellValue(recalc.value, numFmt);

      // Ensure unmerged column width is wide enough so labels like "დოლარის კურსი" never get cut off
      if (
        !masterAddr &&
        colSpan === 1 &&
        formattedValue.length > 0 &&
        formattedValue.length <= 22 &&
        !cell.alignment?.wrapText
      ) {
        const neededPx = Math.min(Math.round(formattedValue.length * (fontSize * 0.62)) + 14, 150);
        if (neededPx > colWidthsPx[c]) {
          colWidthsPx[c] = neededPx;
        }
      }

      cells[addr] = {
        address: addr,
        row: r,
        col: c,
        value: recalc.value,
        formattedValue,
        formula: recalc.formula,
        numFmt,
        isMergedSlave: Boolean(masterAddr),
        masterAddress: masterAddr,
        rowSpan,
        colSpan,
        fontFamily,
        fontSize,
        bold,
        italic,
        textColor,
        bgColor,
        alignHorizontal,
        alignVertical,
        wrapText: Boolean(cell.alignment?.wrapText),
        borders: {
          top: topBorder,
          right: rightBorder,
          bottom: bottomBorder,
          left: leftBorder,
        },
      };
    }
  }

  // Extract embedded images from worksheet preserving exact coordinates and image buffers
  const images: EmbeddedSheetImage[] = [];
  try {
    const wsImages = sheet.getImages ? sheet.getImages() : [];
    const mediaList: any[] =
      (workbook as any).media ?? (workbook.model as any)?.media ?? [];

    for (let i = 0; i < wsImages.length; i++) {
      const imgRef = wsImages[i];
      let mediaItem: any = null;
      if (typeof workbook.getImage === 'function') {
        try {
          mediaItem = workbook.getImage(Number(imgRef.imageId));
        } catch {
          mediaItem = null;
        }
      }
      if (!mediaItem) {
        mediaItem =
          mediaList.find((m) => String(m.index) === String(imgRef.imageId)) ||
          mediaList[Number(imgRef.imageId)];
      }

      if (!mediaItem) continue;

      let dataUrl = '';
      const rawExt = String(mediaItem.extension || 'png').toLowerCase();
      const mimeSub = rawExt === 'jpg' ? 'jpeg' : rawExt;

      if (typeof mediaItem.base64 === 'string' && mediaItem.base64.length > 0) {
        dataUrl = mediaItem.base64.startsWith('data:')
          ? mediaItem.base64
          : `data:image/${mimeSub};base64,${mediaItem.base64}`;
      } else if (mediaItem.buffer) {
        const bytes =
          mediaItem.buffer instanceof Uint8Array
            ? mediaItem.buffer
            : new Uint8Array(mediaItem.buffer);
        const b64 = uint8ArrayToBase64(bytes);
        if (b64) {
          dataUrl = `data:image/${mimeSub};base64,${b64}`;
        }
      }

      if (dataUrl && imgRef.range) {
        const rangeAny = imgRef.range as any;
        const tlCol =
          typeof rangeAny.tl?.col === 'number'
            ? rangeAny.tl.col
            : rangeAny.tl?.nativeCol ?? 0;
        const tlRow =
          typeof rangeAny.tl?.row === 'number'
            ? rangeAny.tl.row
            : rangeAny.tl?.nativeRow ?? 0;
        const brCol =
          typeof rangeAny.br?.col === 'number'
            ? rangeAny.br.col
            : rangeAny.br?.nativeCol;
        const brRow =
          typeof rangeAny.br?.row === 'number'
            ? rangeAny.br.row
            : rangeAny.br?.nativeRow;
        const widthPx = rangeAny.ext?.width;
        const heightPx = rangeAny.ext?.height;

        images.push({
          id: `img-${i}-${imgRef.imageId}`,
          dataUrl,
          tlCol,
          tlRow,
          brCol,
          brRow,
          widthPx,
          heightPx,
        });
      }
    }
  } catch {
    // Ignore image extraction errors on non-standard workbooks
  }

  return {
    sheetName: sheet.name || 'ინვოისი',
    rowCount: maxRow,
    colCount: maxCol,
    colWidthsPx,
    rowHeightsPx,
    cells,
    images,
  };
}

/**
 * Generates a new `.xlsx` file from the template buffer:
 * 1. Modifies ONLY the designated editable cells with `editableOverrides`.
 * 2. Keeps all other static cells, styles, colors, merged cells, and images 100% untouched.
 * 3. Recalculates formulas across the sheet and updates their cached results in the `.xlsx` file.
 * 4. Returns the new `.xlsx` ArrayBuffer and the updated ParsedSheetModel ready for Excel or PDF download.
 */
export async function generateUpdatedExcelPackage(
  templateBuffer: ArrayBuffer,
  editableOverrides: Record<string, string | number>,
  fileNameBase = 'ინვოისი',
  affixRules?: Record<string, { text: string; position: 'start' | 'end' }>,
  currencyRules?: Record<string, { mode: 'number' | 'currency'; symbol?: CurrencySymbol }>
): Promise<GeneratedInvoicePackage> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('შაბლონის ფაილში სამუშაო გვერდი ვერ მოიძებნა.');
  }

  let maxRow = Math.min(Math.max(sheet.rowCount || 36, 36), 100);
  let maxCol = Math.min(Math.max(sheet.columnCount || 6, 6), 26);

  for (const rawAddr of Object.keys(editableOverrides)) {
    const coord = parseCellAddress(rawAddr);
    if (coord) {
      if (coord.row > maxRow && coord.row <= 100) maxRow = coord.row;
      if (coord.col > maxCol && coord.col <= 26) maxCol = coord.col;
    }
  }

  // 1. Collect all initial cells and formulas
  const initialCells: Record<string, FormulaCellData> = {};
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const addr = `${colNumberToLetter(c)}${r}`;
      const cell = sheet.getCell(addr);
      initialCells[addr] = extractRawCellValueAndFormula(cell);
    }
  }

  // 2. Run our deterministic formula recalculation engine with the user's editable overrides
  const recalculated = recalculateSheetCells(initialCells, editableOverrides);

  // 3. Apply editable cell values (including user-entered formulas!) to the workbook
  let changedCellsCount = 0;
  for (const [rawAddr, newVal] of Object.entries(editableOverrides)) {
    const coord = parseCellAddress(rawAddr);
    if (!coord) continue;
    const cell = sheet.getCell(coord.address);
    const calcEntry = recalculated[coord.address];
    if (calcEntry && calcEntry.formula) {
      cell.value = {
        formula: calcEntry.formula,
        result: calcEntry.value,
      };
    } else {
      const finalVal = calcEntry ? calcEntry.value : newVal;
      cell.value = finalVal === '' ? null : finalVal;
      if (
        (coord.col === 1 && coord.row >= 17 && coord.row <= 24) ||
        (typeof finalVal === 'string' && finalVal !== '' && !cell.alignment?.horizontal)
      ) {
        cell.alignment = {
          ...(cell.alignment || {}),
          horizontal: 'left',
        };
      }
    }

    // Apply number vs currency numFmt in the Excel file if specified
    const cRule = currencyRules?.[coord.address.toUpperCase()];
    if (cRule) {
      if (cRule.mode === 'number') {
        cell.numFmt = '0.####';
      } else if (cRule.mode === 'currency') {
        const sym = cRule.symbol || '₾';
        const existingDec = cell.numFmt?.match(/\.(0+)/)?.[1];
        cell.numFmt = existingDec ? `#,##0.${existingDec} "${sym}"` : `#,##0.## "${sym}"`;
      }
    }

    changedCellsCount++;
  }

  // Ensure all cells in currencyRules have their Excel numFmt updated (even if not in editableOverrides)
  if (currencyRules) {
    for (const [rawAddr, cRule] of Object.entries(currencyRules)) {
      const coord = parseCellAddress(rawAddr);
      if (!coord) continue;
      const normAddr = coord.address.toUpperCase();
      const cell = sheet.getCell(normAddr);
      const affix = affixRules?.[normAddr];
      if (cRule.mode === 'number') {
        if (affix && affix.text.trim() !== '') {
          const safeAffix = affix.text.trim().replace(/"/g, '');
          cell.numFmt =
            affix.position === 'end'
              ? `0.#### "${safeAffix}"`
              : `"${safeAffix} "0.####`;
        } else {
          cell.numFmt = '0.####';
        }
      } else if (cRule.mode === 'currency') {
        const sym = cRule.symbol || '₾';
        const existingDec = cell.numFmt?.match(/\.(0+)/)?.[1];
        cell.numFmt = existingDec
          ? `#,##0.${existingDec} "${sym}"`
          : `#,##0.## "${sym}"`;
      }
    }
  }

  // 4. Write recalculated results into template formula cells
  let recalculatedFormulasCount = 0;
  for (const [addr, entry] of Object.entries(recalculated)) {
    if (entry.formula) {
      const cell = sheet.getCell(addr);
      cell.value = {
        formula: entry.formula,
        result: entry.value,
      };
      recalculatedFormulasCount++;
    }
  }

  // 5. Write the newly generated Excel workbook to an ArrayBuffer
  const outBuffer = (await workbook.xlsx.writeBuffer()) as ArrayBuffer;

  // 6. Parse the newly generated Excel buffer so our preview & PDF match 100%
  const updatedSheetModel = await parseExcelBufferToSheetModel(outBuffer);

  // Apply explicit 'number' vs 'currency' display rules to updatedSheetModel
  if (currencyRules) {
    for (const [addr, cRule] of Object.entries(currencyRules)) {
      const norm = addr.toUpperCase();
      const targetCell = updatedSheetModel.cells[norm];
      if (
        !targetCell ||
        targetCell.value === '' ||
        targetCell.value === null ||
        targetCell.value === undefined
      ) {
        continue;
      }
      if (cRule.mode === 'number') {
        targetCell.formattedValue = formatExcelCellValue(
          targetCell.value,
          'General'
        );
      } else if (cRule.mode === 'currency') {
        targetCell.formattedValue = formatCurrencyCellValue(
          targetCell.value,
          cRule.symbol || '₾',
          targetCell.numFmt
        );
      }
    }
  }

  // If any cell has a configured start/end word (e.g. "ლარი" at the end), apply it to formattedValue
  if (affixRules) {
    for (const [addr, rule] of Object.entries(affixRules)) {
      const norm = addr.toUpperCase();
      const targetCell = updatedSheetModel.cells[norm];
      if (targetCell && targetCell.formattedValue) {
        targetCell.formattedValue = applyAffixToDisplayValue(
          targetCell.formattedValue,
          true,
          rule.text,
          rule.position
        );
      }
    }
  }

  const rawInvNum = String(
    updatedSheetModel.cells['C8']?.formattedValue ||
      updatedSheetModel.cells['F3']?.formattedValue ||
      'INV'
  )
    .replace(/^ინვოისის ნომერი:?\s*/i, '')
    .trim();
  const invNum = rawInvNum || 'INV';
  const clientName = String(
    updatedSheetModel.cells['A8']?.formattedValue ||
      updatedSheetModel.cells['C7']?.formattedValue ||
      ''
  );
  const subtotalFormatted = String(
    updatedSheetModel.cells['D36']?.formattedValue ||
      updatedSheetModel.cells['F16']?.formattedValue ||
      '—'
  );
  const vatFormatted = String(updatedSheetModel.cells['F17']?.formattedValue || '—');
  const grandTotalFormatted = String(
    updatedSheetModel.cells['D36']?.formattedValue ||
      updatedSheetModel.cells['F18']?.formattedValue ||
      '—'
  );

  const cleanFileName = String(fileNameBase || '').trim() || 'ინვოისი';

  return {
    xlsxBuffer: outBuffer,
    fileNameBase: cleanFileName,
    generatedAt: new Date().toLocaleTimeString('ka-GE', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    }),
    sheetModel: updatedSheetModel,
    summaryMetrics: {
      invoiceNumber: invNum,
      clientName,
      subtotalFormatted,
      vatFormatted,
      grandTotalFormatted,
      changedCellsCount,
      recalculatedFormulasCount,
    },
  };
}
