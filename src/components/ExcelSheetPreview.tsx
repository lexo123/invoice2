import React from 'react';
import {
  CellBorderSide,
  EditableCellInstruction,
  ParsedSheetCell,
  ParsedSheetModel,
} from '../types/invoice';
import { colNumberToLetter } from '../utils/formulaEngine';

interface ExcelSheetPreviewProps {
  sheetModel: ParsedSheetModel;
  editableInstructions: EditableCellInstruction[];
  selectedAddress: string | null;
  onSelectCell: (address: string) => void;
  showExcelGridHeaders: boolean;
  highlightEditableCells: boolean;
  isExportingPdf?: boolean;
  sheetPrintRef: React.RefObject<HTMLDivElement | null>;
}

function borderSideToCss(side?: CellBorderSide, showDefaultGrid = false): string {
  if (!side || !side.style) {
    return showDefaultGrid ? '1px solid #F1F5F9' : 'none';
  }
  const color = side.color || '#CBD5E1';
  if (side.style === 'double') {
    return `3px double ${color}`;
  }
  if (side.style === 'medium' || side.style === 'thick') {
    return `2px solid ${color}`;
  }
  if (side.style === 'dashed' || side.style === 'dotted') {
    return `1px dashed ${color}`;
  }
  return `1px solid ${color}`;
}

export const ExcelSheetPreview: React.FC<ExcelSheetPreviewProps> = ({
  sheetModel,
  editableInstructions,
  selectedAddress,
  onSelectCell,
  showExcelGridHeaders,
  highlightEditableCells,
  isExportingPdf = false,
  sheetPrintRef,
}) => {
  const editableMap = React.useMemo(() => {
    const map: Record<string, EditableCellInstruction> = {};
    for (const item of editableInstructions) {
      map[item.address.toUpperCase()] = item;
    }
    return map;
  }, [editableInstructions]);

  // Compute cumulative X and Y pixel offsets for positioning embedded Excel images accurately
  const colOffsetsPx = React.useMemo(() => {
    const offsets: number[] = [0];
    let acc = 0;
    for (let c = 1; c <= sheetModel.colCount; c++) {
      offsets[c - 1] = acc;
      acc += sheetModel.colWidthsPx[c] || 80;
    }
    offsets[sheetModel.colCount] = acc;
    return offsets;
  }, [sheetModel]);

  const rowOffsetsPx = React.useMemo(() => {
    const offsets: number[] = [0];
    let acc = 0;
    for (let r = 1; r <= sheetModel.rowCount; r++) {
      offsets[r - 1] = acc;
      acc += sheetModel.rowHeightsPx[r] || 24;
    }
    offsets[sheetModel.rowCount] = acc;
    return offsets;
  }, [sheetModel]);

  const totalTableWidth = colOffsetsPx[sheetModel.colCount] || 720;

  const interpolateOffset = (offsets: number[], indexFloat: number): number => {
    const baseIdx = Math.floor(Math.max(0, indexFloat));
    const frac = indexFloat - baseIdx;
    const startPx = offsets[baseIdx] ?? offsets[offsets.length - 1] ?? 0;
    const nextPx = offsets[baseIdx + 1] ?? startPx + 60;
    return Math.round(startPx + (nextPx - startPx) * frac);
  };

  const rowsArray = Array.from({ length: sheetModel.rowCount }, (_, i) => i + 1);
  const colsArray = Array.from({ length: sheetModel.colCount }, (_, i) => i + 1);

  const showInteractiveOverlays = !isExportingPdf;

  return (
    <div className="inline-block bg-white border border-slate-200 rounded-lg shadow-xs overflow-hidden">
      {/* Optional Top Coordinate Bar (A, B, C, D...) outside the PDF print target */}
      {showExcelGridHeaders && (
        <div className="flex bg-slate-100 border-b border-slate-300 select-none text-[11px] font-mono text-slate-600">
          <div
            style={{ width: 36, minWidth: 36 }}
            className="h-6 flex items-center justify-center border-r border-slate-300 bg-slate-200/70 font-semibold"
          >
            ◢
          </div>
          {colsArray.map((colIdx) => {
            const w = sheetModel.colWidthsPx[colIdx] || 80;
            const letter = colNumberToLetter(colIdx);
            return (
              <div
                key={`col-hdr-${colIdx}`}
                style={{ width: w, minWidth: w }}
                className="h-6 flex items-center justify-center border-r border-slate-200 font-medium"
              >
                {letter}
              </div>
            );
          })}
        </div>
      )}

      <div className="flex">
        {/* Optional Left Row Numbers (1, 2, 3...) outside the PDF print target */}
        {showExcelGridHeaders && (
          <div
            style={{ width: 36, minWidth: 36 }}
            className="bg-slate-100 border-r border-slate-300 select-none text-[11px] font-mono text-slate-500"
          >
            {rowsArray.map((rowIdx) => {
              const h = sheetModel.rowHeightsPx[rowIdx] || 24;
              return (
                <div
                  key={`row-hdr-${rowIdx}`}
                  style={{ height: h }}
                  className="flex items-center justify-center border-b border-slate-200"
                >
                  {rowIdx}
                </div>
              );
            })}
          </div>
        )}

        {/* Exact 1:1 Printable Invoice Sheet Area (#invoice-sheet-print-target) */}
        <div
          id="invoice-sheet-print-target"
          ref={sheetPrintRef}
          style={{ width: totalTableWidth + 32 }}
          className="relative bg-white p-4"
        >
          <div className="relative" style={{ width: totalTableWidth }}>
            {/* Embedded Worksheet Images (Company Logo, Official Stamp, etc.) */}
            {sheetModel.images.map((img) => {
              const leftPx = interpolateOffset(colOffsetsPx, img.tlCol);
              const topPx = interpolateOffset(rowOffsetsPx, img.tlRow);
              const widthPx =
                img.widthPx ||
                (img.brCol !== undefined
                  ? Math.max(40, interpolateOffset(colOffsetsPx, img.brCol) - leftPx)
                  : 160);
              const heightPx =
                img.heightPx ||
                (img.brRow !== undefined
                  ? Math.max(30, interpolateOffset(rowOffsetsPx, img.brRow) - topPx)
                  : 64);

              return (
                <img
                  key={img.id}
                  src={img.dataUrl}
                  alt="Excel embedded graphic"
                  referrerPolicy="no-referrer"
                  style={{
                    position: 'absolute',
                    left: `${leftPx}px`,
                    top: `${topPx}px`,
                    width: `${widthPx}px`,
                    height: `${heightPx}px`,
                    objectFit: 'contain',
                    pointerEvents: 'none',
                    zIndex: 10,
                  }}
                />
              );
            })}

            {/* Exact Cell Grid Table — No text truncation (...) so all words and company names render in full */}
            <table
              style={{
                width: totalTableWidth,
                borderCollapse: 'collapse',
                tableLayout: 'fixed',
              }}
            >
              <colgroup>
                {colsArray.map((colIdx) => (
                  <col
                    key={`col-width-${colIdx}`}
                    style={{ width: `${sheetModel.colWidthsPx[colIdx] || 80}px` }}
                  />
                ))}
              </colgroup>
              <tbody>
                {rowsArray.map((rowIdx) => {
                  const rowHeight = sheetModel.rowHeightsPx[rowIdx] || 24;
                  return (
                    <tr key={`row-${rowIdx}`} style={{ height: `${rowHeight}px` }}>
                      {colsArray.map((colIdx) => {
                        const addr = `${colNumberToLetter(colIdx)}${rowIdx}`;
                        const cell: ParsedSheetCell | undefined = sheetModel.cells[addr];
                        if (!cell || cell.isMergedSlave) {
                          return null;
                        }

                        const isEditable = Boolean(editableMap[addr]);
                        const isFormula = Boolean(cell.formula);
                        const isSelected = showInteractiveOverlays && selectedAddress === addr;
                        const isRightAligned = cell.alignHorizontal === 'right';

                        const bgStyle = cell.bgColor || 'transparent';
                        const showGridFallback = showInteractiveOverlays && showExcelGridHeaders;

                        return (
                          <td
                            key={addr}
                            rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                            colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                            onClick={() => onSelectCell(addr)}
                            title={
                              isEditable
                                ? `${addr} — შესაცვლელი უჯრა (${editableMap[addr].label})`
                                : isFormula
                                ? `${addr} — ფორმულა: =${cell.formula}`
                                : `${addr} — უცვლელი უჯრა`
                            }
                            style={{
                              minHeight: `${rowHeight}px`,
                              height: `${rowHeight}px`,
                              backgroundColor: bgStyle,
                              color: cell.textColor || '#0F172A',
                              fontFamily:
                                cell.fontFamily?.includes('Mono') || typeof cell.value === 'number'
                                  ? 'var(--font-mono)'
                                  : 'var(--font-sans)',
                              fontSize: `${cell.fontSize || 11}px`,
                              fontWeight: cell.bold ? 700 : 400,
                              fontStyle: cell.italic ? 'italic' : 'normal',
                              textAlign: cell.alignHorizontal || 'left',
                              verticalAlign: cell.alignVertical || 'middle',
                              borderTop: borderSideToCss(cell.borders.top, showGridFallback),
                              borderRight: borderSideToCss(cell.borders.right, showGridFallback),
                              borderBottom: borderSideToCss(cell.borders.bottom, showGridFallback),
                              borderLeft: borderSideToCss(cell.borders.left, showGridFallback),
                              padding: '2px 6px',
                              position: 'relative',
                              cursor: 'pointer',
                              whiteSpace: 'pre-wrap',
                              wordBreak: 'break-word',
                              overflowWrap: 'break-word',
                              lineHeight: 1.25,
                              fontVariantNumeric: 'tabular-nums',
                              outline: isSelected
                                ? '2px solid #0284C7'
                                : showInteractiveOverlays && highlightEditableCells && isEditable
                                ? '1.5px dashed #0284C7'
                                : 'none',
                              outlineOffset: '-1.5px',
                            }}
                          >
                            {cell.formattedValue}
                            {showInteractiveOverlays && highlightEditableCells && isEditable && (
                              <span
                                style={{
                                  position: 'absolute',
                                  top: 1,
                                  left: isRightAligned ? 3 : undefined,
                                  right: isRightAligned ? undefined : 3,
                                  fontSize: '8px',
                                  lineHeight: 1,
                                  color: '#0369A1',
                                  fontFamily: 'var(--font-mono)',
                                  fontWeight: 600,
                                  pointerEvents: 'none',
                                }}
                              >
                                {addr}
                              </span>
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};
