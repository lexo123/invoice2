export type CellInputType = 'text' | 'number' | 'currency' | 'date' | 'formula';
export type AffixPosition = 'start' | 'end';
export type CurrencySymbol = '₾' | '$';
export type FormulaResultFormat = 'number' | 'currency';

export type MainHeaderFieldKey =
  | 'companyName'
  | 'cadastralCode'
  | 'address'
  | 'usdRate'
  | 'date'
  | 'invoiceNumber';

export interface MainHeaderFieldConfig {
  key: MainHeaderFieldKey;
  label: string;
  address: string; // e.g. "A8", "C7", etc.
  inputType: 'text' | 'number' | 'date';
  placeholder: string;
  usePrefix: boolean;
  prefixText: string;
  affixPosition: AffixPosition;
}

export interface ProductRowItem {
  id: string;
  productName: string; // Written to A17..A24
  quantity: string; // Written to B17..B24
  unitPrice: string; // Written to C17..C24
  currency: CurrencySymbol; // '₾' (default) or '$'
}

export interface CalculatedProductRow {
  rowIndex: number; // 0..7
  excelRow: number; // 17..24
  productCell: string; // A17..A24
  qtyCell: string; // B17..B24
  priceCell: string; // C17..C24
  rowTotalCell: string; // D17..D24
  hasValues: boolean;
  qtyNumeric: number;
  priceNumeric: number;
  currency: CurrencySymbol;
  lineTotalOriginal: number;
  lineTotalGel: number;
  formattedOriginal: string;
  formattedGel: string;
}

export interface ProductsCalculationSummary {
  rows: CalculatedProductRow[];
  usdRateNumeric: number;
  hasValidUsdRate: boolean;
  hasAnyUsdProduct: boolean;
  grandTotalGel: number;
  formattedGrandTotalGel: string;
  grandTotalFormula: string;
}

export interface EditableCellInstruction {
  id: string;
  address: string; // e.g. "F3", "C7", "D12"
  label: string; // Georgian human-readable label
  group: 'header' | 'client' | 'items' | 'other';
  inputType: CellInputType;
  currencySymbol?: CurrencySymbol; // Used when inputType === 'currency' or formulaResultFormat === 'currency' (default: '₾')
  formulaResultFormat?: FormulaResultFormat; // Used when inputType === 'formula': 'number' (რიცხვი) or 'currency' (თანხა)
  placeholder?: string;
  usePrefix?: boolean; // If true, attaches prefixText at the start or end of user input
  prefixText?: string; // e.g. "მისამართი" (at start) or "ლარი" (at end)
  affixPosition?: AffixPosition; // 'start' (თავში, default) or 'end' (ბოლოში)
  optional?: boolean; // If true, user may leave this cell blank; formulas ignore empty optional cells
  formulaExpression?: string; // Formula expression configured directly in instructions when inputType === 'formula'
  mainFieldKey?: MainHeaderFieldKey;
}

export interface CellBorderSide {
  style: string; // thin, medium, thick, double, dashed
  color: string; // CSS hex color
}

export interface ParsedSheetCell {
  address: string;
  row: number;
  col: number;
  value: string | number;
  formattedValue: string;
  formula?: string;
  numFmt?: string;
  // Merge info
  isMergedSlave?: boolean;
  masterAddress?: string;
  rowSpan: number;
  colSpan: number;
  // Styling
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  textColor?: string;
  bgColor?: string;
  alignHorizontal?: 'left' | 'center' | 'right';
  alignVertical?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  borders: {
    top?: CellBorderSide;
    right?: CellBorderSide;
    bottom?: CellBorderSide;
    left?: CellBorderSide;
  };
}

export interface EmbeddedSheetImage {
  id: string;
  dataUrl: string;
  tlCol: number; // 0-indexed column (can be fractional)
  tlRow: number; // 0-indexed row (can be fractional)
  brCol?: number;
  brRow?: number;
  widthPx?: number;
  heightPx?: number;
}

export interface ParsedSheetModel {
  sheetName: string;
  rowCount: number;
  colCount: number;
  colWidthsPx: number[]; // 1-indexed (index 1..colCount)
  rowHeightsPx: number[]; // 1-indexed (index 1..rowCount)
  cells: Record<string, ParsedSheetCell>;
  images: EmbeddedSheetImage[];
}

export interface GeneratedInvoicePackage {
  xlsxBuffer: ArrayBuffer;
  fileNameBase: string;
  generatedAt: string;
  sheetModel: ParsedSheetModel;
  summaryMetrics: {
    invoiceNumber: string;
    clientName: string;
    subtotalFormatted: string;
    vatFormatted: string;
    grandTotalFormatted: string;
    changedCellsCount: number;
    recalculatedFormulasCount: number;
  };
}
