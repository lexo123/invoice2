/**
 * Standard Excel Indexed Color Palette (0..63) per ECMA-376 SpreadsheetML specification.
 */
const EXCEL_INDEXED_COLORS: Record<number, string> = {
  0: '#000000',
  1: '#FFFFFF',
  2: '#FF0000',
  3: '#00FF00',
  4: '#0000FF',
  5: '#FFFF00',
  6: '#FF00FF',
  7: '#00FFFF',
  8: '#000000',
  9: '#FFFFFF',
  10: '#FF0000',
  11: '#00FF00',
  12: '#0000FF',
  13: '#FFFF00',
  14: '#FF00FF',
  15: '#00FFFF',
  16: '#800000',
  17: '#008000',
  18: '#000080',
  19: '#808000',
  20: '#800080',
  21: '#008080',
  22: '#C0C0C0',
  23: '#808080',
  24: '#9999FF',
  25: '#993366',
  26: '#FFFFCC',
  27: '#CCFFFF',
  28: '#660066',
  29: '#FF8080',
  30: '#0066CC',
  31: '#CCCCFF',
  32: '#000080',
  33: '#FF00FF',
  34: '#FFFF00',
  35: '#00FFFF',
  36: '#800080',
  37: '#800000',
  38: '#008080',
  39: '#0000FF',
  40: '#00CCFF',
  41: '#CCFFFF',
  42: '#CCFFCC',
  43: '#FFFF99',
  44: '#99CCFF',
  45: '#FF99CC',
  46: '#CC99FF',
  47: '#FFCC99',
  48: '#3366FF',
  49: '#33CCCC',
  50: '#99CC00',
  51: '#FFCC00',
  52: '#FF9900',
  53: '#FF6600',
  54: '#666699',
  55: '#969696',
  56: '#003366',
  57: '#339966',
  58: '#003300',
  59: '#333300',
  60: '#993300',
  61: '#993366',
  62: '#333399',
  63: '#333333',
};

/**
 * Default Office Theme Palette mapped to SpreadsheetML theme indices:
 * 0: lt1, 1: dk1, 2: lt2, 3: dk2, 4: accent1, 5: accent2, 6: accent3, 7: accent4, 8: accent5, 9: accent6
 */
const DEFAULT_THEME_PALETTE: string[] = [
  'FFFFFF', // 0: lt1
  '000000', // 1: dk1
  'E7E6E6', // 2: lt2
  '44546A', // 3: dk2
  '4472C4', // 4: accent1
  'ED7D31', // 5: accent2
  'A5A5A5', // 6: accent3
  'FFC000', // 7: accent4
  '5B9BD5', // 8: accent5
  '70AD47', // 9: accent6
];

/**
 * Extracts the 10-slot SpreadsheetML theme palette from OpenXML theme1.xml string if present.
 */
export function extractThemePaletteFromXml(themeXml?: string): string[] {
  const palette = [...DEFAULT_THEME_PALETTE];
  if (!themeXml || typeof themeXml !== 'string') return palette;

  const extractTagHex = (tagName: string): string | null => {
    const regex = new RegExp(`<a:${tagName}[^>]*>([\\s\\S]*?)<\\/a:${tagName}>`, 'i');
    const match = themeXml.match(regex);
    if (!match) return null;
    const inner = match[1];
    const srgb = inner.match(/<a:srgbClr[^>]*val="([0-9A-Fa-f]{6})"/i);
    if (srgb) return srgb[1].toUpperCase();
    const sys = inner.match(/<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/i);
    if (sys) return sys[1].toUpperCase();
    return null;
  };

  // SpreadsheetML order: 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..accent6
  const tagMap: [number, string][] = [
    [0, 'lt1'],
    [1, 'dk1'],
    [2, 'lt2'],
    [3, 'dk2'],
    [4, 'accent1'],
    [5, 'accent2'],
    [6, 'accent3'],
    [7, 'accent4'],
    [8, 'accent5'],
    [9, 'accent6'],
  ];

  for (const [idx, tag] of tagMap) {
    const hex = extractTagHex(tag);
    if (hex) palette[idx] = hex;
  }

  return palette;
}

/**
 * Applies Excel tint (-1.0 to +1.0) to a 6-character hex color string.
 */
export function applyExcelTint(hex6: string, tint?: number): string {
  const clean = hex6.replace(/^#/, '');
  if (clean.length !== 6) return `#${clean}`;
  if (tint === undefined || tint === null || tint === 0 || Number.isNaN(tint)) {
    return `#${clean}`;
  }

  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);

  const adjustChannel = (c: number): number => {
    if (tint < 0) {
      return Math.max(0, Math.min(255, Math.round(c * (1 + tint))));
    }
    return Math.max(0, Math.min(255, Math.round(c * (1 - tint) + 255 * tint)));
  };

  const nr = adjustChannel(r).toString(16).padStart(2, '0');
  const ng = adjustChannel(g).toString(16).padStart(2, '0');
  const nb = adjustChannel(b).toString(16).padStart(2, '0');
  return `#${nr}${ng}${nb}`.toUpperCase();
}

/**
 * Resolves any ExcelJS color object ({ argb, theme, tint, indexed }) into a CSS hex color.
 */
export function resolveExcelColor(
  colorObj: any,
  themePalette: string[] = DEFAULT_THEME_PALETTE,
  fallback = ''
): string {
  if (!colorObj || typeof colorObj !== 'object') return fallback;

  // 1. Direct ARGB string (e.g. "FF1F497D" or "1F497D")
  if (typeof colorObj.argb === 'string' && colorObj.argb.trim() !== '') {
    const clean = colorObj.argb.replace(/^#/, '').trim();
    const hex6 = clean.length === 8 ? clean.slice(2) : clean.length === 6 ? clean : '';
    if (hex6) {
      return applyExcelTint(hex6, colorObj.tint);
    }
  }

  // 2. Theme color index (0..9) + optional tint
  if (typeof colorObj.theme === 'number' && colorObj.theme >= 0) {
    const baseHex = themePalette[colorObj.theme] ?? DEFAULT_THEME_PALETTE[colorObj.theme] ?? '000000';
    return applyExcelTint(baseHex, colorObj.tint);
  }

  // 3. Indexed color (0..63)
  if (typeof colorObj.indexed === 'number') {
    if (EXCEL_INDEXED_COLORS[colorObj.indexed]) {
      return EXCEL_INDEXED_COLORS[colorObj.indexed];
    }
  }

  return fallback;
}
