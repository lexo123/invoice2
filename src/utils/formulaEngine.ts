import {
  CalculatedProductRow,
  CurrencySymbol,
  EditableCellInstruction,
  MainHeaderFieldConfig,
  MainHeaderFieldKey,
  ParsedSheetModel,
  ProductRowItem,
  ProductsCalculationSummary,
} from '../types/invoice';

export interface CellCoord {
  col: number; // 1-indexed
  row: number; // 1-indexed
  address: string; // e.g. "C12"
}

export interface FormulaCellData {
  value: string | number;
  formula?: string;
}

/**
 * Converts YYYY-MM-DD to DD.MM.YYYY
 */
export function isoDateToGeorgianDot(iso: string): string {
  const match = String(iso ?? '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return iso;
  return `${match[3]}.${match[2]}.${match[1]}`;
}

/**
 * Returns today's date formatted as DD.MM.YYYY
 */
export function getTodayGeorgianDate(now: Date = new Date()): string {
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

/**
 * Converts DD.MM.YYYY or DD/MM/YYYY to YYYY-MM-DD for native <input type="date">
 */
export function georgianDotToIsoDate(dotStr: string): string {
  const clean = String(dotStr ?? '').trim();
  const m1 = clean.match(/^(\d{2})[./-](\d{2})[./-](\d{4})$/);
  if (m1) {
    return `${m1[3]}-${m1[2]}-${m1[1]}`;
  }
  const m2 = clean.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) {
    return clean;
  }
  return '';
}

/**
 * Synchronizes the File Name ("ფაილის სახელი") field when Company Name ("კომპანიის სახელი") changes.
 * By default (when the user has not manually customized the file name), the file name mirrors Company Name.
 */
export function syncFileNameWithCompanyName(
  prevCompanyName: string,
  nextCompanyName: string,
  currentFileName: string,
  isFileNameCustomized: boolean
): { fileName: string; isFileNameCustomized: boolean } {
  if (
    !isFileNameCustomized ||
    currentFileName.trim() === '' ||
    currentFileName === prevCompanyName
  ) {
    return {
      fileName: nextCompanyName,
      isFileNameCustomized: false,
    };
  }
  return {
    fileName: currentFileName,
    isFileNameCustomized: true,
  };
}

/**
 * Resolves the final base file name (without .pdf / .xlsx extension) for downloading generated files.
 * Priority:
 * 1. User's fileNameInput (if non-empty)
 * 2. companyNameInput (default when fileNameInput is not manually changed or left blank)
 * 3. fallbackBase ('ინვოისი')
 */
export function resolveInvoiceFileNameBase(
  fileNameInput: string,
  companyNameInput: string,
  fallbackBase = 'ინვოისი'
): string {
  const rawCandidate =
    String(fileNameInput ?? '').trim() ||
    String(companyNameInput ?? '').trim() ||
    fallbackBase;

  // Strip trailing .pdf or .xlsx if user typed an extension manually
  const withoutExt = rawCandidate.replace(/\.(pdf|xlsx)$/i, '').trim();

  // Remove only characters forbidden in filenames by OS/browsers while keeping Georgian text & spaces intact
  const sanitized = withoutExt
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return sanitized || fallbackBase;
}

/**
 * Composes the final Excel cell value when a cell has an optional prefix/suffix ("სავალდებულო სიტყვა").
 * If inputType is 'formula' or userInput starts with '=', returns a normalized formula string starting with '='.
 */
export function composeCellValueWithPrefix(
  userInput: string | number,
  usePrefix?: boolean,
  prefixText?: string,
  inputType: 'text' | 'number' | 'currency' | 'date' | 'formula' = 'text',
  optional?: boolean,
  affixPosition: 'start' | 'end' = 'start'
): string | number {
  const rawStr = String(userInput ?? '').trim();

  // If cell is a formula type or starts with '=', treat as an Excel formula
  if (inputType === 'formula' || (rawStr.startsWith('=') && rawStr.length > 1)) {
    const cleanFormula = rawStr.replace(/^=+/, '').trim();
    return cleanFormula ? `=${cleanFormula}` : '';
  }

  // If cell is marked optional and user left it empty, return empty string even if prefix/suffix is configured
  if (optional && rawStr === '') {
    return '';
  }

  if (!usePrefix || !prefixText || prefixText.trim() === '') {
    if (
      (inputType === 'number' || inputType === 'currency') &&
      typeof userInput === 'string' &&
      rawStr !== ''
    ) {
      const cleanedNumStr = rawStr.replace(/[₾$,\s]/g, '');
      const n = Number(cleanedNumStr);
      if (cleanedNumStr !== '' && !Number.isNaN(n)) return n;
    }
    return userInput;
  }

  const cleanAffix = prefixText.trim();
  if (!rawStr) {
    return cleanAffix;
  }

  if (affixPosition === 'end') {
    // Avoid duplicating suffix if user already typed it at the end
    if (rawStr.toLowerCase().endsWith(cleanAffix.toLowerCase())) {
      return rawStr;
    }
    return `${rawStr} ${cleanAffix}`;
  }

  // Default: 'start' (თავში)
  if (rawStr.toLowerCase().startsWith(cleanAffix.toLowerCase())) {
    return rawStr;
  }

  return `${cleanAffix} ${rawStr}`;
}

/**
 * Formats a numeric/amount value with the selected currency symbol ('₾' default, or '$').
 * Integer values like 100 format as "100 ₾" or "100 $", while decimals format with 2 decimal places.
 */
export function formatCurrencyCellValue(
  value: string | number | null | undefined,
  symbol: '₾' | '$' = '₾',
  _numFmt?: string
): string {
  if (value === null || value === undefined || value === '') return '';
  const cleaned =
    typeof value === 'number'
      ? value
      : Number(String(value).replace(/[₾$,\s]/g, ''));
  if (!Number.isFinite(cleaned) || String(value).trim() === '') {
    const raw = String(value).trim();
    if (!raw) return '';
    return raw.includes(symbol) ? raw : `${raw} ${symbol}`;
  }

  const isInt = Number.isInteger(cleaned);
  const minDecimals = isInt ? 0 : 2;
  const maxDecimals = isInt ? 0 : 2;

  const formattedNum = cleaned.toLocaleString('en-US', {
    minimumFractionDigits: minDecimals,
    maximumFractionDigits: maxDecimals,
  });

  return `${formattedNum} ${symbol}`;
}

/**
 * Attaches an optional start/end word to an already evaluated display string (e.g. for formula cells).
 */
export function applyAffixToDisplayValue(
  displayValue: string | number,
  usePrefix?: boolean,
  prefixText?: string,
  affixPosition: 'start' | 'end' = 'start'
): string {
  const str = String(displayValue ?? '').trim();
  if (!str) return '';
  if (!usePrefix || !prefixText || !prefixText.trim()) return str;
  const cleanAffix = prefixText.trim();
  if (affixPosition === 'end') {
    if (str.toLowerCase().endsWith(cleanAffix.toLowerCase())) return str;
    return `${str} ${cleanAffix}`;
  }
  if (str.toLowerCase().startsWith(cleanAffix.toLowerCase())) return str;
  return `${cleanAffix} ${str}`;
}

const KNOWN_MULTI_WORD_PREFIXES = [
  'კომპანიის სახელი',
  'დოლარის კურსი',
  'ინვოისის ნომერი',
  'საკადასტრო კოდი',
  'საიდენტიფიკაციო კოდი',
  'გამოწერის თარიღი',
  'გადახდის ვადა',
  'მყიდველი ორგანიზაცია',
  'ერთეულის ფასი',
];

/**
 * Extracts a suggested prefix word/phrase and remaining value from an existing cell text.
 * Preserves multi-word Georgian labels such as "დოლარის კურსი", "კომპანიის სახელი", "ინვოისის ნომერი".
 */
export function detectCellPrefixCandidate(rawCellText: string): {
  suggestedPrefix: string;
  remainder: string;
} {
  const text = String(rawCellText ?? '').trim();
  if (!text) {
    return { suggestedPrefix: '', remainder: '' };
  }
  const colonIdx = text.indexOf(':');
  if (colonIdx > 0 && colonIdx < 35) {
    return {
      suggestedPrefix: text.slice(0, colonIdx + 1).trim(),
      remainder: text.slice(colonIdx + 1).trim(),
    };
  }
  const lower = text.toLowerCase();
  for (const phrase of KNOWN_MULTI_WORD_PREFIXES) {
    if (lower.startsWith(phrase.toLowerCase())) {
      const rest = text.slice(phrase.length).trim();
      return {
        suggestedPrefix: text.slice(0, phrase.length).trim(),
        remainder: rest,
      };
    }
  }
  const firstSpace = text.indexOf(' ');
  if (firstSpace > 0) {
    return {
      suggestedPrefix: text.slice(0, firstSpace).trim(),
      remainder: text.slice(firstSpace + 1).trim(),
    };
  }
  return {
    suggestedPrefix: text,
    remainder: '',
  };
}

/**
 * Strips a configured prefix or suffix from a raw cell value so the input box only shows the editable value.
 */
export function stripPrefixFromCellValue(
  fullValue: string | number,
  usePrefix?: boolean,
  prefixText?: string,
  affixPosition: 'start' | 'end' = 'start'
): string | number {
  if (!usePrefix || !prefixText || !prefixText.trim()) {
    return fullValue;
  }
  const str = String(fullValue ?? '').trim();
  const affix = prefixText.trim();
  if (affixPosition === 'end') {
    if (str.toLowerCase().endsWith(affix.toLowerCase())) {
      return str.slice(0, str.length - affix.length).trim();
    }
    return fullValue;
  }
  if (str.toLowerCase().startsWith(affix.toLowerCase())) {
    return str.slice(affix.length).trim();
  }
  return fullValue;
}

/**
 * Converts 1-indexed column number to Excel column letters (1 -> A, 26 -> Z, 27 -> AA)
 */
export function colNumberToLetter(col: number): string {
  let temp = col;
  let letter = '';
  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
}

/**
 * Converts Excel column letters to 1-indexed column number (A -> 1, Z -> 26)
 */
export function letterToColNumber(letters: string): number {
  const clean = letters.toUpperCase().replace(/[^A-Z]/g, '');
  let col = 0;
  for (let i = 0; i < clean.length; i++) {
    col = col * 26 + (clean.charCodeAt(i) - 64);
  }
  return col;
}

/**
 * Parses an Excel address like "C12" or "$D$14" into 1-indexed col, row, and normalized address.
 */
export function parseCellAddress(ref: string): CellCoord | null {
  const clean = ref.trim().replace(/\$/g, '').toUpperCase();
  const match = clean.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const col = letterToColNumber(match[1]);
  const row = parseInt(match[2], 10);
  if (col < 1 || row < 1 || Number.isNaN(row)) return null;
  return {
    col,
    row,
    address: `${match[1]}${row}`,
  };
}

/**
 * Expands a cell range like "F12:F15" into an array of cell addresses ["F12", "F13", "F14", "F15"]
 */
export function expandRange(rangeStr: string): string[] {
  const parts = rangeStr.split(':');
  if (parts.length !== 2) {
    const single = parseCellAddress(rangeStr);
    return single ? [single.address] : [];
  }
  const start = parseCellAddress(parts[0]);
  const end = parseCellAddress(parts[1]);
  if (!start || !end) return [];

  const minCol = Math.min(start.col, end.col);
  const maxCol = Math.max(start.col, end.col);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  const addresses: string[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      addresses.push(`${colNumberToLetter(c)}${r}`);
    }
  }
  return addresses;
}

/**
 * Checks whether a cell value is empty (blank string, null, or undefined).
 * Note: 0 or "0" is NOT empty.
 */
export function isCellValueEmpty(val: unknown): boolean {
  if (val === null || val === undefined) return true;
  if (typeof val === 'string') return val.trim() === '';
  return false;
}

function toNumeric(val: unknown): number {
  if (typeof val === 'number') return Number.isFinite(val) ? val : 0;
  if (typeof val === 'string') {
    const cleaned = val.replace(/[^\d.-]/g, '');
    const parsed = parseFloat(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

/**
 * Safe arithmetic expression evaluator supporting +, -, *, /, parentheses, and numbers.
 * Returns null if the expression contains unsupported syntax so caller can fallback to cached Excel result.
 */
function evaluateArithmeticExpression(expr: string): number | null {
  const sanitized = expr.replace(/\s+/g, '');
  if (!sanitized) return null;
  if (!/^[0-9+\-*/().]+$/.test(sanitized)) {
    return null;
  }
  try {
    const tokens: (number | string)[] = [];
    let i = 0;
    while (i < sanitized.length) {
      const ch = sanitized[i];
      if (ch === '(' || ch === ')' || ch === '+' || ch === '*' || ch === '/') {
        tokens.push(ch);
        i++;
      } else if (ch === '-') {
        const prev = tokens[tokens.length - 1];
        if (
          tokens.length === 0 ||
          prev === '(' ||
          prev === '+' ||
          prev === '-' ||
          prev === '*' ||
          prev === '/'
        ) {
          let numStr = '-';
          i++;
          while (i < sanitized.length && /[0-9.]/.test(sanitized[i])) {
            numStr += sanitized[i];
            i++;
          }
          if (numStr === '-') {
            tokens.push(0);
            tokens.push('-');
          } else {
            tokens.push(parseFloat(numStr));
          }
        } else {
          tokens.push('-');
          i++;
        }
      } else if (/[0-9.]/.test(ch)) {
        let numStr = '';
        while (i < sanitized.length && /[0-9.]/.test(sanitized[i])) {
          numStr += sanitized[i];
          i++;
        }
        tokens.push(parseFloat(numStr));
      } else {
        i++;
      }
    }

    const precedence: Record<string, number> = { '+': 1, '-': 1, '*': 2, '/': 2 };
    const values: number[] = [];
    const ops: string[] = [];

    const applyOp = () => {
      const op = ops.pop();
      const b = values.pop() ?? 0;
      const a = values.pop() ?? 0;
      if (op === '+') values.push(a + b);
      else if (op === '-') values.push(a - b);
      else if (op === '*') values.push(a * b);
      else if (op === '/') values.push(b === 0 ? 0 : a / b);
    };

    for (const token of tokens) {
      if (typeof token === 'number') {
        values.push(token);
      } else if (token === '(') {
        ops.push(token);
      } else if (token === ')') {
        while (ops.length > 0 && ops[ops.length - 1] !== '(') {
          applyOp();
        }
        ops.pop();
      } else if (token in precedence) {
        while (
          ops.length > 0 &&
          ops[ops.length - 1] in precedence &&
          precedence[ops[ops.length - 1]] >= precedence[token]
        ) {
          applyOp();
        }
        ops.push(token);
      }
    }

    while (ops.length > 0) {
      applyOp();
    }

    const res = values[0];
    if (res === undefined || !Number.isFinite(res)) return null;
    return Math.round(res * 1e6) / 1e6;
  } catch {
    return null;
  }
}

/**
 * Evaluates common Excel formulas used in invoices:
 * SUM, AVERAGE, MIN, MAX, ROUND, IF, and arithmetic with cell coordinates.
 * Handles empty/optional cells gracefully:
 * - If a cell is empty, SUM/AVERAGE/MIN/MAX ignore it.
 * - In arithmetic products/quotients (e.g. =D12*E12*B6), if at least one factor in the term has a value
 *   and another cell is empty, the empty cell is ignored (acts as 1 for * and /, or 0 for + and -).
 * - If all cells referenced by the formula are empty, returns '' (empty).
 */
export function evaluateExcelFormula(
  rawFormula: string,
  getCellValue: (address: string) => string | number
): number | string | null {
  const formula = rawFormula.trim().replace(/^=/, '').trim();
  if (!formula) return null;

  let sawAnyCellRef = false;
  let sawNonEmptyCellRef = false;

  const resolveArgValues = (argStr: string): number[] => {
    const segments = argStr.split(/[,;]/).map((s) => s.trim()).filter(Boolean);
    const nums: number[] = [];
    for (const seg of segments) {
      if (seg.includes(':')) {
        const addrs = expandRange(seg);
        for (const addr of addrs) {
          sawAnyCellRef = true;
          const rawCell = getCellValue(addr);
          if (!isCellValueEmpty(rawCell)) {
            sawNonEmptyCellRef = true;
            nums.push(toNumeric(rawCell));
          }
        }
      } else {
        const parsedCoord = parseCellAddress(seg);
        if (parsedCoord) {
          sawAnyCellRef = true;
          const rawCell = getCellValue(parsedCoord.address);
          if (!isCellValueEmpty(rawCell)) {
            sawNonEmptyCellRef = true;
            nums.push(toNumeric(rawCell));
          }
        } else {
          const sub = evaluateExcelFormula(seg, getCellValue);
          if (!isCellValueEmpty(sub)) {
            nums.push(toNumeric(sub ?? seg));
          }
        }
      }
    }
    return nums;
  };

  let expr = formula;

  // Iteratively resolve innermost functions: SUM(...), AVERAGE(...), MIN(...), MAX(...), ROUND(...), IF(...)
  const fnRegex = /\b(SUM|AVERAGE|MIN|MAX|ROUND|IF)\s*\(([^()]*)\)/gi;
  let safety = 0;
  while (fnRegex.test(expr) && safety < 20) {
    safety++;
    expr = expr.replace(fnRegex, (_, fnName: string, innerArgs: string) => {
      const upper = fnName.toUpperCase();
      if (upper === 'IF') {
        const parts = innerArgs.split(/[,;]/).map((s) => s.trim());
        const condStr = parts[0] || '';
        const trueExpr = parts[1] ?? '0';
        const falseExpr = parts[2] ?? '0';
        let condResult = false;
        const cmpMatch = condStr.match(/^(.+?)(>=|<=|<>|!=|>|<|=)(.+)$/);
        if (cmpMatch) {
          const leftRaw = cmpMatch[1].trim();
          const op = cmpMatch[2];
          const rightRaw = cmpMatch[3].trim();
          const leftCoord = parseCellAddress(leftRaw);
          if (leftCoord) sawAnyCellRef = true;
          const leftVal = leftCoord
            ? getCellValue(leftCoord.address)
            : evaluateExcelFormula(leftRaw, getCellValue) ?? leftRaw;
          if (!isCellValueEmpty(leftVal)) sawNonEmptyCellRef = true;
          const rightVal =
            rightRaw === '""' || rightRaw === "''"
              ? ''
              : evaluateExcelFormula(rightRaw, getCellValue) ?? rightRaw;
          if (rightRaw === '""' || rightRaw === "''") {
            const empty = isCellValueEmpty(leftVal);
            condResult = op === '=' ? empty : !empty;
          } else {
            const lNum = toNumeric(leftVal);
            const rNum = toNumeric(rightVal);
            if (op === '>') condResult = lNum > rNum;
            else if (op === '<') condResult = lNum < rNum;
            else if (op === '>=') condResult = lNum >= rNum;
            else if (op === '<=') condResult = lNum <= rNum;
            else if (op === '=') condResult = lNum === rNum;
            else condResult = lNum !== rNum;
          }
        } else {
          const cCoord = parseCellAddress(condStr);
          if (cCoord) {
            sawAnyCellRef = true;
            const raw = getCellValue(cCoord.address);
            if (!isCellValueEmpty(raw)) sawNonEmptyCellRef = true;
            condResult = !isCellValueEmpty(raw) && toNumeric(raw) !== 0;
          } else {
            const sub = evaluateExcelFormula(condStr, getCellValue);
            condResult = !isCellValueEmpty(sub) && toNumeric(sub) !== 0;
          }
        }
        const chosen = condResult ? trueExpr : falseExpr;
        if (chosen === '""' || chosen === "''") return '0';
        const chosenEval = evaluateExcelFormula(chosen, getCellValue);
        if (!isCellValueEmpty(chosenEval)) sawNonEmptyCellRef = true;
        return String(toNumeric(chosenEval ?? chosen));
      }
      if (upper === 'ROUND') {
        const parts = innerArgs.split(/[,;]/);
        const subEval = evaluateExcelFormula(parts[0] || '0', getCellValue);
        if (!isCellValueEmpty(subEval)) {
          sawNonEmptyCellRef = true;
        }
        const val = toNumeric(subEval ?? parts[0]);
        const digits = parts[1] !== undefined ? parseInt(parts[1].trim(), 10) : 2;
        const factor = Math.pow(10, Number.isNaN(digits) ? 2 : digits);
        return String(Math.round(val * factor) / factor);
      }
      const vals = resolveArgValues(innerArgs);
      if (upper === 'SUM') {
        return String(vals.reduce((acc, v) => acc + v, 0));
      }
      if (upper === 'AVERAGE') {
        return String(vals.length ? vals.reduce((acc, v) => acc + v, 0) / vals.length : 0);
      }
      if (upper === 'MIN') {
        return String(vals.length ? Math.min(...vals) : 0);
      }
      if (upper === 'MAX') {
        return String(vals.length ? Math.max(...vals) : 0);
      }
      return '0';
    });
  }

  // Inspect remaining direct cell references in expr
  const cellRefMatches = expr.match(/\$?[A-Z]{1,3}\$?\d{1,5}/gi) || [];
  for (const m of cellRefMatches) {
    const coord = parseCellAddress(m);
    if (coord) {
      sawAnyCellRef = true;
      if (!isCellValueEmpty(getCellValue(coord.address))) {
        sawNonEmptyCellRef = true;
      }
    }
  }

  // If the formula references cells and EVERY referenced cell is empty, return '' (empty)
  if (sawAnyCellRef && !sawNonEmptyCellRef) {
    return '';
  }

  // Replace remaining cell references (e.g. D12, $E$12) with their numeric values.
  // If a cell is empty in a multiplicative sub-term that has at least one non-empty cell/number,
  // treat the empty cell as 1 (multiplicative identity) so empty optional cells are ignored!
  expr = expr.replace(/\$?[A-Z]{1,3}\$?\d{1,5}/gi, (match, offset: number, fullStr: string) => {
    const coord = parseCellAddress(match);
    if (!coord) return match;
    const rawVal = getCellValue(coord.address);
    if (!isCellValueEmpty(rawVal)) {
      const val = toNumeric(rawVal);
      return val < 0 ? `(${val})` : String(val);
    }

    // The cell is empty while at least one other input in the formula is non-empty.
    // Check whether this cell reference is connected via * or /
    const beforeStr = fullStr.slice(0, offset).trim();
    const afterStr = fullStr.slice(offset + match.length).trim();
    const prevChar = beforeStr[beforeStr.length - 1] || '';
    const nextChar = afterStr[0] || '';

    const isMultiplicativeContext =
      prevChar === '*' ||
      prevChar === '/' ||
      ((nextChar === '*' || nextChar === '/') && prevChar !== '+' && prevChar !== '-');

    if (isMultiplicativeContext) {
      // Find the current additive term (bounded by top-level + or -) to see if anything in this product term is non-empty
      let termStart = offset;
      while (termStart > 0 && fullStr[termStart - 1] !== '+' && fullStr[termStart - 1] !== '-') {
        termStart--;
      }
      let termEnd = offset + match.length;
      while (termEnd < fullStr.length && fullStr[termEnd] !== '+' && fullStr[termEnd] !== '-') {
        termEnd++;
      }
      const termSlice = fullStr.slice(termStart, termEnd);
      const termRefs = termSlice.match(/\$?[A-Z]{1,3}\$?\d{1,5}/gi) || [];
      const hasOtherNonEmptyFactor = termRefs.some((r) => {
        const c = parseCellAddress(r);
        return c && c.address !== coord.address && !isCellValueEmpty(getCellValue(c.address));
      });
      // Also check if there is a numeric literal factor in the term (e.g. already resolved SUM(...) * B6)
      const strippedRefs = termSlice.replace(/\$?[A-Z]{1,3}\$?\d{1,5}/gi, '');
      const hasNumericLiteralFactor = /\d/.test(strippedRefs);

      if (hasOtherNonEmptyFactor || hasNumericLiteralFactor) {
        return '1';
      }
    }

    return '0';
  });

  return evaluateArithmeticExpression(expr);
}

/**
 * Applies user overrides to editable cells while leaving unchanged cells intact.
 * If a user override starts with '=' (a custom user-entered formula), it is registered
 * as an active Excel formula on that cell and evaluated along with all sheet formulas!
 */
export function recalculateSheetCells(
  initialCells: Record<string, FormulaCellData>,
  editableOverrides: Record<string, string | number>
): Record<string, FormulaCellData> {
  const result: Record<string, FormulaCellData> = {};
  const nonFormulaOverriddenAddresses = new Set<string>();

  for (const [addr, cell] of Object.entries(initialCells)) {
    const norm = parseCellAddress(addr)?.address ?? addr.toUpperCase();
    result[norm] = { ...cell };
  }

  // Apply overrides to editable cells
  for (const [addr, newValue] of Object.entries(editableOverrides)) {
    const norm = parseCellAddress(addr)?.address ?? addr.toUpperCase();
    const existing = result[norm];

    // Check if the user entered a formula (e.g. "=D12*E12" or "=SUM(F12:F15)")
    if (typeof newValue === 'string' && newValue.trim().startsWith('=') && newValue.trim().length > 1) {
      const cleanFormula = newValue.trim().replace(/^=+/, '').trim();
      result[norm] = {
        value: existing?.value ?? 0,
        formula: cleanFormula,
      };
      continue;
    }

    nonFormulaOverriddenAddresses.add(norm);
    let coercedValue: string | number = newValue;
    if (typeof newValue === 'string') {
      const trimmed = newValue.trim();
      if (trimmed === '') {
        coercedValue = '';
      } else if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
        coercedValue = parseFloat(trimmed);
      }
    }
    result[norm] = {
      value: coercedValue,
      // If user explicitly overwrote a cell with a static value, clear any old formula on it
      formula: undefined,
    };
  }

  // Multi-pass recalculation for all formula cells (both template formulas and user-entered formulas!)
  const formulaAddresses = Object.keys(result).filter(
    (addr) => Boolean(result[addr].formula) && !nonFormulaOverriddenAddresses.has(addr)
  );

  for (let pass = 0; pass < 6; pass++) {
    for (const addr of formulaAddresses) {
      const f = result[addr].formula;
      if (f) {
        const evaluated = evaluateExcelFormula(f, (refAddr) => {
          const normRef = parseCellAddress(refAddr)?.address ?? refAddr.toUpperCase();
          return result[normRef]?.value ?? '';
        });
        if (evaluated !== null) {
          result[addr].value = evaluated;
        }
      }
    }
  }

  return result;
}

/**
 * Formats a cell value according to Excel numFmt for visual display and PDF output.
 * NOTE: Never auto-injects '$' from Excel locale/accounting numFmt strings (such as [$-409] or $#,##0),
 * which previously caused plain numbers like 166 in D36 to display as $166.
 */
export function formatExcelCellValue(value: string | number | null | undefined, numFmt?: string): string {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'string') {
    return value;
  }
  if (!Number.isFinite(value)) return '';

  if (!numFmt || numFmt === 'General') {
    if (Number.isInteger(value)) return String(value);
    const rounded4 = Math.round(value * 10000) / 10000;
    return String(rounded4);
  }

  const hasGel = numFmt.includes('₾') || numFmt.includes('GEL');
  const hasPercent = numFmt.includes('%');

  if (hasPercent) {
    return `${(value * 100).toFixed(0)}%`;
  }

  const decMatch = numFmt.match(/\.(0+)/);
  const minDecimals = decMatch ? decMatch[1].length : 0;
  const maxDecimals = decMatch ? Math.max(decMatch[1].length, 2) : 4;

  const formattedNum = value.toLocaleString('en-US', {
    minimumFractionDigits: minDecimals,
    maximumFractionDigits: maxDecimals,
  });

  if (hasGel) return `${formattedNum} ₾`;
  return formattedNum;
}

/**
 * Pure calculation for up to 8 product rows (A17:C17 to A24:C24) + live conversion of USD ($) prices
 * to Lari (₾) using the Dollar Exchange Rate (`usdRateRaw`).
 * Grand total (`ჯამი`) is ALWAYS in Lari (₾).
 */
export function calculateProductsSummary(
  productRows: ProductRowItem[],
  usdRateRaw: string | number
): ProductsCalculationSummary {
  const cleanedRateStr = String(usdRateRaw ?? '')
    .trim()
    .replace(/,/g, '.')
    .replace(/[^\d.-]/g, '');
  const parsedRate = parseFloat(cleanedRateStr);
  const hasValidUsdRate = Number.isFinite(parsedRate) && parsedRate > 0;
  const effectiveUsdRate = hasValidUsdRate ? parsedRate : 1;

  const calculatedRows: CalculatedProductRow[] = [];
  const formulaTerms: string[] = [];
  let grandTotalGel = 0;
  let hasAnyUsdProduct = false;

  const boundedRows = productRows.slice(0, 8);

  for (let i = 0; i < boundedRows.length; i++) {
    const item = boundedRows[i];
    const excelRow = 17 + i; // A17..A24, B17..B24, C17..C24, D17..D24
    const productCell = `A${excelRow}`;
    const qtyCell = `B${excelRow}`;
    const priceCell = `C${excelRow}`;
    const rowTotalCell = `D${excelRow}`;

    const currency: CurrencySymbol = item.currency === '$' ? '$' : '₾';
    const rawQtyStr = String(item.quantity ?? '').trim().replace(/,/g, '.');
    const rawPriceStr = String(item.unitPrice ?? '').trim().replace(/,/g, '.');

    const hasPrice = rawPriceStr !== '' && !Number.isNaN(Number(rawPriceStr));
    const hasQty = rawQtyStr !== '' && !Number.isNaN(Number(rawQtyStr));
    const hasName = String(item.productName ?? '').trim() !== '';
    const hasValues = hasName || hasQty || hasPrice;

    // If user entered a price before typing quantity, treat effective quantity as 1 for live preview
    // (consistent with multiplicative formula evaluation in formulaEngine)
    const qtyNumeric = hasQty ? Number(rawQtyStr) : hasPrice ? 1 : 0;
    const priceNumeric = hasPrice ? Number(rawPriceStr) : 0;

    if (currency === '$' && (hasPrice || hasValues)) {
      hasAnyUsdProduct = true;
    }

    const lineTotalOriginal = Math.round(qtyNumeric * priceNumeric * 100) / 100;
    const lineTotalGel =
      currency === '$'
        ? Math.round(lineTotalOriginal * effectiveUsdRate * 100) / 100
        : lineTotalOriginal;

    if (hasPrice || hasQty) {
      grandTotalGel += lineTotalGel;
      if (currency === '$') {
        formulaTerms.push(`B${excelRow}*C${excelRow}*${effectiveUsdRate}`);
      } else {
        formulaTerms.push(`B${excelRow}*C${excelRow}`);
      }
    }

    calculatedRows.push({
      rowIndex: i,
      excelRow,
      productCell,
      qtyCell,
      priceCell,
      rowTotalCell,
      hasValues,
      qtyNumeric,
      priceNumeric,
      currency,
      lineTotalOriginal,
      lineTotalGel,
      formattedOriginal: formatCurrencyCellValue(lineTotalOriginal, currency),
      formattedGel: formatCurrencyCellValue(lineTotalGel, '₾'),
    });
  }

  grandTotalGel = Math.round(grandTotalGel * 100) / 100;
  const grandTotalFormula =
    formulaTerms.length > 0 ? `=${formulaTerms.join('+')}` : '=SUM(D17:D24)';

  return {
    rows: calculatedRows,
    usdRateNumeric: hasValidUsdRate ? parsedRate : 0,
    hasValidUsdRate,
    hasAnyUsdProduct,
    grandTotalGel,
    formattedGrandTotalGel: formatCurrencyCellValue(grandTotalGel, '₾'),
    grandTotalFormula,
  };
}

export const DEFAULT_MAIN_HEADER_CONFIGS: MainHeaderFieldConfig[] = [
  {
    key: 'companyName',
    label: 'კომპანიის სახელი',
    address: 'A8',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'კომპანიის სახელი',
    affixPosition: 'start',
  },
  {
    key: 'cadastralCode',
    label: 'საკადასტრო',
    address: 'A9',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'საკადასტრო',
    affixPosition: 'start',
  },
  {
    key: 'address',
    label: 'მისამართი',
    address: 'A10',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'მისამართი',
    affixPosition: 'start',
  },
  {
    key: 'usdRate',
    label: 'დოლარის კურსი',
    address: 'C10',
    inputType: 'number',
    placeholder: '',
    usePrefix: true,
    prefixText: 'დოლარის კურსი',
    affixPosition: 'start',
  },
  {
    key: 'date',
    label: 'თარიღი',
    address: 'C9',
    inputType: 'date',
    placeholder: '',
    usePrefix: true,
    prefixText: 'თარიღი',
    affixPosition: 'start',
  },
  {
    key: 'invoiceNumber',
    label: 'ინვოისის ნომერი',
    address: 'C8',
    inputType: 'text',
    placeholder: '',
    usePrefix: true,
    prefixText: 'ინვოისის ნომერი',
    affixPosition: 'start',
  },
];

const FIELD_KEYWORDS: Record<MainHeaderFieldKey, string[]> = {
  companyName: [
    'კომპანიის სახელი',
    'კომპანიის',
    'კომპანია',
    'მყიდველი ორგანიზაცია',
    'მყიდველი',
    'დამკვეთი',
    'ორგანიზაცია',
  ],
  cadastralCode: [
    'საკადასტრო კოდი',
    'საკადასტრო',
    'საიდენტიფიკაციო კოდი',
    'საიდენტიფიკაციო',
    'ს/კ',
    'ს.კ',
  ],
  address: ['მისამართი'],
  usdRate: ['დოლარის კურსი', 'დოლარის', 'ვალუტის კურსი', 'კურსი'],
  date: ['გამოწერის თარიღი', 'თარიღი'],
  invoiceNumber: [
    'ინვოისის ნომერი',
    'ინვოისის №',
    'ინვოისი №',
    'ინვოისი #',
    'ინვოისის #',
    'ნომერი',
  ],
};

function matchesAnyKeyword(text: string, keywords: string[]): boolean {
  const lower = String(text ?? '').toLowerCase().trim();
  if (!lower) return false;
  return keywords.some((kw) => lower.includes(kw.toLowerCase()));
}

function repairTruncatedPrefix(key: MainHeaderFieldKey, prefixText?: string): string {
  const clean = String(prefixText ?? '').trim();
  if (key === 'usdRate' && (clean === 'დოლარის' || clean === 'დოლარი' || !clean)) {
    return 'დოლარის კურსი';
  }
  if (key === 'companyName' && (clean === 'კომპანიის' || !clean)) {
    return 'კომპანიის სახელი';
  }
  if (key === 'invoiceNumber' && (clean === 'ინვოისის' || !clean)) {
    return 'ინვოისის ნომერი';
  }
  if (key === 'cadastralCode' && !clean) {
    return 'საკადასტრო';
  }
  if (key === 'address' && !clean) {
    return 'მისამართი';
  }
  if (key === 'date' && !clean) {
    return 'თარიღი';
  }
  return clean;
}

/**
 * Automatically resolves the 6 main header field configurations (Company Name, Cadastral Code,
 * Address, USD Rate, Date, Invoice Number) from saved instructions and/or by scanning the Excel sheet's cells.
 */
export function resolveMainHeaderFieldConfigs(
  sheetModel: ParsedSheetModel | null,
  savedInstructions: EditableCellInstruction[],
  explicitConfigs?: MainHeaderFieldConfig[] | null
): MainHeaderFieldConfig[] {
  if (explicitConfigs && explicitConfigs.length === 6) {
    return explicitConfigs.map((cfg) => ({
      ...cfg,
      placeholder: '',
      prefixText: repairTruncatedPrefix(cfg.key, cfg.prefixText),
    }));
  }

  const usedAddresses = new Set<string>();
  const resolved: Partial<Record<MainHeaderFieldKey, MainHeaderFieldConfig>> = {};

  // 1. Check savedInstructions first
  for (const def of DEFAULT_MAIN_HEADER_CONFIGS) {
    const key = def.key;
    const keywords = FIELD_KEYWORDS[key];

    const matchInst = savedInstructions.find((inst) => {
      const addr = inst.address.toUpperCase();
      if (usedAddresses.has(addr)) return false;
      const coord = parseCellAddress(addr);
      if (coord && coord.row >= 17) return false; // Skip product rows 17+

      if (inst.mainFieldKey === key) return true;
      const cellText = String(
        sheetModel?.cells[addr]?.formattedValue || sheetModel?.cells[addr]?.value || ''
      );
      // Also check cell to the left in case inst.address is a value cell next to a label cell
      let leftText = '';
      if (coord && coord.col > 1 && sheetModel) {
        for (let c = coord.col - 1; c >= 1; c--) {
          const lAddr = `${colNumberToLetter(c)}${coord.row}`;
          const lCell = sheetModel.cells[lAddr];
          if (lCell && !lCell.isMergedSlave && String(lCell.value).trim() !== '') {
            leftText = String(lCell.formattedValue || lCell.value || '');
            break;
          }
        }
      }
      return (
        matchesAnyKeyword(inst.label, keywords) ||
        matchesAnyKeyword(inst.prefixText || '', keywords) ||
        matchesAnyKeyword(cellText, keywords) ||
        matchesAnyKeyword(leftText, keywords)
      );
    });

    if (matchInst) {
      const addr = matchInst.address.toUpperCase();
      usedAddresses.add(addr);
      resolved[key] = {
        ...def,
        address: addr,
        usePrefix:
          matchInst.usePrefix !== undefined ? Boolean(matchInst.usePrefix) : def.usePrefix,
        prefixText: repairTruncatedPrefix(key, matchInst.prefixText || def.prefixText),
        affixPosition: matchInst.affixPosition || 'start',
      };
    }
  }

  // 2. Scan sheetModel.cells (rows 1..16) for any remaining unresolved fields
  if (sheetModel) {
    // First collect all non-slave cells in rows 1..16
    const headerCells = Object.values(sheetModel.cells).filter(
      (c) => !c.isMergedSlave && c.row >= 1 && c.row <= 16
    );

    // Helper to check if a cell's text matches ANY of the 6 field keywords
    const isAnyFieldLabelCell = (text: string): boolean => {
      return Object.values(FIELD_KEYWORDS).some((kws) => matchesAnyKeyword(text, kws));
    };

    for (const def of DEFAULT_MAIN_HEADER_CONFIGS) {
      if (resolved[def.key]) continue;
      const keywords = FIELD_KEYWORDS[def.key];

      const foundCell = headerCells.find((c) => {
        if (usedAddresses.has(c.address)) return false;
        const txt = String(c.formattedValue || c.value || '').trim();
        if (!txt) return false;
        // Avoid matching Seller static row ("მიმწოდებელი:") when searching for buyer company/id
        if (
          (def.key === 'companyName' || def.key === 'cadastralCode') &&
          txt.includes('მიმწოდებელი')
        ) {
          return false;
        }
        return matchesAnyKeyword(txt, keywords);
      });

      if (foundCell) {
        const cellTxt = String(foundCell.formattedValue || foundCell.value || '').trim();
        // Check if there is an adjacent cell to the right on the same row that holds a separate value
        const rightCol = foundCell.col + (foundCell.colSpan || 1);
        const rightAddr = `${colNumberToLetter(rightCol)}${foundCell.row}`;
        const rightCell =
          rightCol <= sheetModel.colCount ? sheetModel.cells[rightAddr] : undefined;
        const rightTxt = String(rightCell?.formattedValue || rightCell?.value || '').trim();

        const rightHasSeparateValue =
          rightCell &&
          !rightCell.isMergedSlave &&
          !usedAddresses.has(rightAddr) &&
          rightTxt !== '' &&
          !isAnyFieldLabelCell(rightTxt);

        if (rightHasSeparateValue && rightCell) {
          usedAddresses.add(rightAddr);
          const rightStartsWithLabel = matchesAnyKeyword(rightTxt, keywords);
          resolved[def.key] = {
            ...def,
            address: rightAddr,
            usePrefix: def.key === 'address' ? true : rightStartsWithLabel,
            prefixText: def.prefixText,
            affixPosition: 'start',
          };
        } else {
          usedAddresses.add(foundCell.address);
          const candidate = detectCellPrefixCandidate(cellTxt);
          resolved[def.key] = {
            ...def,
            address: foundCell.address,
            usePrefix: true,
            prefixText: repairTruncatedPrefix(
              def.key,
              candidate.suggestedPrefix || def.prefixText
            ),
            affixPosition: 'start',
          };
        }
      }
    }
  }

  return DEFAULT_MAIN_HEADER_CONFIGS.map((def) => resolved[def.key] || def);
}

/**
 * Builds the combined workbook overrides, affix rules, and currency/number formatting rules.
 * Ensures that any cell configured as 'number' ("რიცხვი") — whether as a direct number cell
 * or a formula cell with formulaResultFormat === 'number' (such as D36 in the template) —
 * keeps `{ mode: 'number' }` and never has the Lari symbol ('₾') automatically appended.
 */
export function buildCombinedWorkbookPayload(
  instList: EditableCellInstruction[],
  valsMap: Record<string, string | number>,
  hdrConfigs: MainHeaderFieldConfig[],
  hdrValues: Record<MainHeaderFieldKey, string>,
  prodRows: ProductRowItem[],
  totalCellAddr: string,
  currentSheet: ParsedSheetModel | null,
  maxProductRows = 8
): {
  effectiveOverrides: Record<string, string | number>;
  affixRules: Record<string, { text: string; position: 'start' | 'end' }>;
  currencyRules: Record<
    string,
    { mode: 'number' | 'currency'; symbol?: CurrencySymbol }
  >;
} {
  const effectiveOverrides: Record<string, string | number> = {};
  const affixRules: Record<string, { text: string; position: 'start' | 'end' }> = {};
  const currencyRules: Record<
    string,
    { mode: 'number' | 'currency'; symbol?: CurrencySymbol }
  > = {};

  // 1. Apply Settings instructions first
  for (const inst of instList) {
    const addr = inst.address.toUpperCase();
    const rawVal =
      inst.inputType === 'formula'
        ? valsMap[addr] ?? inst.formulaExpression ?? ''
        : valsMap[addr];
    if (rawVal !== undefined) {
      effectiveOverrides[addr] = composeCellValueWithPrefix(
        rawVal,
        inst.usePrefix,
        inst.prefixText,
        inst.inputType,
        inst.optional,
        inst.affixPosition || 'start'
      );
    }
    if (inst.usePrefix && inst.prefixText && inst.prefixText.trim() !== '') {
      affixRules[addr] = {
        text: inst.prefixText.trim(),
        position: inst.affixPosition || 'start',
      };
    }
    if (inst.inputType === 'number') {
      currencyRules[addr] = { mode: 'number' };
    } else if (inst.inputType === 'currency') {
      currencyRules[addr] = {
        mode: 'currency',
        symbol: inst.currencySymbol || '₾',
      };
    } else if (inst.inputType === 'formula') {
      const fmt = inst.formulaResultFormat || 'number';
      currencyRules[addr] =
        fmt === 'currency'
          ? { mode: 'currency', symbol: inst.currencySymbol || '₾' }
          : { mode: 'number' };
    }
  }

  // 2. Apply the 6 Main Header fields
  for (const cfg of hdrConfigs) {
    const addr = cfg.address.trim().toUpperCase();
    if (!addr) continue;
    const rawInput = hdrValues[cfg.key] ?? '';
    effectiveOverrides[addr] = composeCellValueWithPrefix(
      rawInput,
      cfg.usePrefix,
      cfg.prefixText,
      cfg.inputType,
      false,
      cfg.affixPosition || 'start'
    );
    if (cfg.usePrefix && cfg.prefixText && cfg.prefixText.trim() !== '') {
      affixRules[addr] = {
        text: cfg.prefixText.trim(),
        position: cfg.affixPosition || 'start',
      };
      // When prefix text is present in the same cell (e.g. "დოლარის კურსი 2.6"), do not force numeric-only format
      delete currencyRules[addr];
    } else if (cfg.inputType === 'number') {
      currencyRules[addr] = { mode: 'number' };
    }
  }

  // 3. Apply the 8 Product Rows (A17:C17 .. A24:C24) & GEL conversions
  const calcSummary = calculateProductsSummary(prodRows, hdrValues.usdRate);
  const hasColDRowTotals = Boolean(
    currentSheet?.cells['D17']?.formula ||
      String(currentSheet?.cells['D16']?.value ?? '').trim() !== '' ||
      String(currentSheet?.cells['D36']?.formula ?? '')
        .toUpperCase()
        .includes('D17') ||
      instList.some((inst) => inst.address.toUpperCase() === 'D17')
  );

  const effectiveRate = calcSummary.hasValidUsdRate
    ? calcSummary.usdRateNumeric
    : 1;

  for (let i = 0; i < maxProductRows; i++) {
    const r = 17 + i;
    const aAddr = `A${r}`;
    const bAddr = `B${r}`;
    const cAddr = `C${r}`;
    const dAddr = `D${r}`;

    if (i < prodRows.length) {
      const rowItem = prodRows[i];
      const calcRow = calcSummary.rows[i];
      const cleanName = rowItem.productName.trim();
      const cleanQty = rowItem.quantity.trim().replace(/,/g, '.');
      const cleanPrice = rowItem.unitPrice.trim().replace(/,/g, '.');

      effectiveOverrides[aAddr] = cleanName;
      effectiveOverrides[bAddr] =
        cleanQty !== '' && !Number.isNaN(Number(cleanQty))
          ? Number(cleanQty)
          : cleanQty;
      if (!currencyRules[bAddr]) {
        currencyRules[bAddr] = { mode: 'number' };
      }

      effectiveOverrides[cAddr] =
        cleanPrice !== '' && !Number.isNaN(Number(cleanPrice))
          ? Number(cleanPrice)
          : cleanPrice;
      if (currencyRules[cAddr]?.mode === 'number' && calcRow.currency !== '$') {
        currencyRules[cAddr] = { mode: 'number' };
      } else {
        currencyRules[cAddr] = {
          mode: 'currency',
          symbol: calcRow.currency, // '₾' (default) or '$'
        };
      }

      if (hasColDRowTotals) {
        if (cleanPrice !== '' || cleanQty !== '') {
          effectiveOverrides[dAddr] =
            calcRow.currency === '$'
              ? `=B${r}*C${r}*${effectiveRate}`
              : `=B${r}*C${r}`;
          if (!currencyRules[dAddr]) {
            currencyRules[dAddr] = { mode: 'currency', symbol: '₾' };
          }
        } else {
          effectiveOverrides[dAddr] = '';
        }
      }
    } else {
      // Clear unused product rows (e.g. rows 18..24 when only 1 product is active)
      effectiveOverrides[aAddr] = '';
      effectiveOverrides[bAddr] = '';
      effectiveOverrides[cAddr] = '';
      if (hasColDRowTotals) {
        effectiveOverrides[dAddr] = '';
      }
    }
  }

  // 4. Apply Grand Total formula to D36 (or configured total cell),
  // while respecting the user's configured format ('number' vs 'currency') in Settings!
  const cleanTotalAddr = (totalCellAddr || 'D36').trim().toUpperCase();
  if (cleanTotalAddr) {
    effectiveOverrides[cleanTotalAddr] = calcSummary.grandTotalFormula;
    if (!currencyRules[cleanTotalAddr]) {
      currencyRules[cleanTotalAddr] = { mode: 'currency', symbol: '₾' };
    }
  }

  return {
    effectiveOverrides,
    affixRules,
    currencyRules,
  };
}


