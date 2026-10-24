import test from 'node:test';
import assert from 'node:assert/strict';
import {
  colNumberToLetter,
  letterToColNumber,
  parseCellAddress,
  expandRange,
  evaluateExcelFormula,
  recalculateSheetCells,
  formatExcelCellValue,
  formatCurrencyCellValue,
  composeCellValueWithPrefix,
  detectCellPrefixCandidate,
  stripPrefixFromCellValue,
  isoDateToGeorgianDot,
  georgianDotToIsoDate,
  getTodayGeorgianDate,
  DEFAULT_MAIN_HEADER_CONFIGS,
  calculateProductsSummary,
  resolveMainHeaderFieldConfigs,
  syncFileNameWithCompanyName,
  resolveInvoiceFileNameBase,
  buildCombinedWorkbookPayload,
  applyAffixToDisplayValue,
} from './formulaEngine';
import {
  extractThemePaletteFromXml,
  resolveExcelColor,
  applyExcelTint,
} from './excelColorUtils';

test('colNumberToLetter and letterToColNumber convert accurately', () => {
  assert.equal(colNumberToLetter(1), 'A');
  assert.equal(colNumberToLetter(6), 'F');
  assert.equal(colNumberToLetter(26), 'Z');
  assert.equal(colNumberToLetter(27), 'AA');
  assert.equal(letterToColNumber('A'), 1);
  assert.equal(letterToColNumber('F'), 6);
  assert.equal(letterToColNumber('AA'), 27);
});

test('date conversion helpers convert between YYYY-MM-DD and DD.MM.YYYY and getTodayGeorgianDate formats DD.MM.YYYY', () => {
  assert.equal(isoDateToGeorgianDot('2026-10-23'), '23.10.2026');
  assert.equal(georgianDotToIsoDate('23.10.2026'), '2026-10-23');
  assert.equal(georgianDotToIsoDate('23/10/2026'), '2026-10-23');
  assert.equal(getTodayGeorgianDate(new Date(2026, 9, 5)), '05.10.2026');
  for (const cfg of DEFAULT_MAIN_HEADER_CONFIGS) {
    assert.equal(cfg.placeholder, '');
  }
});

test('composeCellValueWithPrefix handles prefix words (start), suffix words (end), dates, and user formulas', () => {
  assert.equal(
    composeCellValueWithPrefix('გორის ქუჩა 2', true, 'მისამართი', 'text', false, 'start'),
    'მისამართი გორის ქუჩა 2'
  );
  // Mandatory word at the END of the cell (e.g. "10" + "ლარი" -> "10 ლარი")
  assert.equal(
    composeCellValueWithPrefix('10', true, 'ლარი', 'number', false, 'end'),
    '10 ლარი'
  );
  assert.equal(
    stripPrefixFromCellValue('10 ლარი', true, 'ლარი', 'end'),
    '10'
  );
  assert.equal(
    composeCellValueWithPrefix('23.10.2026', true, 'თარიღი:', 'date'),
    'თარიღი: 23.10.2026'
  );
  assert.equal(
    composeCellValueWithPrefix('D12*E12', false, '', 'formula'),
    '=D12*E12'
  );
  assert.equal(
    composeCellValueWithPrefix('=SUM(F12:F15)*2.72', false, '', 'formula'),
    '=SUM(F12:F15)*2.72'
  );
  // Optional empty cell with prefix/suffix returns '' instead of dangling word
  assert.equal(
    composeCellValueWithPrefix('', true, 'მისამართი:', 'text', true, 'start'),
    ''
  );
  assert.equal(
    composeCellValueWithPrefix('', true, 'ლარი', 'number', true, 'end'),
    ''
  );
});

test('recalculateSheetCells evaluates custom user-entered formulas referencing other cells', () => {
  const initialCells = {
    D12: { value: 4 },
    E12: { value: 250 },
    B6: { value: 2.7 }, // Exchange rate
    G12: { value: 0 }, // Target cell where user writes a custom formula
    G13: { value: 0 }, // Chained formula cell
  };

  const updated = recalculateSheetCells(initialCells, {
    G12: '=D12*E12',
    G13: '=G12*B6',
  });

  assert.equal(updated.G12.formula, 'D12*E12');
  assert.equal(updated.G12.value, 1000);
  assert.equal(updated.G13.formula, 'G12*B6');
  assert.equal(updated.G13.value, 2700);
});

test('formulas ignore empty optional cells when left blank and include them when filled', () => {
  const initialCells = {
    D12: { value: 2 },
    E12: { value: 300 },
    D13: { value: 1 },
    E13: { value: 400 },
    D14: { value: '' }, // Optional row left empty
    E14: { value: '' }, // Optional row left empty
    B6: { value: '' }, // Optional multiplier/exchange rate left empty
    F12: { value: 0, formula: 'D12*E12' },
    F13: { value: 0, formula: 'D13*E13' },
    F14: { value: 0, formula: 'D14*E14' },
    F16: { value: 0, formula: 'SUM(F12:F14)' },
    F18: { value: 0, formula: 'F16*B6' },
  };

  // Case 1: Optional B6 and D14/E14 are empty
  const whenEmpty = recalculateSheetCells(initialCells, {
    D14: '',
    E14: '',
    B6: '',
  });
  assert.equal(whenEmpty.F12.value, 600);
  assert.equal(whenEmpty.F13.value, 400);
  assert.equal(whenEmpty.F14.value, ''); // Both D14 & E14 empty -> F14 is empty
  assert.equal(whenEmpty.F16.value, 1000); // SUM ignores empty F14
  assert.equal(whenEmpty.F18.value, 1000); // F16*B6 ignores empty B6 instead of zeroing out!

  // Case 2: Optional B6 is filled with 2.5, and optional E14 is filled with 150 (while D14 is left empty)
  const whenFilled = recalculateSheetCells(initialCells, {
    D14: '',
    E14: 150,
    B6: 2.5,
  });
  assert.equal(whenFilled.F14.value, 150); // D14 is empty, E14=150 -> 150
  assert.equal(whenFilled.F16.value, 1150); // 600 + 400 + 150
  assert.equal(whenFilled.F18.value, 2875); // 1150 * 2.5
});

test('detectCellPrefixCandidate preserves multi-word Georgian prefixes without truncating', () => {
  assert.deepEqual(detectCellPrefixCandidate('მისამართი: თბილისი, წერეთლის #114'), {
    suggestedPrefix: 'მისამართი:',
    remainder: 'თბილისი, წერეთლის #114',
  });
  assert.deepEqual(detectCellPrefixCandidate('დოლარის კურსი 2.72'), {
    suggestedPrefix: 'დოლარის კურსი',
    remainder: '2.72',
  });
  assert.deepEqual(detectCellPrefixCandidate('კომპანიის სახელი შპს ალფა'), {
    suggestedPrefix: 'კომპანიის სახელი',
    remainder: 'შპს ალფა',
  });
  assert.equal(
    stripPrefixFromCellValue('მისამართი გორის ქუჩა 2', true, 'მისამართი'),
    'გორის ქუჩა 2'
  );
});

test('resolveExcelColor resolves ARGB, theme indices with tint, and indexed colors', () => {
  const palette = extractThemePaletteFromXml();
  assert.equal(resolveExcelColor({ argb: 'FF1F497D' }, palette), '#1F497D');
  assert.equal(resolveExcelColor({ theme: 0 }, palette), '#FFFFFF');
  assert.equal(resolveExcelColor({ theme: 1 }, palette), '#000000');
  assert.equal(resolveExcelColor({ theme: 4 }, palette), '#4472C4');
  assert.equal(applyExcelTint('000000', 0.5), '#808080');
});

test('formatExcelCellValue never injects a spurious $ sign from Excel locale numFmt tags, while formatCurrencyCellValue formats Lari (₾ default) and Dollar ($)', () => {
  assert.equal(formatExcelCellValue(166, '[$-409]#,##0'), '166');
  assert.equal(formatExcelCellValue(166, '$#,##0'), '166');
  assert.equal(formatExcelCellValue(9676, '#,##0.00 "₾"'), '9,676.00 ₾');
  assert.equal(formatExcelCellValue(2.715, 'General'), '2.715');
  assert.equal(formatExcelCellValue('INV-001'), 'INV-001');

  // Default currency symbol is Lari ('₾')
  assert.equal(formatCurrencyCellValue(100), '100 ₾');
  assert.equal(formatCurrencyCellValue(2400, '₾', '#,##0.00 "₾"'), '2,400 ₾');
  assert.equal(formatCurrencyCellValue(100.5, '₾'), '100.50 ₾');
  // Selectable Dollar ('$') currency symbol
  assert.equal(formatCurrencyCellValue(100, '$'), '100 $');
  assert.equal(formatCurrencyCellValue(2400, '$', '#,##0.00 "$"'), '2,400 $');
  assert.equal(formatCurrencyCellValue(100.5, '$'), '100.50 $');
});

test('calculateProductsSummary maps products to A17:C24 and always computes Grand Total in Lari (₾), multiplying Dollar ($) prices by USD rate', () => {
  const summary = calculateProductsSummary(
    [
      {
        id: 'p1',
        productName: 'სერვერი',
        quantity: '2',
        unitPrice: '100',
        currency: '$', // 2 * 100 $ * 2.7 = 540 ₾
      },
      {
        id: 'p2',
        productName: 'მონიტორი',
        quantity: '3',
        unitPrice: '80',
        currency: '₾', // 3 * 80 ₾ = 240 ₾
      },
    ],
    '2.7'
  );

  assert.equal(summary.rows.length, 2);
  assert.equal(summary.rows[0].productCell, 'A17');
  assert.equal(summary.rows[0].qtyCell, 'B17');
  assert.equal(summary.rows[0].priceCell, 'C17');
  assert.equal(summary.rows[0].lineTotalOriginal, 200);
  assert.equal(summary.rows[0].formattedOriginal, '200 $');
  assert.equal(summary.rows[0].lineTotalGel, 540);
  assert.equal(summary.rows[0].formattedGel, '540 ₾');

  assert.equal(summary.rows[1].productCell, 'A18');
  assert.equal(summary.rows[1].qtyCell, 'B18');
  assert.equal(summary.rows[1].priceCell, 'C18');
  assert.equal(summary.rows[1].lineTotalGel, 240);
  assert.equal(summary.rows[1].formattedGel, '240 ₾');

  assert.equal(summary.grandTotalGel, 780);
  assert.equal(summary.formattedGrandTotalGel, '780 ₾');
  assert.equal(summary.grandTotalFormula, '=B17*C17*2.7+B18*C18');
});

test('resolveMainHeaderFieldConfigs repairs truncated prefixes and resolves 6 main fields', () => {
  const resolved = resolveMainHeaderFieldConfigs(null, [
    {
      id: '1',
      address: 'B5',
      label: 'დოლარის კურსი',
      group: 'header',
      inputType: 'number',
      usePrefix: true,
      prefixText: 'დოლარის', // Previously truncated in Turn 1
    },
    {
      id: '2',
      address: 'A6',
      label: 'კომპანიის სახელი',
      group: 'client',
      inputType: 'text',
      usePrefix: true,
      prefixText: 'კომპანიის', // Previously truncated in Turn 1
    },
  ]);

  const usdField = resolved.find((f) => f.key === 'usdRate');
  const compField = resolved.find((f) => f.key === 'companyName');
  assert.equal(usdField?.address, 'B5');
  assert.equal(usdField?.prefixText, 'დოლარის კურსი');
  assert.equal(compField?.address, 'A6');
  assert.equal(compField?.prefixText, 'კომპანიის სახელი');
});

test('syncFileNameWithCompanyName defaults fileName to companyName until manually customized', () => {
  // When not customized, typing "შპს გია" into companyName sets fileName to "შპს გია"
  const step1 = syncFileNameWithCompanyName('', 'შპს გია', '', false);
  assert.deepEqual(step1, {
    fileName: 'შპს გია',
    isFileNameCustomized: false,
  });

  // Continuing to edit companyName updates fileName automatically when not customized
  const step2 = syncFileNameWithCompanyName(
    'შპს გია',
    'შპს გია ჯგუფი',
    'შპს გია',
    false
  );
  assert.deepEqual(step2, {
    fileName: 'შპს გია ჯგუფი',
    isFileNameCustomized: false,
  });

  // When user manually customized fileName (e.g. "ინვოისი 01"), changing companyName preserves custom fileName
  const step3 = syncFileNameWithCompanyName(
    'შპს გია ჯგუფი',
    'შპს ახალი',
    'ინვოისი 01',
    true
  );
  assert.deepEqual(step3, {
    fileName: 'ინვოისი 01',
    isFileNameCustomized: true,
  });

  // resolveInvoiceFileNameBase produces "შპს გია" so PDF becomes "შპს გია.pdf"
  assert.equal(resolveInvoiceFileNameBase('შპს გია', 'შპს გია'), 'შპს გია');
  assert.equal(resolveInvoiceFileNameBase('', 'შპს გია'), 'შპს გია');
  assert.equal(resolveInvoiceFileNameBase('შპს გია.pdf', 'შპს გია'), 'შპს გია');
  assert.equal(resolveInvoiceFileNameBase('ჩემი_ფაილი', 'შპს გია'), 'ჩემი_ფაილი');
  assert.equal(resolveInvoiceFileNameBase('', ''), 'ინვოისი');
});

test('buildCombinedWorkbookPayload keeps mode: "number" for cells marked as რიცხვი (e.g. D36) so ₾ is never added automatically', async () => {
  const { buildCombinedWorkbookPayload, applyAffixToDisplayValue } = await import('./formulaEngine');
  const instructions = [
    {
      id: 'custom-D36',
      address: 'D36',
      label: 'ჯამი',
      group: 'other' as const,
      inputType: 'formula' as const,
      usePrefix: true,
      prefixText: ' ლარი',
      affixPosition: 'end' as const,
      formulaExpression: '=D17+D18+D19+D20+D21+D22+D23+D24',
      formulaResultFormat: 'number' as const,
      currencySymbol: '₾' as const,
    },
    {
      id: 'custom-D17',
      address: 'D17',
      label: 'უჯრა D17',
      group: 'other' as const,
      inputType: 'formula' as const,
      usePrefix: false,
      prefixText: '',
      formulaExpression: '=B17*C17',
      // formulaResultFormat omitted (defaults to 'number' / რიცხვი in UI)
      currencySymbol: '₾' as const,
    },
  ];

  const payload = buildCombinedWorkbookPayload(
    instructions,
    {},
    DEFAULT_MAIN_HEADER_CONFIGS,
    {
      companyName: 'შპს ტესტი',
      cadastralCode: '123456789',
      address: 'თბილისი',
      usdRate: '2.6',
      date: '23.10.2026',
      invoiceNumber: 'INV-1',
    },
    [
      {
        id: 'prod-1',
        productName: 'მომსახურება',
        quantity: '2',
        unitPrice: '150',
        currency: '₾',
      },
    ],
    'D36',
    null
  );

  // D36 and D17 are marked as რიცხვი ('number'), so currencyRules must stay { mode: 'number' }
  assert.deepEqual(payload.currencyRules.D36, { mode: 'number' });
  assert.deepEqual(payload.currencyRules.D17, { mode: 'number' });
  assert.deepEqual(payload.affixRules.D36, { text: 'ლარი', position: 'end' });

  // Formatted display for D36 is plain number + suffix ("300 ლარი"), with NO '₾' symbol added!
  const numFormatted = formatExcelCellValue(300, 'General');
  const finalD36Display = applyAffixToDisplayValue(
    numFormatted,
    true,
    payload.affixRules.D36.text,
    payload.affixRules.D36.position
  );
  assert.equal(finalD36Display, '300 ლარი');
  assert.equal(finalD36Display.includes('₾'), false);
});

test('cells marked as number (რიცხვი) — including D36 — do NOT automatically receive the Lari symbol (₾)', () => {
  const instructions = [
    {
      id: 'd17-inst',
      address: 'D17',
      label: 'უჯრა D17',
      group: 'other' as const,
      inputType: 'formula' as const,
      formulaExpression: '=B17*C17',
      // formulaResultFormat omitted or 'number' -> means "რიცხვი"
    },
    {
      id: 'd36-inst',
      address: 'D36',
      label: 'ჯამი',
      group: 'other' as const,
      inputType: 'formula' as const,
      usePrefix: true,
      prefixText: ' ლარი',
      affixPosition: 'end' as const,
      formulaExpression: '=D17+D18+D19+D20+D21+D22+D23+D24',
      formulaResultFormat: 'number' as const,
      currencySymbol: '₾' as const,
    },
  ];

  const payload = buildCombinedWorkbookPayload(
    instructions,
    {},
    DEFAULT_MAIN_HEADER_CONFIGS,
    {
      companyName: 'შპს გია',
      cadastralCode: '123456789',
      address: 'თბილისი',
      usdRate: '2.6',
      date: '23.10.2026',
      invoiceNumber: '01',
    },
    [
      {
        id: 'prod-1',
        productName: 'პროდუქტი 1',
        quantity: '2',
        unitPrice: '50',
        currency: '₾',
      },
    ],
    'D36',
    null
  );

  // D17 and D36 must remain mode: 'number' (რიცხვი) and NOT be overwritten with mode: 'currency'
  assert.deepEqual(payload.currencyRules['D17'], { mode: 'number' });
  assert.deepEqual(payload.currencyRules['D36'], { mode: 'number' });
  assert.deepEqual(payload.affixRules['D36'], { text: 'ლარი', position: 'end' });

  // C17 (unit price) has currency '₾', so it gets mode: 'currency'
  assert.deepEqual(payload.currencyRules['C17'], { mode: 'currency', symbol: '₾' });

  // Verify final formatted string for D36 has no '₾' (becomes "100 ლარი", not "100 ₾ ლარი")
  const baseNumberStr = formatExcelCellValue(100, 'General');
  const finalD36Display = applyAffixToDisplayValue(
    baseNumberStr,
    true,
    payload.affixRules['D36'].text,
    payload.affixRules['D36'].position
  );
  assert.equal(finalD36Display, '100 ლარი');
  assert.equal(finalD36Display.includes('₾'), false);
});

test('product cells (A17:A24) keep numeric input as string so they always align to the left like text', () => {
  const payload = buildCombinedWorkbookPayload(
    [],
    {},
    DEFAULT_MAIN_HEADER_CONFIGS,
    {
      companyName: 'შპს გია',
      cadastralCode: '123456789',
      address: 'თბილისი',
      usdRate: '2.6',
      date: '24.10.2026',
      invoiceNumber: '01',
    },
    [
      {
        id: 'prod-1',
        productName: '12345',
        quantity: '2',
        unitPrice: '50',
        currency: '₾',
      },
    ],
    'D36',
    null
  );

  assert.equal(typeof payload.effectiveOverrides['A17'], 'string');
  assert.equal(payload.effectiveOverrides['A17'], '12345');

  const recalculated = recalculateSheetCells(
    { A17: { value: '' }, B17: { value: 1 }, C17: { value: 0 } },
    payload.effectiveOverrides
  );

  // A17 must remain a string ('12345'), NOT coerced to number 12345 (which would trigger right-alignment)
  assert.equal(typeof recalculated['A17'].value, 'string');
  assert.equal(recalculated['A17'].value, '12345');
  // Meanwhile B17 and C17 are numbers
  assert.equal(typeof recalculated['B17'].value, 'number');
  assert.equal(typeof recalculated['C17'].value, 'number');
});






