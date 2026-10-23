import React, { useState } from 'react';
import { Plus, Trash2, Lock, RotateCcw } from 'lucide-react';
import {
  AffixPosition,
  CellInputType,
  CurrencySymbol,
  EditableCellInstruction,
  FormulaResultFormat,
  ParsedSheetModel,
} from '../types/invoice';
import {
  applyAffixToDisplayValue,
  detectCellPrefixCandidate,
  evaluateExcelFormula,
  formatCurrencyCellValue,
  parseCellAddress,
} from '../utils/formulaEngine';

interface InstructionsConfigPanelProps {
  instructions: EditableCellInstruction[];
  sheetModel: ParsedSheetModel | null;
  selectedAddress: string | null;
  editableValues: Record<string, string | number>;
  liveCellValueLookup: Record<string, string | number>;
  onSelectAddress: (addr: string) => void;
  onAddInstruction: (inst: EditableCellInstruction, initialUserValue?: string | number) => void;
  onUpdateInstruction: (id: string, patch: Partial<EditableCellInstruction>) => void;
  onChangeCellValue: (address: string, rawVal: string, inputType: CellInputType) => void;
  onRemoveInstruction: (id: string) => void;
  onClearAllInstructions: () => void;
  onResetDefaultInstructions: () => void;
}

export const InstructionsConfigPanel: React.FC<InstructionsConfigPanelProps> = ({
  instructions,
  sheetModel,
  selectedAddress,
  editableValues,
  liveCellValueLookup,
  onSelectAddress,
  onAddInstruction,
  onUpdateInstruction,
  onChangeCellValue,
  onRemoveInstruction,
  onClearAllInstructions,
  onResetDefaultInstructions,
}) => {
  const [newAddress, setNewAddress] = useState('');
  const [newLabel, setNewLabel] = useState('');
  const [newType, setNewType] = useState<CellInputType>('text');
  const [newCurrencySymbol, setNewCurrencySymbol] = useState<CurrencySymbol>('₾');
  const [newFormulaResultFormat, setNewFormulaResultFormat] =
    useState<FormulaResultFormat>('currency');
  const [newFormulaValue, setNewFormulaValue] = useState('');
  const [newUsePrefix, setNewUsePrefix] = useState(false);
  const [newPrefixText, setNewPrefixText] = useState('');
  const [newAffixPosition, setNewAffixPosition] = useState<AffixPosition>('start');
  const [newOptional, setNewOptional] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  // Resolve master address if user clicked a merged slave cell
  const effectiveSelectedAddress = React.useMemo(() => {
    if (!selectedAddress || !sheetModel) return selectedAddress;
    const cell = sheetModel.cells[selectedAddress];
    if (cell?.isMergedSlave && cell.masterAddress) {
      return cell.masterAddress;
    }
    return selectedAddress;
  }, [selectedAddress, sheetModel]);

  const selectedCellObj =
    effectiveSelectedAddress && sheetModel
      ? sheetModel.cells[effectiveSelectedAddress]
      : null;

  const existingInstructionForSelected = instructions.find(
    (i) => i.address.toUpperCase() === effectiveSelectedAddress?.toUpperCase()
  );

  const prefixCandidate = React.useMemo(() => {
    const rawStr = String(selectedCellObj?.formattedValue || selectedCellObj?.value || '');
    return detectCellPrefixCandidate(rawStr);
  }, [selectedCellObj]);

  const handleManualAdd = (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMsg('');
    const parsed = parseCellAddress(newAddress);
    if (!parsed) {
      setErrorMsg('გთხოვთ მიუთითოთ სწორი უჯრის კოორდინატი (მაგ: C7, F3, D12)');
      return;
    }
    const targetAddr =
      sheetModel?.cells[parsed.address]?.masterAddress || parsed.address;

    if (!newLabel.trim()) {
      setErrorMsg('გთხოვთ მიუთითოთ უჯრის დასახელება');
      return;
    }
    if (instructions.some((i) => i.address.toUpperCase() === targetAddr)) {
      setErrorMsg(`უჯრა ${targetAddr} უკვე დამატებულია შესაცვლელ სიაში`);
      return;
    }

    const isFormula = newType === 'formula';
    const initialFormulaStr = isFormula
      ? newFormulaValue.trim()
        ? newFormulaValue.trim().startsWith('=')
          ? newFormulaValue.trim()
          : `=${newFormulaValue.trim()}`
        : sheetModel?.cells[targetAddr]?.formula
        ? `=${sheetModel.cells[targetAddr].formula}`
        : '='
      : undefined;

    onAddInstruction(
      {
        id: `custom-${targetAddr}-${Date.now()}`,
        address: targetAddr,
        label: newLabel.trim(),
        group: 'other',
        inputType: newType,
        currencySymbol:
          newType === 'currency' ||
          (isFormula && newFormulaResultFormat === 'currency')
            ? newCurrencySymbol
            : undefined,
        formulaResultFormat: isFormula ? newFormulaResultFormat : undefined,
        usePrefix: newUsePrefix && Boolean(newPrefixText.trim()),
        prefixText: newPrefixText.trim(),
        affixPosition: newAffixPosition,
        optional: isFormula ? false : newOptional,
        formulaExpression: isFormula ? initialFormulaStr : undefined,
      },
      initialFormulaStr
    );
    setNewAddress('');
    setNewLabel('');
    setNewCurrencySymbol('₾');
    setNewFormulaResultFormat('currency');
    setNewFormulaValue('');
    setNewUsePrefix(false);
    setNewPrefixText('');
    setNewAffixPosition('start');
    setNewOptional(false);
  };

  const handleAddSelectedCell = (
    withPrefix: boolean,
    explicitType?: CellInputType,
    affixPos: AffixPosition = 'start'
  ) => {
    if (!effectiveSelectedAddress) return;
    if (existingInstructionForSelected) {
      onRemoveInstruction(existingInstructionForSelected.id);
      return;
    }
    const cellVal = selectedCellObj?.value ?? '';
    const existingFormula = selectedCellObj?.formula;
    const hasGelFmt =
      selectedCellObj?.numFmt?.includes('₾') ||
      String(selectedCellObj?.formattedValue || '').includes('₾');
    const guessedType: CellInputType =
      explicitType ||
      (existingFormula
        ? 'formula'
        : !withPrefix && typeof cellVal === 'number'
        ? hasGelFmt
          ? 'currency'
          : 'number'
        : 'text');

    const isFormula = guessedType === 'formula';
    const pref = withPrefix
      ? affixPos === 'end'
        ? 'ლარი'
        : prefixCandidate.suggestedPrefix || 'მისამართი'
      : '';
    const initVal = isFormula
      ? existingFormula
        ? `=${existingFormula}`
        : '='
      : withPrefix && affixPos === 'start'
      ? prefixCandidate.remainder
      : cellVal;
    const cleanLabel =
      pref.replace(/:$/, '').trim() && affixPos === 'start'
        ? pref.replace(/:$/, '').trim()
        : `უჯრა ${effectiveSelectedAddress}`;

    onAddInstruction(
      {
        id: `cell-${effectiveSelectedAddress}-${Date.now()}`,
        address: effectiveSelectedAddress,
        label: cleanLabel,
        group: 'other',
        inputType: guessedType,
        currencySymbol: guessedType === 'currency' || isFormula ? '₾' : undefined,
        formulaResultFormat: isFormula ? (hasGelFmt ? 'currency' : 'number') : undefined,
        usePrefix: withPrefix,
        prefixText: pref,
        affixPosition: affixPos,
        formulaExpression: isFormula ? String(initVal) : undefined,
      },
      initVal
    );
  };

  // Count total populated cells, editable cells, formula cells, and unchanged static cells
  const stats = React.useMemo(() => {
    if (!sheetModel) {
      return { editable: instructions.length, formula: 0, unchanged: 0 };
    }
    let formula = 0;
    let unchanged = 0;
    const editSet = new Set(instructions.map((i) => i.address.toUpperCase()));

    for (const cell of Object.values(sheetModel.cells)) {
      if (cell.isMergedSlave) continue;
      if (cell.formula) {
        formula++;
      } else if (!editSet.has(cell.address) && String(cell.value).trim() !== '') {
        unchanged++;
      }
    }
    return {
      editable: instructions.length,
      formula,
      unchanged,
    };
  }, [sheetModel, instructions]);

  return (
    <div className="space-y-5">
      {/* Summary Strip */}
      <div className="grid grid-cols-3 gap-2.5 bg-slate-50 p-3 rounded-lg border border-slate-200 text-xs">
        <div>
          <div className="text-slate-500">შესაცვლელი</div>
          <div className="text-base font-semibold text-sky-700 font-mono-tabular mt-0.5">
            {stats.editable} უჯრა
          </div>
        </div>
        <div>
          <div className="text-slate-500">უცვლელი</div>
          <div className="text-base font-semibold text-slate-800 font-mono-tabular mt-0.5">
            {stats.unchanged} უჯრა
          </div>
        </div>
        <div>
          <div className="text-slate-500">ფორმულები</div>
          <div className="text-base font-semibold text-emerald-700 font-mono-tabular mt-0.5">
            {stats.formula} უჯრა
          </div>
        </div>
      </div>

      {/* Selected Cell Inspector from Sheet Preview */}
      {effectiveSelectedAddress && (
        <div className="p-3.5 rounded-lg border border-slate-200 bg-white space-y-2.5">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-xs text-slate-600">
              <span className="font-mono font-semibold text-sm text-slate-900">
                უჯრა {effectiveSelectedAddress}
              </span>
              <span aria-hidden="true">·</span>
              <span>
                {existingInstructionForSelected
                  ? existingInstructionForSelected.inputType === 'formula'
                    ? 'ფორმულის უჯრა'
                    : 'შესაცვლელი (რედაქტირებადი)'
                  : selectedCellObj?.formula
                  ? 'ავტომატური ფორმულა'
                  : 'უცვლელი (დაცული)'}
              </span>
            </div>

            {existingInstructionForSelected ? (
              <button
                type="button"
                onClick={() => handleAddSelectedCell(false)}
                className="px-3 py-1.5 text-xs font-medium rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                უცვლელად დატოვება
              </button>
            ) : (
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => handleAddSelectedCell(false)}
                  className="px-2.5 py-1.5 text-xs font-medium rounded-md bg-sky-600 text-white hover:bg-sky-700 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  შესაცვლელად მონიშვნა
                </button>
                <button
                  type="button"
                  onClick={() => handleAddSelectedCell(true, undefined, 'start')}
                  className="px-2.5 py-1.5 text-xs font-medium rounded-md bg-slate-900 text-white hover:bg-slate-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  + სიტყვა თავში
                </button>
                <button
                  type="button"
                  onClick={() => handleAddSelectedCell(true, undefined, 'end')}
                  className="px-2.5 py-1.5 text-xs font-medium rounded-md bg-slate-700 text-white hover:bg-slate-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  + სიტყვა ბოლოში
                </button>
                <button
                  type="button"
                  onClick={() => handleAddSelectedCell(false, 'formula')}
                  className="px-2.5 py-1.5 text-xs font-medium rounded-md bg-emerald-700 text-white hover:bg-emerald-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  ფორმულის ჩაწერა (=)
                </button>
              </div>
            )}
          </div>

          <div className="text-xs text-slate-500 flex items-center justify-between gap-4 pt-1 border-t border-slate-100">
            <span className="break-words">
              მიმდინარე ტექსტი:{' '}
              <strong className="text-slate-800 font-mono-tabular">
                {selectedCellObj?.formattedValue || '— (ცარიელი)'}
              </strong>
            </span>
            {selectedCellObj?.formula && (
              <span className="font-mono text-emerald-700 shrink-0">
                ={selectedCellObj.formula}
              </span>
            )}
          </div>

          {/* Direct formula editor + Result Format (რიცხვი vs თანხა ₾/$) inside Selected Cell Inspector */}
          {existingInstructionForSelected &&
            existingInstructionForSelected.inputType === 'formula' && (
              <div className="pt-2 border-t border-emerald-100 space-y-2">
                <label
                  htmlFor={`inspector-formula-${effectiveSelectedAddress}`}
                  className="block text-[11px] font-medium text-emerald-900"
                >
                  ფორმულა უჯრა {effectiveSelectedAddress}-ისთვის (იწერება პირდაპირ აქ):
                </label>
                <input
                  id={`inspector-formula-${effectiveSelectedAddress}`}
                  type="text"
                  value={String(
                    editableValues[effectiveSelectedAddress.toUpperCase()] ??
                      existingInstructionForSelected.formulaExpression ??
                      '='
                  )}
                  onChange={(e) => {
                    const val = e.target.value;
                    onChangeCellValue(effectiveSelectedAddress, val, 'formula');
                    onUpdateInstruction(existingInstructionForSelected.id, {
                      formulaExpression: val,
                    });
                  }}
                  placeholder="მაგ: =D12*E12 ან =SUM(F12:F15)*B6"
                  className="w-full px-2.5 py-1.5 text-xs font-mono bg-emerald-50/50 border border-emerald-400 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-600 text-slate-900"
                />

                <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                  <span className="text-[11px] text-slate-600 font-medium">
                    ფორმულის შედეგის ტიპი:
                  </span>
                  <div className="flex items-center gap-1.5">
                    <select
                      aria-label={`უჯრა ${effectiveSelectedAddress} ფორმულის შედეგის ტიპი`}
                      value={existingInstructionForSelected.formulaResultFormat || 'number'}
                      onChange={(e) => {
                        const fmt = e.target.value as FormulaResultFormat;
                        onUpdateInstruction(existingInstructionForSelected.id, {
                          formulaResultFormat: fmt,
                          currencySymbol:
                            fmt === 'currency'
                              ? existingInstructionForSelected.currencySymbol || '₾'
                              : existingInstructionForSelected.currencySymbol,
                        });
                      }}
                      className="text-[11px] text-slate-800 bg-slate-100 border border-slate-300 rounded px-2 py-1"
                    >
                      <option value="number">რიცხვი</option>
                      <option value="currency">თანხა</option>
                    </select>

                    {(existingInstructionForSelected.formulaResultFormat || 'number') ===
                      'currency' && (
                      <div className="inline-flex items-center bg-slate-200/80 p-0.5 rounded">
                        <button
                          type="button"
                          onClick={() =>
                            onUpdateInstruction(existingInstructionForSelected.id, {
                              currencySymbol: '₾',
                            })
                          }
                          className={`px-2 py-0.5 text-[11px] font-semibold rounded cursor-pointer ${
                            (existingInstructionForSelected.currencySymbol || '₾') === '₾'
                              ? 'bg-sky-600 text-white'
                              : 'text-slate-700 hover:text-slate-900'
                          }`}
                        >
                          ₾ ლარი
                        </button>
                        <button
                          type="button"
                          onClick={() =>
                            onUpdateInstruction(existingInstructionForSelected.id, {
                              currencySymbol: '$',
                            })
                          }
                          className={`px-2 py-0.5 text-[11px] font-semibold rounded cursor-pointer ${
                            existingInstructionForSelected.currencySymbol === '$'
                              ? 'bg-sky-600 text-white'
                              : 'text-slate-700 hover:text-slate-900'
                          }`}
                        >
                          $ დოლარი
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
        </div>
      )}

      {/* Add New Editable Cell Instruction Form */}
      <form
        onSubmit={handleManualAdd}
        className="p-3.5 rounded-lg border border-slate-200 bg-slate-50/70 space-y-3"
      >
        <div className="text-xs font-semibold text-slate-800">
          ახალი შესაცვლელი უჯრის დამატება
        </div>
        <div className="grid grid-cols-12 gap-2">
          <div className="col-span-3">
            <label htmlFor="new-cell-coord" className="sr-only">
              უჯრის კოორდინატი
            </label>
            <input
              id="new-cell-coord"
              type="text"
              value={newAddress}
              onChange={(e) => setNewAddress(e.target.value.toUpperCase())}
              placeholder="მაგ: C9"
              className="w-full px-2.5 py-1.5 text-xs font-mono bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
            />
          </div>
          <div className="col-span-5">
            <label htmlFor="new-cell-label" className="sr-only">
              ველის დასახელება
            </label>
            <input
              id="new-cell-label"
              type="text"
              value={newLabel}
              onChange={(e) => setNewLabel(e.target.value)}
              placeholder="სახელი (მაგ: მისამართი)"
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
            />
          </div>
          <div className="col-span-4">
            <label htmlFor="new-cell-type" className="sr-only">
              მონაცემის ტიპი
            </label>
            <select
              id="new-cell-type"
              value={newType}
              onChange={(e) => {
                const nextT = e.target.value as CellInputType;
                setNewType(nextT);
                if (nextT === 'currency' || nextT === 'formula') {
                  setNewCurrencySymbol('₾');
                }
              }}
              className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
            >
              <option value="text">ტექსტი</option>
              <option value="number">რიცხვი</option>
              <option value="currency">თანხა</option>
              <option value="date">თარიღი</option>
              <option value="formula">ფორმულა (=)</option>
            </select>
          </div>
        </div>

        {/* Currency symbol selector when adding a 'currency' (თანხა) cell */}
        {newType === 'currency' && (
          <div className="pt-1 flex items-center justify-between gap-2 bg-white px-2.5 py-1.5 rounded-md border border-slate-200">
            <span className="text-xs text-slate-700 font-medium">
              ვალუტის სიმბოლო (Default: ლარი):
            </span>
            <div className="inline-flex items-center gap-1 bg-slate-100 p-0.5 rounded-md">
              <button
                type="button"
                onClick={() => setNewCurrencySymbol('₾')}
                className={`px-2.5 py-1 text-xs font-medium rounded transition-colors cursor-pointer ${
                  newCurrencySymbol === '₾'
                    ? 'bg-sky-600 text-white'
                    : 'text-slate-700 hover:text-slate-900'
                }`}
              >
                ₾ (ლარი)
              </button>
              <button
                type="button"
                onClick={() => setNewCurrencySymbol('$')}
                className={`px-2.5 py-1 text-xs font-medium rounded transition-colors cursor-pointer ${
                  newCurrencySymbol === '$'
                    ? 'bg-sky-600 text-white'
                    : 'text-slate-700 hover:text-slate-900'
                }`}
              >
                $ (დოლარი)
              </button>
            </div>
          </div>
        )}

        {/* Formula input & Result Format (რიცხვი vs თანხა ₾/$) when adding a formula cell */}
        {newType === 'formula' && (
          <div className="pt-1 space-y-2">
            <div className="space-y-1">
              <label htmlFor="new-cell-formula" className="block text-[11px] text-slate-600">
                ფორმულა (სხვა უჯრებიდან გამოსათვლელად):
              </label>
              <input
                id="new-cell-formula"
                type="text"
                value={newFormulaValue}
                onChange={(e) => setNewFormulaValue(e.target.value)}
                placeholder="მაგ: =D12*E12 ან =SUM(F12:F15)*B6"
                className="w-full px-2.5 py-1.5 text-xs font-mono bg-white border border-emerald-400 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-600"
              />
            </div>

            <div className="flex flex-wrap items-center justify-between gap-2 bg-white px-2.5 py-1.5 rounded-md border border-slate-200">
              <span className="text-xs text-slate-700 font-medium">
                ფორმულის შედეგი ჩაიწეროს როგორც:
              </span>
              <div className="flex items-center gap-2">
                <select
                  aria-label="ახალი ფორმულის შედეგის ტიპი"
                  value={newFormulaResultFormat}
                  onChange={(e) =>
                    setNewFormulaResultFormat(e.target.value as FormulaResultFormat)
                  }
                  className="px-2 py-1 text-xs bg-slate-100 border border-slate-300 rounded-md text-slate-800"
                >
                  <option value="number">რიცხვი (მაგ: 100)</option>
                  <option value="currency">თანხა (მაგ: 100 ₾ ან 100 $)</option>
                </select>

                {newFormulaResultFormat === 'currency' && (
                  <div className="inline-flex items-center gap-1 bg-slate-100 p-0.5 rounded-md">
                    <button
                      type="button"
                      onClick={() => setNewCurrencySymbol('₾')}
                      className={`px-2 py-1 text-xs font-semibold rounded transition-colors cursor-pointer ${
                        newCurrencySymbol === '₾'
                          ? 'bg-sky-600 text-white'
                          : 'text-slate-700 hover:text-slate-900'
                      }`}
                    >
                      ₾ ლარი
                    </button>
                    <button
                      type="button"
                      onClick={() => setNewCurrencySymbol('$')}
                      className={`px-2 py-1 text-xs font-semibold rounded transition-colors cursor-pointer ${
                        newCurrencySymbol === '$'
                          ? 'bg-sky-600 text-white'
                          : 'text-slate-700 hover:text-slate-900'
                      }`}
                    >
                      $ დოლარი
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {/* Mandatory Word ("სავალდებულო სიტყვა — თავში ან ბოლოში") & Optional ("არასავალდებულო") setup */}
        <div className="pt-1 space-y-2">
          <div className="flex flex-wrap items-center gap-4">
            <label className="inline-flex items-center gap-2 text-xs text-slate-700 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={newUsePrefix}
                onChange={(e) => setNewUsePrefix(e.target.checked)}
                className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
              />
              <span>სავალდებულო სიტყვა (თავში ან ბოლოში)</span>
            </label>

            {newType !== 'formula' && (
              <label className="inline-flex items-center gap-2 text-xs text-slate-700 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={newOptional}
                  onChange={(e) => setNewOptional(e.target.checked)}
                  className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                />
                <span>არასავალდებულო (Optional)</span>
              </label>
            )}
          </div>

          {newUsePrefix && (
            <div className="grid grid-cols-12 gap-2">
              <div className="col-span-4">
                <select
                  aria-label="სავალდებულო სიტყვის პოზიცია"
                  value={newAffixPosition}
                  onChange={(e) => setNewAffixPosition(e.target.value as AffixPosition)}
                  className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
                >
                  <option value="start">თავში (წინ)</option>
                  <option value="end">ბოლოში (უკან)</option>
                </select>
              </div>
              <div className="col-span-8">
                <input
                  type="text"
                  value={newPrefixText}
                  onChange={(e) => setNewPrefixText(e.target.value)}
                  placeholder={
                    newAffixPosition === 'end'
                      ? 'ჩაწერეთ ბოლო სიტყვა (მაგ: ლარი, ₾, კგ)'
                      : 'ჩაწერეთ საწყისი სიტყვა (მაგ: მისამართი)'
                  }
                  className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
                />
              </div>
            </div>
          )}
        </div>

        {errorMsg && <p className="text-xs text-red-600">{errorMsg}</p>}
        <button
          type="submit"
          className="w-full flex items-center justify-center gap-1.5 py-2 px-3 bg-slate-900 hover:bg-slate-800 text-white text-xs font-medium rounded-md transition-colors whitespace-nowrap cursor-pointer"
        >
          <Plus className="w-3.5 h-3.5" />
          <span>ინსტრუქციაში დამატება</span>
        </button>
      </form>

      {/* Editable Instructions List with Formula Editor, Optional Toggle, Currency & Prefix/Suffix Controls */}
      <div className="space-y-2">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <h3 className="text-xs font-semibold text-slate-700">
            შესაცვლელი უჯრების სია ({instructions.length})
          </h3>
          <div className="flex items-center gap-3">
            {instructions.length > 0 && (
              <button
                type="button"
                onClick={onClearAllInstructions}
                className="text-xs text-red-600 hover:text-red-700 transition-colors whitespace-nowrap cursor-pointer"
              >
                სიის გასუფთავება
              </button>
            )}
            <button
              type="button"
              onClick={onResetDefaultInstructions}
              className="inline-flex items-center gap-1 text-xs text-slate-500 hover:text-slate-900 transition-colors whitespace-nowrap cursor-pointer"
            >
              <RotateCcw className="w-3 h-3" />
              <span>სტანდარტულის აღდგენა</span>
            </button>
          </div>
        </div>

        <div className="divide-y divide-slate-200 border border-slate-200 rounded-lg bg-white max-h-[460px] overflow-y-auto">
          {instructions.length === 0 ? (
            <div className="p-6 text-center text-xs text-slate-500">
              შესაცვლელი უჯრები არ არის მონიშნული. დააჭირეთ სასურველ უჯრას მარჯვენა
              ცხრილში ან დაამატეთ კოორდინატი ზემოთ.
            </div>
          ) : (
            instructions.map((inst) => {
              const addr = inst.address.toUpperCase();
              const isSelected = effectiveSelectedAddress === inst.address;
              const currentAffixPos: AffixPosition = inst.affixPosition || 'start';
              const activeCurrency: CurrencySymbol = inst.currencySymbol || '₾';
              const activeFormulaFormat: FormulaResultFormat =
                inst.formulaResultFormat || 'number';
              const currentFormulaStr = String(
                editableValues[addr] ?? inst.formulaExpression ?? ''
              );
              const normalizedFormulaForEval = currentFormulaStr.trim()
                ? currentFormulaStr.trim().startsWith('=')
                  ? currentFormulaStr.trim()
                  : `=${currentFormulaStr.trim()}`
                : '';
              const liveFormulaResult =
                inst.inputType === 'formula' && normalizedFormulaForEval
                  ? evaluateExcelFormula(
                      normalizedFormulaForEval,
                      (refAddr) => liveCellValueLookup[refAddr.toUpperCase()] ?? ''
                    )
                  : null;

              const formattedFormulaPreview =
                liveFormulaResult !== null && liveFormulaResult !== ''
                  ? activeFormulaFormat === 'currency'
                    ? formatCurrencyCellValue(liveFormulaResult, activeCurrency)
                    : typeof liveFormulaResult === 'number'
                    ? liveFormulaResult.toLocaleString('en-US', {
                        minimumFractionDigits: Number.isInteger(liveFormulaResult) ? 0 : 2,
                        maximumFractionDigits: 2,
                      })
                    : String(liveFormulaResult ?? '')
                  : '';

              const finalFormulaPreview =
                inst.usePrefix && inst.prefixText && formattedFormulaPreview
                  ? applyAffixToDisplayValue(
                      formattedFormulaPreview,
                      true,
                      inst.prefixText,
                      currentAffixPos
                    )
                  : formattedFormulaPreview;

              return (
                <div
                  key={inst.id}
                  onClick={() => onSelectAddress(inst.address)}
                  className={`p-2.5 space-y-2 text-xs cursor-pointer transition-colors ${
                    isSelected ? 'bg-sky-50/70' : 'hover:bg-slate-50'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2 min-w-0 flex-1">
                      <span className="font-mono font-semibold text-sky-800 bg-sky-100/80 px-2 py-0.5 rounded text-[11px] shrink-0">
                        {inst.address}
                      </span>
                      <input
                        type="text"
                        aria-label={`${inst.address} უჯრის დასახელება`}
                        value={inst.label}
                        onChange={(e) =>
                          onUpdateInstruction(inst.id, { label: e.target.value })
                        }
                        onClick={(e) => e.stopPropagation()}
                        className="w-full bg-transparent text-slate-800 font-medium focus:bg-white focus:px-1.5 focus:py-0.5 focus:outline-none focus:ring-1 focus:ring-sky-600 rounded"
                      />
                    </div>

                    <div className="flex items-center gap-1.5 shrink-0">
                      <select
                        aria-label={`${inst.address} უჯრის ტიპი`}
                        value={inst.inputType}
                        onChange={(e) => {
                          const nextType = e.target.value as CellInputType;
                          onUpdateInstruction(inst.id, {
                            inputType: nextType,
                            currencySymbol:
                              nextType === 'currency' || nextType === 'formula'
                                ? inst.currencySymbol || '₾'
                                : inst.currencySymbol,
                            formulaResultFormat:
                              nextType === 'formula'
                                ? inst.formulaResultFormat || 'currency'
                                : inst.formulaResultFormat,
                          });
                        }}
                        onClick={(e) => e.stopPropagation()}
                        className="text-[11px] text-slate-700 bg-slate-100 border border-slate-200 rounded px-1.5 py-1"
                      >
                        <option value="text">ტექსტი</option>
                        <option value="number">რიცხვი</option>
                        <option value="currency">თანხა</option>
                        <option value="date">თარიღი</option>
                        <option value="formula">ფორმულა (=)</option>
                      </select>

                      {inst.inputType === 'currency' && (
                        <select
                          aria-label={`${inst.address} ვალუტის სიმბოლო`}
                          value={activeCurrency}
                          onChange={(e) =>
                            onUpdateInstruction(inst.id, {
                              currencySymbol: e.target.value as CurrencySymbol,
                            })
                          }
                          onClick={(e) => e.stopPropagation()}
                          className="text-[11px] font-semibold text-sky-900 bg-sky-100 border border-sky-300 rounded px-1.5 py-1"
                          title="აირჩიეთ თანხის ვალუტა: ლარი (₾) ან დოლარი ($)"
                        >
                          <option value="₾">₾ (ლარი)</option>
                          <option value="$">$ (დოლარი)</option>
                        </select>
                      )}

                      <button
                        type="button"
                        aria-label={`უჯრა ${inst.address}-ის წაშლა შესაცვლელი სიიდან`}
                        onClick={(e) => {
                          e.stopPropagation();
                          onRemoveInstruction(inst.id);
                        }}
                        className="p-1 text-slate-400 hover:text-red-600 rounded transition-colors cursor-pointer"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>

                  {/* Formula Row when Formula Type is selected: includes Formula Input AND Result Format (რიცხვი vs თანხა ₾ / $) */}
                  {inst.inputType === 'formula' && (
                    <div
                      onClick={(e) => e.stopPropagation()}
                      className="pt-1.5 border-t border-slate-100 space-y-2"
                    >
                      <div className="flex items-center gap-2">
                        <span className="text-[11px] font-medium text-emerald-800 shrink-0">
                          ფორმულა:
                        </span>
                        <input
                          type="text"
                          aria-label={`${inst.address} უჯრის ფორმულა`}
                          value={currentFormulaStr}
                          onChange={(e) => {
                            const val = e.target.value;
                            onChangeCellValue(inst.address, val, 'formula');
                            onUpdateInstruction(inst.id, {
                              formulaExpression: val,
                            });
                          }}
                          placeholder="მაგ: =D12*E12 ან =SUM(F12:F15)*B6"
                          className="flex-1 min-w-0 px-2 py-1 text-xs font-mono bg-emerald-50/50 border border-emerald-400 rounded focus:outline-none focus:ring-1 focus:ring-emerald-600 text-slate-900"
                        />
                      </div>

                      <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-50 px-2 py-1 rounded border border-slate-200/80">
                        <div className="flex items-center gap-1.5">
                          <span className="text-[11px] text-slate-600">შედეგი:</span>
                          <select
                            aria-label={`${inst.address} ფორმულის შედეგის ფორმატი`}
                            value={activeFormulaFormat}
                            onChange={(e) => {
                              const nextFmt = e.target.value as FormulaResultFormat;
                              onUpdateInstruction(inst.id, {
                                formulaResultFormat: nextFmt,
                                currencySymbol:
                                  nextFmt === 'currency'
                                    ? inst.currencySymbol || '₾'
                                    : inst.currencySymbol,
                              });
                            }}
                            className="text-[11px] text-slate-800 bg-white border border-slate-300 rounded px-1.5 py-0.5"
                          >
                            <option value="number">რიცხვი</option>
                            <option value="currency">თანხა</option>
                          </select>

                          {activeFormulaFormat === 'currency' && (
                            <div className="inline-flex items-center bg-slate-200/80 p-0.5 rounded">
                              <button
                                type="button"
                                onClick={() =>
                                  onUpdateInstruction(inst.id, {
                                    currencySymbol: '₾',
                                  })
                                }
                                className={`px-1.5 py-0.5 text-[11px] font-semibold rounded cursor-pointer ${
                                  activeCurrency === '₾'
                                    ? 'bg-sky-600 text-white'
                                    : 'text-slate-700 hover:text-slate-900'
                                }`}
                              >
                                ₾ ლარი
                              </button>
                              <button
                                type="button"
                                onClick={() =>
                                  onUpdateInstruction(inst.id, {
                                    currencySymbol: '$',
                                  })
                                }
                                className={`px-1.5 py-0.5 text-[11px] font-semibold rounded cursor-pointer ${
                                  activeCurrency === '$'
                                    ? 'bg-sky-600 text-white'
                                    : 'text-slate-700 hover:text-slate-900'
                                }`}
                              >
                                $ დოლარი
                              </button>
                            </div>
                          )}
                        </div>

                        {normalizedFormulaForEval && (
                          <span className="text-[11px] text-emerald-900">
                            ჩაიწერება:{' '}
                            <strong className="font-mono-tabular">
                              {finalFormulaPreview || '— (ცარიელი)'}
                            </strong>
                          </span>
                        )}
                      </div>
                    </div>
                  )}

                  {/* Mandatory Word ("სავალდებულო სიტყვა — თავში ან ბოლოში") & Optional ("არასავალდებულო") Toggles */}
                  <div
                    onClick={(e) => e.stopPropagation()}
                    className="flex flex-wrap items-center justify-between gap-2 pl-1 pt-1 border-t border-slate-100"
                  >
                    <div className="flex flex-wrap items-center gap-3 flex-1 min-w-0">
                      <label className="inline-flex items-center gap-1.5 text-[11px] text-slate-600 cursor-pointer select-none shrink-0">
                        <input
                          type="checkbox"
                          checked={Boolean(inst.usePrefix)}
                          onChange={(e) => {
                            const checked = e.target.checked;
                            const defaultPref =
                              inst.prefixText ||
                              (inst.affixPosition === 'end'
                                ? 'ლარი'
                                : inst.label.replace(/^პოზიცია #\d+ — /, '').trim());
                            onUpdateInstruction(inst.id, {
                              usePrefix: checked,
                              prefixText: checked ? defaultPref : inst.prefixText,
                            });
                          }}
                          className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                        />
                        <span>სავალდებულო სიტყვა</span>
                      </label>

                      {inst.inputType !== 'formula' && (
                        <label
                          className="inline-flex items-center gap-1.5 text-[11px] text-slate-600 cursor-pointer select-none shrink-0"
                          title="თუ ჩართულია, ამ უჯრის შევსება სავალდებულო არ არის; ცარიელის დატოვებისას ფორმულა მას არ გაითვალისწინებს"
                        >
                          <input
                            type="checkbox"
                            checked={Boolean(inst.optional)}
                            onChange={(e) =>
                              onUpdateInstruction(inst.id, {
                                optional: e.target.checked,
                              })
                            }
                            className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                          />
                          <span>არასავალდებულო (Optional)</span>
                        </label>
                      )}
                    </div>

                    {inst.usePrefix && (
                      <div className="flex items-center gap-1.5 w-full sm:w-auto flex-1 min-w-[190px]">
                        <select
                          aria-label={`${inst.address} სავალდებულო სიტყვის პოზიცია`}
                          value={currentAffixPos}
                          onChange={(e) =>
                            onUpdateInstruction(inst.id, {
                              affixPosition: e.target.value as AffixPosition,
                            })
                          }
                          className="text-[11px] text-slate-700 bg-slate-100 border border-slate-300 rounded px-1.5 py-1 shrink-0 focus:outline-none focus:ring-1 focus:ring-sky-600"
                        >
                          <option value="start">თავში</option>
                          <option value="end">ბოლოში</option>
                        </select>
                        <input
                          type="text"
                          aria-label={`${inst.address} სავალდებულო სიტყვა`}
                          value={inst.prefixText ?? ''}
                          onChange={(e) =>
                            onUpdateInstruction(inst.id, { prefixText: e.target.value })
                          }
                          placeholder={
                            currentAffixPos === 'end' ? 'მაგ: ლარი' : 'მაგ: მისამართი'
                          }
                          className="flex-1 min-w-[90px] px-2 py-1 text-[11px] bg-white border border-sky-300 rounded focus:outline-none focus:ring-1 focus:ring-sky-600 text-slate-900"
                        />
                      </div>
                    )}
                  </div>
                </div>
              );
            })
          )}
        </div>
      </div>

      {/* Explanation of Unchanged, Optional & Formula Cells */}
      <div className="p-3 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-600 space-y-1.5">
        <div className="flex items-center gap-1.5 font-medium text-slate-800">
          <Lock className="w-3.5 h-3.5 text-slate-500" />
          <span>სავალდებულო სიტყვა (თავში / ბოლოში), ფორმულები და Optional უჯრები</span>
        </div>
        <p className="leading-relaxed">
          „სავალდებულო სიტყვა“ შეგიძლიათ დააყენოთ როგორც <strong>თავში</strong> (მაგ: „მისამართი გორის ქუჩა 2“),
          ისე <strong>ბოლოში</strong> (მაგ: „10 ლარი“). „არასავალდებულო (Optional)“ უჯრები ცარიელის დატოვებისას
          ფორმულაში იგნორირდება, ხოლო შევსებისას — ითვლება.
        </p>
      </div>
    </div>
  );
};

