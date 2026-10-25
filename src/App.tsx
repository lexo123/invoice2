import React, { useEffect, useRef, useState } from 'react';
import {
  Calculator,
  Calendar,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  FileText,
  Home,
  Plus,
  RefreshCw,
  Settings,
  Trash2,
  Upload,
} from 'lucide-react';
import {
  AffixPosition,
  CellInputType,
  CurrencySymbol,
  EditableCellInstruction,
  FormulaResultFormat,
  GeneratedInvoicePackage,
  InvoiceProfileId,
  MainHeaderFieldConfig,
  MainHeaderFieldKey,
  ParsedSheetModel,
  ProductRowItem,
} from './types/invoice';
import {
  createDefaultInvoiceTemplateBuffer,
  DEFAULT_EDITABLE_INSTRUCTIONS,
} from './utils/defaultInvoiceTemplate';
import {
  generateUpdatedExcelPackage,
  parseExcelBufferToSheetModel,
} from './utils/excelParserAndGenerator';
import {
  downloadExcelBuffer,
  exportSheetElementToPdf,
} from './utils/exportFileUtils';
import {
  applyAffixToDisplayValue,
  buildCombinedWorkbookPayload,
  calculateProductsSummary,
  composeCellValueWithPrefix,
  DEFAULT_MAIN_HEADER_CONFIGS,
  detectCellPrefixCandidate,
  evaluateExcelFormula,
  formatCurrencyCellValue,
  georgianDotToIsoDate,
  getTodayGeorgianDate,
  isoDateToGeorgianDot,
  parseCellAddress,
  resolveInvoiceFileNameBase,
  resolveMainHeaderFieldConfigs,
  stripPrefixFromCellValue,
  syncFileNameWithCompanyName,
} from './utils/formulaEngine';
import {
  clearCustomDefaultTemplate,
  loadCustomDefaultTemplate,
  loadSavedGrandTotalAddress,
  loadSavedInstructionsForTemplate,
  loadSavedMainHeaderConfigs,
  saveCustomDefaultTemplate,
  saveGrandTotalAddress,
  saveInstructionsForTemplate,
  saveMainHeaderConfigs,
} from './utils/templateStorage';
import { ExcelSheetPreview } from './components/ExcelSheetPreview';
import { InstructionsConfigPanel } from './components/InstructionsConfigPanel';

type AppMainView = 'home' | 'settings';
type SettingsSubTab = 'fields' | 'instructions' | 'template';

const DEFAULT_TEMPLATE_FILENAME = 'სტანდარტული_ინვოისის_შაბლონი.xlsx';
const MAX_PRODUCT_ROWS = 8; // 1 initial (A17:C17) + up to 7 added (A18:C18 .. A24:C24)

const PROFILE_LABELS: Record<InvoiceProfileId, string> = {
  mypos: 'მაიპოს',
  btrade: 'ბითრეიდ',
};

interface ProfileSessionSnapshot {
  templateBuffer: ArrayBuffer | null;
  templateName: string;
  isCustomDefault: boolean;
  mainHeaderConfigs: MainHeaderFieldConfig[];
  mainHeaderValues: Record<MainHeaderFieldKey, string>;
  customFileName: string;
  isFileNameCustomized: boolean;
  productRows: ProductRowItem[];
  grandTotalAddress: string;
  instructions: EditableCellInstruction[];
  editableValues: Record<string, string | number>;
  sheetModel: ParsedSheetModel | null;
  generatedPackage: GeneratedInvoicePackage | null;
  isDirtySinceLastGenerate: boolean;
}

export default function App() {
  // Active company/template profile: 'mypos' ("მაიპოს") vs 'btrade' ("ბითრეიდ")
  const [activeProfile, setActiveProfile] = useState<InvoiceProfileId>('mypos');
  // Top-level navigation: 'home' (streamlined invoice generator for activeProfile) vs 'settings' (configuration tools)
  const [activeView, setActiveView] = useState<AppMainView>('home');
  const [settingsTab, setSettingsTab] = useState<SettingsSubTab>('fields');

  const [templateBuffer, setTemplateBuffer] = useState<ArrayBuffer | null>(null);
  const [templateName, setTemplateName] = useState<string>(DEFAULT_TEMPLATE_FILENAME);
  const [isCustomDefault, setIsCustomDefault] = useState<boolean>(false);

  // Main Page 6 Header Fields configuration & values
  const [mainHeaderConfigs, setMainHeaderConfigs] = useState<MainHeaderFieldConfig[]>(
    DEFAULT_MAIN_HEADER_CONFIGS
  );
  const [mainHeaderValues, setMainHeaderValues] = useState<
    Record<MainHeaderFieldKey, string>
  >(() => ({
    companyName: '',
    cadastralCode: '',
    address: '',
    usdRate: '2.6',
    date: getTodayGeorgianDate(),
    invoiceNumber: '',
  }));

  // File Name ("ფაილის სახელი") — defaults to Company Name ("კომპანიის სახელი") unless manually changed
  const [customFileName, setCustomFileName] = useState<string>('');
  const [isFileNameCustomized, setIsFileNameCustomized] = useState<boolean>(false);

  // Main Page Dynamic Product Rows (starts with 1 initial product at A17, B17, C17; up to 7 more can be added -> max 8)
  const [productRows, setProductRows] = useState<ProductRowItem[]>([
    {
      id: 'prod-1',
      productName: '',
      quantity: '1',
      unitPrice: '',
      currency: '₾', // Default currency is Lari (₾)
    },
  ]);

  // Grand total cell address (default D36)
  const [grandTotalAddress, setGrandTotalAddress] = useState<string>('D36');

  // Full instructions & editableValues (accessible & configurable in Settings)
  const [instructions, setInstructions] = useState<EditableCellInstruction[]>([]);
  const [editableValues, setEditableValues] = useState<Record<string, string | number>>({});
  const [sheetModel, setSheetModel] = useState<ParsedSheetModel | null>(null);
  const [selectedAddress, setSelectedAddress] = useState<string | null>('A17');

  // Generated Package state
  const [generatedPackage, setGeneratedPackage] =
    useState<GeneratedInvoicePackage | null>(null);
  const [isDirtySinceLastGenerate, setIsDirtySinceLastGenerate] =
    useState<boolean>(false);
  const [isGeneratingExcel, setIsGeneratingExcel] = useState<boolean>(false);
  const [isExportingPdf, setIsExportingPdf] = useState<boolean>(false);
  const [statusBanner, setStatusBanner] = useState<string | null>(null);
  const [errorBanner, setErrorBanner] = useState<string | null>(null);

  // Preview display toggles (used in Settings preview)
  const [showExcelGridHeaders, setShowExcelGridHeaders] = useState<boolean>(true);
  const [highlightEditableCells, setHighlightEditableCells] = useState<boolean>(true);

  const sheetPrintRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const profileSessionCacheRef = useRef<
    Partial<Record<InvoiceProfileId, ProfileSessionSnapshot>>
  >({});

  // Live calculation of all product rows (A17:C17 .. A24:C24) and Grand Total in Lari (₾) BEFORE clicking "შექმნა"
  const liveProductsSummary = React.useMemo(() => {
    return calculateProductsSummary(productRows, mainHeaderValues.usdRate);
  }, [productRows, mainHeaderValues.usdRate]);

  /**
   * Initializes Main Page header values and initial product row with clean defaults:
   * - usdRate defaults to '2.6'
   * - date defaults to today's date (DD.MM.YYYY)
   * - all other fillable fields start empty
   */
  const extractInitialMainFormState = (
    _parsed?: ParsedSheetModel,
    _resolvedConfigs?: MainHeaderFieldConfig[],
    _existingInstructions?: EditableCellInstruction[]
  ) => {
    const nextHeaderVals: Record<MainHeaderFieldKey, string> = {
      companyName: '',
      cadastralCode: '',
      address: '',
      usdRate: '2.6',
      date: getTodayGeorgianDate(),
      invoiceNumber: '',
    };

    const initialRows: ProductRowItem[] = [
      {
        id: 'prod-1',
        productName: '',
        quantity: '1',
        unitPrice: '',
        currency: '₾',
      },
    ];

    return {
      nextHeaderVals,
      initialRows,
    };
  };

  const loadTemplateForProfile = async (
    targetProfile: InvoiceProfileId,
    isCancelled?: () => boolean
  ) => {
    try {
      const cached = profileSessionCacheRef.current[targetProfile];
      if (cached && cached.templateBuffer && cached.sheetModel) {
        setTemplateBuffer(cached.templateBuffer);
        setTemplateName(cached.templateName);
        setIsCustomDefault(cached.isCustomDefault);
        setMainHeaderConfigs(cached.mainHeaderConfigs);
        setMainHeaderValues(cached.mainHeaderValues);
        setCustomFileName(cached.customFileName);
        setIsFileNameCustomized(cached.isFileNameCustomized);
        setProductRows(cached.productRows);
        setGrandTotalAddress(cached.grandTotalAddress);
        setInstructions(cached.instructions);
        setEditableValues(cached.editableValues);
        setSheetModel(cached.sheetModel);
        setGeneratedPackage(cached.generatedPackage);
        setIsDirtySinceLastGenerate(cached.isDirtySinceLastGenerate);
        return;
      }

      const savedCustom = loadCustomDefaultTemplate(targetProfile);
      let buf: ArrayBuffer;
      let name: string;
      let customFlag = false;

      if (savedCustom) {
        buf = savedCustom.buffer;
        name = savedCustom.fileName;
        customFlag = true;
      } else {
        buf = await createDefaultInvoiceTemplateBuffer();
        name =
          targetProfile === 'btrade'
            ? 'btrade_სტანდარტული_შაბლონი.xlsx'
            : DEFAULT_TEMPLATE_FILENAME;
      }

      const parsed = await parseExcelBufferToSheetModel(buf);
      if (isCancelled?.()) return;

      const loadedInstructions = loadSavedInstructionsForTemplate(
        name,
        customFlag ? [] : DEFAULT_EDITABLE_INSTRUCTIONS,
        targetProfile
      );

      const savedMainConfigs = loadSavedMainHeaderConfigs(name, targetProfile);
      const resolvedHeaderConfigs = resolveMainHeaderFieldConfigs(
        parsed,
        loadedInstructions,
        savedMainConfigs
      );

      const loadedTotalAddr = loadSavedGrandTotalAddress(
        name,
        targetProfile,
        'D36'
      );

      const { nextHeaderVals, initialRows } = extractInitialMainFormState(
        parsed,
        resolvedHeaderConfigs,
        loadedInstructions
      );

      const initVals: Record<string, string | number> = {};
      for (const inst of loadedInstructions) {
        const addr = inst.address.toUpperCase();
        const cell = parsed.cells[addr];
        if (inst.inputType === 'formula') {
          initVals[addr] =
            inst.formulaExpression ||
            (cell?.formula ? `=${cell.formula}` : '=');
        } else {
          const rawCellVal = cell?.value ?? '';
          initVals[addr] = stripPrefixFromCellValue(
            rawCellVal,
            inst.usePrefix,
            inst.prefixText,
            inst.affixPosition || 'start'
          );
        }
      }

      const { effectiveOverrides, affixRules, currencyRules } =
        buildCombinedWorkbookPayload(
          loadedInstructions,
          initVals,
          resolvedHeaderConfigs,
          nextHeaderVals,
          initialRows,
          loadedTotalAddr,
          parsed
        );

      const pkg = await generateUpdatedExcelPackage(
        buf,
        effectiveOverrides,
        'ინვოისი',
        affixRules,
        currencyRules
      );
      if (isCancelled?.()) return;

      setTemplateBuffer(buf);
      setTemplateName(name);
      setIsCustomDefault(customFlag);
      setMainHeaderConfigs(resolvedHeaderConfigs);
      setMainHeaderValues(nextHeaderVals);
      setCustomFileName(nextHeaderVals.companyName);
      setIsFileNameCustomized(false);
      setProductRows(initialRows);
      setGrandTotalAddress(loadedTotalAddr);
      setInstructions(loadedInstructions);
      setEditableValues(initVals);
      setGeneratedPackage(pkg);
      setSheetModel(pkg.sheetModel);
      setIsDirtySinceLastGenerate(false);

      profileSessionCacheRef.current[targetProfile] = {
        templateBuffer: buf,
        templateName: name,
        isCustomDefault: customFlag,
        mainHeaderConfigs: resolvedHeaderConfigs,
        mainHeaderValues: nextHeaderVals,
        customFileName: nextHeaderVals.companyName,
        isFileNameCustomized: false,
        productRows: initialRows,
        grandTotalAddress: loadedTotalAddr,
        instructions: loadedInstructions,
        editableValues: initVals,
        sheetModel: pkg.sheetModel,
        generatedPackage: pkg,
        isDirtySinceLastGenerate: false,
      };
    } catch (err: any) {
      if (!isCancelled?.()) {
        setErrorBanner(
          err?.message || 'შაბლონის ჩატვირთვისას დაფიქსირდა შეცდომა.'
        );
      }
    }
  };

  // Switch between 'mypos' ("მაიპოს") and 'btrade' ("ბითრეიდ") saving current session snapshot
  const handleSwitchProfile = async (
    nextProfile: InvoiceProfileId,
    nextView?: AppMainView
  ) => {
    if (nextView) {
      setActiveView(nextView);
    }
    setStatusBanner(null);
    setErrorBanner(null);
    if (nextProfile === activeProfile) return;

    profileSessionCacheRef.current[activeProfile] = {
      templateBuffer,
      templateName,
      isCustomDefault,
      mainHeaderConfigs,
      mainHeaderValues,
      customFileName,
      isFileNameCustomized,
      productRows,
      grandTotalAddress,
      instructions,
      editableValues,
      sheetModel,
      generatedPackage,
      isDirtySinceLastGenerate,
    };

    setActiveProfile(nextProfile);
    await loadTemplateForProfile(nextProfile);
  };

  // Initialize embedded/custom template on mount and ensure Btrade custom template in localStorage syncs
  useEffect(() => {
    let cancelled = false;
    loadCustomDefaultTemplate('btrade');
    loadTemplateForProfile('mypos', () => cancelled);
    return () => {
      cancelled = true;
    };
  }, []);

  // Save instructions scoped to current template & profile
  const updateInstructionsList = (
    next: EditableCellInstruction[],
    targetTemplateName = templateName,
    targetProfile = activeProfile
  ) => {
    setInstructions(next);
    saveInstructionsForTemplate(targetTemplateName, next, targetProfile);
  };

  // Update one of the 6 Main Header field values
  const handleMainHeaderValueChange = (key: MainHeaderFieldKey, value: string) => {
    if (key === 'companyName') {
      const synced = syncFileNameWithCompanyName(
        mainHeaderValues.companyName,
        value,
        customFileName,
        isFileNameCustomized
      );
      setCustomFileName(synced.fileName);
      setIsFileNameCustomized(synced.isFileNameCustomized);
    }
    setMainHeaderValues((prev) => ({
      ...prev,
      [key]: value,
    }));
    const cfg = mainHeaderConfigs.find((c) => c.key === key);
    if (cfg) {
      const addr = cfg.address.toUpperCase();
      setEditableValues((prev) => ({
        ...prev,
        [addr]: value,
      }));
    }
    setIsDirtySinceLastGenerate(true);
    setStatusBanner(null);
  };

  // Update configuration (e.g. cell coordinate or prefix toggle) of one of the 6 Main Header fields
  const handleUpdateMainHeaderConfig = (
    key: MainHeaderFieldKey,
    patch: Partial<MainHeaderFieldConfig>
  ) => {
    const next = mainHeaderConfigs.map((cfg) =>
      cfg.key === key ? { ...cfg, ...patch } : cfg
    );
    setMainHeaderConfigs(next);
    saveMainHeaderConfigs(templateName, next, activeProfile);
    setIsDirtySinceLastGenerate(true);
  };

  // Product Row Handlers (Initial product at A17:C17 + up to 7 added products A18:C18 .. A24:C24)
  const handleAddProductRow = () => {
    if (productRows.length >= MAX_PRODUCT_ROWS) return;
    const nextIndex = productRows.length + 1;
    const nextExcelRow = 16 + nextIndex;
    setProductRows((prev) => [
      ...prev,
      {
        id: `prod-${Date.now()}-${nextIndex}`,
        productName: '',
        quantity: '1',
        unitPrice: '',
        currency: '₾', // Default currency is Lari (₾)
      },
    ]);
    setSelectedAddress(`A${nextExcelRow}`);
    setIsDirtySinceLastGenerate(true);
    setStatusBanner(null);
  };

  const handleRemoveProductRow = (indexToRemove: number) => {
    if (productRows.length <= 1) return;
    setProductRows((prev) => prev.filter((_, idx) => idx !== indexToRemove));
    setIsDirtySinceLastGenerate(true);
    setStatusBanner(null);
  };

  const handleProductRowChange = (
    index: number,
    patch: Partial<ProductRowItem>
  ) => {
    setProductRows((prev) =>
      prev.map((row, idx) => (idx === index ? { ...row, ...patch } : row))
    );
    setIsDirtySinceLastGenerate(true);
    setStatusBanner(null);
  };

  const handleFieldChange = (address: string, rawVal: string, inputType: string) => {
    const norm = address.toUpperCase();
    const isNumericType = inputType === 'number' || inputType === 'currency';
    const nextVal: string | number =
      isNumericType && rawVal.trim() !== '' && !Number.isNaN(Number(rawVal))
        ? Number(rawVal)
        : rawVal;

    setEditableValues((prev) => ({
      ...prev,
      [norm]: nextVal,
    }));

    // Also sync if this cell is one of the 6 Main Header fields
    const matchedHdr = mainHeaderConfigs.find(
      (c) => c.address.toUpperCase() === norm
    );
    if (matchedHdr) {
      if (matchedHdr.key === 'companyName') {
        const synced = syncFileNameWithCompanyName(
          mainHeaderValues.companyName,
          rawVal,
          customFileName,
          isFileNameCustomized
        );
        setCustomFileName(synced.fileName);
        setIsFileNameCustomized(synced.isFileNameCustomized);
      }
      setMainHeaderValues((prev) => ({
        ...prev,
        [matchedHdr.key]: rawVal,
      }));
    }

    setIsDirtySinceLastGenerate(true);
    setStatusBanner(null);
  };

  const handleUpdateInstructionAndSync = (
    id: string,
    patch: Partial<EditableCellInstruction>
  ) => {
    const targetInst = instructions.find((item) => item.id === id);
    let enrichedPatch = { ...patch };

    if (
      patch.inputType === 'currency' &&
      !patch.currencySymbol &&
      !targetInst?.currencySymbol
    ) {
      enrichedPatch.currencySymbol = '₾';
    }

    if (targetInst && patch.inputType === 'formula') {
      const addr = targetInst.address.toUpperCase();
      const existingFormula = sheetModel?.cells[addr]?.formula;
      const curVal = String(
        editableValues[addr] ?? targetInst.formulaExpression ?? ''
      ).trim();
      const nextFormula = curVal.startsWith('=')
        ? curVal
        : existingFormula
        ? `=${existingFormula}`
        : curVal
        ? `=${curVal}`
        : '=';
      enrichedPatch.formulaExpression = nextFormula;
      if (!patch.formulaResultFormat && !targetInst.formulaResultFormat) {
        enrichedPatch.formulaResultFormat = 'currency';
      }
      if (!patch.currencySymbol && !targetInst.currencySymbol) {
        enrichedPatch.currencySymbol = '₾';
      }
      setEditableValues((prev) => ({
        ...prev,
        [addr]: nextFormula,
      }));
    }

    const next = instructions.map((item) =>
      item.id === id ? { ...item, ...enrichedPatch } : item
    );
    updateInstructionsList(next);
    setIsDirtySinceLastGenerate(true);
  };

  /**
   * Generates the updated Excel (.xlsx) package in memory and updates the 1:1 preview
   */
  const buildLatestPackage = async (): Promise<GeneratedInvoicePackage | null> => {
    if (!templateBuffer) return null;
    const { effectiveOverrides, affixRules, currencyRules } =
      buildCombinedWorkbookPayload(
        instructions,
        editableValues,
        mainHeaderConfigs,
        mainHeaderValues,
        productRows,
        grandTotalAddress,
        sheetModel
      );
    const resolvedBaseName = resolveInvoiceFileNameBase(
      customFileName,
      mainHeaderValues.companyName
    );
    const pkg = await generateUpdatedExcelPackage(
      templateBuffer,
      effectiveOverrides,
      resolvedBaseName,
      affixRules,
      currencyRules
    );
    setGeneratedPackage(pkg);
    setSheetModel(pkg.sheetModel);
    setIsDirtySinceLastGenerate(false);
    return pkg;
  };

  const handleGenerateNewExcel = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!templateBuffer) return;

    setIsGeneratingExcel(true);
    setErrorBanner(null);
    try {
      const pkg = await buildLatestPackage();
      if (pkg) {
        const resolvedBaseName = resolveInvoiceFileNameBase(
          customFileName,
          mainHeaderValues.companyName
        );
        setStatusBanner(
          `ინვოისი მომზადდა (${resolvedBaseName}.pdf / ${resolvedBaseName}.xlsx)! სულ ჯამი ლარში: ${liveProductsSummary.formattedGrandTotalGel}.`
        );
      }
    } catch (err: any) {
      setErrorBanner(
        err?.message || 'ახალი Excel ფაილის გენერაციისას დაფიქსირდა შეცდომა.'
      );
    } finally {
      setIsGeneratingExcel(false);
    }
  };

  const handleDownloadGeneratedExcel = async () => {
    if (!templateBuffer) return;
    setErrorBanner(null);
    try {
      const pkg =
        !generatedPackage || isDirtySinceLastGenerate
          ? await buildLatestPackage()
          : generatedPackage;
      if (!pkg) return;
      const resolvedBaseName = resolveInvoiceFileNameBase(
        customFileName,
        mainHeaderValues.companyName
      );
      downloadExcelBuffer(pkg.xlsxBuffer, `${resolvedBaseName}.xlsx`);
      setStatusBanner(`ფაილი ${resolvedBaseName}.xlsx გადმოიწერა.`);
    } catch (err: any) {
      setErrorBanner(err?.message || 'Excel ფაილის გადმოწერისას დაფიქსირდა შეცდომა.');
    }
  };

  const handleDownloadGeneratedPdf = async () => {
    if (!templateBuffer || !sheetPrintRef.current) return;
    setIsExportingPdf(true);
    setErrorBanner(null);
    try {
      const pkg =
        !generatedPackage || isDirtySinceLastGenerate
          ? await buildLatestPackage()
          : generatedPackage;
      if (!pkg) return;
      const resolvedBaseName = resolveInvoiceFileNameBase(
        customFileName,
        mainHeaderValues.companyName
      );
      await new Promise((r) =>
        requestAnimationFrame(() => requestAnimationFrame(r))
      );
      await exportSheetElementToPdf(
        sheetPrintRef.current,
        `${resolvedBaseName}.pdf`
      );
      setStatusBanner(
        `PDF ფაილი (${resolvedBaseName}.pdf) ორიგინალი ფერებითა და სურათებით გადმოიწერა.`
      );
    } catch (err: any) {
      setErrorBanner(err?.message || 'PDF-ის გენერაციისას დაფიქსირდა შეცდომა.');
    } finally {
      setIsExportingPdf(false);
    }
  };

  /**
   * Handle custom `.xlsx` template upload in Settings and embed it as default for the active profile
   */
  const handleUploadTemplateFile = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setErrorBanner(null);
    setStatusBanner(null);

    try {
      const arrayBuf = await file.arrayBuffer();
      const parsed = await parseExcelBufferToSheetModel(arrayBuf);
      setTemplateBuffer(arrayBuf);
      setTemplateName(file.name);
      setIsCustomDefault(true);
      saveCustomDefaultTemplate(arrayBuf, file.name, activeProfile);

      const savedForThisFile = loadSavedInstructionsForTemplate(
        file.name,
        instructions.length > 0 ? instructions : DEFAULT_EDITABLE_INSTRUCTIONS,
        activeProfile
      );
      setInstructions(savedForThisFile);
      saveInstructionsForTemplate(file.name, savedForThisFile, activeProfile);

      const savedMainCfg =
        loadSavedMainHeaderConfigs(file.name, activeProfile) || mainHeaderConfigs;
      const resolvedHdr = resolveMainHeaderFieldConfigs(
        parsed,
        savedForThisFile,
        savedMainCfg
      );
      setMainHeaderConfigs(resolvedHdr);
      saveMainHeaderConfigs(file.name, resolvedHdr, activeProfile);

      const loadedTotalAddr = loadSavedGrandTotalAddress(
        file.name,
        activeProfile,
        grandTotalAddress
      );
      setGrandTotalAddress(loadedTotalAddr);
      saveGrandTotalAddress(file.name, loadedTotalAddr, activeProfile);

      const { nextHeaderVals, initialRows } = extractInitialMainFormState(
        parsed,
        resolvedHdr,
        savedForThisFile
      );
      setMainHeaderValues(nextHeaderVals);
      setCustomFileName(nextHeaderVals.companyName);
      setIsFileNameCustomized(false);
      setProductRows(initialRows);

      const nextVals: Record<string, string | number> = {};
      for (const inst of savedForThisFile) {
        const addr = inst.address.toUpperCase();
        const c = parsed.cells[addr];
        if (inst.inputType === 'formula' && c?.formula) {
          nextVals[addr] = `=${c.formula}`;
        } else {
          nextVals[addr] = stripPrefixFromCellValue(
            c?.value ?? '',
            inst.usePrefix,
            inst.prefixText,
            inst.affixPosition || 'start'
          );
        }
      }
      setEditableValues(nextVals);

      const { effectiveOverrides, affixRules, currencyRules } =
        buildCombinedWorkbookPayload(
          savedForThisFile,
          nextVals,
          resolvedHdr,
          nextHeaderVals,
          initialRows,
          loadedTotalAddr,
          parsed
        );

      const pkg = await generateUpdatedExcelPackage(
        arrayBuf,
        effectiveOverrides,
        file.name.replace(/\.xlsx$/i, ''),
        affixRules,
        currencyRules
      );
      setGeneratedPackage(pkg);
      setSheetModel(pkg.sheetModel);
      setIsDirtySinceLastGenerate(false);
      profileSessionCacheRef.current[activeProfile] = {
        templateBuffer: arrayBuf,
        templateName: file.name,
        isCustomDefault: true,
        mainHeaderConfigs: resolvedHdr,
        mainHeaderValues: nextHeaderVals,
        customFileName: nextHeaderVals.companyName,
        isFileNameCustomized: false,
        productRows: initialRows,
        grandTotalAddress: loadedTotalAddr,
        instructions: savedForThisFile,
        editableValues: nextVals,
        sheetModel: pkg.sheetModel,
        generatedPackage: pkg,
        isDirtySinceLastGenerate: false,
      };
      setStatusBanner(
        `„${PROFILE_LABELS[activeProfile]}“-ის შაბლონი "${file.name}" აიტვირთა და ჩაშენდა საიტზე.`
      );
    } catch (err: any) {
      setErrorBanner(
        err?.message ||
          'ფაილის წაკითხვა ვერ მოხერხდა. დარწმუნდით, რომ ტვირთავთ სწორ .xlsx ფაილს.'
      );
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleRestoreFactoryTemplate = async () => {
    clearCustomDefaultTemplate(templateName, activeProfile);
    delete profileSessionCacheRef.current[activeProfile];
    await loadTemplateForProfile(activeProfile);
    setStatusBanner(
      `აღდგენილია „${PROFILE_LABELS[activeProfile]}“-ის ჩაშენებული სტანდარტული შაბლონი.`
    );
  };

  // Resolve master address when clicking merged cells in preview
  const effectiveSelectedAddress = React.useMemo(() => {
    if (!selectedAddress || !sheetModel) return selectedAddress;
    const cell = sheetModel.cells[selectedAddress];
    if (cell?.isMergedSlave && cell.masterAddress) {
      return cell.masterAddress;
    }
    return selectedAddress;
  }, [selectedAddress, sheetModel]);

  const selectedCellInfo =
    effectiveSelectedAddress && sheetModel
      ? sheetModel.cells[effectiveSelectedAddress]
      : null;

  const selectedCellInstruction = instructions.find(
    (i) => i.address.toUpperCase() === effectiveSelectedAddress?.toUpperCase()
  );

  // Build combined list of instructions for highlighting on the sheet preview
  const previewHighlightInstructions = React.useMemo(() => {
    const list: EditableCellInstruction[] = [...instructions];
    const seen = new Set(list.map((i) => i.address.toUpperCase()));

    for (const cfg of mainHeaderConfigs) {
      const addr = cfg.address.toUpperCase();
      if (!seen.has(addr)) {
        seen.add(addr);
        list.push({
          id: `hdr-${cfg.key}`,
          address: addr,
          label: cfg.label,
          group: 'header',
          inputType: cfg.inputType,
          usePrefix: cfg.usePrefix,
          prefixText: cfg.prefixText,
          affixPosition: cfg.affixPosition,
        });
      }
    }

    for (let i = 0; i < productRows.length; i++) {
      const r = 17 + i;
      for (const [col, lbl] of [
        ['A', `პროდუქტი #${i + 1}`],
        ['B', `რაოდენობა #${i + 1}`],
        ['C', `ფასი #${i + 1}`],
      ]) {
        const addr = `${col}${r}`;
        if (!seen.has(addr)) {
          seen.add(addr);
          list.push({
            id: `prod-${addr}`,
            address: addr,
            label: lbl,
            group: 'items',
            inputType: col === 'A' ? 'text' : col === 'B' ? 'number' : 'currency',
          });
        }
      }
    }

    const totalAddr = (grandTotalAddress || 'D36').toUpperCase();
    if (!seen.has(totalAddr)) {
      list.push({
        id: 'total-d36',
        address: totalAddr,
        label: 'ჯამი (₾)',
        group: 'other',
        inputType: 'formula',
      });
    }

    return list;
  }, [instructions, mainHeaderConfigs, productRows.length, grandTotalAddress]);

  const handleQuickAddSelectedCell = (
    withPrefix: boolean,
    explicitType?: CellInputType,
    affixPos: AffixPosition = 'start'
  ) => {
    if (!effectiveSelectedAddress || !sheetModel) return;
    if (selectedCellInstruction) {
      const next = instructions.filter((i) => i.id !== selectedCellInstruction.id);
      updateInstructionsList(next);
      setIsDirtySinceLastGenerate(true);
      return;
    }
    const cellObj = sheetModel.cells[effectiveSelectedAddress];
    const rawStr = String(cellObj?.formattedValue || cellObj?.value || '');
    const candidate = detectCellPrefixCandidate(rawStr);
    const existingFormula = cellObj?.formula;
    const hasGelFmt =
      cellObj?.numFmt?.includes('₾') ||
      String(cellObj?.formattedValue || '').includes('₾');
    const guessedType: CellInputType =
      explicitType ||
      (existingFormula
        ? 'formula'
        : !withPrefix && typeof cellObj?.value === 'number'
        ? hasGelFmt
          ? 'currency'
          : 'number'
        : 'text');

    const isFormula = guessedType === 'formula';
    const pref = withPrefix
      ? affixPos === 'end'
        ? 'ლარი'
        : candidate.suggestedPrefix || 'მისამართი'
      : '';
    const initVal = isFormula
      ? existingFormula
        ? `=${existingFormula}`
        : '='
      : withPrefix && affixPos === 'start'
      ? candidate.remainder
      : cellObj?.value ?? '';
    const cleanLabel =
      pref.replace(/:$/, '').trim() && affixPos === 'start'
        ? pref.replace(/:$/, '').trim()
        : `უჯრა ${effectiveSelectedAddress}`;

    const newInst: EditableCellInstruction = {
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
    };

    const next = [...instructions, newInst];
    updateInstructionsList(next);
    setEditableValues((prev) => ({
      ...prev,
      [effectiveSelectedAddress.toUpperCase()]: initVal,
    }));
    if (isFormula) {
      setActiveView('settings');
      setSettingsTab('instructions');
    }
    setIsDirtySinceLastGenerate(true);
  };

  // Live map of cell values for instant formula evaluation preview in Settings
  const liveCellValueLookup = React.useMemo(() => {
    const map: Record<string, string | number> = {};
    if (sheetModel) {
      for (const [k, cell] of Object.entries(sheetModel.cells)) {
        map[k.toUpperCase()] = cell.value;
      }
    }
    const { effectiveOverrides } = buildCombinedWorkbookPayload(
      instructions,
      editableValues,
      mainHeaderConfigs,
      mainHeaderValues,
      productRows,
      grandTotalAddress,
      sheetModel
    );
    for (const [k, v] of Object.entries(effectiveOverrides)) {
      if (typeof v === 'string' && v.trim().startsWith('=')) continue;
      map[k.toUpperCase()] = v;
    }
    return map;
  }, [
    sheetModel,
    instructions,
    editableValues,
    mainHeaderConfigs,
    mainHeaderValues,
    productRows,
    grandTotalAddress,
  ]);

  // Renders an editable cell card inside Settings -> "1. უჯრების მართვა"
  const renderEditableFieldControl = (inst: EditableCellInstruction) => {
    const addr = inst.address.toUpperCase();
    const currentAffixPos: AffixPosition = inst.affixPosition || 'start';
    const activeCurrency: CurrencySymbol = inst.currencySymbol || '₾';
    const activeFormulaFormat: FormulaResultFormat =
      inst.formulaResultFormat || 'number';
    const currentInputVal =
      inst.inputType === 'formula'
        ? editableValues[addr] ?? inst.formulaExpression ?? ''
        : editableValues[addr] ?? '';
    const isFormulaType =
      inst.inputType === 'formula' ||
      (!inst.usePrefix && String(currentInputVal).trim().startsWith('='));
    const composedPreview = composeCellValueWithPrefix(
      currentInputVal,
      inst.usePrefix,
      inst.prefixText,
      inst.inputType,
      inst.optional,
      currentAffixPos
    );

    const liveFormulaEval = isFormulaType
      ? evaluateExcelFormula(
          String(composedPreview),
          (refAddr) => liveCellValueLookup[refAddr.toUpperCase()] ?? ''
        )
      : null;

    const formattedFormulaResult =
      liveFormulaEval !== null && liveFormulaEval !== ''
        ? activeFormulaFormat === 'currency'
          ? formatCurrencyCellValue(liveFormulaEval, activeCurrency)
          : typeof liveFormulaEval === 'number'
          ? liveFormulaEval.toLocaleString('en-US', {
              minimumFractionDigits: Number.isInteger(liveFormulaEval) ? 0 : 2,
              maximumFractionDigits: 2,
            })
          : String(liveFormulaEval ?? '')
        : '';

    const finalFormulaDisplay =
      inst.usePrefix && inst.prefixText && formattedFormulaResult
        ? applyAffixToDisplayValue(
            formattedFormulaResult,
            true,
            inst.prefixText,
            currentAffixPos
          )
        : formattedFormulaResult;

    const currencyFormattedPreview =
      inst.inputType === 'currency' && String(composedPreview) !== ''
        ? formatCurrencyCellValue(
            composedPreview,
            activeCurrency,
            sheetModel?.cells[addr]?.numFmt
          )
        : '';

    const isoDateValue =
      inst.inputType === 'date' ? georgianDotToIsoDate(String(currentInputVal)) : '';

    const affixInputControl = inst.usePrefix ? (
      <div className="w-2/5 shrink-0 flex items-center gap-1">
        <select
          aria-label={`${addr} სავალდებულო სიტყვის პოზიცია`}
          value={currentAffixPos}
          onChange={(e) =>
            handleUpdateInstructionAndSync(inst.id, {
              affixPosition: e.target.value as AffixPosition,
            })
          }
          className="px-1.5 py-1.5 text-[11px] bg-sky-100/80 border border-sky-300 text-sky-950 font-medium rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600 shrink-0 cursor-pointer"
        >
          <option value="start">თავში</option>
          <option value="end">ბოლოში</option>
        </select>
        <input
          type="text"
          aria-label={`${addr} სავალდებულო სიტყვა`}
          value={inst.prefixText ?? ''}
          onChange={(e) =>
            handleUpdateInstructionAndSync(inst.id, {
              prefixText: e.target.value,
            })
          }
          placeholder={currentAffixPos === 'end' ? 'მაგ: ლარი' : 'მაგ: მისამართი'}
          className="w-full min-w-0 px-2 py-1.5 text-xs bg-sky-50 border border-sky-300 text-sky-950 font-medium rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
        />
      </div>
    ) : null;

    return (
      <div
        key={inst.id}
        className="p-2.5 rounded-lg border border-slate-200 bg-slate-50/50 space-y-2"
      >
        <div className="flex flex-wrap items-center justify-between gap-2">
          <label
            htmlFor={`cell-input-${addr}`}
            className="text-xs font-medium text-slate-700 leading-snug flex items-center gap-1.5 flex-wrap"
          >
            <span className="font-mono font-semibold text-sky-700">
              [{addr}]
            </span>
            <span>{inst.label}</span>
            {inst.optional && inst.inputType !== 'formula' && (
              <span className="text-[11px] font-normal text-slate-500">
                (არასავალდებულო)
              </span>
            )}
          </label>

          <div className="flex flex-wrap items-center gap-2">
            <select
              aria-label={`${addr} უჯრის ტიპი`}
              value={inst.inputType}
              onChange={(e) => {
                const nextType = e.target.value as CellInputType;
                handleUpdateInstructionAndSync(inst.id, {
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
              className="text-[11px] text-slate-700 bg-white border border-slate-200 rounded px-1.5 py-0.5 focus:outline-none focus:ring-1 focus:ring-sky-600"
            >
              <option value="text">ტექსტი</option>
              <option value="number">რიცხვი</option>
              <option value="currency">თანხა</option>
              <option value="date">თარიღი</option>
              <option value="formula">ფორმულა (=)</option>
            </select>

            <label className="inline-flex items-center gap-1 text-[11px] text-slate-600 cursor-pointer select-none">
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
                  handleUpdateInstructionAndSync(inst.id, {
                    usePrefix: checked,
                    prefixText: checked ? defaultPref : inst.prefixText,
                  });
                }}
                className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
              />
              <span>სავალდებულო სიტყვა</span>
            </label>

            {inst.inputType !== 'formula' && (
              <label className="inline-flex items-center gap-1 text-[11px] text-slate-600 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={Boolean(inst.optional)}
                  onChange={(e) =>
                    handleUpdateInstructionAndSync(inst.id, {
                      optional: e.target.checked,
                    })
                  }
                  className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                />
                <span>არასავალდებულო</span>
              </label>
            )}
          </div>
        </div>

        <div className="flex items-center gap-2">
          {currentAffixPos === 'start' && affixInputControl}

          {inst.inputType === 'date' ? (
            <div className="flex-1 min-w-0 flex items-center gap-1.5">
              <input
                id={`cell-input-${addr}`}
                type="text"
                placeholder={inst.placeholder || 'DD.MM.YYYY'}
                value={currentInputVal}
                onFocus={() => setSelectedAddress(addr)}
                onChange={(e) => handleFieldChange(addr, e.target.value, 'date')}
                className="flex-1 min-w-0 px-2.5 py-1.5 text-xs font-mono-tabular bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
              />
              <input
                type="date"
                aria-label={`${addr} კალენდრიდან თარიღის არჩევა`}
                value={isoDateValue}
                onChange={(e) => {
                  const formatted = isoDateToGeorgianDot(e.target.value);
                  if (formatted) {
                    handleFieldChange(addr, formatted, 'date');
                  }
                }}
                className="px-2 py-1 text-xs bg-white border border-slate-300 rounded-md text-slate-700 focus:outline-none focus:ring-2 focus:ring-sky-600 shrink-0 cursor-pointer"
              />
            </div>
          ) : inst.inputType === 'formula' ? (
            <div className="flex-1 min-w-0 flex flex-wrap items-center gap-1.5">
              <input
                id={`cell-input-${addr}`}
                type="text"
                placeholder="მაგ: =B17*C17 ან =SUM(D17:D24)"
                value={currentInputVal}
                onFocus={() => setSelectedAddress(addr)}
                onChange={(e) => {
                  const val = e.target.value;
                  handleFieldChange(addr, val, 'formula');
                  handleUpdateInstructionAndSync(inst.id, {
                    formulaExpression: val,
                  });
                }}
                className="flex-1 min-w-[140px] px-2.5 py-1.5 text-xs font-mono bg-emerald-50/40 border border-emerald-400 text-slate-900 rounded-md focus:outline-none focus:ring-2 focus:ring-emerald-600"
              />

              <select
                aria-label={`${addr} ფორმულის შედეგის ფორმატი`}
                value={activeFormulaFormat}
                onChange={(e) => {
                  const fmt = e.target.value as FormulaResultFormat;
                  handleUpdateInstructionAndSync(inst.id, {
                    formulaResultFormat: fmt,
                    currencySymbol:
                      fmt === 'currency'
                        ? inst.currencySymbol || '₾'
                        : inst.currencySymbol,
                  });
                }}
                className="px-2 py-1.5 text-xs bg-white border border-slate-300 rounded-md text-slate-800 shrink-0 cursor-pointer"
              >
                <option value="number">რიცხვი</option>
                <option value="currency">თანხა</option>
              </select>

              {activeFormulaFormat === 'currency' && (
                <div
                  role="group"
                  aria-label={`${addr} ფორმულის ვალუტის არჩევა`}
                  className="inline-flex items-center bg-slate-200/80 p-0.5 rounded-md shrink-0"
                >
                  <button
                    type="button"
                    onClick={() =>
                      handleUpdateInstructionAndSync(inst.id, {
                        currencySymbol: '₾',
                      })
                    }
                    className={`px-2 py-1 text-xs font-semibold rounded transition-colors cursor-pointer whitespace-nowrap ${
                      activeCurrency === '₾'
                        ? 'bg-sky-600 text-white shadow-2xs'
                        : 'text-slate-700 hover:text-slate-900'
                    }`}
                  >
                    ₾ ლარი
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleUpdateInstructionAndSync(inst.id, {
                        currencySymbol: '$',
                      })
                    }
                    className={`px-2 py-1 text-xs font-semibold rounded transition-colors cursor-pointer whitespace-nowrap ${
                      activeCurrency === '$'
                        ? 'bg-sky-600 text-white shadow-2xs'
                        : 'text-slate-700 hover:text-slate-900'
                    }`}
                  >
                    $ დოლარი
                  </button>
                </div>
              )}
            </div>
          ) : (
            <div className="flex-1 min-w-0 flex items-center gap-1.5">
              <input
                id={`cell-input-${addr}`}
                type={
                  (inst.inputType === 'number' || inst.inputType === 'currency') &&
                  !inst.usePrefix
                    ? 'number'
                    : 'text'
                }
                step="any"
                placeholder={inst.placeholder || 'ჩაწერეთ მნიშვნელობა...'}
                value={currentInputVal}
                onFocus={() => setSelectedAddress(addr)}
                onChange={(e) =>
                  handleFieldChange(addr, e.target.value, inst.inputType)
                }
                className="flex-1 min-w-0 px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-sky-600"
              />

              {inst.inputType === 'currency' && (
                <div
                  role="group"
                  aria-label={`${addr} თანხის ვალუტის არჩევა`}
                  className="inline-flex items-center bg-slate-200/80 p-0.5 rounded-md shrink-0"
                >
                  <button
                    type="button"
                    onClick={() =>
                      handleUpdateInstructionAndSync(inst.id, {
                        currencySymbol: '₾',
                      })
                    }
                    className={`px-2 py-1 text-xs font-semibold rounded transition-colors cursor-pointer whitespace-nowrap ${
                      activeCurrency === '₾'
                        ? 'bg-sky-600 text-white shadow-2xs'
                        : 'text-slate-700 hover:text-slate-900'
                    }`}
                  >
                    ₾ ლარი
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      handleUpdateInstructionAndSync(inst.id, {
                        currencySymbol: '$',
                      })
                    }
                    className={`px-2 py-1 text-xs font-semibold rounded transition-colors cursor-pointer whitespace-nowrap ${
                      activeCurrency === '$'
                        ? 'bg-sky-600 text-white shadow-2xs'
                        : 'text-slate-700 hover:text-slate-900'
                    }`}
                  >
                    $ დოლარი
                  </button>
                </div>
              )}
            </div>
          )}

          {currentAffixPos === 'end' && affixInputControl}
        </div>

        {isFormulaType && String(currentInputVal).trim() !== '' && (
          <div className="text-[11px] text-emerald-900 bg-emerald-50/80 px-2.5 py-1 rounded border border-emerald-200 flex items-center justify-between gap-2">
            <span className="font-mono">ფორმულა: {String(composedPreview)}</span>
            <span>
              უჯრაში ჩაიწერება:{' '}
              <strong className="font-mono-tabular">
                {finalFormulaDisplay || '— (ცარიელი)'}
              </strong>
            </span>
          </div>
        )}

        {!isFormulaType && inst.inputType === 'currency' && !inst.usePrefix && (
          <div className="text-[11px] text-slate-600 bg-white px-2 py-1 rounded border border-slate-200/80 flex items-center justify-between">
            <span>
              ვალუტა:{' '}
              <strong className="text-sky-800">
                {activeCurrency === '₾' ? 'ლარი (₾)' : 'დოლარი ($)'}
              </strong>
            </span>
            <span>
              უჯრაში ჩაიწერება:{' '}
              <strong className="text-slate-900 font-mono-tabular">
                {currencyFormattedPreview || '— (ცარიელი დარჩება)'}
              </strong>
            </span>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Top Bar: Brand and Main Profile Pages (მაიპოს, ბითრეიდ) + Settings Navigation */}
      <header className="flex flex-wrap items-center justify-between gap-3 px-4 sm:px-6 py-3 bg-white border-b border-slate-200">
        <a
          href="#top"
          onClick={(e) => {
            e.preventDefault();
            handleSwitchProfile('mypos', 'home');
          }}
          className="text-base sm:text-lg font-bold tracking-tight text-slate-900 font-display whitespace-nowrap shrink-0"
        >
          InvoExcel
        </a>

        <nav
          aria-label="მთავარი ნავიგაცია"
          className="flex items-center gap-1.5 sm:gap-2.5 text-xs sm:text-sm font-medium"
        >
          <button
            type="button"
            onClick={() => handleSwitchProfile('mypos', 'home')}
            className={`px-3 py-1.5 rounded-lg inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeView === 'home' && activeProfile === 'mypos'
                ? 'bg-sky-50 text-sky-800 font-semibold border border-sky-200'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <FileSpreadsheet className="w-4 h-4 shrink-0" />
            <span>მაიპოს</span>
          </button>

          <button
            type="button"
            onClick={() => handleSwitchProfile('btrade', 'home')}
            className={`px-3 py-1.5 rounded-lg inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeView === 'home' && activeProfile === 'btrade'
                ? 'bg-sky-50 text-sky-800 font-semibold border border-sky-200'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <FileSpreadsheet className="w-4 h-4 shrink-0" />
            <span>ბითრეიდ</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveView('settings')}
            className={`px-3 py-1.5 rounded-lg inline-flex items-center gap-1.5 transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              activeView === 'settings'
                ? 'bg-sky-50 text-sky-800 font-semibold border border-sky-200'
                : 'text-slate-600 hover:text-slate-900 hover:bg-slate-100'
            }`}
          >
            <Settings className="w-4 h-4 shrink-0" />
            <span>პარამეტრები</span>
          </button>
        </nav>
      </header>

      {/* Main Workspace Container */}
      <main className="flex-1 max-w-[1440px] w-full mx-auto px-3 sm:px-6 py-4 sm:py-6">
        {errorBanner && (
          <div
            role="alert"
            className="max-w-2xl mx-auto mb-4 p-3.5 rounded-lg bg-red-50 border border-red-200 text-xs text-red-800 flex items-center justify-between gap-2"
          >
            <span>{errorBanner}</span>
            <button
              type="button"
              onClick={() => setErrorBanner(null)}
              className="text-red-700 font-medium underline whitespace-nowrap shrink-0"
            >
              დახურვა
            </button>
          </div>
        )}

        {statusBanner && (
          <div
            role="status"
            className="max-w-2xl mx-auto mb-4 p-3.5 rounded-lg bg-sky-50 border border-sky-200 text-xs text-sky-900 flex items-center justify-between gap-2"
          >
            <div className="flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-sky-700 shrink-0" />
              <span>{statusBanner}</span>
            </div>
            <button
              type="button"
              onClick={() => setStatusBanner(null)}
              className="text-sky-800 font-medium underline whitespace-nowrap shrink-0"
            >
              დახურვა
            </button>
          </div>
        )}

        {activeView === 'home' ? (
          /* MAIN PAGE ('home'): Clean, mobile-adapted fillable fields ONLY for activeProfile (მაიპოს or ბითრეიდ) */
          <div className="max-w-2xl mx-auto w-full space-y-4 sm:space-y-5">
            {activeProfile === 'btrade' && !isCustomDefault && (
              <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-200 text-xs text-amber-950 flex flex-wrap items-center justify-between gap-3">
                <span>
                  <strong>ბითრეიდ:</strong> საკუთარი Excel (.xlsx) ფაილის დასაყენებლად და მისი უჯრების ინსტრუქციების მოსანიშნად გადადით პარამეტრებში.
                </span>
                <button
                  type="button"
                  onClick={() => {
                    setActiveView('settings');
                    setSettingsTab('template');
                  }}
                  className="px-3 py-1.5 bg-amber-900 hover:bg-amber-800 text-white font-semibold rounded-lg transition-colors cursor-pointer shrink-0"
                >
                  ბითრეიდის შაბლონის დაყენება →
                </button>
              </div>
            )}

            <div className="bg-white p-4 sm:p-6 rounded-xl border border-slate-200 space-y-5">
              <form onSubmit={handleGenerateNewExcel} className="space-y-6">
                {/* SECTION 1: 6 MAIN HEADER FIELDS */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
                  {mainHeaderConfigs.map((cfg) => {
                    const val = mainHeaderValues[cfg.key] ?? '';
                    const isoDate =
                      cfg.inputType === 'date' ? georgianDotToIsoDate(val) : '';

                    return (
                      <div key={cfg.key} className="space-y-1.5 min-w-0">
                        <label
                          htmlFor={`main-field-${cfg.key}`}
                          className="block text-xs font-semibold text-slate-800"
                        >
                          {cfg.label}
                        </label>

                        {cfg.inputType === 'date' ? (
                          <div className="flex items-center gap-1.5 min-w-0">
                            <input
                              id={`main-field-${cfg.key}`}
                              type="text"
                              value={val}
                              onChange={(e) =>
                                handleMainHeaderValueChange(
                                  cfg.key,
                                  e.target.value
                                )
                              }
                              className="flex-1 min-w-0 px-3 py-2 text-sm sm:text-xs font-mono-tabular bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-600"
                            />
                            <input
                              type="date"
                              aria-label="თარიღის კალენდრიდან არჩევა"
                              value={isoDate}
                              onChange={(e) => {
                                const dot = isoDateToGeorgianDot(e.target.value);
                                if (dot) {
                                  handleMainHeaderValueChange(cfg.key, dot);
                                }
                              }}
                              className="w-9 px-1.5 py-2 sm:py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-700 cursor-pointer shrink-0"
                              title="კალენდრიდან არჩევა"
                            />
                          </div>
                        ) : (
                          <input
                            id={`main-field-${cfg.key}`}
                            type={cfg.inputType === 'number' ? 'number' : 'text'}
                            step="any"
                            value={val}
                            onChange={(e) =>
                              handleMainHeaderValueChange(cfg.key, e.target.value)
                            }
                            className={`w-full min-w-0 px-3 py-2 text-sm sm:text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-600 ${
                              cfg.inputType === 'number'
                                ? 'font-mono-tabular'
                                : ''
                            }`}
                          />
                        )}
                      </div>
                    );
                  })}

                  {/* ფაილის სახელი (defaults to კომპანიის სახელი unless manually changed) */}
                  <div className="space-y-1.5 min-w-0 sm:col-span-2">
                    <label
                      htmlFor="main-field-fileName"
                      className="block text-xs font-semibold text-slate-800"
                    >
                      ფაილის სახელი
                    </label>
                    <input
                      id="main-field-fileName"
                      type="text"
                      value={customFileName}
                      onChange={(e) => {
                        const nextVal = e.target.value;
                        setCustomFileName(nextVal);
                        setIsFileNameCustomized(
                          nextVal.trim() !== '' &&
                            nextVal !== mainHeaderValues.companyName
                        );
                        setStatusBanner(null);
                      }}
                      onBlur={() => {
                        if (
                          customFileName.trim() === '' &&
                          mainHeaderValues.companyName.trim() !== ''
                        ) {
                          setCustomFileName(mainHeaderValues.companyName);
                          setIsFileNameCustomized(false);
                        }
                      }}
                      className="w-full min-w-0 px-3 py-2 text-sm sm:text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-600"
                    />
                  </div>
                </div>

                {/* SECTION 2: PRODUCTS (Max 8 Products: 1 initial + up to 7 added) */}
                <div className="space-y-3.5 pt-4 border-t border-slate-200">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <h2 className="text-xs font-semibold text-slate-800 uppercase tracking-wider">
                      პროდუქტები ({productRows.length} / {MAX_PRODUCT_ROWS})
                    </h2>

                    <button
                      type="button"
                      onClick={handleAddProductRow}
                      disabled={productRows.length >= MAX_PRODUCT_ROWS}
                      className={`px-3 py-2 sm:py-1.5 text-xs font-semibold rounded-lg inline-flex items-center gap-1.5 transition-colors cursor-pointer whitespace-nowrap ${
                        productRows.length >= MAX_PRODUCT_ROWS
                          ? 'bg-slate-100 text-slate-400 cursor-not-allowed'
                          : 'bg-sky-600 hover:bg-sky-700 text-white'
                      }`}
                    >
                      <Plus className="w-3.5 h-3.5 shrink-0" />
                      <span>
                        {productRows.length >= MAX_PRODUCT_ROWS
                          ? 'მაქსიმუმ 8 პროდუქტი'
                          : 'პროდუქტის დამატება'}
                      </span>
                    </button>
                  </div>

                  <div className="space-y-3">
                    {productRows.map((row, idx) => {
                      const calcRow = liveProductsSummary.rows[idx];
                      const rowNum = idx + 1;
                      const isFirst = idx === 0;
                      const priceTitle = isFirst ? 'ფასი' : 'ერთეულის ფასი';

                      return (
                        <div
                          key={row.id}
                          className="p-3 sm:p-3.5 rounded-lg border border-slate-200 bg-slate-50/70 space-y-3"
                        >
                          {!isFirst && (
                            <div className="flex items-center justify-between gap-2">
                              <span className="text-xs font-semibold text-slate-700">
                                პროდუქტი #{rowNum}
                              </span>
                              <button
                                type="button"
                                aria-label={`პროდუქტი #${rowNum}-ის წაშლა`}
                                onClick={() => handleRemoveProductRow(idx)}
                                className="text-xs text-slate-400 hover:text-red-600 inline-flex items-center gap-1 transition-colors cursor-pointer"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                                <span>წაშლა</span>
                              </button>
                            </div>
                          )}

                          <div className="grid grid-cols-2 sm:grid-cols-12 gap-2.5">
                            {/* პროდუქტი */}
                            <div className="col-span-2 sm:col-span-5 min-w-0">
                              <label
                                htmlFor={`prod-name-${rowNum}`}
                                className="block text-xs font-medium text-slate-700 mb-1"
                              >
                                პროდუქტი
                              </label>
                              <input
                                id={`prod-name-${rowNum}`}
                                type="text"
                                value={row.productName}
                                onChange={(e) =>
                                  handleProductRowChange(idx, {
                                    productName: e.target.value,
                                  })
                                }
                                className="w-full min-w-0 px-3 py-2 text-sm sm:text-xs bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-600"
                              />
                            </div>

                            {/* რაოდენობა */}
                            <div className="col-span-1 sm:col-span-3 min-w-0">
                              <label
                                htmlFor={`prod-qty-${rowNum}`}
                                className="block text-xs font-medium text-slate-700 mb-1"
                              >
                                რაოდენობა
                              </label>
                              <input
                                id={`prod-qty-${rowNum}`}
                                type="number"
                                step="any"
                                value={row.quantity}
                                onChange={(e) =>
                                  handleProductRowChange(idx, {
                                    quantity: e.target.value,
                                  })
                                }
                                className="w-full min-w-0 px-3 py-2 text-sm sm:text-xs font-mono-tabular bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-600"
                              />
                            </div>

                            {/* ფასი / ერთეულის ფასი + ლარი/დოლარი */}
                            <div className="col-span-1 sm:col-span-4 min-w-0">
                              <label
                                htmlFor={`prod-price-${rowNum}`}
                                className="block text-xs font-medium text-slate-700 mb-1"
                              >
                                {priceTitle}
                              </label>
                              <div className="flex items-center gap-1 min-w-0">
                                <input
                                  id={`prod-price-${rowNum}`}
                                  type="number"
                                  step="any"
                                  value={row.unitPrice}
                                  onChange={(e) =>
                                    handleProductRowChange(idx, {
                                      unitPrice: e.target.value,
                                    })
                                  }
                                  className="w-full min-w-0 px-2.5 py-2 text-sm sm:text-xs font-mono-tabular bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-sky-600"
                                />

                                {/* Currency Toggle: Default Lari (₾) vs Dollar ($) */}
                                <div
                                  role="group"
                                  aria-label={`პროდუქტი #${rowNum} ვალუტა`}
                                  className="inline-flex items-center bg-slate-200/90 p-0.5 rounded-lg shrink-0"
                                >
                                  <button
                                    type="button"
                                    onClick={() =>
                                      handleProductRowChange(idx, {
                                        currency: '₾',
                                      })
                                    }
                                    title="ლარი (₾)"
                                    className={`px-2 py-1.5 sm:py-1 text-xs font-bold rounded-md transition-colors cursor-pointer ${
                                      row.currency === '₾'
                                        ? 'bg-sky-600 text-white shadow-2xs'
                                        : 'text-slate-700 hover:text-slate-900'
                                    }`}
                                  >
                                    ₾
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      handleProductRowChange(idx, {
                                        currency: '$',
                                      })
                                    }
                                    title="დოლარი ($)"
                                    className={`px-2 py-1.5 sm:py-1 text-xs font-bold rounded-md transition-colors cursor-pointer ${
                                      row.currency === '$'
                                        ? 'bg-sky-600 text-white shadow-2xs'
                                        : 'text-slate-700 hover:text-slate-900'
                                    }`}
                                  >
                                    $
                                  </button>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* Instant Live Row Price Calculation (before clicking "შექმნა") */}
                          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-slate-200/80 text-xs">
                            <span className="text-slate-600">
                              {calcRow.currency === '$' ? (
                                <>
                                  <strong className="font-mono-tabular text-slate-800">
                                    {calcRow.qtyNumeric} ×{' '}
                                    {formatCurrencyCellValue(
                                      calcRow.priceNumeric,
                                      '$'
                                    )}{' '}
                                    = {calcRow.formattedOriginal}
                                  </strong>{' '}
                                  <span className="text-sky-700 font-medium">
                                    (×{' '}
                                    {liveProductsSummary.hasValidUsdRate
                                      ? liveProductsSummary.usdRateNumeric
                                      : '1'}
                                    )
                                  </span>
                                </>
                              ) : (
                                <strong className="font-mono-tabular text-slate-800">
                                  {calcRow.qtyNumeric} ×{' '}
                                  {formatCurrencyCellValue(
                                    calcRow.priceNumeric,
                                    '₾'
                                  )}
                                </strong>
                              )}
                            </span>

                            <span className="text-slate-800 font-medium">
                              ჯამი:{' '}
                              <strong className="font-mono-tabular text-sky-800">
                                {calcRow.formattedGel}
                              </strong>
                            </span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* SECTION 3: LIVE TOTAL SUMMARY BOX (Always in Lari ₾, calculated BEFORE clicking Create) */}
                <div
                  data-testid="live-total-summary"
                  className="p-3.5 sm:p-4 rounded-xl bg-sky-50/90 border border-sky-200 flex items-center justify-between gap-3"
                >
                  <div className="text-xs sm:text-sm font-bold text-slate-900">
                    სულ ჯამი (ლარში):
                  </div>
                  <div
                    data-testid="live-grand-total-gel"
                    className="text-lg sm:text-xl font-bold font-mono-tabular text-sky-900"
                  >
                    {liveProductsSummary.formattedGrandTotalGel}
                  </div>
                </div>

                {/* SECTION 4: CREATE & DOWNLOAD BUTTONS */}
                <div className="space-y-2.5">
                  <button
                    type="submit"
                    disabled={isGeneratingExcel}
                    className="w-full py-2.5 px-4 bg-sky-600 hover:bg-sky-700 text-white text-xs sm:text-sm font-semibold rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer"
                  >
                    <Calculator className="w-4 h-4 shrink-0" />
                    <span>
                      {isGeneratingExcel ? 'მიმდინარეობს შექმნა...' : 'შექმნა'}
                    </span>
                  </button>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                    <button
                      type="button"
                      onClick={handleDownloadGeneratedExcel}
                      className="py-2.5 px-4 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer"
                    >
                      <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0" />
                      <span>Excel (.xlsx) გადმოწერა</span>
                    </button>

                    <button
                      type="button"
                      onClick={handleDownloadGeneratedPdf}
                      disabled={isExportingPdf}
                      className="py-2.5 px-4 bg-slate-800 hover:bg-slate-700 text-white text-xs font-semibold rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer"
                    >
                      <FileText className="w-4 h-4 text-sky-300 shrink-0" />
                      <span>
                        {isExportingPdf ? 'მზადდება PDF...' : 'PDF (.pdf) გადმოწერა'}
                      </span>
                    </button>
                  </div>
                </div>
              </form>
            </div>

            {/* Offscreen hidden render container so PDF export works directly from the Main Page without displaying the Excel file */}
            {sheetModel && (
              <div
                aria-hidden="true"
                className="fixed -left-[10000px] top-0 pointer-events-none select-none"
              >
                <ExcelSheetPreview
                  sheetModel={sheetModel}
                  editableInstructions={previewHighlightInstructions}
                  selectedAddress={effectiveSelectedAddress}
                  onSelectCell={() => {}}
                  showExcelGridHeaders={false}
                  highlightEditableCells={false}
                  isExportingPdf={true}
                  sheetPrintRef={sheetPrintRef}
                />
              </div>
            )}
          </div>
        ) : (
          /* SETTINGS VIEW ('settings'): Contains ALL existing configuration tools, instructions & 1:1 Excel sheet preview */
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
            <div className="lg:col-span-5 space-y-5">
              <div className="bg-white p-4 rounded-xl border border-slate-200 space-y-4">
                <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-3">
                  <div>
                    <h1 className="text-base font-bold text-slate-900">
                      პარამეტრები — {PROFILE_LABELS[activeProfile]}
                    </h1>
                    <p className="text-xs text-slate-500 mt-0.5">
                      შაბლონი: {templateName} · უჯრების ინსტრუქციები და კოორდინატები
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setActiveView('home')}
                    className="px-2.5 py-1.5 text-xs font-medium text-sky-700 bg-sky-50 hover:bg-sky-100 rounded-md cursor-pointer shrink-0"
                  >
                    ← {PROFILE_LABELS[activeProfile]}-ზე დაბრუნება
                  </button>
                </div>

                {/* Profile Selector inside Settings: switch between მაიპოს and ბითრეიდ */}
                <div className="p-2.5 rounded-lg bg-sky-50/70 border border-sky-200/80 space-y-1.5">
                  <div className="text-[11px] font-semibold text-sky-950">
                    აირჩიეთ გვერდი, რომლის შაბლონს და ინსტრუქციებსაც აყენებთ:
                  </div>
                  <div
                    role="group"
                    aria-label="შაბლონის გვერდის არჩევა"
                    className="grid grid-cols-2 gap-1.5"
                  >
                    <button
                      type="button"
                      onClick={() => handleSwitchProfile('mypos')}
                      className={`py-1.5 px-3 text-xs font-semibold rounded-md transition-colors cursor-pointer ${
                        activeProfile === 'mypos'
                          ? 'bg-sky-600 text-white shadow-2xs'
                          : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
                      }`}
                    >
                      მაიპოს (MyPos)
                    </button>
                    <button
                      type="button"
                      onClick={() => handleSwitchProfile('btrade')}
                      className={`py-1.5 px-3 text-xs font-semibold rounded-md transition-colors cursor-pointer ${
                        activeProfile === 'btrade'
                          ? 'bg-sky-600 text-white shadow-2xs'
                          : 'bg-white text-slate-700 hover:bg-slate-100 border border-slate-200'
                      }`}
                    >
                      ბითრეიდ (Btrade)
                    </button>
                  </div>
                </div>

                {/* Settings Sub-Tabs */}
                <div
                  role="tablist"
                  aria-label="პარამეტრების რეჟიმები"
                  className="grid grid-cols-3 gap-1 p-1 bg-slate-100 rounded-lg"
                >
                  <button
                    type="button"
                    role="tab"
                    aria-selected={settingsTab === 'fields'}
                    onClick={() => setSettingsTab('fields')}
                    className={`px-2 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                      settingsTab === 'fields'
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    1. ველები
                  </button>
                  <button
                    type="button"
                    role="tab"
                    aria-selected={settingsTab === 'instructions'}
                    onClick={() => setSettingsTab('instructions')}
                    className={`px-2 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                      settingsTab === 'instructions'
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    2. ინსტრუქციები
                  </button>
                  <button
                    type="button"
                    role="tab"
                    aria-selected={settingsTab === 'template'}
                    onClick={() => setSettingsTab('template')}
                    className={`px-2 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap cursor-pointer ${
                      settingsTab === 'template'
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    3. შაბლონი
                  </button>
                </div>

                {/* SETTINGS TAB 1: Main Page Coordinate Mapping + All Editable Cells */}
                {settingsTab === 'fields' && (
                  <div className="space-y-4">
                    <div className="p-3.5 rounded-lg border border-slate-200 bg-slate-50/70 space-y-3">
                      <div className="text-xs font-bold text-slate-900">
                        „{PROFILE_LABELS[activeProfile]}“-ის გვერდის ველების უჯრების კოორდინატები
                      </div>
                      <p className="text-[11px] text-slate-500">
                        აქ შეგიძლიათ შეცვალოთ, რომელ უჯრაში ჩაიწეროს „{PROFILE_LABELS[activeProfile]}“-ის გვერდის თითოეული ველი და სავალდებულო საწყისი სიტყვა:
                      </p>

                      <div className="space-y-2">
                        {mainHeaderConfigs.map((cfg) => (
                          <div
                            key={cfg.key}
                            className="p-2 rounded bg-white border border-slate-200 flex flex-wrap items-center justify-between gap-2 text-xs"
                          >
                            <span className="font-medium text-slate-800 min-w-[125px]">
                              {cfg.label}
                            </span>
                            <div className="flex items-center gap-1.5 flex-1">
                              <input
                                type="text"
                                aria-label={`${cfg.label} უჯრის კოორდინატი`}
                                value={cfg.address}
                                onChange={(e) =>
                                  handleUpdateMainHeaderConfig(cfg.key, {
                                    address: e.target.value.toUpperCase(),
                                  })
                                }
                                className="w-16 px-2 py-1 text-xs font-mono bg-slate-50 border border-slate-300 rounded text-center"
                                title="უჯრის კოორდინატი (მაგ: A8, C9)"
                              />
                              <label className="inline-flex items-center gap-1 text-[11px] text-slate-600 cursor-pointer">
                                <input
                                  type="checkbox"
                                  checked={cfg.usePrefix}
                                  onChange={(e) =>
                                    handleUpdateMainHeaderConfig(cfg.key, {
                                      usePrefix: e.target.checked,
                                    })
                                  }
                                  className="rounded border-slate-300 text-sky-600"
                                />
                                <span>სიტყვა:</span>
                              </label>
                              {cfg.usePrefix && (
                                <input
                                  type="text"
                                  aria-label={`${cfg.label} საწყისი სიტყვა`}
                                  value={cfg.prefixText}
                                  onChange={(e) =>
                                    handleUpdateMainHeaderConfig(cfg.key, {
                                      prefixText: e.target.value,
                                    })
                                  }
                                  className="flex-1 min-w-[90px] px-2 py-1 text-xs bg-sky-50 border border-sky-300 rounded"
                                />
                              )}
                            </div>
                          </div>
                        ))}

                        {/* Grand Total Cell Coordinate */}
                        <div className="p-2 rounded bg-white border border-slate-200 flex items-center justify-between gap-2 text-xs">
                          <span className="font-medium text-slate-800">
                            სულ ჯამის უჯრა (ლარში ₾):
                          </span>
                          <input
                            type="text"
                            aria-label="სულ ჯამის უჯრის კოორდინატი"
                            value={grandTotalAddress}
                            onChange={(e) => {
                              const nextAddr = e.target.value.toUpperCase();
                              setGrandTotalAddress(nextAddr);
                              saveGrandTotalAddress(
                                templateName,
                                nextAddr,
                                activeProfile
                              );
                              setIsDirtySinceLastGenerate(true);
                            }}
                            className="w-20 px-2 py-1 text-xs font-mono bg-emerald-50 border border-emerald-400 rounded text-center font-semibold"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Additional custom instructions list */}
                    {instructions.length > 0 && (
                      <div className="space-y-2">
                        <h3 className="text-xs font-semibold text-slate-700">
                          დამატებითი ინსტრუქციის უჯრები ({instructions.length})
                        </h3>
                        <div className="space-y-2 max-h-[380px] overflow-y-auto pr-1">
                          {instructions.map(renderEditableFieldControl)}
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* SETTINGS TAB 2: INSTRUCTIONS CONFIG PANEL */}
                {settingsTab === 'instructions' && (
                  <InstructionsConfigPanel
                    instructions={instructions}
                    sheetModel={sheetModel}
                    selectedAddress={effectiveSelectedAddress}
                    editableValues={editableValues}
                    liveCellValueLookup={liveCellValueLookup}
                    onSelectAddress={(addr) => setSelectedAddress(addr)}
                    onAddInstruction={(newInst, initialUserVal) => {
                      const next = [...instructions, newInst];
                      updateInstructionsList(next);
                      const addr = newInst.address.toUpperCase();
                      const cellVal =
                        initialUserVal !== undefined
                          ? initialUserVal
                          : stripPrefixFromCellValue(
                              sheetModel?.cells[addr]?.value ?? '',
                              newInst.usePrefix,
                              newInst.prefixText,
                              newInst.affixPosition || 'start'
                            );
                      setEditableValues((prev) => ({
                        ...prev,
                        [addr]: cellVal,
                      }));
                      setIsDirtySinceLastGenerate(true);
                    }}
                    onUpdateInstruction={handleUpdateInstructionAndSync}
                    onChangeCellValue={(addr, rawVal, type) =>
                      handleFieldChange(addr, rawVal, type)
                    }
                    onRemoveInstruction={(id) => {
                      const next = instructions.filter((i) => i.id !== id);
                      updateInstructionsList(next);
                      setIsDirtySinceLastGenerate(true);
                    }}
                    onClearAllInstructions={() => {
                      updateInstructionsList([]);
                      setEditableValues({});
                      setIsDirtySinceLastGenerate(true);
                    }}
                    onResetDefaultInstructions={() => {
                      updateInstructionsList(DEFAULT_EDITABLE_INSTRUCTIONS);
                      setMainHeaderConfigs(DEFAULT_MAIN_HEADER_CONFIGS);
                      saveMainHeaderConfigs(
                        templateName,
                        DEFAULT_MAIN_HEADER_CONFIGS,
                        activeProfile
                      );
                      setIsDirtySinceLastGenerate(true);
                    }}
                  />
                )}

                {/* SETTINGS TAB 3: TEMPLATE (.XLSX) MANAGEMENT */}
                {settingsTab === 'template' && (
                  <div className="space-y-4 text-xs">
                    <div className="p-3.5 rounded-lg bg-slate-50 border border-slate-200 space-y-2">
                      <div className="font-semibold text-slate-800">
                        „{PROFILE_LABELS[activeProfile]}“-ის ჩაშენებული (Default) შაბლონი
                      </div>
                      <div className="flex items-center justify-between gap-2 text-slate-600">
                        <span className="font-mono break-all">{templateName}</span>
                        <span className="text-sky-700 font-medium shrink-0">
                          {isCustomDefault
                            ? 'თქვენი ატვირთული (ჩაშენებული)'
                            : 'სისტემური სტანდარტი'}
                        </span>
                      </div>
                      <p className="text-slate-500 leading-relaxed">
                        „{PROFILE_LABELS[activeProfile]}“-ის გვერდზე მონაცემების შეყვანისას პირდაპირ ამ შაბლონის მიხედვით
                        გენერირდება Excel (.xlsx) და PDF (.pdf) ფაილები. ფერები, სურათები და
                        უჯრების სტილები ნარჩუნდება 100%-ით.
                      </p>
                    </div>

                    <div className="space-y-2.5">
                      <label
                        htmlFor="upload-custom-xlsx"
                        className="w-full py-2.5 px-4 bg-slate-900 hover:bg-slate-800 text-white font-medium rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer"
                      >
                        <Upload className="w-4 h-4 shrink-0" />
                        <span>საკუთარი Excel (.xlsx) შაბლონის ატვირთვა / შეცვლა</span>
                      </label>
                      <input
                        id="upload-custom-xlsx"
                        ref={fileInputRef}
                        type="file"
                        accept=".xlsx"
                        onChange={handleUploadTemplateFile}
                        className="sr-only"
                      />

                      {templateBuffer && (
                        <button
                          type="button"
                          onClick={() =>
                            downloadExcelBuffer(
                              templateBuffer,
                              `შაბლონი_${templateName}`
                            )
                          }
                          className="w-full py-2 px-4 bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 font-medium rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer"
                        >
                          <Download className="w-3.5 h-3.5 shrink-0" />
                          <span>მიმდინარე საწყისი შაბლონის გადმოწერა</span>
                        </button>
                      )}

                      {isCustomDefault && (
                        <button
                          type="button"
                          onClick={handleRestoreFactoryTemplate}
                          className="w-full py-2 px-4 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer"
                        >
                          <RefreshCw className="w-3.5 h-3.5 shrink-0" />
                          <span>ქარხნული ინვოისის შაბლონის დაბრუნება</span>
                        </button>
                      )}
                    </div>
                  </div>
                )}
              </div>

              {/* DOWNLOAD CARD in Settings */}
              <section
                aria-label="გენერირებული ფაილის გადმოწერის არჩევანი"
                className="bg-white p-4 rounded-xl border border-slate-200 space-y-3.5"
              >
                <div className="flex items-center justify-between gap-2 border-b border-slate-100 pb-2.5">
                  <div>
                    <h2 className="text-sm font-bold text-slate-900">
                      მზა ფაილების გადმოწერა (Excel და PDF)
                    </h2>
                    <div className="flex items-center gap-2 text-xs text-slate-500 mt-0.5">
                      <span>
                        ჯამი ლარში:{' '}
                        <strong className="text-slate-900 font-mono-tabular">
                          {liveProductsSummary.formattedGrandTotalGel}
                        </strong>
                      </span>
                      {generatedPackage && (
                        <>
                          <span aria-hidden="true">·</span>
                          <span>განახლდა: {generatedPackage.generatedAt}</span>
                        </>
                      )}
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <button
                    type="button"
                    onClick={handleDownloadGeneratedExcel}
                    className="py-2.5 px-4 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer whitespace-nowrap"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-400 shrink-0" />
                    <span>Excel (.xlsx) გადმოწერა</span>
                  </button>

                  <button
                    type="button"
                    onClick={handleDownloadGeneratedPdf}
                    disabled={isExportingPdf}
                    className="py-2.5 px-4 bg-sky-600 hover:bg-sky-700 text-white text-xs font-semibold rounded-lg transition-colors flex items-center justify-center gap-2 cursor-pointer whitespace-nowrap"
                  >
                    <FileText className="w-4 h-4 text-white shrink-0" />
                    <span>
                      {isExportingPdf ? 'მზადდება PDF...' : 'PDF (.pdf) გადმოწერა'}
                    </span>
                  </button>
                </div>
              </section>
            </div>

            {/* RIGHT WORKSPACE COLUMN IN SETTINGS: 1:1 Single-Page Excel & PDF Sheet Preview */}
            <div id="sheet-preview-section" className="lg:col-span-7 space-y-3">
              <div className="bg-white px-4 py-3 rounded-xl border border-slate-200 space-y-2.5 text-xs">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-2 text-slate-700">
                    <span className="font-semibold text-slate-900">
                      1:1 ფურცლის ესკიზი (A4)
                    </span>
                    <span aria-hidden="true">·</span>
                    <span className="text-slate-600 font-medium">
                      {effectiveSelectedAddress
                        ? `არჩეულია უჯრა ${effectiveSelectedAddress}`
                        : 'აირჩიეთ უჯრა'}
                    </span>
                    {selectedCellInfo?.formula && (
                      <>
                        <span aria-hidden="true">·</span>
                        <span className="font-mono text-emerald-700">
                          ={selectedCellInfo.formula}
                        </span>
                      </>
                    )}
                  </div>

                  <div className="flex items-center gap-4">
                    <label className="inline-flex items-center gap-1.5 text-slate-600 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={showExcelGridHeaders}
                        onChange={(e) => setShowExcelGridHeaders(e.target.checked)}
                        className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                      />
                      <span>კოორდინატების ბადე</span>
                    </label>

                    <label className="inline-flex items-center gap-1.5 text-slate-600 cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={highlightEditableCells}
                        onChange={(e) =>
                          setHighlightEditableCells(e.target.checked)
                        }
                        className="rounded border-slate-300 text-sky-600 focus:ring-sky-600"
                      />
                      <span>შესაცვლელი უჯრების ჩარჩო</span>
                    </label>
                  </div>
                </div>

                {effectiveSelectedAddress && (
                  <div className="pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
                    <div className="text-slate-600">
                      <span className="font-mono font-semibold text-slate-900 mr-1.5">
                        [{effectiveSelectedAddress}]
                      </span>
                      <span>
                        {selectedCellInfo?.formattedValue
                          ? `„${selectedCellInfo.formattedValue}“`
                          : '(ცარიელი უჯრა)'}
                      </span>
                    </div>

                    <div className="flex flex-wrap items-center gap-1.5">
                      {selectedCellInstruction ? (
                        <button
                          type="button"
                          onClick={() => handleQuickAddSelectedCell(false)}
                          className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium rounded transition-colors inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                        >
                          <Trash2 className="w-3 h-3 text-red-600" />
                          <span>შესაცვლელი სიიდან მოხსნა</span>
                        </button>
                      ) : (
                        <>
                          <button
                            type="button"
                            onClick={() => handleQuickAddSelectedCell(false)}
                            className="px-2.5 py-1 bg-sky-600 hover:bg-sky-700 text-white font-medium rounded transition-colors inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                          >
                            <Plus className="w-3 h-3" />
                            <span>შესაცვლელ უჯრად დამატება</span>
                          </button>
                          <button
                            type="button"
                            onClick={() =>
                              handleQuickAddSelectedCell(true, undefined, 'start')
                            }
                            className="px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-white font-medium rounded transition-colors inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                          >
                            <Plus className="w-3 h-3" />
                            <span>+ სიტყვა თავში</span>
                          </button>
                          <button
                            type="button"
                            onClick={() =>
                              handleQuickAddSelectedCell(true, undefined, 'end')
                            }
                            className="px-2.5 py-1 bg-slate-700 hover:bg-slate-800 text-white font-medium rounded transition-colors inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                          >
                            <Plus className="w-3 h-3" />
                            <span>+ სიტყვა ბოლოში</span>
                          </button>
                          <button
                            type="button"
                            onClick={() =>
                              handleQuickAddSelectedCell(false, 'formula')
                            }
                            className="px-2.5 py-1 bg-emerald-700 hover:bg-emerald-800 text-white font-medium rounded transition-colors inline-flex items-center gap-1 cursor-pointer whitespace-nowrap"
                          >
                            <Calculator className="w-3 h-3" />
                            <span>ფორმულის ჩაწერა (=)</span>
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                )}
              </div>

              {/* 1:1 Sheet Surface in Settings */}
              <div className="bg-slate-200/70 p-3 rounded-xl border border-slate-200 overflow-x-auto">
                <div className="w-fit mx-auto">
                  {sheetModel ? (
                    <ExcelSheetPreview
                      sheetModel={sheetModel}
                      editableInstructions={previewHighlightInstructions}
                      selectedAddress={effectiveSelectedAddress}
                      onSelectCell={(addr) => setSelectedAddress(addr)}
                      showExcelGridHeaders={showExcelGridHeaders}
                      highlightEditableCells={highlightEditableCells}
                      isExportingPdf={isExportingPdf}
                      sheetPrintRef={sheetPrintRef}
                    />
                  ) : (
                    <div className="w-full max-w-[740px] h-[860px] bg-white rounded-lg border border-slate-200 p-8 flex items-center justify-center text-sm text-slate-500">
                      იტვირთება ექსელის შაბლონი...
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
