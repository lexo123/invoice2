import {
  EditableCellInstruction,
  InvoiceProfileId,
  MainHeaderFieldConfig,
} from '../types/invoice';
import { BTRADE_VALID_B64_SUFFIX } from './btradeTemplatePatch';
import embeddedBtradeTemplate from './embeddedBtradeTemplate.json';
import embeddedCustomTemplate from './embeddedCustomTemplate.json';

interface EmbeddedTemplateData {
  base64?: string;
  fileName?: string;
  instructions?: EditableCellInstruction[];
  mainHeaderConfigs?: MainHeaderFieldConfig[];
  grandTotalAddress?: string;
}

const STORAGE_KEYS = {
  mypos: {
    INSTRUCTIONS_PREFIX: 'invoexcel_instructions_v2_',
    MAIN_FIELDS_PREFIX: 'invoexcel_main_fields_v1_',
    GRAND_TOTAL_PREFIX: 'invoexcel_grand_total_v1_',
    CUSTOM_TEMPLATE_B64: 'invoexcel_custom_template_b64_v1',
    CUSTOM_TEMPLATE_NAME: 'invoexcel_custom_template_name_v1',
  },
  btrade: {
    INSTRUCTIONS_PREFIX: 'invoexcel_btrade_instructions_v3_',
    MAIN_FIELDS_PREFIX: 'invoexcel_btrade_main_fields_v3_',
    GRAND_TOTAL_PREFIX: 'invoexcel_btrade_grand_total_v3_',
    CUSTOM_TEMPLATE_B64: 'invoexcel_btrade_template_b64_v3',
    CUSTOM_TEMPLATE_NAME: 'invoexcel_btrade_template_name_v3',
  },
};

function getEmbeddedForProfile(profileId: InvoiceProfileId): EmbeddedTemplateData {
  if (profileId === 'btrade') {
    const raw = embeddedBtradeTemplate as EmbeddedTemplateData;
    const rawB64 = raw.base64 || '';
    const validBase64 =
      rawB64.length === 43555 || rawB64.length % 4 !== 0
        ? rawB64.slice(0, 8596) + BTRADE_VALID_B64_SUFFIX
        : rawB64;
    return {
      ...raw,
      base64: validBase64,
    };
  }
  return embeddedCustomTemplate as EmbeddedTemplateData;
}

export function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  const chunks: string[] = [];
  for (let i = 0; i < bytes.byteLength; i += chunkSize) {
    const sub = bytes.subarray(i, i + chunkSize);
    chunks.push(String.fromCharCode.apply(null, Array.from(sub)));
  }
  return btoa(chunks.join(''));
}

export function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function resolveCurrentBase64AndName(
  profileId: InvoiceProfileId,
  fallbackFileName?: string
): { base64: string; fileName: string } {
  const keys = STORAGE_KEYS[profileId];
  const embedded = getEmbeddedForProfile(profileId);
  let b64 = '';
  let name = '';
  try {
    b64 = localStorage.getItem(keys.CUSTOM_TEMPLATE_B64) || '';
    name = localStorage.getItem(keys.CUSTOM_TEMPLATE_NAME) || '';
  } catch {
    // ignore
  }
  if (!b64 && embedded?.base64) {
    b64 = embedded.base64;
  }
  if (!name) {
    name =
      fallbackFileName ||
      embedded?.fileName ||
      (profileId === 'btrade' ? 'btrade.xlsx' : 'pos.xlsx');
  }
  return { base64: b64, fileName: name };
}

/**
 * Syncs the active template, instructions, and header coordinates for a given profile ('mypos' or 'btrade')
 * to the repository's JSON files during development so static builds (like GitHub Pages) include everything.
 */
export function syncCustomTemplateToRepo(
  base64: string,
  fileName: string,
  instructions?: EditableCellInstruction[],
  profileId: InvoiceProfileId = 'mypos',
  mainHeaderConfigs?: MainHeaderFieldConfig[] | null,
  grandTotalAddress?: string
): void {
  if (typeof fetch === 'undefined') return;
  if (!base64 && (!instructions || instructions.length === 0)) return;
  try {
    fetch('/__sync_embedded_template', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        profileId,
        base64,
        fileName,
        instructions,
        mainHeaderConfigs: mainHeaderConfigs || undefined,
        grandTotalAddress: grandTotalAddress || undefined,
      }),
    }).catch(() => {
      // Ignore in static production environment (e.g. GitHub Pages)
    });
  } catch {
    // Ignore
  }
}

export function loadSavedInstructionsForTemplate(
  templateName: string,
  fallback: EditableCellInstruction[],
  profileId: InvoiceProfileId = 'mypos'
): EditableCellInstruction[] {
  const keys = STORAGE_KEYS[profileId];
  const embedded = getEmbeddedForProfile(profileId);
  try {
    const raw = localStorage.getItem(`${keys.INSTRUCTIONS_PREFIX}${templateName}`);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch {
    // Ignore localStorage error and fall back to embedded instructions
  }
  if (
    Array.isArray(embedded?.instructions) &&
    embedded.instructions.length > 0
  ) {
    return embedded.instructions as EditableCellInstruction[];
  }
  return fallback;
}

export function saveInstructionsForTemplate(
  templateName: string,
  instructions: EditableCellInstruction[],
  profileId: InvoiceProfileId = 'mypos'
): void {
  const keys = STORAGE_KEYS[profileId];
  try {
    localStorage.setItem(
      `${keys.INSTRUCTIONS_PREFIX}${templateName}`,
      JSON.stringify(instructions)
    );
    const { base64, fileName } = resolveCurrentBase64AndName(profileId, templateName);
    const savedMain = loadSavedMainHeaderConfigs(templateName, profileId);
    const savedTotal = loadSavedGrandTotalAddress(templateName, profileId);
    syncCustomTemplateToRepo(
      base64,
      fileName,
      instructions,
      profileId,
      savedMain,
      savedTotal
    );
  } catch {
    // Ignore storage quota errors
  }
}

export function loadSavedMainHeaderConfigs(
  templateName: string,
  profileId: InvoiceProfileId = 'mypos'
): MainHeaderFieldConfig[] | null {
  const keys = STORAGE_KEYS[profileId];
  const embedded = getEmbeddedForProfile(profileId);
  try {
    const raw = localStorage.getItem(`${keys.MAIN_FIELDS_PREFIX}${templateName}`);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length === 6) {
        return parsed;
      }
    }
  } catch {
    // Ignore localStorage error and fall back to embedded mainHeaderConfigs
  }
  if (
    Array.isArray(embedded?.mainHeaderConfigs) &&
    embedded.mainHeaderConfigs.length === 6
  ) {
    return embedded.mainHeaderConfigs;
  }
  return null;
}

export function saveMainHeaderConfigs(
  templateName: string,
  configs: MainHeaderFieldConfig[],
  profileId: InvoiceProfileId = 'mypos'
): void {
  const keys = STORAGE_KEYS[profileId];
  try {
    localStorage.setItem(
      `${keys.MAIN_FIELDS_PREFIX}${templateName}`,
      JSON.stringify(configs)
    );
    const { base64, fileName } = resolveCurrentBase64AndName(profileId, templateName);
    const savedInst = loadSavedInstructionsForTemplate(templateName, [], profileId);
    const savedTotal = loadSavedGrandTotalAddress(templateName, profileId);
    syncCustomTemplateToRepo(
      base64,
      fileName,
      savedInst,
      profileId,
      configs,
      savedTotal
    );
  } catch {
    // Ignore
  }
}

export function loadSavedGrandTotalAddress(
  templateName: string,
  profileId: InvoiceProfileId = 'mypos',
  fallback = 'D36'
): string {
  const keys = STORAGE_KEYS[profileId];
  const embedded = getEmbeddedForProfile(profileId);
  try {
    const raw = localStorage.getItem(`${keys.GRAND_TOTAL_PREFIX}${templateName}`);
    if (raw && raw.trim()) {
      return raw.trim().toUpperCase();
    }
  } catch {
    // Ignore localStorage error and fall back to embedded grandTotalAddress
  }
  if (embedded?.grandTotalAddress && embedded.grandTotalAddress.trim()) {
    return embedded.grandTotalAddress.trim().toUpperCase();
  }
  return fallback;
}

export function saveGrandTotalAddress(
  templateName: string,
  grandTotalAddress: string,
  profileId: InvoiceProfileId = 'mypos'
): void {
  const keys = STORAGE_KEYS[profileId];
  const clean = grandTotalAddress.trim().toUpperCase() || 'D36';
  try {
    localStorage.setItem(`${keys.GRAND_TOTAL_PREFIX}${templateName}`, clean);
    const { base64, fileName } = resolveCurrentBase64AndName(profileId, templateName);
    const savedInst = loadSavedInstructionsForTemplate(templateName, [], profileId);
    const savedMain = loadSavedMainHeaderConfigs(templateName, profileId);
    syncCustomTemplateToRepo(
      base64,
      fileName,
      savedInst,
      profileId,
      savedMain,
      clean
    );
  } catch {
    // Ignore
  }
}

export function saveCustomDefaultTemplate(
  buffer: ArrayBuffer,
  fileName: string,
  profileId: InvoiceProfileId = 'mypos'
): void {
  const keys = STORAGE_KEYS[profileId];
  try {
    const b64 = arrayBufferToBase64(buffer);
    localStorage.setItem(keys.CUSTOM_TEMPLATE_B64, b64);
    localStorage.setItem(keys.CUSTOM_TEMPLATE_NAME, fileName);
    const savedInst = loadSavedInstructionsForTemplate(fileName, [], profileId);
    const savedMain = loadSavedMainHeaderConfigs(fileName, profileId);
    const savedTotal = loadSavedGrandTotalAddress(fileName, profileId);
    syncCustomTemplateToRepo(
      b64,
      fileName,
      savedInst,
      profileId,
      savedMain,
      savedTotal
    );
  } catch {
    // Ignore storage quota errors
  }
}

export function loadCustomDefaultTemplate(
  profileId: InvoiceProfileId = 'mypos'
): { buffer: ArrayBuffer; fileName: string } | null {
  const keys = STORAGE_KEYS[profileId];
  const embedded = getEmbeddedForProfile(profileId);
  try {
    const b64 = localStorage.getItem(keys.CUSTOM_TEMPLATE_B64);
    const fileName = localStorage.getItem(keys.CUSTOM_TEMPLATE_NAME);
    if (b64 && fileName) {
      const buffer = base64ToArrayBuffer(b64);
      const savedInst = loadSavedInstructionsForTemplate(fileName, [], profileId);
      const savedMain = loadSavedMainHeaderConfigs(fileName, profileId);
      const savedTotal = loadSavedGrandTotalAddress(fileName, profileId);
      syncCustomTemplateToRepo(
        b64,
        fileName,
        savedInst,
        profileId,
        savedMain,
        savedTotal
      );
      return {
        buffer,
        fileName,
      };
    }
  } catch {
    // Ignore localStorage error and fall back to embedded template
  }

  try {
    // Fallback to bundled embedded JSON (used when deployed to GitHub Pages!)
    if (
      embedded &&
      typeof embedded.base64 === 'string' &&
      embedded.base64.length > 0
    ) {
      const embeddedName =
        embedded.fileName ||
        (profileId === 'btrade' ? 'btrade_შაბლონი.xlsx' : 'ჩაშენებული_შაბლონი.xlsx');
      const savedInst = loadSavedInstructionsForTemplate(
        embeddedName,
        embedded.instructions || [],
        profileId
      );
      const savedMain = loadSavedMainHeaderConfigs(embeddedName, profileId);
      const savedTotal = loadSavedGrandTotalAddress(embeddedName, profileId);
      syncCustomTemplateToRepo(
        embedded.base64,
        embeddedName,
        savedInst,
        profileId,
        savedMain,
        savedTotal
      );
      return {
        buffer: base64ToArrayBuffer(embedded.base64),
        fileName: embeddedName,
      };
    }
    return null;
  } catch {
    return null;
  }
}

export function clearCustomDefaultTemplate(
  templateName?: string,
  profileId: InvoiceProfileId = 'mypos'
): void {
  const keys = STORAGE_KEYS[profileId];
  try {
    localStorage.removeItem(keys.CUSTOM_TEMPLATE_B64);
    localStorage.removeItem(keys.CUSTOM_TEMPLATE_NAME);
    if (templateName) {
      localStorage.removeItem(`${keys.INSTRUCTIONS_PREFIX}${templateName}`);
      localStorage.removeItem(`${keys.MAIN_FIELDS_PREFIX}${templateName}`);
      localStorage.removeItem(`${keys.GRAND_TOTAL_PREFIX}${templateName}`);
    }
  } catch {
    // Ignore
  }
}

