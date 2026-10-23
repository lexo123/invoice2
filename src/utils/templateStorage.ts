import { EditableCellInstruction, MainHeaderFieldConfig } from '../types/invoice';
import embeddedCustomTemplate from './embeddedCustomTemplate.json';

const STORAGE_KEYS = {
  INSTRUCTIONS_PREFIX: 'invoexcel_instructions_v2_',
  MAIN_FIELDS_PREFIX: 'invoexcel_main_fields_v1_',
  CUSTOM_TEMPLATE_B64: 'invoexcel_custom_template_b64_v1',
  CUSTOM_TEMPLATE_NAME: 'invoexcel_custom_template_name_v1',
};

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

/**
 * Syncs the active custom template from the browser's localStorage to the repository's
 * src/utils/embeddedCustomTemplate.json file during development so that static builds
 * (such as GitHub Pages) have the uploaded Excel file baked into the bundle.
 */
export function syncCustomTemplateToRepo(
  base64: string,
  fileName: string,
  instructions?: EditableCellInstruction[]
): void {
  if (typeof fetch === 'undefined' || !base64) return;
  try {
    fetch('/__sync_embedded_template', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        base64,
        fileName,
        instructions,
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
  fallback: EditableCellInstruction[]
): EditableCellInstruction[] {
  try {
    const raw = localStorage.getItem(`${STORAGE_KEYS.INSTRUCTIONS_PREFIX}${templateName}`);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
    if (
      embeddedCustomTemplate?.fileName === templateName &&
      Array.isArray(embeddedCustomTemplate?.instructions) &&
      embeddedCustomTemplate.instructions.length > 0
    ) {
      return embeddedCustomTemplate.instructions as EditableCellInstruction[];
    }
    return fallback;
  } catch {
    return fallback;
  }
}

export function saveInstructionsForTemplate(
  templateName: string,
  instructions: EditableCellInstruction[]
): void {
  try {
    localStorage.setItem(
      `${STORAGE_KEYS.INSTRUCTIONS_PREFIX}${templateName}`,
      JSON.stringify(instructions)
    );
    const b64 = localStorage.getItem(STORAGE_KEYS.CUSTOM_TEMPLATE_B64);
    const name = localStorage.getItem(STORAGE_KEYS.CUSTOM_TEMPLATE_NAME);
    if (b64 && name === templateName) {
      syncCustomTemplateToRepo(b64, name, instructions);
    }
  } catch {
    // Ignore storage quota errors
  }
}

export function loadSavedMainHeaderConfigs(
  templateName: string
): MainHeaderFieldConfig[] | null {
  try {
    const raw = localStorage.getItem(`${STORAGE_KEYS.MAIN_FIELDS_PREFIX}${templateName}`);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length === 6) {
      return parsed;
    }
    return null;
  } catch {
    return null;
  }
}

export function saveMainHeaderConfigs(
  templateName: string,
  configs: MainHeaderFieldConfig[]
): void {
  try {
    localStorage.setItem(
      `${STORAGE_KEYS.MAIN_FIELDS_PREFIX}${templateName}`,
      JSON.stringify(configs)
    );
  } catch {
    // Ignore
  }
}

export function saveCustomDefaultTemplate(buffer: ArrayBuffer, fileName: string): void {
  try {
    const b64 = arrayBufferToBase64(buffer);
    localStorage.setItem(STORAGE_KEYS.CUSTOM_TEMPLATE_B64, b64);
    localStorage.setItem(STORAGE_KEYS.CUSTOM_TEMPLATE_NAME, fileName);
    const savedInst = loadSavedInstructionsForTemplate(fileName, []);
    syncCustomTemplateToRepo(b64, fileName, savedInst);
  } catch {
    // Ignore storage quota errors
  }
}

export function loadCustomDefaultTemplate(): { buffer: ArrayBuffer; fileName: string } | null {
  try {
    const b64 = localStorage.getItem(STORAGE_KEYS.CUSTOM_TEMPLATE_B64);
    const fileName = localStorage.getItem(STORAGE_KEYS.CUSTOM_TEMPLATE_NAME);
    if (b64 && fileName) {
      const savedInst = loadSavedInstructionsForTemplate(fileName, []);
      syncCustomTemplateToRepo(b64, fileName, savedInst);
      return {
        buffer: base64ToArrayBuffer(b64),
        fileName,
      };
    }
    // Fallback to bundled embeddedCustomTemplate.json (used when deployed to GitHub Pages!)
    if (
      embeddedCustomTemplate &&
      typeof embeddedCustomTemplate.base64 === 'string' &&
      embeddedCustomTemplate.base64.length > 0
    ) {
      return {
        buffer: base64ToArrayBuffer(embeddedCustomTemplate.base64),
        fileName: embeddedCustomTemplate.fileName || 'ჩაშენებული_შაბლონი.xlsx',
      };
    }
    return null;
  } catch {
    return null;
  }
}

export function clearCustomDefaultTemplate(templateName?: string): void {
  try {
    localStorage.removeItem(STORAGE_KEYS.CUSTOM_TEMPLATE_B64);
    localStorage.removeItem(STORAGE_KEYS.CUSTOM_TEMPLATE_NAME);
    if (templateName) {
      localStorage.removeItem(`${STORAGE_KEYS.INSTRUCTIONS_PREFIX}${templateName}`);
      localStorage.removeItem(`${STORAGE_KEYS.MAIN_FIELDS_PREFIX}${templateName}`);
    }
  } catch {
    // Ignore
  }
}
