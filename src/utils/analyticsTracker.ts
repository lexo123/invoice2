import { InvoiceProfileId } from '../types/invoice';
import embeddedAnalyticsConfig from './embeddedAnalyticsConfig.json';

export type AnalyticsEventType = 'visit' | 'download_excel' | 'download_pdf' | 'test';
export type DeviceCategory = 'Android' | 'iOS' | 'კომპიუტერი';

export interface AnalyticsEventRecord {
  id: string;
  timestamp: string;
  isoTimestamp: string;
  eventType: AnalyticsEventType;
  actionLabel: string;
  profileId: InvoiceProfileId;
  profileLabel: string;
  deviceCategory: DeviceCategory;
  deviceLabel: string;
  browser: string;
  fileName?: string;
}

export interface AnalyticsSummary {
  totalVisits: number;
  totalDownloads: number;
  excelDownloads: number;
  pdfDownloads: number;
  androidVisits: number;
  iosVisits: number;
  desktopVisits: number;
  androidDownloads: number;
  iosDownloads: number;
  desktopDownloads: number;
}

const SHEETS_URL_STORAGE_KEY = 'invoexcel_google_sheets_url_v1';
const LOCAL_EVENTS_STORAGE_KEY = 'invoexcel_analytics_events_v1';
const SESSION_VISIT_LOGGED_KEY = 'invoexcel_visit_logged_session_v1';

export function detectVisitorDeviceAndBrowser(
  userAgent: string,
  maxTouchPoints = 0
): {
  deviceCategory: DeviceCategory;
  deviceLabel: string;
  browser: string;
} {
  const ua = userAgent || '';

  let deviceCategory: DeviceCategory = 'კომპიუტერი';
  let deviceLabel = 'კომპიუტერი';

  const isIOS =
    /iPhone|iPad|iPod/i.test(ua) ||
    (/Macintosh/i.test(ua) && maxTouchPoints > 1);
  const isAndroid = /Android/i.test(ua);

  if (isAndroid) {
    deviceCategory = 'Android';
    deviceLabel = 'Android';
  } else if (isIOS) {
    deviceCategory = 'iOS';
    deviceLabel = /iPad/i.test(ua) || maxTouchPoints > 1 ? 'iOS (iPad)' : 'iOS (iPhone)';
  } else if (/Windows/i.test(ua)) {
    deviceCategory = 'კომპიუტერი';
    deviceLabel = 'კომპიუტერი (Windows)';
  } else if (/Macintosh|Mac OS X/i.test(ua)) {
    deviceCategory = 'კომპიუტერი';
    deviceLabel = 'კომპიუტერი (Mac)';
  } else if (/Linux/i.test(ua)) {
    deviceCategory = 'კომპიუტერი';
    deviceLabel = 'კომპიუტერი (Linux)';
  }

  let browser = 'სხვა';
  if (/Edg\//i.test(ua)) {
    browser = 'Edge';
  } else if (/OPR\/|Opera/i.test(ua)) {
    browser = 'Opera';
  } else if (/Chrome\/|CriOS\//i.test(ua)) {
    browser = 'Chrome';
  } else if (/Firefox\/|FxiOS\//i.test(ua)) {
    browser = 'Firefox';
  } else if (/Safari\//i.test(ua)) {
    browser = 'Safari';
  }

  return {
    deviceCategory,
    deviceLabel,
    browser,
  };
}

export function formatGeorgianTimestamp(date: Date = new Date()): string {
  try {
    const parts = new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Asia/Tbilisi',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false,
    }).formatToParts(date);

    const map: Record<string, string> = {};
    for (const part of parts) {
      map[part.type] = part.value;
    }
    return `${map.day}.${map.month}.${map.year} ${map.hour}:${map.minute}:${map.second}`;
  } catch {
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${pad(date.getDate())}.${pad(date.getMonth() + 1)}.${date.getFullYear()} ${pad(
      date.getHours()
    )}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
  }
}

export function getActionLabel(eventType: AnalyticsEventType): string {
  switch (eventType) {
    case 'visit':
      return 'საიტზე შესვლა';
    case 'download_excel':
      return 'Excel (.xlsx) გადმოწერა';
    case 'download_pdf':
      return 'PDF (.pdf) გადმოწერა';
    case 'test':
      return 'ტესტ-ჩანაწერი (შემოწმება)';
  }
}

export function summarizeAnalyticsEvents(
  events: AnalyticsEventRecord[]
): AnalyticsSummary {
  const summary: AnalyticsSummary = {
    totalVisits: 0,
    totalDownloads: 0,
    excelDownloads: 0,
    pdfDownloads: 0,
    androidVisits: 0,
    iosVisits: 0,
    desktopVisits: 0,
    androidDownloads: 0,
    iosDownloads: 0,
    desktopDownloads: 0,
  };

  for (const ev of events) {
    const isVisit = ev.eventType === 'visit';
    const isExcel = ev.eventType === 'download_excel';
    const isPdf = ev.eventType === 'download_pdf';
    const isDownload = isExcel || isPdf;

    if (isVisit) {
      summary.totalVisits += 1;
      if (ev.deviceCategory === 'Android') summary.androidVisits += 1;
      else if (ev.deviceCategory === 'iOS') summary.iosVisits += 1;
      else summary.desktopVisits += 1;
    }

    if (isDownload) {
      summary.totalDownloads += 1;
      if (isExcel) summary.excelDownloads += 1;
      if (isPdf) summary.pdfDownloads += 1;

      if (ev.deviceCategory === 'Android') summary.androidDownloads += 1;
      else if (ev.deviceCategory === 'iOS') summary.iosDownloads += 1;
      else summary.desktopDownloads += 1;
    }
  }

  return summary;
}

const DEFAULT_GOOGLE_SHEETS_WEB_APP_URL =
  'https://script.google.com/macros/s/AKfycbwFfqjZ70bFeCUQD5imwpFRlhqjNbnvBc3tXpIOlEAiS3XT5xu2d4c20V5KphpxsC8H/exec';

export function loadGoogleSheetsWebhookUrl(): string {
  try {
    const localUrl = localStorage.getItem(SHEETS_URL_STORAGE_KEY);
    if (localUrl && localUrl.trim()) {
      return localUrl.trim();
    }
  } catch {
    // Ignore localStorage error
  }
  const embeddedUrl = (embeddedAnalyticsConfig as { googleSheetsWebAppUrl?: string })
    ?.googleSheetsWebAppUrl;
  return typeof embeddedUrl === 'string' && embeddedUrl.trim()
    ? embeddedUrl.trim()
    : DEFAULT_GOOGLE_SHEETS_WEB_APP_URL;
}

export function saveGoogleSheetsWebhookUrl(url: string): void {
  const clean = url.trim();
  try {
    if (clean) {
      localStorage.setItem(SHEETS_URL_STORAGE_KEY, clean);
    } else {
      localStorage.removeItem(SHEETS_URL_STORAGE_KEY);
    }
  } catch {
    // Ignore
  }

  if (typeof fetch !== 'undefined') {
    fetch('/__sync_embedded_template', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ googleSheetsWebAppUrl: clean }),
    }).catch(() => {
      // Ignore in static GitHub Pages environment
    });
  }
}

export function loadLocalAnalyticsEvents(): AnalyticsEventRecord[] {
  try {
    const raw = localStorage.getItem(LOCAL_EVENTS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function clearLocalAnalyticsEvents(): void {
  try {
    localStorage.removeItem(LOCAL_EVENTS_STORAGE_KEY);
  } catch {
    // Ignore
  }
}

function appendLocalAnalyticsEvent(event: AnalyticsEventRecord): void {
  try {
    const existing = loadLocalAnalyticsEvents();
    const next = [event, ...existing].slice(0, 150);
    localStorage.setItem(LOCAL_EVENTS_STORAGE_KEY, JSON.stringify(next));
  } catch {
    // Ignore
  }
}

export function createAnalyticsEventRecord(params: {
  eventType: AnalyticsEventType;
  profileId: InvoiceProfileId;
  fileName?: string;
  userAgent?: string;
  maxTouchPoints?: number;
  now?: Date;
}): AnalyticsEventRecord {
  const ua =
    params.userAgent ??
    (typeof navigator !== 'undefined' ? navigator.userAgent : '');
  const touchPoints =
    params.maxTouchPoints ??
    (typeof navigator !== 'undefined' ? navigator.maxTouchPoints || 0 : 0);
  const now = params.now ?? new Date();
  const { deviceCategory, deviceLabel, browser } = detectVisitorDeviceAndBrowser(
    ua,
    touchPoints
  );

  return {
    id: `ev-${now.getTime()}-${Math.random().toString(36).slice(2, 7)}`,
    timestamp: formatGeorgianTimestamp(now),
    isoTimestamp: now.toISOString(),
    eventType: params.eventType,
    actionLabel: getActionLabel(params.eventType),
    profileId: params.profileId,
    profileLabel: params.profileId === 'btrade' ? 'ბითრეიდ' : 'მაიპოს',
    deviceCategory,
    deviceLabel,
    browser,
    fileName: params.fileName || '',
  };
}

export async function sendEventToGoogleSheets(
  event: AnalyticsEventRecord,
  explicitWebhookUrl?: string
): Promise<boolean> {
  const webhookUrl = (explicitWebhookUrl ?? loadGoogleSheetsWebhookUrl()).trim();
  if (!webhookUrl || typeof fetch === 'undefined') {
    return false;
  }

  try {
    await fetch(webhookUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: JSON.stringify(event),
    });
    return true;
  } catch {
    return false;
  }
}

export async function trackSiteEvent(params: {
  eventType: AnalyticsEventType;
  profileId: InvoiceProfileId;
  fileName?: string;
}): Promise<AnalyticsEventRecord> {
  const event = createAnalyticsEventRecord(params);
  appendLocalAnalyticsEvent(event);
  await sendEventToGoogleSheets(event);
  return event;
}

export function trackInitialSessionVisit(profileId: InvoiceProfileId): void {
  try {
    if (typeof sessionStorage !== 'undefined') {
      const alreadyLogged = sessionStorage.getItem(SESSION_VISIT_LOGGED_KEY);
      if (alreadyLogged === '1') {
        return;
      }
      sessionStorage.setItem(SESSION_VISIT_LOGGED_KEY, '1');
    }
  } catch {
    // Ignore sessionStorage restrictions
  }
  void trackSiteEvent({
    eventType: 'visit',
    profileId,
  });
}

export async function fetchEventsFromGoogleSheets(
  explicitWebhookUrl?: string
): Promise<AnalyticsEventRecord[] | null> {
  const webhookUrl = (explicitWebhookUrl ?? loadGoogleSheetsWebhookUrl()).trim();
  if (!webhookUrl || typeof fetch === 'undefined') {
    return null;
  }

  try {
    const separator = webhookUrl.includes('?') ? '&' : '?';
    const response = await fetch(`${webhookUrl}${separator}action=read&t=${Date.now()}`, {
      method: 'GET',
      redirect: 'follow',
    });
    if (!response.ok) return null;
    const data = await response.json();
    if (data && Array.isArray(data.events)) {
      return data.events as AnalyticsEventRecord[];
    }
    return null;
  } catch {
    return null;
  }
}

export const GOOGLE_APPS_SCRIPT_CODE = `function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "თარიღი და დრო",
        "ქმედება",
        "გვერდი",
        "მოწყობილობა",
        "სისტემა / დეტალები",
        "ბრაუზერი",
        "ფაილის სახელი"
      ]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e0f2fe");
      sheet.setFrozenRows(1);
    }

    var data = JSON.parse(e.postData.contents);
    sheet.appendRow([
      data.timestamp || new Date().toLocaleString("ka-GE", { timeZone: "Asia/Tbilisi" }),
      data.actionLabel || "",
      data.profileLabel || "",
      data.deviceCategory || "",
      data.deviceLabel || "",
      data.browser || "",
      data.fileName || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, events: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var events = [];
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var actionLabel = String(row[1] || "");
      var eventType = "visit";
      if (actionLabel.indexOf("Excel") !== -1) eventType = "download_excel";
      else if (actionLabel.indexOf("PDF") !== -1) eventType = "download_pdf";
      else if (actionLabel.indexOf("ტესტ") !== -1) eventType = "test";

      var devCat = String(row[3] || "კომპიუტერი");
      events.push({
        id: "row-" + (i + 2),
        timestamp: String(row[0] || ""),
        isoTimestamp: "",
        eventType: eventType,
        actionLabel: actionLabel,
        profileId: String(row[2] || "").indexOf("ბითრეიდ") !== -1 ? "btrade" : "mypos",
        profileLabel: String(row[2] || ""),
        deviceCategory: (devCat === "Android" || devCat === "iOS") ? devCat : "კომპიუტერი",
        deviceLabel: String(row[4] || devCat),
        browser: String(row[5] || ""),
        fileName: String(row[6] || "")
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, events: events }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;
