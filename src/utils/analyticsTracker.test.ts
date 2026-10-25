import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createAnalyticsEventRecord,
  detectVisitorDeviceAndBrowser,
  formatGeorgianTimestamp,
  summarizeAnalyticsEvents,
} from './analyticsTracker';

test('detectVisitorDeviceAndBrowser identifies Android, iOS (iPhone/iPad), and Desktop platforms', () => {
  const android = detectVisitorDeviceAndBrowser(
    'Mozilla/5.0 (Linux; Android 14; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36'
  );
  assert.equal(android.deviceCategory, 'Android');
  assert.equal(android.deviceLabel, 'Android');
  assert.equal(android.browser, 'Chrome');

  const iphone = detectVisitorDeviceAndBrowser(
    'Mozilla/5.0 (iPhone; CPU iPhone OS 17_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1'
  );
  assert.equal(iphone.deviceCategory, 'iOS');
  assert.equal(iphone.deviceLabel, 'iOS (iPhone)');
  assert.equal(iphone.browser, 'Safari');

  const ipadDesktopUA = detectVisitorDeviceAndBrowser(
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15',
    5
  );
  assert.equal(ipadDesktopUA.deviceCategory, 'iOS');
  assert.equal(ipadDesktopUA.deviceLabel, 'iOS (iPad)');

  const windowsPc = detectVisitorDeviceAndBrowser(
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36 Edg/125.0.0.0',
    0
  );
  assert.equal(windowsPc.deviceCategory, 'კომპიუტერი');
  assert.equal(windowsPc.deviceLabel, 'კომპიუტერი (Windows)');
  assert.equal(windowsPc.browser, 'Edge');
});

test('summarizeAnalyticsEvents aggregates visits and downloads by device category', () => {
  const events = [
    createAnalyticsEventRecord({
      eventType: 'visit',
      profileId: 'mypos',
      userAgent: 'Mozilla/5.0 (Linux; Android 14)',
    }),
    createAnalyticsEventRecord({
      eventType: 'visit',
      profileId: 'btrade',
      userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0)',
    }),
    createAnalyticsEventRecord({
      eventType: 'visit',
      profileId: 'mypos',
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    }),
    createAnalyticsEventRecord({
      eventType: 'download_pdf',
      profileId: 'btrade',
      fileName: 'შპს_ტესტი.pdf',
      userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0)',
    }),
    createAnalyticsEventRecord({
      eventType: 'download_excel',
      profileId: 'mypos',
      fileName: 'შპს_მაიპოს.xlsx',
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    }),
  ];

  const summary = summarizeAnalyticsEvents(events);
  assert.equal(summary.totalVisits, 3);
  assert.equal(summary.androidVisits, 1);
  assert.equal(summary.iosVisits, 1);
  assert.equal(summary.desktopVisits, 1);

  assert.equal(summary.totalDownloads, 2);
  assert.equal(summary.pdfDownloads, 1);
  assert.equal(summary.excelDownloads, 1);
  assert.equal(summary.iosDownloads, 1);
  assert.equal(summary.desktopDownloads, 1);
  assert.equal(summary.androidDownloads, 0);
});

test('formatGeorgianTimestamp returns DD.MM.YYYY HH:mm:ss format', () => {
  const formatted = formatGeorgianTimestamp(new Date('2026-10-25T10:30:45.000Z'));
  assert.match(formatted, /^\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2}$/);
});
