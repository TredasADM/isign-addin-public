/* global Office, OfficeRuntime, fetch */

// Bu dosya YALNIZCA event tabanlı otomatik aktivasyon (LaunchEvent/OnNewMessageCompose) için --
// yeni/yanıtla/ilet yazma penceresi her açıldığında Outlook bunu kısıtlı bir runtime'da
// çalıştırır (bkz. manifest > Runtimes). Manuel "İmzamı Ekle" butonu bu dosyayı DEĞİL,
// normal bir sayfa olan taskpane.html/taskpane.js'i açıyor (ShowTaskpane / openPage).
//
// ÖNEMLİ -- OfficeRuntime.auth vs Office.auth: kısıtlı runtime'da Office.auth.getAccessToken
// YASAKLI, yalnızca OfficeRuntime.auth.getAccessToken destekleniyor.
// bkz. https://learn.microsoft.com/office/dev/add-ins/develop/autolaunch#unsupported-apis

const API_BASE_URL = 'https://isign-api.tredas.com.tr';

// GECICI TESHIS -- iki paralel iz bırakma yöntemi:
//  1) debugPing: canlı izleme için sunucuya fetch ile sinyal -- ama bu fetch'in KENDİSİ
//     o an ağ/CORS/zamanlama sorunu yaşarsa SESSİZCE kaybolur, hiç iz kalmaz.
//  2) traceLog: Office.context.roamingSettings'e (mailbox'a bağlı, ağdan bağımsız, kalıcı
//     yerel depolama) yazar -- taskpane.js açıldığında bu izi okuyup ekranda gösterir, böylece
//     "en son otomatik çalışma ne zaman, nereye kadar gitti" sorusuna ağ/log'a bakmadan,
//     her zaman cevap verebiliyoruz. İkisi birbirinin yedeği.
function debugPing(stage, detail) {
  try {
    fetch(`${API_BASE_URL}/api/addin/debug-ping?stage=${encodeURIComponent(stage)}&detail=${encodeURIComponent(detail || '')}`, { keepalive: true }).catch(() => {});
  } catch (e) { /* yut */ }
}

function traceLog(stage, detail) {
  debugPing(stage, detail);
  try {
    const settings = Office && Office.context && Office.context.roamingSettings;
    if (!settings) return;
    let trace = [];
    try {
      const raw = settings.get('isignDebugTrace');
      if (raw) trace = JSON.parse(raw);
    } catch (e) { trace = []; }
    trace.push({ t: new Date().toISOString(), s: stage, d: String(detail || '').slice(0, 200) });
    if (trace.length > 30) trace = trace.slice(trace.length - 30);
    settings.set('isignDebugTrace', JSON.stringify(trace));
    settings.saveAsync(() => { /* sonucu onemli degil, best-effort */ });
  } catch (e) {
    // roamingSettings bu baglamda kullanilamiyorsa sessizce gec -- debugPing zaten denendi.
  }
}

// Global, beklenmedik hatalari da yakalayip iz birakmak icin -- normalde try/catch'lerin
// disinda kalabilecek (ornegin script parse/yukleme asamasindaki) hatalari da gormek icin.
try {
  self.addEventListener('error', (e) => {
    traceLog('window-error', (e && e.message) || 'bilinmeyen hata');
  });
  self.addEventListener('unhandledrejection', (e) => {
    traceLog('unhandled-rejection', (e && e.reason && e.reason.message) || String(e && e.reason));
  });
} catch (e) { /* bu ortamda self/addEventListener olmayabilir, sessizce gec */ }

// Script yuklendigi anda (handler cagrilmadan once) -- kisitli runtime'in bu dosyayi
// gercekten calistirip calistirmadigini gormek icin.
traceLog('script-loaded', typeof Office !== 'undefined' ? 'office-defined' : 'office-undefined');

async function getSignatureHtml() {
  traceLog('handler-start');
  let accessToken;
  try {
    accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });
    traceLog('token-ok');
  } catch (err) {
    traceLog('token-fail', (err && err.message) || String(err));
    throw err;
  }

  const response = await fetch(`${API_BASE_URL}/api/addin/my-signature`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  traceLog('fetch-status', String(response.status));
  if (!response.ok) {
    throw new Error(`iSign API ${response.status}`);
  }
  const data = await response.json();
  if (!data.found) return null;
  return data.html;
}

function setSignature(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(result.error);
      }
    );
  });
}

/** Otomatik: yazma penceresi (yeni/yanıtla/ilet) açıldığında çalışır. */
function onNewMessageComposeHandler(event) {
  // Ust seviye senkron try/catch -- .then/.catch zinciri sadece ASYNC hatalari yakalar;
  // buraya girer girmez (getSignatureHtml() cagrisindan once) senkron bir hata olursa
  // (ornegin beklenmedik bir ReferenceError) bu olmadan hic yakalanmazdi.
  try {
    traceLog('handler-invoked');
    getSignatureHtml()
      .then((html) => {
        traceLog('signature-received', html ? 'has-html' : 'no-html');
        if (html) return setSignature(html).then(() => traceLog('set-signature-ok'));
      })
      .catch((err) => {
        traceLog('handler-catch', (err && err.message) || String(err));
        console.error('iSign otomatik imza ekleme başarısız:', err);
      })
      .finally(() => {
        traceLog('handler-completed');
        event.completed();
      });
  } catch (err) {
    traceLog('handler-sync-throw', (err && err.message) || String(err));
    event.completed();
  }
}

traceLog('before-associate', `actions-${typeof Office.actions}-associate-${typeof (Office.actions && Office.actions.associate)}`);
try {
  Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
  traceLog('associate-ok');
} catch (err) {
  traceLog('associate-fail', (err && err.message) || String(err));
}
