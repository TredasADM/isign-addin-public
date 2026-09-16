/* global Office, OfficeRuntime, fetch */

// Bu dosya YALNIZCA event tabanlı otomatik aktivasyon (LaunchEvent/OnNewMessageCompose) için --
// yeni/yanıtla/ilet yazma penceresi her açıldığında Outlook bunu kısıtlı, JS-only bir runtime'da
// çalıştırır (bkz. manifest.xml > Runtimes). Manuel "İmzamı Ekle" butonu bu dosyayı DEĞİL,
// normal bir sayfa olan taskpane.html/taskpane.js'i açıyor (ShowTaskpane).
//
// ÖNEMLİ -- OfficeRuntime.auth vs Office.auth: kısıtlı runtime'da Office.auth.getAccessToken
// YASAKLI, yalnızca OfficeRuntime.auth.getAccessToken destekleniyor.
// bkz. https://learn.microsoft.com/office/dev/add-ins/develop/autolaunch#unsupported-apis

const API_BASE_URL = 'https://isign-api.tredas.com.tr';

// GECICI TESHIS -- LaunchEvent'in gercekten tetiklenip tetiklenmedigini ve hangi asamada
// durdugunu gormek icin. Auth gerektirmeyen ayri bir uc noktaya sinyal gonderir, sonuc
// dogrulanip sorun cozulunce kaldirilacak.
function debugPing(stage, detail) {
  try {
    fetch(`${API_BASE_URL}/api/addin/debug-ping?stage=${encodeURIComponent(stage)}&detail=${encodeURIComponent(detail || '')}`, { keepalive: true }).catch(() => {});
  } catch (e) { /* yut */ }
}

// Script yuklendigi anda (handler cagrilmadan once) -- kisitli runtime'in bu dosyayi
// gercekten calistirip calistirmadigini gormek icin.
debugPing('script-loaded', typeof Office !== 'undefined' ? 'office-defined' : 'office-undefined');

async function getSignatureHtml() {
  debugPing('handler-start');
  let accessToken;
  try {
    accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });
    debugPing('token-ok');
  } catch (err) {
    debugPing('token-fail', (err && err.message) || String(err));
    throw err;
  }

  const response = await fetch(`${API_BASE_URL}/api/addin/my-signature`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  debugPing('fetch-status', String(response.status));
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
  debugPing('handler-invoked');
  getSignatureHtml()
    .then((html) => {
      debugPing('signature-received', html ? 'has-html' : 'no-html');
      if (html) return setSignature(html).then(() => debugPing('set-signature-ok'));
    })
    .catch((err) => {
      debugPing('handler-catch', (err && err.message) || String(err));
      console.error('iSign otomatik imza ekleme başarısız:', err);
    })
    .finally(() => {
      debugPing('handler-completed');
      event.completed();
    });
}

debugPing('before-associate', `actions-${typeof Office.actions}-associate-${typeof (Office.actions && Office.actions.associate)}`);
try {
  Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
  debugPing('associate-ok');
} catch (err) {
  debugPing('associate-fail', (err && err.message) || String(err));
}
