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

async function getSignatureHtml() {
  const accessToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });

  const response = await fetch(`${API_BASE_URL}/api/addin/my-signature`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
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
  getSignatureHtml()
    .then((html) => {
      if (html) return setSignature(html);
    })
    .catch((err) => {
      console.error('iSign otomatik imza ekleme başarısız:', err);
    })
    .finally(() => event.completed());
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
