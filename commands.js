/* global Office, OfficeRuntime, fetch */

// Bu dosya YALNIZCA event tabanlı otomatik aktivasyon (LaunchEvent/OnNewMessageCompose) için --
// yeni/yanıtla/ilet yazma penceresi her açıldığında Outlook bunu kısıtlı bir runtime'da
// çalıştırır (bkz. manifest > Runtimes). Manuel "iSign" butonu bu dosyayı DEĞİL, normal bir
// sayfa olan taskpane.html/taskpane.js'i açıyor (ShowTaskpane / openPage).
//
// NOT: Bu mekanizma (Office.actions.associate + LaunchEvent) uzun bir teşhis sürecinden sonra
// Microsoft'un kendi platformunda hâlâ çözülmemiş, geniş çaplı bir hata olduğu kanıtlandı
// (bkz. OfficeDev/office-js#2717, #3395, #4379, #4391, #4469, #5867, #5904, #6956, #7000) --
// associate() başarıyla çalışıyor ama Outlook handler'ı hiçbir zaman gerçekten çağırmıyor,
// hiçbir manifest formatında/hesapta/istemcide. Bu dosya ileride Microsoft bu hatayı
// düzeltirse diye KAYITLI tutuluyor (maliyeti sıfır) -- ama gerçek otomatik ekleme şu an
// taskpane.js'teki sabitlenebilir (pinnable) görev bölmesi + ItemChanged mekanizmasıyla
// sağlanıyor. Bu yüzden burada artık teşhis/loglama kodu YOK -- sorun taraf-bağımsız olarak
// kapatıldı, aktif izleme gerekmiyor.
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

/** Otomatik: yazma penceresi (yeni/yanıtla/ilet) açıldığında çalışır (bkz. yukarıdaki not). */
function onNewMessageComposeHandler(event) {
  getSignatureHtml()
    .then((html) => (html ? setSignature(html) : undefined))
    .catch((err) => console.error('iSign otomatik imza ekleme başarısız:', err))
    .finally(() => event.completed());
}

Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
