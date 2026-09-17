/* global Office, document, fetch */

const API_BASE_URL = 'https://isign-api.tredas.com.tr';

Office.onReady(() => {
  document.getElementById('insertBtn').addEventListener('click', () => insertSignature({ manual: true }));

  // Sabitlenebilir (pinnable) görev bölmesi -- LaunchEvent'ten TAMAMEN BAĞIMSIZ bir otomatik
  // ekleme yolu. Kullanıcı bu paneli bir kez sabitlerse (pin), her yeni yazma penceresinde
  // panel otomatik açık kalır ve Office.EventType.ItemChanged olayı tetiklenir -- biz de o
  // olayda imzayı otomatik ekliyoruz. LaunchEvent (commands.js) hâlâ paralel olarak
  // deneniyor, ama bu, ondan bağımsız, çok daha eski/olgun bir mekanizma (requirement set 1.5).
  // bkz. https://learn.microsoft.com/office/dev/add-ins/outlook/pinnable-taskpane
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      insertSignature({ manual: false });
    });
  } catch (e) {
    // addHandlerAsync bu ortamda yoksa (cok eski istemci) sessizce gec, manuel buton kalir.
  }

  // Panel zaten sabitlenmiş olarak (ilk kez degil, onceden pinlenmis) acilmis olabilir --
  // bu durumda ItemChanged bu ilk oge icin hic tetiklenmeyebilir, o yuzden yuklendiginde
  // de bir kez deniyoruz.
  insertSignature({ manual: false, silent: true });
});

function setStatus(text) {
  document.getElementById('status').textContent = text;
}

/**
 * Hem manuel buton hem otomatik (ItemChanged/ilk yükleme) yol tarafından kullanılan tek
 * ortak imza ekleme fonksiyonu. `silent`, ilk otomatik denemede kullanıcıya gereksiz
 * "imza bulunamadı" gibi mesajlar göstermemek için (örn. panel henüz sabitlenmemişken de
 * çalışıyor olabilir, bu durumda sessiz kalmak daha doğru).
 */
async function insertSignature({ manual, silent }) {
  const btn = document.getElementById('insertBtn');
  if (manual) btn.disabled = true;
  if (!silent) setStatus('İmza alınıyor...');
  try {
    const accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: manual, allowConsentPrompt: manual });
    const response = await fetch(`${API_BASE_URL}/api/addin/my-signature`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(`API ${response.status}: ${body}`);
    }
    const data = await response.json();
    if (!data.found) {
      if (!silent) setStatus('Bu hesap için tanımlı bir imza bulunamadı (dizin senkronizasyonu veya atama kuralı kontrol edilmeli).');
      return;
    }
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.setSignatureAsync(
        data.html,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(result.error);
        }
      );
    });
    if (!silent) setStatus('İmza eklendi.');
  } catch (err) {
    if (!silent) setStatus('Hata: ' + (err && err.message ? err.message : String(err)));
  } finally {
    if (manual) btn.disabled = false;
  }
}
