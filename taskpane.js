/* global Office, document, fetch */

const API_BASE_URL = 'https://isign-api.tredas.com.tr';

Office.onReady(() => {
  document.getElementById('insertBtn').addEventListener('click', onInsertClick);
});

function setStatus(text) {
  document.getElementById('status').textContent = text;
}

async function onInsertClick() {
  const btn = document.getElementById('insertBtn');
  btn.disabled = true;
  setStatus('İmza alınıyor...');
  try {
    const accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });
    const response = await fetch(`${API_BASE_URL}/api/addin/my-signature`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(`API ${response.status}: ${body}`);
    }
    const data = await response.json();
    if (!data.found) {
      setStatus('Bu hesap için tanımlı bir imza bulunamadı (dizin senkronizasyonu veya atama kuralı kontrol edilmeli).');
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
    setStatus('İmza eklendi.');
  } catch (err) {
    setStatus('Hata: ' + (err && err.message ? err.message : String(err)));
  } finally {
    btn.disabled = false;
  }
}
