/* global Office, document, fetch */

const API_BASE_URL = 'https://isign-api.tredas.com.tr';

Office.onReady(() => {
  document.getElementById('insertBtn').addEventListener('click', onInsertClick);
  renderDebugTrace();
});

/**
 * commands.js'in (otomatik LaunchEvent handler'ı) roamingSettings'e yazdığı izi okuyup
 * gösterir -- bu sayede "en son otomatik çalışma ne zaman, nereye kadar gitti" sorusuna
 * sunucu loglarına bakmadan, her zaman, herkes tarafından cevap verilebiliyor.
 * GECICI TESHIS amaçlı, sorun çözülünce kaldırılacak.
 */
function renderDebugTrace() {
  const el = document.getElementById('debugTrace');
  try {
    const raw = Office.context.roamingSettings.get('isignDebugTrace');
    if (!raw) {
      el.textContent = '(henüz hiç otomatik çalışma izi yok -- LaunchEvent hiç tetiklenmemiş demektir)';
      return;
    }
    const trace = JSON.parse(raw);
    el.textContent = trace.map((e) => `${e.t}  ${e.s}${e.d ? '  ' + e.d : ''}`).join('\n');
  } catch (err) {
    el.textContent = 'İz okunamadı: ' + (err && err.message ? err.message : String(err));
  }
}

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
