/**
 * Koray - Rafta · Google E-Tablo eşitleme kodu
 * Bu kodu E-Tablo > Uzantılar > Apps Script içine yapıştırın,
 * Dağıt > Yeni dağıtım > Web uygulaması olarak yayınlayın.
 * "Urunler" sayfasını Excel gibi düzenleyebilirsiniz; telefonlar değişikliği alır.
 */
const SH_U = 'Urunler', SH_A = 'Ayar';
const COLS = ['id', 'ad', 'kod', 'raf', 'bolme', 'kat', 'taraf', 'adet', 'guncelleyen', 'guncelleme'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sayfa_(ad, baslik) {
  let s = ss_().getSheetByName(ad);
  if (!s) { s = ss_().insertSheet(ad); if (baslik) s.appendRow(baslik); s.setFrozenRows(1); }
  return s;
}
function ayarOku_() {
  const v = sayfa_(SH_A, ['anahtar', 'deger']).getDataRange().getValues(), o = {};
  for (let i = 1; i < v.length; i++) o[v[i][0]] = v[i][1];
  return o;
}
function ayarYaz_(k, val) {
  const s = sayfa_(SH_A, ['anahtar', 'deger']), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === k) { s.getRange(i + 1, 2).setValue(val); return; }
  s.appendRow([k, val]);
}
function sayi_(x, d) { const n = Number(x); return isFinite(n) && x !== '' ? n : d; }
function urunler_() {
  const v = sayfa_(SH_U, COLS).getDataRange().getValues(), out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!String(r[1] || '').trim()) continue;
    out.push({
      id: String(r[0] || ('g' + i)), ad: String(r[1]).trim(), kod: String(r[2] || ''), raf: String(r[3] || '').trim(),
      bolme: sayi_(r[4], 1), kat: sayi_(r[5], 1), taraf: String(r[6]) === '2' ? 2 : '', adet: r[7] === '' ? '' : sayi_(r[7], '')
    });
  }
  return out;
}
function durum_() {
  const a = ayarOku_(); let plan = null, users = [];
  try { plan = a.plan ? JSON.parse(a.plan) : null; } catch (e) {}
  try { users = a.kullanicilar ? JSON.parse(a.kullanicilar) : []; } catch (e) {}
  const govde = { plan: plan, prods: urunler_(), users: users };
  const ver = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(govde), Utilities.Charset.UTF_8));
  return Object.assign({ ok: true, ver: ver }, govde);
}
function cevap_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() { return cevap_(durum_()); }

function doPost(e) {
  const kilit = LockService.getScriptLock(); kilit.waitLock(20000);
  try {
    const q = JSON.parse(e.postData.contents || '{}'), a = ayarOku_();
    let users = []; try { users = a.kullanicilar ? JSON.parse(a.kullanicilar) : []; } catch (x) {}
    const yetkili = users.length === 0 || users.some(u => u.ad === q.kullanici && u.h === q.h);
    if (!yetkili) return cevap_({ ok: false, hata: 'yetki' });
    const s = sayfa_(SH_U, COLS), zaman = new Date(), kim = String(q.kullanici || '');
    const satir = p => COLS.map(c => c === 'guncelleyen' ? kim : c === 'guncelleme' ? zaman : (p[c] === undefined || p[c] === null ? '' : p[c]));
    for (const op of (q.ops || [])) {
      if (op.t === 'users' && Array.isArray(op.users)) ayarYaz_('kullanicilar', JSON.stringify(op.users.slice(0, 3)));
      else if (op.t === 'plan' && op.plan) ayarYaz_('plan', JSON.stringify(op.plan));
      else if (op.t === 'hepsi') {
        s.clearContents(); s.getRange(1, 1, 1, COLS.length).setValues([COLS]);
        const rows = (op.prods || []).map(satir);
        if (rows.length) s.getRange(2, 1, rows.length, COLS.length).setValues(rows);
      } else if (op.t === 'urun' && op.p && op.p.id) {
        const v = s.getDataRange().getValues(); let bulundu = false;
        for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(op.p.id)) { s.getRange(i + 1, 1, 1, COLS.length).setValues([satir(op.p)]); bulundu = true; break; }
        if (!bulundu) s.appendRow(satir(op.p));
      } else if (op.t === 'sil' && op.id) {
        const v = s.getDataRange().getValues();
        for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(op.id)) s.deleteRow(i + 1);
      }
    }
    return cevap_(durum_());
  } finally { kilit.releaseLock(); }
}
