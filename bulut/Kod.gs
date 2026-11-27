/**
 * Koray - Rafta · Google E-Tablo eşitleme kodu
 * Bu kodu E-Tablo > Uzantılar > Apps Script içine yapıştırın,
 * Dağıt > Yeni dağıtım > Web uygulaması olarak yayınlayın.
 * "Urunler" sayfasını Excel gibi düzenleyebilirsiniz; telefonlar değişikliği alır.
 */
const SH_U = 'Urunler', SH_A = 'Ayar', SH_S = 'Sorular', SH_C = 'Cihazlar';
const SCOLS = ['id', 'urunId', 'urun', 'yer', 'soran', 'zaman', 'durum', 'adet', 'cevaplayan', 'cevapZaman', 'kapali', 'cihaz'];
const CCOLS = ['cihaz', 'token', 'tip', 'ad', 'zaman'];
const FB_PROJE = 'koray-rafta', UYG_ADRES = 'https://mohbai-code.github.io/depo-bul/';
const COLS = ['id', 'ad', 'kod', 'raf', 'bolme', 'kat', 'taraf', 'adet', 'guncelleyen', 'guncelleme'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sayfa_(ad, baslik) {
  let s = ss_().getSheetByName(ad);
  if (!s) { s = ss_().insertSheet(ad); if (baslik) s.appendRow(baslik); s.setFrozenRows(1); }
  else if (baslik && s.getLastColumn() < baslik.length) s.getRange(1, 1, 1, baslik.length).setValues([baslik]);
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
  const s = sayfa_(SH_U, COLS), v = s.getDataRange().getValues(), out = [], gorulen = {};
  for (let i = 1; i < v.length; i++) {
    const r = v[i]; if (!String(r[1] || '').trim()) continue;
    let id = String(r[0] || '').trim();
    // id boşsa ya da tekrar ediyorsa kalıcı bir kimlik yaz
    if (!id || gorulen[id]) { id = 'g' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); s.getRange(i + 1, 1).setValue(id); }
    gorulen[id] = 1;
    out.push({
      id: id, ad: String(r[1]).trim(), kod: String(r[2] || ''), raf: String(r[3] || '').trim(),
      bolme: sayi_(r[4], 1), kat: sayi_(r[5], 1), taraf: String(r[6]) === '2' ? 2 : '', adet: r[7] === '' ? '' : sayi_(r[7], '')
    });
  }
  return out;
}
function sorular_() {
  const v = sayfa_(SH_S, SCOLS).getDataRange().getValues(), out = [], sinir = Date.now() - 12 * 3600 * 1000;
  for (let i = v.length - 1; i >= 1 && out.length < 40; i--) {
    const r = v[i], z = r[5] instanceof Date ? r[5].getTime() : Date.parse(r[5]);
    if (!r[0] || r[10] === true || String(r[10]) === 'evet' || !(z > sinir)) continue;
    out.push({ id: String(r[0]), urunId: String(r[1]), urun: String(r[2]), yer: String(r[3]), soran: String(r[4]), zaman: z,
      durum: String(r[6] || 'bekliyor'), adet: r[7] === '' ? '' : r[7], cevaplayan: String(r[8] || ''), cevapZaman: r[9] instanceof Date ? r[9].getTime() : (r[9] ? Date.parse(r[9]) : 0) });
  }
  return out;
}
function durum_(v) {
  const a = ayarOku_(); let plan = null, users = [];
  try { plan = a.plan ? JSON.parse(a.plan) : null; } catch (e) {}
  try { users = a.kullanicilar ? JSON.parse(a.kullanicilar) : []; } catch (e) {}
  const govde = { plan: plan, prods: urunler_(), users: users };
  const ver = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(govde), Utilities.Charset.UTF_8));
  const sorular = sorular_(), simdi = Date.now();
  if (v && v === ver) return { ok: true, ver: ver, ayni: true, sorular: sorular, simdi: simdi };
  return Object.assign({ ok: true, ver: ver, sorular: sorular, simdi: simdi }, govde);
}
function cevap_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  // Okuma: telefonlar birbirini beklemesin diye kilitsiz
  return cevap_(durum_(e && e.parameter && e.parameter.v));
}
const SERBEST = { kayit: true,  soru: 1, cevap: 1, kapat: 1 };
function soruIslem_(op) {
  const s = sayfa_(SH_S, SCOLS), simdi = new Date();
  if (op.t === 'soru' && op.s && op.s.id) {
    const cid = String(op.s.cid || '').slice(0, 40);
    s.appendRow([String(op.s.id), String(op.s.urunId || ''), String(op.s.urun || '').slice(0, 80), String(op.s.yer || '').slice(0, 80), String(op.s.soran || '').slice(0, 30), simdi, 'bekliyor', '', '', '', '', cid]);
    bildirimKuyruk_.push({ haric: cid, baslik: 'Var mı? ' + String(op.s.urun || '').slice(0, 60), metin: String(op.s.soran || 'Tezgâh') + ' soruyor' + (op.s.yer ? ' · ' + String(op.s.yer).slice(0, 60) : ''), etiket: String(op.s.id), tur: 'soru' });
    return;
  }
  if (op.t === 'kayit') { cihazKaydet_(op); return; }
  const v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(op.id)) {
    if (op.t === 'cevap' && String(v[i][6] || 'bekliyor') === 'bekliyor') {
      const adet = op.adet === '' || op.adet == null ? '' : Number(op.adet) || '', varMi = op.durum === 'var';
      s.getRange(i + 1, 7, 1, 4).setValues([[varMi ? 'var' : 'yok', adet, String(op.kim || '').slice(0, 30), simdi]]);
      if (v[i][11]) bildirimKuyruk_.push({ sadece: String(v[i][11]), baslik: (varMi ? '✅ Var' + (adet !== '' ? ' · ' + adet + ' adet' : '') : '❌ Yok') + ' — ' + String(v[i][2]).slice(0, 60), metin: String(op.kim || 'Depo').slice(0, 30) + ' cevapladı', etiket: String(op.id), tur: 'cevap' });
    }
    if (op.t === 'kapat') s.getRange(i + 1, 11).setValue('evet');
    return;
  }
}

function doPost(e) {
  const kilit = LockService.getScriptLock(); kilit.waitLock(20000);
  try {
    const q = JSON.parse(e.postData.contents || '{}'), a = ayarOku_();
    let users = []; try { users = a.kullanicilar ? JSON.parse(a.kullanicilar) : []; } catch (x) {}
    const yetkili = users.length === 0 || users.some(u => u.ad === q.kullanici && u.h === q.h);
    const sadeceSoru = (q.ops || []).every(o => SERBEST[o.t]);
    if (!yetkili && !sadeceSoru) return cevap_({ ok: false, hata: 'yetki' });
    const s = sayfa_(SH_U, COLS), zaman = new Date(), kim = String(q.kullanici || '');
    const satir = p => COLS.map(c => c === 'guncelleyen' ? kim : c === 'guncelleme' ? zaman : (p[c] === undefined || p[c] === null ? '' : p[c]));
    for (const op of (q.ops || [])) {
      if (SERBEST[op.t]) { soruIslem_(op); continue; }
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
  } finally { kilit.releaseLock(); try { bildirimGonder_(); } catch (x) { console.error(x); } }
}

/* ---------- Bildirimler (Firebase Cloud Messaging) ----------
 * Gizli anahtar: Proje Ayarları (⚙️) > Komut dosyası özellikleri > FCM_ANAHTAR
 * değerine Firebase'den indirilen hizmet hesabı .json dosyasının TAMAMI yapıştırılır. */
var bildirimKuyruk_ = [];

function cihazKaydet_(op) {
  const cid = String(op.cid || '').slice(0, 40), token = String(op.token || '');
  if (!cid || token.length < 20) return;
  const s = sayfa_(SH_C, CCOLS), v = s.getDataRange().getValues(), satir = [cid, token, String(op.tip || '').slice(0, 10), String(op.ad || '').slice(0, 30), new Date()];
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][1]) === token && String(v[i][0]) !== cid) s.deleteRow(i + 1);
  const v2 = s.getDataRange().getValues();
  for (let i = 1; i < v2.length; i++) if (String(v2[i][0]) === cid) { s.getRange(i + 1, 1, 1, CCOLS.length).setValues([satir]); return; }
  s.appendRow(satir);
}

function erisimAnahtari_() {
  const c = CacheService.getScriptCache(), eski = c.get('fcm_erisim'); if (eski) return eski;
  const ham = PropertiesService.getScriptProperties().getProperty('FCM_ANAHTAR');
  if (!ham) throw new Error('FCM_ANAHTAR komut dosyası özelliği eksik');
  let k; try { k = JSON.parse(ham); } catch (x) { throw new Error('FCM_ANAHTAR içeriği okunamadı. .json dosyasının TAMAMINI ({ ile başlayıp } ile biten) yapıştırın.'); }
  if (!k.client_email || !k.private_key) throw new Error('FCM_ANAHTAR içinde client_email / private_key yok. Firebase > Hizmet hesapları > Yeni özel anahtar dosyasını kullanın.');
  // Yapıştırırken bozulan satır sonlarını onar: anahtarı baştan düzgün PEM biçiminde kur
  const govdeAnahtar = String(k.private_key).replace(/\\n/g, '\n').replace(/-----(BEGIN|END) PRIVATE KEY-----/g, '').replace(/[^A-Za-z0-9+\/=]/g, '');
  k.private_key = '-----BEGIN PRIVATE KEY-----\n' + govdeAnahtar.match(/.{1,64}/g).join('\n') + '\n-----END PRIVATE KEY-----\n';
  const simdi = Math.floor(Date.now() / 1000);
  const b64 = o => Utilities.base64EncodeWebSafe(JSON.stringify(o)).replace(/=+$/, '');
  const govde = b64({ alg: 'RS256', typ: 'JWT' }) + '.' + b64({ iss: k.client_email, scope: 'https://www.googleapis.com/auth/firebase.messaging', aud: 'https://oauth2.googleapis.com/token', iat: simdi, exp: simdi + 3600 });
  const imza = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(govde, k.private_key)).replace(/=+$/, '');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: govde + '.' + imza }, muteHttpExceptions: true });
  const d = JSON.parse(r.getContentText()); if (!d.access_token) throw new Error('Firebase anahtarı alınamadı: ' + r.getContentText());
  c.put('fcm_erisim', d.access_token, 3000); return d.access_token;
}

function bildirimGonder_() {
  const is = bildirimKuyruk_; bildirimKuyruk_ = []; if (!is.length) return;
  if (!PropertiesService.getScriptProperties().getProperty('FCM_ANAHTAR')) return;
  const s = sayfa_(SH_C, CCOLS), v = s.getDataRange().getValues().slice(1).filter(r => r[0] && r[1]);
  const istekler = [], hedefler = [], erisim = erisimAnahtari_();
  for (const b of is) for (const r of v) {
    const cid = String(r[0]); if (b.haric && cid === b.haric) continue; if (b.sadece && cid !== b.sadece) continue;
    const mesaj = { token: String(r[1]), notification: { title: b.baslik, body: b.metin }, data: { tur: b.tur, id: b.etiket },
      android: { priority: 'high', notification: { channel_id: 'sorular', icon: 'ic_stat_rafta', color: '#058a8e', sound: 'default', tag: b.etiket } },
      apns: { headers: { 'apns-priority': '10' }, payload: { aps: { sound: 'default' } } },
      webpush: { headers: { Urgency: 'high', TTL: '3600' }, notification: { icon: UYG_ADRES + 'icon-192.png', badge: UYG_ADRES + 'icon-192.png', tag: b.etiket, renotify: true }, fcm_options: { link: UYG_ADRES } } };
    istekler.push({ url: 'https://fcm.googleapis.com/v1/projects/' + FB_PROJE + '/messages:send', method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + erisim }, payload: JSON.stringify({ message: mesaj }), muteHttpExceptions: true });
    hedefler.push(String(r[1]));
  }
  if (!istekler.length) return;
  const sonuclar = UrlFetchApp.fetchAll(istekler), olu = {};
  sonuclar.forEach((r, i) => { const k = r.getResponseCode(); if (k === 404 || (k === 400 && /UNREGISTERED|registration token/i.test(r.getContentText()))) olu[hedefler[i]] = 1; else if (k >= 300) console.warn(k, r.getContentText()); });
  if (Object.keys(olu).length) { const t = s.getDataRange().getValues(); for (let i = t.length - 1; i >= 1; i--) if (olu[String(t[i][1])]) s.deleteRow(i + 1); }
}

/** Kurulumdan sonra bir kez çalıştırın: izin ister ve kayıtlı tüm cihazlara deneme bildirimi yollar. */
function bildirimDene() {
  sayfa_(SH_C, CCOLS); UrlFetchApp.fetch('https://oauth2.googleapis.com/', { muteHttpExceptions: true });
  CacheService.getScriptCache().remove('fcm_erisim'); erisimAnahtari_();
  bildirimKuyruk_.push({ baslik: 'Koray - Rafta', metin: 'Bildirimler çalışıyor ✓', etiket: 'deneme', tur: 'deneme' });
  bildirimGonder_(); console.log('Tamam. Kayıtlı cihaz sayısı: ' + Math.max(0, sayfa_(SH_C, CCOLS).getLastRow() - 1));
}
