/**
 * L-TEX SMM Звіт — бекенд синхронізації з Google-таблицею.
 * ─────────────────────────────────────────────────────────────
 * Працює як проста хмарна «база даних»: зберігає 4 ключі
 * (щоденні звіти, вірусні відео, ROAS-кампанії, цілі KPI)
 * у вигляді рядків key | value(JSON) | updated(мс).
 *
 * НАЛАШТУВАННЯ (детально — у CLOUD_SETUP.md):
 *  1. Створіть Google-таблицю.
 *  2. Розширення → Apps Script.
 *  3. Вставте цей код, замініть SECRET на свій.
 *  4. Розгорнути → Новий розгортання → тип «Веб-застосунок».
 *     Виконувати як: Я.  Хто має доступ: Усі.
 *  5. Скопіюйте URL (…/exec) і вставте у звіті: кнопка ☁️ → URL + той самий ключ.
 */

const SHEET_NAME = 'LTEX_STORE';
const SECRET     = 'ltex-2026'; // ⚠️ ЗАМІНІТЬ на свій ключ і впишіть такий самий у звіті

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'value', 'updated']);
  }
  return sh;
}

function out_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ЧИТАННЯ: повертає всі ключі ──
function doGet(e) {
  try {
    if (SECRET && (e.parameter.k || '') !== SECRET) {
      return out_({ ok: false, error: 'unauthorized' });
    }
    const rows = getSheet_().getDataRange().getValues();
    const data = {}, meta = {};
    for (let i = 1; i < rows.length; i++) {
      const k = rows[i][0];
      if (!k) continue;
      data[k] = rows[i][1];
      meta[k] = Number(rows[i][2]) || 0;
    }
    return out_({ ok: true, data: data, meta: meta });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// ── ЗАПИС: upsert ключів ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (SECRET && (body.secret || '') !== SECRET) {
      return out_({ ok: false, error: 'unauthorized' });
    }
    const sh = getSheet_();
    const rows = sh.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) rowByKey[rows[i][0]] = i + 1; // номер рядка в аркуші
    }
    const data = body.data || {};
    const meta = body.meta || {};
    Object.keys(data).forEach(function (k) {
      const v = typeof data[k] === 'string' ? data[k] : JSON.stringify(data[k]);
      const ts = Number(meta[k]) || Date.now();
      if (rowByKey[k]) {
        sh.getRange(rowByKey[k], 2, 1, 2).setValues([[v, ts]]);
      } else {
        sh.appendRow([k, v, ts]);
      }
    });
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
