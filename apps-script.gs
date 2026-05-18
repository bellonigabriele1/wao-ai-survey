/* ============================================================================
   WAO AI Survey — Google Apps Script backend (hardened)
   ----------------------------------------------------------------------------
   COME INSTALLARLO (primo setup, ~5 min)
   1. Sul tuo Google Sheet (account personale): Estensioni → Apps Script.
   2. Cancella il codice di default e incolla TUTTO questo file dentro Code.gs.
   3. Mostra il manifest: clicca sull'icona "Impostazioni progetto" (ingranaggio
      a sinistra) → spunta "Mostra il file manifest 'appsscript.json' nell'editor".
   4. Apri il file `appsscript.json` apparso a sinistra e incolla esattamente
      il contenuto del file `appsscript.json` di questo repo.
   5. Salva il progetto (dischetto), assegna un nome (es. "wao-survey").
   6. Esegui la funzione `setup` (selettore in alto). Concedi i permessi: vedrai
      richiedere SOLO l'accesso a questo Sheet (currentonly), non a tutto Drive.
   7. Distribuisci → Nuova distribuzione → Tipo "App web"
        - Esegui come: "Me"
        - Chi ha accesso: "Chiunque"
        Distribuisci e copia l'URL.

   AGGIORNAMENTO (se modifichi il codice in futuro)
   - Distribuisci → Gestisci distribuzioni → matita → "Nuova versione" → Distribuisci.
     L'URL del Web App rimane la stessa.

   HARDENING ATTIVI
   - Scope OAuth ristretto (spreadsheets.currentonly): anche se modificassi questo
     codice in futuro, Google NON permetterebbe più di accedere ad altri Sheets.
   - Token segreto condiviso (SECRET_TOKEN qui sotto) verificato in doPost.
     Senza token corretto, le POST vengono rifiutate.
   ========================================================================== */

const SHEET_NAME = "Responses";

// IMPORTANTE: questo token deve essere identico a quello in index.html (campo TOKEN).
// Se lo rigenerate, va aggiornato in entrambi i file.
const SECRET_TOKEN = "3ced0ef27e8067d7bb2c34b0a1d0b8df9268b57ab2983616";

const HEADERS = [
  "submittedAt", "anonymous", "userAgent",
  "A1_name", "A1_email",
  "A2_area", "A3_line", "A4_tenure",
  "B1_freq", "B2_tools", "B3_tasks", "B4_builder",
  "C1_blockers", "C2_literacy",
  "D1_timesinks", "D2_wish", "D3_risks",
  "E1_readiness", "E2_format", "E3_budget",
  "F1_ambassador", "F2_free",
  "_maturity", "_literacy", "_readiness", "_modules"
];

function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#3c3b30").setFontColor("#fcfaee");
  sh.autoResizeColumns(1, HEADERS.length);
  SpreadsheetApp.getUi().alert("Setup completato. Ora distribuisci come App web.");
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.__token !== SECRET_TOKEN) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: "unauthorized" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME) || setupSheet_();
    const p = payload._profile || {};
    const row = HEADERS.map(h => {
      if (h === "submittedAt")  return payload.__submittedAt || new Date().toISOString();
      if (h === "anonymous")    return !!payload.__anonymous;
      if (h === "userAgent")    return payload.__userAgent || "";
      if (h === "_maturity")    return p.maturity ?? "";
      if (h === "_literacy")    return p.literacy ?? "";
      if (h === "_readiness")   return p.readiness ?? "";
      if (h === "_modules")     return (p.modules || []).join(" | ");
      const v = payload[h];
      if (Array.isArray(v)) return v.join(" | ");
      return v ?? "";
    });
    sh.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupSheet_() {
  setup();
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function doGet() {
  return ContentService.createTextOutput("WAO AI Survey endpoint is alive.");
}
