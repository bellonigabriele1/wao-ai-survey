/* ============================================================================
   WAO AI Survey — Google Apps Script backend
   ----------------------------------------------------------------------------
   COME INSTALLARLO (5 minuti)
   1. Vai su https://sheets.google.com e crea un nuovo Google Sheet.
      Rinominalo "WAO AI Survey — Responses".
   2. Dal menu: Estensioni → Apps Script.
   3. Cancella il codice di default e incolla TUTTO questo file.
   4. Clicca "Salva" (icona dischetto), dai un nome al progetto (es. "wao-survey").
   5. Clicca "Esegui" sulla funzione `setup` (selettore in alto). Concedi i permessi
      al tuo account quando richiesto. Questo crea l'header del foglio.
   6. Clicca "Distribuisci" (in alto a destra) → "Nuova distribuzione".
      - Tipo: "App web"
      - Descrizione: "WAO AI Survey endpoint"
      - Esegui come: "Me (tuo@waospaces.it)"
      - Chi ha accesso: "Chiunque" (la URL fa da chiave)
      Clicca "Distribuisci" e copia la "URL app web".
   7. Incolla quella URL nel file index.html, costante `ENDPOINT`.
   8. Fatto. Le risposte arrivano sul Sheet in tempo reale.

   NOTE
   - L'endpoint accetta solo POST con corpo JSON (text/plain per evitare CORS).
   - In caso di nuove domande, aggiungi la chiave a HEADERS qui sotto e ri-esegui
     setup() per aggiornare l'intestazione (non sovrascrive le righe già esistenti).
   ========================================================================== */

const SHEET_NAME = "Responses";

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
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0E0E10").setFontColor("#E8FF5E");
  sh.autoResizeColumns(1, HEADERS.length);
  SpreadsheetApp.getUi().alert("Setup completato. Ora distribuisci come App web.");
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
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
