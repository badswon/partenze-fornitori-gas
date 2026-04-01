/****************************************************
 * HEADER OGGI – testo stabile (NO formattazioni)
 * Celle giorni: A2, C2, E2, G2, I2, K2 (Lun..Sab)
 * - oggi:  ▶️ OGGI – MARTEDÌ 27/01 ◀️
 * - altri: MARTEDÌ 27/01
 ****************************************************/
function AGGIORNA_HEADER_OGGI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Settimana Attiva");
  if (!sh) return;

  const base = sh.getRange("M2").getValue();
  if (!(base instanceof Date)) return;

  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);

  const headerCells = ["A2", "C2", "E2", "G2", "I2", "K2"];
  const dayNames = ["LUNEDÌ", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ", "SABATO"];

  const base0 = new Date(base);
  base0.setHours(0, 0, 0, 0);

  const diffDays = Math.floor((oggi.getTime() - base0.getTime()) / 86400000);
  const todayIndex = (diffDays >= 0 && diffDays <= 5) ? diffDays : -1;

  const tz = Session.getScriptTimeZone();

  for (let i = 0; i < 6; i++) {
    const d = new Date(base0);
    d.setDate(d.getDate() + i);

    const ddmm = Utilities.formatDate(d, tz, "dd/MM");
    const normal = `${dayNames[i]} ${ddmm}`;

    const txt = (i === todayIndex)
      ? `▶️ OGGI – ${dayNames[i]} ${ddmm} ◀️`
      : normal;

    sh.getRange(headerCells[i]).setValue(txt);
  }
}

/** Versione safe: non rompe nulla anche se chiamata spesso */
function AGGIORNA_HEADER_OGGI_SAFE_() {
  try { AGGIORNA_HEADER_OGGI(); } catch (e) {}
}

/** Wrapper giornaliero per trigger */
function AggiornaOggiGiornaliero() {
  try { AGGIORNA_HEADER_OGGI_SAFE_(); } catch (e) {}
  try { RIPRISTINA_EVIDENZIA_OGGI(); } catch (e) {}
}

function TEST_CREA_PROSSIMA_SILENT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let before = !!ss.getSheetByName("Prossima settimana");
  Logger.log("Prima esiste? " + before);

  let res;
  try {
    res = creaProssimaSettimana_Silent_();
    Logger.log("Risultato funzione: " + res);
  } catch (e) {
    Logger.log("ERRORE creaProssimaSettimana_Silent_: " + (e && e.message ? e.message : e));
  }

  let after = !!ss.getSheetByName("Prossima settimana");
  Logger.log("Dopo esiste? " + after);
}

function TEST_N2_CREAZIONE_PROSSIMA_LIGHT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = !!ss.getSheetByName("Prossima settimana");

  let aggiornata = "FUNZIONE_MANCANTE";
  try {
    if (typeof _settimanaAttivaAggiornata_ === "function") {
      aggiornata = _settimanaAttivaAggiornata_();
    }
  } catch (e) {
    aggiornata = "ERRORE: " + (e && e.message ? e.message : e);
  }

  Logger.log("Prossima settimana esiste? " + exists);
  Logger.log("Settimana attiva aggiornata? " + aggiornata);
}

function TEST_DIAGNOSI_HEADER_OGGI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Settimana Attiva");

  if (!sh) {
    Logger.log("ERRORE: Settimana Attiva non trovata");
    return;
  }

  const base = sh.getRange("M2").getValue();
  Logger.log("M2 raw: " + base);
  Logger.log("M2 type: " + Object.prototype.toString.call(base));
  Logger.log("M2 instanceof Date: " + (base instanceof Date));

  const oggi = new Date();
  oggi.setHours(0, 0, 0, 0);
  Logger.log("Oggi: " + oggi);

  if (!(base instanceof Date)) {
    Logger.log("STOP: M2 non è una Date valida");
    return;
  }

  const base0 = new Date(base);
  base0.setHours(0, 0, 0, 0);

  const diffDays = Math.floor((oggi.getTime() - base0.getTime()) / 86400000);
  const todayIndex = (diffDays >= 0 && diffDays <= 5) ? diffDays : -1;

  Logger.log("base0: " + base0);
  Logger.log("diffDays: " + diffDays);
  Logger.log("todayIndex: " + todayIndex);

  const headerCells = ["A2", "C2", "E2", "G2", "I2", "K2"];
  const dayNames = ["LUNEDÌ", "MARTEDÌ", "MERCOLEDÌ", "GIOVEDÌ", "VENERDÌ", "SABATO"];
  const tz = Session.getScriptTimeZone();

  for (let i = 0; i < 6; i++) {
    const d = new Date(base0);
    d.setDate(d.getDate() + i);

    const ddmm = Utilities.formatDate(d, tz, "dd/MM");
    const normal = `${dayNames[i]} ${ddmm}`;
    const txt = (i === todayIndex)
      ? `▶️ OGGI – ${dayNames[i]} ${ddmm} ◀️`
      : normal;

    Logger.log(headerCells[i] + " -> " + txt);
  }
}
