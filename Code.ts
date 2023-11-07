interface MapS2I {
  [others: string]: number;
}
interface MapI2S {
  [others: number]: string;
}
interface MapS2S {
  [others: string]: string;
}
interface HeaderMap {
  [others: string]: MapS2I;
}

let inited = false;
let phase = 1;
let total = 0;
let antworten = 0;
let inaktiv = 0;
let emails = 0;
let headers: HeaderMap = {};
let aktDbSheet: GoogleAppsScript.Spreadsheet.Sheet;
let antwortSheet: GoogleAppsScript.Spreadsheet.Sheet;
let entries: MapI2S = {};
let antwortMap: MapS2I = {};
let dbMap: MapS2I = {};

let nachnameIndex: number; // Nachname
let vornameIndex: number; // Vorname
let geschlechtIndex: number; // Geschlecht
let geburtsjahrIndex: number; // Geburtsjahr
let postleitzahlIndex: number; // Postleitzahl
let mitgliedsnrIndex: number; // ADFC-Mitgliedsnummer
let emailAdfcIndex: number; // Email-ADFC
let emailPrivIndex: number; // Email-Privat
let telefonIndex: number; // Telefon
let telefonAltIndex: number; // Telefon-Alternative
let agsIndex: number; // AGs
let interessenIndex: number; // Interessen
let lastFirstAidIndex: number; // Letztes Erste-Hilfe-Training
let nextFirstAidIndex: number; // Nächstes Erste-Hilfe-Training
// let registriertIndex: number; // Registriert für Erste-Hilfe-Training
let aktivIndex: number; // Aktiv
let statusIndex: number; // Status

let nowDate = any2Str(Date());

function main() {
  init();
  let nrAktdbRows = aktDbSheet.getLastRow() - 1; // first row = headers
  let nrAktdbCols = aktDbSheet.getLastColumn();
  let aktdbRows = aktDbSheet
    .getRange(2, 1, nrAktdbRows, nrAktdbCols)
    .getValues();
  for (let row of aktdbRows) {
    let name =
      row[nachnameIndex - 1].trim() + "," + row[vornameIndex - 1].trim(); // Nachname,Vorname
    dbMap[name] = 1;
  }

  let nrAntwortRows = antwortSheet.getLastRow() - 1; // first row = headers
  let antwortRows =
    nrAntwortRows == 0
      ? []
      : antwortSheet.getRange(2, 1, nrAntwortRows, 3).getValues();
  for (let row of antwortRows) {
    let name = row[1].trim() + "," + row[2].trim(); // Nachname,Vorname
    if (antwortMap[name]) {
      antwortMap[name] = antwortMap[name] + 1;
    } else {
      antwortMap[name] = 1; // Nachname,Vorname
    }
    if (!dbMap[name]) {
      Logger.log("Antwortname %s nicht in der DB", name);
    }
  }

  // antwortMap.size() returns always 0!
  // https://stackoverflow.com/questions/54518951/how-to-find-the-size-of-map-in-javascript
  let sz = 0;
  for (const k in antwortMap) {
    sz += 1;
    let v = antwortMap[k];
    if (v > 1) {
      Logger.log("Doppelte Antwort %s: %d", k, v);
    }
  }
  Logger.log("Größe antwortMap %d", sz);

  for (let row of aktdbRows) {
    sendeEmail(row);
  }
  Logger.log(
    "total %d antworten %d inaktiv %d emails %d",
    total,
    antworten,
    inaktiv,
    emails,
  );
}

function isEmpty(str: string | undefined | null) {
  if (typeof str == "number") return false;
  return !str || 0 === str.length; // I think !str is sufficient...
}

function init() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  for (let sheet of sheets) {
    let sheetName = sheet.getName();
    let sheetHeaders: MapS2I = {};
    Logger.log("sheetName %s", sheetName);
    if (sheetName != "AktivenDB" && sheetName != "Formularantworten 1")
      continue;
    headers[sheetName] = sheetHeaders;
    let numCols = sheet.getLastColumn();
    Logger.log("numCols %s", numCols);
    let row1Vals = sheet.getRange(1, 1, 1, numCols).getValues();
    Logger.log("sheetName %s row1 %s", sheetName, row1Vals);
    for (let i = 0; i < numCols; i++) {
      let v: string = row1Vals[0][i];
      if (isEmpty(v)) continue;
      sheetHeaders[v] = i + 1;
    }
    // Sheet aus AktivenDB mit Export erzeugt
    if (sheetName == "AktivenDB") {
      aktDbSheet = sheet;
      nachnameIndex = sheetHeaders["Nachname"];
      vornameIndex = sheetHeaders["Vorname"];
      geschlechtIndex = sheetHeaders["Geschlecht"];
      geburtsjahrIndex = sheetHeaders["Geburtsjahr"];
      postleitzahlIndex = sheetHeaders["Postleitzahl"];
      mitgliedsnrIndex = sheetHeaders["ADFC-Mitgliedsnummer"];
      emailAdfcIndex = sheetHeaders["Email-ADFC"];
      emailPrivIndex = sheetHeaders["Email-Privat"];
      telefonIndex = sheetHeaders["Telefon"];
      telefonAltIndex = sheetHeaders["Telefon-Alternative"];
      agsIndex = sheetHeaders["AGs"];
      interessenIndex = sheetHeaders["Interessen"];
      lastFirstAidIndex = sheetHeaders["Letztes Erste-Hilfe-Training"];
      nextFirstAidIndex = sheetHeaders["Nächstes Erste-Hilfe-Training"];
      aktivIndex = sheetHeaders["Aktiv"];
      statusIndex = sheetHeaders["Status"];

      entries[nachnameIndex] = "entry.1985977124";
      entries[vornameIndex] = "entry.666565320";
      entries[geschlechtIndex] = "entry.1638875874";
      entries[geburtsjahrIndex] = "entry.931621781";
      entries[postleitzahlIndex] = "entry.1777914664";
      entries[mitgliedsnrIndex] = "entry.98896261";
      entries[emailAdfcIndex] = "entry.2076354113";
      entries[emailPrivIndex] = "entry.440890410";
      entries[telefonIndex] = "entry.329829470";
      entries[telefonAltIndex] = "entry.1481160666";
      entries[agsIndex] = "entry.1781476495";
      entries[interessenIndex] = "entry.1674515812";
      entries[lastFirstAidIndex] = "entry.1254417615";
      entries[nextFirstAidIndex] = "entry.285304371";
      entries[statusIndex] = "entry.583933307";
      // entries[] = "entry.273898972"; // Einverstanden mit Speicherung
      // entries[] = "entry.2103848384"; // Aktiv
    }
    if (sheetName == "Formularantworten 1") {
      antwortSheet = sheet;
    }
  }
}

function sendeEmail(row: Array<string>) {
  let vorname = row[vornameIndex - 1].trim();
  let nachname = row[nachnameIndex - 1].trim();

  total++;
  let name = nachname + "," + vorname;
  if (antwortMap[name]) {
    Logger.log("Antwort %s", name);
    antworten++;
    return;
  }
  if (row[aktivIndex] == "FALSCH") {
    Logger.log("Inaktiv %s", name);
    inaktiv++;
    return;
  }

  let emailTo1 = row[emailPrivIndex - 1];
  let emailTo2 = row[emailAdfcIndex - 1];
  let emailTo = "";
  if (!isEmpty(emailTo1)) {
    emailTo = emailTo1;
    if (!isEmpty(emailTo2)) {
      emailTo = emailTo + "," + emailTo2;
    }
  } else if (!isEmpty(emailTo2)) {
    emailTo = emailTo2;
  } else {
    Logger.log("Keine Email Adresse für " + vorname + " " + nachname);
    return;
  }
  Logger.log("emailTo %s = %s", name, emailTo);

  emails++;

  let templateFile = "email.html";

  let template: GoogleAppsScript.HTML.HtmlTemplate =
    HtmlService.createTemplateFromFile(templateFile);
  template.anrede = "Liebe(r) " + vorname + " " + nachname;
  template.verifLink = verifLinkUrl + row2Params(row); // encodeURIComponent(verifLinkParams);
  Logger.log("verifLink %s", template.verifLink);
  if (phase == 1) return;
  let htmlText: string = template.evaluate().getContent();
  let subject = "Aktualisierung Deiner Daten in der AktivenDB";
  let textbody = "HTML only";
  let options = {
    htmlBody: htmlText,
    name: "ADFC München e.V.",
    replyTo: "aktive@adfc-muenchen.de",
  };
  GmailApp.sendEmail(emailTo, subject, textbody, options);
}

function any2Str(val: any): string {
  if (typeof val == "object" && "getUTCHours" in val) {
    return Utilities.formatDate(
      val,
      SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
      "YYYY-MM-dd",
    );
  }
  return val.toString();
}

function row2Params(row: Array<string>) {
  let params = [];
  for (let idx = 1; idx <= row.length; idx++) {
    if (idx > nextFirstAidIndex) continue;
    let entry = entries[idx];
    if (!entry) continue;
    let v = row[idx - 1];
    if (isEmpty(v)) continue;
    Logger.log("row[%s] = %s %s", idx, v, typeof v);
    if (idx == mitgliedsnrIndex) {
      // v = 12345.0, number
      params.push("&" + entry + "=" + encodeURIComponent(v.toString()));
    } /* else if (idx == nextFirstAidIndex) { // v = true, boolean  
            if (v) params.push("&" + entry + "=" + "ja/nein");
        } */ else if (idx == agsIndex) {
      // v = "ag1,ag2,ag3"
      let ags = v.split(",");
      for (let ag of ags) {
        ag = ag.trim();
        if (isEmpty(ag)) continue;
        params.push("&" + entries[idx] + "=" + encodeURIComponent(ag));
      }
    } else if (idx == lastFirstAidIndex) {
      // v = date
      params.push("&" + entry + "=" + any2Str(v));
    } else if (idx == nextFirstAidIndex) {
      let ndate = any2Str(v);
      if (ndate.length >= 10 && ndate <= nowDate) ndate = "";
      params.push("&" + entry + "=" + ndate);
    } else if (idx == telefonIndex || idx == telefonAltIndex) {
      // v = string, remove blank,-
      params.push("&" + entry + "=" + any2Str(v).replace(/[\s-]/g, ""));
    } else {
      // v = simple string or number param
      if (typeof v != "number") v = v.trim();
      params.push("&" + entry + "=" + encodeURIComponent(v));
    }
  }
  let res = params.join("");
  Logger.log("res %s", res);
  return res;
}

// URL of user aktive:
let verifLinkUrl =
  "https://docs.google.com/forms/d/e/1FAIpQLSfDjK7m42eofskS164D2qTj8e-7ngHZeoiSgwsMWzB-AG-xfA/viewform?usp=pp_url";

/*
https://docs.google.com/forms/d/e/1FAIpQLSfDjK7m42eofskS164D2qTj8e-7ngHZeoiSgwsMWzB-AG-xfA/viewform?usp=pp_url&entry.1985977124=Nachname&entry.666565320=Vorname&entry.1638875874=M&entry.931621781=1111&entry.1777914664=2222&entry.2076354113=mail@adfc-muenchen.de&entry.440890410=mail@mail.de&entry.329829470=012345678&entry.1481160666=0987654321&entry.1781476495=AG+Aktionen&entry.1781476495=AG+Asyl&entry.1781476495=AG+Codierung&entry.1781476495=OG+Putzbrunn&entry.1781476495=OG+Stra%C3%9Flach-Dingharting&entry.1781476495=OG+Unterhaching&entry.1674515812=Interessen&entry.98896261=345678&entry.1254417615=2022-08-11&entry.285304371=1999-09-09&entry.273898972=Ja&entry.2103848384=Ja
*/
