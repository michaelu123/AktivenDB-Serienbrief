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
let emails = 0;
let headers: HeaderMap = {};
let aktDbSheet: GoogleAppsScript.Spreadsheet.Sheet;
let entries: MapI2S = {};
let antwortMap: MapS2I = {};

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
let registriertIndex: number; // Registriert für Erste-Hilfe-Training


function main() {
    init();
    let nrRows = aktDbSheet.getLastRow() - 1; // first row = headers
    let nrCols = aktDbSheet.getLastColumn();
    let rows = aktDbSheet.getRange(2, 1, nrRows, nrCols).getValues();
    for (let row of rows) { 
        sendeEmail(row);
    }
    Logger.log("total %d antworten %d emails %d", total, antworten, emails);
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
        if (sheetName != "AktivenDB" && sheetName != "Formularantworten 1") continue;
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
            registriertIndex = sheetHeaders["Registriert für Erste-Hilfe-Training"];

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
            entries[registriertIndex] = "entry.285304371";
            // entries[] = "entry.273898972"; // Einverstanden
            // entries[] = "entry.2103848384"; // Aktiv
        } 
        if (sheetName == "Formularantworten 1") { 
            let nrRows = sheet.getLastRow() - 1; // first row = headers
            let rows = sheet.getRange(2, 1, nrRows, 3).getValues();
            for (let row of rows) { 
              let name = row[1].trim() + "," + row[2].trim(); // Nachname,Vorname
              if (antwortMap[name]) {
                antwortMap[name] = antwortMap[name] + 1;
              } else {
                antwortMap[name] = 1; // Nachname,Vorname
              }
            }
            // antwortMap.size() returns always 0!
            // https://stackoverflow.com/questions/54518951/how-to-find-the-size-of-map-in-javascript
            let sz = 0;
            for (const k in antwortMap) {
              sz += 1;
              let v = antwortMap[k];
              if (v > 1) {
                Logger.log("duplicate antwort %s: %d", k, v);
              }
            } 
            Logger.log("size antwortMap %d", sz);
        }
    }
}

function sendeEmail(row: Array<string>) {
  let vorname = row[vornameIndex-1].trim();
  let nachname = row[nachnameIndex-1].trim();

  total++;
  let name = nachname + "," + vorname;
  if (antwortMap[name]) {
    Logger.log("Antwort %s", name);
    antworten++;
    return;   
  }

  let emailTo1 = row[emailAdfcIndex-1];
  let emailTo2 = row[emailPrivIndex-1];
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
  if (phase == 1) return;

  let templateFile = "email.html";

  let template: GoogleAppsScript.HTML.HtmlTemplate = HtmlService.createTemplateFromFile(
    templateFile
  );
  template.anrede = "Liebe(r) " + vorname + " " + nachname;
  template.verifLink = verifLinkUrl + row2Params(row); // encodeURIComponent(verifLinkParams);
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
      "YYYY-MM-dd"
    );
  }
  return val.toString();
}

function row2Params(row: Array<string>) {
    let params = [];
    for (let idx = 1; idx <= row.length; idx++) {
        if (idx > registriertIndex) continue; 
        let v = row[idx-1];
        if (isEmpty(v)) continue;
        Logger.log("row[%s] = %s %s", idx, v, typeof v);
        if (idx == mitgliedsnrIndex) { // v = 12345.0, number
            params.push("&" + entries[idx] + "=" + encodeURIComponent(v.toString()));
        } else if (idx == registriertIndex) { // v = true, boolean
            if (v) params.push("&" + entries[idx] + "=" + "ja/nein");
        } else if (idx == agsIndex) { // v = "ag1,ag2,ag3"
            let ags = v.split(",");
            for (let ag of ags) {
                ag = ag.trim();
                if (isEmpty(ag)) continue;
                params.push("&" + entries[idx] + "=" + encodeURIComponent(ag));
            }
        } else if (idx == lastFirstAidIndex) { // v = date
            params.push("&" + entries[idx] + "=" + any2Str(v));
        } else { // v = simple string or number param
            if (typeof v != "number") v = v.trim();
            params.push("&" + entries[idx] + "=" + encodeURIComponent(v));
        }     
    }
    let res = params.join("");
    Logger.log("res %s", res);
    return res;
} 

// URL of user aktive:
let verifLinkUrl = "https://docs.google.com/forms/d/e/1FAIpQLSfDjK7m42eofskS164D2qTj8e-7ngHZeoiSgwsMWzB-AG-xfA/viewform?usp=pp_url";
// URL of user mu:
// let verifLinkUrl = "https://docs.google.com/forms/d/e/1FAIpQLSdh7q00OHbeQdJ1ZMEy_LhXRPnMT3TJw-TeWcsVjZboWwJ2zA/viewform?usp=pp_url"

/*
let verifLink = "https://docs.google.com/forms/d/e/1FAIpQLSdh7q00OHbeQdJ1ZMEy_LhXRPnMT3TJw-TeWcsVjZboWwJ2zA/viewform?usp=pp_url&entry.1985977124=Nach+Name&entry.666565320=Vor+Name&entry.2076354113=email@adfc-muenchen.de&entry.440890410=email@t-online.de&entry.329829470=1234567&entry.1481160666=7654321&entry.1781476495=AG+Aktionen&entry.1781476495=AG+Asyl&entry.1781476495=Fundraising&entry.1674515812=Erstens+saufen,+zweitens+fressen&entry.1777914664=Kleine+Strasse+1,+34567+Klein-Kleckersdorf&entry.98896261=23456789&entry.1254417615=2017-01-08&entry.285304371=ja/nein&entry.1638875874=M&entry.931621781=1999&entry.273898972=Ja&entry.2103848384=Ja";

https://docs.google.com/forms/d/e/1FAIpQLSdh7q00OHbeQdJ1ZMEy_LhXRPnMT3TJw-TeWcsVjZboWwJ2zA/viewform
?usp=pp_url
&entry.1985977124=Nach+Name
&entry.666565320=Vor+Name
&entry.2076354113=email@adfc-muenchen.de
&entry.440890410=email@t-online.de
&entry.329829470=1234567
&entry.1481160666=7654321
&entry.1781476495=AG+Aktionen
&entry.1781476495=AG+Asyl
&entry.1781476495=AG+Codierung
&entry.1781476495=AG+Infoladen
&entry.1781476495=AG+IT
&entry.1781476495=AG+Leitungen
&entry.1781476495=AG+Mehrtagestouren
&entry.1781476495=AG+Navigation
&entry.1781476495=AG+Radfahrschule
&entry.1781476495=AG+Rikscha
&entry.1781476495=AG+Tagestouren
&entry.1781476495=AG+Tandem
&entry.1781476495=AG+Technik
&entry.1781476495=AG+Verkehr
&entry.1781476495=Event+Team
&entry.1781476495=Fundraising
&entry.1781476495=AG+Landkreis
&entry.1674515812=Erstens+saufen,+zweitens+fressen
&entry.1777914664=Kleine+Stra%C3%9Fe+1,+34567+Klein-Kleckersdorf
&entry.98896261=23456789
&entry.1254417615=2017-01-08
&entry.285304371=ja/nein
&entry.1638875874=M
&entry.931621781=1999
&entry.273898972=Ja
&entry.2103848384=Ja
*/