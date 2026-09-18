// ============================================================
// DATA LUUK - Vocht, Vaste feiten, Taxi & Dialyse (Code.gs)
// ============================================================

const SPREADSHEET_ID = ""; 
const ID_KOLOM = "_ID";

const SHEETS = {
  "Vocht": ["Datum/tijd", "Soort", "Hoeveelheid ml", "Opmerking"],
  "Vaste feiten": ["Instelling", "Waarde", "Eenheid"],
  "Taxi": ["Afspraak datum/tijd", "Bestemming", "Taxibedrijf", "Start", "Stop", "Status", "Herhaling", "Einddatum herhaling", "Opmerking"],
  "Dialyse": ["Afspraak Datum/tijd", "Start Datum/tijd", "Stop Datum/tijd", "Type", "Status", "Pre-gewicht kg", "Post-gewicht kg", "UF L", "Bloeddruk metingen", "Opmerking", "Herhaling", "Einddatum herhaling"]
};

// Vaste wachtwoorden (pas deze gerust aan naar eigen wens)
const WACHTWOORDEN = {
  "admin": "luuk123",  // Volledige rechten (jij)
  "gast": "kijker123"  // Alleen leesrechten
};

function doGet() {
  setupDataLuuk();
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("DATA LUUK - Dashboard & Planning")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function login(gebruikersnaam, wachtwoord) {
  const gb = String(gebruikersnaam || "").trim().toLowerCase();
  const ww = String(wachtwoord || "").trim();

  if (WACHTWOORDEN[gb] && WACHTWOORDEN[gb] === ww) {
    return {
      succes: true,
      rol: gb === "admin" ? "admin" : "kijker",
      naam: gb === "admin" ? "Luuk (Beheerder)" : "Gast (Alleen lezen)"
    };
  }
  return { succes: false, fout: "Onjuiste gebruikersnaam of wachtwoord!" };
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
  }
  const actieveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!actieveSpreadsheet) {
    throw new Error("Google Sheet kan niet worden gevonden.");
  }
  return actieveSpreadsheet;
}

function setupDataLuuk() {
  const ss = getSpreadsheet_();
  Object.keys(SHEETS).forEach(function(sheetNaam) {
    let sheet = ss.getSheetByName(sheetNaam);
    if (!sheet) {
      sheet = ss.insertSheet(sheetNaam);
    }
    const headers = SHEETS[sheetNaam].concat([ID_KOLOM]);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const bestaandeHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
      headers.forEach(function(header, index) {
        if (String(bestaandeHeaders[index] || "").trim() !== header) {
          sheet.getRange(1, index + 1).setValue(header);
        }
      });
    }
    const idKolom = headers.indexOf(ID_KOLOM) + 1;
    if (idKolom > 0) {
      sheet.getRange(1, idKolom).setValue(ID_KOLOM);
    }
    sheet.setFrozenRows(1);
  });

  const vasteFeiten = ss.getSheetByName("Vaste feiten");
  if (vasteFeiten && vasteFeiten.getLastRow() <= 1) {
    vasteFeiten.getRange(2, 1, 13, 3).setValues([
      ["Drooggewicht", "109", "kg"],
      ["Vochtlimiet", "1000", "ml"],
      ["Drank", "Thee", ""],
      ["Drank", "Water", ""],
      ["Drank", "Koffie", ""],
      ["Drank", "Anders", ""],
      ["Taxibedrijf", "Connexxion", ""],
      ["Taxibedrijf", "Transvision", ""],
      ["Bestemming", "Ziekenhuis / Dialyse", ""],
      ["Bestemming", "Huis", ""],
      ["Bestemming", "Fysiotherapie", ""],
      ["Bloeddruk", "Standaard", ""]
    ]);
  }
  return true;
}

function getOverzicht() {
  setupDataLuuk();
  return {
    vocht: getData("Vocht"),
    vasteFeiten: getData("Vaste feiten"),
    taxi: getData("Taxi"),
    dialyse: getData("Dialyse")
  };
}

function getData(sheetNaam) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetNaam);
  if (!sheet) return [];
  const laatsteRij = sheet.getLastRow();
  const laatsteKolom = sheet.getLastColumn();
  if (laatsteRij < 2 || laatsteKolom < 1) return [];

  const range = sheet.getRange(1, 1, laatsteRij, laatsteKolom).getValues();
  const headers = range[0];
  const waarden = range.slice(1);

  return waarden
    .filter(rij => rij.some(waarde => waarde !== "" && waarde !== null))
    .map(rij => {
      const object = {};
      headers.forEach((header, index) => {
        let waarde = rij[index];
        if (waarde instanceof Date) {
          if (header.toLowerCase().includes("tijd") && !header.toLowerCase().includes("datum")) {
            waarde = Utilities.formatDate(waarde, Session.getScriptTimeZone(), "HH:mm");
          } else {
            waarde = formatDatumTijd_(waarde);
          }
        } else if (waarde && String(waarde).includes("1899")) {
          let match = String(waarde).match(/(\d{2}:\d{2})/);
          if (match) waarde = match[1];
        }
        object[header] = waarde;
      });
      return object;
    });
}

function addRecord(sheetNaam, data) {
  setupDataLuuk();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetNaam);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const nieuweRij = headers.map(header => {
    if (header === ID_KOLOM) return Utilities.getUuid();
    let waarde = data && Object.prototype.hasOwnProperty.call(data, header) ? data[header] : "";
    return verwerkWaarde_(header, waarde);
  });

  sheet.appendRow(nieuweRij);
  return true;
}

function addPlanningMetHerhaling(sheetNaam, data, datumVeldNaam) {
  const herhaling = data["Herhaling"] || "Eenmalig op datum";
  const basisDatumStr = data[datumVeldNaam]; 
  const eindDatumStr = data["Einddatum herhaling"]; 
  
  if (!basisDatumStr || herhaling === "Eenmalig op datum") { 
    addRecord(sheetNaam, data); 
    return; 
  }

  const afspraakDtMs = data["Afspraak Datum/tijd"] ? new Date(data["Afspraak Datum/tijd"]).getTime() : 0;
  const startDtMs = data["Start Datum/tijd"] ? new Date(data["Start Datum/tijd"]).getTime() : 0;
  const stopDtMs = data["Stop Datum/tijd"] ? new Date(data["Stop Datum/tijd"]).getTime() : 0;

  const startOffsetMs = startDtMs && afspraakDtMs ? startDtMs - afspraakDtMs : 0;
  const duurMs = stopDtMs && startDtMs ? stopDtMs - startDtMs : (5.5 * 3600 * 1000);

  const [datumDeel, tijdDeel] = basisDatumStr.split("T");
  const startDatum = new Date(datumDeel);
  
  let eindDatum = eindDatumStr ? new Date(eindDatumStr) : new Date(startDatum);
  if (!eindDatumStr) {
    eindDatum.setMonth(startDatum.getMonth() + 6);
  }

  let d = new Date(startDatum);
  while (d <= eindDatum) {
    let dagVanDeWeek = d.getDay();
    let overslaan = false;

    if (herhaling === "Om de dag") {
      let diffTime = Math.abs(d - startDatum);
      let diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays % 2 !== 0) overslaan = true;
    } else if (herhaling === "Ma, wo, vrij" && ![1, 3, 5].includes(dagVanDeWeek)) {
      overslaan = true;
    } else if (herhaling === "Di, do, za" && ![2, 4, 6].includes(dagVanDeWeek)) {
      overslaan = true;
    } else if (herhaling === "Elke maand") {
      if (d.getDate() !== startDatum.getDate()) {
        overslaan = true;
      }
    }

    if (!overslaan) {
      let jaar = d.getFullYear();
      let maand = String(d.getMonth() + 1).padStart(2, '0');
      let dag = String(d.getDate()).padStart(2, '0');
      let nieuweAfspraakTijd = `${jaar}-${maand}-${dag}T${tijdDeel || "00:00"}`;

      let kopieData = Object.assign({}, data);
      kopieData["Afspraak Datum/tijd"] = nieuweAfspraakTijd;

      let nieuweAfspraakMs = new Date(nieuweAfspraakTijd).getTime();
      if (nieuweAfspraakMs) {
        if (sheetNaam === "Dialyse") {
          let nieuweStartMs = nieuweAfspraakMs + startOffsetMs;
          let nieuweStopMs = nieuweStartMs + duurMs;

          kopieData["Start Datum/tijd"] = formatIsoDatetime_(new Date(nieuweStartMs));
          kopieData["Stop Datum/tijd"] = formatIsoDatetime_(new Date(nieuweStopMs));
        }
      }

      addRecord(sheetNaam, kopieData);
    }

    if (herhaling === "Elke maand") {
      d.setMonth(d.getMonth() + 1);
      d.setDate(startDatum.getDate());
    } else {
      d.setDate(d.getDate() + 1);
    }
  }
}

function updateRecord(sheetNaam, id, data) {
  setupDataLuuk();
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetNaam);
  const laatsteRij = sheet.getLastRow();
  if (laatsteRij < 2) throw new Error("Geen gegevens gevonden.");

  const range = sheet.getRange(1, 1, laatsteRij, sheet.getLastColumn()).getValues();
  const headers = range[0];
  const idIndex = headers.indexOf(ID_KOLOM);
  const gegevens = range.slice(1);

  let gevondenRijIndex = -1;
  for (let i = 0; i < gegevens.length; i++) {
    if (String(gegevens[i][idIndex]) === String(id)) {
      gevondenRijIndex = i + 2;
      break;
    }
  }

  if (gevondenRijIndex < 0) throw new Error("Registratie niet gevonden.");

  const oudeRij = sheet.getRange(gevondenRijIndex, 1, 1, headers.length).getValues()[0];
  const nieuweRij = oudeRij.slice();

  headers.forEach((header, index) => {
    if (header === ID_KOLOM) return;
    if (data && Object.prototype.hasOwnProperty.call(data, header)) {
      nieuweRij[index] = verwerkWaarde_(header, data[header]);
    }
  });

  sheet.getRange(gevondenRijIndex, 1, 1, headers.length).setValues([nieuweRij]);
  return true;
}

function deleteRecord(sheetNaam, id) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetNaam);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf(ID_KOLOM);
  if (idIndex < 0) return false;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idIndex]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function verwerkWaarde_(header, waarde) {
  if (waarde === null || waarde === undefined || waarde === "") return "";
  const tekst = String(waarde).trim();

  if (header.includes("Datum/tijd") || header === "Datum" || header === "Afspraak datum/tijd") {
    if (waarde instanceof Date) return formatDatumTijd_(waarde);
    if (tekst.includes("T")) {
      let delen = tekst.split("T");
      let datumDelen = delen[0].split("-");
      if (datumDelen.length === 3) {
        return `${datumDelen[2]}-${datumDelen[1]}-${datumDelen[0]}${delen[1] ? ' ' + delen[1] : ''}`;
      }
    }
    return tekst.replace("T", " ");
  }

  if (header.toLowerCase().includes("tijd") || header === "Start" || header === "Stop") {
    return tekst;
  }

  if (header === "Hoeveelheid ml" || header === "Pre-gewicht kg" || header === "Post-gewicht kg" || header === "UF L") {
    const getal = Number(tekst.replace(",", "."));
    return !isNaN(getal) ? getal : tekst;
  }
  return waarde;
}

function formatDatumTijd_(datum) {
  return Utilities.formatDate(datum, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
}

function formatIsoDatetime_(d) {
  let jaar = d.getFullYear();
  let maand = String(d.getMonth() + 1).padStart(2, '0');
  let dag = String(d.getDate()).padStart(2, '0');
  let uur = String(d.getHours()).padStart(2, '0');
  let min = String(d.getMinutes()).padStart(2, '0');
  return `${jaar}-${maand}-${dag}T${uur}:${min}`;
}
