const SHEET_NAME = "FC Pro-T e.V Elternkasse";
const ARCHIV_NAME = "Gelöscht";

const ART_KASSE = "Bar"; // Bestehende Daten bleiben kompatibel
const ART_RUECKLAGE = "Rücklage";
const ART_PAYPAL = "PayPal";

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = p.action || "getEntries";
  let ergebnis;

  try {
    if (action === "getEntries") {
      ergebnis = getEntries();
    } else if (action === "add") {
      ergebnis = addEntry(p.betrag, p.beschreibung, p.erfasser, p.art, p.typ, p.buchungsdatum);
    } else if (action === "zaehle") {
      ergebnis = zaehleBar(p.betrag, p.erfasser, p.buchungsdatum, p.spieltag);
    } else if (action === "umbuch") {
      ergebnis = umbuchen(p.betrag, p.von, p.nach, p.erfasser, p.buchungsdatum);
    } else if (action === "delete") {
      ergebnis = storniereEntry(p.id, p.erfasser);
    } else if (action === "deleteTransfer") {
      ergebnis = storniereUmbuchung(p.ids, p.erfasser);
    } else if (action === "update") {
      ergebnis = updateEntry(p.id, p.betrag, p.beschreibung, p.erfasser, p.art, p.typ, p.buchungsdatum);
    } else {
      ergebnis = { fehler: "unbekannte Aktion" };
    }
  } catch (err) {
    ergebnis = { fehler: String(err && err.message ? err.message : err) };
  }

  const json = JSON.stringify(ergebnis);
  const cb = p.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getArchiv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let a = ss.getSheetByName(ARCHIV_NAME);

  if (!a) {
    a = ss.insertSheet(ARCHIV_NAME);
  }

  if (a.getLastRow() === 0) {
    a.appendRow([
      "Gelöscht am",
      "Gelöscht von",
      "Orig-Zeitstempel",
      "Datum",
      "Betrag",
      "Beschreibung",
      "Erfasser",
      "Art",
      "Typ",
      "Buchungsdatum",
      "ID",
      "Transfer-ID"
    ]);
  } else if (a.getLastColumn() < 12) {
    a.getRange(1, 12).setValue("Transfer-ID");
  }

  return a;
}

function zahl(wert) {
  return Number(String(wert).replace(",", ".")) || 0;
}

function runde2(wert) {
  return Math.round((Number(wert) || 0) * 100) / 100;
}

function formatDatum(wert) {
  if (wert === "" || wert === null || wert === undefined) return "";

  if (Object.prototype.toString.call(wert) === "[object Date]") {
    return Utilities.formatDate(wert, "Europe/Berlin", "dd.MM.yyyy");
  }

  return String(wert);
}

function isoVon(deDatum) {
  const t = String(deDatum).split(".");
  if (t.length !== 3) return "";

  return t[2] + "-" + t[1] + "-" + t[0];
}

function neueId() {
  return String(new Date().getTime()) + "-" + Math.floor(Math.random() * 100000);
}

function neueTransferId() {
  return "TR-" + neueId();
}

function artName(art) {
  if (art === ART_KASSE) return "Kasse";
  return art;
}

function istGueltigeArt(art) {
  return [
    ART_KASSE,
    ART_RUECKLAGE,
    ART_PAYPAL
  ].indexOf(String(art)) !== -1;
}

function appendBuchung(
  sheet,
  now,
  betrag,
  beschreibung,
  erfasser,
  art,
  typ,
  datum,
  status,
  transferId
) {
  sheet.appendRow([
    now,
    Utilities.formatDate(now, "Europe/Berlin", "dd.MM.yyyy HH:mm"),
    runde2(betrag),
    beschreibung || "",
    erfasser || "",
    art || ART_KASSE,
    typ || "Einnahme",
    datum || Utilities.formatDate(now, "Europe/Berlin", "dd.MM.yyyy"),
    neueId(),
    status || "",
    transferId || ""
  ]);
}

function addEntry(betrag, beschreibung, erfasser, art, typ, buchungsdatum) {
  const b = zahl(betrag);

  art = art || ART_KASSE;
  typ = String(typ || "Einnahme");

  if (b <= 0) {
    throw new Error("Der Betrag muss größer als 0 sein.");
  }

  if (!istGueltigeArt(art)) {
    throw new Error("Unbekannter Geldort.");
  }

  if (["Einnahme", "Ausgabe"].indexOf(typ) === -1) {
    throw new Error("Unbekannter Buchungstyp.");
  }

  if (typ === "Ausgabe") {
    const best = bestaende();

    if (b > (best[art] || 0) + 0.0001) {
      throw new Error(
        "Die Ausgabe ist höher als der Bestand bei " + artName(art) + "."
      );
    }
  }

  const sheet = getSheet();
  const now = new Date();

  appendBuchung(
    sheet,
    now,
    b,
    beschreibung,
    erfasser,
    art,
    typ,
    buchungsdatum,
    "",
    ""
  );

  return getEntries();
}

function bestaende() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  values.shift();

  const result = {};

  result[ART_KASSE] = 0;
  result[ART_RUECKLAGE] = 0;
  result[ART_PAYPAL] = 0;

  values.forEach(r => {
    if (r[2] === "") return;

    const art = String(r[5] || ART_KASSE);

    if (result[art] === undefined) {
      result[art] = 0;
    }

    const betrag = Number(r[2]) || 0;

    result[art] += String(r[6]) === "Ausgabe"
      ? -betrag
      : betrag;
  });

  Object.keys(result).forEach(k => {
    result[k] = runde2(result[k]);
  });

  return result;
}

function barBestand() {
  return bestaende()[ART_KASSE] || 0;
}

function zaehleBar(gezaehlt, erfasser, buchungsdatum, spieltag) {
  const neu = zahl(gezaehlt);

  if (neu < 0) {
    throw new Error("Der gezählte Bestand darf nicht negativ sein.");
  }

  const alt = barBestand();
  const diff = runde2(neu - alt);

  if (diff === 0) {
    return getEntries();
  }

  const sheet = getSheet();
  const now = new Date();

  const typ = diff >= 0
    ? "Einnahme"
    : "Ausgabe";

  const st = spieltag
    ? ("Spieltag " + spieltag + " – ")
    : "";

  appendBuchung(
    sheet,
    now,
    Math.abs(diff),
    st + "Kassenzählung (Stand: " + neu.toFixed(2).replace(".", ",") + " €)",
    erfasser,
    ART_KASSE,
    typ,
    buchungsdatum,
    "",
    ""
  );

  return getEntries();
}

function umbuchen(betrag, von, nach, erfasser, buchungsdatum) {
  const b = zahl(betrag);

  von = String(von || "");
  nach = String(nach || "");

  if (b <= 0) {
    throw new Error("Der Betrag muss größer als 0 sein.");
  }

  if (!istGueltigeArt(von) || !istGueltigeArt(nach)) {
    throw new Error("Unbekannter Geldort.");
  }

  if (von === nach) {
    throw new Error("Quelle und Ziel dürfen nicht gleich sein.");
  }

  const best = bestaende();

  if (b > (best[von] || 0) + 0.0001) {
    throw new Error(
      "Nicht genügend Guthaben bei " + artName(von) + "."
    );
  }

  const sheet = getSheet();
  const now = new Date();

  const datum =
    buchungsdatum ||
    Utilities.formatDate(now, "Europe/Berlin", "dd.MM.yyyy");

  const beschreibung =
    "Umbuchung " +
    artName(von) +
    " → " +
    artName(nach);

  const transferId = neueTransferId();

  appendBuchung(
    sheet,
    now,
    b,
    beschreibung,
    erfasser,
    von,
    "Ausgabe",
    datum,
    "",
    transferId
  );

  appendBuchung(
    sheet,
    now,
    b,
    beschreibung,
    erfasser,
    nach,
    "Einnahme",
    datum,
    "",
    transferId
  );

  return getEntries();
}

function archivieren(archiv, now, storniertVon, r) {
  archiv.appendRow([
    Utilities.formatDate(now, "Europe/Berlin", "dd.MM.yyyy HH:mm"),
    storniertVon || "",
    r[0],
    r[1],
    r[2],
    r[3],
    r[4],
    r[5],
    r[6],
    r[7],
    r[8],
    r[10] || ""
  ]);
}

function storniereEntry(id, storniertVon) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][8]) !== String(id)) {
      continue;
    }

    const r = values[i];

    if (String(r[10] || "")) {
      throw new Error(
        "Umbuchungen können nur vollständig storniert werden."
      );
    }

    if (/^Umbuchung\s+.+\s+→\s+.+$/i.test(String(r[3] || ""))) {
      throw new Error(
        "Umbuchungen können nur vollständig storniert werden."
      );
    }

    if (String(r[9]) === "storniert") {
      return getEntries();
    }

    if (String(r[9]) === "storno") {
      throw new Error(
        "Eine Stornozeile kann nicht erneut storniert werden."
      );
    }

    const row = i + 1;
    const now = new Date();

    sheet.getRange(row, 10).setValue("storniert");

    const origBetrag = Number(r[2]) || 0;
    const origTyp = String(r[6]);

    const gegenTyp =
      origTyp === "Ausgabe"
        ? "Einnahme"
        : "Ausgabe";

    appendBuchung(
      sheet,
      now,
      origBetrag,
      "Storno: " + String(r[3] || ""),
      storniertVon || "",
      String(r[5] || ART_KASSE),
      gegenTyp,
      formatDatum(r[7]) || formatDatum(r[0]),
      "storno",
      ""
    );

    archivieren(
      getArchiv(),
      now,
      storniertVon,
      r
    );

    return getEntries();
  }

  throw new Error("Buchung wurde nicht gefunden.");
}

function storniereUmbuchung(idsText, storniertVon) {
  const ids = String(idsText || "")
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);

  if (ids.length !== 2) {
    throw new Error(
      "Die Umbuchung ist unvollständig und kann nicht sicher storniert werden."
    );
  }

  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  const treffer = [];

  for (let i = 1; i < values.length; i++) {
    if (ids.indexOf(String(values[i][8])) !== -1) {
      treffer.push({
        index: i,
        row: values[i]
      });
    }
  }

  if (treffer.length !== 2) {
    throw new Error(
      "Die beiden Teile der Umbuchung wurden nicht gefunden."
    );
  }

  const a = treffer[0].row;
  const b = treffer[1].row;

  if (
    String(a[9]) === "storniert" &&
    String(b[9]) === "storniert"
  ) {
    return getEntries();
  }

  if (String(a[9]) || String(b[9])) {
    throw new Error(
      "Die Umbuchung hat bereits einen abweichenden Status. Bitte nicht manuell verändern."
    );
  }

  if (runde2(a[2]) !== runde2(b[2])) {
    throw new Error(
      "Die Umbuchung ist nicht konsistent."
    );
  }

  if (String(a[3]) !== String(b[3])) {
    throw new Error(
      "Die Umbuchung ist nicht konsistent."
    );
  }

  if (String(a[6]) === String(b[6])) {
    throw new Error(
      "Die Umbuchung ist nicht konsistent."
    );
  }

  const now = new Date();
  const archiv = getArchiv();

  const vorhandeneTransferId =
    String(a[10] || b[10] || "");

  const gemeinsameTransferId =
    vorhandeneTransferId ||
    ("TR-LEGACY-" + neueId());

  treffer.forEach(t => {
    const r = t.row;

    sheet
      .getRange(t.index + 1, 10)
      .setValue("storniert");

    sheet
      .getRange(t.index + 1, 11)
      .setValue(gemeinsameTransferId);

    const gegenTyp =
      String(r[6]) === "Ausgabe"
        ? "Einnahme"
        : "Ausgabe";

    appendBuchung(
      sheet,
      now,
      Number(r[2]) || 0,
      "Storno: " + String(r[3] || ""),
      storniertVon || "",
      String(r[5] || ART_KASSE),
      gegenTyp,
      formatDatum(r[7]) || formatDatum(r[0]),
      "storno",
      gemeinsameTransferId
    );

    const archivZeile = r.slice();
    archivZeile[10] = gemeinsameTransferId;

    archivieren(
      archiv,
      now,
      storniertVon,
      archivZeile
    );
  });

  return getEntries();
}

function updateEntry(
  id,
  betrag,
  beschreibung,
  erfasser,
  art,
  typ,
  buchungsdatum
) {
  const b = zahl(betrag);

  art = String(art || "");
  typ = String(typ || "");

  if (b <= 0) {
    throw new Error(
      "Der Betrag muss größer als 0 sein."
    );
  }

  if (!istGueltigeArt(art)) {
    throw new Error(
      "Unbekannter Geldort."
    );
  }

  if (["Einnahme", "Ausgabe"].indexOf(typ) === -1) {
    throw new Error(
      "Unbekannter Buchungstyp."
    );
  }

  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][8]) !== String(id)) {
      continue;
    }

    const alt = values[i];

    if (
      String(alt[9]) === "storniert" ||
      String(alt[9]) === "storno"
    ) {
      throw new Error(
        "Stornierte Buchungen können nicht bearbeitet werden."
      );
    }

    if (
      String(alt[10] || "") ||
      /^Umbuchung\s+.+\s+→\s+.+$/i.test(String(alt[3] || ""))
    ) {
      throw new Error(
        "Umbuchungen können nicht einzeln bearbeitet werden."
      );
    }

    const best = bestaende();

    const altArt =
      String(alt[5] || ART_KASSE);

    const altBetrag =
      Number(alt[2]) || 0;

    const altFaktor =
      String(alt[6]) === "Ausgabe"
        ? -1
        : 1;

    best[altArt] = runde2(
      (best[altArt] || 0) -
      altFaktor * altBetrag
    );

    const neuFaktor =
      typ === "Ausgabe"
        ? -1
        : 1;

    best[art] = runde2(
      (best[art] || 0) +
      neuFaktor * b
    );

    if ((best[art] || 0) < -0.0001) {
      throw new Error(
        "Durch diese Änderung würde " +
        artName(art) +
        " ins Minus rutschen."
      );
    }

    const row = i + 1;

    sheet.getRange(row, 3).setValue(b);
    sheet.getRange(row, 4).setValue(beschreibung || "");
    sheet.getRange(row, 5).setValue(erfasser || "");
    sheet.getRange(row, 6).setValue(art);
    sheet.getRange(row, 7).setValue(typ);
    sheet.getRange(row, 8).setValue(buchungsdatum || "");

    return getEntries();
  }

  throw new Error("Buchung wurde nicht gefunden.");
}

function legacyTransferKey(r) {
  const beschreibung =
    String(r[3] || "");

  if (
    !/^Umbuchung\s+.+\s+→\s+.+$/i.test(beschreibung)
  ) {
    return "";
  }

  const ts =
    Object.prototype.toString.call(r[0]) === "[object Date]"
      ? r[0].getTime()
      : String(r[0]);

  return (
    "LEGACY|" +
    ts +
    "|" +
    runde2(r[2]) +
    "|" +
    beschreibung +
    "|" +
    formatDatum(r[7]) +
    "|" +
    String(r[4] || "")
  );
}

function getEntries() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();

  values.shift();

  let kasse = 0;
  let ruecklage = 0;
  let paypal = 0;
  let maxSt = 0;

  const normale = [];
  const transferGruppen = {};

  for (let i = 0; i < values.length; i++) {
    const r = values[i];

    if (r[2] === "") {
      continue;
    }

    const betrag =
      Number(r[2]) || 0;

    const art =
      String(r[5] || ART_KASSE);

    const typ =
      String(r[6] || "Einnahme");

    const beschreibung =
      String(r[3] || "");

    const status =
      String(r[9] || "");

    const transferId =
      String(r[10] || "");

    const faktor =
      typ === "Ausgabe"
        ? -1
        : 1;

    if (art === ART_KASSE) {
      kasse += faktor * betrag;
    } else if (art === ART_RUECKLAGE) {
      ruecklage += faktor * betrag;
    } else if (art === ART_PAYPAL) {
      paypal += faktor * betrag;
    }

    const m =
      beschreibung.match(/Spieltag\s+(\d+)/i);

    if (m) {
      const n =
        parseInt(m[1], 10);

      if (n > maxSt) {
        maxSt = n;
      }
    }

    let d =
      formatDatum(r[7]);

    if (!d) {
      d = formatDatum(r[0]);
    }

    const basis = {
      datum: d,
      iso: isoVon(d),
      betrag: betrag,
      beschreibung: beschreibung,
      erfasser: String(r[4] || ""),
      art: art,
      typ: typ,
      id: String(r[8] || ""),
      status: status,
      transferId: transferId,
      zeitwert:
        Object.prototype.toString.call(r[0]) === "[object Date]"
          ? r[0].getTime()
          : i
    };

    const legacyKey =
      legacyTransferKey(r);

    const gruppenKey =
      transferId
        ? ("TR|" + transferId)
        : legacyKey;

    if (
      gruppenKey &&
      status === "storno"
    ) {
      if (!transferGruppen[gruppenKey]) {
        transferGruppen[gruppenKey] = [];
      }

      transferGruppen[gruppenKey].push(basis);
      continue;
    }

    if (gruppenKey) {
      if (!transferGruppen[gruppenKey]) {
        transferGruppen[gruppenKey] = [];
      }

      transferGruppen[gruppenKey].push(basis);
    } else {
      normale.push(basis);
    }
  }

  const entries = normale.slice();

  Object.keys(transferGruppen).forEach(key => {
    const gruppe =
      transferGruppen[key];

    const originals =
      gruppe.filter(e => e.status !== "storno");

    if (originals.length < 2) {
      originals.forEach(e => entries.push(e));
      return;
    }

    const ausgabe =
      originals.find(e => e.typ === "Ausgabe");

    const einnahme =
      originals.find(e => e.typ === "Einnahme");

    if (
      !ausgabe ||
      !einnahme ||
      runde2(ausgabe.betrag) !== runde2(einnahme.betrag)
    ) {
      originals.forEach(e => entries.push(e));
      return;
    }

    const komplettStorniert =
      originals.every(
        e => e.status === "storniert"
      );

    const teilweiseStorniert =
      originals.some(
        e => e.status === "storniert"
      ) && !komplettStorniert;

    entries.push({
      kind: "transfer",
      datum: ausgabe.datum,
      iso: ausgabe.iso,
      betrag: ausgabe.betrag,
      beschreibung: ausgabe.beschreibung,
      erfasser:
        ausgabe.erfasser ||
        einnahme.erfasser,
      von: ausgabe.art,
      nach: einnahme.art,
      ids: [
        ausgabe.id,
        einnahme.id
      ],
      transferId:
        ausgabe.transferId ||
        einnahme.transferId ||
        "",
      status:
        komplettStorniert
          ? "storniert"
          : (
              teilweiseStorniert
                ? "teilweise"
                : ""
            ),
      zeitwert:
        Math.max(
          ausgabe.zeitwert,
          einnahme.zeitwert
        )
    });
  });

  kasse = runde2(kasse);
  ruecklage = runde2(ruecklage);
  paypal = runde2(paypal);

  entries.sort(
    (a, b) =>
      (b.zeitwert || 0) -
      (a.zeitwert || 0)
  );

  entries.forEach(e => {
    delete e.zeitwert;
  });

  return {
    entries: entries,

    barBestand: kasse,
    kasseBestand: kasse,

    ruecklageBestand: ruecklage,
    paypalBestand: paypal,

    gesamtBestand:
      runde2(
        kasse +
        ruecklage +
        paypal
      ),

    maxSpieltag: maxSt
  };
}