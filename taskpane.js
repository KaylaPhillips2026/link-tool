// JLL Link Tool - task pane logic
// Reads a Forbury model's named ranges directly (client-side, via SheetJS)
// and populates matching bookmarks in the active Word document.
//
// This mirrors LinkToolReport.bas exactly:
//   - single-value bookmarks: TXT_/DOL_/NUM_/PCT_/DAT_/TDP_/ODP_ prefix + named range name
//   - table bookmarks: TAB_ prefix + Table_/Chart_ named range name

const SINGLE_VALUE_PREFIXES = ["TXT_", "DOL_", "NUM_", "PCT_", "DAT_", "TDP_", "ODP_"];

let workbook = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("fileInput").addEventListener("change", handleFileSelect);
    document.getElementById("runButton").addEventListener("click", runLinkTool);
  }
});

function setStatus(text) {
  document.getElementById("status").textContent = text;
}

function handleFileSelect(event) {
  const file = event.target.files[0];
  if (!file) {
    workbook = null;
    document.getElementById("runButton").disabled = true;
    return;
  }

  setStatus("Reading " + file.name + "...");
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array" });
      setStatus("Loaded " + file.name + ". Ready to run.");
      document.getElementById("runButton").disabled = false;
    } catch (err) {
      setStatus("Could not read this file: " + err.message);
      workbook = null;
      document.getElementById("runButton").disabled = true;
    }
  };
  reader.onerror = function () {
    setStatus("Failed to read the file.");
  };
  reader.readAsArrayBuffer(file);
}

// ---------------------------------------------------------------
// Named range resolution (mirrors fWb.Names(name).RefersToRange in VBA)
// ---------------------------------------------------------------

function parseNamedRangeRef(refText) {
  // refText looks like: General!$D$15  or  'Sheet Name'!$D$15:$H$20
  const m = refText.match(/^'?([^'!]+)'?!(\$?[A-Z]+\$?[0-9]+)(?::(\$?[A-Z]+\$?[0-9]+))?/);
  if (!m) return null;
  return {
    sheetName: m[1],
    startRef: m[2].replace(/\$/g, ""),
    endRef: m[3] ? m[3].replace(/\$/g, "") : m[2].replace(/\$/g, ""),
  };
}

function resolveNamedRangeValue(name) {
  const names = (workbook.Workbook && workbook.Workbook.Names) || [];
  const match = names.find((x) => x.Name === name);
  if (!match) return { found: false };

  const parsed = parseNamedRangeRef(match.Ref);
  if (!parsed) return { found: false };

  const ws = workbook.Sheets[parsed.sheetName];
  if (!ws) return { found: false };

  const cell = ws[parsed.startRef];
  return { found: true, value: cell ? cell.v : undefined };
}

function resolveNamedRangeTable(name) {
  const names = (workbook.Workbook && workbook.Workbook.Names) || [];
  const match = names.find((x) => x.Name === name);
  if (!match) return { found: false };

  const parsed = parseNamedRangeRef(match.Ref);
  if (!parsed) return { found: false };

  const ws = workbook.Sheets[parsed.sheetName];
  if (!ws) return { found: false };

  const range = XLSX.utils.decode_range(parsed.startRef + ":" + parsed.endRef);
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      row.push(cell !== undefined ? String(cell.v !== undefined ? cell.v : "") : "");
    }
    rows.push(row);
  }
  return { found: true, rows };
}

// ---------------------------------------------------------------
// Formatting (mirrors FormatByPrefix in VBA)
// ---------------------------------------------------------------

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  return new Date(utcDays * 86400 * 1000);
}

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

function formatByPrefix(prefix, value) {
  if (value === undefined || value === null) return "";

  try {
    switch (prefix) {
      case "DOL_": {
        const n = Number(value);
        const rounded = Math.round(n);
        const abs = Math.abs(rounded).toLocaleString("en-AU");
        return rounded < 0 ? "($" + abs + ")" : "$" + abs;
      }
      case "NUM_": {
        return Math.round(Number(value)).toLocaleString("en-AU");
      }
      case "PCT_": {
        return (Number(value) * 100).toFixed(1) + "%";
      }
      case "DAT_": {
        const d = typeof value === "number" ? excelSerialToDate(value) : new Date(value);
        return d.getDate() + " " + MONTH_NAMES[d.getMonth()] + " " + d.getFullYear();
      }
      case "TDP_":
      case "ODP_": {
        return Number(value).toFixed(1);
      }
      default: // TXT_
        return String(value);
    }
  } catch (e) {
    return String(value);
  }
}

// ---------------------------------------------------------------
// Main routine
// ---------------------------------------------------------------

async function runLinkTool() {
  if (!workbook) {
    setStatus("Select a Forbury model first.");
    return;
  }

  document.getElementById("runButton").disabled = true;
  document.getElementById("results").hidden = true;
  setStatus("Populating fields...");

  const singleOK = [];
  const singleMiss = [];
  const tableOK = [];
  const tableMiss = [];

  try {
    await Word.run(async (context) => {
      const bodyRange = context.document.body.getRange();
      const bookmarkNames = bodyRange.getBookmarks();
      await context.sync();

      const names = bookmarkNames.value;

      // --- Part 1: single-value bookmarks ---
      for (const bmName of names) {
        const prefix = SINGLE_VALUE_PREFIXES.find((p) => bmName.startsWith(p));
        if (!prefix) continue;

        const rangeName = bmName.substring(prefix.length);
        const resolved = resolveNamedRangeValue(rangeName);

        if (!resolved.found) {
          singleMiss.push(bmName);
          continue;
        }

        const formatted = formatByPrefix(prefix, resolved.value);
        const bmRange = context.document.getBookmarkRangeOrNullObject(bmName);
        bmRange.load("isNullObject");
        await context.sync();

        if (bmRange.isNullObject) {
          singleMiss.push(bmName + " (bookmark missing at write time)");
          continue;
        }

        bmRange.insertText(formatted, Word.InsertLocation.replace);
        await context.sync();

        // Re-anchor the bookmark - inserting text destroys it, same as VBA
        const newRange = context.document.getBookmarkRangeOrNullObject(bmName);
        newRange.load("isNullObject");
        await context.sync();
        if (newRange.isNullObject) {
          // insertText's returned range needs to be re-bookmarked manually
          const sel = context.document.getSelection();
          sel.insertBookmark(bmName);
          await context.sync();
        }

        singleOK.push(bmName);
      }

      setStatus("Inserting tables...");

      // --- Part 2: table bookmarks ---
      for (const bmName of names) {
        if (!bmName.startsWith("TAB_")) continue;

        const sourceName = bmName.substring(4);
        const resolved = resolveNamedRangeTable(sourceName);

        if (!resolved.found) {
          tableMiss.push(bmName + " (source name: " + sourceName + ")");
          continue;
        }

        const bmRange = context.document.getBookmarkRangeOrNullObject(bmName);
        bmRange.load("isNullObject");
        await context.sync();

        if (bmRange.isNullObject) {
          tableMiss.push(bmName + " (bookmark missing at write time)");
          continue;
        }

        bmRange.insertText("", Word.InsertLocation.replace);
        await context.sync();

        const sel = context.document.getSelection();
        const table = sel.insertTable(
          resolved.rows.length,
          resolved.rows[0].length,
          Word.InsertLocation.replace,
          resolved.rows
        );
        table.getRange().insertBookmark(bmName);
        await context.sync();

        tableOK.push(bmName);
      }
    });

    showResults(singleOK, singleMiss, tableOK, tableMiss);
    setStatus("Done.");
  } catch (err) {
    setStatus("Error: " + err.message);
    console.error(err);
  }

  document.getElementById("runButton").disabled = false;
}

function showResults(singleOK, singleMiss, tableOK, tableMiss) {
  const resultsDiv = document.getElementById("results");
  resultsDiv.hidden = false;

  document.getElementById("fieldsSummary").innerHTML =
    "Fields: <span class=\"ok\">" + singleOK.length + " populated</span>" +
    (singleMiss.length ? ", <span class=\"fail\">" + singleMiss.length + " missing</span>" : "");

  document.getElementById("tablesSummary").innerHTML =
    "Tables: <span class=\"ok\">" + tableOK.length + " inserted</span>" +
    (tableMiss.length ? ", <span class=\"fail\">" + tableMiss.length + " missing</span>" : "");

  const missList = document.getElementById("missList");
  missList.innerHTML = "";

  if (singleMiss.length) {
    const h = document.createElement("div");
    h.className = "miss-heading";
    h.textContent = "Missing fields:";
    missList.appendChild(h);
    singleMiss.forEach((m) => {
      const item = document.createElement("div");
      item.className = "miss-item";
      item.textContent = m;
      missList.appendChild(item);
    });
  }

  if (tableMiss.length) {
    const h = document.createElement("div");
    h.className = "miss-heading";
    h.textContent = "Missing tables:";
    missList.appendChild(h);
    tableMiss.forEach((m) => {
      const item = document.createElement("div");
      item.className = "miss-item";
      item.textContent = m;
      missList.appendChild(item);
    });
  }
}
