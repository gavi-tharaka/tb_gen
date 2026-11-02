/* ==========================================================================
   GL Trial Balance Generator - excel-tools.js   (ES module, imported by app.js)

   Pure helpers with no DOM access:
     - number helpers          parseAmount, toCents, formatMoney
     - numberToWords           "Zero Only" / "... Rupees Only" / "... Rupees and ... Cents"
     - normalizeDate           any common date input -> YYYY.MM.DD
     - parsePastedRows         Excel clipboard text -> schedule items
     - parseWorkbookFile       previous-month .xlsx -> GL blocks (scans "GL Account Number and Name:")
     - buildWorkbook           selected GLs -> one ExcelJS worksheet (Sheet1)

   ExcelJS is loaded as a global (<script> in index.html) and read via globalThis.ExcelJS.
   ========================================================================== */


/* --------------------------------------------------------------------------
   Export style. These values were read from final_trail_sample.xlsx, so the
   output matches the sample. Change a value here to change every block.
   -------------------------------------------------------------------------- */
export const STYLE = {
  fontName: "Sansation",          // falls back automatically if the font is not installed
  fontSize: 10,
  certFontSize: 11,
  boldLabels: false,              // sample: header labels are not bold (spec says bold -> set true)
  headerValueAlign: "left",       // sample: GL / TB balance values are left aligned (spec says right)
  amountFormat: "#,##0.00;\\(#,##0.00\\)",   // item amounts, as in the sample
  totalFormat: "#,##0.00",
  columnWidths: [14, 16, 22, 42, 20],         // A..E
  tableHeaderHeight: 25.5,
  sigBoxRows: 6,                  // height of each signature box, in rows
  sigRowHeight: 18,               // points per signature row (room for pen and rubber stamp)
  gapRows: 3,                     // blank rows between GL blocks
  pageBreakPerBlock: true,        // each GL block starts on a new printed page
  sheetName: "Sheet1"
};


/* ==========================================================================
   NUMBER HELPERS
   ========================================================================== */

/** "1,234.50", "(500.00)", "-20", "Rs. 1,000" -> number. Returns NaN when it is not a number. */
export function parseAmount(value) {
  if (typeof value === "number") return isFinite(value) ? value : NaN;
  let s = String(value == null ? "" : value).trim();
  if (!s) return NaN;
  let negative = false;
  if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1); }
  s = s.replace(/^(rs\.?|lkr)\s*/i, "").replace(/[,\s]/g, "");
  if (s.charAt(0) === "-") { negative = true; s = s.slice(1); }
  if (!/^(\d+\.?\d*|\.\d+)$/.test(s)) return NaN;
  const n = parseFloat(s);
  return negative ? -n : n;
}

/** Money -> integer cents, so sums never drift by floating point error. */
export function toCents(n) {
  return Math.round((n + (n < 0 ? -Number.EPSILON : Number.EPSILON)) * 100);
}

/** Integer cents -> "1,234.50" */
export function formatMoney(cents) {
  return (cents / 100).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}


/* ==========================================================================
   NUMBER TO WORDS
   0                -> "Zero Only"
   1793000          -> "One Million Seven Hundred Ninety Three Thousand Rupees Only"
   674581.25        -> "Six Hundred Seventy Four Thousand Five Hundred Eighty One Rupees and Twenty Five Cents"
   No "and" is used inside the rupee figure.
   ========================================================================== */
const ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
  "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
const TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
const SCALES = [
  { value: 1e9, name: "Billion" },
  { value: 1e6, name: "Million" },
  { value: 1e3, name: "Thousand" }
];

function underThousand(n) {
  const parts = [];
  if (n >= 100) { parts.push(ONES[Math.floor(n / 100)], "Hundred"); n %= 100; }
  if (n >= 20)  { parts.push(TENS[Math.floor(n / 10)]); n %= 10; }
  if (n > 0)    { parts.push(ONES[n]); }
  return parts.join(" ");
}

function integerToWords(n) {
  if (n === 0) return "Zero";
  const parts = [];
  SCALES.forEach(function (s) {
    if (n >= s.value) {
      parts.push(underThousand(Math.floor(n / s.value)), s.name);
      n %= s.value;
    }
  });
  if (n > 0) parts.push(underThousand(n));
  return parts.join(" ");
}

export function numberToWords(value) {
  const num = Number(value);
  const cents = Math.round(Math.abs(isFinite(num) ? num : 0) * 100);
  if (cents === 0) return "Zero Only";

  const rupees = Math.floor(cents / 100);
  const c = cents % 100;
  const sign = num < 0 ? "Minus " : "";
  const rupeeWord = rupees === 1 ? "Rupee" : "Rupees";
  const centWord = c === 1 ? "Cent" : "Cents";

  if (c === 0)      return sign + integerToWords(rupees) + " " + rupeeWord + " Only";
  if (rupees === 0) return sign + integerToWords(c) + " " + centWord;
  return sign + integerToWords(rupees) + " " + rupeeWord + " and " + integerToWords(c) + " " + centWord;
}


/* ==========================================================================
   DATE NORMALISER  ->  "YYYY.MM.DD"
   Accepts Date objects, Excel serial numbers, 2026-09-30, 2026/9/3, 30/09/2026 (day first).
   Anything unrecognised is returned trimmed and unchanged.
   ========================================================================== */
const pad2 = function (n) { return String(n).padStart(2, "0"); };

export function normalizeDate(input) {
  if (input == null || input === "") return "";

  if (input instanceof Date) {
    if (isNaN(input.getTime())) return "";
    return input.getUTCFullYear() + "." + pad2(input.getUTCMonth() + 1) + "." + pad2(input.getUTCDate());
  }

  if (typeof input === "number") {
    // Excel serial date (days since 1899-12-30)
    if (input > 20000 && input < 80000) {
      return normalizeDate(new Date(Math.round((input - 25569) * 86400 * 1000)));
    }
    return String(input);
  }

  const s = String(input).trim();
  let m = s.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
  if (m) return m[1] + "." + pad2(m[2]) + "." + pad2(m[3]);
  m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/);
  if (m) return m[3] + "." + pad2(m[2]) + "." + pad2(m[1]);
  return s;
}

export function isValidDate(s) {
  const m = /^(\d{4})\.(\d{2})\.(\d{2})$/.exec(String(s || ""));
  if (!m) return false;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
}


/* ==========================================================================
   PASTE PARSER (tab separated text copied from Excel)
   Layout A: Date | Reference | Description | Amount
   Layout B: Serial No | Date | Reference | Description | Amount
   A header row is skipped automatically.
   Returns { items: [{date, ref, desc, amount}], skipped: number }
   ========================================================================== */
const looksLikeDate = function (s) {
  return /^\d{4}[.\-\/]\d{1,2}[.\-\/]\d{1,2}/.test(s) || /^\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{4}/.test(s);
};

export function parsePastedRows(text) {
  const items = [];
  let skipped = 0;
  const lines = String(text || "").replace(/\r/g, "").split("\n");

  lines.forEach(function (line) {
    if (!line.trim()) return;
    const cols = line.split("\t").map(function (c) { return c.trim(); });

    // Header row?
    if (/^(serial|date)\b/i.test(cols[0]) && /amount/i.test(cols[cols.length - 1])) return;

    let date, ref, desc, amount;
    const withSerial = cols.length >= 5 && /^\d+$/.test(cols[0]) && looksLikeDate(cols[1]);

    if (withSerial) {
      date = cols[1]; ref = cols[2]; amount = cols[cols.length - 1];
      desc = cols.slice(3, cols.length - 1).join(" ");
    } else if (cols.length >= 4) {
      date = cols[0]; ref = cols[1]; amount = cols[cols.length - 1];
      desc = cols.slice(2, cols.length - 1).join(" ");
    } else {
      skipped++;
      return;
    }

    const n = parseAmount(amount);
    if (isNaN(n)) { skipped++; return; }
    items.push({ date: normalizeDate(date), ref: ref, desc: desc, amount: String(Math.round(n * 100) / 100) });
  });

  return { items: items, skipped: skipped };
}


/* ==========================================================================
   READ A PREVIOUS-MONTH WORKBOOK
   Finds every "GL Account Number and Name:" block on every sheet and returns
   [{ glNo, glName, items, glBalance, diffText, sheet }]
   ========================================================================== */
function rawValue(cell) {
  let v = cell.value;
  if (v == null) return null;
  if (typeof v === "object" && !(v instanceof Date)) {
    if ("result" in v) v = v.result;
    else if (Array.isArray(v.richText)) v = v.richText.map(function (t) { return t.text; }).join("");
    else if ("text" in v) v = v.text;
    else v = null;
  }
  return v;
}

function textOf(cell) {
  const v = rawValue(cell);
  if (v == null) return "";
  if (v instanceof Date) return normalizeDate(v);
  return String(v).trim();
}

function rowValueAfterLabel(row) {
  const labelText = textOf(row.getCell(1));
  let val = null;
  [4, 5, 2, 3].forEach(function (col) {
    if (val !== null && val !== "") return;
    const rv = rawValue(row.getCell(col));
    const tv = textOf(row.getCell(col));
    if (tv !== labelText && rv !== null && String(rv).trim() !== "") {
      val = rv;
    }
  });
  return val;
}

export async function parseWorkbookFile(file) {
  const ExcelJS = globalThis.ExcelJS;
  if (!ExcelJS) throw new Error("ExcelJS library did not load. Check your internet connection.");

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await file.arrayBuffer());

  const blocks = [];

  wb.eachSheet(function (ws) {
    const last = ws.rowCount;
    const starts = [];
    for (let r = 1; r <= last; r++) {
      if (/^gl account number and name/i.test(textOf(ws.getRow(r).getCell(1)))) starts.push(r);
    }

    starts.forEach(function (startRow, idx) {
      const endRow = idx + 1 < starts.length ? starts[idx + 1] - 1 : last;
      const labelText = textOf(ws.getRow(startRow).getCell(1));

      // The value sits in D:E (merged). Merged slaves repeat the label, so start at column D.
      let value = "";
      [4, 5, 2, 3].forEach(function (col) {
        if (value) return;
        const t = textOf(ws.getRow(startRow).getCell(col));
        if (t && t !== labelText) value = t;
      });
      const m = value.match(/^(\d{6})\s*-\s*(.*)$/);
      if (!m) return;

      const block = {
        glNo: m[1],
        glName: m[2].trim(),
        tbDate: "",
        items: [],
        glBalance: null,
        tbBalance: null,
        diffText: "",
        sheet: ws.name
      };
      let headerRow = 0;

      // Also check 1-2 rows above startRow in case Trial Balance Date is above GL Account row
      for (let preR = Math.max(1, startRow - 2); preR < startRow; preR++) {
        const preRow = ws.getRow(preR);
        const preA = textOf(preRow.getCell(1)).toLowerCase();
        if ((preA.indexOf("trial balance date") === 0 || preA.indexOf("date as at") !== -1) && !block.tbDate) {
          const dv = rowValueAfterLabel(preRow);
          if (dv != null) block.tbDate = normalizeDate(dv);
        }
      }

      for (let r = startRow + 1; r <= endRow; r++) {
        const row = ws.getRow(r);
        const a = textOf(row.getCell(1)).toLowerCase();

        if ((a.indexOf("trial balance date") === 0 || a.indexOf("date as at") !== -1) && !block.tbDate) {
          const dv = rowValueAfterLabel(row);
          if (dv != null) block.tbDate = normalizeDate(dv);
        } else if (a.indexOf("general ledger balance") === 0 && block.glBalance === null) {
          const n = parseAmount(rowValueAfterLabel(row));
          if (!isNaN(n)) block.glBalance = n;
        } else if (a.indexOf("balance of trial") === 0 && block.tbBalance === null) {
          const n = parseAmount(rowValueAfterLabel(row));
          if (!isNaN(n)) block.tbBalance = n;
        } else if (a === "serial no" && !headerRow) {
          headerRow = r;
        } else if (headerRow && a.indexOf("total balance") === 0) {
          headerRow = -1;                      // data rows finished
        } else if (a.indexOf("difference") === 0) {
          const t = textOf(row.getCell(5)) || textOf(row.getCell(4));
          block.diffText = t;
        } else if (headerRow > 0) {
          const dateV = rawValue(row.getCell(2));
          const refV  = rawValue(row.getCell(3));
          const descV = rawValue(row.getCell(4));
          const amtV  = rawValue(row.getCell(5));
          const empty = [dateV, refV, descV, amtV].every(function (x) { return x == null || String(x).trim() === ""; });
          if (empty) continue;
          const n = parseAmount(amtV);
          block.items.push({
            date: normalizeDate(dateV),
            ref: refV == null ? "" : String(refV).trim(),
            desc: descV == null ? "" : String(descV).trim(),
            amount: isNaN(n) ? "" : String(Math.round(n * 100) / 100)
          });
        } else if (!block.tbDate && a.indexOf("we hereby certify") === 0) {
          // Fallback: extract date at the end of certification text ("... as at YYYY.MM.DD")
          const certDateMatch = textOf(row.getCell(1)).match(/as at\s+(\d{4}[.\-\/]\d{1,2}[.\-\/]\d{1,2})/i);
          if (certDateMatch) block.tbDate = normalizeDate(certDateMatch[1]);
        }
      }
      blocks.push(block);
    });
  });

  return blocks;
}


/* ==========================================================================
   BUILD THE FINAL WORKBOOK
   params = {
     officerName, serviceNo, tbDate,
     blocks: [{
       glNo, glName,
       mode: "zero" | "balance",
       items: [{ date, ref, desc, amount:Number }],
       tbCents, glCents,           // Trial Balance total and General Ledger balance (integer cents)
       diffText                    // "Nil" or the user's explanation
     }]
   }
   Returns an ExcelJS.Workbook (call workbook.xlsx.writeBuffer() to save it).
   ========================================================================== */
const THIN = { style: "thin", color: { argb: "FF000000" } };

function font(bold, size) {
  return { name: STYLE.fontName, size: size || STYLE.fontSize, bold: !!bold };
}

/** Thin outline around a rectangular range (works for merged ranges and single cells). */
function outline(ws, r1, c1, r2, c2) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      ws.getCell(r, c).border = {
        top:    r === r1 ? THIN : undefined,
        bottom: r === r2 ? THIN : undefined,
        left:   c === c1 ? THIN : undefined,
        right:  c === c2 ? THIN : undefined
      };
    }
  }
}

/** Merge (if more than one cell), write a value, style it, and draw its border. */
function put(ws, r1, c1, r2, c2, value, opts) {
  opts = opts || {};
  if (r1 !== r2 || c1 !== c2) ws.mergeCells(r1, c1, r2, c2);
  const cell = ws.getCell(r1, c1);
  cell.value = value;
  cell.font = font(opts.bold, opts.size);
  cell.alignment = {
    horizontal: opts.align || "left",
    vertical: opts.valign || "middle",
    wrapText: !!opts.wrap,
    indent: opts.indent || 0
  };
  if (opts.numFmt) cell.numFmt = opts.numFmt;
  if (opts.border !== false) outline(ws, r1, c1, r2, c2);
  return cell;
}

/** Long numeric references (e.g. 196100130016076) are stored as numbers like the sample; others stay text. */
function referenceValue(ref) {
  const s = String(ref == null ? "" : ref).trim();
  if (/^[1-9]\d{0,14}$/.test(s)) return Number(s);
  return s;
}

function writeBlock(ws, startRow, params, block) {
  const isZero = block.mode !== "balance";
  const items = isZero ? [] : block.items;
  const tbValue = block.tbCents / 100;
  const glValue = block.glCents / 100;
  const labelBold = STYLE.boldLabels;
  let r = startRow;

  // ---- Header rows 1-5 (A:C label | D:E value) ----
  const headerRows = [
    ["Name and Service Number:",       params.officerName + " - " + params.serviceNo, null],
    ["GL Account Number and Name:",   block.glNo + " - " + block.glName,             null],
    ["Trial Balance Date As at:",     params.tbDate,                                  null],
    ["General Ledger Balance :",      isZero ? 0 : glValue,                           STYLE.totalFormat],
    ["Balance of Trial :",            isZero ? 0 : tbValue,                           STYLE.totalFormat]
  ];
  headerRows.forEach(function (h) {
    put(ws, r, 1, r, 3, h[0], { bold: labelBold });
    put(ws, r, 4, r, 5, h[1], { numFmt: h[2], align: h[2] ? STYLE.headerValueAlign : "left" });
    r++;
  });

  // ---- Table header ----
  ["Serial No", "Date", "Reference", "Description", "Amount"].forEach(function (t, i) {
    put(ws, r, i + 1, r, i + 1, t, { bold: true, align: "center" });
  });
  ws.getRow(r).height = STYLE.tableHeaderHeight;
  r++;

  // ---- Data rows ----
  const firstDataRow = r;
  items.forEach(function (it, i) {
    put(ws, r, 1, r, 1, i + 1, { align: "center" });
    put(ws, r, 2, r, 2, it.date, { align: "left" });
    put(ws, r, 3, r, 3, referenceValue(it.ref), { align: "left" });
    put(ws, r, 4, r, 4, it.desc, { align: "left", wrap: true });
    put(ws, r, 5, r, 5, it.amount, { align: "right", numFmt: STYLE.amountFormat });
    r++;
  });
  const lastDataRow = r - 1;

  // ---- Total row ----
  put(ws, r, 1, r, 4, "Total Balance of Trial Balance", { align: "center" });
  let totalValue;
  if (items.length > 1) {
    totalValue = { formula: "SUM(E" + firstDataRow + ":E" + lastDataRow + ")", result: tbValue };
  } else if (items.length === 1) {
    totalValue = items[0].amount;
  } else {
    totalValue = 0;
  }
  put(ws, r, 5, r, 5, totalValue, { bold: true, numFmt: STYLE.totalFormat, align: "right" });
  r++;

  // ---- Blank row, then amount-in-words and difference ----
  r++;
  const tbWords = isZero ? "Zero Only" : numberToWords(tbValue);
  const glWords = isZero ? "Zero Only" : numberToWords(glValue);
  const diffText = isZero ? "NIL" : (block.diffText || "Nil");

  [
    ["Trial Balance Amount in Words",              tbWords],
    ["General Ledger Account Balance in Words  :", glWords],
    ["Difference (If Any):",                       diffText]
  ].forEach(function (w) {
    if (isZero) {
      put(ws, r, 1, r, 4, w[0], { align: "left" });
      put(ws, r, 5, r, 5, w[1], { align: "left" });
    } else {
      put(ws, r, 1, r, 3, w[0], { align: "left" });
      put(ws, r, 4, r, 5, w[1], { align: "left", wrap: true });
    }
    r++;
  });

  // ---- Blank row, then certification (3 rows x 5 columns, wrapped)
  //      ONLY included when GL Balance and Trial Balance tally (no difference) ----
  if (!block.hasDiscrepancy) {
    r++;
    put(ws, r, 1, r + 2, 5,
      "We hereby certify that the total balance of GL Account No." + block.glNo +
      " has been duly verified and tallied with the total amount of the corresponding Trial Balance as at " + params.tbDate,
      { align: "left", valign: "middle", wrap: true, size: STYLE.certFontSize, border: false });
    r += 3;
  }

  // ---- Blank row, then 1 single merged signature box (2*h rows x 5 columns)
  //      with 4 floating Excel Text Boxes placed on top of it ----
  r++;
  const h = STYLE.sigBoxRows;
  const totalSigRows = h * 2;
  put(ws, r, 1, r + totalSigRows - 1, 5, "", { align: "left", valign: "top", border: true });
  for (let i = 0; i < totalSigRows; i++) ws.getRow(r + i).height = STYLE.sigRowHeight;

  if (Array.isArray(ws._sigBoxesMeta)) {
    ws._sigBoxesMeta.push({
      startRow: r,
      halfRows: h,
      totalRows: totalSigRows,
      boxes: [
        { pos: "tl", text: "Prepared by : " + params.officerName },
        { pos: "tr", text: "Checked by : Operations manager" },
        { pos: "bl", text: "Checked by : Manager" },
        { pos: "br", text: "Audit officer :" }
      ]
    });
  }
  r += totalSigRows;

  // ---- Gap before the next block ----
  return r + STYLE.gapRows;
}

function xmlEscape(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" }[c];
  });
}

function buildDrawingXml(sigBoxesMeta) {
  let shapesXml = "";
  let shapeId = 2;

  sigBoxesMeta.forEach(function (meta) {
    const r0 = meta.startRow - 1; // 0-based row index of the top of the merged signature container
    const h = meta.halfRows;      // 6 rows per tier

    // Coordinates inside the merged A..E (col 0..4) container, with clean margins
    const coords = {
      tl: { fromCol: 0, fromColOff: 133350, fromRow: r0,     fromRowOff: 95250, toCol: 3, toColOff: 152400,  toRow: r0 + h - 1,     toRowOff: 133350 },
      tr: { fromCol: 3, fromColOff: 419100, fromRow: r0,     fromRowOff: 95250, toCol: 4, toColOff: 1181100, toRow: r0 + h - 1,     toRowOff: 133350 },
      bl: { fromCol: 0, fromColOff: 133350, fromRow: r0 + h, fromRowOff: 76200, toCol: 3, toColOff: 152400,  toRow: r0 + h * 2 - 1, toRowOff: 114300 },
      br: { fromCol: 3, fromColOff: 419100, fromRow: r0 + h, fromRowOff: 76200, toCol: 4, toColOff: 1181100, toRow: r0 + h * 2 - 1, toRowOff: 114300 }
    };

    meta.boxes.forEach(function (box) {
      const c = coords[box.pos];
      if (!c) return;
      const id = shapeId++;
      shapesXml +=
        '<xdr:twoCellAnchor>' +
          '<xdr:from>' +
            '<xdr:col>' + c.fromCol + '</xdr:col>' +
            '<xdr:colOff>' + c.fromColOff + '</xdr:colOff>' +
            '<xdr:row>' + c.fromRow + '</xdr:row>' +
            '<xdr:rowOff>' + c.fromRowOff + '</xdr:rowOff>' +
          '</xdr:from>' +
          '<xdr:to>' +
            '<xdr:col>' + c.toCol + '</xdr:col>' +
            '<xdr:colOff>' + c.toColOff + '</xdr:colOff>' +
            '<xdr:row>' + c.toRow + '</xdr:row>' +
            '<xdr:rowOff>' + c.toRowOff + '</xdr:rowOff>' +
          '</xdr:to>' +
          '<xdr:sp macro="" textlink="">' +
            '<xdr:nvSpPr>' +
              '<xdr:cNvPr id="' + id + '" name="TextBox ' + id + '"/>' +
              '<xdr:cNvSpPr txBox="1"/>' +
            '</xdr:nvSpPr>' +
            '<xdr:spPr>' +
              '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
              '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
              '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
            '</xdr:spPr>' +
            '<xdr:txBody>' +
              '<a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>' +
              '<a:lstStyle/>' +
              '<a:p>' +
                '<a:pPr algn="l"/>' +
                '<a:r>' +
                  '<a:rPr lang="en-US" sz="1000" b="0">' +
                    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>' +
                    '<a:latin typeface="' + xmlEscape(STYLE.fontName) + '"/>' +
                    '<a:cs typeface="' + xmlEscape(STYLE.fontName) + '"/>' +
                  '</a:rPr>' +
                  '<a:t>' + xmlEscape(box.text) + '</a:t>' +
                '</a:r>' +
              '</a:p>' +
            '</xdr:txBody>' +
          '</xdr:sp>' +
          '<xdr:clientData/>' +
        '</xdr:twoCellAnchor>';
    });
  });

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      shapesXml +
    '</xdr:wsDr>'
  );
}

export function buildWorkbook(params) {
  const ExcelJS = globalThis.ExcelJS;
  if (!ExcelJS) throw new Error("ExcelJS library did not load. Check your internet connection.");

  const wb = new ExcelJS.Workbook();
  wb.creator = params.officerName;
  wb.created = new Date();

  const ws = wb.addWorksheet(STYLE.sheetName, {
    pageSetup: {
      paperSize: 9,                 // A4
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
    }
  });

  ws._sigBoxesMeta = [];
  wb._sigBoxesMeta = ws._sigBoxesMeta;

  STYLE.columnWidths.forEach(function (w, i) { ws.getColumn(i + 1).width = w; });

  let row = 1;
  params.blocks.forEach(function (block, i) {
    row = writeBlock(ws, row, params, block);
    if (STYLE.pageBreakPerBlock && i < params.blocks.length - 1) {
      ws.getRow(row - 1).addPageBreak();
    }
  });

  return wb;
}

export async function buildWorkbookBuffer(params) {
  const wb = buildWorkbook(params);
  const rawBuffer = await wb.xlsx.writeBuffer();
  const sigBoxesMeta = wb._sigBoxesMeta || [];
  const JSZip = globalThis.JSZip;

  if (!JSZip || !sigBoxesMeta.length) {
    return rawBuffer;
  }

  try {
    const zip = await JSZip.loadAsync(rawBuffer);
    const drawingXml = buildDrawingXml(sigBoxesMeta);
    zip.file("xl/drawings/drawing1.xml", drawingXml);

    // 1. Link drawing1.xml in xl/worksheets/sheet1.xml
    const sheetPath = "xl/worksheets/sheet1.xml";
    const sheetFile = zip.file(sheetPath);
    if (sheetFile) {
      let sheetXml = await sheetFile.async("string");
      if (sheetXml.indexOf("xmlns:r=") === -1) {
        sheetXml = sheetXml.replace(
          "<worksheet ",
          '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        );
      }
      if (sheetXml.indexOf("<drawing ") === -1) {
        const drawingTag = '<drawing r:id="rIdSigDrawing1"/>';
        if (sheetXml.indexOf("<extLst") !== -1) {
          sheetXml = sheetXml.replace("<extLst", drawingTag + "<extLst");
        } else {
          sheetXml = sheetXml.replace("</worksheet>", drawingTag + "</worksheet>");
        }
      }
      zip.file(sheetPath, sheetXml);
    }

    // 2. Add relationship in xl/worksheets/_rels/sheet1.xml.rels
    const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
    const relTag = '<Relationship Id="rIdSigDrawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>';
    const relsFile = zip.file(relsPath);
    if (relsFile) {
      let relsXml = await relsFile.async("string");
      if (relsXml.indexOf('Id="rIdSigDrawing1"') === -1) {
        relsXml = relsXml.replace("</Relationships>", relTag + "</Relationships>");
        zip.file(relsPath, relsXml);
      }
    } else {
      zip.file(
        relsPath,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          relTag +
        '</Relationships>'
      );
    }

    // 3. Register content type in [Content_Types].xml
    const ctPath = "[Content_Types].xml";
    const ctFile = zip.file(ctPath);
    if (ctFile) {
      let ctXml = await ctFile.async("string");
      if (ctXml.indexOf("/xl/drawings/drawing1.xml") === -1) {
        const overrideTag = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
        ctXml = ctXml.replace("</Types>", overrideTag + "</Types>");
        zip.file(ctPath, ctXml);
      }
    }

    return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  } catch (err) {
    console.warn("Could not attach floating textboxes, returning standard workbook buffer:", err);
    return rawBuffer;
  }
}


/* ==========================================================================
   SANSATION FONT LOADER FOR PDF (jsPDF VFS)
   Fetches Sansation-Regular.ttf and Sansation-Bold.ttf once, caches their
   Base64 strings, and registers "Sansation" ("normal" & "bold") on jsPDF.
   ========================================================================== */
const SANSATION_FONT_URLS = {
  normal: [
    "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/sansation/Sansation-Regular.ttf",
    "https://raw.githubusercontent.com/google/fonts/main/ofl/sansation/Sansation-Regular.ttf"
  ],
  bold: [
    "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/sansation/Sansation-Bold.ttf",
    "https://raw.githubusercontent.com/google/fonts/main/ofl/sansation/Sansation-Bold.ttf"
  ]
};
const fontBase64Cache = { normal: null, bold: null };

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  const chunks = [];
  for (let i = 0; i < bytes.length; i += chunkSize) {
    chunks.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize)));
  }
  return btoa(chunks.join(""));
}

async function fetchFontBase64(urls) {
  for (let i = 0; i < urls.length; i++) {
    try {
      const res = await fetch(urls[i], { mode: "cors", cache: "force-cache" });
      if (res.ok) {
        const buf = await res.arrayBuffer();
        if (buf && buf.byteLength > 1000) return arrayBufferToBase64(buf);
      }
    } catch (e) { /* try next mirror */ }
  }
  return null;
}

async function registerSansationFont(doc) {
  try {
    if (!fontBase64Cache.normal) {
      const loaded = await Promise.all([
        fetchFontBase64(SANSATION_FONT_URLS.normal),
        fetchFontBase64(SANSATION_FONT_URLS.bold)
      ]);
      fontBase64Cache.normal = loaded[0];
      fontBase64Cache.bold = loaded[1] || loaded[0];
    }
    if (fontBase64Cache.normal) {
      doc.addFileToVFS("Sansation-Regular.ttf", fontBase64Cache.normal);
      doc.addFont("Sansation-Regular.ttf", "Sansation", "normal");
      doc.addFileToVFS("Sansation-Bold.ttf", fontBase64Cache.bold || fontBase64Cache.normal);
      doc.addFont("Sansation-Bold.ttf", "Sansation", "bold");
      return "Sansation";
    }
  } catch (e) {
    console.warn("Could not register Sansation font for PDF, using fallback:", e);
  }
  return "helvetica";
}


/* ==========================================================================
   BUILD & DOWNLOAD FORMATTED PDF
   Uses jsPDF + jspdf-autotable with the Sansation font to create an A4 PDF:
     - Top header: "<glNo> - <glName>" (e.g. "393630 - S/A - PREMISES")
     - Header summary table (5 rows)
     - Schedule items table with total row
     - Amount in words & difference table
     - Certification statement (omitted automatically when a difference exists)
     - Single outer merged box with 4 inner signature text boxes
     - Footer: "Prepared by: <officerName> (<serviceNo>)" only
   ========================================================================== */
export async function buildAndSavePdf(params, fileName) {
  const jspdfNS = globalThis.jspdf;
  if (!jspdfNS || !jspdfNS.jsPDF) {
    throw new Error("PDF library (jsPDF) did not load. Check your internet connection.");
  }

  const doc = new jspdfNS.jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pdfFont = await registerSansationFont(doc);
  const pageWidth = doc.internal.pageSize.getWidth();   // 210
  const pageHeight = doc.internal.pageSize.getHeight(); // 297
  const margin = 14;
  const contentWidth = pageWidth - margin * 2;          // 182

  params.blocks.forEach(function (block, idx) {
    if (idx > 0) doc.addPage();

    const isZero = block.mode !== "balance";
    const items = isZero ? [] : block.items;
    const tbValue = block.tbCents / 100;
    const glValue = block.glCents / 100;
    const tbFormatted = formatMoney(isZero ? 0 : block.tbCents);
    const glFormatted = formatMoney(isZero ? 0 : block.glCents);
    const tbWords = isZero ? "Zero Only" : numberToWords(tbValue);
    const glWords = isZero ? "Zero Only" : numberToWords(glValue);
    const diffText = isZero ? "NIL" : (block.diffText || "Nil");

    let curY = margin;

    // ---- Top Banner: "<glNo> - <glName>" (e.g. "393630 - S/A - PREMISES") ----
    doc.setFillColor(15, 92, 99); // #0F5C63 ledger
    doc.rect(margin, curY, contentWidth, 8.5, "F");
    doc.setFont(pdfFont, "bold");
    doc.setFontSize(10.5);
    doc.setTextColor(255, 255, 255);
    doc.text(block.glNo + " - " + block.glName, margin + 3.5, curY + 5.7);
    curY += 10.5;

    // ---- 1. Header Details Table (5 rows) ----
    doc.autoTable({
      startY: curY,
      margin: { left: margin, right: margin },
      theme: "grid",
      styles: {
        font: pdfFont,
        fontSize: 9.5,
        cellPadding: { top: 1.8, right: 2.5, bottom: 1.8, left: 2.5 },
        textColor: [22, 35, 46],
        lineColor: [0, 0, 0],
        lineWidth: 0.2
      },
      columnStyles: {
        0: { cellWidth: 68, fontStyle: "bold", fillColor: [244, 247, 248] },
        1: { cellWidth: contentWidth - 68 }
      },
      body: [
        ["Name and Service Number:", params.officerName + " - " + params.serviceNo],
        ["GL Account Number and Name:", { content: block.glNo + " - " + block.glName, styles: { fontStyle: "bold" } }],
        ["Trial Balance Date As at:", params.tbDate],
        ["General Ledger Balance :", { content: glFormatted, styles: { fontStyle: "bold" } }],
        ["Balance of Trial :", { content: tbFormatted, styles: { fontStyle: "bold" } }]
      ]
    });
    curY = doc.lastAutoTable.finalY + 3.5;

    // ---- 2. Schedule Items Table ----
    const itemRows = items.length
      ? items.map(function (it, i) {
          return [
            String(i + 1),
            it.date || "",
            String(it.ref == null ? "" : it.ref),
            it.desc || "",
            formatMoney(toCents(it.amount || 0))
          ];
        })
      : [["-", "-", "-", isZero ? "Zero Balance Account (No Schedule Items)" : "No items in Trial Balance book", "0.00"]];

    doc.autoTable({
      startY: curY,
      margin: { left: margin, right: margin },
      theme: "grid",
      head: [["Serial No", "Date", "Reference", "Description", "Amount"]],
      body: itemRows,
      foot: [[
        { content: "Total Balance of Trial Balance", colSpan: 4, styles: { halign: "center", fontStyle: "bold", fillColor: [238, 242, 244], textColor: [22, 35, 46] } },
        { content: tbFormatted, styles: { halign: "right", fontStyle: "bold", fillColor: [238, 242, 244], textColor: [22, 35, 46] } }
      ]],
      showFoot: "lastPage",
      styles: {
        font: pdfFont,
        fontSize: 9,
        cellPadding: { top: 1.8, right: 2.2, bottom: 1.8, left: 2.2 },
        textColor: [22, 35, 46],
        lineColor: [0, 0, 0],
        lineWidth: 0.2,
        valign: "middle"
      },
      headStyles: {
        font: pdfFont,
        fillColor: [225, 240, 239],
        textColor: [10, 65, 71],
        fontStyle: "bold",
        halign: "center",
        lineColor: [0, 0, 0],
        lineWidth: 0.2
      },
      footStyles: {
        font: pdfFont
      },
      columnStyles: {
        0: { cellWidth: 18, halign: "center" },
        1: { cellWidth: 25, halign: "left" },
        2: { cellWidth: 34, halign: "left" },
        3: { cellWidth: contentWidth - 18 - 25 - 34 - 32, halign: "left" },
        4: { cellWidth: 32, halign: "right" }
      }
    });
    curY = doc.lastAutoTable.finalY + 4;

    // ---- 3. Amount in Words & Difference Table ----
    doc.autoTable({
      startY: curY,
      margin: { left: margin, right: margin },
      theme: "grid",
      styles: {
        font: pdfFont,
        fontSize: 9,
        cellPadding: { top: 1.8, right: 2.5, bottom: 1.8, left: 2.5 },
        textColor: [22, 35, 46],
        lineColor: [0, 0, 0],
        lineWidth: 0.2
      },
      columnStyles: {
        0: { cellWidth: 72, fontStyle: "bold", fillColor: [244, 247, 248] },
        1: { cellWidth: contentWidth - 72 }
      },
      body: [
        ["Trial Balance Amount in Words", tbWords],
        ["General Ledger Account Balance in Words :", glWords],
        ["Difference (If Any):", diffText]
      ]
    });
    curY = doc.lastAutoTable.finalY + 4;

    // ---- Ensure Certification (if tallied) + Single Merged Signature Box fit on current page ----
    const neededHeight = block.hasDiscrepancy ? 50 : 64;
    if (curY + neededHeight > pageHeight - margin) {
      doc.addPage();
      curY = margin;
    }

    // ---- 4. Certification Statement (ONLY when GL Balance and Trial Balance tally) ----
    if (!block.hasDiscrepancy) {
      const certText =
        "We hereby certify that the total balance of GL Account No." + block.glNo +
        " has been duly verified and tallied with the total amount of the corresponding Trial Balance as at " + params.tbDate;
      doc.setFont(pdfFont, "normal");
      doc.setFontSize(9.5);
      doc.setTextColor(22, 35, 46);
      const splitCert = doc.splitTextToSize(certText, contentWidth - 2);
      doc.text(splitCert, margin + 1, curY + 4);
      curY += Math.max(11, splitCert.length * 4.5 + 3);
    } else {
      curY += 2;
    }

    // ---- 5. Single Merged Signature Box with 4 Inner Text Boxes ----
    const outerH = 46;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(0.3);
    doc.setFillColor(255, 255, 255);
    doc.rect(margin, curY, contentWidth, outerH, "FD");

    // 4 inner text boxes placed inside the single merged box
    const padX = 3.5;
    const padY = 3.5;
    const gapX = 5;
    const gapY = 3.5;
    const boxW = (contentWidth - padX * 2 - gapX) / 2;
    const boxH = (outerH - padY * 2 - gapY) / 2;

    const innerBoxes = [
      { x: margin + padX,               y: curY + padY,               text: "Prepared by : " + params.officerName },
      { x: margin + padX + boxW + gapX, y: curY + padY,               text: "Checked by : Operations manager" },
      { x: margin + padX,               y: curY + padY + boxH + gapY, text: "Checked by : Manager" },
      { x: margin + padX + boxW + gapX, y: curY + padY + boxH + gapY, text: "Audit officer :" }
    ];

    doc.setLineWidth(0.22);
    doc.setFont(pdfFont, "normal");
    doc.setFontSize(9.5);
    doc.setTextColor(0, 0, 0);

    innerBoxes.forEach(function (b) {
      doc.setFillColor(255, 255, 255);
      doc.rect(b.x, b.y, boxW, boxH, "FD");
      const lines = doc.splitTextToSize(b.text, boxW - 5);
      doc.text(lines, b.x + 2.5, b.y + 4.8);
    });
  });

  // ---- Footer: ONLY "Prepared by: <officerName> (<serviceNo>)" ----
  const totalPages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPages; p++) {
    doc.setPage(p);
    doc.setFont(pdfFont, "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(74, 91, 104);
    doc.text(
      "Prepared by: " + params.officerName + " (" + params.serviceNo + ")",
      margin,
      pageHeight - 6
    );
  }

  doc.save(fileName);
}
