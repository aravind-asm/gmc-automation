// GMC report email summary — an Office Script (Excel on the web → Automate tab).
//
// What it does: opens the report workbook the app just dropped, reads the figures already on the
// Overview, Spike Radar and Top 10 Issues sheets, and returns the email body as HTML. Power
// Automate runs it with "Excel Online (Business) → Run script" and puts the result in the body
// of "Send an email (V2)". It only reads: the workbook is never changed.
//
// It finds things by their label text ("TOTAL PRODUCTS", a header row starting "Region"), not by
// fixed cell positions, so small layout changes don't break it. If a sheet's labels or column
// headings are renamed in backend/app/report.py, update the matching text here.
//
// This copy is kept in the repo so it can be reviewed and restored. The live copy is the one
// saved in Excel (OneDrive → Documents → Office Scripts); edit both together.

// --- Wording: edit freely ---------------------------------------------------------------------
const GREETING = "Hi Team,";
const SIGN_OFF = "Please refer to the attached report for full details.<br>Best regards,";
const TOP_N = 5; // rows in the "biggest drops" and "key issues" tables

// Signature under "Best regards,". The first line is bold. The Stagwell Media logo goes under it,
// taken from the report itself (the wide picture at the top of the Overview sheet).
// Replace these lines with your own signature (INSTRUCTIONS.md step 10).
const SIGNATURE_LINES = [
  "Your Name",
  "Your job title",
  "M. +00-0000000000",
  `<a href="mailto:you@example.com" style="color:#1F4E79;">you@example.com</a>`,
  `<a href="https://example.com" style="color:#1F4E79;">example.com</a>`,
];
const LOGO_WIDTH_PX = 200; // shown size in the email; the height follows

// --- Colours: the same as the report's status colours ------------------------------------------
const STATUS_COLOURS: { [status: string]: string } = {
  Critical: "#9C0006",
  Warning: "#9C3B00",
  Minor: "#9C5700",
};
// The status definitions under table ② also name Healthy, in the report's green. Kept apart from
// STATUS_COLOURS, which decides which accounts are underperforming.
const DEFINITION_COLOURS: { [status: string]: string } = {
  Healthy: "#006100",
  Critical: "#9C0006",
  Warning: "#9C3B00",
  Minor: "#9C5700",
};
const NOTE_TEXT = "#404040"; // the report's note text colour
// Tables ②, ③ and ④ share one width so their right edges line up every day, whatever names are
// in them (ADR-011 amendment 2026-11-23). Table ③ with its Account ID column needs ~935 px with
// today's names and ~960 px with the longest account names, so 1000 px leaves room.
const TABLE_WIDTH_PX = 1000;
// One font and size for every piece of text (operator: uniform 10 pt). Each paragraph, list item and
// table cell carries it itself, because Outlook desktop does not pass a font down from the outer box
// and would show its own default size instead.
const FONT = "font-family:'Segoe UI',Calibri,Arial,sans-serif;font-size:10pt;color:#1A1A1A;";
const CELL = "border:1px solid #BFBFBF;padding:3px 8px;";
const HEADING = `${FONT}font-weight:bold;margin:18px 0 6px 0;`;

type Cell = string | number | boolean;

function main(workbook: ExcelScript.Workbook): string {
  const overview = sheetValues(workbook, "Overview");
  const radar = sheetValues(workbook, "Spike Radar");
  const issues = sheetValues(workbook, "Top 10 Issues");

  const parts: string[] = [];
  parts.push(`<div style="${FONT}">`);
  parts.push(`<p style="${FONT}margin:0 0 12px 0;">${GREETING}<br>Here is the Daily Global GMC Feed Governance Summary as of ${esc(generatedAt(overview))}.</p>`);
  parts.push(globalOverview(overview, radar));
  parts.push(underperformingAccounts(overview));
  parts.push(biggestDrops(radar));
  parts.push(keyIssues(issues));
  parts.push(`<p style="${FONT}margin:18px 0 0 0;">${SIGN_OFF}</p>`);
  parts.push(signature(workbook));
  parts.push("</div>");
  return parts.join("\n");
}

// Name and details, then the logo. If the logo can't be read, the email still goes, without it.
function signature(workbook: ExcelScript.Workbook): string {
  const lines = SIGNATURE_LINES.map((line, i) => (i === 0 ? `<b>${line}</b>` : line)).join("<br>");
  let html = `<p style="${FONT}margin:12px 0 8px 0;">${lines}</p>`;
  const logo = logoBase64(workbook);
  if (logo) {
    html += `<img src="data:image/png;base64,${logo}" width="${LOGO_WIDTH_PX}" alt="Stagwell Media"`
      + ` style="display:block;border:0;">`;
  }
  return html;
}

// The Stagwell Media logo is the widest-for-its-height picture on the Overview sheet (the LG logo
// beside it is square).
function logoBase64(workbook: ExcelScript.Workbook): string {
  try {
    const sheet = workbook.getWorksheet("Overview");
    if (!sheet) {
      return "";
    }
    // Office Scripts doesn't allow Excel calls inside sort() (and is fussy about other array
    // callbacks), so a plain loop reads each shape once.
    let widest: ExcelScript.Shape | undefined = undefined;
    let widestRatio = 0;
    for (const shape of sheet.getShapes()) {
      if (shape.getType() !== ExcelScript.ShapeType.image) {
        continue;
      }
      const ratio = shape.getWidth() / shape.getHeight();
      if (ratio > widestRatio) {
        widest = shape;
        widestRatio = ratio;
      }
    }
    return widest ? widest.getImageAsBase64(ExcelScript.PictureFormat.png) : "";
  } catch (e) {
    console.log(`Logo not added: ${e}`);
    return "";
  }
}

// ① Global overview: the Overview cards, plus the Spike Radar "Approval change" card.
function globalOverview(overview: Cell[][], radar: Cell[][]): string {
  const total = cardValue(overview, "TOTAL PRODUCTS");
  const disapproved = cardValue(overview, "DISAPPROVED PRODUCTS");
  const average = cardValue(overview, "AVERAGE APPROVAL RATE");
  const change = cardValue(radar, "APPROVAL CHANGE");
  const earlier = comparedWith(radar);

  let averageText = typeof average === "number" ? pct(average) : "—";
  if (typeof change === "number" && earlier) {
    averageText += ` (${pp(change)} vs. ${esc(earlier)})`;
  }
  return heading("&#9312; Global Overview")
    + `<ul style="${FONT}margin:0;">`
    + `<li style="${FONT}">Total Products: ${num(total)} (Disapproved: ${num(disapproved)})</li>`
    + `<li style="${FONT}">Avg. Approval Rate: ${averageText}</li>`
    + `</ul>`;
}

// ② Every account whose Status is Critical, Warning or Minor, worst approval rate first, then the
// status definitions (copied from the report, so the two never disagree).
function underperformingAccounts(overview: Cell[][]): string {
  const title = heading("&#9313; Underperforming Accounts (status Critical, Warning or Minor)");
  const table = findTable(overview, ["Region", "Country", "Account ID"]);
  if (!table) {
    return title + note("No account table in this report.");
  }
  const rows = table.rows
    .filter((r) => STATUS_COLOURS[String(r["Status"])] !== undefined)
    .sort((a, b) => Number(a["Approval rate"]) - Number(b["Approval rate"]));
  if (rows.length === 0) {
    return title + note("None — every account is Healthy.");
  }
  return title + htmlTable(
    ["Region", "Country", "Account ID", "Account", "Total products", "Disapproved", "Approval rate", "Status"],
    rows.map((r) => [
      esc(r["Region"]),
      esc(r["Country"]),
      esc(r["Account ID"]),
      esc(r["Account"]),
      num(r["Total products"]),
      num(r["Disapproved"]),
      pct(r["Approval rate"]),
      `<b style="color:${STATUS_COLOURS[String(r["Status"])]};">${esc(r["Status"])}</b>`,
    ]),
    [4, 5, 6],
  ) + statusDefinitions(overview);
}

// The Overview sheet's status line ("Healthy = approval rate above 95.0%  ·  Critical = …"), with
// each status name in its bold colour as in the report. Empty if the line can't be found: the email
// still goes, without it.
function statusDefinitions(overview: Cell[][]): string {
  let line = "";
  for (const row of overview) {
    for (const cell of row) {
      if (String(cell).indexOf("Healthy = ") === 0) {
        line = String(cell);
      }
    }
  }
  if (!line) {
    return "";
  }
  const parts = line.split("  ·  ").map((part) => {
    const m = part.match(/^(Healthy|Critical|Warning|Minor)( = .*)$/);
    return m
      ? `<b style="color:${DEFINITION_COLOURS[m[1]]};">${m[1]}</b>${esc(m[2])}`
      : esc(part);
  });
  // As wide as the table above it, so the line wraps under the table rather than past it.
  return `<p style="${FONT}color:${NOTE_TEXT};margin:6px 0 0 0;width:${TABLE_WIDTH_PX}px;max-width:${TABLE_WIDTH_PX}px;">`
    + `${parts.join("&nbsp; &middot; &nbsp;")}</p>`;
}

// ③ The accounts whose approval rate fell the most since the earlier day (Spike Radar sheet).
function biggestDrops(radar: Cell[][]): string {
  const earlier = comparedWith(radar);
  const title = heading(`&#9314; Top ${TOP_N} Biggest Approval Rate Drops`
    + (earlier ? ` (vs. ${esc(earlier)})` : ""));
  const table = findTable(radar, ["Region", "Country", "Account ID"]);
  if (!table) {
    return title + note("Not available yet: the comparison needs a sync on an earlier day (Korea time).");
  }
  const drops = table.rows
    .filter((r) => typeof r["Change (pp)"] === "number" && (r["Change (pp)"] as number) < 0)
    .sort((a, b) => (a["Change (pp)"] as number) - (b["Change (pp)"] as number))
    .slice(0, TOP_N);
  if (drops.length === 0) {
    return title + note("No account's approval rate fell.");
  }
  return title + htmlTable(
    ["Region", "Country", "Account ID", "Account", "Total products", "Disapproved now", "Approval before",
      "Approval now", "Change (pp)"],
    drops.map((r) => [
      esc(r["Region"]),
      esc(r["Country"]),
      esc(r["Account ID"]),
      esc(r["Account"]),
      num(r["Total products"]),
      num(r["Disapproved now"]),
      pct(r["Approval before"]),
      pct(r["Approval now"]),
      pp(r["Change (pp)"]),
    ]),
    [4, 5, 6, 7, 8],
  );
}

// ④ The first rows of the Top 10 Issues sheet (already ranked by products affected).
function keyIssues(issues: Cell[][]): string {
  const title = heading(`&#9315; Top ${TOP_N} Key Issues Summary`);
  const table = findTable(issues, ["Rank", "Issue", "Attribute"]);
  if (!table || table.rows.length === 0) {
    return title + note("No disapproved products in this sync.");
  }
  return title + htmlTable(
    ["Issue", "Attribute", "Issue code", "Products", "Countries", "Accounts"],
    table.rows.slice(0, TOP_N).map((r) => [
      esc(r["Issue"]) + (r["Tag"] ? ` <span style="color:#7F7F7F;">(pending review)</span>` : ""),
      esc(r["Attribute"]),
      esc(r["Issue code"]),
      num(r["Products"]),
      num(r["Countries"]),
      num(r["Accounts"]),
    ]),
    [3, 4, 5],
  );
}

// --- Reading the workbook ---------------------------------------------------------------------

function sheetValues(workbook: ExcelScript.Workbook, name: string): Cell[][] {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) {
    return [];
  }
  const used = sheet.getUsedRange(true);
  return used ? (used.getValues() as Cell[][]) : [];
}

// A card is a label cell with its value in the cell directly below it.
function cardValue(values: Cell[][], label: string): Cell | undefined {
  for (let r = 0; r < values.length - 1; r++) {
    const c = values[r].indexOf(label);
    if (c >= 0) {
      return values[r + 1][c];
    }
  }
  return undefined;
}

// A table is a header row starting with the given headings; its rows run until the first
// blank first cell. Each row comes back keyed by its column heading.
function findTable(values: Cell[][], firstHeaders: string[]):
  { rows: { [header: string]: Cell }[] } | undefined {
  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const start = row.indexOf(firstHeaders[0]);
    if (start < 0 || !firstHeaders.every((h, i) => row[start + i] === h)) {
      continue;
    }
    const headers = row.slice(start).map((h) => String(h));
    const rows: { [header: string]: Cell }[] = [];
    for (let d = r + 1; d < values.length && values[d][start] !== ""; d++) {
      const record: { [header: string]: Cell } = {};
      headers.forEach((h, i) => {
        if (h) {
          record[h] = values[d][start + i];
        }
      });
      rows.push(record);
    }
    return { rows };
  }
  return undefined;
}

// "2026-11-16 13:21 KST" from the "Generated: …" line under the title.
function generatedAt(values: Cell[][]): string {
  for (const row of values) {
    for (const cell of row) {
      const m = String(cell).match(/^Generated: (\d{4}-\d{2}-\d{2} \d{2}:\d{2} KST)/);
      if (m) {
        return m[1];
      }
    }
  }
  return "the latest sync";
}

// "Fri 13 Nov 2026" from the Spike Radar subtitle "… compared with Fri 13 Nov 2026 (3 days earlier)".
function comparedWith(values: Cell[][]): string | undefined {
  for (const row of values) {
    for (const cell of row) {
      const m = String(cell).match(/compared with (.+?) \(/);
      if (m) {
        return m[1];
      }
    }
  }
  return undefined;
}

// --- Writing HTML -----------------------------------------------------------------------------

function heading(text: string): string {
  return `<p style="${HEADING}">${text}</p>`;
}

function note(text: string): string {
  return `<p style="${FONT}margin:0;color:#595959;">${esc(text)}</p>`;
}

// Inline styles only: Outlook ignores <style> blocks. Columns listed in `rightAligned` are numbers.
function htmlTable(headers: string[], rows: string[][], rightAligned: number[]): string {
  const align = (i: number) => (rightAligned.indexOf(i) >= 0 ? "text-align:right;" : "text-align:left;");
  const head = headers
    .map((h, i) => `<th style="${FONT}${CELL}${align(i)}background:#F2F2F2;white-space:nowrap;">${h}</th>`)
    .join("");
  const body = rows
    .map((r) => "<tr>" + r.map((v, i) => `<td style="${FONT}${CELL}${align(i)}">${v}</td>`).join("") + "</tr>")
    .join("\n");
  // Both the width attribute (Outlook desktop) and the style (web and phone mail) set the width.
  return `<table cellspacing="0" cellpadding="0" width="${TABLE_WIDTH_PX}"`
    + ` style="border-collapse:collapse;width:${TABLE_WIDTH_PX}px;${FONT}">`
    + `<tr>${head}</tr>\n${body}</table>`;
}

function esc(value: Cell | undefined): string {
  return String(value === undefined ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// 17441 → "17,441"
function num(value: Cell | undefined): string {
  if (typeof value !== "number") {
    return esc(value);
  }
  return String(Math.round(value)).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// 0.9446 → "94.5%" (one decimal, as in the report)
function pct(value: Cell | undefined): string {
  return typeof value === "number" ? `${(value * 100).toFixed(1)}%` : esc(value);
}

// -0.3 → "-0.3 pp", 1.2 → "+1.2 pp"
function pp(value: Cell | undefined): string {
  if (typeof value !== "number") {
    return esc(value);
  }
  return `${value > 0 ? "+" : ""}${value.toFixed(1)} pp`;
}
