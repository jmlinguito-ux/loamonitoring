import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const workspaceDir = 'C:\\Users\\Jinguito\\Desktop\\CSV\\LOA';
const SKILL_DIR = 'C:\\Users\\Jinguito\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.909.12148\\skills\\presentations';
const TMP_DIR = path.join(workspaceDir, '.codex-build-loa-intro');
const FINAL_PPTX = path.join(workspaceDir, 'output', 'LOA-Monitoring-Tool-Introduction.pptx');
const RUNTIME_PYTHON = 'C:\\Users\\Jinguito\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe';
const FONT = 'Arial';

const C = {
  navy: '#102A43', navy2: '#183B56', cyan: '#0B91C8', cyan2: '#28B5D7',
  ink: '#21384F', muted: '#637C92', pale: '#F2F7FA', line: '#D6E3EC', white: '#FFFFFF',
  green: '#159A66', amber: '#C67A10', lightCyan: '#DDF3FA', lightGreen: '#E8F7F1',
};
const pres = Presentation.create({ slideSize: { width: 1280, height: 720 } });

function addText(slide, value, x, y, w, h, options = {}) {
  const s = slide.shapes.add({
    geometry: 'textbox', position: { left: x, top: y, width: w, height: h },
    fill: 'none', line: { fill: 'none', width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: FONT, fontSize: options.size ?? 24, color: options.color ?? C.ink,
    bold: options.bold ?? false, alignment: options.align ?? 'left',
    verticalAlignment: options.valign ?? 'middle', wrap: 'square',
    autoFit: 'shrinkText', insets: { left: 0, right: 0, top: 0, bottom: 0 },
  };
  return s;
}

function addRect(slide, x, y, w, h, fill, radius = 0, lineFill = 'none', lineWidth = 0) {
  return slide.shapes.add({
    geometry: radius ? 'roundRect' : 'rect',
    position: { left: x, top: y, width: w, height: h },
    fill, line: { fill: lineFill, width: lineWidth },
    ...(radius ? { borderRadius: radius } : {}),
  });
}

function addLine(slide, x, y, w, h, color = C.line, width = 2) {
  return slide.shapes.add({ geometry: 'line', position: { left: x, top: y, width: w, height: h }, fill: 'none', line: { fill: color, width } });
}

function header(slide, title, number, dark = false) {
  const primary = dark ? C.white : C.navy;
  addText(slide, title, 72, 44, 1120, 66, { size: 43, bold: true, color: primary });
  addRect(slide, 72, 122, 56, 5, C.cyan2);
  addText(slide, 'LOA MONITORING TOOL', 72, 675, 400, 20, { size: 14, bold: true, color: dark ? '#9DC5D6' : C.muted });
  addText(slide, String(number).padStart(2, '0'), 1172, 675, 36, 20, { size: 14, bold: true, color: dark ? '#9DC5D6' : C.muted, align: 'right' });
}

function note(slide, text) { slide.speakerNotes.textFrame.setText(text); }

// 1. Cover
{
  const s = pres.slides.add(); s.background.fill = C.navy;
  addRect(s, 82, 100, 8, 500, C.cyan2);
  addText(s, 'LOA Monitoring Tool', 124, 170, 1010, 100, { size: 66, bold: true, color: C.white });
  addText(s, 'Introduction, core functions, and process flow', 128, 286, 950, 58, { size: 31, color: '#C6D8E5' });
  addLine(s, 128, 386, 730, 0, '#42647D', 2);
  addText(s, 'Land operations workspace', 128, 414, 600, 42, { size: 25, bold: true, color: C.cyan2 });
  addText(s, 'Land Master  /  Reference data  /  Entry Logs', 128, 464, 760, 38, { size: 22, color: '#C6D8E5' });
  addText(s, 'A practical guide to reviewing accounts and recording changes', 128, 606, 900, 30, { size: 18, color: '#9DB5C7' });
  note(s, 'Introduce the app as a desktop workspace for managing LOA land records, applying reference information, and reviewing auditable changes.');
}

// 2. Purpose
{
  const s = pres.slides.add(); s.background.fill = C.white; header(s, 'What the tool supports', 2);
  addText(s, 'One account workspace from current record to change history', 72, 160, 1080, 66, { size: 36, bold: true, color: C.navy });
  addText(s, 'LOA Monitoring brings land-account records together with finance and conversion references, then keeps a traceable record of saved changes.', 72, 235, 1050, 78, { size: 25, color: C.muted });
  const items = [
    ['01', 'Review accounts', 'Search and filter Land Master records.'],
    ['02', 'Apply references', 'Compare price, deductible, and conversion status.'],
    ['03', 'Track progress', 'Monitor payment, mode, and conversion summaries.'],
    ['04', 'Keep an audit trail', 'See which field changed, when, and from what value.'],
  ];
  const xs = [72, 360, 648, 936];
  items.forEach(([n, title, desc], i) => {
    const x = xs[i];
    addText(s, n, x, 385, 54, 42, { size: 30, bold: true, color: C.cyan });
    addLine(s, x, 438, 230, 0, C.line, 2);
    addText(s, title, x, 455, 245, 44, { size: 24, bold: true, color: C.navy });
    addText(s, desc, x, 505, 238, 92, { size: 21, color: C.muted, valign: 'top' });
  });
  note(s, 'The four functions describe the day-to-day value of the application without implying that all information is automatically synchronized. Users review and explicitly save staged changes.');
}

// 3. Workspace map
{
  const s = pres.slides.add(); s.background.fill = C.pale; header(s, 'Workspace map', 3);
  addText(s, 'Daily work', 72, 170, 240, 42, { size: 19, bold: true, color: C.cyan });
  const daily = [
    ['Overview', 'Account mix, payment progress, conversion progress, and recent changes.'],
    ['Land Master', 'Main account table for search, review, entry, and updates.'],
    ['Entry Logs', 'Field-level history read from the Dataverse logs table.'],
  ];
  daily.forEach(([title, desc], i) => {
    const y = 228 + i * 118;
    addText(s, title, 72, y, 300, 36, { size: 28, bold: true, color: C.navy });
    addText(s, desc, 380, y, 790, 56, { size: 23, color: C.ink });
    if (i < daily.length - 1) addLine(s, 72, y + 76, 1096, 0, C.line, 1.5);
  });
  addLine(s, 72, 590, 1096, 0, C.line, 2);
  addText(s, 'Reference & setup', 72, 605, 220, 34, { size: 19, bold: true, color: C.cyan });
  addText(s, 'Finance  ·  Conversion  ·  Settings', 320, 601, 840, 40, { size: 25, bold: true, color: C.navy });
  note(s, 'Finance and Conversion maintain reference data used in account work. Settings holds Dataverse connection and table configuration.');
}

// 4. Overview
{
  const s = pres.slides.add(); s.background.fill = C.white; header(s, 'Overview dashboard', 4);
  addText(s, 'A quick read on account mix, payment, and conversion', 72, 160, 1090, 50, { size: 32, bold: true, color: C.navy });
  const groups = [
    { y: 252, label: 'ACCOUNTS BY MODE', desc: 'See how the account portfolio is distributed across modes. Select a mode to focus the account list.' },
    { y: 350, label: 'PAYMENT PROGRESS', desc: 'Review contract value, payments deducted, outstanding balance, and settlement distribution.' },
    { y: 448, label: 'CONVERSION PROGRESS', desc: 'Track DAR status counts, secured progress, and applied area.' },
    { y: 546, label: 'WEEKLY CHANGES', desc: 'Search recent audit activity by account or changed field.' },
  ];
  groups.forEach((g, i) => {
    addText(s, g.label, 72, g.y, 265, 35, { size: 18, bold: true, color: i === 1 ? C.green : C.cyan });
    addText(s, g.desc, 360, g.y - 1, 820, 54, { size: 22, color: C.ink });
    if (i < groups.length - 1) addLine(s, 72, g.y + 65, 1096, 0, C.line, 1.2);
  });
  note(s, 'The dashboard is intended for scanning operational status. Filters can focus the related account view, and the weekly changes area provides a direct search over recent activity.');
}

// 5. Land Master and reference update rules
{
  const s = pres.slides.add(); s.background.fill = C.pale; header(s, 'Land Master account work', 5);
  addText(s, 'The account table is the working surface', 72, 160, 1030, 50, { size: 33, bold: true, color: C.navy });
  addText(s, 'Filter the list, inspect a record, then make a focused change or stage a reference update.', 72, 220, 1020, 45, { size: 23, color: C.muted });
  addText(s, 'Reference update scope', 72, 315, 360, 42, { size: 25, bold: true, color: C.cyan });
  addText(s, 'Total Contract Price\nPayment Deductible\nConversion Status', 72, 372, 420, 152, { size: 26, bold: true, color: C.ink, valign: 'top' });
  addLine(s, 548, 306, 0, 252, '#BDCFDA', 2);
  addText(s, 'Calculated financial fields', 600, 315, 520, 42, { size: 25, bold: true, color: C.green });
  addText(s, 'Balance and % Paid recalculate from price and deductible. They are saved with the account update, but excluded as separate fields in the audit trail.', 600, 372, 555, 132, { size: 23, color: C.ink, valign: 'top' });
  addRect(s, 72, 577, 1096, 2, C.line);
  addText(s, 'Changes remain staged until the user selects “Save to Dataverse”.', 72, 599, 1090, 36, { size: 23, bold: true, color: C.navy });
  note(s, 'Update from References uses the approved comparison scope: Total Contract Price, Payment Deductible, and Conversion Status. Balance and % Paid are calculated from the financial values and saved, but not listed as independent changed fields.');
}

// 6. Reference data and new entries
{
  const s = pres.slides.add(); s.background.fill = C.white; header(s, 'Reference data and new entries', 6);
  addText(s, 'Reference data', 72, 165, 470, 48, { size: 30, bold: true, color: C.navy });
  addText(s, 'Finance', 72, 238, 190, 34, { size: 23, bold: true, color: C.cyan });
  addText(s, 'Holds financial reference values used during account lookup and reference updates.', 260, 236, 370, 74, { size: 21, color: C.ink, valign: 'top' });
  addText(s, 'Conversion', 72, 336, 190, 34, { size: 23, bold: true, color: C.cyan });
  addText(s, 'Holds DAR status values used to update and monitor conversion progress.', 260, 334, 370, 74, { size: 21, color: C.ink, valign: 'top' });
  addText(s, 'Load or upload reference data, review it, then sync it to Dataverse when appropriate.', 72, 445, 550, 72, { size: 21, bold: true, color: C.muted });
  addLine(s, 660, 166, 0, 452, C.line, 2);
  addText(s, 'New account entry', 710, 165, 470, 48, { size: 30, bold: true, color: C.navy });
  addText(s, 'Enter account identifiers and details. Finance lookup can fill matching financial values.', 710, 236, 470, 82, { size: 22, color: C.ink, valign: 'top' });
  addText(s, 'Acquisition Sub-Level', 710, 342, 440, 36, { size: 23, bold: true, color: C.cyan });
  addText(s, 'Required for new records. Account Status is derived from the selected code and shown in Land Master and the entry form.', 710, 384, 470, 84, { size: 21, color: C.ink, valign: 'top' });
  addText(s, 'Saving creates the Land Master record and an Initial Entry audit log.', 710, 515, 470, 64, { size: 21, bold: true, color: C.green, valign: 'top' });
  note(s, 'The Acquisition Sub-Level database field is optional at Dataverse schema level so existing records may remain blank. The app requires it only when creating a new entry. Account Status is a derived display value, not a separate stored or logged field.');
}

// 7. Process flow, native editable elements
{
  const s = pres.slides.add(); s.background.fill = C.navy; header(s, 'End-to-end update process', 7, true);
  addText(s, 'From reference comparison to visible audit history', 72, 158, 1100, 46, { size: 29, color: '#C6D8E5' });
  const steps = [
    ['1', 'Load account data', 'Land Master from Dataverse'],
    ['2', 'Load references', 'Finance and Conversion'],
    ['3', 'Compare values', 'Price, deductible, DAR status'],
    ['4', 'Review changes', 'Changed rows are staged'],
    ['5', 'Save update', 'Master values are written'],
    ['6', 'Check Entry Logs', 'Field changes are shown'],
  ];
  const startX = 72, gap = 19, w = 174, y = 300, h = 218;
  const nodes = [];
  steps.forEach(([n, title, desc], i) => {
    const x = startX + i * (w + gap);
    const node = addRect(s, x, y, w, h, C.navy2, 14, '#3D6681', 1.5);
    nodes.push(node);
    const badge = s.shapes.add({ geometry: 'ellipse', position: { left: x + 20, top: y + 20, width: 42, height: 42 }, fill: C.cyan, line: { fill: 'none', width: 0 } });
    badge.text = n;
    badge.text.style = { typeface: FONT, fontSize: 23, bold: true, color: C.white, alignment: 'center', verticalAlignment: 'middle', insets: { left: 0, right: 0, top: 0, bottom: 0 } };
    addText(s, title, x + 20, y + 78, w - 40, 60, { size: 20, bold: true, color: C.white, valign: 'top' });
    addText(s, desc, x + 20, y + 148, w - 40, 60, { size: 18, color: '#C6D8E5', valign: 'top' });
  });
  for (let i = 0; i < nodes.length - 1; i++) {
    s.shapes.connect(nodes[i], nodes[i + 1], { kind: 'straight', fromSide: 'right', toSide: 'left', line: { fill: C.cyan2, width: 2 }, tail: { type: 'arrow', width: 'sm', length: 'sm' } });
  }
  addText(s, 'Manual edits follow the same review, save, and audit steps.', 72, 572, 1100, 40, { size: 22, bold: true, color: C.cyan2 });
  note(s, 'The process diagram is editable. If reviewing an update, compare the staged values before saving. Reference updates do not alter the Acquisition Sub-Level.');
}

// 8. Entry Logs and key controls
{
  const s = pres.slides.add(); s.background.fill = C.white; header(s, 'Entry Logs and operating controls', 8);
  addText(s, 'A field-level record of saved account changes', 72, 158, 1080, 52, { size: 33, bold: true, color: C.navy });
  addText(s, 'Entry Logs reads the Dataverse log source and presents comparisons for each changed field.', 72, 218, 1060, 42, { size: 23, color: C.muted });
  addText(s, 'Each change shows', 72, 304, 290, 36, { size: 21, bold: true, color: C.cyan });
  addText(s, 'Field name\nPrevious value\nNew value\nAction type and work week', 72, 354, 430, 180, { size: 26, bold: true, color: C.ink, valign: 'top' });
  addLine(s, 574, 296, 0, 260, C.line, 2);
  addText(s, 'Use the table safely', 626, 304, 480, 36, { size: 21, bold: true, color: C.cyan });
  addText(s, 'Select a cell and press Ctrl+C to copy its displayed value.\n\nRevert is available only for eligible field changes. Calculated Balance and % Paid are not independent revert targets.', 626, 354, 520, 178, { size: 22, color: C.ink, valign: 'top' });
  addRect(s, 72, 574, 1096, 2, C.line);
  addText(s, 'Recommended rhythm', 72, 595, 260, 36, { size: 20, bold: true, color: C.navy });
  addText(s, 'Check connection  /  review changes  /  save explicitly  /  verify the audit trail', 350, 594, 820, 38, { size: 21, color: C.ink });
  note(s, 'Entry Logs are fetched from Dataverse and held in application memory for the open session. A refresh failure should preserve the displayed results and show a Could not refresh notice. Revert affects only eligible field changes; calculated Balance and % Paid are excluded.');
}

await fs.mkdir(TMP_DIR, { recursive: true });
await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });
const candidatePath = path.join(TMP_DIR, 'candidate.pptx');
await (await PresentationFile.exportPptx(pres)).save(candidatePath);
for (let i = 0; i < pres.slides.items.length; i++) {
  const slide = pres.slides.items[i];
  const png = await pres.export({ slide, format: 'png', scale: 1 });
  await fs.writeFile(path.join(TMP_DIR, `slide-${i + 1}.png`), new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: 'layout' });
  await fs.writeFile(path.join(TMP_DIR, `slide-${i + 1}.layout.json`), await layout.text());
}

const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, 'container_tools/artifact_tool_utils.mjs')).href);
const stagingDir = path.join(TMP_DIR, 'finalizer');
await fs.mkdir(stagingDir, { recursive: true });
const result = await finalizePresentation({
  workspaceDir, candidatePath, finalPath: FINAL_PPTX, pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, 'container_tools/inspect_presentation_package_integrity.py'),
  layoutValidatorPath: path.join(SKILL_DIR, 'container_tools/inspect_presentation_layout_geometry.py'),
  layoutArgs: ['--expected-slide-size-emu', '12192000,6858000', '--validate-bullet-geometry', '--validate-heading-fit'],
  explicitTotalSlideCount: 8,
  requiredNativeTableOwnerSlides: [], requiredNativeChartOwnerSlides: [],
  fontPolicy: { basis: 'design', families: [FONT] }, verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, 'LOA-Monitoring-Tool-Introduction.validation.json'),
});
console.log(JSON.stringify({ finalPath: FINAL_PPTX, result }, null, 2));
