require('dotenv').config();

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const _ = require('lodash');
const XLSX = require('xlsx');
const Papa = require('papaparse');
const officeCrypto = require('officecrypto-tool');
const nodemailer = require('nodemailer');

const OUTPUT_DIR = path.join(__dirname, 'output');
const LOM_REP_EMAIL_PATH = path.join(__dirname, 'data', 'lom-rep-email.csv');

/** @type {Record<string, { jvcRows: object[], maRows: object[], files: Record<string, string>, password?: string }>} */
const lomData = {};

function createPrompt() {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const ask = (question) =>
    new Promise((resolve) => rl.question(question, (answer) => resolve(answer.trim())));
  return { rl, ask };
}

function normalizeEmail(email) {
  if (!email || typeof email !== 'string') return '';
  return email.trim().toLowerCase();
}

function isJciLomName(name) {
  return /^JCI .+/.test(String(name || '').trim());
}

function omitSheetJsEmptyColumns(row) {
  return _.omitBy(row, (_value, key) => /^__EMPTY(_\d+)?$/.test(key));
}

function readTextFile(filePath, label) {
  try {
    return fs.readFileSync(filePath, 'utf8');
  } catch (err) {
    throw new Error(`Cannot read ${label}: ${filePath} (${err.message})`);
  }
}

function assertRequiredColumns(rows, requiredColumns, label, filePath) {
  if (!rows.length) {
    throw new Error(`${label} is empty: ${filePath}`);
  }

  const keys = new Set(Object.keys(rows[0]));
  const missing = requiredColumns.filter((col) => !keys.has(col));
  if (missing.length) {
    throw new Error(
      `${label} is missing required column(s): ${missing.join(', ')} (${filePath})`
    );
  }
}

function readJvcXlsx(filePath) {
  let workbook;
  try {
    workbook = XLSX.readFile(filePath);
  } catch (err) {
    throw new Error(`Cannot open JVC file: ${filePath} (${err.message})`);
  }

  if (!workbook.SheetNames.length) {
    throw new Error(`JVC file has no worksheets: ${filePath}`);
  }

  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
  const normalizedRows = rows.map(omitSheetJsEmptyColumns);

  assertRequiredColumns(
    normalizedRows,
    ['Email', 'Organization Name', 'Member Type', 'User ID'],
    'JVC file',
    filePath
  );

  return normalizedRows;
}

function normalizeMaRow(row) {
  const normalized = {};
  for (const [key, value] of Object.entries(row)) {
    normalized[key.trim()] = value;
  }
  return normalized;
}

function readMaCsv(filePath) {
  const content = readTextFile(filePath, 'MA file');

  if (/^\x50\x4B/.test(content.slice(0, 2)) || content.startsWith('PK')) {
    throw new Error(
      `Cannot open MA file as CSV (file looks like Excel): ${filePath}`
    );
  }

  const parsed = Papa.parse(content, { header: true, skipEmptyLines: true });
  if (parsed.errors.length) {
    const detail = parsed.errors.map((error) => error.message).join('; ');
    throw new Error(`Cannot parse MA CSV: ${filePath} (${detail})`);
  }

  const rows = parsed.data.map(normalizeMaRow);
  assertRequiredColumns(rows, ['NOM', 'email'], 'MA file', filePath);

  return rows;
}

function applyWideColumns(worksheet, { min = 14, max = 56, padding = 3, startRow } = {}) {
  if (!worksheet['!ref']) return;

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const firstRow = startRow == null ? range.s.r : startRow;
  const cols = [];
  for (let C = range.s.c; C <= range.e.c; C++) {
    let width = min;
    for (let R = firstRow; R <= range.e.r; R++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: C })];
      if (cell == null || cell.v == null) continue;
      const len = String(cell.v).length + padding;
      if (len > width) width = Math.min(max, len);
    }
    cols.push({ wch: width });
  }
  worksheet['!cols'] = cols;
}

/** Copy column width from one header to another (e.g. LOM ← Email). */
function syncColumnWidthByHeader(worksheet, headerRowIndex, fromHeader, toHeader, { scale = 1 } = {}) {
  if (!worksheet['!cols'] || !worksheet['!ref']) return;

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  let fromIdx = -1;
  let toIdx = -1;
  for (let C = range.s.c; C <= range.e.c; C++) {
    const header = worksheet[XLSX.utils.encode_cell({ r: headerRowIndex, c: C })]?.v;
    if (header === fromHeader) fromIdx = C;
    if (header === toHeader) toIdx = C;
  }
  if (fromIdx < 0 || toIdx < 0 || !worksheet['!cols'][fromIdx]) return;
  const source = worksheet['!cols'][fromIdx];
  const wch = Math.max(8, Math.round((source.wch ?? 14) * scale));
  worksheet['!cols'][toIdx] = { ...source, wch };
}

function jsonToWideSheet(rows) {
  const worksheet = XLSX.utils.json_to_sheet(rows);
  applyWideColumns(worksheet);
  syncColumnWidthByHeader(worksheet, 0, 'Email', 'LOM', { scale: 0.5 });
  syncColumnWidthByHeader(worksheet, 0, 'email', 'LOM', { scale: 0.5 });
  return worksheet;
}

const ACCOUNT_MEMBER_TYPES = new Set(['full', 'prospective', 'senior']);

function isJvcGuest(jvcRow) {
  return String(jvcRow?.['Member Type'] ?? '').trim().toLowerCase() === 'guest';
}

function isAccountMemberType(maRow) {
  return ACCOUNT_MEMBER_TYPES.has(String(maRow?.['NOM Member Type'] ?? '').trim().toLowerCase());
}

function buildJvcEmailToRow(jvcRows) {
  const jvcEmailToRow = {};
  for (const jvcRow of jvcRows) {
    const jvcEmail = normalizeEmail(jvcRow.Email);
    if (!jvcEmail) continue;

    const existing = jvcEmailToRow[jvcEmail];
    if (!existing || (isJvcGuest(existing) && !isJvcGuest(jvcRow))) {
      jvcEmailToRow[jvcEmail] = jvcRow;
    }
  }
  return jvcEmailToRow;
}

function matchMaMemberToJvc(candidates, jvcEmailToRow, consumedJvcEmails) {
  let guestMatch = null;
  let guestEmail = null;

  for (const candidate of candidates) {
    if (consumedJvcEmails.has(candidate)) continue;

    const jvcRow = jvcEmailToRow[candidate];
    if (!jvcRow) continue;

    if (isJvcGuest(jvcRow)) {
      if (!guestMatch) {
        guestMatch = jvcRow;
        guestEmail = candidate;
      }
      continue;
    }

    consumedJvcEmails.add(candidate);
    return { jvcRow, isGuest: false };
  }

  if (guestMatch) {
    consumedJvcEmails.add(guestEmail);
    return { jvcRow: guestMatch, isGuest: true };
  }

  return null;
}

/** Find email on a JVC row whose Organization Name is not this LOM. */
function findWrongChapterJvcMatch(candidates, lomName, allJvcEmailToRows) {
  for (const candidate of candidates) {
    const rows = allJvcEmailToRows[candidate] || [];
    for (const jvcRow of rows) {
      const org = String(jvcRow['Organization Name'] ?? '').trim();
      if (org && org !== lomName) {
        return { jvcRow, email: candidate };
      }
    }
  }
  return null;
}

function buildJvcEmailToRows(jvcRows) {
  const jvcEmailToRows = {};
  for (const jvcRow of jvcRows) {
    const jvcEmail = normalizeEmail(jvcRow.Email);
    if (!jvcEmail) continue;
    if (!jvcEmailToRows[jvcEmail]) jvcEmailToRows[jvcEmail] = [];
    jvcEmailToRows[jvcEmail].push(jvcRow);
  }
  return jvcEmailToRows;
}

function formatWrongChapterRemark(wrongChapterMatch) {
  if (!wrongChapterMatch) return '';

  const org = String(wrongChapterMatch.jvcRow['Organization Name'] ?? '').trim();
  const userId = String(wrongChapterMatch.jvcRow['User ID'] ?? '').trim();
  const email = wrongChapterMatch.jvcRow.Email || wrongChapterMatch.email;
  const parts = [`JVC account email found under wrong chapter: ${org}`];
  if (email) parts.push(`email ${email}`);
  if (userId) parts.push(`User ID ${userId}`);
  return parts.join(' | ');
}

function buildAccountReportSheet(lomName, maRows, jvcRows, allJvcEmailToRows, dateStamp) {
  const accountMaRows = maRows.filter(isAccountMemberType);
  const jvcEmailToRow = buildJvcEmailToRow(jvcRows);
  const consumedJvcEmails = new Set();

  const tableRows = accountMaRows.map((row) => {
    const email = normalizeEmail(row.email);
    const email2 = normalizeEmail(row['email 2 ( backup )']);
    const candidates = [email, email2].filter(Boolean);

    const match = matchMaMemberToJvc(candidates, jvcEmailToRow, consumedJvcEmails);
    const matchedJvcRow = match?.jvcRow ?? null;
    const isGuestMatch = Boolean(match?.isGuest);

    let remarks = '';
    if (!matchedJvcRow) {
      const existsInChapter = candidates.some((candidate) => jvcEmailToRow[candidate]);
      if (!existsInChapter) {
        const wrongChapterMatch = findWrongChapterJvcMatch(
          candidates,
          lomName,
          allJvcEmailToRows
        );
        remarks = formatWrongChapterRemark(wrongChapterMatch);
      }
    }

    return {
      LOM: row.NOM || lomName,
      'Member Type': row['NOM Member Type'],
      'First Name': row['First Name'],
      'Last Name': row['Last Name'],
      Email: row.email,
      'Email 2': row['email 2 ( backup )'],
      'JVC Account Matched': matchedJvcRow ? 'Yes' : 'No',
      'Validity': !matchedJvcRow
        ? 'No'
        : isGuestMatch
          ? 'Invalid, Member Type should not be Guest'
          : 'Valid',
      'JVC Account Email': matchedJvcRow?.Email ?? '',
      'JVC User ID': matchedJvcRow ? String(matchedJvcRow['User ID'] ?? '').trim() : '',
      Remarks: remarks,
      _accountMatched: Boolean(matchedJvcRow),
      _accountOk: Boolean(matchedJvcRow && !isGuestMatch),
    };
  });

  const memCount = tableRows.length;
  const jvcMatchCount = _.sumBy(tableRows, (row) => (row._accountMatched ? 1 : 0));
  const jvcOkCount = _.sumBy(tableRows, (row) => (row._accountOk ? 1 : 0));
  const percentage =
    memCount === 0 ? 0 : _.round((jvcOkCount / memCount) * 100, 1);

  const summary = `JVC Account Report as of ${dateStamp} | Members (Full/Prospective/Senior): ${memCount} | JVC Accounts Matched: ${jvcMatchCount} | JVC Accounts Valid: ${jvcOkCount} | Valid Percentage: ${percentage}%`;

  const exportRows = accountMatchingExportRows(tableRows);
  const headerKeys = ACCOUNT_MATCHING_HEADERS;

  const aoa = [
    [summary],
    [],
    headerKeys,
    ...exportRows.map((row) => headerKeys.map((key) => row[key] ?? '')),
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(aoa);
  // Header is on row 3 (0-based index 2); skip summary line so LOM is not inflated.
  applyWideColumns(worksheet, { min: 14, max: 64, padding: 3, startRow: 2 });
  syncColumnWidthByHeader(worksheet, 2, 'Email', 'LOM', { scale: 0.5 });

  return {
    worksheet,
    memCount,
    jvcMatchCount,
    jvcOkCount,
    percentage,
    tableRows,
    exportRows,
  };
}

function buildMaRecordRows(maRows) {
  return maRows.filter(isAccountMemberType).map((row) => {
    const out = {};
    for (const [key, value] of Object.entries(row)) {
      out[key === 'NOM' ? 'LOM' : key] = value;
    }
    return out;
  });
}

const ACCOUNT_MATCHING_HEADERS = [
  'LOM',
  'Member Type',
  'First Name',
  'Last Name',
  'Email',
  'Email 2',
  'JVC Account Matched',
  'Validity',
  'JVC Account Email',
  'JVC User ID',
  'Remarks',
];

function accountMatchingExportRows(tableRows) {
  return tableRows.map(({ _accountMatched, _accountOk, ...rest }) => rest);
}

function writeLomCombinedReport(filePath, { jvcRows, maRows, accountSheet }) {
  const workbook = XLSX.utils.book_new();
  let sheetCount = 0;
  const maRecordRows = buildMaRecordRows(maRows);

  if (accountSheet) {
    XLSX.utils.book_append_sheet(workbook, accountSheet, 'JVC Summary');
    sheetCount += 1;
  }
  if (jvcRows.length) {
    XLSX.utils.book_append_sheet(workbook, jsonToWideSheet(jvcRows), 'JVC Accounts');
    sheetCount += 1;
  }
  if (maRecordRows.length) {
    XLSX.utils.book_append_sheet(workbook, jsonToWideSheet(maRecordRows), 'MA Record');
    sheetCount += 1;
  }

  if (!sheetCount) return false;

  XLSX.writeFile(workbook, filePath);
  return true;
}

function writeAllLomsMasterReport(filePath, { dateStamp, lomSummaries, matchingRows, jvcRows, maRows }) {
  const workbook = XLSX.utils.book_new();

  const summaryHeaders = [
    'LOM',
    'Members (Full/Prospective/Senior)',
    'JVC Accounts Matched',
    'JVC Accounts Valid',
    'Valid Percentage',
    'JVC Accounts (raw)',
    'MA Record',
  ];

  const totalMembers = _.sumBy(lomSummaries, 'memCount');
  const totalMatched = _.sumBy(lomSummaries, 'jvcMatchCount');
  const totalValid = _.sumBy(lomSummaries, 'jvcOkCount');
  const totalJvc = _.sumBy(lomSummaries, 'jvcRawCount');
  const totalMa = _.sumBy(lomSummaries, 'maRecordCount');
  const totalPercentage =
    totalMembers === 0 ? 0 : _.round((totalValid / totalMembers) * 100, 1);

  const summaryAoa = [
    [`All LOMs JVC Account Summary as of ${dateStamp}`],
    [],
    summaryHeaders,
    ...lomSummaries.map((row) => [
      row.lomName,
      row.memCount,
      row.jvcMatchCount,
      row.jvcOkCount,
      `${row.percentage}%`,
      row.jvcRawCount,
      row.maRecordCount,
    ]),
    [],
    [
      'TOTAL',
      totalMembers,
      totalMatched,
      totalValid,
      `${totalPercentage}%`,
      totalJvc,
      totalMa,
    ],
  ];

  const summarySheet = XLSX.utils.aoa_to_sheet(summaryAoa);
  applyWideColumns(summarySheet, { min: 12, max: 40, padding: 2, startRow: 2 });
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'LOM summaries');

  XLSX.utils.book_append_sheet(
    workbook,
    matchingRows.length
      ? jsonToWideSheet(matchingRows)
      : jsonToWideSheet([Object.fromEntries(ACCOUNT_MATCHING_HEADERS.map((h) => [h, '']))]),
    'JVC Account Matching'
  );
  XLSX.utils.book_append_sheet(
    workbook,
    jvcRows.length ? jsonToWideSheet(jvcRows) : XLSX.utils.aoa_to_sheet([['(no rows)']]),
    'JVC Accounts'
  );
  XLSX.utils.book_append_sheet(
    workbook,
    maRows.length ? jsonToWideSheet(maRows) : XLSX.utils.aoa_to_sheet([['(no rows)']]),
    'MA Record'
  );

  XLSX.writeFile(workbook, filePath);
}

const LOM_FILE_NAME_MAP = {
  'JCI Victoria': 'VJC',
  'JCI Kowloon': 'KJC',
  'JCI Island': 'IJC',
  'JCI Peninsula': 'PJC',
  'JCI Hong Kong Jayceettes': 'HKJTT',
  'JCI Lion Rock': 'LRJC',
  'JCI Harbour': 'HJC',
  'JCI Yuen Long': 'YLJC',
  'JCI Tai Ping Shan': 'TPSJC',
  'JCI Bauhinia': 'BJC',
  'JCI Dragon': 'DJC',
  'JCI East Kowloon': 'EKJC',
  'JCI City': 'CJC',
  'JCI Queensway': 'QJC',
  'JCI North District': 'NDJC',
  'JCI Ocean': 'OJC',
  'JCI Sha Tin': 'STJC',
  'JCI Apex': 'AJC',
  'JCI City Lady': 'CLJC',
  'JCI Tsuen Wan': 'TWJC',
  'JCI Lantau': 'LJC',
};

function lomFileBase(lomName) {
  return LOM_FILE_NAME_MAP[lomName] ?? lomName;
}

function formatDateStamp(date = new Date()) {
  return date.toISOString().slice(0, 10);
}

function exportLomSheets(groupedJvc, groupedMa, dateStamp) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const allLoms = _.union(Object.keys(groupedJvc), Object.keys(groupedMa)).sort();
  const allJvcRows = allLoms.flatMap((lomName) => groupedJvc[lomName] || []);
  const allJvcEmailToRows = buildJvcEmailToRows(allJvcRows);

  const lomSummaries = [];
  const allMatchingRows = [];
  const allMaRecordRows = [];

  for (const lomName of allLoms) {
    const jvcRows = groupedJvc[lomName] || [];
    const maRows = groupedMa[lomName] || [];
    const maRecordRows = buildMaRecordRows(maRows);
    const base = lomFileBase(lomName);
    const reportPath = path.join(OUTPUT_DIR, `${base} JVC Account Report ${dateStamp}.xlsx`);

    let accountResult = null;
    if (maRows.some(isAccountMemberType)) {
      accountResult = buildAccountReportSheet(
        lomName,
        maRows,
        jvcRows,
        allJvcEmailToRows,
        dateStamp
      );
    }

    const files = {};
    const wrote = writeLomCombinedReport(reportPath, {
      jvcRows,
      maRows,
      accountSheet: accountResult?.worksheet ?? null,
    });
    if (wrote) files.report = reportPath;

    lomSummaries.push({
      lomName,
      memCount: accountResult?.memCount ?? 0,
      jvcMatchCount: accountResult?.jvcMatchCount ?? 0,
      jvcOkCount: accountResult?.jvcOkCount ?? 0,
      percentage: accountResult?.percentage ?? 0,
      jvcRawCount: jvcRows.length,
      maRecordCount: maRecordRows.length,
    });

    if (accountResult?.exportRows?.length) {
      allMatchingRows.push(...accountResult.exportRows);
    }
    if (maRecordRows.length) {
      allMaRecordRows.push(...maRecordRows);
    }

    lomData[lomName] = { jvcRows, maRows, files };
  }

  const masterPath = path.join(OUTPUT_DIR, `All LOMs JVC Account Report ${dateStamp}.xlsx`);
  writeAllLomsMasterReport(masterPath, {
    dateStamp,
    lomSummaries,
    matchingRows: allMatchingRows,
    jvcRows: allJvcRows,
    maRows: allMaRecordRows,
  });
  console.log(`Master summary (not emailed): ${path.basename(masterPath)}`);

  return allLoms;
}

function generatePassword() {
  const chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  return _.times(8, () => chars[_.random(0, chars.length - 1)]).join('');
}

function protectXlsx(inputPath, outputPath, password) {
  let input;
  try {
    input = fs.readFileSync(inputPath);
  } catch (err) {
    throw new Error(`Cannot open file for password protection: ${inputPath} (${err.message})`);
  }

  let output;
  try {
    output = officeCrypto.encrypt(input, { password });
  } catch (err) {
    throw new Error(`Cannot password-protect file: ${inputPath} (${err.message})`);
  }

  fs.writeFileSync(outputPath, output);
}

function readLomRepEmails(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`LOM rep email file not found: ${filePath}`);
  }

  const content = readTextFile(filePath, 'LOM rep email file');
  const lines = content.split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) {
    throw new Error(`LOM rep email file is empty: ${filePath}`);
  }

  const map = {};
  for (const line of lines) {
    const [lom, email] = line.split(',').map((part) => part.trim());
    if (lom && email) map[lom] = email;
  }

  if (!Object.keys(map).length) {
    throw new Error(`LOM rep email file has no valid entries: ${filePath}`);
  }

  return map;
}

async function selectLomsForEmail(loms, lomRepEmails, ask) {
  const eligibleLoms = loms.filter(
    (lom) => lomRepEmails[lom] && Object.keys(lomData[lom]?.files || {}).length
  );
  const skippedNoEmail = loms.filter((lom) => !lomRepEmails[lom]);
  const skippedNoFiles = loms.filter(
    (lom) => lomRepEmails[lom] && !Object.keys(lomData[lom]?.files || {}).length
  );

  if (skippedNoEmail.length) {
    console.log(`\nNo rep email (will not be offered): ${skippedNoEmail.join(', ')}`);
  }
  if (skippedNoFiles.length) {
    console.log(`No generated files (will not be offered): ${skippedNoFiles.join(', ')}`);
  }
  if (!eligibleLoms.length) {
    console.log('\nNo LOMs eligible for email.');
    return [];
  }

  const selected = new Set(eligibleLoms);

  while (true) {
    console.log('\nSelect LOMs to send email:');
    eligibleLoms.forEach((lom, index) => {
      const mark = selected.has(lom) ? 'x' : ' ';
      console.log(`  [${mark}] ${index + 1}. ${lom} (${lomRepEmails[lom]})`);
    });
    console.log('\nToggle: enter number | a = select all | n = select none | Enter = confirm');

    const answer = await ask('> ');
    if (!answer) {
      return eligibleLoms.filter((lom) => selected.has(lom));
    }
    if (/^a(ll)?$/i.test(answer)) {
      eligibleLoms.forEach((lom) => selected.add(lom));
      continue;
    }
    if (/^n(one)?$/i.test(answer)) {
      selected.clear();
      continue;
    }

    const index = Number.parseInt(answer, 10);
    if (index >= 1 && index <= eligibleLoms.length) {
      const lom = eligibleLoms[index - 1];
      if (selected.has(lom)) selected.delete(lom);
      else selected.add(lom);
      continue;
    }

    console.log('Invalid input. Use a number, a, n, or Enter to confirm.');
  }
}

async function sendReportEmails(lomRepEmails, selectedLoms) {
  const user = process.env.GMAIL_SMTP_USER;
  const pass = process.env.GMAIL_SMTP_PW;
  if (!user || !pass) {
    throw new Error('GMAIL_SMTP_USER and GMAIL_SMTP_PW must be set in .env');
  }

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user, pass },
  });

  const selectedSet = new Set(selectedLoms);

  for (const [lomName, data] of Object.entries(lomData)) {
    if (!selectedSet.has(lomName)) continue;

    const repEmail = lomRepEmails[lomName];
    if (!repEmail) {
      console.log(`Skipping ${lomName}: no representative email`);
      continue;
    }

    if (!Object.keys(data.files).length) {
      console.log(`Skipping ${lomName}: no generated files`);
      continue;
    }

    const password = generatePassword();
    data.password = password;

    const reportPath = data.files.report;
    if (!reportPath) {
      console.log(`Skipping ${lomName}: no generated report`);
      continue;
    }

    const dir = path.dirname(reportPath);
    const ext = path.extname(reportPath);
    const baseName = path.basename(reportPath, ext);
    const protectedPath = path.join(dir, `${baseName} (protected)${ext}`);
    protectXlsx(reportPath, protectedPath, password);
    data.files.reportProtected = protectedPath;

    const attachments = [
      {
        filename: path.basename(protectedPath),
        path: protectedPath,
      },
    ];

    const subject = `${lomName} JVC Account Reports`;
    const body = [
      `Dear ${lomName} president and representative(s),`,
      '',
      `Please find attached the JVC Account Report for your chapter as of today (${new Date().toISOString().slice(0, 10)}).`,
      '',
      'The Excel file contains three tabs:',
      '- JVC Summary — account matching report (Full / Prospective / Senior)',
      '- JVC Accounts — raw JVC member list for your chapter',
      '- MA Record — MA member list (Full / Prospective / Senior)',
      '',
      `The attached Excel file is password-protected. Use this password to open it: ${password}`,
      '',
      'Best regards,',
      'Kenneth Law',
      '2026 National Digital Development Director',
      'JCI Hong Kong, China',
    ].join('\n');

    await transporter.sendMail({
      from: user,
      to: repEmail,
      subject,
      text: body,
      attachments,
    });

    console.log(`Sent report email to ${repEmail} (${lomName})`);
  }
}

async function main() {
  const args = process.argv.slice(2);
  const sendEmailFlag = args.includes('--send-email');
  const positional = args.filter((arg) => !arg.startsWith('--'));

  let jvcPath;
  let maPath;
  let lomRepPath;
  let sendAnswer;
  let rl;
  let ask;

  if (positional.length >= 2) {
    [jvcPath, maPath] = positional;
    lomRepPath = positional[2] || LOM_REP_EMAIL_PATH;
    sendAnswer = sendEmailFlag ? 'yes' : 'no';
  } else {
    const prompt = createPrompt();
    rl = prompt.rl;
    ask = prompt.ask;
    jvcPath = await ask('JVC member list path (xlsx): ');
    maPath = await ask('MA system member list path (csv): ');
    lomRepPath = await ask(`LOM rep email list path (csv) [${LOM_REP_EMAIL_PATH}]: `);
    if (!lomRepPath) lomRepPath = LOM_REP_EMAIL_PATH;
  }

  try {

    if (!fs.existsSync(jvcPath)) throw new Error(`JVC file not found: ${jvcPath}`);
    if (!fs.existsSync(maPath)) throw new Error(`MA file not found: ${maPath}`);

    console.log('\nReading files...');
    const jvcRowsAll = readJvcXlsx(jvcPath);
    const maRows = readMaCsv(maPath);
    const jvcRows = jvcRowsAll.filter((row) => isJciLomName(row['Organization Name']));

    if (!jvcRows.length) {
      throw new Error(
        `JVC file has no rows with Organization Name matching "JCI XXXXX": ${jvcPath}`
      );
    }

    const groupedJvc = _.groupBy(jvcRows, (row) => row['Organization Name'].trim());
    const groupedMa = _.groupBy(maRows, (row) => row.NOM || 'Unknown');

    delete groupedMa.Unknown;

    const jvcSkipped = jvcRowsAll.length - jvcRows.length;
    console.log(`JVC rows: ${jvcRows.length} (${jvcSkipped} ignored — not "JCI XXXXX"), MA rows: ${maRows.length}`);
    console.log(`LOMs (JVC): ${Object.keys(groupedJvc).length}, LOMs (MA): ${Object.keys(groupedMa).length}`);

    const dateStamp = formatDateStamp();
    const loms = exportLomSheets(groupedJvc, groupedMa, dateStamp);
    console.log(`\nExported reports for ${loms.length} LOM(s) to ${OUTPUT_DIR}`);

    for (const lom of loms) {
      const files = lomData[lom]?.files || {};
      const names = Object.values(files).map((f) => path.basename(f));
      if (names.length) console.log(`  ${lom}: ${names.join(', ')}`);
    }

    if (!positional.length) {
      sendAnswer = await ask('\nSend report emails? (yes/no): ');
    }
    if (/^y(es)?$/i.test(sendAnswer)) {
      const lomRepEmails = readLomRepEmails(lomRepPath);
      let selectedLoms;

      if (!positional.length) {
        selectedLoms = await selectLomsForEmail(loms, lomRepEmails, ask);
      } else {
        selectedLoms = loms.filter(
          (lom) => lomRepEmails[lom] && Object.keys(lomData[lom]?.files || {}).length
        );
      }

      if (!selectedLoms.length) {
        console.log('No LOMs selected. Skipped sending emails.');
      } else {
        console.log(`\nSending emails to ${selectedLoms.length} LOM(s)...`);
        await sendReportEmails(lomRepEmails, selectedLoms);
        console.log('Done sending emails.');
      }
    } else {
      console.log('Skipped sending emails.');
    }
  } finally {
    rl?.close();
  }
}

main().catch((err) => {
  console.error('Error:', err.message);
  process.exit(1);
});
