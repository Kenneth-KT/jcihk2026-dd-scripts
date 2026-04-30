#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const MEMBER_TYPE_TAG_MAP = {
  Senior: "MEM_SM",
  Prospective: "MEM_PM",
  Full: "MEM_FM",
};

const LOM_TAG_MAP = {
  "JCI Victoria": "LOM_VJC",
  "JCI Kowloon": "LOM_KJC",
  "JCI Island": "LOM_IJC",
  "JCI Peninsula": "LOM_PJC",
  "JCI Hong Kong Jayceettes": "LOM_HKJTT",
  "JCI Lion Rock": "LOM_LRJC",
  "JCI Harbour": "LOM_HJC",
  "JCI Yuen Long": "LOM_YLJC",
  "JCI Tai Ping Shan": "LOM_TPSJC",
  "JCI Bauhinia": "LOM_BJC",
  "JCI Dragon": "LOM_DJC",
  "JCI East Kowloon": "LOM_EKJC",
  "JCI City": "LOM_CJC",
  "JCI Queensway": "LOM_QJC",
  "JCI North District": "LOM_NDJC",
  "JCI Ocean": "LOM_OJC",
  "JCI Sha Tin": "LOM_STJC",
  "JCI Apex": "LOM_AJC",
  "JCI City Lady": "LOM_CLJC",
  "JCI Tsuen Wan": "LOM_TWJC",
  "JCI Lantau": "LOM_LTJC",
};

function parseCSV(content) {
  const rows = [];
  let row = [];
  let field = "";
  let i = 0;
  let inQuotes = false;

  while (i < content.length) {
    const char = content[i];

    if (inQuotes) {
      if (char === '"') {
        if (content[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += char;
      }
    } else if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (char === "\r") {
      // Ignore CR in CRLF.
    } else {
      field += char;
    }

    i += 1;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

function escapeCSV(value) {
  if (value == null) {
    return "";
  }

  const text = String(value);
  if (/[",\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function getValue(record, key) {
  return (record[key] || "").trim();
}

function getMemberType(record, rowNumber) {
  const nomType = getValue(record, "NOM Member Type");
  if (!nomType) {
    throw new Error(`Row ${rowNumber}: missing NOM Member Type`);
  }
  return nomType;
}

function main() {
  const inputArg = process.argv[2] || "members.csv";
  const outputArg = process.argv[3] || "mailchimp.csv";
  const yearArg = process.argv[4] || process.env.MAILCHIMP_TAG_YEAR || "2026";
  const summaryArg = process.argv[5] || "lom-summary.csv";
  const tagYear = String(yearArg).trim();

  const inputPath = path.resolve(process.cwd(), inputArg);
  const outputPath = path.resolve(process.cwd(), outputArg);
  const summaryPath = path.resolve(process.cwd(), summaryArg);

  if (!/^\d{4}$/.test(tagYear)) {
    throw new Error(`Invalid tag year "${tagYear}". Expected 4 digits, e.g. 2026`);
  }

  if (!fs.existsSync(inputPath)) {
    throw new Error(`Input file not found: ${inputPath}`);
  }

  const content = fs.readFileSync(inputPath, "utf8");
  const rows = parseCSV(content);

  if (rows.length === 0) {
    throw new Error("Input CSV is empty");
  }

  const headers = rows[0].map((h) => h.trim());
  const requiredHeaders = ["NOM", "First Name", "Last Name", "email"];
  for (const requiredHeader of requiredHeaders) {
    if (!headers.includes(requiredHeader)) {
      throw new Error(`Missing required column: ${requiredHeader}`);
    }
  }

  if (!headers.includes("NOM Member Type")) {
    throw new Error("Missing required column: NOM Member Type");
  }

  const dataRows = rows.slice(1).filter((r) => r.some((v) => v.trim() !== ""));
  const records = dataRows.map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index] || "";
    });
    return record;
  });

  const outputRows = [["Email", "First Name", "Last Name", "Tags"]];
  const summaryByLom = {};

  records.forEach((record, index) => {
    const rowNumber = index + 2;
    const email = getValue(record, "email");
    const firstName = getValue(record, "First Name");
    const lastName = getValue(record, "Last Name");
    const lomName = getValue(record, "NOM");
    const memberType = getMemberType(record, rowNumber);

    if (!lomName) {
      throw new Error(`Row ${rowNumber}: missing NOM (LOM name)`);
    }

    const memberTag = MEMBER_TYPE_TAG_MAP[memberType];
    if (!memberTag) {
      throw new Error(`Row ${rowNumber}: unmapped member type "${memberType}"`);
    }

    const lomTag = LOM_TAG_MAP[lomName];
    if (!lomTag) {
      throw new Error(`Row ${rowNumber}: unmapped LOM "${lomName}"`);
    }

    const typeShort = memberTag.replace("MEM_", "");
    if (!summaryByLom[lomName]) {
      summaryByLom[lomName] = { SM: 0, PM: 0, FM: 0, Total: 0 };
    }
    if (!Object.prototype.hasOwnProperty.call(summaryByLom[lomName], typeShort)) {
      throw new Error(`Row ${rowNumber}: unexpected member tag "${memberTag}"`);
    }
    summaryByLom[lomName][typeShort] += 1;
    summaryByLom[lomName].Total += 1;

    const tags = `${tagYear}_MEM,${tagYear}_${memberTag},${tagYear}_${lomTag}`;
    outputRows.push([email, firstName, lastName, tags]);
  });

  const output = outputRows
    .map((row) => row.map((v) => escapeCSV(v)).join(","))
    .join("\n");

  const summaryRows = [["LOM", "SM", "PM", "FM", "Total"]];
  const loms = Object.keys(summaryByLom).sort((a, b) => a.localeCompare(b));
  const grandTotal = { SM: 0, PM: 0, FM: 0, Total: 0 };
  loms.forEach((lomName) => {
    const counts = summaryByLom[lomName];
    summaryRows.push([lomName, counts.SM, counts.PM, counts.FM, counts.Total]);
    grandTotal.SM += counts.SM;
    grandTotal.PM += counts.PM;
    grandTotal.FM += counts.FM;
    grandTotal.Total += counts.Total;
  });
  summaryRows.push(["ALL", grandTotal.SM, grandTotal.PM, grandTotal.FM, grandTotal.Total]);
  const summaryOutput = summaryRows
    .map((row) => row.map((v) => escapeCSV(v)).join(","))
    .join("\n");

  fs.writeFileSync(outputPath, `${output}\n`, "utf8");
  fs.writeFileSync(summaryPath, `${summaryOutput}\n`, "utf8");
  console.log(
    `Wrote ${outputRows.length - 1} rows to ${outputPath} (tag year: ${tagYear})`
  );
  console.log(`Wrote LOM summary to ${summaryPath}`);
}

try {
  main();
} catch (error) {
  console.error(error.message);
  process.exit(1);
}
