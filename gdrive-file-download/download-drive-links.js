#!/usr/bin/env node

const fs = require("node:fs/promises");
const path = require("node:path");
const readline = require("node:readline/promises");

function printUsage() {
  console.log("Usage:");
  console.log("  node download-drive-links.js");
  console.log("  # then paste comma-separated links when prompted");
  console.log("");
  console.log("Or pipe input:");
  console.log('  echo "<link1>,<link2>,<link3>" | node download-drive-links.js');
  console.log("");
  console.log("Example:");
  console.log(
    '  echo "https://drive.google.com/file/d/FILE_ID/view?usp=sharing,https://drive.google.com/open?id=ANOTHER_ID" | node download-drive-links.js',
  );
}

function getTimestampFolderName() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const min = String(now.getMinutes()).padStart(2, "0");
  const sec = String(now.getSeconds()).padStart(2, "0");
  return `downloads-${yyyy}${mm}${dd}-${hh}${min}${sec}`;
}

function extractDriveFileId(link) {
  try {
    const url = new URL(link.trim());
    if (!url.hostname.includes("drive.google.com")) {
      return null;
    }

    const openId = url.searchParams.get("id");
    if (openId) {
      return openId;
    }

    const pathMatch = url.pathname.match(/\/file\/d\/([^/]+)/);
    if (pathMatch && pathMatch[1]) {
      return pathMatch[1];
    }

    const ucId = url.searchParams.get("id");
    if (url.pathname.includes("/uc") && ucId) {
      return ucId;
    }

    return null;
  } catch {
    return null;
  }
}

function safeFilename(name) {
  return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").trim() || "downloaded-file";
}

function resolveFilenameFromHeaders(headers, fallback) {
  const disposition = headers.get("content-disposition");
  if (!disposition) {
    return fallback;
  }

  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match && utf8Match[1]) {
    try {
      return safeFilename(decodeURIComponent(utf8Match[1]));
    } catch {
      return safeFilename(utf8Match[1]);
    }
  }

  const basicMatch = disposition.match(/filename="?([^"]+)"?/i);
  if (basicMatch && basicMatch[1]) {
    return safeFilename(basicMatch[1]);
  }

  return fallback;
}

function getIndexedFilename(originalName, index) {
  const indexedPrefix = `${String(index).padStart(2, "0")}-`;
  return safeFilename(`${indexedPrefix}${originalName}`);
}

async function downloadDriveFile(fileId, outputDir, index) {
  const downloadUrl = `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`;
  const response = await fetch(downloadUrl, { redirect: "follow" });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}`);
  }

  const contentType = (response.headers.get("content-type") || "").toLowerCase();
  if (contentType.includes("text/html")) {
    throw new Error("Received HTML page instead of file. Link may not be publicly downloadable.");
  }

  const fallbackName = `file-${fileId}`;
  const resolvedName = resolveFilenameFromHeaders(response.headers, fallbackName);
  const finalName = getIndexedFilename(resolvedName, index);
  const outputPath = path.join(outputDir, finalName);

  const arrayBuffer = await response.arrayBuffer();
  await fs.writeFile(outputPath, Buffer.from(arrayBuffer));

  return {
    fileName: finalName,
    size: arrayBuffer.byteLength,
  };
}

function parseLinksArg(rawArg) {
  return rawArg
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function readStdin() {
  return new Promise((resolve, reject) => {
    const chunks = [];
    process.stdin.setEncoding("utf8");
    process.stdin.on("data", (chunk) => chunks.push(chunk));
    process.stdin.on("end", () => resolve(chunks.join("")));
    process.stdin.on("error", reject);
  });
}

async function readLinksInput() {
  if (process.stdin.isTTY) {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    try {
      const answer = await rl.question(
        "Paste comma-separated Google Drive public links, then press Enter:\n> ",
      );
      return answer.trim();
    } finally {
      rl.close();
    }
  }

  return (await readStdin()).trim();
}

async function main() {
  const rawInput = await readLinksInput();
  if (!rawInput) {
    console.error("No input received.");
    printUsage();
    process.exitCode = 1;
    return;
  }

  const links = parseLinksArg(rawInput);
  if (links.length === 0) {
    console.error("No links found in input.");
    process.exitCode = 1;
    return;
  }

  const outputDir = path.resolve(process.cwd(), getTimestampFolderName());
  await fs.mkdir(outputDir, { recursive: true });

  console.log(`Downloading ${links.length} file(s) into: ${outputDir}`);

  let successCount = 0;
  let failCount = 0;

  for (let i = 0; i < links.length; i += 1) {
    const link = links[i];
    const displayIndex = i + 1;

    const fileId = extractDriveFileId(link);
    if (!fileId) {
      console.error(`[${displayIndex}] Skipped: invalid Google Drive file link -> ${link}`);
      failCount += 1;
      continue;
    }

    try {
      const result = await downloadDriveFile(fileId, outputDir, displayIndex);
      console.log(
        `[${displayIndex}] Downloaded: ${result.fileName} (${result.size.toLocaleString()} bytes)`,
      );
      successCount += 1;
    } catch (error) {
      console.error(`[${displayIndex}] Failed (${fileId}): ${error.message}`);
      failCount += 1;
    }
  }

  console.log("");
  console.log(`Completed. Success: ${successCount}, Failed: ${failCount}`);
  console.log(`Output folder: ${outputDir}`);

  if (failCount > 0) {
    process.exitCode = 2;
  }
}

main().catch((error) => {
  console.error("Unexpected error:", error);
  process.exitCode = 1;
});
