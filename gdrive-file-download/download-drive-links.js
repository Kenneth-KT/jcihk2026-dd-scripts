#!/usr/bin/env node

const fs = require("node:fs/promises");
const path = require("node:path");
const readline = require("node:readline/promises");

function printUsage() {
  console.log("Usage:");
  console.log('  node download-drive-links.js "<link1>,<link2>,..." [folder-name]');
  console.log("  node download-drive-links.js");
  console.log("  # then paste comma-separated links when prompted");
  console.log("");
  console.log("Or pipe input:");
  console.log('  echo "<link1>,<link2>,<link3>" | node download-drive-links.js');
  console.log("");
  console.log("Examples:");
  console.log(
    '  node download-drive-links.js "https://drive.google.com/file/d/FILE_ID/view,https://drive.google.com/open?id=ANOTHER_ID" 01-BJC1',
  );
  console.log(
    '  echo "https://drive.google.com/file/d/FILE_ID/view?usp=sharing,https://drive.google.com/open?id=ANOTHER_ID" | node download-drive-links.js',
  );
}

function safeDirName(name) {
  const cleaned = name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").trim();
  if (!cleaned || cleaned === "." || cleaned === "..") {
    throw new Error(`Invalid folder name: ${name}`);
  }
  return cleaned;
}

function parseCliArgs(argv) {
  const args = argv.slice(2);
  if (args.length === 0) {
    return { links: null, folder: null };
  }

  return {
    links: args[0],
    folder: args[1] ?? null,
  };
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

function decodeRfc5987Filename(value) {
  try {
    return decodeURIComponent(value.replace(/\+/g, " "));
  } catch {
    return value;
  }
}

// HTTP headers are Latin-1; servers often put UTF-8 bytes in filename="...".
function decodeLegacyFilename(value) {
  return Buffer.from(value, "latin1").toString("utf8");
}

function resolveFilenameFromHeaders(headers, fallback) {
  const disposition = headers.get("content-disposition");
  if (!disposition) {
    return fallback;
  }

  const utf8Match = disposition.match(/filename\*=(?:UTF-8|utf-8)''([^;]+)/i);
  if (utf8Match?.[1]) {
    return safeFilename(decodeRfc5987Filename(utf8Match[1].trim()));
  }

  const quotedMatch = disposition.match(/filename="([^"]+)"/i);
  if (quotedMatch?.[1]) {
    return safeFilename(decodeLegacyFilename(quotedMatch[1]));
  }

  const unquotedMatch = disposition.match(/filename=([^;]+)/i);
  if (unquotedMatch?.[1]) {
    return safeFilename(decodeLegacyFilename(unquotedMatch[1].trim()));
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
  const cli = parseCliArgs(process.argv);
  const rawInput = cli.links ?? (await readLinksInput());
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

  let folderName;
  try {
    folderName = cli.folder ? safeDirName(cli.folder) : getTimestampFolderName();
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
    return;
  }

  const outputDir = path.resolve(process.cwd(), folderName);
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
