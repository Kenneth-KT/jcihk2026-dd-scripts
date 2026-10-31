#!/usr/bin/env node

const fs = require("node:fs/promises");
const path = require("node:path");

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

async function runWithConcurrency(items, concurrency, worker) {
  let nextIndex = 0;
  const results = new Array(items.length);

  async function runWorker() {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await worker(items[currentIndex], currentIndex);
    }
  }

  const workerCount = Math.min(concurrency, items.length);
  await Promise.all(Array.from({ length: workerCount }, () => runWorker()));
  return results;
}

const DRIVE_URL_PREFIXES = [
  "https://drive.google.com/open?id=",
  "http://drive.google.com/open?id=",
  "https://drive.google.com/file/d/",
  "http://drive.google.com/file/d/",
  "https://drive.google.com/uc?export=download&id=",
  "http://drive.google.com/uc?export=download&id=",
  "https://drive.google.com/uc?id=",
  "http://drive.google.com/uc?id=",
];

const DRIVE_FILE_ID_LENGTH = 33;

function matchDriveUrlPrefix(source, start) {
  const rest = source.slice(start).toLowerCase();
  for (const prefix of DRIVE_URL_PREFIXES) {
    if (rest.startsWith(prefix)) {
      return {
        prefix: source.slice(start, start + prefix.length),
        end: start + prefix.length,
      };
    }
  }
  return null;
}

function readDriveFileId(source, start) {
  let end = start;
  while (end < source.length && /[A-Za-z0-9_-]/.test(source[end])) {
    end += 1;
  }

  const greedy = source.slice(start, end);
  if (greedy.length > DRIVE_FILE_ID_LENGTH) {
    return {
      id: greedy.slice(0, DRIVE_FILE_ID_LENGTH),
      end: start + DRIVE_FILE_ID_LENGTH,
    };
  }

  return { id: greedy, end };
}

function skipDriveUrlSuffix(source, start) {
  const ch = source[start];
  if (ch !== "/" && ch !== "?") {
    return start;
  }

  let i = start;
  while (i < source.length && source[i] !== "," && source[i] !== " " && source[i] !== "\t") {
    i += 1;
  }
  return i;
}

function skipSpaces(source, start) {
  let i = start;
  while (i < source.length && (source[i] === " " || source[i] === "\t")) {
    i += 1;
  }
  return i;
}

function parseInputText(text) {
  // Terminals insert wrap newlines in bulk paste. Drop them and recover
  // folder rows from the original separators:
  //   FOLDER,https://...     (no space after comma)
  //   , https://...          (space after comma = next file in the same folder)
  const source = String(text).replace(/[\r\n]/g, "");
  const records = [];
  let folder = null;
  let links = [];
  let i = 0;

  function flush() {
    if (links.length > 0) {
      records.push({ folder, links });
    }
    folder = null;
    links = [];
  }

  while (i < source.length) {
    i = skipSpaces(source, i);
    if (i >= source.length) {
      break;
    }

    const prefixMatch = matchDriveUrlPrefix(source, i);
    if (prefixMatch) {
      const fileId = readDriveFileId(source, prefixMatch.end);
      const urlEnd = skipDriveUrlSuffix(source, fileId.end);
      if (fileId.id) {
        links.push(prefixMatch.prefix + fileId.id + source.slice(fileId.end, urlEnd));
      }
      i = skipSpaces(source, urlEnd);
      if (source[i] === ",") {
        i += 1;
        continue;
      }
      flush();
      continue;
    }

    const commaAt = source.indexOf(",", i);
    if (commaAt === -1) {
      break;
    }

    flush();
    folder = source.slice(i, commaAt).trim() || null;
    i = commaAt + 1;
  }

  flush();
  return records.filter((record) => record.links.length > 0);
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

async function readPastedText() {
  console.log(
    "Paste comma-separated Google Drive public links, one line will create one folder. \nYou may put the folder name in front of the comma list, files will be put under the folder:",
  );
  process.stdout.write("> ");

  return new Promise((resolve, reject) => {
    process.stdin.setEncoding("utf8");
    let buffer = "";

    const onData = (chunk) => {
      buffer += chunk;
      if (/(?:\r?\n)[ \t]*(?:\r?\n)$/.test(buffer)) {
        cleanup();
        resolve(buffer.replace(/(?:\r?\n)[ \t]*(?:\r?\n)+$/, "\n"));
        return;
      }
      // A finished line is not the blank line that starts the download.
      // Print another prompt so it is obvious the script is still waiting.
      if (/(?:\r?\n)$/.test(buffer)) {
        process.stdout.write("> ");
      }
    };

    const onEnd = () => {
      cleanup();
      resolve(buffer);
    };

    const onError = (error) => {
      cleanup();
      reject(error);
    };

    function cleanup() {
      process.stdin.off("data", onData);
      process.stdin.off("end", onEnd);
      process.stdin.off("error", onError);
      // A resumed TTY stdin keeps the event loop alive after main() returns.
      process.stdin.pause();
      process.stdin.unref();
    }

    process.stdin.on("data", onData);
    process.stdin.on("end", onEnd);
    process.stdin.on("error", onError);
  });
}

async function readInputText() {
  if (process.stdin.isTTY) {
    return readPastedText();
  }

  return readStdin();
}

function buildDownloadTasks(records, cliFolder) {
  const defaultFolderName = cliFolder ? safeDirName(cliFolder) : getTimestampFolderName();
  const tasks = [];

  for (const parsed of records) {
    if (!parsed.links.length) {
      continue;
    }

    const folderName = parsed.folder ? safeDirName(parsed.folder) : defaultFolderName;
    const outputDir = path.resolve(process.cwd(), folderName);

    for (let i = 0; i < parsed.links.length; i += 1) {
      tasks.push({
        link: parsed.links[i],
        outputDir,
        fileIndex: i + 1,
      });
    }
  }

  return tasks;
}

async function main() {
  const cli = parseCliArgs(process.argv);
  const rawText = cli.links ?? (await readInputText());
  const records = parseInputText(rawText ?? "");

  if (records.length === 0) {
    console.error("No input received.");
    printUsage();
    process.exitCode = 1;
    return;
  }

  let tasks;
  try {
    tasks = buildDownloadTasks(records, cli.folder);
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
    return;
  }

  if (tasks.length === 0) {
    console.error("No links found in input.");
    process.exitCode = 1;
    return;
  }

  const outputDirs = [...new Set(tasks.map((task) => task.outputDir))];
  await Promise.all(outputDirs.map((dir) => fs.mkdir(dir, { recursive: true })));

  console.log("");
  console.log(`Downloading ${tasks.length} files`);

  const results = await runWithConcurrency(tasks, 5, async (task) => {
    const { link, outputDir, fileIndex } = task;
    const fileId = extractDriveFileId(link);

    if (!fileId) {
      console.error(`[${fileIndex}] Skipped: invalid Google Drive file link -> ${link}`);
      return { success: false };
    }

    try {
      const result = await downloadDriveFile(fileId, outputDir, fileIndex);
      console.log(
        `[${fileIndex}] Downloaded: ${result.fileName} (${result.size.toLocaleString()} bytes)`,
      );
      return { success: true };
    } catch (error) {
      console.error(`[${fileIndex}] Failed (${fileId}): ${error.message}`);
      return { success: false };
    }
  });

  const successCount = results.filter((result) => result.success).length;
  const failCount = results.length - successCount;

  console.log("");
  console.log(`Completed. Success: ${successCount}, Failed: ${failCount}`);
  for (const dir of outputDirs) {
    console.log(`Output folder: ${dir}`);
  }

  if (failCount > 0) {
    process.exitCode = 2;
  }
}

if (require.main === module) {
  main().catch((error) => {
    console.error("Unexpected error:", error);
    process.exitCode = 1;
  });
} else {
  module.exports = { parseInputText, extractDriveFileId };
}
