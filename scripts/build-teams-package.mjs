#!/usr/bin/env node
/*
 * Build the Teams app package (.zip) for the My Apps launcher.
 *
 * Usage (run from the project folder):
 *   node scripts/build-teams-package.mjs --host <your-site-hostname> --org "<organization name>" [--version 1.0.0]
 *
 * Example:
 *   node scripts/build-teams-package.mjs --host gentle-sand-0b1c2d30f.5.azurestaticapps.net --org "Contoso IT" --version 1.0.0
 *
 * Add --option-b to build a separate TEST app ("My Apps B") whose tab tries to open My Apps
 * automatically when it loads. It gets its own app ID so you can test it next to the normal app.
 *
 * What it does, in plain words:
 *   1. Reads the template: appPackage/manifest.json
 *   2. Fills in your website hostname, your organization name and the version number.
 *   3. Checks for common mistakes (leftover placeholders, wrong icon sizes, bad version format).
 *   4. Writes dist/MyAppsTeamsApp-<version>.zip with manifest.json, color.png and outline.png
 *      at the top level of the zip (Teams requires that).
 * It needs only Node.js. It does not sign in anywhere or change anything online.
 */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { deflateRawSync } from "node:zlib";
import { createHash } from "node:crypto";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const HOST_PLACEHOLDER = "REPLACE-WITH-YOUR-HOSTNAME";
const ORG_PLACEHOLDER = "REPLACE-WITH-YOUR-ORGANIZATION";

function fail(message) {
  console.error(`\nSTOPPED: ${message}\n`);
  process.exit(1);
}

const SWITCHES = new Set(["option-b"]); // flags that need no value

function readArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const key = argv[i];
    if (!key.startsWith("--")) fail(`Unexpected value "${key}". Use --host, --org, --version and --option-b.`);
    if (SWITCHES.has(key.slice(2))) {
      args[key.slice(2)] = true;
      continue;
    }
    const value = argv[i + 1];
    if (value === undefined || value.startsWith("--")) fail(`Missing a value after ${key}.`);
    args[key.slice(2)] = value;
    i += 1;
  }
  return args;
}

const args = readArgs(process.argv.slice(2));
if (!args.host) fail("Add your website hostname, for example: --host gentle-sand-0b1c2d30f.5.azurestaticapps.net");
if (!args.org) fail('Add your organization name in quotes, for example: --org "Contoso IT"');

const host = args.host.trim().replace(/^https?:\/\//i, "").replace(/\/.*$/, "").toLowerCase();
if (!/^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/.test(host)) {
  fail(`"${args.host}" doesn't look like a hostname. Copy it from the Azure portal (Overview > URL) and leave out https://`);
}
if (host.includes("replace-with")) fail("The hostname still contains the placeholder text.");

const org = args.org.trim();
if (org.length < 1 || org.length > 32) fail("The organization name must be 1 to 32 characters long.");

const version = (args.version || "1.0.0").trim();
if (!/^\d+\.\d+\.\d+$/.test(version)) fail(`Version "${version}" must look like 1.0.0 (three numbers separated by dots).`);

let manifest;
try {
  manifest = JSON.parse(readFileSync(join(root, "appPackage", "manifest.json"), "utf8"));
} catch (error) {
  fail(`appPackage/manifest.json could not be read as JSON: ${error.message}`);
}

function fillPlaceholders(value) {
  if (typeof value === "string") return value.split(HOST_PLACEHOLDER).join(host).split(ORG_PLACEHOLDER).join(org);
  if (Array.isArray(value)) return value.map(fillPlaceholders);
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, fillPlaceholders(item)]));
  }
  return value;
}

manifest = fillPlaceholders(manifest);
manifest.version = version;

// Option B test build: a separate app (its own ID and name) whose tab opens My Apps without a click.
// The ID is worked out from the main ID, so rebuilding Option B always gives the same ID.
function derivedGuid(baseId, label) {
  const hex = createHash("sha1").update(`${String(baseId).toLowerCase()}:${label}`).digest("hex").slice(0, 32).split("");
  hex[12] = "5";
  hex[16] = ((parseInt(hex[16], 16) & 0x3) | 0x8).toString(16);
  const h = hex.join("");
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20, 32)}`;
}

const optionB = Boolean(args["option-b"]);
if (optionB) {
  manifest.id = derivedGuid(manifest.id, "option-b");
  manifest.name.short = "My Apps B";
  manifest.name.full = "My Apps launcher (Option B automatic test)";
  manifest.description.short = "Test build: tries to open My Apps as soon as the tab loads.";
  for (const staticTab of manifest.staticTabs || []) {
    const address = new URL(staticTab.contentUrl);
    address.searchParams.set("launch", "auto");
    staticTab.contentUrl = address.toString();
  }
}

function pngInfo(buffer) {
  if (buffer.subarray(0, 8).toString("hex") !== "89504e470d0a1a0a") return null;
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20), colorType: buffer[25] };
}

const problems = [];
if (JSON.stringify(manifest).includes("REPLACE-WITH")) problems.push("A placeholder is still in the manifest.");
if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(manifest.id || "")) problems.push('The "id" is not a valid GUID.');
const tab = (manifest.staticTabs || [])[0];
if (!tab) {
  problems.push("No personal tab (staticTabs) was found.");
} else {
  for (const field of ["contentUrl", "websiteUrl"]) {
    let url;
    try {
      url = new URL(tab[field]);
    } catch (error) {
      problems.push(`${field} is not a valid web address.`);
      continue;
    }
    if (url.protocol !== "https:") problems.push(`${field} must start with https://`);
    if (url.hostname !== host) problems.push(`${field} points to ${url.hostname}, but your hostname is ${host}.`);
  }
  if (!(tab.scopes || []).includes("personal")) problems.push('The tab scope must include "personal".');
}
if (!(manifest.validDomains || []).includes(host)) problems.push("validDomains must include your hostname.");
for (const field of ["websiteUrl", "privacyUrl", "termsOfUseUrl"]) {
  if (!/^https:\/\//.test((manifest.developer || {})[field] || "")) problems.push(`developer.${field} must start with https://`);
}
if (((manifest.name || {}).short || "").length > 30) problems.push("name.short is longer than 30 characters.");
if (((manifest.description || {}).short || "").length > 80) problems.push("description.short is longer than 80 characters.");
if (((manifest.description || {}).full || "").length > 4000) problems.push("description.full is longer than 4000 characters.");
if (!/^#[0-9a-fA-F]{6}$/.test(manifest.accentColor || "")) problems.push("accentColor must look like #7B1FA2.");

const colorPng = readFileSync(join(root, "appPackage", manifest.icons.color));
const outlinePng = readFileSync(join(root, "appPackage", manifest.icons.outline));
const colorInfo = pngInfo(colorPng);
const outlineInfo = pngInfo(outlinePng);
if (!colorInfo || colorInfo.width !== 192 || colorInfo.height !== 192) problems.push("color.png must be a 192 x 192 PNG.");
if (!outlineInfo || outlineInfo.width !== 32 || outlineInfo.height !== 32) problems.push("outline.png must be a 32 x 32 PNG.");
if (outlineInfo && ![4, 6].includes(outlineInfo.colorType)) problems.push("outline.png must be white on a transparent background.");

if (problems.length > 0) {
  fail(`The package was NOT built. Fix these first:\n  - ${problems.join("\n  - ")}`);
}

// ---- Tiny ZIP writer (no extra tools needed) ----
const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buffer) {
  let c = 0xffffffff;
  for (const byte of buffer) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function dosStamp(date) {
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    day: ((date.getFullYear() - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()
  };
}

function makeZip(files) {
  const stamp = dosStamp(new Date());
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const { name, data } of files) {
    const nameBytes = Buffer.from(name, "utf8");
    const packed = deflateRawSync(data, { level: 9 });
    const crc = crc32(data);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(stamp.time, 10);
    local.writeUInt16LE(stamp.day, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28);
    localParts.push(local, nameBytes, packed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(stamp.time, 12);
    central.writeUInt16LE(stamp.day, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(packed.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBytes.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, nameBytes);

    offset += local.length + nameBytes.length + packed.length;
  }
  const centralSize = centralParts.reduce((total, part) => total + part.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralSize, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, ...centralParts, end]);
}

const manifestBytes = Buffer.from(`${JSON.stringify(manifest, null, 2)}\n`, "utf8");
const zip = makeZip([
  { name: "manifest.json", data: manifestBytes },
  { name: "color.png", data: colorPng },
  { name: "outline.png", data: outlinePng }
]);

const distFolder = join(root, "dist");
mkdirSync(distFolder, { recursive: true });
const label = optionB ? `OptionB-${version}` : version;
const zipName = `MyAppsTeamsApp-${label}.zip`;
writeFileSync(join(distFolder, zipName), zip);
writeFileSync(join(distFolder, `manifest-${label}.json`), manifestBytes);

console.log(optionB ? "\nTeams app package built (Option B automatic-launch TEST build)." : "\nTeams app package built.");
console.log(`  File:      dist/${zipName}`);
console.log(`  App ID:    ${manifest.id}`);
console.log(`  Version:   ${version}`);
console.log(`  Website:   https://${host}/`);
console.log(`  SHA-256:   ${createHash("sha256").update(zip).digest("hex")}`);
console.log("\nNext: upload the zip in Teams (Apps > Manage your apps > Upload an app > Upload a custom app).\n");
