/**
 * fetch-icons.ts
 *
 * Clones loryanstrant/MicrosoftCloudLogos, extracts only the latest
 * official SVG icons (root-level per product folder), and organises
 * them into a clean icons/ directory with a manifest.json index.
 */

import { execSync } from "node:child_process";
import {
  cpSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { basename, extname, join, relative } from "node:path";

// ── Config ──────────────────────────────────────────────────────────

const SOURCE_REPO = "https://github.com/loryanstrant/MicrosoftCloudLogos.git";
const TEMP_DIR = join(process.cwd(), ".tmp-source");
const ICONS_DIR = join(process.cwd(), "icons");
const MANIFEST_PATH = join(process.cwd(), "manifest.json");
const README_PATH = join(process.cwd(), "README.md");

/** Top-level folders in the source repo to process */
const CATEGORIES = [
  "Azure",
  "Copilot (not M365)",
  "Dynamics 365",
  "Entra",
  "Fabric",
  "Microsoft 365",
  "Power Platform",
  "Viva",
  "other",
];

/** Folders to skip entirely */
const SKIP_FOLDERS = new Set([
  "zzLEGACY logos",
  "docs",
  ".github",
  ".devcontainer",
  ".git",
]);

/** Sanitise folder name for filesystem (no spaces/special chars) */
function sanitise(name: string): string {
  return name
    .replace(/\(not M365\)/g, "")
    .replace(/\s+/g, "-")
    .replace(/[()[\]]/g, "")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
}

/** Check if a subfolder name is a year-range prefix */
function parseYearRange(name: string): { start: number; end: number } | null {
  const match = name.match(/^(\d{4})[-–](\d{4})/);
  if (!match) return null;
  return { start: parseInt(match[1]), end: parseInt(match[2]) };
}

/** Check if a subfolder should be skipped */
function shouldSkipSubfolder(name: string): boolean {
  // Known renamed-product subfolders (old branding kept under new product)
  const renamedProducts = new Set([
    "Power Virtual Agents",
    "Azure Active Directory",
    "MyAnalytics",
    "Workplace Analytics",
    "Yammer",
    "Office",
    "Skype for Business",
  ]);
  if (renamedProducts.has(name)) return true;

  // Year-range subfolders: skip if end year is before 2025
  const range = parseYearRange(name);
  if (range && range.end < 2025) return true;

  return false;
}

/** Check if a filename has a year-range prefix (e.g. "2020-2025 PowerApps_scalable.svg") */
function stripYearPrefix(filename: string): string {
  return filename.replace(/^\d{4}[-–]\d{4}\s+/, "");
}

// ── Types ───────────────────────────────────────────────────────────

interface IconEntry {
  name: string;
  category: string;
  product: string | null;
  path: string;
  sourceFile: string;
}

// ── Main ────────────────────────────────────────────────────────────

function cloneSource(): void {
  if (existsSync(TEMP_DIR)) {
    console.log("🗑️  Cleaning previous temp clone...");
    rmSync(TEMP_DIR, { recursive: true, force: true });
  }

  console.log("📦 Shallow-cloning source repo...");
  execSync(`git clone --depth 1 "${SOURCE_REPO}" "${TEMP_DIR}"`, {
    stdio: "inherit",
  });
}

function collectSvgs(): IconEntry[] {
  const icons: IconEntry[] = [];

  for (const category of CATEGORIES) {
    const categoryDir = join(TEMP_DIR, category);
    if (!existsSync(categoryDir)) {
      console.warn(`⚠️  Category folder not found: ${category}`);
      continue;
    }

    const sanitisedCategory = sanitise(category);

    // Collect root-level SVGs directly in the category folder
    const rootEntries = readdirSync(categoryDir);
    for (const entry of rootEntries) {
      const fullPath = join(categoryDir, entry);
      const stat = statSync(fullPath);

      if (stat.isFile() && extname(entry).toLowerCase() === ".svg") {
        const cleanName = stripYearPrefix(entry);
        icons.push({
          name: basename(cleanName, ".svg"),
          category: sanitisedCategory,
          product: null,
          path: `icons/${sanitisedCategory}/${cleanName}`,
          sourceFile: relative(TEMP_DIR, fullPath).replace(/\\/g, "/"),
        });
      }
    }

    // Collect SVGs from product subfolders
    for (const entry of rootEntries) {
      const fullPath = join(categoryDir, entry);
      const stat = statSync(fullPath);

      if (!stat.isDirectory()) continue;
      if (SKIP_FOLDERS.has(entry)) continue;
      if (shouldSkipSubfolder(entry)) continue;

      // If this is a year-range subfolder at category level (not a product), skip
      // (we already picked up root-level SVGs above)
      const isYearRange = parseYearRange(entry) !== null;
      if (isYearRange) continue;

      const sanitisedProduct = sanitise(entry);
      const productEntries = readdirSync(fullPath);

      // First: collect root-level SVGs from the product folder
      for (const file of productEntries) {
        const filePath = join(fullPath, file);
        const fileStat = statSync(filePath);

        if (fileStat.isFile() && extname(file).toLowerCase() === ".svg") {
          const cleanName = stripYearPrefix(file);
          icons.push({
            name: basename(cleanName, ".svg"),
            category: sanitisedCategory,
            product: sanitisedProduct,
            path: `icons/${sanitisedCategory}/${sanitisedProduct}/${cleanName}`,
            sourceFile: relative(TEMP_DIR, filePath).replace(/\\/g, "/"),
          });
        }
      }

      // Second: collect SVGs from current-era year-range subfolders
      for (const subEntry of productEntries) {
        const subPath = join(fullPath, subEntry);
        const subStat = statSync(subPath);
        if (!subStat.isDirectory()) continue;
        if (shouldSkipSubfolder(subEntry)) continue;

        const range = parseYearRange(subEntry);
        if (!range) continue; // not a year-range folder — skip (could be a renamed product)

        // Include SVGs from this current-era subfolder
        const subFiles = readdirSync(subPath);
        for (const file of subFiles) {
          const filePath = join(subPath, file);
          const fileStat = statSync(filePath);

          if (fileStat.isFile() && extname(file).toLowerCase() === ".svg") {
            icons.push({
              name: basename(file, ".svg"),
              category: sanitisedCategory,
              product: sanitisedProduct,
              path: `icons/${sanitisedCategory}/${sanitisedProduct}/${file}`,
              sourceFile: relative(TEMP_DIR, filePath).replace(/\\/g, "/"),
            });
          }
        }
      }
    }
  }

  // Deduplicate: if we have both a year-prefixed and non-prefixed version
  // with the same cleaned name, prefer the non-prefixed (latest) one
  const seen = new Map<string, IconEntry>();
  for (const icon of icons) {
    const key = icon.path;
    if (!seen.has(key)) {
      seen.set(key, icon);
    }
  }

  return Array.from(seen.values());
}

function copyIcons(icons: IconEntry[]): void {
  // Clean existing icons
  if (existsSync(ICONS_DIR)) {
    rmSync(ICONS_DIR, { recursive: true, force: true });
  }

  for (const icon of icons) {
    const destPath = join(process.cwd(), icon.path);
    const destDir = join(destPath, "..");
    mkdirSync(destDir, { recursive: true });

    const sourcePath = join(TEMP_DIR, icon.sourceFile);
    cpSync(sourcePath, destPath);
  }
}

function writeManifest(icons: IconEntry[]): void {
  const manifest = {
    generated: new Date().toISOString(),
    source: "https://github.com/loryanstrant/MicrosoftCloudLogos",
    totalIcons: icons.length,
    categories: {} as Record<string, { count: number; products: Record<string, number> }>,
    icons: icons.map((i) => ({
      name: i.name,
      category: i.category,
      product: i.product,
      path: i.path,
    })),
  };

  for (const icon of icons) {
    if (!manifest.categories[icon.category]) {
      manifest.categories[icon.category] = { count: 0, products: {} };
    }
    manifest.categories[icon.category].count++;
    const prod = icon.product || "(root)";
    manifest.categories[icon.category].products[prod] =
      (manifest.categories[icon.category].products[prod] || 0) + 1;
  }

  writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2));
}

function writeReadme(icons: IconEntry[]): void {
  const grouped = new Map<string, Map<string | null, IconEntry[]>>();

  for (const icon of icons) {
    if (!grouped.has(icon.category)) grouped.set(icon.category, new Map());
    const catMap = grouped.get(icon.category)!;
    if (!catMap.has(icon.product)) catMap.set(icon.product, []);
    catMap.get(icon.product)!.push(icon);
  }

  const lines: string[] = [
    "# Official Microsoft Cloud Icons (SVG)",
    "",
    "A curated collection of the **latest official SVG icons** for Microsoft Cloud products.",
    "",
    `> Auto-generated from [loryanstrant/MicrosoftCloudLogos](https://github.com/loryanstrant/MicrosoftCloudLogos) — only current, root-level SVGs are included. No legacy, no PNGs, no JPGs.`,
    "",
    `**${icons.length} icons** across **${grouped.size} categories**`,
    "",
    "## Quick Reference",
    "",
    "| Category | Icons |",
    "|----------|-------|",
  ];

  for (const [category, products] of grouped) {
    let count = 0;
    for (const entries of products.values()) count += entries.length;
    lines.push(`| ${category} | ${count} |`);
  }

  lines.push("", "## Icons by Category", "");

  for (const [category, products] of grouped) {
    lines.push(`### ${category}`, "");

    // Root-level icons (no product)
    const rootIcons = products.get(null);
    if (rootIcons) {
      for (const icon of rootIcons.sort((a, b) => a.name.localeCompare(b.name))) {
        lines.push(`- \`${icon.name}\` — [\`${icon.path}\`](${icon.path})`);
      }
      lines.push("");
    }

    // Product folders
    for (const [product, entries] of products) {
      if (product === null) continue;
      lines.push(`#### ${product}`, "");
      for (const icon of entries.sort((a, b) => a.name.localeCompare(b.name))) {
        lines.push(`- \`${icon.name}\` — [\`${icon.path}\`](${icon.path})`);
      }
      lines.push("");
    }
  }

  lines.push(
    "---",
    "",
    "## Refresh Icons",
    "",
    "```bash",
    "npm run fetch",
    "```",
    "",
    "This re-clones the source repo and rebuilds the collection with only the latest SVGs.",
    "",
    "## Credits",
    "",
    "Source: [loryanstrant/MicrosoftCloudLogos](https://github.com/loryanstrant/MicrosoftCloudLogos)",
    "",
    "All logos are the property of Microsoft Corporation.",
    "",
  );

  writeFileSync(README_PATH, lines.join("\n"));
}

function cleanup(): void {
  if (existsSync(TEMP_DIR)) {
    console.log("🧹 Cleaning up temp files...");
    rmSync(TEMP_DIR, { recursive: true, force: true });
  }
}

// ── Run ─────────────────────────────────────────────────────────────

try {
  console.log("🚀 Official Icons — Fetch & Curate\n");

  cloneSource();

  console.log("\n🔍 Scanning for latest SVGs...");
  const icons = collectSvgs();
  console.log(`   Found ${icons.length} current SVG icons\n`);

  console.log("📂 Copying icons to icons/ directory...");
  copyIcons(icons);

  console.log("📋 Writing manifest.json...");
  writeManifest(icons);

  console.log("📝 Writing README.md...");
  writeReadme(icons);

  cleanup();

  console.log(`\n✅ Done! ${icons.length} SVG icons curated into icons/`);
  console.log("   See manifest.json for the full index.");
} catch (error) {
  cleanup();
  console.error("❌ Failed:", error);
  process.exit(1);
}
