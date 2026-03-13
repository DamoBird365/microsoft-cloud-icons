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
import { basename, dirname, extname, join, relative } from "node:path";
import https from "node:https";

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

// ── Fluent UI CDN Brand Icons ───────────────────────────────────────

const CDN_BASE =
  "https://res.cdn.office.net/midgard/versionless/fluentui-resources/1.1.6/assets/brand-icons/product/svg";
const CDN_SIZES = [16, 24, 32, 48];

/** All available product brand icons on the Office CDN */
const CDN_PRODUCTS: { id: string; displayName: string }[] = [
  // Microsoft 365 core apps
  { id: "word", displayName: "Word" },
  { id: "excel", displayName: "Excel" },
  { id: "powerpoint", displayName: "PowerPoint" },
  { id: "outlook", displayName: "Outlook" },
  { id: "onenote", displayName: "OneNote" },
  { id: "teams", displayName: "Teams" },
  { id: "sharepoint", displayName: "SharePoint" },
  { id: "onedrive", displayName: "OneDrive" },
  { id: "access", displayName: "Access" },
  { id: "publisher", displayName: "Publisher" },
  { id: "visio", displayName: "Visio" },
  { id: "project", displayName: "Project" },
  { id: "loop", displayName: "Loop" },
  { id: "lists", displayName: "Lists" },
  { id: "forms", displayName: "Forms" },
  { id: "planner", displayName: "Planner" },
  { id: "stream", displayName: "Stream" },
  { id: "clipchamp", displayName: "Clipchamp" },
  { id: "sway", displayName: "Sway" },
  { id: "bookings", displayName: "Bookings" },
  { id: "todo", displayName: "To Do" },
  { id: "whiteboard", displayName: "Whiteboard" },
  { id: "delve", displayName: "Delve" },

  // Power Platform
  { id: "powerapps", displayName: "Power Apps" },
  { id: "powerautomate", displayName: "Power Automate" },
  { id: "powerbi", displayName: "Power BI" },
  { id: "powerpages", displayName: "Power Pages" },

  // Copilot & AI
  { id: "copilot", displayName: "Copilot" },
  { id: "designer", displayName: "Designer" },

  // Viva
  { id: "vivaconnections", displayName: "Viva Connections" },
  { id: "vivainsights", displayName: "Viva Insights" },
  { id: "vivalearning", displayName: "Viva Learning" },
  { id: "vivaengage", displayName: "Viva Engage" },
  { id: "vivapulse", displayName: "Viva Pulse" },
  { id: "vivaamplify", displayName: "Viva Amplify" },

  // Security & Compliance
  { id: "defender", displayName: "Defender" },
  { id: "purview", displayName: "Purview" },

  // Other Microsoft
  { id: "dynamics365", displayName: "Dynamics 365" },
  { id: "edge", displayName: "Edge" },
  { id: "bing", displayName: "Bing" },
  { id: "yammer", displayName: "Yammer" },
  { id: "skype", displayName: "Skype" },
  { id: "m365", displayName: "Microsoft 365" },
  { id: "office", displayName: "Office" },
  { id: "msn", displayName: "MSN" },
  { id: "familysafety", displayName: "Family Safety" },
  { id: "kaizala", displayName: "Kaizala" },
];

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
    sources: [
      "https://github.com/loryanstrant/MicrosoftCloudLogos",
      "https://res.cdn.office.net (Fluent UI Brand Icons)",
    ],
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
    `> Auto-generated from two sources:`,
    `> - [loryanstrant/MicrosoftCloudLogos](https://github.com/loryanstrant/MicrosoftCloudLogos) — community-curated SVGs (Azure, Entra, Fabric, etc.)`,
    `> - **Microsoft Office CDN** — official Fluent 2 brand icons (Word, Excel, Teams, etc.) in 4 sizes`,
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
    "Sources:",
    "- [loryanstrant/MicrosoftCloudLogos](https://github.com/loryanstrant/MicrosoftCloudLogos) — community-curated collection",
    "- [Microsoft Office CDN](https://res.cdn.office.net) — official Fluent 2 brand icons",
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

// ── Fluent UI CDN Download ───────────────────────────────────────────

function downloadFile(url: string): Promise<string | null> {
  return new Promise((resolve) => {
    https
      .get(url, (res) => {
        if (res.statusCode !== 200) {
          res.resume();
          resolve(null);
          return;
        }
        let data = "";
        res.on("data", (chunk: Buffer) => (data += chunk.toString()));
        res.on("end", () => resolve(data));
      })
      .on("error", () => resolve(null));
  });
}

async function fetchCdnBrandIcons(): Promise<IconEntry[]> {
  const icons: IconEntry[] = [];
  const category = "Fluent-Brand";

  console.log(`   Downloading from Office CDN (${CDN_PRODUCTS.length} products × ${CDN_SIZES.length} sizes)...`);

  let downloaded = 0;
  let failed = 0;

  for (const product of CDN_PRODUCTS) {
    for (const size of CDN_SIZES) {
      const filename = `${product.id}_${size}x1.svg`;
      const url = `${CDN_BASE}/${filename}`;
      const destPath = `icons/${category}/${product.displayName}/${filename}`;

      const svg = await downloadFile(url);
      if (svg) {
        const fullPath = join(process.cwd(), destPath);
        mkdirSync(dirname(fullPath), { recursive: true });
        writeFileSync(fullPath, svg);
        icons.push({
          name: `${product.displayName} (${size}px)`,
          category,
          product: product.displayName,
          path: destPath,
          sourceFile: url,
        });
        downloaded++;
      } else {
        failed++;
      }
    }
  }

  console.log(`   Downloaded ${downloaded} brand icons (${failed} unavailable)`);
  return icons;
}

// ── Run ─────────────────────────────────────────────────────────────

(async () => {
  try {
    console.log("🚀 Official Icons — Fetch & Curate\n");

    cloneSource();

    console.log("\n🔍 Scanning for latest SVGs from GitHub...");
    const repoIcons = collectSvgs();
    console.log(`   Found ${repoIcons.length} icons from loryanstrant/MicrosoftCloudLogos\n`);

    console.log("📂 Copying repo icons to icons/ directory...");
    copyIcons(repoIcons);

    console.log("\n🌐 Fetching Fluent UI brand icons from Office CDN...");
    const cdnIcons = await fetchCdnBrandIcons();

    const allIcons = [...repoIcons, ...cdnIcons];

    console.log("\n📋 Writing manifest.json...");
    writeManifest(allIcons);

    console.log("📝 Writing README.md...");
    writeReadme(allIcons);

    cleanup();

    console.log(`\n✅ Done! ${allIcons.length} SVG icons curated into icons/`);
    console.log(`   📦 ${repoIcons.length} from GitHub repo`);
    console.log(`   🌐 ${cdnIcons.length} from Office CDN (Fluent Brand)`);
    console.log("   See manifest.json for the full index.");
  } catch (error) {
    cleanup();
    console.error("❌ Failed:", error);
    process.exit(1);
  }
})();
