/**
 * fetch-icons.ts
 *
 * Fetches curated Microsoft product/application SVG icons from two sources:
 *   1. Microsoft Office CDN (Fluent 2 brand icons) — preferred
 *   2. loryanstrant/MicrosoftCloudLogos GitHub repo — fallback
 *
 * Outputs a clean, flat folder structure with consistent kebab-case naming.
 */

import { execSync } from "node:child_process";
import {
  cpSync,
  existsSync,
  mkdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join } from "node:path";
import https from "node:https";

// ── Config ──────────────────────────────────────────────────────────

const SOURCE_REPO = "https://github.com/loryanstrant/MicrosoftCloudLogos.git";
const TEMP_DIR = join(process.cwd(), ".tmp-source");
const ICONS_DIR = join(process.cwd(), "icons");
const MANIFEST_PATH = join(process.cwd(), "manifest.json");
const README_PATH = join(process.cwd(), "README.md");
const CDN_BASE =
  "https://res.cdn.office.net/midgard/versionless/fluentui-resources/1.1.6/assets/brand-icons/product/svg";
const CDN_SIZE = 48;

// ── Product Catalogue ───────────────────────────────────────────────

interface ProductDef {
  /** kebab-case filename (without .svg) */
  id: string;
  /** Human-readable name */
  displayName: string;
  /** Category folder */
  category: string;
  /** CDN product id (for Fluent Brand download), or null */
  cdnId: string | null;
  /** Path within the GitHub repo (fallback), or null */
  repoPath: string | null;
}

const PRODUCTS: ProductDef[] = [
  // ── Microsoft 365 ──
  { id: "word", displayName: "Word", category: "microsoft-365", cdnId: "word", repoPath: null },
  { id: "excel", displayName: "Excel", category: "microsoft-365", cdnId: "excel", repoPath: null },
  { id: "powerpoint", displayName: "PowerPoint", category: "microsoft-365", cdnId: "powerpoint", repoPath: null },
  { id: "outlook", displayName: "Outlook", category: "microsoft-365", cdnId: "outlook", repoPath: null },
  { id: "onenote", displayName: "OneNote", category: "microsoft-365", cdnId: "onenote", repoPath: null },
  { id: "teams", displayName: "Teams", category: "microsoft-365", cdnId: "teams", repoPath: null },
  { id: "sharepoint", displayName: "SharePoint", category: "microsoft-365", cdnId: "sharepoint", repoPath: null },
  { id: "onedrive", displayName: "OneDrive", category: "microsoft-365", cdnId: "onedrive", repoPath: null },
  { id: "access", displayName: "Access", category: "microsoft-365", cdnId: "access", repoPath: null },
  { id: "publisher", displayName: "Publisher", category: "microsoft-365", cdnId: "publisher", repoPath: null },
  { id: "visio", displayName: "Visio", category: "microsoft-365", cdnId: "visio", repoPath: null },
  { id: "project", displayName: "Project", category: "microsoft-365", cdnId: "project", repoPath: null },
  { id: "loop", displayName: "Loop", category: "microsoft-365", cdnId: "loop", repoPath: null },
  { id: "lists", displayName: "Lists", category: "microsoft-365", cdnId: "lists", repoPath: null },
  { id: "forms", displayName: "Forms", category: "microsoft-365", cdnId: "forms", repoPath: null },
  { id: "planner", displayName: "Planner", category: "microsoft-365", cdnId: "planner", repoPath: null },
  { id: "stream", displayName: "Stream", category: "microsoft-365", cdnId: "stream", repoPath: null },
  { id: "clipchamp", displayName: "Clipchamp", category: "microsoft-365", cdnId: "clipchamp", repoPath: null },
  { id: "sway", displayName: "Sway", category: "microsoft-365", cdnId: "sway", repoPath: null },
  { id: "bookings", displayName: "Bookings", category: "microsoft-365", cdnId: "bookings", repoPath: null },
  { id: "to-do", displayName: "To Do", category: "microsoft-365", cdnId: "todo", repoPath: "Microsoft 365/To Do/To_Do.svg" },
  { id: "whiteboard", displayName: "Whiteboard", category: "microsoft-365", cdnId: "whiteboard", repoPath: null },
  { id: "delve", displayName: "Delve", category: "microsoft-365", cdnId: "delve", repoPath: null },
  { id: "places", displayName: "Places", category: "microsoft-365", cdnId: null, repoPath: "Microsoft 365/Places/Microsoft Places.svg" },
  { id: "microsoft-365", displayName: "Microsoft 365", category: "microsoft-365", cdnId: "m365", repoPath: null },

  // ── Power Platform ──
  { id: "power-apps", displayName: "Power Apps", category: "power-platform", cdnId: "powerapps", repoPath: "Power Platform/Power Apps/PowerApps_scalable.svg" },
  { id: "power-automate", displayName: "Power Automate", category: "power-platform", cdnId: "powerautomate", repoPath: "Power Platform/Power Automate/PowerAutomate_scalable.svg" },
  { id: "power-bi", displayName: "Power BI", category: "power-platform", cdnId: "powerbi", repoPath: "Power Platform/Power BI/PowerBI_scalable.svg" },
  { id: "power-pages", displayName: "Power Pages", category: "power-platform", cdnId: "powerpages", repoPath: "Power Platform/Power Pages/PowerPages_scalable.svg" },
  { id: "power-platform", displayName: "Power Platform", category: "power-platform", cdnId: null, repoPath: "Power Platform/Power Platform/PowerPlatform_scalable.svg" },
  { id: "copilot-studio", displayName: "Copilot Studio", category: "power-platform", cdnId: null, repoPath: "Power Platform/Copilot Studio/CopilotStudio_scalable.svg" },
  { id: "ai-builder", displayName: "AI Builder", category: "power-platform", cdnId: null, repoPath: "Power Platform/AI Builder/AIBuilder_scalable.svg" },
  { id: "dataverse", displayName: "Dataverse", category: "power-platform", cdnId: null, repoPath: "Power Platform/Dataverse/Dataverse_scalable.svg" },
  { id: "power-fx", displayName: "Power Fx", category: "power-platform", cdnId: null, repoPath: "Power Platform/PowerFx_scalable.svg" },
  { id: "connectors", displayName: "Connectors", category: "power-platform", cdnId: null, repoPath: "Power Platform/PowerPlatform_Connectors.svg" },
  { id: "agent-365", displayName: "Agent 365", category: "power-platform", cdnId: null, repoPath: "Power Platform/Agent 365_scalable.svg" },

  // ── Dynamics 365 ──
  { id: "dynamics-365", displayName: "Dynamics 365", category: "dynamics-365", cdnId: "dynamics365", repoPath: "Dynamics 365/Dynamics 365 Product Family Icon/Dynamics365_scalable.svg" },
  { id: "business-central", displayName: "Business Central", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Business Central/BusinessCentral_scalable.svg" },
  { id: "customer-service", displayName: "Customer Service", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Customer Service/CustomerService_scalable.svg" },
  { id: "field-service", displayName: "Field Service", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Field Service/FieldService_scalable.svg" },
  { id: "finance", displayName: "Finance", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Finance/Finance_scalable.svg" },
  { id: "sales", displayName: "Sales", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Sales/Sales_scalable.svg" },
  { id: "supply-chain", displayName: "Supply Chain Management", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Supply Chain Management/SupplyChainManagement_scalable.svg" },
  { id: "commerce", displayName: "Commerce", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Commerce/Commerce_scalable.svg" },
  { id: "remote-assist", displayName: "Remote Assist", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Dynamics 365 Mixed Reality Icons/RemoteAssist_scalable.svg" },
  { id: "guides", displayName: "Guides", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Dynamics 365 Mixed Reality Icons/Guides_scalable.svg" },
  { id: "customer-voice", displayName: "Customer Voice", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Customer Voice/CustomerVoice_scalable.svg" },
  { id: "project-operations", displayName: "Project Operations", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Project Operations/ProjectOperations_scalable.svg" },
  { id: "fraud-protection", displayName: "Fraud Protection", category: "dynamics-365", cdnId: null, repoPath: "Dynamics 365/Fraud Protection/FraudProtection_scalable.svg" },

  // ── Entra ──
  { id: "entra", displayName: "Microsoft Entra", category: "entra", cdnId: null, repoPath: "Entra/Microsoft Entra Product Family.svg" },
  { id: "entra-id", displayName: "Entra ID", category: "entra", cdnId: null, repoPath: "Entra/Microsoft Entra ID color icon.svg" },
  { id: "entra-id-governance", displayName: "Entra ID Governance", category: "entra", cdnId: null, repoPath: "Entra/Microsoft Entra ID Governance color icon.svg" },
  { id: "entra-verified-id", displayName: "Entra Verified ID", category: "entra", cdnId: null, repoPath: "Entra/Microsoft Entra Verified ID color icon.svg" },

  // ── Viva ──
  { id: "viva-connections", displayName: "Viva Connections", category: "viva", cdnId: "vivaconnections", repoPath: "Viva/Viva Connections/Viva Connections.svg" },
  { id: "viva-insights", displayName: "Viva Insights", category: "viva", cdnId: "vivainsights", repoPath: "Viva/Viva Insights/Viva Insights.svg" },
  { id: "viva-learning", displayName: "Viva Learning", category: "viva", cdnId: "vivalearning", repoPath: "Viva/Viva Learning/Viva Learning.svg" },
  { id: "viva-engage", displayName: "Viva Engage", category: "viva", cdnId: "vivaengage", repoPath: "Viva/Viva Engage/Viva Engage.svg" },
  { id: "viva-pulse", displayName: "Viva Pulse", category: "viva", cdnId: "vivapulse", repoPath: "Viva/Viva Pulse/Viva Pulse.svg" },
  { id: "viva-amplify", displayName: "Viva Amplify", category: "viva", cdnId: "vivaamplify", repoPath: "Viva/Viva Amplify/Viva Amplify.svg" },
  { id: "viva-glint", displayName: "Viva Glint", category: "viva", cdnId: null, repoPath: "Viva/Viva Glint/Glint.svg" },
  { id: "viva-suite", displayName: "Viva Suite", category: "viva", cdnId: null, repoPath: "Viva/Viva Suite/Viva Suite.svg" },

  // ── Security ──
  { id: "defender", displayName: "Defender", category: "security", cdnId: "defender", repoPath: null },
  { id: "purview", displayName: "Purview", category: "security", cdnId: "purview", repoPath: null },

  // ── Copilot ──
  { id: "copilot", displayName: "Microsoft Copilot", category: "copilot", cdnId: "copilot", repoPath: null },
  { id: "copilot-365", displayName: "Microsoft 365 Copilot", category: "copilot", cdnId: null, repoPath: "Microsoft 365/Copilot in [app]/Microsoft_365_Copilot.svg" },

  // ── Fabric ──
  { id: "fabric", displayName: "Microsoft Fabric", category: "fabric", cdnId: null, repoPath: "Fabric/Fabric_256.svg" },

  // ── Other ──
  { id: "edge", displayName: "Microsoft Edge", category: "other", cdnId: "edge", repoPath: null },
  { id: "bing", displayName: "Bing", category: "other", cdnId: "bing", repoPath: null },
  { id: "designer", displayName: "Designer", category: "other", cdnId: "designer", repoPath: null },
  { id: "office", displayName: "Office", category: "other", cdnId: "office", repoPath: null },

  { id: "foundry", displayName: "Microsoft Foundry", category: "other", cdnId: null, repoPath: "other/Microsoft Foundry.svg" },
  { id: "family-safety", displayName: "Family Safety", category: "other", cdnId: "familysafety", repoPath: null },
];

// ── Types ───────────────────────────────────────────────────────────

interface IconEntry {
  id: string;
  displayName: string;
  category: string;
  path: string;
  source: "cdn" | "repo";
}

const STANDARD_SIZE = 256;

// ── Helpers ─────────────────────────────────────────────────────────

/** Normalise SVG to a standard display size while preserving the original viewBox */
function normaliseSvgSize(svgContent: string): string {
  // Extract and modify only the <svg ...> opening tag
  const svgTagMatch = svgContent.match(/<svg\b[^>]*>/);
  if (!svgTagMatch) return svgContent;

  let svgTag = svgTagMatch[0];
  const beforeTag = svgContent.slice(0, svgTagMatch.index!);
  const afterTag = svgContent.slice(svgTagMatch.index! + svgTag.length);

  // Ensure viewBox exists — infer from existing width/height if needed
  if (!/viewBox\s*=\s*"[^"]+?"/.test(svgTag)) {
    const wMatch = svgTag.match(/\bwidth\s*=\s*"([^"]+)"/);
    const hMatch = svgTag.match(/\bheight\s*=\s*"([^"]+)"/);
    if (wMatch && hMatch) {
      const w = parseFloat(wMatch[1]);
      const h = parseFloat(hMatch[1]);
      if (!isNaN(w) && !isNaN(h)) {
        svgTag = svgTag.replace(/<svg\b/, `<svg viewBox="0 0 ${w} ${h}"`);
      }
    }
  }

  // Replace or add width on the <svg> tag only
  if (/\bwidth\s*=\s*"[^"]*"/.test(svgTag)) {
    svgTag = svgTag.replace(/\bwidth\s*=\s*"[^"]*"/, `width="${STANDARD_SIZE}"`);
  } else {
    svgTag = svgTag.replace(/<svg\b/, `<svg width="${STANDARD_SIZE}"`);
  }

  // Replace or add height on the <svg> tag only
  if (/\bheight\s*=\s*"[^"]*"/.test(svgTag)) {
    svgTag = svgTag.replace(/\bheight\s*=\s*"[^"]*"/, `height="${STANDARD_SIZE}"`);
  } else {
    svgTag = svgTag.replace(/<svg\b/, `<svg height="${STANDARD_SIZE}"`);
  }

  return beforeTag + svgTag + afterTag;
}

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

function cloneSource(): void {
  if (existsSync(TEMP_DIR)) {
    rmSync(TEMP_DIR, { recursive: true, force: true });
  }
  console.log("📦 Shallow-cloning source repo...");
  execSync(`git clone --depth 1 "${SOURCE_REPO}" "${TEMP_DIR}"`, {
    stdio: "inherit",
  });
}

function cleanup(): void {
  if (existsSync(TEMP_DIR)) {
    console.log("🧹 Cleaning up temp files...");
    rmSync(TEMP_DIR, { recursive: true, force: true });
  }
}

// ── Fetch Icons ─────────────────────────────────────────────────────

async function fetchAllIcons(): Promise<IconEntry[]> {
  const icons: IconEntry[] = [];
  let cdnCount = 0;
  let repoCount = 0;
  let skipped = 0;

  // Phase 1: Clone repo (needed for fallbacks)
  cloneSource();

  // Phase 2: Process each product
  console.log(`\n🔍 Processing ${PRODUCTS.length} products...\n`);

  for (const product of PRODUCTS) {
    const destPath = `icons/${product.category}/${product.id}.svg`;
    const fullDest = join(process.cwd(), destPath);
    let fetched = false;

    // Try CDN first (preferred — latest Fluent 2 brand icons)
    // Fall back to smaller sizes if 48px not available
    if (product.cdnId) {
      for (const size of [48, 32, 24]) {
        const url = `${CDN_BASE}/${product.cdnId}_${size}x1.svg`;
        const svg = await downloadFile(url);
        if (svg) {
          mkdirSync(dirname(fullDest), { recursive: true });
          writeFileSync(fullDest, normaliseSvgSize(svg));
          icons.push({
            id: product.id,
            displayName: product.displayName,
            category: product.category,
            path: destPath,
            source: "cdn",
          });
          cdnCount++;
          fetched = true;
          break;
        }
      }
    }

    // Fall back to repo
    if (!fetched && product.repoPath) {
      const repoFile = join(TEMP_DIR, product.repoPath);
      if (existsSync(repoFile)) {
        mkdirSync(dirname(fullDest), { recursive: true });
        const svg = readFileSync(repoFile, "utf-8");
        writeFileSync(fullDest, normaliseSvgSize(svg));
        icons.push({
          id: product.id,
          displayName: product.displayName,
          category: product.category,
          path: destPath,
          source: "repo",
        });
        repoCount++;
        fetched = true;
      }
    }

    if (!fetched) {
      console.warn(`   ⚠️  ${product.displayName} — no icon found`);
      skipped++;
    }
  }

  console.log(`\n   🌐 ${cdnCount} from Office CDN`);
  console.log(`   📦 ${repoCount} from GitHub repo`);
  if (skipped > 0) console.log(`   ⚠️  ${skipped} unavailable`);

  return icons;
}

// ── Manifest & README ───────────────────────────────────────────────

function writeManifest(icons: IconEntry[]): void {
  const categories: Record<string, string[]> = {};
  for (const icon of icons) {
    if (!categories[icon.category]) categories[icon.category] = [];
    categories[icon.category].push(icon.id);
  }

  const manifest = {
    generated: new Date().toISOString(),
    sources: [
      "https://res.cdn.office.net (Fluent 2 Brand Icons)",
      "https://github.com/loryanstrant/MicrosoftCloudLogos",
    ],
    totalIcons: icons.length,
    categories,
    icons: icons.map((i) => ({
      id: i.id,
      displayName: i.displayName,
      category: i.category,
      path: i.path,
      source: i.source,
    })),
  };

  writeFileSync(MANIFEST_PATH, JSON.stringify(manifest, null, 2));
}

function writeReadme(icons: IconEntry[]): void {
  const grouped = new Map<string, IconEntry[]>();
  for (const icon of icons) {
    if (!grouped.has(icon.category)) grouped.set(icon.category, []);
    grouped.get(icon.category)!.push(icon);
  }

  const categoryNames: Record<string, string> = {
    "microsoft-365": "Microsoft 365",
    "power-platform": "Power Platform",
    "dynamics-365": "Dynamics 365",
    entra: "Entra",
    viva: "Viva",
    security: "Security",
    copilot: "Copilot",
    fabric: "Fabric",
    other: "Other",
  };

  const lines: string[] = [
    "# Microsoft Cloud Product Icons (SVG)",
    "",
    "A curated collection of **latest official SVG icons** for Microsoft Cloud products.",
    "One icon per product. Clean, consistent, ready to use.",
    "",
    "> Sources: [Microsoft Office CDN](https://res.cdn.office.net) (Fluent 2 brand icons) + [loryanstrant/MicrosoftCloudLogos](https://github.com/loryanstrant/MicrosoftCloudLogos)",
    "",
    `**${icons.length} product icons** across **${grouped.size} categories**`,
    "",
    "## Quick Reference",
    "",
    "| Category | Icons | Products |",
    "|----------|-------|----------|",
  ];

  for (const [category, entries] of grouped) {
    const names = entries.map((e) => e.displayName).join(", ");
    lines.push(
      `| ${categoryNames[category] || category} | ${entries.length} | ${names} |`
    );
  }

  lines.push("", "## All Icons", "");

  for (const [category, entries] of grouped) {
    lines.push(`### ${categoryNames[category] || category}`, "");
    lines.push("| Icon | Product | File |");
    lines.push("|------|---------|------|");
    for (const icon of entries.sort((a, b) =>
      a.displayName.localeCompare(b.displayName)
    )) {
      lines.push(
        `| <img src="${icon.path}" width="32" /> | **${icon.displayName}** | \`${icon.path}\` |`
      );
    }
    lines.push("");
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
    "## Adding Products",
    "",
    "Edit the `PRODUCTS` array in `fetch-icons.ts` to add or remove products.",
    "Each entry specifies a CDN id (preferred) and/or a GitHub repo path (fallback).",
    "",
    "## Credits",
    "",
    "- [Microsoft Office CDN](https://res.cdn.office.net) — Fluent 2 brand icons",
    "- [loryanstrant/MicrosoftCloudLogos](https://github.com/loryanstrant/MicrosoftCloudLogos)",
    "",
    "All logos are the property of Microsoft Corporation.",
    ""
  );

  writeFileSync(README_PATH, lines.join("\n"));
}

// ── Run ─────────────────────────────────────────────────────────────

(async () => {
  try {
    console.log("🚀 Official Icons — Fetch & Curate\n");

    // Clean existing icons
    if (existsSync(ICONS_DIR)) {
      rmSync(ICONS_DIR, { recursive: true, force: true });
    }

    const icons = await fetchAllIcons();

    console.log("\n📋 Writing manifest.json...");
    writeManifest(icons);

    console.log("📝 Writing README.md...");
    writeReadme(icons);

    cleanup();

    console.log(`\n✅ Done! ${icons.length} product icons in icons/`);
  } catch (error) {
    cleanup();
    console.error("❌ Failed:", error);
    process.exit(1);
  }
})();
