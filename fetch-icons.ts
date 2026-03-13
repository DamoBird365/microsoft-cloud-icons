/**
 * fetch-icons.ts
 *
 * Fetches curated Microsoft product/application SVG icons from three sources:
 *   1. Official Microsoft download (Power Platform icons ZIP) — highest priority
 *   2. Microsoft Office CDN (Fluent 2 brand icons) — preferred for M365 apps
 *   3. loryanstrant/MicrosoftCloudLogos GitHub repo — fallback
 *   4. Direct URLs (e.g. powerapps.com CDN for Agent Builder)
 *
 * Outputs a clean, flat folder structure with consistent kebab-case naming.
 */

import { execSync } from "node:child_process";
import {
  cpSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { dirname, join } from "node:path";
import https from "node:https";

// ── Config ──────────────────────────────────────────────────────────

const SOURCE_REPO = "https://github.com/loryanstrant/MicrosoftCloudLogos.git";
const TEMP_DIR = join(process.cwd(), ".tmp-source");
const TEMP_PP_DIR = join(process.cwd(), ".tmp-pp-icons");
const ICONS_DIR = join(process.cwd(), "icons");
const MANIFEST_PATH = join(process.cwd(), "manifest.json");
const README_PATH = join(process.cwd(), "README.md");
const CDN_BASE =
  "https://res.cdn.office.net/midgard/versionless/fluentui-resources/1.1.6/assets/brand-icons/product/svg";

/** Official Microsoft Power Platform icons download */
const PP_ICONS_ZIP =
  "https://download.microsoft.com/download/498606aa-6d27-4f13-aa5c-1401078c153b/Power-Platform-icons-scalable.zip";

// ── Product Catalogue ───────────────────────────────────────────────

interface ProductDef {
  /** kebab-case filename (without .svg) */
  id: string;
  /** Human-readable name */
  displayName: string;
  /** Category folder */
  category: string;
  /** Filename in the official Microsoft PP icons ZIP, or null */
  officialFile: string | null;
  /** CDN product id (for Fluent Brand download), or null */
  officialFile: null, cdnId: string | null;
  /** Direct URL to download SVG from, or null */
  directUrl: string | null;
  /** Path within the GitHub repo (fallback), or null */
  repoPath: string | null;
}

const PRODUCTS: ProductDef[] = [
  // ── Microsoft 365 ──
  { id: "word", displayName: "Word", category: "microsoft-365", officialFile: null, cdnId: "word", directUrl: null, repoPath: null },
  { id: "excel", displayName: "Excel", category: "microsoft-365", officialFile: null, cdnId: "excel", directUrl: null, repoPath: null },
  { id: "powerpoint", displayName: "PowerPoint", category: "microsoft-365", officialFile: null, cdnId: "powerpoint", directUrl: null, repoPath: null },
  { id: "outlook", displayName: "Outlook", category: "microsoft-365", officialFile: null, cdnId: "outlook", directUrl: null, repoPath: null },
  { id: "onenote", displayName: "OneNote", category: "microsoft-365", officialFile: null, cdnId: "onenote", directUrl: null, repoPath: null },
  { id: "teams", displayName: "Teams", category: "microsoft-365", officialFile: null, cdnId: "teams", directUrl: null, repoPath: null },
  { id: "sharepoint", displayName: "SharePoint", category: "microsoft-365", officialFile: null, cdnId: "sharepoint", directUrl: null, repoPath: null },
  { id: "onedrive", displayName: "OneDrive", category: "microsoft-365", officialFile: null, cdnId: "onedrive", directUrl: null, repoPath: null },
  { id: "access", displayName: "Access", category: "microsoft-365", officialFile: null, cdnId: "access", directUrl: null, repoPath: null },
  { id: "publisher", displayName: "Publisher", category: "microsoft-365", officialFile: null, cdnId: "publisher", directUrl: null, repoPath: null },
  { id: "visio", displayName: "Visio", category: "microsoft-365", officialFile: null, cdnId: "visio", directUrl: null, repoPath: null },
  { id: "project", displayName: "Project", category: "microsoft-365", officialFile: null, cdnId: "project", directUrl: null, repoPath: null },
  { id: "loop", displayName: "Loop", category: "microsoft-365", officialFile: null, cdnId: "loop", directUrl: null, repoPath: null },
  { id: "lists", displayName: "Lists", category: "microsoft-365", officialFile: null, cdnId: "lists", directUrl: null, repoPath: null },
  { id: "forms", displayName: "Forms", category: "microsoft-365", officialFile: null, cdnId: "forms", directUrl: null, repoPath: null },
  { id: "planner", displayName: "Planner", category: "microsoft-365", officialFile: null, cdnId: "planner", directUrl: null, repoPath: null },
  { id: "stream", displayName: "Stream", category: "microsoft-365", officialFile: null, cdnId: "stream", directUrl: null, repoPath: null },
  { id: "clipchamp", displayName: "Clipchamp", category: "microsoft-365", officialFile: null, cdnId: "clipchamp", directUrl: null, repoPath: null },
  { id: "sway", displayName: "Sway", category: "microsoft-365", officialFile: null, cdnId: "sway", directUrl: null, repoPath: null },
  { id: "bookings", displayName: "Bookings", category: "microsoft-365", officialFile: null, cdnId: "bookings", directUrl: null, repoPath: null },
  { id: "to-do", displayName: "To Do", category: "microsoft-365", officialFile: null, cdnId: "todo", directUrl: null, repoPath: "Microsoft 365/To Do/To_Do.svg" },
  { id: "whiteboard", displayName: "Whiteboard", category: "microsoft-365", officialFile: null, cdnId: "whiteboard", directUrl: null, repoPath: null },
  { id: "delve", displayName: "Delve", category: "microsoft-365", officialFile: null, cdnId: "delve", directUrl: null, repoPath: null },
  { id: "places", displayName: "Places", category: "microsoft-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Microsoft 365/Places/Microsoft Places.svg" },
  { id: "microsoft-365", displayName: "Microsoft 365", category: "microsoft-365", officialFile: null, cdnId: "m365", directUrl: null, repoPath: null },

  // ── Power Platform (official MS download preferred — Dec 2025 update) ──
  { id: "power-apps", displayName: "Power Apps", category: "power-platform", officialFile: "PowerApps_scalable.svg", cdnId: "powerapps", directUrl: null, repoPath: "Power Platform/Power Apps/PowerApps_scalable.svg" },
  { id: "power-automate", displayName: "Power Automate", category: "power-platform", officialFile: "PowerAutomate_scalable.svg", cdnId: "powerautomate", directUrl: null, repoPath: "Power Platform/Power Automate/PowerAutomate_scalable.svg" },
  { id: "power-bi", displayName: "Power BI", category: "power-platform", officialFile: null, cdnId: "powerbi", directUrl: null, repoPath: "Power Platform/Power BI/PowerBI_scalable.svg" },
  { id: "power-pages", displayName: "Power Pages", category: "power-platform", officialFile: "PowerPages_scalable.svg", cdnId: "powerpages", directUrl: null, repoPath: "Power Platform/Power Pages/PowerPages_scalable.svg" },
  { id: "power-platform", displayName: "Power Platform", category: "power-platform", officialFile: "PowerPlatform_scalable.svg", cdnId: null, directUrl: null, repoPath: "Power Platform/Power Platform/PowerPlatform_scalable.svg" },
  { id: "copilot-studio", displayName: "Copilot Studio", category: "power-platform", officialFile: "CopilotStudio_scalable.svg", cdnId: null, directUrl: null, repoPath: "Power Platform/Copilot Studio/CopilotStudio_scalable.svg" },
  { id: "ai-builder", displayName: "AI Builder", category: "power-platform", officialFile: "AIBuilder_scalable.svg", cdnId: null, directUrl: null, repoPath: "Power Platform/AI Builder/AIBuilder_scalable.svg" },
  { id: "dataverse", displayName: "Dataverse", category: "power-platform", officialFile: "Dataverse_scalable.svg", cdnId: null, directUrl: null, repoPath: "Power Platform/Dataverse/Dataverse_scalable.svg" },
  { id: "power-fx", displayName: "Power Fx", category: "power-platform", officialFile: null, cdnId: null, directUrl: null, repoPath: "Power Platform/PowerFx_scalable.svg" },
  { id: "connectors", displayName: "Connectors", category: "power-platform", officialFile: null, cdnId: null, directUrl: null, repoPath: "Power Platform/PowerPlatform_Connectors.svg" },
  { id: "agent-365", displayName: "Agent 365", category: "power-platform", officialFile: "Agent365_scalable.svg", cdnId: null, directUrl: null, repoPath: "Power Platform/Agent 365_scalable.svg" },
  { id: "agent-builder", displayName: "Agent Builder", category: "power-platform", officialFile: null, cdnId: null, directUrl: "https://content.powerapps.com/resource/makerx/static/media/agentbuilder-brand-icon.2a792a41.svg", repoPath: null },

  // ── Dynamics 365 ──
  { id: "dynamics-365", displayName: "Dynamics 365", category: "dynamics-365", officialFile: null, cdnId: "dynamics365", directUrl: null, repoPath: "Dynamics 365/Dynamics 365 Product Family Icon/Dynamics365_scalable.svg" },
  { id: "business-central", displayName: "Business Central", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Business Central/BusinessCentral_scalable.svg" },
  { id: "customer-service", displayName: "Customer Service", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Customer Service/CustomerService_scalable.svg" },
  { id: "field-service", displayName: "Field Service", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Field Service/FieldService_scalable.svg" },
  { id: "finance", displayName: "Finance", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Finance/Finance_scalable.svg" },
  { id: "sales", displayName: "Sales", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Sales/Sales_scalable.svg" },
  { id: "supply-chain", displayName: "Supply Chain Management", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Supply Chain Management/SupplyChainManagement_scalable.svg" },
  { id: "commerce", displayName: "Commerce", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Commerce/Commerce_scalable.svg" },
  { id: "remote-assist", displayName: "Remote Assist", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Dynamics 365 Mixed Reality Icons/RemoteAssist_scalable.svg" },
  { id: "guides", displayName: "Guides", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Dynamics 365 Mixed Reality Icons/Guides_scalable.svg" },
  { id: "customer-voice", displayName: "Customer Voice", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Customer Voice/CustomerVoice_scalable.svg" },
  { id: "project-operations", displayName: "Project Operations", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Project Operations/ProjectOperations_scalable.svg" },
  { id: "fraud-protection", displayName: "Fraud Protection", category: "dynamics-365", officialFile: null, cdnId: null, directUrl: null, repoPath: "Dynamics 365/Fraud Protection/FraudProtection_scalable.svg" },

  // ── Entra ──
  { id: "entra", displayName: "Microsoft Entra", category: "entra", officialFile: null, cdnId: null, directUrl: null, repoPath: "Entra/Microsoft Entra Product Family.svg" },
  { id: "entra-id", displayName: "Entra ID", category: "entra", officialFile: null, cdnId: null, directUrl: null, repoPath: "Entra/Microsoft Entra ID color icon.svg" },
  { id: "entra-id-governance", displayName: "Entra ID Governance", category: "entra", officialFile: null, cdnId: null, directUrl: null, repoPath: "Entra/Microsoft Entra ID Governance color icon.svg" },
  { id: "entra-verified-id", displayName: "Entra Verified ID", category: "entra", officialFile: null, cdnId: null, directUrl: null, repoPath: "Entra/Microsoft Entra Verified ID color icon.svg" },

  // ── Viva ──
  { id: "viva-connections", displayName: "Viva Connections", category: "viva", officialFile: null, cdnId: "vivaconnections", directUrl: null, repoPath: "Viva/Viva Connections/Viva Connections.svg" },
  { id: "viva-insights", displayName: "Viva Insights", category: "viva", officialFile: null, cdnId: "vivainsights", directUrl: null, repoPath: "Viva/Viva Insights/Viva Insights.svg" },
  { id: "viva-learning", displayName: "Viva Learning", category: "viva", officialFile: null, cdnId: "vivalearning", directUrl: null, repoPath: "Viva/Viva Learning/Viva Learning.svg" },
  { id: "viva-engage", displayName: "Viva Engage", category: "viva", officialFile: null, cdnId: "vivaengage", directUrl: null, repoPath: "Viva/Viva Engage/Viva Engage.svg" },
  { id: "viva-pulse", displayName: "Viva Pulse", category: "viva", officialFile: null, cdnId: "vivapulse", directUrl: null, repoPath: "Viva/Viva Pulse/Viva Pulse.svg" },
  { id: "viva-amplify", displayName: "Viva Amplify", category: "viva", officialFile: null, cdnId: "vivaamplify", directUrl: null, repoPath: "Viva/Viva Amplify/Viva Amplify.svg" },
  { id: "viva-glint", displayName: "Viva Glint", category: "viva", officialFile: null, cdnId: null, directUrl: null, repoPath: "Viva/Viva Glint/Glint.svg" },
  { id: "viva-suite", displayName: "Viva Suite", category: "viva", officialFile: null, cdnId: null, directUrl: null, repoPath: "Viva/Viva Suite/Viva Suite.svg" },

  // ── Security ──
  { id: "defender", displayName: "Defender", category: "security", officialFile: null, cdnId: "defender", directUrl: null, repoPath: null },
  { id: "purview", displayName: "Purview", category: "security", officialFile: null, cdnId: "purview", directUrl: null, repoPath: null },

  // ── Copilot ──
  { id: "copilot", displayName: "Microsoft Copilot", category: "copilot", officialFile: null, cdnId: "copilot", directUrl: null, repoPath: null },
  { id: "copilot-365", displayName: "Microsoft 365 Copilot", category: "copilot", officialFile: null, cdnId: null, directUrl: null, repoPath: "Microsoft 365/Copilot in [app]/Microsoft_365_Copilot.svg" },

  // ── Fabric ──
  { id: "fabric", displayName: "Microsoft Fabric", category: "fabric", officialFile: null, cdnId: null, directUrl: null, repoPath: "Fabric/Fabric_256.svg" },

  // ── Other ──
  { id: "edge", displayName: "Microsoft Edge", category: "other", officialFile: null, cdnId: "edge", directUrl: null, repoPath: null },
  { id: "bing", displayName: "Bing", category: "other", officialFile: null, cdnId: "bing", directUrl: null, repoPath: null },
  { id: "designer", displayName: "Designer", category: "other", officialFile: null, cdnId: "designer", directUrl: null, repoPath: null },
  { id: "office", displayName: "Office", category: "other", officialFile: null, cdnId: "office", directUrl: null, repoPath: null },

  { id: "foundry", displayName: "Microsoft Foundry", category: "other", officialFile: null, cdnId: null, directUrl: null, repoPath: "other/Microsoft Foundry.svg" },
  { id: "family-safety", displayName: "Family Safety", category: "other", officialFile: null, cdnId: "familysafety", directUrl: null, repoPath: null },
];

// ── Types ───────────────────────────────────────────────────────────

interface IconEntry {
  id: string;
  displayName: string;
  category: string;
  path: string;
  source: "official" | "cdn" | "direct" | "repo";
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
    rmSync(TEMP_DIR, { recursive: true, force: true });
  }
  if (existsSync(TEMP_PP_DIR)) {
    rmSync(TEMP_PP_DIR, { recursive: true, force: true });
  }
  const zipFile = join(process.cwd(), ".tmp-pp-icons.zip");
  if (existsSync(zipFile)) {
    rmSync(zipFile);
  }
  console.log("🧹 Cleaned up temp files");
}

function downloadOfficialPPIcons(): void {
  const zipPath = join(process.cwd(), ".tmp-pp-icons.zip");
  if (existsSync(TEMP_PP_DIR)) {
    rmSync(TEMP_PP_DIR, { recursive: true, force: true });
  }
  console.log("📥 Downloading official Power Platform icons from Microsoft...");
  execSync(
    `powershell -Command "$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '${PP_ICONS_ZIP}' -OutFile '${zipPath}'; Expand-Archive -Path '${zipPath}' -DestinationPath '${TEMP_PP_DIR}' -Force"`,
    { stdio: "inherit" }
  );
}

/** Find an SVG file by name in the official PP icons directory (recursive) */
function findOfficialFile(filename: string): string | null {
  if (!existsSync(TEMP_PP_DIR)) return null;
  const search = (dir: string): string | null => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (entry.isFile() && entry.name === filename) {
        return join(dir, entry.name);
      }
      if (entry.isDirectory()) {
        const found = search(join(dir, entry.name));
        if (found) return found;
      }
    }
    return null;
  };
  return search(TEMP_PP_DIR);
}

// ── Fetch Icons ─────────────────────────────────────────────────────

async function fetchAllIcons(): Promise<IconEntry[]> {
  const icons: IconEntry[] = [];
  let officialCount = 0;
  let cdnCount = 0;
  let directCount = 0;
  let repoCount = 0;
  let skipped = 0;

  // Phase 1: Download official PP icons + clone repo
  downloadOfficialPPIcons();
  cloneSource();

  // Phase 2: Process each product
  console.log(`\n🔍 Processing ${PRODUCTS.length} products...\n`);

  for (const product of PRODUCTS) {
    const destPath = `icons/${product.category}/${product.id}.svg`;
    const fullDest = join(process.cwd(), destPath);
    let fetched = false;

    // Source 1: Official Microsoft download (highest priority)
    if (!fetched && product.officialFile) {
      const officialPath = findOfficialFile(product.officialFile);
      if (officialPath) {
        mkdirSync(dirname(fullDest), { recursive: true });
        const svg = readFileSync(officialPath, "utf-8");
        writeFileSync(fullDest, normaliseSvgSize(svg));
        icons.push({
          id: product.id,
          displayName: product.displayName,
          category: product.category,
          path: destPath,
          source: "official",
        });
        officialCount++;
        fetched = true;
      }
    }

    // Source 2: Office CDN (Fluent 2 brand icons)
    if (!fetched && product.cdnId) {
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

    // Source 3: Direct URL
    if (!fetched && product.directUrl) {
      const svg = await downloadFile(product.directUrl);
      if (svg) {
        mkdirSync(dirname(fullDest), { recursive: true });
        writeFileSync(fullDest, normaliseSvgSize(svg));
        icons.push({
          id: product.id,
          displayName: product.displayName,
          category: product.category,
          path: destPath,
          source: "direct",
        });
        directCount++;
        fetched = true;
      }
    }

    // Source 4: GitHub repo (fallback)
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

  console.log(`\n   📋 ${officialCount} from Official Microsoft download`);
  console.log(`   🌐 ${cdnCount} from Office CDN`);
  if (directCount > 0) console.log(`   🔗 ${directCount} from direct URL`);
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
      "https://learn.microsoft.com/power-platform/guidance/icons (Official Microsoft PP Icons)",
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
