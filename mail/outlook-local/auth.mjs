import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createInterface } from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { chromium } from "playwright";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const CONFIG_PATH = path.join(__dirname, "config.json");

async function exists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function loadConfig() {
  if (!(await exists(CONFIG_PATH))) {
    throw new Error(
      "Missing config.json. Copy config.example.json to config.json and update your Outlook URLs/selectors first."
    );
  }

  return JSON.parse(await fs.readFile(CONFIG_PATH, "utf8"));
}

async function ensureParentDir(targetPath) {
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
}

function resolveProjectPath(projectRelativePath) {
  return path.resolve(__dirname, projectRelativePath);
}

async function main() {
  const config = await loadConfig();
  const statePath = resolveProjectPath(config.statePath || "./state/storage-state.json");
  const metaPath = path.join(path.dirname(statePath), "auth-meta.json");
  const loginUrl = config.loginUrl || "https://outlook.office.com/mail/";
  const readySelector = config.selectors?.postLoginReady;

  await ensureParentDir(statePath);

  console.log("Opening a local browser for a one-time Outlook Web login.");
  console.log("Do not paste or store your password in chat, config files, or terminal prompts.");
  console.log("Complete sign-in and MFA only in the browser window that opens on this machine.");

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto(loginUrl, { waitUntil: "domcontentloaded" });

  const rl = createInterface({ input, output });
  await rl.question(
    "\nAfter Outlook is fully signed in and your mailbox is visible, press Enter here to save local session state..."
  );
  rl.close();

  if (readySelector) {
    try {
      await page.waitForSelector(readySelector, { timeout: config.defaultTimeoutMs || 45000 });
    } catch {
      console.warn(
        `Post-login selector "${readySelector}" was not found before timeout. Saving state anyway.`
      );
    }
  }

  await context.storageState({ path: statePath });
  await fs.writeFile(
    metaPath,
    JSON.stringify(
      {
        savedAt: new Date().toISOString(),
        loginUrl,
        currentUrl: page.url()
      },
      null,
      2
    )
  );

  console.log(`Saved storage state to ${statePath}`);
  console.log("Reuse it for sync/draft workflows until Outlook asks for a fresh login.");

  await browser.close();
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
