import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import process from "node:process";

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

function resolveProjectPath(projectRelativePath) {
  return path.resolve(__dirname, projectRelativePath);
}

async function ensureDir(targetPath) {
  await fs.mkdir(targetPath, { recursive: true });
}

async function gotoWithRetries(page, url, { timeout, waitUntil = "domcontentloaded", attempts = 3 } = {}) {
  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      await page.goto(url, { waitUntil, timeout });
      return;
    } catch (error) {
      lastError = error;
      const message = String(error?.message || error);
      const currentUrl = page.url();
      const navigationWasInterrupted = /net::ERR_ABORTED/i.test(message);
      const outlookNavigationContinued =
        navigationWasInterrupted && currentUrl && currentUrl !== "about:blank" && /outlook|office|microsoft/i.test(currentUrl);

      if (outlookNavigationContinued) {
        await page.waitForLoadState(waitUntil, { timeout: Math.min(timeout || 5000, 5000) }).catch(() => {});
        return;
      }

      if (attempt === attempts) {
        break;
      }

      await page.waitForTimeout(1000 * attempt);
    }
  }

  throw lastError;
}

function compactWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function safeSlug(value) {
  return compactWhitespace(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "") || "item";
}

function summarizeEmail(email) {
  const parts = [];

  if (email.unread) {
    parts.push("Unread");
  }

  if (email.from) {
    parts.push(`from ${email.from}`);
  }

  if (email.subject) {
    parts.push(`about "${email.subject}"`);
  }

  if (email.preview) {
    parts.push(`preview: ${email.preview}`);
  }

  return parts.join(", ");
}

function buildMarkdown(report) {
  const lines = [
    "# Outlook Inbox Snapshot",
    "",
    `- Synced at: ${report.syncedAt}`,
    `- Inbox URL: ${report.inboxUrl}`,
    `- Email count captured: ${report.emails.length}`,
    `- Email count ignored by rules: ${report.ignoredCount || 0}`,
    "",
    "## Quick summaries",
    ""
  ];

  if (report.emails.length === 0) {
    lines.push("No inbox items were extracted. Review the saved screenshot/HTML and adjust selectors in config.json.");
    return lines.join("\n");
  }

  for (const email of report.emails) {
    lines.push(`### ${email.index}. ${email.subject || "(no subject)"}`);
    lines.push("");
    lines.push(`- From: ${email.from || "Unknown"}`);
    lines.push(`- Received: ${email.receivedAt || "Unknown"}`);
    lines.push(`- Unread: ${email.unread ? "yes" : "no"}`);
    lines.push(`- Summary: ${email.summary}`);
    lines.push(`- Reply ID: ${email.id}`);
    lines.push("");
  }

  lines.push("## Notes");
  lines.push("");
  lines.push("- These summaries are local snapshots of the message list, not full message bodies.");
  lines.push("- If Outlook's DOM differs in your tenant, tweak the selectors in config.json and rerun sync.");
  return lines.join("\n");
}

async function extractVisibleInboxRows(page, rowSelector, selectors, maxItems) {
  return page.locator(rowSelector).evaluateAll((nodes, config) => {
    const normalize = (value) => String(value || "").replace(/\s+/g, " ").trim();
    const looksLikeRealText = (value) => /[A-Za-z0-9]/.test(value) && !/^mark as /i.test(value);
    const datePattern =
      /\b(?:\d{1,2}\/\d{1,2}\/\d{2,4}|\d{1,2}:\d{2}\s?(?:AM|PM)|Today|Yesterday|Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i;
    const pickText = (node, selector) => {
      if (!selector) {
        return "";
      }

      const found = node.querySelector(selector);
      return normalize(found?.textContent || found?.getAttribute("title") || "");
    };
    const extractMeaningfulSpans = (node) =>
      Array.from(node.querySelectorAll("span"))
        .map((span) => normalize(span.textContent || ""))
        .filter((value) => looksLikeRealText(value));
    const parsePreviewFromAria = (ariaLabel, from, subject) => {
      if (!ariaLabel) {
        return "";
      }

      let preview = ariaLabel;

      if (from) {
        const fromIndex = preview.indexOf(from);
        if (fromIndex !== -1) {
          preview = preview.slice(fromIndex + from.length).trim();
        }
      }

      if (subject) {
        const subjectIndex = preview.indexOf(subject);
        if (subjectIndex !== -1) {
          preview = preview.slice(subjectIndex + subject.length).trim();
        }
      }

      const dateMatch = preview.match(datePattern);
      if (dateMatch && typeof dateMatch.index === "number") {
        preview = preview.slice(dateMatch.index + dateMatch[0].length).trim();
      }

      preview = preview.replace(/\bNo items selected\b/i, "").trim();
      return preview.slice(0, 220);
    };

    return nodes.slice(0, config.maxItems).map((node, index) => {
      const ariaLabel = normalize(node.getAttribute("aria-label") || "");
      const textSpans = extractMeaningfulSpans(node);
      const from = textSpans[0] || pickText(node, config.from);
      const subjectCandidate = textSpans.find((value) => value !== from) || pickText(node, config.subject);
      const subject = /^mark as /i.test(subjectCandidate) ? "" : subjectCandidate;
      const buttonTitles = Array.from(node.querySelectorAll("button"))
        .map((button) => normalize(button.getAttribute("title") || button.getAttribute("aria-label") || ""))
        .filter(Boolean);
      const previewCandidate = pickText(node, config.preview);
      const preview =
        previewCandidate && previewCandidate !== from && previewCandidate !== subject
          ? previewCandidate
          : parsePreviewFromAria(ariaLabel, from, subject);
      const receivedAt =
        pickText(node, config.time) ||
        normalize(node.querySelector("time")?.getAttribute("datetime") || "") ||
        ariaLabel.match(datePattern)?.[0] ||
        "";
      const unread =
        buttonTitles.some((value) => /^mark as read$/i.test(value)) ||
        Boolean(config.unreadIndicator && node.querySelector(config.unreadIndicator)) ||
        /\bunread\b/i.test(ariaLabel);
      const rawKey =
        node.getAttribute("data-convid") ||
        node.getAttribute("data-item-id") ||
        node.getAttribute("id") ||
        `${subject}-${from}-${receivedAt}-${index + 1}`;
      const messageId = normalize(node.getAttribute("id") || node.getAttribute("data-item-id") || "");
      const conversationId = normalize(node.getAttribute("data-convid") || "");

      return {
        index: index + 1,
        rawKey,
        messageId,
        conversationId,
        subject,
        from,
        preview,
        receivedAt,
        unread
      };
    }).filter((email) => email.subject || email.from);
  }, {
    maxItems,
    subject: selectors.subject,
    from: selectors.from,
    preview: selectors.preview,
    time: selectors.time,
    unreadIndicator: selectors.unreadIndicator
  });
}

async function resolveInboxScroller(page) {
  const preferredScroller = page.locator("#MailList [data-testid='virtuoso-scroller']").first();
  if ((await preferredScroller.count().catch(() => 0)) > 0) {
    return preferredScroller;
  }

  return page.locator("#MailList [role='listbox']").first();
}

async function extractInbox(page, selectors, maxItems) {
  const defaultRowSelector =
    "#MailList [role='listbox'][aria-label*='Message list'] [role='option'][data-convid][data-focusable-row='true']";
  const rowSelector =
    !selectors.messageRow || selectors.messageRow === "[role='option'], [data-convid], div[aria-selected]"
      ? defaultRowSelector
      : selectors.messageRow;
  const scroller = await resolveInboxScroller(page);

  try {
    await page.locator(rowSelector).first().waitFor({ timeout: 15000 });
  } catch {
    return [];
  }

  let previousCount = -1;
  let stableCountPasses = 0;
  for (let attempt = 0; attempt < 8 && stableCountPasses < 2; attempt += 1) {
    const count = await page.locator(rowSelector).count().catch(() => 0);
    if (count > 0 && count === previousCount) {
      stableCountPasses += 1;
    } else {
      stableCountPasses = 0;
      previousCount = count;
    }

    await page.waitForTimeout(500);
  }

  const collected = new Map();
  let lastSignature = "";
  let stalledPasses = 0;
  const maxPasses = Math.max(6, maxItems * 2);

  for (let pass = 0; pass < maxPasses && collected.size < maxItems; pass += 1) {
    const rows = await extractVisibleInboxRows(page, rowSelector, selectors, maxItems);

    for (const row of rows) {
      const key = row.rawKey || `${row.subject}-${row.from}-${row.receivedAt}`;
      if (!collected.has(key)) {
        collected.set(key, {
          ...row,
          index: collected.size + 1
        });
      }
    }

    const signature = rows.map((row) => row.rawKey).filter(Boolean).join("|");
    if (signature === lastSignature) {
      stalledPasses += 1;
    } else {
      lastSignature = signature;
      stalledPasses = 0;
    }

    if (collected.size >= maxItems || stalledPasses >= 3) {
      break;
    }

    const scrollState = await scroller.evaluate((node) => {
      const previousTop = node.scrollTop || 0;
      const delta = Math.max(200, Math.floor((node.clientHeight || 600) * 0.75));
      node.scrollTop = previousTop + delta;
      return {
        previousTop,
        nextTop: node.scrollTop || 0
      };
    }).catch(() => null);

    if (!scrollState || scrollState.nextTop === scrollState.previousTop) {
      stalledPasses += 1;
    }

    await page.waitForTimeout(500);
  }

  await scroller.evaluate((node) => {
    node.scrollTop = 0;
  }).catch(() => {});
  await page.waitForTimeout(250);

  return Array.from(collected.values()).slice(0, maxItems);
}

async function writeJson(targetPath, value) {
  await fs.writeFile(targetPath, JSON.stringify(value, null, 2));
}

function applyIgnoreRules(emails, ignoreRules = {}) {
  const normalize = (value) => compactWhitespace(String(value || "")).toLowerCase();
  const subjectIncludes = Array.isArray(ignoreRules.subjectIncludes)
    ? ignoreRules.subjectIncludes.map(normalize).filter(Boolean)
    : [];
  const fromIncludes = Array.isArray(ignoreRules.fromIncludes)
    ? ignoreRules.fromIncludes.map(normalize).filter(Boolean)
    : [];
  const previewIncludes = Array.isArray(ignoreRules.previewIncludes)
    ? ignoreRules.previewIncludes.map(normalize).filter(Boolean)
    : [];

  const filtered = emails.filter((email) => {
    const subject = normalize(email.subject);
    const from = normalize(email.from);
    const preview = normalize(email.preview);

    if (subjectIncludes.some((value) => subject.includes(value))) {
      return false;
    }

    if (fromIncludes.some((value) => from.includes(value))) {
      return false;
    }

    if (previewIncludes.some((value) => preview.includes(value))) {
      return false;
    }

    return true;
  });

  return {
    emails: filtered,
    ignoredCount: emails.length - filtered.length
  };
}

export async function runSync(options = {}) {
  const quiet = options.quiet === true;
  const config = await loadConfig();
  const statePath = resolveProjectPath(config.statePath || "./state/storage-state.json");
  const outputDir = resolveProjectPath(config.outputDir || "./output");
  const inboxDir = path.join(outputDir, "inbox");
  const timeoutMs = config.defaultTimeoutMs || 45000;

  if (!(await exists(statePath))) {
    throw new Error("Missing saved browser state. Run `npm run auth` first.");
  }

  await ensureDir(inboxDir);

  const browser = await chromium.launch({
    headless: config.headless !== false,
    args: ["--disable-dev-shm-usage"]
  });
  const viewport = {
    width: Number(config.viewport?.width) || 1440,
    height: Number(config.viewport?.height) || 1100
  };
  const context = await browser.newContext({
    storageState: statePath,
    viewport
  });

  try {
    const page = await context.newPage();

    const inboxUrl = config.inboxUrl || "https://outlook.office.com/mail/";
    const selectors = config.selectors || {};

    if (!quiet) {
      console.log(`Opening Outlook inbox: ${inboxUrl}`);
    }
    await gotoWithRetries(page, inboxUrl, { waitUntil: "domcontentloaded", timeout: timeoutMs });

    if (selectors.inboxReady) {
      await page.waitForSelector(selectors.inboxReady, { timeout: timeoutMs });
    }

    await page.waitForTimeout(4000);
    const rawRows = await extractInbox(page, selectors, config.maxInboxItems || 15);

    if (rawRows.length === 0) {
      const bodyText = compactWhitespace(
        await page.locator("body").innerText({ timeout: 5000 }).catch(() => "")
      );
      const authChallengeVisible =
        /verify your identity|approve sign in request|enter code|text .*\*+\d+|call .*\*+\d+|sign in/i.test(bodyText) ||
        /login\.microsoftonline\.com|microsoftonline|signin/i.test(page.url());

      if (authChallengeVisible) {
        throw new Error("Outlook auth requires verification/MFA. Open Outlook in the saved browser session and complete verification, then rerun the checker.");
      }
    }

    const mappedEmails = rawRows.map((email) => {
      const id = `${email.index}-${safeSlug(email.subject || email.rawKey).slice(0, 40)}`;
      return {
        ...email,
        id,
        summary: summarizeEmail(email)
      };
    });
    const { emails, ignoredCount } = applyIgnoreRules(mappedEmails, config.ignoreRules);

    const report = {
      syncedAt: new Date().toISOString(),
      inboxUrl: page.url(),
      selectorsUsed: {
        inboxReady: selectors.inboxReady,
        messageList: selectors.messageList,
        messageRow: selectors.messageRow,
        subject: selectors.subject,
        from: selectors.from,
        preview: selectors.preview,
        time: selectors.time,
        unreadIndicator: selectors.unreadIndicator
      },
      ignoreRules: config.ignoreRules || {},
      ignoredCount,
      emails
    };

    const historyStem = report.syncedAt.replace(/[:.]/g, "-");
    const latestJsonPath = path.join(inboxDir, "latest.json");
    const latestMdPath = path.join(inboxDir, "latest.md");
    const latestHtmlPath = path.join(inboxDir, "latest.html");
    const latestPngPath = path.join(inboxDir, "latest.png");

    await writeJson(latestJsonPath, report);
    await fs.writeFile(latestMdPath, buildMarkdown(report));

    if (config.captureHtml !== false) {
      const html = await page.content();
      await fs.writeFile(latestHtmlPath, html);
      await fs.writeFile(path.join(inboxDir, `${historyStem}.html`), html);
    }

    if (config.captureScreenshots !== false) {
      const screenshot = await page.screenshot({ fullPage: true, type: "png" });
      await fs.writeFile(latestPngPath, screenshot);
      await fs.writeFile(path.join(inboxDir, `${historyStem}.png`), screenshot);
    }

    await writeJson(path.join(inboxDir, `${historyStem}.json`), report);
    await fs.writeFile(path.join(inboxDir, `${historyStem}.md`), buildMarkdown(report));

    const perEmailDir = path.join(inboxDir, "emails");
    await ensureDir(perEmailDir);
    for (const email of emails) {
      await writeJson(path.join(perEmailDir, `${email.id}.json`), email);
    }

    if (!quiet) {
      console.log(`Inbox sync complete. Captured ${emails.length} email rows.`);
      if (ignoredCount > 0) {
        console.log(`Ignored ${ignoredCount} email rows based on local rules.`);
      }
      console.table(
        emails.map((email) => ({
          id: email.id,
          from: email.from,
          subject: email.subject,
          unread: email.unread
        }))
      );
    }

    return report;
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }
}

async function main() {
  await runSync();
}

const isDirectRun = process.argv[1] && path.resolve(process.argv[1]) === __filename;

if (isDirectRun) {
  main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}
