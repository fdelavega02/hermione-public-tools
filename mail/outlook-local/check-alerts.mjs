import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { runSync } from "./sync-inbox.mjs";
import { fetchMessageBodiesForEmails, summarizePreviewText } from "./outlook-session-api.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const CONFIG_PATH = path.join(__dirname, "config.json");
const DEFAULT_ALERT_STATE_PATH = "./state/alert-state.json";
const MAX_STORED_KEYS = 250;
const MAX_ALERT_ITEMS = 5;
const DEFAULT_ALERT_RULES = {
  timeZone: "America/Indianapolis",
  weekdays: ["Mon", "Tue", "Wed", "Thu", "Fri"],
  startHour: 7,
  startMinute: 0,
  endHour: 16,
  endMinute: 0,
  subjectIncludes: ["IT", "issue", "request", "printer", "radio", "excel", "powerpoint", "word", "outlook"],
  senderIncludes: ["Hogle, Karenah"]
};

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
      "Missing config.json. Copy config.example.json to config.json and update your Outlook settings first."
    );
  }

  return JSON.parse(await fs.readFile(CONFIG_PATH, "utf8"));
}

function resolveProjectPath(projectRelativePath) {
  return path.resolve(__dirname, projectRelativePath);
}

async function ensureParentDir(targetPath) {
  await fs.mkdir(path.dirname(targetPath), { recursive: true });
}

function compactWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalize(value) {
  return compactWhitespace(value).toLowerCase();
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getAlertRules(config) {
  const rules = config?.alertRules || {};
  return {
    ...DEFAULT_ALERT_RULES,
    ...rules,
    weekdays: Array.isArray(rules.weekdays) && rules.weekdays.length ? rules.weekdays : DEFAULT_ALERT_RULES.weekdays,
    subjectIncludes:
      Array.isArray(rules.subjectIncludes) && rules.subjectIncludes.length
        ? rules.subjectIncludes
        : DEFAULT_ALERT_RULES.subjectIncludes,
    senderIncludes:
      Array.isArray(rules.senderIncludes) && rules.senderIncludes.length
        ? rules.senderIncludes
        : DEFAULT_ALERT_RULES.senderIncludes
  };
}

function buildEmailKey(email) {
  const rawKey = compactWhitespace(email.rawKey);

  return [
    rawKey ? `raw:${rawKey}` : "fallback",
    compactWhitespace(email.subject).toLowerCase(),
    compactWhitespace(email.from).toLowerCase(),
    compactWhitespace(email.receivedAt).toLowerCase(),
    compactWhitespace(email.preview).toLowerCase(),
    email.unread ? "unread" : "read"
  ].join("|");
}

function summarizeEmail(email) {
  const sender = compactWhitespace(email.from) || "Unknown sender";
  const subject = compactWhitespace(email.subject) || "(no subject)";
  const summaryText = compactWhitespace(
    email.summaryText || summarizePreviewText(subject, email.preview || "") || "No preview is available."
  );
  return `From ${sender}. Subject: "${subject}". Summary: ${summaryText}`;
}

function buildAlertMessage(alertEmails) {
  const summaries = alertEmails
    .slice(0, MAX_ALERT_ITEMS)
    .map((email) => summarizeEmail(email))
    .filter(Boolean);

  if (summaries.length === 0) {
    return "NO_REPLY";
  }

  if (summaries.length === 1) {
    return summaries[0];
  }

  return summaries.map((summary) => `- ${summary}`).join("\n");
}

function isWithinAlertWindow(now, rules) {
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone: rules.timeZone,
    weekday: "short",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  });
  const parts = formatter.formatToParts(now);
  const values = Object.fromEntries(parts.filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
  const weekday = normalize(values.weekday).slice(0, 3);
  const hour = Number.parseInt(values.hour || "0", 10);
  const minute = Number.parseInt(values.minute || "0", 10);
  const allowedWeekdays = new Set((rules.weekdays || []).map((value) => normalize(value).slice(0, 3)));

  if (!allowedWeekdays.has(weekday)) {
    return false;
  }

  if (hour < rules.startHour || (hour === rules.startHour && minute < rules.startMinute)) {
    return false;
  }

  if (hour > rules.endHour || (hour === rules.endHour && minute > rules.endMinute)) {
    return false;
  }

  return true;
}

function matchesSubjectKeyword(subject, keywords) {
  const subjectText = subject || "";
  return (keywords || []).some((keyword) => {
    const cleaned = compactWhitespace(keyword);
    if (!cleaned) {
      return false;
    }

    return new RegExp(`\\b${escapeRegExp(cleaned)}\\b`, "i").test(subjectText);
  });
}

function isAlertableEmail(email, rules) {
  if (!email?.unread) {
    return false;
  }

  const sender = normalize(email.from);
  const subject = email.subject || "";
  const senderMatches = (rules.senderIncludes || []).some((value) => {
    const needle = normalize(value);
    return needle && sender.includes(needle);
  });

  if (senderMatches) {
    return true;
  }

  return matchesSubjectKeyword(subject, rules.subjectIncludes);
}

async function readAlertEmailSummaries(config, report, alertableEmails) {
  if (!alertableEmails.length) {
    return alertableEmails;
  }

  const statePath = resolveProjectPath(config.statePath || "./state/storage-state.json");
  if (!(await exists(statePath))) {
    return alertableEmails;
  }

  const emailsForBodyFetch = alertableEmails.slice(0, MAX_ALERT_ITEMS);
  const remainingEmails = alertableEmails.slice(MAX_ALERT_ITEMS).map((email) => ({
    ...email,
    summarySource: "preview",
    summaryText: compactWhitespace(email.preview) || "No preview is available."
  }));
  const timeoutMs = Math.min(config.defaultTimeoutMs || 45000, 10000);
  const hardDeadlineMs = Math.min(timeoutMs + 2000, 12000);

  const enrichedHead = await Promise.race([
    fetchMessageBodiesForEmails({
      statePath,
      inboxUrl: report?.inboxUrl || config.inboxUrl,
      emails: emailsForBodyFetch,
      timeoutMs,
      viewport: {
        ...(config.viewport || {}),
        maxInboxItems: config.maxInboxItems || 15
      }
    }),
    new Promise((resolve) => {
      setTimeout(() => {
        resolve(
          emailsForBodyFetch.map((email) => ({
            ...email,
            summarySource: "preview-timeout",
            summaryText: compactWhitespace(email.preview) || "No preview is available."
          }))
        );
      }, hardDeadlineMs);
    })
  ]);

  return [...enrichedHead, ...remainingEmails];
}

async function loadAlertState(alertStatePath) {
  if (!(await exists(alertStatePath))) {
    return null;
  }

  return JSON.parse(await fs.readFile(alertStatePath, "utf8"));
}

async function writeAlertState(alertStatePath, state) {
  await ensureParentDir(alertStatePath);
  await fs.writeFile(alertStatePath, JSON.stringify(state, null, 2));
}

export async function main() {
  const config = await loadConfig();
  const alertRules = getAlertRules(config);
  if (!isWithinAlertWindow(new Date(), alertRules)) {
    process.stdout.write("NO_REPLY\n");
    return;
  }

  const alertStatePath = resolveProjectPath(config.alertStatePath || DEFAULT_ALERT_STATE_PATH);
  const previousState = await loadAlertState(alertStatePath);

  try {
    const report = await runSync({ quiet: true });
    const emails = Array.isArray(report?.emails) ? report.emails : [];
    const alertableEmails = emails.filter((email) => isAlertableEmail(email, alertRules));
    const currentKeys = alertableEmails.map((email) => buildEmailKey(email));

    await writeAlertState(alertStatePath, {
      initializedAt: previousState?.initializedAt || new Date().toISOString(),
      lastCheckedAt: new Date().toISOString(),
      lastSyncedAt: report.syncedAt,
      lastError: null,
      seenKeys: currentKeys.slice(0, MAX_STORED_KEYS)
    });

    if (alertableEmails.length === 0) {
      process.stdout.write("NO_REPLY\n");
      return;
    }

    const summarizedEmails = await readAlertEmailSummaries(config, report, alertableEmails);
    process.stdout.write(`${buildAlertMessage(summarizedEmails, report)}\n`);
  } catch (error) {
    const message = compactWhitespace(error?.message || error);
    const errorFingerprint = message.toLowerCase();
    const lastErrorFingerprint = compactWhitespace(previousState?.lastError).toLowerCase();

    await writeAlertState(alertStatePath, {
      initializedAt: previousState?.initializedAt || null,
      lastCheckedAt: new Date().toISOString(),
      lastSyncedAt: previousState?.lastSyncedAt || null,
      lastError: message,
      seenKeys: Array.isArray(previousState?.seenKeys)
        ? previousState.seenKeys.slice(0, MAX_STORED_KEYS)
        : []
    });

    if (lastErrorFingerprint === errorFingerprint) {
      process.stdout.write("NO_REPLY\n");
      return;
    }

    process.stdout.write(`Outlook alert check failed: ${message}\n`);
  }
}

const isDirectRun = process.argv[1] && path.resolve(process.argv[1]) === __filename;

if (isDirectRun) {
  main().catch((error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}
