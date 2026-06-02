import fs from "node:fs/promises";
import { chromium } from "playwright";

function compactWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function decodeHtmlEntities(value) {
  return String(value || "")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number.parseInt(code, 10)))
    .replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCodePoint(Number.parseInt(code, 16)))
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

function htmlToText(value) {
  const normalized = String(value || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<(br|\/p|\/div|\/li|\/tr|\/h[1-6])\b[^>]*>/gi, "\n")
    .replace(/<li\b[^>]*>/gi, "- ")
    .replace(/<[^>]+>/g, " ");

  return decodeHtmlEntities(normalized)
    .replace(/\r/g, "")
    .replace(/[ \t\f\v]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function loadStorageState(statePath) {
  return JSON.parse(await fs.readFile(statePath, "utf8"));
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

function readLocalStorageEntries(storageState) {
  return (storageState.origins || []).flatMap((origin) => origin.localStorage || []);
}

function parseMsalTokenEntry(entry) {
  if (!/\|accesstoken\|/i.test(entry?.name || "")) {
    return null;
  }

  let value;
  try {
    value = JSON.parse(entry.value);
  } catch {
    return null;
  }

  const expiresOn = Number.parseInt(String(value.expiresOn || "0"), 10);
  if (!value.secret || !expiresOn) {
    return null;
  }

  return {
    key: entry.name,
    token: value.secret,
    target: String(value.target || ""),
    expiresOn
  };
}

function chooseAccessToken(storageState, matcher) {
  const nowSeconds = Math.floor(Date.now() / 1000);
  const candidates = readLocalStorageEntries(storageState)
    .map(parseMsalTokenEntry)
    .filter(Boolean)
    .filter((entry) => entry.expiresOn > nowSeconds + 60)
    .filter((entry) => matcher(entry.target, entry.key))
    .sort((left, right) => right.expiresOn - left.expiresOn);

  return candidates[0] || null;
}

function getCookieValue(storageState, cookieName) {
  return (
    (storageState.cookies || []).find((cookie) => cookie.name === cookieName)?.value || ""
  );
}

function getSessionApiTokensFromStorageState(storageState) {
  return {
    outlook:
      chooseAccessToken(
        storageState,
        (target) => /https:\/\/outlook\.office\.com\//i.test(target) && /\b(mail\.readwrite|owa\.accessasuser\.all)\b/i.test(target)
      ) || null,
    graph:
      chooseAccessToken(
        storageState,
        (target) => /https:\/\/graph\.microsoft\.com\//i.test(target) && /\b(user\.read|mail\.read|mail\.readwrite)\b/i.test(target)
      ) || null
  };
}

export async function getSessionApiTokens(statePath) {
  const storageState = await loadStorageState(statePath);
  return getSessionApiTokensFromStorageState(storageState);
}

function normalize(value) {
  return compactWhitespace(value).toLowerCase();
}

function getOwaMailboxName(item) {
  return compactWhitespace(
    item?.From?.Mailbox?.Name ||
      item?.Sender?.Mailbox?.Name ||
      item?.From?.Mailbox?.EmailAddress ||
      item?.Sender?.Mailbox?.EmailAddress ||
      ""
  );
}

function getOwaPreview(item) {
  return compactWhitespace(item?.Preview || item?.BodyPreview || "");
}

function truncate(value, maxLength) {
  const text = compactWhitespace(value);
  if (!text || text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, Math.max(0, maxLength - 1)).trimEnd()}…`;
}

function trimQuotedText(value) {
  const text = String(value || "");
  const markers = [
    /\n\s*From:\s/i,
    /\n\s*Sent:\s/i,
    /\n\s*To:\s/i,
    /\n\s*Subject:\s/i,
    /\n\s*On .+ wrote:\s*/i,
    /\n\s*Begin forwarded message:\s*/i,
    /\n\s*-{2,}\s*Original Message\s*-{2,}\s*/i
  ];

  let earliestIndex = -1;
  for (const marker of markers) {
    const match = marker.exec(text);
    if (!match || typeof match.index !== "number") {
      continue;
    }

    if (earliestIndex === -1 || match.index < earliestIndex) {
      earliestIndex = match.index;
    }
  }

  return earliestIndex === -1 ? text : text.slice(0, earliestIndex);
}

function stripSignatureAndNoiseLines(lines) {
  const cleaned = [];

  for (const rawLine of lines) {
    const line = compactWhitespace(rawLine);
    if (!line) {
      continue;
    }

    if (/^image removed by sender\.?$/i.test(line)) {
      continue;
    }

    if (/^(thanks|thank you|regards|best|sincerely|sent from my)/i.test(line)) {
      break;
    }

    if (/^[OMFWP]\.?\s*[:.]?\s*\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/i.test(line)) {
      break;
    }

    if (/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(line)) {
      break;
    }

    if (/\b(human resources|manager|specialist|manufacturing facility|phone|mobile|office)\b/i.test(line) && cleaned.length) {
      break;
    }

    if (/^from:\s|^sent:\s|^to:\s|^subject:\s/i.test(line)) {
      break;
    }

    cleaned.push(line);
  }

  return cleaned;
}

function splitIntoSummaryUnits(text) {
  return String(text || "")
    .split(/(?<=[.!?])\s+|\n+/)
    .map((part) => compactWhitespace(part))
    .filter(Boolean);
}

function isUsefulSummaryUnit(unit) {
  if (!unit) {
    return false;
  }

  if (unit.length < 4) {
    return false;
  }

  if (/^(hi|hello|team|all|thanks|thank you|regards|best)\b/i.test(unit)) {
    return false;
  }

  if (/\b(?:confidential|privileged|disclaimer|facebook|instagram|linkedin|www\.)\b/i.test(unit)) {
    return false;
  }

  return true;
}

function buildIssueSummary(text) {
  const compact = compactWhitespace(text);
  if (!compact) {
    return "";
  }

  const versionMatches = Array.from(compact.matchAll(/\bversion\s+(\d+(?:\.\d+)?)/gi)).map((match) => match[1]);
  const hasFirmwareMismatch = /\b(?:mismatch|keying)\b/i.test(compact) && /\b(?:drive|plc|revision|firmware)\b/i.test(compact);
  const hasStudioBlock = /\bstudio 5000\b/i.test(compact) && /\b(?:cannot|can't|unable|not able|missing|without)\b/i.test(compact);
  const hasAccessBlock = /\b(?:access|download|install|update|flash|permissions|rights)\b/i.test(compact) && /\b(?:cannot|can't|unable|not able|missing|without|need)\b/i.test(compact);

  const parts = [];
  if (hasFirmwareMismatch) {
    if (versionMatches.length >= 2) {
      parts.push(`Firmware mismatch on the replacement drive (${versionMatches[0]} vs ${versionMatches[1]}).`);
    } else {
      parts.push("Firmware mismatch on the replacement drive.");
    }
  }

  if (hasStudioBlock || hasAccessBlock) {
    const blockers = [];
    if (hasStudioBlock) {
      blockers.push("Studio 5000/admin access");
    }
    if (/\b(?:download|install|update|flash)\b/i.test(compact)) {
      blockers.push("updates/downloads");
    }

    if (blockers.length) {
      parts.push(`He is blocked by ${Array.from(new Set(blockers)).join(" and ")}.`);
    }
  }

  if (parts.length) {
    return parts.join(" ");
  }

  if (/\btest\b/i.test(compact) && compact.length < 80) {
    return "Simple test message.";
  }

  return "";
}

function buildReminderSummary(subject, bodyText) {
  const subjectText = compactWhitespace(subject);
  const body = compactWhitespace(bodyText);
  const combined = `${subjectText} ${body}`.trim();
  if (!combined) {
    return "";
  }

  if (/\brestart(?:ing)? (?:your|their) computer every day before (?:you|they) leave the plant\b/i.test(combined) || /\beveryone should restart (?:their|the) computer every day before (?:you|they) leave the plant\b/i.test(combined)) {
    const mentionsBestPractice = /\bbest practice\b/i.test(combined);
    const mentionsPerformance = /\boptimal computer performance\b/i.test(combined) || /\bperformance\b/i.test(combined);
    const mentionsShutdownDifference = /\bshut(?:ting)? down is different\b/i.test(combined) || /\bshutdown is different\b/i.test(combined);

    const parts = [
      `Team reminder to restart computers each day before leaving the plant${mentionsBestPractice ? " as a best practice" : ""}${mentionsPerformance ? " to support optimal performance" : ""}.`
    ];

    if (/\bclear temporary issues\b/i.test(combined) || /\bbackground updates\b/i.test(combined) || /\bclean start\b/i.test(combined)) {
      parts.push("It says this helps clear temporary issues, finish updates, and give the system a clean start.");
    }

    if (mentionsShutdownDifference || /\brestart is usually the better first step\b/i.test(combined)) {
      parts.push("It also explains that restarting is better than shutting down when a computer is acting up because it fully reloads the system.");
    }

    return parts.join(" ");
  }


  return "";
}

function extractInlineField(text, label, stopLabels = []) {
  const compact = compactWhitespace(text);
  if (!compact || !label) {
    return "";
  }

  const stopPattern = stopLabels.length
    ? `(?=\\s+\\b(?:${stopLabels.map((value) => escapeRegExp(value)).join("|")})\\b(?:\\s*:|\\b)|$)`
    : "$";
  const match = compact.match(new RegExp(`\\b${escapeRegExp(label)}\\b\\s*:\\s*(.+?)${stopPattern}`, "i"));
  return compactWhitespace(match?.[1] || "");
}

function buildGenericRequestSummary(subject, bodyText) {
  const subjectText = compactWhitespace(subject);
  const body = compactWhitespace(bodyText);
  const combined = `${subjectText} ${body}`.trim();
  if (!combined || !/\b(?:service request|ticket|request created|new request)\b/i.test(combined)) {
    return "";
  }

  const requestId = compactWhitespace(subjectText.match(/\b(?:SR|REQ|TICKET)[- ]?\d+\b/i)?.[0] || body.match(/\b(?:SR|REQ|TICKET)[- ]?\d+\b/i)?.[0] || "");
  const dueDate = compactWhitespace(body.match(/\bdue\s+(.+?)(?:\.|$)/i)?.[1] || "");
  const assignedTo = compactWhitespace(body.match(/\bassigned to\s+(.+?)(?:\.|$)/i)?.[1] || "");

  const parts = [
    `${requestId ? `Request ${requestId}` : "A request"} was created${assignedTo ? ` and assigned to ${assignedTo}` : ""}.`
  ];

  if (dueDate) {
    parts.push(`It is due ${dueDate}.`);
  }

  return parts.join(" ");
}

function buildOutOfOfficeSummary(subject, bodyText) {
  const subjectText = compactWhitespace(subject);
  const body = compactWhitespace(bodyText);
  const combined = `${subjectText} ${body}`.trim();
  if (!combined || !/\bout of (?:the )?office today\b/i.test(combined)) {
    return "";
  }

  const person = compactWhitespace(
    body.match(/\b(?:As\s+)?([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+is out of (?:the )?office today\b/i)?.[1] ||
      subjectText.match(/\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+Out Today\b/i)?.[1] ||
      ""
  );
  const contact = compactWhitespace(body.match(/\breach out to\s+(.+?)(?=\s+(?:with|for|if|regarding|about)\b|[.,]|$)/i)?.[1] || "");
  const coveragePerson = compactWhitespace(body.match(/\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+is working on\b/i)?.[1] || "");
  const coverageWork = compactWhitespace(body.match(/\bis working on\s+(.+?)(?:\.|$)/i)?.[1] || "");

  const parts = [];
  if (person) {
    parts.push(`${person} is out today.`);
  }
  if (contact) {
    parts.push(`Contact ${contact} for current needs.`);
  }
  if (coveragePerson && coverageWork) {
    parts.push(`${coveragePerson} is working on ${coverageWork}.`);
  }

  return parts.join(" ");
}

function buildComputerRequestSummary(subject, bodyText) {
  const subjectText = compactWhitespace(subject);
  const body = compactWhitespace(bodyText);
  const combined = `${subjectText} ${body}`.trim();
  if (!combined) {
    return "";
  }

  const hasRequestLanguage = /\b(?:can you|could you|please|take a look|look at|check|fix|help)\b/i.test(combined);
  const hasProblemLanguage = /\b(?:takes? forever to load|slow(?:ing|downs?)?|issue|problem|not working|down|won't load|will not load)\b/i.test(combined);
  const isComputerRequest = /\bcomputer\b/i.test(combined) && hasRequestLanguage && hasProblemLanguage;
  if (!isComputerRequest) {
    return "";
  }

  const area = compactWhitespace(body.match(/\bcomputer in the ([^.]+?) area\b/i)?.[1] || "");
  const mentionsNightShift = /\bnight\s*shift\b/i.test(combined);
  const slowLoad = /\b(?:takes?|taking) forever to load\b/i.test(combined) || /\bslow(?:ing|downs?)?\b/i.test(combined);

  const parts = [
    `Request to check ${area ? `the ${area.replace(/^the\s+/i, "")}-area ` : "a "}computer${slowLoad ? " because it is loading very slowly" : ""}.`
  ];

  if (mentionsNightShift) {
    parts.push("Night shift reported the problem.");
  }

  return parts.join(" ");
}

function buildComponentsSummary(subject, bodyText) {
  const subjectText = compactWhitespace(subject);
  const body = compactWhitespace(bodyText);
  const combined = `${subjectText} ${body}`.trim();
  if (!combined) {
    return "";
  }

  const componentHits = ["processor", "ram", "storage", "motherboard", "power supply", "cooling", "fans"]
    .filter((term) => new RegExp(`\\b${escapeRegExp(term)}\\b`, "i").test(combined)).length;
  const hasFailureLanguage = /\b(?:slow|freeze|crash|startup failure|performance drop|overheating|maintenance|hardware checks?|issues?|problems?)\b/i.test(combined);

  if (componentHits < 3 || (!hasFailureLanguage && !/\b(?:it issues|computer issues|components?)\b/i.test(combined))) {
    return "";
  }

  return "Explains how core computer parts like the processor, RAM, storage, motherboard, power supply, and cooling affect performance, and how failures can cause slowdowns, crashes, and startup problems.";
}

function extractTicketName(subject) {
  const subjectText = compactWhitespace(subject);
  if (!subjectText) {
    return "";
  }

  const patterns = [
    /SR-\d+\s*-\s*(.+)$/i,
    /Re:\s*(.+?)\s*-\s*SR-\d+/i,
    /Ticket Updated\s+SR-\d+\s*-\s*(.+)$/i,
    /Service Request Approved\/Rejected\s+SR-\d+\s*-\s*(.+)$/i
  ];

  for (const pattern of patterns) {
    const match = subjectText.match(pattern);
    const value = compactWhitespace(match?.[1] || "");
    if (value) {
      return value;
    }
  }

  return "";
}

function buildTicketSummary(subject, bodyText, sender) {
  const subjectText = compactWhitespace(subject);
  const body = compactWhitespace(bodyText);
  const senderText = compactWhitespace(sender);
  const ticketName = extractTicketName(subjectText);
  const label = ticketName || "the ticket";

  if (!subjectText && !body) {
    return "";
  }

  if (/\bservice request\b/i.test(subjectText) && /\bapproved\b/i.test(subjectText + " " + body)) {
    return `${label} was approved.`;
  }

  if (/\breceived your ticket\b/i.test(subjectText + " " + body) || /\bsupport team has been notified\b/i.test(body)) {
    return `${label} was received and a support team was notified.`;
  }

  if (/\bticket resolved\b/i.test(subjectText) || /\bticket has been resolved\b/i.test(body)) {
    return `${label} was marked resolved.`;
  }

  if (/\bticket updated\b/i.test(subjectText) || /\bnew comment on your ticket\b/i.test(body)) {
    return `New comment on ${label}.`;
  }

  if (/\bassigned to an agent\b/i.test(body) || /\bbeing worked on\b/i.test(body)) {
    return `${label} was assigned to an agent and is being worked on.`;
  }

  if (/\btemporary\s+usb\s+exception\s+has\s+been\s+added\b/i.test(body)) {
    const untilMatch = body.match(/will end on\s+(.+?)(?:\.\s|$)/i);
    const untilText = compactWhitespace(untilMatch?.[1] || "");
    return untilText
      ? `Temporary USB exception added for ${label} until ${untilText}.`
      : `Temporary USB exception added for ${label}.`;
  }

  return "";
}

function squashRunOn(text, maxLength) {
  const compact = compactWhitespace(text);
  if (!compact) {
    return "";
  }

  const craftedSummary = buildIssueSummary(compact);
  if (craftedSummary) {
    return truncate(craftedSummary, maxLength);
  }

  const focusPatterns = [
    /\b(?:issue|problem|request|need|needs|needed|cannot|can't|unable|not able|trying|tried)\b[^.!?;]{0,180}/gi,
    /\b(?:version|mismatch|error|install|download|access|unable|failed)\b[^.!?;]{0,180}/gi
  ];

  const snippets = [];
  for (const pattern of focusPatterns) {
    const matches = compact.match(pattern) || [];
    for (const match of matches) {
      const snippet = compactWhitespace(match);
      if (snippet && !snippets.includes(snippet)) {
        snippets.push(snippet);
      }
      if (snippets.length >= 2) {
        break;
      }
    }
    if (snippets.length >= 2) {
      break;
    }
  }

  const joined = snippets.join("; ");
  return truncate(joined || compact, maxLength);
}

export function summarizePreviewText(subject, previewText, maxLength = 220) {
  const preview = compactWhitespace(previewText);
  const normalizedSubject = normalize(subject);

  if (!preview) {
    return "No preview is available.";
  }

  const reminderSummary = buildReminderSummary(subject, preview);
  if (reminderSummary) {
    return truncate(reminderSummary, maxLength);
  }

  const ticketSummary = buildTicketSummary(subject, preview, "");
  if (ticketSummary) {
    return truncate(ticketSummary, maxLength);
  }

  const genericRequestSummary = buildGenericRequestSummary(subject, preview);
  if (genericRequestSummary) {
    return truncate(genericRequestSummary, maxLength);
  }

  const outOfOfficeSummary = buildOutOfOfficeSummary(subject, preview);
  if (outOfOfficeSummary) {
    return truncate(outOfOfficeSummary, maxLength);
  }

  const computerRequestSummary = buildComputerRequestSummary(subject, preview);
  if (computerRequestSummary) {
    return truncate(computerRequestSummary, maxLength);
  }

  const componentsSummary = buildComponentsSummary(subject, preview);
  if (componentsSummary) {
    return truncate(componentsSummary, maxLength);
  }

  if ((normalizedSubject === "test" || normalizedSubject === "re: test") && /^test\b/i.test(preview)) {
    return "Simple test message.";
  }

  if (/\b(?:manager|specialist|manufacturing facility)\b/i.test(preview) || /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(preview)) {
    const firstChunk = compactWhitespace(preview.split(/(?<=[.!?])\s+|\b(?:manager|specialist|manufacturing facility)\b/i)[0]);
    if (firstChunk) {
      return truncate(firstChunk, maxLength);
    }
  }

  return truncate(preview, maxLength);
}

function summarizeBodyText(subject, bodyText, previewText, maxLength = 220) {
  const body = trimQuotedText(bodyText);
  const preview = compactWhitespace(previewText);
  const normalizedSubject = normalize(subject);

  if (!body) {
    return truncate(preview, maxLength);
  }

  const lines = stripSignatureAndNoiseLines(body.split(/\n+/));
  const content = lines.join("\n");

  if ((normalizedSubject === "test" || normalizedSubject === "re: test") && /^test\b/i.test(compactWhitespace(content))) {
    return "Simple test message.";
  }

  const reminderSummary = buildReminderSummary(subject, content);
  if (reminderSummary) {
    return truncate(reminderSummary, maxLength);
  }

  const ticketSummary = buildTicketSummary(subject, content, "");
  if (ticketSummary) {
    return truncate(ticketSummary, maxLength);
  }

  const genericRequestSummary = buildGenericRequestSummary(subject, content);
  if (genericRequestSummary) {
    return truncate(genericRequestSummary, maxLength);
  }

  const outOfOfficeSummary = buildOutOfOfficeSummary(subject, content);
  if (outOfOfficeSummary) {
    return truncate(outOfOfficeSummary, maxLength);
  }

  const computerRequestSummary = buildComputerRequestSummary(subject, content);
  if (computerRequestSummary) {
    return truncate(computerRequestSummary, maxLength);
  }

  const componentsSummary = buildComponentsSummary(subject, content);
  if (componentsSummary) {
    return truncate(componentsSummary, maxLength);
  }

  const units = splitIntoSummaryUnits(content).filter(isUsefulSummaryUnit);

  if (!units.length) {
    return summarizePreviewText(subject, preview || body, maxLength);
  }

  const issueSummary = buildIssueSummary(content);
  if (issueSummary) {
    return truncate(issueSummary, maxLength);
  }

  if (units.length === 1 && !/[.!?]/.test(units[0]) && units[0].length > 100) {
    return truncate(squashRunOn(content, maxLength) || summarizePreviewText(subject, preview || content, maxLength), maxLength);
  }

  const topUnits = units
    .map((unit, index) => ({
      unit,
      score:
        (index === 0 ? 2 : 0) +
        (/\b(?:request|need|issue|problem|error|slow|down|out of the office|due|assigned|working on|computer|restart|ticket|approved|resolved|created)\b/i.test(unit) ? 3 : 0) +
        (/\b(?:processor|ram|storage|motherboard|power supply|cooling|components?)\b/i.test(unit) ? 2 : 0)
    }))
    .sort((left, right) => right.score - left.score || left.unit.length - right.unit.length)
    .slice(0, 2)
    .map((entry) => entry.unit);

  const stitched = compactWhitespace(topUnits.join(" "));
  if (stitched && stitched.length >= 30 && normalize(stitched) !== normalize(units[0])) {
    return truncate(stitched, maxLength);
  }

  let summary = "";
  for (const unit of units) {
    const candidate = compactWhitespace(summary ? `${summary} ${unit}` : unit);
    if (candidate.length > maxLength) {
      break;
    }

    summary = candidate;
    if (summary.length >= Math.min(140, maxLength)) {
      break;
    }
  }

  if (summary && summary.length >= 30) {
    return truncate(summary, maxLength);
  }

  return truncate(squashRunOn(content, maxLength) || summarizePreviewText(subject, preview || content, maxLength), maxLength);
}

function previewOverlapScore(left, right) {
  const a = normalize(left);
  const b = normalize(right);

  if (!a || !b) {
    return 0;
  }

  if (a === b) {
    return 6;
  }

  if (a.includes(b) || b.includes(a)) {
    return 4;
  }

  const aWords = new Set(a.split(/\s+/).filter((word) => word.length >= 4));
  const bWords = new Set(b.split(/\s+/).filter((word) => word.length >= 4));
  let overlap = 0;

  for (const word of aWords) {
    if (bWords.has(word)) {
      overlap += 1;
    }
  }

  return Math.min(overlap, 3);
}

function scoreCandidate(email, item) {
  const subject = normalize(email.subject);
  const from = normalize(email.from);
  const preview = normalize(email.preview);
  const itemSubject = normalize(item.subject);
  const itemFrom = normalize(item.from);
  const itemPreview = normalize(item.preview);

  let score = 0;

  if (subject && itemSubject) {
    if (subject === itemSubject) {
      score += 8;
    } else if (itemSubject.includes(subject) || subject.includes(itemSubject)) {
      score += 4;
    }
  }

  if (from && itemFrom) {
    if (from === itemFrom) {
      score += 7;
    } else if (itemFrom.includes(from) || from.includes(itemFrom)) {
      score += 4;
    }
  }

  if (typeof email.unread === "boolean" && typeof item.unread === "boolean" && email.unread === item.unread) {
    score += 2;
  }

  score += previewOverlapScore(preview, itemPreview);

  const emailReceived = normalize(email.receivedAt);
  const itemReceived = normalize(item.receivedAt);
  if (emailReceived && itemReceived && (itemReceived.includes(emailReceived) || emailReceived.includes(itemReceived))) {
    score += 2;
  }

  return score;
}

function matchEmailsToItems(emails, items) {
  const remaining = items.map((item) => ({ ...item }));
  const matches = new Map();

  for (const email of emails) {
    let bestIndex = -1;
    let bestScore = -1;

    for (let index = 0; index < remaining.length; index += 1) {
      const score = scoreCandidate(email, remaining[index]);
      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    }

    if (bestIndex >= 0 && bestScore >= 8) {
      const [matched] = remaining.splice(bestIndex, 1);
      matches.set(email.rawKey || email.messageId || `${email.subject}-${email.from}`, matched);
    }
  }

  return matches;
}

function normalizeFindItem(item) {
  return {
    itemId: compactWhitespace(item?.ItemId?.Id),
    changeKey: compactWhitespace(item?.ItemId?.ChangeKey),
    subject: compactWhitespace(item?.Subject),
    from: getOwaMailboxName(item),
    preview: getOwaPreview(item),
    receivedAt: compactWhitespace(item?.DateTimeReceived),
    unread: item?.IsRead === false
  };
}

function normalizeGetItem(item, fallbackPreview = "") {
  const bodyHtml = String(item?.NormalizedBody?.Value || item?.NormalizedBody?.Content || item?.Body?.Value || item?.Body?.Content || "");
  const bodyText = htmlToText(bodyHtml);
  const previewText = compactWhitespace(item?.Preview || item?.BodyPreview || fallbackPreview || "");
  const subject = compactWhitespace(item?.Subject);

  return {
    itemId: compactWhitespace(item?.ItemId?.Id),
    subject,
    bodyHtml,
    bodyText,
    previewText,
    summaryText: summarizeBodyText(subject, bodyText, previewText)
  };
}

async function withOutlookPage({ statePath, inboxUrl, viewport, timeoutMs }, callback) {
  const browser = await chromium.launch({
    headless: true,
    args: ["--disable-dev-shm-usage"]
  });
  const context = await browser.newContext({
    storageState: statePath,
    viewport: {
      width: Number(viewport?.width) || 1440,
      height: Number(viewport?.height) || 1100
    }
  });

  try {
    const page = await context.newPage();
    await gotoWithRetries(page, inboxUrl || "https://outlook.cloud.microsoft/mail/", {
      waitUntil: "domcontentloaded",
      timeout: timeoutMs
    });
    await page.waitForTimeout(1500);
    return await callback(page);
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }
}

async function callOwaAction(page, { action, payload, token, anchorMailbox, requestNumber }) {
  const result = await page.evaluate(
    async ({ actionName, requestBody, accessToken, mailbox, requestNumberValue }) => {
      const response = await fetch(`/owa/service.svc?action=${encodeURIComponent(actionName)}&app=Mail&n=${requestNumberValue}`, {
        method: "POST",
        credentials: "same-origin",
        headers: {
          accept: "*/*",
          action: actionName,
          authorization: `Bearer ${accessToken}`,
          "content-type": "application/json; charset=utf-8",
          prefer: 'IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"',
          "x-anchormailbox": mailbox,
          "x-req-source": "Mail"
        },
        body: JSON.stringify(requestBody)
      });

      const text = await response.text();
      return {
        ok: response.ok,
        status: response.status,
        text
      };
    },
    {
      actionName: action,
      requestBody: payload,
      accessToken: token,
      mailbox: anchorMailbox,
      requestNumberValue: requestNumber
    }
  );

  let json = null;
  try {
    json = result.text ? JSON.parse(result.text) : null;
  } catch {
    json = null;
  }

  return {
    ...result,
    json
  };
}

function buildFindInboxItemsPayload(maxEntriesReturned) {
  return {
    __type: "FindItemJsonRequest:#Exchange",
    Header: {
      __type: "JsonRequestHeaders:#Exchange",
      RequestServerVersion: "V2018_01_08",
      TimeZoneContext: {
        __type: "TimeZoneContext:#Exchange",
        TimeZoneDefinition: {
          __type: "TimeZoneDefinitionType:#Exchange",
          Id: "Eastern Standard Time"
        }
      }
    },
    Body: {
      __type: "FindItemRequest:#Exchange",
      ParentFolderIds: [{ __type: "DistinguishedFolderId:#Exchange", Id: "inbox" }],
      ItemShape: {
        __type: "ItemResponseShape:#Exchange",
        BaseShape: "IdOnly",
        AdditionalProperties: [
          { __type: "PropertyUri:#Exchange", FieldURI: "Subject" },
          { __type: "PropertyUri:#Exchange", FieldURI: "Preview" },
          { __type: "PropertyUri:#Exchange", FieldURI: "DateTimeReceived" },
          { __type: "PropertyUri:#Exchange", FieldURI: "From" },
          { __type: "PropertyUri:#Exchange", FieldURI: "IsRead" }
        ]
      },
      Paging: {
        __type: "IndexedPageView:#Exchange",
        BasePoint: "Beginning",
        Offset: 0,
        MaxEntriesReturned: Math.max(10, maxEntriesReturned)
      },
      Traversal: "Shallow",
      ViewFilter: "All"
    }
  };
}

function buildGetItemsPayload(itemIds) {
  return {
    __type: "GetItemJsonRequest:#Exchange",
    Header: {
      __type: "JsonRequestHeaders:#Exchange",
      RequestServerVersion: "V2017_08_18",
      TimeZoneContext: {
        __type: "TimeZoneContext:#Exchange",
        TimeZoneDefinition: {
          __type: "TimeZoneDefinitionType:#Exchange",
          Id: "Eastern Standard Time"
        }
      }
    },
    Body: {
      __type: "GetItemRequest:#Exchange",
      ItemShape: {
        __type: "ItemResponseShape:#Exchange",
        BaseShape: "IdOnly",
        AddBlankTargetToLinks: true,
        BlockContentFromUnknownSenders: false,
        BlockExternalImagesIfSenderUntrusted: true,
        ClientSupportsIrm: true,
        CssScopeClassName: "rps_openclaw",
        FilterHtmlContent: true,
        FilterInlineSafetyTips: true,
        InlineImageCustomDataTemplate: "{id}",
        InlineImageUrlTemplate: "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAEALAAAAAABAAEAAAIBTAA7",
        MaximumBodySize: 2097152,
        MaximumRecipientsToReturn: 20,
        ImageProxyCapability: "OwaAndConnectorsProxy",
        AdditionalProperties: [{ __type: "PropertyUri:#Exchange", FieldURI: "CanDelete" }],
        InlineImageUrlOnLoadTemplate: ""
      },
      ItemIds: itemIds.map((itemId) => ({ __type: "ItemId:#Exchange", Id: itemId })),
      ShapeName: "ItemNormalizedBody"
    }
  };
}

export async function fetchMessageBodiesForEmails({
  statePath,
  inboxUrl,
  emails,
  viewport,
  timeoutMs = 45000
}) {
  const safeEmails = Array.isArray(emails) ? emails.filter(Boolean) : [];
  if (!safeEmails.length) {
    return [];
  }

  const storageState = await loadStorageState(statePath);
  const fallbackTokens = getSessionApiTokensFromStorageState(storageState);
  const fallbackAnchorMailbox = getCookieValue(storageState, "DefaultAnchorMailbox");
  const requestLimit = Math.max((viewport?.maxInboxItems || safeEmails.length || 1) * 2, 15);

  try {
    return await withOutlookPage({ statePath, inboxUrl, viewport, timeoutMs }, async (page) => {
      const liveStorageState = await page.context().storageState().catch(() => storageState);
      const liveTokens = getSessionApiTokensFromStorageState(liveStorageState);
      const outlookToken = liveTokens.outlook?.token || fallbackTokens.outlook?.token;
      const anchorMailbox = getCookieValue(liveStorageState, "DefaultAnchorMailbox") || fallbackAnchorMailbox;

      if (!outlookToken || !anchorMailbox) {
        return safeEmails.map((email) => ({
          ...email,
          summarySource: "preview",
          summaryText: summarizePreviewText(email.subject, email.preview),
          summaryError: "Missing refreshed Outlook session token."
        }));
      }

      const findResponse = await callOwaAction(page, {
        action: "FindItem",
        payload: buildFindInboxItemsPayload(requestLimit),
        token: outlookToken,
        anchorMailbox,
        requestNumber: 1
      });

      if (!findResponse.ok) {
        throw new Error(`FindItem failed with ${findResponse.status}`);
      }

      const inboxItems = (findResponse.json?.Body?.ResponseMessages?.Items?.[0]?.RootFolder?.Items || [])
        .map(normalizeFindItem)
        .filter((item) => item.itemId && (item.subject || item.from));
      const emailMatches = matchEmailsToItems(safeEmails, inboxItems);
      const matchedItemIds = Array.from(new Set(Array.from(emailMatches.values()).map((item) => item.itemId).filter(Boolean)));

      let bodiesById = new Map();
      if (matchedItemIds.length) {
        const getResponse = await callOwaAction(page, {
          action: "GetItem",
          payload: buildGetItemsPayload(matchedItemIds),
          token: outlookToken,
          anchorMailbox,
          requestNumber: 2
        });

        if (getResponse.ok) {
          const items = getResponse.json?.Body?.ResponseMessages?.Items?.[0]?.Items || [];
          bodiesById = new Map(
            items
              .map((item) => normalizeGetItem(item, getOwaPreview(item)))
              .filter((item) => item.itemId)
              .map((item) => [item.itemId, item])
          );
        }
      }

      return safeEmails.map((email) => {
        const matchKey = email.rawKey || email.messageId || `${email.subject}-${email.from}`;
        const matchedItem = emailMatches.get(matchKey);
        const body = matchedItem?.itemId ? bodiesById.get(matchedItem.itemId) : null;
        const previewText = compactWhitespace(email.preview);

        if (body?.summaryText) {
          return {
            ...email,
            summarySource: "owa-getitem",
            summaryText: body.summaryText,
            bodyText: body.bodyText
          };
        }

        if (matchedItem?.preview) {
          return {
            ...email,
            summarySource: "owa-finditem-preview",
            summaryText: summarizePreviewText(email.subject, matchedItem.preview)
          };
        }

        return {
          ...email,
          summarySource: "preview",
          summaryText: summarizePreviewText(email.subject, previewText)
        };
      });
    });
  } catch (error) {
    return safeEmails.map((email) => ({
      ...email,
      summarySource: "preview",
      summaryText: summarizePreviewText(email.subject, email.preview),
      summaryError: compactWhitespace(error?.message || error)
    }));
  }
}
