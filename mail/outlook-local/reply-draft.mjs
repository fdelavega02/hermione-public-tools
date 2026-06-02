import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

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
      "Missing config.json. Copy config.example.json to config.json and update your Outlook settings first."
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

function compactWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function parseArgs(argv) {
  const result = {};

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      continue;
    }

    const key = token.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith("--")) {
      result[key] = "true";
      continue;
    }

    result[key] = next;
    index += 1;
  }

  return result;
}

function buildAssistantPrompt(email, guidance, style) {
  return [
    "# Reply Draft Request",
    "",
    "Use this as local context for drafting a reply. Do not send anything automatically.",
    "",
    "## Email summary",
    "",
    `- Reply ID: ${email.id}`,
    `- From: ${email.from || "Unknown"}`,
    `- Subject: ${email.subject || "(no subject)"}`,
    `- Received: ${email.receivedAt || "Unknown"}`,
    `- Unread: ${email.unread ? "yes" : "no"}`,
    `- Preview: ${email.preview || "(none captured)"}`,
    "",
    "## User guidance",
    "",
    guidance || "No guidance supplied.",
    "",
    "## Draft constraints",
    "",
    `- Tone: ${style.tone}`,
    `- Length: ${style.length}`,
    `- Signature: ${style.signature}`,
    "- The draft should stay clear and should not invent facts not present in the email or guidance.",
    "- This is a draft only. User approval is required before any send action.",
    ""
  ].join("\n");
}

function buildStarterDraft(email, guidance, style) {
  const greetingName = email.from ? email.from.split(/[<,(]/)[0].trim() : "";
  const greeting = greetingName ? `Hi ${greetingName},` : "Hello,";
  const toneLine =
    style.length === "short"
      ? "Thanks for the note."
      : "Thanks for reaching out. I reviewed your message.";
  const guidanceParagraph = guidance
    ? compactWhitespace(guidance)
    : "I wanted to follow up on your message and respond directly.";

  return [
    greeting,
    "",
    toneLine,
    "",
    guidanceParagraph,
    "",
    "Please let me know if you need anything else from me.",
    "",
    style.signature
  ].join("\n");
}

async function loadLatestInbox(outputDir) {
  const inboxPath = path.join(outputDir, "inbox", "latest.json");
  if (!(await exists(inboxPath))) {
    throw new Error("Missing inbox snapshot. Run `npm run sync` first.");
  }

  return JSON.parse(await fs.readFile(inboxPath, "utf8"));
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const config = await loadConfig();
  const outputDir = resolveProjectPath(config.outputDir || "./output");
  const inbox = await loadLatestInbox(outputDir);
  const emails = Array.isArray(inbox.emails) ? inbox.emails : [];

  if (emails.length === 0) {
    throw new Error("The latest inbox snapshot has no emails to draft against.");
  }

  const targetId = args["email-id"] || emails[0].id;
  const email = emails.find((entry) => entry.id === targetId);

  if (!email) {
    throw new Error(`Email ID "${targetId}" was not found in output/inbox/latest.json.`);
  }

  let guidance = args.idea || "";
  if (args["guidance-file"]) {
    guidance = await fs.readFile(resolveProjectPath(args["guidance-file"]), "utf8");
  }

  if (!compactWhitespace(guidance)) {
    throw new Error(
      "Provide reply direction with `--idea \"...\"` or `--guidance-file path/to/file.txt`."
    );
  }

  const style = {
    tone: args.tone || config.replyStyleDefaults?.tone || "clear",
    length: args.length || config.replyStyleDefaults?.length || "short",
    signature: args.signature || config.replyStyleDefaults?.signature || "Francisco"
  };

  const draftDir = path.join(outputDir, "replies", email.id);
  await ensureDir(draftDir);

  const request = {
    createdAt: new Date().toISOString(),
    sendApproved: false,
    email,
    guidance: compactWhitespace(guidance),
    style
  };

  const assistantPrompt = buildAssistantPrompt(email, request.guidance, style);
  const starterDraft = buildStarterDraft(email, request.guidance, style);
  const checklist = [
    "# Approval Checklist",
    "",
    "- Review the starter draft and edit wording as needed.",
    "- Confirm all facts, dates, and commitments are correct.",
    "- Keep `sendApproved` false until you explicitly want a future send script to act.",
    "- This starter scaffold does not send emails.",
    ""
  ].join("\n");

  await fs.writeFile(path.join(draftDir, "request.json"), JSON.stringify(request, null, 2));
  await fs.writeFile(path.join(draftDir, "assistant-prompt.md"), assistantPrompt);
  await fs.writeFile(path.join(draftDir, "starter-draft.txt"), starterDraft);
  await fs.writeFile(path.join(draftDir, "approval-checklist.md"), checklist);

  console.log(`Draft scaffold created for ${email.id}`);
  console.log(`Review files in ${draftDir}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
