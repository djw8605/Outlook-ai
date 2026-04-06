/* global Office */
import "./taskpane.css";
import { AIClient } from "../lib/ai-client";
import {
  getEmailBody,
  getEmailSubject,
  getEmailSender,
  getDraftBody,
  setDraftBody,
  insertReply,
} from "../lib/outlook-helper";

// ─── Constants ────────────────────────────────────────────────────────────────
const DEFAULT_BASE_URL = "https://api.openai.com/v1";
const DEFAULT_MODEL = "gpt-4o";

// ─── Settings helpers (roamingSettings) ───────────────────────────────────────
function loadSettings(): void {
  const settings = Office.context.roamingSettings;

  const apiKey = (settings.get("apiKey") as string) || "";
  const baseUrl = (settings.get("baseUrl") as string) || DEFAULT_BASE_URL;
  const model = (settings.get("model") as string) || DEFAULT_MODEL;

  (document.getElementById("api-key") as HTMLInputElement).value = apiKey;
  (document.getElementById("base-url") as HTMLInputElement).value = baseUrl;
  (document.getElementById("model") as HTMLInputElement).value = model;
}

function saveSettings(): void {
  const settings = Office.context.roamingSettings;
  const apiKey = (document.getElementById("api-key") as HTMLInputElement).value.trim();
  const baseUrl =
    (document.getElementById("base-url") as HTMLInputElement).value.trim() || DEFAULT_BASE_URL;
  const model =
    (document.getElementById("model") as HTMLInputElement).value.trim() || DEFAULT_MODEL;

  settings.set("apiKey", apiKey);
  settings.set("baseUrl", baseUrl);
  settings.set("model", model);

  settings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      showStatus("Settings saved.", "success");
    } else {
      showStatus("Failed to save settings: " + result.error.message, "error");
    }
  });
}

function getAIClient(): AIClient {
  const apiKey = (document.getElementById("api-key") as HTMLInputElement).value.trim();
  const baseUrl =
    (document.getElementById("base-url") as HTMLInputElement).value.trim() || DEFAULT_BASE_URL;
  const model =
    (document.getElementById("model") as HTMLInputElement).value.trim() || DEFAULT_MODEL;

  if (!apiKey) {
    throw new Error("Please enter your API key in Settings.");
  }

  return new AIClient({ apiKey, baseUrl, model });
}

// ─── UI helpers ───────────────────────────────────────────────────────────────
function showSpinner(visible: boolean): void {
  const el = document.getElementById("spinner") as HTMLElement;
  el.classList.toggle("hidden", !visible);
}

function setButtonsDisabled(disabled: boolean): void {
  ["btn-summarize", "btn-reply", "btn-improve"].forEach((id) => {
    (document.getElementById(id) as HTMLButtonElement).disabled = disabled;
  });
}

function showOutput(text: string): void {
  const area = document.getElementById("output-area") as HTMLElement;
  const copyBtn = document.getElementById("copy-btn") as HTMLButtonElement;

  area.textContent = text;
  area.classList.remove("hidden");
  copyBtn.classList.remove("hidden");
}

function showStatus(message: string, type: "error" | "success" | ""): void {
  const bar = document.getElementById("status-bar") as HTMLElement;
  bar.textContent = message;
  bar.className = type;
}

function clearStatus(): void {
  showStatus("", "");
}

// ─── Email context helper ─────────────────────────────────────────────────────
async function buildEmailContext(): Promise<string> {
  const subject = getEmailSubject();
  const sender = getEmailSender();
  const body = await getEmailBody();

  const parts: string[] = [];
  if (subject) parts.push(`Subject: ${subject}`);
  if (sender.displayName || sender.emailAddress) {
    parts.push(`From: ${sender.displayName} <${sender.emailAddress}>`);
  }
  if (body) parts.push(`\n${body}`);

  return parts.join("\n");
}

// ─── Button handlers ──────────────────────────────────────────────────────────
async function handleSummarize(): Promise<void> {
  clearStatus();
  showSpinner(true);
  setButtonsDisabled(true);

  try {
    const client = getAIClient();
    const emailText = await buildEmailContext();
    const summary = await client.summarize(emailText);
    showOutput(summary);
  } catch (err) {
    showStatus((err as Error).message, "error");
  } finally {
    showSpinner(false);
    setButtonsDisabled(false);
  }
}

async function handleReply(): Promise<void> {
  clearStatus();
  showSpinner(true);
  setButtonsDisabled(true);

  try {
    const client = getAIClient();

    let emailText: string;
    try {
      // Try read mode first
      emailText = await buildEmailContext();
    } catch {
      // Fall back to compose (draft) mode
      emailText = await getDraftBody();
    }

    const reply = await client.generateReply(emailText);
    showOutput(reply);
  } catch (err) {
    showStatus((err as Error).message, "error");
  } finally {
    showSpinner(false);
    setButtonsDisabled(false);
  }
}

async function handleImprove(): Promise<void> {
  clearStatus();
  showSpinner(true);
  setButtonsDisabled(true);

  try {
    const client = getAIClient();
    let draft: string;

    try {
      draft = await getDraftBody();
    } catch {
      // If not in compose mode, try reading the current email body
      draft = await getEmailBody();
    }

    if (!draft.trim()) {
      throw new Error("No text found to improve. Please open an email or draft.");
    }

    const improved = await client.improveText(draft);
    showOutput(improved);

    // Try to insert back into compose window (will fail silently in read mode)
    try {
      await setDraftBody(improved);
    } catch {
      // Not in compose mode — output is already shown in the panel
    }
  } catch (err) {
    showStatus((err as Error).message, "error");
  } finally {
    showSpinner(false);
    setButtonsDisabled(false);
  }
}

// ─── Copy to clipboard ────────────────────────────────────────────────────────
async function handleCopy(): Promise<void> {
  const area = document.getElementById("output-area") as HTMLElement;
  const text = area.textContent || "";
  try {
    await navigator.clipboard.writeText(text);
    showStatus("Copied to clipboard!", "success");
  } catch {
    showStatus("Failed to copy. Please select the text manually.", "error");
  }
}

// ─── Initialization ───────────────────────────────────────────────────────────
Office.onReady(() => {
  loadSettings();

  // Settings toggle
  document.getElementById("toggle-settings")!.addEventListener("click", () => {
    const panel = document.getElementById("settings-panel") as HTMLElement;
    panel.classList.toggle("hidden");
  });

  // Auto-save settings on change
  ["api-key", "base-url", "model"].forEach((id) => {
    document.getElementById(id)!.addEventListener("change", saveSettings);
  });

  // Action buttons
  document.getElementById("btn-summarize")!.addEventListener("click", handleSummarize);
  document.getElementById("btn-reply")!.addEventListener("click", handleReply);
  document.getElementById("btn-improve")!.addEventListener("click", handleImprove);

  // Copy button
  document.getElementById("copy-btn")!.addEventListener("click", handleCopy);
});
