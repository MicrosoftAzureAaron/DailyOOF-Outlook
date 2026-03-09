import { getAccessToken } from "./authService";
import type { MailboxSettings, UserProfile, AutoReplyStatus, DateTimeTimeZone } from "../types";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/** Generic Graph API fetch helper. Adds Bearer token and parses JSON. */
async function graphFetch<T>(url: string, init?: RequestInit): Promise<T> {
  const token = await getAccessToken();
  const response = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...init?.headers,
    },
  });
  if (!response.ok) {
    const errBody = await response.text();
    throw new Error(`Graph API ${response.status}: ${errBody}`);
  }
  return response.json();
}

// ─── User Profile ───────────────────────────────────────────────────────────

/** Get the signed-in user's profile (display name, email, job title). */
export async function getUserProfile(): Promise<UserProfile> {
  return graphFetch<UserProfile>(`${GRAPH_BASE}/me?$select=displayName,mail,userPrincipalName,jobTitle`);
}

// ─── Mailbox Settings (Auto-Reply) ─────────────────────────────────────────

/** Fetch the current auto-reply settings from Exchange Online. */
export async function getMailboxSettings(): Promise<MailboxSettings> {
  return graphFetch<MailboxSettings>(`${GRAPH_BASE}/me/mailboxSettings`);
}

/** Build an ISO datetime string for Graph from a Date. */
function toDateTimeTimeZone(date: Date): DateTimeTimeZone {
  return {
    dateTime: date.toISOString(),
    timeZone: "UTC",
  };
}

/** Set the auto-reply state (disabled / alwaysEnabled / scheduled). */
export async function setAutoReplyStatus(status: AutoReplyStatus): Promise<void> {
  await graphFetch(`${GRAPH_BASE}/me/mailboxSettings`, {
    method: "PATCH",
    body: JSON.stringify({
      automaticRepliesSetting: { status },
    }),
  });
}

/** Set a scheduled auto-reply with start/end times. */
export async function setScheduledAutoReply(
  startTime: Date,
  endTime: Date,
  internalMessage?: string,
  externalMessage?: string
): Promise<void> {
  const body: Record<string, unknown> = {
    automaticRepliesSetting: {
      status: "scheduled" as AutoReplyStatus,
      scheduledStartDateTime: toDateTimeTimeZone(startTime),
      scheduledEndDateTime: toDateTimeTimeZone(endTime),
      ...(internalMessage !== undefined && { internalReplyMessage: internalMessage }),
      ...(externalMessage !== undefined && { externalReplyMessage: externalMessage }),
    },
  };
  await graphFetch(`${GRAPH_BASE}/me/mailboxSettings`, {
    method: "PATCH",
    body: JSON.stringify(body),
  });
}

/** Set the auto-reply message body for internal, external, or both. */
export async function setAutoReplyMessage(
  message: string,
  scope: "internal" | "external" | "both"
): Promise<void> {
  const setting: Record<string, string> = {};
  if (scope === "internal" || scope === "both") {
    setting.internalReplyMessage = message;
  }
  if (scope === "external" || scope === "both") {
    setting.externalReplyMessage = message;
  }
  await graphFetch(`${GRAPH_BASE}/me/mailboxSettings`, {
    method: "PATCH",
    body: JSON.stringify({ automaticRepliesSetting: setting }),
  });
}

// ─── Schedule Calculation ───────────────────────────────────────────────────

/**
 * Calculate the OOF schedule window based on shift times and work days.
 * Mirrors the PowerShell logic: OOF is active from end-of-shift today
 * until start-of-shift on the next work day.
 */
export function calculateScheduleWindow(
  startOfShift: string,  // "HH:mm"
  endOfShift: string,    // "HH:mm"
  workDays: string[]     // e.g. ["Monday","Tuesday",...]
): { startTime: Date; endTime: Date } {
  const now = new Date();
  const dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  const [startH, startM] = startOfShift.split(":").map(Number);
  const [endH, endM] = endOfShift.split(":").map(Number);

  // OOF starts at end of today's shift
  const oofStart = new Date(now);
  oofStart.setHours(endH, endM, 0, 0);

  // Find next work day
  const todayName = dayNames[now.getDay()];
  const isWorkDay = workDays.includes(todayName);
  const beforeShift = now.getHours() < startH || (now.getHours() === startH && now.getMinutes() < startM);

  let daysAhead: number;
  if (isWorkDay && beforeShift) {
    daysAhead = 0;
  } else {
    daysAhead = 1;
    const check = new Date(now);
    check.setDate(check.getDate() + 1);
    while (!workDays.includes(dayNames[check.getDay()])) {
      daysAhead++;
      check.setDate(check.getDate() + 1);
    }
  }

  // OOF ends at start of next work day's shift
  const oofEnd = new Date(now);
  oofEnd.setDate(oofEnd.getDate() + daysAhead);
  oofEnd.setHours(startH, startM, 0, 0);

  return { startTime: oofStart, endTime: oofEnd };
}

/**
 * Calculate a vacation OOF window from now until the return date at shift start.
 */
export function calculateVacationWindow(
  returnDate: Date,
  startOfShift: string,
  endOfShift: string
): { startTime: Date; endTime: Date } {
  const [endH, endM] = endOfShift.split(":").map(Number);
  const [startH, startM] = startOfShift.split(":").map(Number);

  const oofStart = new Date();
  oofStart.setHours(endH, endM, 0, 0);

  const oofEnd = new Date(returnDate);
  oofEnd.setHours(startH, startM, 0, 0);

  return { startTime: oofStart, endTime: oofEnd };
}
