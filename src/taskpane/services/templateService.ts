import type { AppConfig } from "../types";
import { TEMPLATES } from "../types";

/**
 * Fetch a built-in HTML template by its id.
 * Templates are served from the /templates/ directory in the webpack output.
 */
export async function loadTemplate(templateId: string): Promise<string> {
  const template = TEMPLATES.find((t) => t.id === templateId);
  if (!template) throw new Error(`Unknown template: ${templateId}`);

  const response = await fetch(`/templates/${template.fileName}`);
  if (!response.ok) throw new Error(`Failed to load template: ${template.fileName}`);
  return response.text();
}

/**
 * Replace placeholders in an HTML template string with live values.
 *
 * Supported placeholders:
 *   [RETURN DATE] — formatted return date
 *   [ROLE]        — user's job title / role
 *   [SIGNATURE]   — auto-generated signature block (or removed if disabled)
 */
export function resolvePlaceholders(
  html: string,
  config: AppConfig,
  returnDate: Date | null,
  userEmail: string
): string {
  // [RETURN DATE]
  if (returnDate) {
    const formatted = returnDate.toLocaleDateString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
    });
    html = html.replace(/\[RETURN DATE\]/g, formatted);
  }

  // [ROLE]
  const role = config.role || "Azure Support Engineer";
  html = html.replace(/\[ROLE\]/g, role);

  // [SIGNATURE]
  if (config.includeSignature) {
    html = html.replace(/\[SIGNATURE\]/g, buildSignature(config, userEmail));
  } else {
    html = html.replace(/^\s*\[SIGNATURE\]\s*\r?\n?/gm, "");
  }

  return html;
}

/** Build the auto-generated signature HTML block. */
function buildSignature(config: AppConfig, userEmail: string): string {
  const displayName = config.fullName || userEmail.split("@")[0] || "User";

  const lines: string[] = [];
  lines.push(`<p><b>Best Regards,</b><br/>`);
  lines.push(`${displayName}</p>`);

  // Office detail parts
  const details: string[] = [];
  if (config.includeOfficeHours && config.startOfShift && config.endOfShift) {
    details.push(`${formatTime(config.startOfShift)} - ${formatTime(config.endOfShift)}`);
  }
  if (config.includeTimezone) {
    details.push(Intl.DateTimeFormat().resolvedOptions().timeZone);
  }
  if (config.includeWorkDays && config.workDays.length > 0) {
    details.push(config.workDays.join(", "));
  }
  if (details.length > 0) {
    lines.push(`<p style='color: #555; font-size: 10pt;'>${details.join(" | ")}</p>`);
  }

  // Email link
  if (userEmail) {
    lines.push(`<p><a href='mailto:${userEmail}'>${userEmail}</a></p>`);
  }

  return lines.join("\n");
}

/** Convert "HH:mm" (24h) to "h:mm AM/PM". */
function formatTime(time24: string): string {
  const [h, m] = time24.split(":").map(Number);
  const ampm = h >= 12 ? "PM" : "AM";
  const h12 = h > 12 ? h - 12 : h === 0 ? 12 : h;
  return `${h12}:${m.toString().padStart(2, "0")} ${ampm}`;
}
