/** Possible auto-reply states from Microsoft Graph. */
export type AutoReplyStatus = "disabled" | "alwaysEnabled" | "scheduled";

/** Auto-reply schedule boundary from Graph (ISO 8601 + timezone). */
export interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

/** The automaticRepliesSetting object returned by GET /me/mailboxSettings. */
export interface AutomaticRepliesSetting {
  status: AutoReplyStatus;
  externalAudience: "none" | "contactsOnly" | "all";
  externalReplyMessage: string;
  internalReplyMessage: string;
  scheduledStartDateTime: DateTimeTimeZone;
  scheduledEndDateTime: DateTimeTimeZone;
}

/** Subset of the Graph /me/mailboxSettings response we use. */
export interface MailboxSettings {
  automaticRepliesSetting: AutomaticRepliesSetting;
}

/** User profile fields we read from Graph /me. */
export interface UserProfile {
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle: string | null;
}

/** Persisted user configuration (stored in Office roaming settings). */
export interface AppConfig {
  startOfShift: string | null;   // ISO time string e.g. "09:00"
  endOfShift: string | null;     // ISO time string e.g. "18:00"
  workDays: string[];            // e.g. ["Monday","Tuesday",...]
  role: string;
  fullName: string;
  includeSignature: boolean;
  includeOfficeHours: boolean;
  includeWorkDays: boolean;
  includeTimezone: boolean;
}

/** Default configuration values. */
export const DEFAULT_CONFIG: AppConfig = {
  startOfShift: "09:00",
  endOfShift: "18:00",
  workDays: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
  role: "Azure Support Engineer",
  fullName: "",
  includeSignature: true,
  includeOfficeHours: true,
  includeWorkDays: true,
  includeTimezone: true,
};

/** Template metadata. */
export interface TemplateInfo {
  id: string;
  name: string;
  fileName: string;
  headerColor: string;
}

/** Built-in template definitions. */
export const TEMPLATES: TemplateInfo[] = [
  { id: "normal",   name: "Normal OOF",   fileName: "normal_oof.html",   headerColor: "#0078D4" },
  { id: "vacation", name: "Vacation OOF",  fileName: "vacation_oof.html", headerColor: "#2E7D32" },
  { id: "sick",     name: "Sick OOF",      fileName: "sick_oof.html",     headerColor: "#D32F2F" },
  { id: "holiday",  name: "Holiday OOF",   fileName: "holiday_oof.html",  headerColor: "#FF8F00" },
];
