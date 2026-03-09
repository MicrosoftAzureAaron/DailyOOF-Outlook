import type { AppConfig } from "../types";
import { DEFAULT_CONFIG } from "../types";

const SETTINGS_KEY = "dailyoof_config";

/**
 * Load the user's configuration from Office roaming settings.
 * Falls back to DEFAULT_CONFIG for any missing fields.
 */
export function loadConfig(): AppConfig {
  try {
    const raw = Office.context.roamingSettings.get(SETTINGS_KEY);
    if (raw) {
      return { ...DEFAULT_CONFIG, ...JSON.parse(raw) };
    }
  } catch {
    // First run or corrupted settings — use defaults
  }
  return { ...DEFAULT_CONFIG };
}

/**
 * Save the user's configuration to Office roaming settings.
 * Roaming settings sync across devices automatically.
 */
export function saveConfig(config: AppConfig): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.set(SETTINGS_KEY, JSON.stringify(config));
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error?.message || "Failed to save settings"));
      }
    });
  });
}
