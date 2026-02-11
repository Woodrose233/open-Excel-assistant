import { AppSettings, DEFAULT_SETTINGS } from "./types";

const SETTINGS_KEY = "excel_ai_assistant_settings";

const isLocalStorageAvailable = () => {
  try {
    const test = "__storage_test__";
    localStorage.setItem(test, test);
    localStorage.removeItem(test);
    return true;
  } catch (e) {
    return false;
  }
};

export const loadSettings = (): AppSettings => {
  if (!isLocalStorageAvailable()) {
    console.warn("LocalStorage is not available. Settings will not be persisted.");
    return DEFAULT_SETTINGS;
  }
  const saved = localStorage.getItem(SETTINGS_KEY);
  if (saved) {
    try {
      return JSON.parse(saved);
    } catch (e) {
      console.error("Failed to parse settings", e);
    }
  }
  return DEFAULT_SETTINGS;
};

export const saveSettings = (settings: AppSettings): void => {
  if (isLocalStorageAvailable()) {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  }
};
