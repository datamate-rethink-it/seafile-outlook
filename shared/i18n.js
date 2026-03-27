/**
 * Lightweight i18n helper for the Seafile Outlook Add-in.
 *
 * Office.js add-ins don't have a built-in i18n API like Thunderbird's
 * browser.i18n. This module provides a simple equivalent using JSON
 * locale files loaded at startup.
 */

const I18N = {
  _strings: {},
  _locale: "en",

  /**
   * Initialize i18n by loading the appropriate locale file.
   * Detects locale from Office.js or falls back to navigator.language.
   */
  async init() {
    // Detect locale
    let locale = "en";
    try {
      if (typeof Office !== "undefined" && Office.context) {
        locale = Office.context.displayLanguage || "en";
      } else {
        locale = navigator.language || "en";
      }
    } catch { /* fallback to en */ }

    // Normalize locale: "de-DE" -> "de", "zh-CN" -> "zh_CN"
    locale = locale.replace("-", "_");
    const lang = locale.split("_")[0];

    // Try to load locale file (exact match first, then base language, then English)
    for (const candidate of [locale, lang, "en"]) {
      try {
        const resp = await fetch(`../_locales/${candidate}.json`);
        if (resp.ok) {
          const data = await resp.json();
          // Support both flat format { key: "value" } and Thunderbird format { key: { message: "value" } }
          for (const [key, val] of Object.entries(data)) {
            this._strings[key] = typeof val === "string" ? val : val.message || "";
          }
          this._locale = candidate;
          return;
        }
      } catch { /* try next */ }
    }
  },

  /**
   * Get a translated string by key.
   * @param {string} key
   * @returns {string} Translated string or the key itself as fallback
   */
  get(key) {
    return this._strings[key] || key;
  },

  /**
   * Apply translations to all elements with data-i18n attributes in the document.
   * Supports:
   * - data-i18n="key" → sets textContent
   * - data-i18n-placeholder="key" → sets placeholder attribute
   * - data-i18n-empty="key" → sets data-empty attribute (for CSS :empty::after)
   * - data-i18n-title="key" → sets title attribute
   */
  applyToDocument() {
    for (const el of document.querySelectorAll("[data-i18n]")) {
      const msg = this.get(el.getAttribute("data-i18n"));
      if (msg) el.textContent = msg;
    }
    for (const el of document.querySelectorAll("[data-i18n-placeholder]")) {
      const msg = this.get(el.dataset.i18nPlaceholder);
      if (msg) el.placeholder = msg;
    }
    for (const el of document.querySelectorAll("[data-i18n-empty]")) {
      const msg = this.get(el.dataset.i18nEmpty);
      if (msg) el.dataset.empty = msg;
    }
    for (const el of document.querySelectorAll("[data-i18n-title]")) {
      const msg = this.get(el.dataset.i18nTitle);
      if (msg) el.title = msg;
    }
  },
};
