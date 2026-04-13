/**
 * Shared utility functions for the Seafile Outlook Add-in.
 */

/**
 * Escape a string for safe insertion into HTML.
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

/**
 * Format file size for display.
 * @param {number} bytes
 * @returns {string}
 */
function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Extract hostname from a URL for display.
 * @param {string} url
 * @returns {string}
 */
function getHostLabel(url) {
  try { return new URL(url).hostname; } catch { return url; }
}

/**
 * Show a status message in a status element.
 * @param {HTMLElement} element
 * @param {string} message
 * @param {"success"|"error"|"info"} type
 */
function showStatus(element, message, type) {
  element.textContent = message;
  element.className = `status ${type}`;
  const closeBtn = document.createElement("button");
  closeBtn.className = "close-btn";
  closeBtn.textContent = "\u00D7";
  closeBtn.addEventListener("click", () => { element.className = "status"; });
  element.appendChild(closeBtn);
}

/**
 * Clear a status element.
 * @param {HTMLElement} element
 */
function clearStatus(element) {
  element.className = "status";
}

/**
 * Build the HTML template for a Seafile share link (for insertion into email body).
 * @param {Object} params
 * @param {string} params.link - Share link URL
 * @param {string} params.fileName - File name
 * @param {string} [params.fileSize] - Formatted file size
 * @param {string} [params.password] - Link password
 * @param {boolean} [params.showPasswordInEmail] - Whether to display password
 * @param {number} [params.expireDays] - Expiration in days
 * @returns {string} HTML string
 */
function buildShareLinkHtml(params) {
  const safeLink = escapeHtml(params.link);
  const safeFileName = escapeHtml(params.fileName);
  const i = (key, fallback) => escapeHtml(I18N.get(key) || fallback);

  let metaLines = "";
  if (params.fileSize) metaLines += `${i("emailSize", "Size:")} ${escapeHtml(params.fileSize)}<br>`;
  metaLines += `${i("emailLink", "Link:")} <a href="${safeLink}" style="color:#0060df;">${safeLink}</a>`;
  if (params.password) {
    metaLines += params.showPasswordInEmail
      ? `<br>${i("emailPassword", "Password:")} <code style="background:#f0f0f0;padding:2px 6px;border-radius:2px;display:inline-block;user-select:all;">${escapeHtml(params.password)}</code>`
      : `<br>${i("emailPasswordProtected", "Password protected (password will be sent separately)")}`;
  }
  if (params.expireDays) {
    const expiryDate = new Date();
    expiryDate.setDate(expiryDate.getDate() + params.expireDays);
    metaLines += `<br>${i("emailExpires", "Expires:")} ${expiryDate.toLocaleDateString()}`;
  }

  const logoUrl = "https://outlook.datamate.org/assets/icon-32.png";

  return `<br><div style="padding:15px;background:#dae3f0;border-radius:4px;font-family:sans-serif;"><div style="font-size:13px;color:#333;margin-bottom:8px;">${i("emailLinkedFile", "I've linked a file to this email:")}</div><div style="background:#fff;border:1px solid #c8cfd6;border-radius:4px;padding:10px 12px;"><table style="width:100%;border-collapse:collapse;"><tr><td style="width:28px;vertical-align:top;padding-right:8px;"><span style="font-size:20px;color:#7b8a99;">&#128206;</span></td><td style="vertical-align:top;font-size:12px;color:#555;"><a href="${safeLink}" style="color:#0060df;font-size:13px;text-decoration:underline;">${safeFileName}</a><br>${metaLines}</td><td style="width:50px;vertical-align:middle;text-align:center;"><img src="${logoUrl}" alt="Seafile" width="28" height="28" style="display:block;margin:0 auto 2px auto;"><div style="font-size:9px;color:#888;">Seafile</div></td></tr></table></div><div style="font-size:11px;color:#555;margin-top:6px;">${i("emailLearnMore", "Learn more about")} <a href="https://www.seafile.com" style="color:#0060df;">Seafile</a>.</div></div>`;
}

/**
 * Build a plain-text version of a share link (for plain-text emails).
 * @param {Object} params - Same as buildShareLinkHtml
 * @returns {string} Plain text string
 */
function buildShareLinkText(params) {
  const t = (key, fallback) => I18N.get(key) || fallback;
  let text = `\n${params.fileName}: ${params.link}`;
  if (params.fileSize) text += `\n${t("emailSize", "Size:")} ${params.fileSize}`;
  if (params.password) {
    text += params.showPasswordInEmail
      ? `\n${t("emailPassword", "Password:")} ${params.password}`
      : `\n${t("emailPasswordProtected", "Password protected (password will be sent separately)")}`;
  }
  if (params.expireDays) {
    const msg = t("emailExpiresInDays", `Expires in ${params.expireDays} days`);
    text += `\n${msg.replace("$1", params.expireDays)}`;
  }
  text += "\n";
  return text;
}

/**
 * Detect the Outlook theme and apply the corresponding data-theme attribute.
 * Uses Office.context.officeTheme.bodyBackgroundColor when available,
 * falls back to prefers-color-scheme media query.
 */
function applyOfficeTheme() {
  let isDark = false;
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.officeTheme) {
      const bg = Office.context.officeTheme.bodyBackgroundColor;
      if (bg) {
        const hex = bg.replace("#", "");
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
        isDark = luminance < 0.5;
      }
    } else {
      isDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
    }
  } catch {
    isDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  }
  document.documentElement.dataset.theme = isDark ? "dark" : "light";
}
