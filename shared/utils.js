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

  let metaLines = "";
  if (params.fileSize) metaLines += `Size: ${escapeHtml(params.fileSize)}<br>`;
  metaLines += `Link: <a href="${safeLink}" style="color:#0060df;">${safeLink}</a>`;
  if (params.password) {
    metaLines += params.showPasswordInEmail
      ? `<br>Password: <code style="background:#f0f0f0;padding:2px 6px;border-radius:2px;display:inline-block;user-select:all;">${escapeHtml(params.password)}</code>`
      : `<br>Password protected (password will be sent separately)`;
  }
  if (params.expireDays) {
    const expiryDate = new Date();
    expiryDate.setDate(expiryDate.getDate() + params.expireDays);
    metaLines += `<br>Expires: ${expiryDate.toLocaleDateString()}`;
  }

  const logoSvg = `data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAwJSIgaGVpZ2h0PSIxMDAlIiB2aWV3Qm94PSIxMCAxIDU2IDUyIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPjxwYXRoIGQ9Ik0xMy44MDYsMjIuMTYxYzAsLTAuOTIgMC43NDYsLTEuNjU4IDEuNjY3LC0xLjY1OGMwLjQzOSwwIDAuODI5LDAuMTc0IDEuMTI4LDAuNDQ4Yy0wLjAwOCwtMC4xNDkgLTAuMDE3LC0wLjI5IC0wLjAxNywtMC40MzljMCwtMy4wNzYgMi40ODcsLTUuNTY0IDUuNTY0LC01LjU2NGMwLjc5NiwwIDEuNTUsMC4xNjYgMi4yMzksMC40NzNjLTAuMDA4LC0wLjE1OCAtMC4wMTcsLTAuMzE1IC0wLjAxNywtMC40NjRjMCwtNC42MSAzLjczMSwtOC4zNDEgOC4zNDEsLTguMzQxYzQuNTc3LDAgOC4yOTEsMy42OSA4LjM0MSw4LjI1OGMtMS41MDksMS4zMjcgLTIuNzExLDMuMDAxIC0zLjQ4Miw0Ljg5MmMtMS40OTIsLTAuOTI5IC0zLjI1LC0xLjQ4NCAtNS4xMzIsLTEuNDg0Yy0zLjg4LDAgLTcuMjMsMi4yMDYgLTguNzk3LDUuNTM5bC01Ljk0NiwwbC0yLjIyMiwwYy0wLjkyMSwtMC4wMTggLTEuNjY3LC0wLjczOSAtMS42NjcsLTEuNjZabTQzLjkyLC00Ljk4M2MtMi4wNjUsLTIuMDY1IC00LjkwOSwtMy4zNDEgLTguMDU5LC0zLjM0MWMtNS43NDYsMCAtMTAuNTA1LDQuMjU0IC0xMS4yODUsOS43OTJjLTEuNDE4LC0xLjg5IC0zLjY4MSwtMy4xMTggLTYuMjI3LC0zLjExOGMtNC4zMDMsMCAtNy43ODYsMy40OTEgLTcuNzg2LDcuNzg2YzAsMS4yNTIgMC4yOTgsMi40MjkgMC44MjEsMy40ODJjLTIuNzAzLDAuNTQ3IC00LjcxLDIuNjQ1IC00LjcxLDUuMTQxYzAsMi45MTkgMi43MzYsNS4yODIgNi4xMTksNS4yODJjMS40OTIsMCAyLjg2MSwtMC40NjQgMy45MjIsLTEuMjM1bDEyLjUxMiwtMTIuMzA1YzEuMzg1LC0xLjI2OSAzLjIyNSwtMi4wNCA1LjI0OCwtMi4wNGM0LjIyOSwwIDcuNjcsMy4zNzUgNy43ODYsNy41NzhjMCwwIDAsLTAuMDA4IC0wLjAwOCwtMC4wMDhjMC4wNjYsMS4yNDQgLTAuNTgsMi40OTYgLTEuNzc0LDMuMTg0Yy0xLjY2NywwLjk2MiAtMy43NTYsMC40NDggLTQuNjY4LC0xLjEzNmMtMC45MiwtMS41OTIgLTAuMzE1LC0zLjY1NyAxLjM1MiwtNC42MThjMC4zOSwtMC4yMjQgMC43OTYsLTAuMzY1IDEuMjExLC0wLjQzOWMtMC4zNTcsLTAuMDc1IC0wLjczLC0wLjEwOCAtMS4xMTEsLTAuMTA4Yy0zLjA2OCwwIC01LjU2NCwyLjQ4NyAtNS41NjQsNS41NjRjMCwzLjA3NiAyLjQ4Nyw1LjU2NCA1LjU2NCw1LjU2NGMwLjEzMywwIDAuMjc0LC0wLjAwOCAwLjQwNiwtMC4wMTdsLTAuMDA4LC0wLjAxN2wwLjE1OCwtMC4wMTdsMTAuOTIsMGwwLDAuMDQxYzMuMzQxLC0wLjE0OSA2LjUsLTMuMjI1IDYuNSwtNi45NzNjMCwtMy44MzkgLTMuMjgzLC02Ljk3MyAtNy4xMjIsLTYuOTczYy0wLjAwOCwwIC0wLjAwOCwwIC0wLjAxNywwYy0wLjYyMiwxLjExMSAtMS4zNzYsMS43NTggLTIuMjMsMi40NTRjMC44OTUsLTEuNjMzIDEuNDEsLTMuNDkxIDEuNDEsLTUuNDgxYy0wLjAxOSwtMy4xNDEgLTEuMjk2LC01Ljk3NyAtMy4zNiwtOC4wNDJaIiBzdHlsZT0iZmlsbDp1cmwoI19MaW5lYXIxKTtmaWxsLXJ1bGU6bm9uemVybzsiLz48ZGVmcz48bGluZWFyR3JhZGllbnQgaWQ9Il9MaW5lYXIxIiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMCIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiIGdyYWRpZW50VHJhbnNmb3JtPSJtYXRyaXgoMi4xNzkxNmUtMTUsMzUuNTg4NCwtMzUuNTg4NCwyLjE3OTE2ZS0xNSw0MS40MjIsNi42MjAxNSkiPjxzdG9wIG9mZnNldD0iMCIgc3R5bGU9InN0b3AtY29sb3I6I2ZhZDk1NjtzdG9wLW9wYWNpdHk6MSIvPjxzdG9wIG9mZnNldD0iMSIgc3R5bGU9InN0b3AtY29sb3I6I2ZmYTEwZjtzdG9wLW9wYWNpdHk6MSIvPjwvbGluZWFyR3JhZGllbnQ+PC9kZWZzPjwvc3ZnPg==`;

  return `<br><div style="padding:15px;background:#dae3f0;border-radius:4px;font-family:sans-serif;"><div style="font-size:13px;color:#333;margin-bottom:8px;">I've linked a file to this email:</div><div style="background:#fff;border:1px solid #c8cfd6;border-radius:4px;padding:10px 12px;"><table style="width:100%;border-collapse:collapse;"><tr><td style="width:28px;vertical-align:top;padding-right:8px;"><span style="font-size:20px;color:#7b8a99;">&#128206;</span></td><td style="vertical-align:top;font-size:12px;color:#555;"><a href="${safeLink}" style="color:#0060df;font-size:13px;text-decoration:underline;">${safeFileName}</a><br>${metaLines}</td><td style="width:50px;vertical-align:middle;text-align:center;"><img src="${logoSvg}" alt="Seafile" width="28" height="28" style="display:block;margin:0 auto 2px auto;"><div style="font-size:9px;color:#888;">Seafile</div></td></tr></table></div><div style="font-size:11px;color:#555;margin-top:6px;">Learn more about <a href="https://www.seafile.com" style="color:#0060df;">Seafile</a>.</div></div>`;
}

/**
 * Build a plain-text version of a share link (for plain-text emails).
 * @param {Object} params - Same as buildShareLinkHtml
 * @returns {string} Plain text string
 */
function buildShareLinkText(params) {
  let text = `\n${params.fileName}: ${params.link}`;
  if (params.fileSize) text += `\nSize: ${params.fileSize}`;
  if (params.password) {
    text += params.showPasswordInEmail
      ? `\nPassword: ${params.password}`
      : `\nPassword protected (password will be sent separately)`;
  }
  if (params.expireDays) text += `\nExpires in ${params.expireDays} days`;
  text += "\n";
  return text;
}
