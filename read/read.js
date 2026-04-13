/**
 * Read task pane for saving email attachments to Seafile.
 * Ported from the Seafile Thunderbird extension's save-attachments popup.
 */

const seafile = new SeafileAPI();

// --- DOM References ---
const loadingEl = document.getElementById("loading");
const noAttachmentsEl = document.getElementById("noAttachments");
const notConfiguredEl = document.getElementById("notConfigured");
const contentEl = document.getElementById("content");
const attachmentListEl = document.getElementById("attachmentList");
const repoSelectEl = document.getElementById("repoSelect");
const folderListEl = document.getElementById("folderList");
const currentPathEl = document.getElementById("currentPath");
const selectAllEl = document.getElementById("selectAll");
const saveBtn = document.getElementById("saveBtn");
const saveStatus = document.getElementById("saveStatus");
const accountSelectorEl = document.getElementById("accountSelector");
const accountSelectEl = document.getElementById("accountSelect");
const folderPicker = document.getElementById("folderPicker");
const folderFilterInput = document.getElementById("folderFilter");
const resetDefaultsEl = document.getElementById("resetDefaults");
const openSettingsBtn = document.getElementById("openSettings");
const setupBtn = document.getElementById("setupBtn");

// --- State ---
let currentAccountId = null;
let accountConfig = null;
let attachments = [];
let currentPath = "/";
let currentRepoId = null;

// --- Settings Navigation ---
function openSettings() {
  window.location.href = "../settings/settings.html";
}
openSettingsBtn.addEventListener("click", openSettings);
if (setupBtn) setupBtn.addEventListener("click", openSettings);

// --- Account Selector ---
accountSelectEl.addEventListener("change", async () => {
  currentAccountId = accountSelectEl.value;
  lastUsedAccount("read", currentAccountId);
  currentPath = "/";
  currentRepoId = null;
  repoSelectEl.innerHTML = "";
  folderListEl.innerHTML = "";
  await loadAccountData();
});

// --- Load Account Data ---
async function loadAccountData() {
  accountConfig = getAccountConfig(currentAccountId);
  if (!accountConfig || !accountConfig.apiToken) {
    loadingEl.style.display = "none";
    notConfiguredEl.style.display = "block";
    return;
  }

  const repos = await seafile.listRepos(accountConfig.serverUrl, accountConfig.apiToken);
  repoSelectEl.innerHTML = "";
  const unencrypted = repos.filter(r => !r.encrypted);
  for (const repo of unencrypted) {
    const option = document.createElement("option");
    option.value = repo.repo_id || repo.id;
    option.textContent = repo.repo_name || repo.name;
    repoSelectEl.appendChild(option);
  }

  const defaultRepoId = accountConfig.saveRepoId || accountConfig.repoId;
  if (defaultRepoId) repoSelectEl.value = defaultRepoId;
  currentRepoId = repoSelectEl.value;

  currentPath = accountConfig.savePath || "/";
  await navigateToFolder(currentPath);
}

// --- Folder Picker ---
const FOLDER_FILTER_THRESHOLD = 8;

async function navigateToFolder(path) {
  currentPath = path;
  currentPathEl.querySelector(".path-text").textContent = path;
  folderFilterInput.value = "";

  try {
    const entries = await seafile.listDir(accountConfig.serverUrl, accountConfig.apiToken, currentRepoId, path);
    const dirs = entries.filter(e => e.type === "dir");
    const fragment = document.createDocumentFragment();

    if (path !== "/") {
      const parentLi = document.createElement("li");
      const parentPath = path.substring(0, path.lastIndexOf("/")) || "/";
      parentLi.innerHTML = `<span class="folder-icon">${FILE_ICONS.folderUp}</span> ..`;
      parentLi.addEventListener("click", () => navigateToFolder(parentPath));
      fragment.appendChild(parentLi);
    }

    for (const dir of dirs) {
      const li = document.createElement("li");
      li.dataset.name = dir.name.toLowerCase();
      const dirPath = path === "/" ? `/${dir.name}` : `${path}/${dir.name}`;
      li.innerHTML = `<span class="folder-icon">${FILE_ICONS.folder}</span> ${escapeHtml(dir.name)}`;
      li.addEventListener("click", () => navigateToFolder(dirPath));
      fragment.appendChild(li);
    }

    folderListEl.replaceChildren(fragment);
    folderFilterInput.style.display = dirs.length > FOLDER_FILTER_THRESHOLD ? "block" : "none";
    updateResetLink();
  } catch (e) {
    console.error("Failed to list directory:", e);
  }
}

repoSelectEl.addEventListener("change", () => {
  currentRepoId = repoSelectEl.value;
  navigateToFolder("/");
});

// --- Folder Filter ---
folderFilterInput.addEventListener("input", () => {
  const query = folderFilterInput.value.toLowerCase().trim();
  for (const li of folderListEl.children) {
    if (!li.dataset.name) { li.style.display = ""; continue; }
    li.style.display = li.dataset.name.includes(query) ? "" : "none";
  }
});

// --- Reset to Defaults ---
function updateResetLink() {
  const defaultRepoId = accountConfig.saveRepoId || accountConfig.repoId;
  const defaultPath = accountConfig.savePath || "/";
  const differs = currentRepoId !== defaultRepoId || currentPath !== defaultPath;
  resetDefaultsEl.style.display = differs ? "inline" : "none";
}

resetDefaultsEl.addEventListener("click", (e) => {
  e.preventDefault();
  const defaultRepoId = accountConfig.saveRepoId || accountConfig.repoId;
  const defaultPath = accountConfig.savePath || "/";
  if (defaultRepoId) {
    repoSelectEl.value = defaultRepoId;
    currentRepoId = defaultRepoId;
  }
  navigateToFolder(defaultPath);
});

// --- Attachment Rendering ---
function fallbackFileName(att) {
  const ext = att.contentType ? { "text/plain": ".txt", "text/html": ".html", "image/png": ".png", "image/jpeg": ".jpg", "application/pdf": ".pdf", "application/zip": ".zip" }[att.contentType] || "" : "";
  return `attachment-${att.id.replace(/[^a-zA-Z0-9]/g, "-").substring(0, 20)}${ext}`;
}

function renderAttachments() {
  attachmentListEl.innerHTML = "";
  for (const att of attachments) {
    if (!att.name || att.name.trim() === "") {
      att.name = fallbackFileName(att);
    }
    if (!att.customName) att.customName = att.name;

    const li = document.createElement("li");
    const resetSvg = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg>`;
    li.innerHTML = `
      <input type="checkbox" class="att-checkbox" data-id="${escapeHtml(att.id)}" checked>
      <span class="att-name${att.customName !== att.name ? ' renamed' : ''}" contenteditable="true" data-id="${escapeHtml(att.id)}" title="${escapeHtml(att.name)}">${escapeHtml(att.customName)}</span>
      <span class="att-reset${att.customName !== att.name ? ' visible' : ''}" title="${escapeHtml(I18N.get("resetFileName") || "Reset to original name")}">${resetSvg}</span>
      <span class="att-size">${formatSize(att.size)}</span>
      <span class="att-status" data-att-status="${escapeHtml(att.id)}"></span>
    `;
    li.querySelector(".att-checkbox").addEventListener("change", syncSelectAll);

    // Rename support
    const nameEl = li.querySelector(".att-name");
    const resetEl = li.querySelector(".att-reset");

    function updateRenamed(isRenamed) {
      nameEl.classList.toggle("renamed", isRenamed);
      resetEl.classList.toggle("visible", isRenamed);
    }

    nameEl.addEventListener("focus", () => {
      const text = nameEl.textContent;
      const dotIndex = text.lastIndexOf(".");
      if (dotIndex > 0) {
        const range = document.createRange();
        const sel = window.getSelection();
        range.setStart(nameEl.firstChild, 0);
        range.setEnd(nameEl.firstChild, dotIndex);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    });
    nameEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); nameEl.blur(); }
      if (e.key === "Escape") { att.customName = att.name; nameEl.textContent = att.name; updateRenamed(false); nameEl.blur(); }
    });
    nameEl.addEventListener("blur", () => {
      const newName = nameEl.textContent.trim();
      if (newName && newName !== att.name) {
        att.customName = newName;
        updateRenamed(true);
      } else {
        att.customName = att.name;
        nameEl.textContent = att.name;
        updateRenamed(false);
      }
    });
    resetEl.addEventListener("click", () => {
      att.customName = att.name;
      nameEl.textContent = att.name;
      updateRenamed(false);
    });

    attachmentListEl.appendChild(li);
  }
}

function syncSelectAll() {
  const checkboxes = attachmentListEl.querySelectorAll(".att-checkbox:not(:disabled)");
  selectAllEl.checked = [...checkboxes].every(cb => cb.checked);
}

selectAllEl.addEventListener("change", () => {
  for (const cb of attachmentListEl.querySelectorAll(".att-checkbox")) {
    cb.checked = selectAllEl.checked;
  }
});

// --- Save Button ---
saveBtn.addEventListener("click", async () => {
  const selected = attachmentListEl.querySelectorAll(".att-checkbox:checked");
  if (selected.length === 0) return;

  saveBtn.disabled = true;
  saveStatus.className = "status";
  let errorCount = 0;

  const config = accountConfig;
  const repoId = currentRepoId;
  const targetDir = currentPath;
  const replace = !!config.saveReplaceExisting;

  for (const cb of selected) {
    const attId = cb.dataset.id;
    const statusEl = document.querySelector(`[data-att-status="${attId}"]`);
    const att = attachments.find(a => a.id === attId);
    statusEl.innerHTML = STATUS_ICONS.pending;

    try {
      const fileName = (att.customName || att.name || "").trim();
      if (!fileName || fileName === ".") {
        throw new Error("Invalid filename");
      }

      // Read attachment content via Office.js
      const content = await new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(attId, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
          else reject(new Error("Failed to read attachment"));
        });
      });

      // Convert base64 to Blob
      const byteChars = atob(content.content);
      const byteArray = new Uint8Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) {
        byteArray[i] = byteChars.charCodeAt(i);
      }
      const blob = new Blob([byteArray]);

      // Ensure target directory exists
      const exists = await seafile.dirExists(config.serverUrl, config.apiToken, repoId, targetDir);
      if (!exists) {
        await seafile.createDir(config.serverUrl, config.apiToken, repoId, targetDir);
      }

      // Upload
      const uploadLink = await seafile.getUploadLink(config.serverUrl, config.apiToken, repoId, targetDir);
      await seafile.uploadFile(uploadLink, config.apiToken, blob, fileName, targetDir, null, replace);

      statusEl.innerHTML = STATUS_ICONS.success;
      cb.disabled = true;
    } catch (e) {
      statusEl.innerHTML = STATUS_ICONS.error;
      statusEl.title = e.message;
      errorCount++;
      console.error(`Failed to upload ${att.name}:`, e);
    }
  }

  if (errorCount === 0) {
    showStatus(saveStatus, I18N.get("saveSuccess") || "All files saved!", "success");
  } else {
    showStatus(saveStatus, I18N.get("savePartialError") || "Some files failed to upload.", "error");
  }
  saveBtn.disabled = false;
});

// --- Initialize ---
Office.onReady(async () => {
  await I18N.init();
  I18N.applyToDocument();

  try {
    // Check for configured accounts
    const accounts = getAllConfiguredAccounts();
    if (accounts.length === 0) {
      loadingEl.style.display = "none";
      notConfiguredEl.style.display = "block";
      return;
    }

    // Account selector
    const lastUsed = lastUsedAccount("read");
    currentAccountId = accounts.find(a => a.accountId === lastUsed)?.accountId || accounts[0].accountId;

    if (accounts.length > 1) {
      accountSelectorEl.style.display = "block";
      for (const acc of accounts) {
        const option = document.createElement("option");
        option.value = acc.accountId;
        const host = getHostLabel(acc.serverUrl);
        option.textContent = acc.displayName ? `${acc.displayName} (${host})` : `${acc.username} (${host})`;
        accountSelectEl.appendChild(option);
      }
      accountSelectEl.value = currentAccountId;
    }

    // Read attachments from current email
    const item = Office.context.mailbox.item;
    loadingEl.style.display = "none";

    // In read mode, attachments is a synchronous property
    const rawAttachments = item.attachments || [];
    if (rawAttachments.length === 0) {
      noAttachmentsEl.style.display = "block";
      return;
    }

    attachments = rawAttachments
      .filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File)
      .map(a => ({ id: a.id, name: a.name, size: a.size, customName: a.name }));

    if (attachments.length === 0) {
      noAttachmentsEl.style.display = "block";
      return;
    }

    contentEl.style.display = "block";
    renderAttachments();
    loadAccountData();
    lastUsedAccount("read", currentAccountId);
  } catch (e) {
    loadingEl.style.display = "none";
    showStatus(saveStatus, `Error: ${e.message}`, "error");
    saveStatus.style.display = "block";
    console.error("Init error:", e);
  }
});
