/**
 * Compose task pane for the Seafile Outlook Add-in.
 * Handles two modes:
 * 1. Upload: Upload email attachments or local files to Seafile, insert share links.
 * 2. Insert: Browse Seafile libraries and insert links to existing files.
 */

const seafile = new SeafileAPI();

// --- DOM References ---
const notConfiguredEl = document.getElementById("notConfigured");
const mainContent = document.getElementById("mainContent");
const setupBtn = document.getElementById("setupBtn");
const openSettingsBtn = document.getElementById("openSettings");
const accountSelectorEl = document.getElementById("accountSelector");
const accountSelectEl = document.getElementById("accountSelect");

// Upload mode
const emailAttachmentsEl = document.getElementById("emailAttachments");
const attachmentListEl = document.getElementById("attachmentList");
const uploadAttBtn = document.getElementById("uploadAttBtn");
const uploadArea = document.getElementById("uploadArea");
const fileInput = document.getElementById("fileInput");
const uploadProgress = document.getElementById("uploadProgress");
const uploadStatus = document.getElementById("uploadStatus");

// Insert mode
const repoSelectEl = document.getElementById("repoSelect");
const fileListEl = document.getElementById("fileList");
const currentPathEl = document.getElementById("currentPath");
const fileFilterInput = document.getElementById("fileFilter");
const detailView = document.getElementById("detailView");
const backBtn = document.getElementById("backBtn");
const selectedFileIcon = document.getElementById("selectedFileIcon");
const selectedFileName = document.getElementById("selectedFileName");
const selectedFileSize = document.getElementById("selectedFileSize");
const existingLinkBar = document.getElementById("existingLinkBar");
const useExistingBtn = document.getElementById("useExistingBtn");
const createNewBtn = document.getElementById("createNewBtn");
const linkOptions = document.getElementById("linkOptions");
const linkPasswordInput = document.getElementById("linkPassword");
const linkExpireDaysInput = document.getElementById("linkExpireDays");
const generatePasswordBtn = document.getElementById("generatePasswordBtn");
const showPasswordInEmailInput = document.getElementById("showPasswordInEmail");
const showPasswordLabel = document.getElementById("showPasswordLabel");
const insertBtn = document.getElementById("insertBtn");
const insertStatus = document.getElementById("insertStatus");

// --- State ---
let currentAccountId = null;
let accountConfig = null;
let currentPath = "/";
let currentRepoId = null;
let selectedFilePath = null;
let selectedFileObj = null;
let existingLink = null;

const FILE_FILTER_THRESHOLD = 8;

// --- Mode Switching ---
for (const tab of document.querySelectorAll(".mode-tab")) {
  tab.addEventListener("click", () => {
    document.querySelector(".mode-tab.active").classList.remove("active");
    document.querySelector(".mode-content.active").classList.remove("active");
    tab.classList.add("active");
    document.getElementById(`mode-${tab.dataset.mode}`).classList.add("active");
  });
}

// --- Settings Navigation ---
function openSettings() {
  // Navigate to settings page in the same task pane
  window.location.href = "../settings/settings.html";
}

openSettingsBtn.addEventListener("click", openSettings);
if (setupBtn) setupBtn.addEventListener("click", openSettings);

// --- Account Selector ---
function initAccountSelector() {
  const accounts = getAllConfiguredAccounts();
  if (accounts.length === 0) {
    notConfiguredEl.style.display = "block";
    mainContent.style.display = "none";
    return false;
  }

  notConfiguredEl.style.display = "none";
  mainContent.style.display = "block";

  if (accounts.length > 1) {
    accountSelectorEl.style.display = "block";
    accountSelectEl.innerHTML = "";
    for (const acc of accounts) {
      const option = document.createElement("option");
      option.value = acc.accountId;
      const host = getHostLabel(acc.serverUrl);
      option.textContent = acc.displayName ? `${acc.displayName} (${host})` : `${acc.username} (${host})`;
      accountSelectEl.appendChild(option);
    }
  }

  const lastUsed = lastUsedAccount("compose");
  currentAccountId = accounts.find(a => a.accountId === lastUsed)?.accountId || accounts[0].accountId;
  if (accounts.length > 1) accountSelectEl.value = currentAccountId;

  return true;
}

accountSelectEl.addEventListener("change", async () => {
  currentAccountId = accountSelectEl.value;
  lastUsedAccount("compose", currentAccountId);
  await loadAccountData();
});

// --- Load Account Data ---
async function loadAccountData() {
  accountConfig = getAccountConfig(currentAccountId);
  if (!accountConfig || !accountConfig.apiToken) {
    notConfiguredEl.style.display = "block";
    mainContent.style.display = "none";
    return;
  }

  // Load repos for insert-link mode
  await loadRepos();
  await navigateToFolder("/");

  // Load email attachments for upload mode
  loadEmailAttachments();
}

// --- Upload Mode: Email Attachments ---
function loadEmailAttachments() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  item.getAttachmentsAsync((result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const attachments = result.value.filter(a => a.attachmentType === Office.MailboxEnums.AttachmentType.File);
    if (attachments.length === 0) return;

    emailAttachmentsEl.style.display = "block";
    attachmentListEl.innerHTML = "";

    for (const att of attachments) {
      const li = document.createElement("li");
      li.innerHTML = `
        <input type="checkbox" class="att-checkbox" data-id="${escapeHtml(att.id)}" checked>
        <span class="file-icon">${getFileIcon(att.name)}</span>
        <span class="file-name">${escapeHtml(att.name)}</span>
        <span class="att-size">${formatSize(att.size)}</span>
        <span class="att-status" data-att-status="${escapeHtml(att.id)}"></span>
      `;
      attachmentListEl.appendChild(li);
    }
  });
}

// --- Upload Mode: Upload Email Attachments ---
uploadAttBtn.addEventListener("click", async () => {
  const selected = attachmentListEl.querySelectorAll(".att-checkbox:checked");
  if (selected.length === 0) return;

  uploadAttBtn.disabled = true;
  uploadStatus.className = "status";

  const config = accountConfig;
  const uploadPath = config.uploadPath || "/";

  for (const cb of selected) {
    const attId = cb.dataset.id;
    const statusEl = document.querySelector(`[data-att-status="${attId}"]`);
    const li = cb.closest("li");
    const fileName = li.querySelector(".file-name").textContent;
    statusEl.innerHTML = STATUS_ICONS.pending;

    try {
      // Read attachment content
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

      // Ensure upload dir exists
      const exists = await seafile.dirExists(config.serverUrl, config.apiToken, config.repoId, uploadPath);
      if (!exists) {
        await seafile.createDir(config.serverUrl, config.apiToken, config.repoId, uploadPath);
      }

      // Upload
      const uploadLink = await seafile.getUploadLink(config.serverUrl, config.apiToken, config.repoId, uploadPath);
      await seafile.uploadFile(uploadLink, config.apiToken, blob, fileName, uploadPath);

      // Create share link
      const filePath = `${uploadPath}/${fileName}`;
      const existingLinks = await seafile.getShareLinks(config.serverUrl, config.apiToken, config.repoId, filePath);
      for (const old of existingLinks) {
        const oldToken = old.token || extractTokenFromUrl(old.link);
        if (oldToken) await seafile.deleteShareLink(config.serverUrl, config.apiToken, oldToken);
      }

      const pw = resolvePassword(config.fileLinkPasswordMode || "none", config.fileLinkPasswordLength || 12, config.fileLinkCustomPassword || "");
      const expDays = config.fileLinkExpireDays || 0;
      const shareResult = await seafile.createShareLink(config.serverUrl, config.apiToken, config.repoId, filePath, {
        password: pw || undefined,
        expireDays: expDays || undefined,
      });

      // Insert link into email body
      await insertShareLinkIntoBody({
        link: shareResult.link,
        fileName,
        password: pw,
        showPasswordInEmail: config.showPasswordInEmail !== false,
        expireDays: expDays,
      });

      // Remove original attachment
      await new Promise((resolve, reject) => {
        Office.context.mailbox.item.removeAttachmentAsync(attId, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error("Failed to remove attachment"));
        });
      });

      statusEl.innerHTML = STATUS_ICONS.success;
      cb.disabled = true;
    } catch (e) {
      statusEl.innerHTML = STATUS_ICONS.error;
      statusEl.title = e.message;
      console.error(`Failed to upload ${fileName}:`, e);
    }
  }

  uploadAttBtn.disabled = false;
});

// --- Upload Mode: File Picker / Drag & Drop ---
uploadArea.addEventListener("click", () => fileInput.click());
uploadArea.addEventListener("dragover", (e) => { e.preventDefault(); uploadArea.classList.add("dragover"); });
uploadArea.addEventListener("dragleave", () => uploadArea.classList.remove("dragover"));
uploadArea.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadArea.classList.remove("dragover");
  if (e.dataTransfer.files.length) handleFileUpload(e.dataTransfer.files);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files.length) handleFileUpload(fileInput.files);
  fileInput.value = "";
});

async function handleFileUpload(files) {
  const config = accountConfig;
  if (!config || !config.apiToken) return;

  const uploadPath = config.uploadPath || "/";
  uploadProgress.style.display = "block";
  uploadStatus.className = "status";

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    uploadProgress.textContent = `Uploading ${i + 1}/${files.length}: ${file.name}...`;

    try {
      const exists = await seafile.dirExists(config.serverUrl, config.apiToken, config.repoId, uploadPath);
      if (!exists) {
        await seafile.createDir(config.serverUrl, config.apiToken, config.repoId, uploadPath);
      }

      const uploadLink = await seafile.getUploadLink(config.serverUrl, config.apiToken, config.repoId, uploadPath);
      await seafile.uploadFile(uploadLink, config.apiToken, file, file.name, uploadPath);

      const filePath = `${uploadPath}/${file.name}`;
      const pw = resolvePassword(config.fileLinkPasswordMode || "none", config.fileLinkPasswordLength || 12, config.fileLinkCustomPassword || "");
      const expDays = config.fileLinkExpireDays || 0;
      const shareResult = await seafile.createShareLink(config.serverUrl, config.apiToken, config.repoId, filePath, {
        password: pw || undefined,
        expireDays: expDays || undefined,
      });

      await insertShareLinkIntoBody({
        link: shareResult.link,
        fileName: file.name,
        fileSize: formatSize(file.size),
        password: pw,
        showPasswordInEmail: config.showPasswordInEmail !== false,
        expireDays: expDays,
      });
    } catch (e) {
      showStatus(uploadStatus, `Error uploading ${file.name}: ${e.message}`, "error");
      console.error(`Failed to upload ${file.name}:`, e);
    }
  }

  uploadProgress.style.display = "none";
  if (uploadStatus.className === "status") {
    showStatus(uploadStatus, I18N.get("uploadSuccess") || "Files uploaded and links inserted!", "success");
  }
}

// --- Insert Link Mode ---
async function loadRepos() {
  const repos = await seafile.listRepos(accountConfig.serverUrl, accountConfig.apiToken);
  repoSelectEl.innerHTML = "";
  const unencrypted = repos.filter(r => !r.encrypted);
  for (const repo of unencrypted) {
    const option = document.createElement("option");
    option.value = repo.repo_id || repo.id;
    option.textContent = repo.repo_name || repo.name;
    repoSelectEl.appendChild(option);
  }
  const defaultRepoId = accountConfig.repoId;
  if (defaultRepoId) repoSelectEl.value = defaultRepoId;
  currentRepoId = repoSelectEl.value;
}

async function navigateToFolder(path) {
  currentPath = path;
  currentPathEl.textContent = path;
  fileFilterInput.value = "";

  try {
    const entries = await seafile.listDir(accountConfig.serverUrl, accountConfig.apiToken, currentRepoId, path);
    const fragment = document.createDocumentFragment();

    if (path !== "/") {
      const parentLi = document.createElement("li");
      const parentPath = path.substring(0, path.lastIndexOf("/")) || "/";
      parentLi.innerHTML = `<span class="file-icon">${FILE_ICONS.folderUp}</span><span class="file-name">..</span>`;
      parentLi.addEventListener("click", () => navigateToFolder(parentPath));
      fragment.appendChild(parentLi);
    }

    const dirs = entries.filter(e => e.type === "dir");
    const files = entries.filter(e => e.type === "file");

    for (const dir of dirs) {
      const li = document.createElement("li");
      li.dataset.name = dir.name.toLowerCase();
      li.dataset.type = "dir";
      const dirPath = path === "/" ? `/${dir.name}` : `${path}/${dir.name}`;
      li.innerHTML = `<span class="file-icon">${FILE_ICONS.folder}</span><span class="file-name">${escapeHtml(dir.name)}</span>`;
      li.addEventListener("click", () => navigateToFolder(dirPath));
      fragment.appendChild(li);
    }

    for (const file of files) {
      const li = document.createElement("li");
      li.dataset.name = file.name.toLowerCase();
      li.dataset.type = "file";
      const filePath = path === "/" ? `/${file.name}` : `${path}/${file.name}`;
      li.innerHTML = `
        <span class="file-icon">${getFileIcon(file.name)}</span>
        <span class="file-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
        <span class="file-size">${formatSize(file.size)}</span>
      `;
      li.addEventListener("click", () => showDetailView(file, filePath));
      fragment.appendChild(li);
    }

    fileListEl.innerHTML = "";
    fileListEl.appendChild(fragment);
    fileFilterInput.style.display = entries.length > FILE_FILTER_THRESHOLD ? "block" : "none";
  } catch (e) {
    showStatus(insertStatus, `Error: ${e.message}`, "error");
    console.error("Failed to list directory:", e);
  }
}

fileFilterInput.addEventListener("input", () => {
  const query = fileFilterInput.value.toLowerCase().trim();
  for (const li of fileListEl.children) {
    if (!li.dataset.name) { li.style.display = ""; continue; }
    li.style.display = li.dataset.name.includes(query) ? "" : "none";
  }
});

repoSelectEl.addEventListener("change", () => {
  currentRepoId = repoSelectEl.value;
  navigateToFolder("/");
});

// --- Detail View ---
function showBrowseView() {
  detailView.classList.remove("active");
  document.querySelector("#mode-insert .file-browser").style.display = "";
  document.querySelector("#mode-insert .form-group").style.display = "";
  document.querySelector("#mode-insert .current-path").style.display = "";
  document.querySelector("#mode-insert label[data-i18n='targetFolder']").style.display = "";
  clearStatus(insertStatus);
  selectedFilePath = null;
  existingLink = null;
}

async function showDetailView(file, filePath) {
  selectedFilePath = filePath;
  selectedFileObj = file;
  existingLink = null;
  insertBtn.disabled = false;

  selectedFileIcon.innerHTML = getFileIcon(file.name);
  selectedFileName.textContent = file.name;
  selectedFileSize.textContent = formatSize(file.size);

  // Pre-fill defaults
  if (!accountConfig.skipLinkOptions) {
    linkPasswordInput.value = resolvePassword(accountConfig.sharePasswordMode || "none", accountConfig.sharePasswordLength || 12, accountConfig.shareCustomPassword || "");
    linkExpireDaysInput.value = accountConfig.shareExpireDays || 0;
    showPasswordInEmailInput.checked = accountConfig.showPasswordInEmail !== false;
    updatePasswordCheckboxVisibility();
  }

  // Hide browse, show detail
  document.querySelector("#mode-insert .file-browser").style.display = "none";
  document.querySelector("#mode-insert .form-group").style.display = "none";
  document.querySelector("#mode-insert .current-path").style.display = "none";
  document.querySelector("#mode-insert label[data-i18n='targetFolder']").style.display = "none";
  detailView.classList.add("active");
  clearStatus(insertStatus);

  if (accountConfig.skipLinkOptions) {
    const pw = resolvePassword(accountConfig.sharePasswordMode || "none", accountConfig.sharePasswordLength || 12, accountConfig.shareCustomPassword || "");
    await doInsert(pw, accountConfig.shareExpireDays || 0, null, accountConfig.showPasswordInEmail !== false);
    return;
  }

  existingLinkBar.classList.remove("visible");
  linkOptions.style.display = "block";

  try {
    const result = await seafile.getShareLinks(accountConfig.serverUrl, accountConfig.apiToken, currentRepoId, filePath);
    if (result && result.length > 0) {
      existingLink = result[0];
      existingLinkBar.classList.add("visible");
      linkOptions.style.display = "none";
    }
  } catch (e) {
    console.error("Failed to check existing links:", e);
  }
}

async function doInsert(password, expireDays, linkUrl, showPw) {
  insertBtn.disabled = true;
  clearStatus(insertStatus);
  const fileName = selectedFilePath.split("/").pop();
  const fileSize = selectedFileObj ? formatSize(selectedFileObj.size) : null;

  try {
    if (!linkUrl) {
      const shareResult = await seafile.createShareLink(
        accountConfig.serverUrl, accountConfig.apiToken,
        currentRepoId, selectedFilePath,
        { password: password || undefined, expireDays: expireDays || undefined }
      );
      linkUrl = shareResult.link;
    }

    await insertShareLinkIntoBody({
      link: linkUrl, fileName, fileSize,
      password: password || "",
      showPasswordInEmail: !!showPw,
      expireDays: expireDays || 0,
    });

    showStatus(insertStatus, I18N.get("linkInserted") || "Link inserted!", "success");
    setTimeout(() => showBrowseView(), 1200);
  } catch (e) {
    showStatus(insertStatus, `Error: ${e.message}`, "error");
    insertBtn.disabled = false;
    console.error("Failed to insert link:", e);
  }
}

// Event handlers for detail view
backBtn.addEventListener("click", showBrowseView);
insertBtn.addEventListener("click", () => {
  const password = linkPasswordInput.value.trim();
  const expireDays = Math.max(0, parseInt(linkExpireDaysInput.value, 10) || 0);
  doInsert(password, expireDays, null, showPasswordInEmailInput.checked);
});
useExistingBtn.addEventListener("click", () => {
  if (existingLink) doInsert(null, null, existingLink.link, false);
});
createNewBtn.addEventListener("click", () => {
  existingLinkBar.classList.remove("visible");
  linkOptions.style.display = "block";
});
generatePasswordBtn.addEventListener("click", () => {
  linkPasswordInput.value = generateRandomPassword();
  updatePasswordCheckboxVisibility();
});
linkPasswordInput.addEventListener("input", updatePasswordCheckboxVisibility);
linkExpireDaysInput.addEventListener("input", () => {
  linkExpireDaysInput.value = linkExpireDaysInput.value.replace(/[^0-9]/g, "");
});

function updatePasswordCheckboxVisibility() {
  showPasswordLabel.style.display = linkPasswordInput.value.trim() ? "flex" : "none";
}

// --- Insert Share Link into Email Body (Office.js) ---
async function insertShareLinkIntoBody(params) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;

    item.body.getTypeAsync((typeResult) => {
      if (typeResult.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error("Failed to get body type"));
      }

      const isHtml = typeResult.value === Office.CoercionType.Html;
      const content = isHtml ? buildShareLinkHtml(params) : buildShareLinkText(params);
      const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;

      item.body.setSelectedDataAsync(content, { coercionType }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "Failed to insert content"));
        }
      });
    });
  });
}

// --- Initialize ---
Office.onReady(async () => {
  applyOfficeTheme();
  await I18N.init();
  I18N.applyToDocument();

  if (!initAccountSelector()) return;

  lastUsedAccount("compose", currentAccountId);
  await loadAccountData();

  // Check URL params for initial mode
  const params = new URLSearchParams(window.location.search);
  if (params.get("mode") === "insert") {
    document.querySelector('.mode-tab[data-mode="insert"]').click();
  }
});
