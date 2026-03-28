/**
 * Settings page for the Seafile Outlook Add-in.
 * Manages account connections, library selection, and default preferences.
 * Ported from the Seafile Thunderbird extension's management page.
 */

const seafile = new SeafileAPI();

// --- DOM References ---
const accountSelectEl = document.getElementById("accountSelect");
const addAccountBtn = document.getElementById("addAccountBtn");
const removeAccountBtn = document.getElementById("removeAccountBtn");
const loginForm = document.getElementById("loginForm");
const connectedInfo = document.getElementById("connectedInfo");
const connectedServer = document.getElementById("connectedServer");
const connectedUser = document.getElementById("connectedUser");
const connectedMethod = document.getElementById("connectedMethod");
const disconnectBtn = document.getElementById("disconnectBtn");
const serverUrlInput = document.getElementById("serverUrl");
const usernameInput = document.getElementById("username");
const passwordInput = document.getElementById("password");
const otpInput = document.getElementById("otp");
const connectBtn = document.getElementById("connectBtn");
const connectStatus = document.getElementById("connectStatus");
const repoSelect = document.getElementById("repoSelect");
const uploadPathEl = document.getElementById("uploadPath");
const uploadFolderList = document.getElementById("uploadFolderList");
const saveRepoSelect = document.getElementById("saveRepoSelect");
const savePathEl = document.getElementById("savePath");
const saveFolderList = document.getElementById("saveFolderList");
const sharePasswordModeSelect = document.getElementById("sharePasswordMode");
const sharePasswordLengthInput = document.getElementById("sharePasswordLength");
const shareCustomPasswordInput = document.getElementById("shareCustomPassword");
const shareExpireDaysInput = document.getElementById("shareExpireDays");
const showPasswordInEmailInput = document.getElementById("showPasswordInEmail");
const skipLinkOptionsInput = document.getElementById("skipLinkOptions");
const saveReplaceExistingInput = document.getElementById("saveReplaceExisting");
const fileLinkPasswordModeSelect = document.getElementById("fileLinkPasswordMode");
const fileLinkPasswordLengthInput = document.getElementById("fileLinkPasswordLength");
const fileLinkCustomPasswordInput = document.getElementById("fileLinkCustomPassword");
const fileLinkExpireDaysInput = document.getElementById("fileLinkExpireDays");
const ssoBtn = document.getElementById("ssoBtn");
const ssoStatus = document.getElementById("ssoStatus");
const uploadFolderPicker = document.getElementById("uploadFolderPicker");
const saveFolderPicker = document.getElementById("saveFolderPicker");

// --- State ---
let currentAccountId = null;
let uploadCurrentPath = "/";
let saveCurrentPath = "/";
let ssoPollingInterval = null;

// --- Tab Switching ---
function switchTab(tabName) {
  const tab = document.querySelector(`.tab[data-tab="${tabName}"]`);
  if (tab.classList.contains("disabled")) return;
  document.querySelector(".tab.active").classList.remove("active");
  document.querySelector(".tab-content.active").classList.remove("active");
  tab.classList.add("active");
  document.getElementById(`tab-${tabName}`).classList.add("active");
  if (["filelink", "sharing", "saving"].includes(tabName)) {
    refreshRepos();
  }
}

for (const tab of document.querySelectorAll(".tab")) {
  tab.addEventListener("click", () => switchTab(tab.dataset.tab));
}

function enableSettingsTabs() {
  document.querySelector('.tab[data-tab="filelink"]').classList.remove("disabled");
  document.querySelector('.tab[data-tab="sharing"]').classList.remove("disabled");
  document.querySelector('.tab[data-tab="saving"]').classList.remove("disabled");
}

function disableSettingsTabs() {
  document.querySelector('.tab[data-tab="filelink"]').classList.add("disabled");
  document.querySelector('.tab[data-tab="sharing"]').classList.add("disabled");
  document.querySelector('.tab[data-tab="saving"]').classList.add("disabled");
}

// --- Account Selector ---
function refreshAccountSelector() {
  const ids = getAccountIds();
  accountSelectEl.innerHTML = "";
  for (const id of ids) {
    const config = getAccountConfig(id);
    const option = document.createElement("option");
    option.value = id;
    if (config && config.serverUrl) {
      const label = config.displayName || config.username || "";
      const host = getHostLabel(config.serverUrl);
      option.textContent = label ? `${label} (${host})` : host;
    } else {
      option.textContent = I18N.get("newAccount") || "New account";
    }
    accountSelectEl.appendChild(option);
  }
  if (currentAccountId) {
    accountSelectEl.value = currentAccountId;
  }
  removeAccountBtn.style.display = ids.length > 1 ? "inline" : "none";
}

accountSelectEl.addEventListener("change", () => {
  currentAccountId = accountSelectEl.value;
  loadAccountUI();
});

addAccountBtn.addEventListener("click", () => {
  const newId = generateAccountId();
  saveAccountConfig(newId, {});
  currentAccountId = newId;
  refreshAccountSelector();
  accountSelectEl.value = newId;
  loadAccountUI();
});

removeAccountBtn.addEventListener("click", () => {
  if (!currentAccountId) return;
  const ids = getAccountIds();
  if (ids.length <= 1) return;
  deleteAccount(currentAccountId);
  currentAccountId = getAccountIds()[0] || null;
  refreshAccountSelector();
  if (currentAccountId) {
    accountSelectEl.value = currentAccountId;
    loadAccountUI();
  }
});

// --- Connection ---
function markConnected(config) {
  loginForm.style.display = "none";
  connectedInfo.style.display = "block";
  connectedServer.textContent = config.serverUrl;
  const name = config.displayName || "";
  const contact = config.contactEmail || "";
  const login = config.username || "";
  if (name && contact && contact !== login) {
    connectedUser.textContent = `${name} (${contact})`;
  } else if (name && name !== login) {
    connectedUser.textContent = `${name} (${login})`;
  } else {
    connectedUser.textContent = login;
  }
  connectedMethod.textContent = config.authMethod === "sso"
    ? (I18N.get("connectedViaSSO") || "Connected via SSO")
    : (I18N.get("connectedViaPassword") || "Connected via password");
}

async function onConnected(newConfig) {
  const existing = getAccountConfig(currentAccountId) || {};
  const config = { ...existing, ...newConfig };

  try {
    const info = await seafile.getAccountInfo(config.serverUrl, config.apiToken);
    config.displayName = info.name || "";
    config.contactEmail = info.contact_email || "";
  } catch (e) {
    console.error("Failed to fetch account info:", e);
  }

  saveAccountConfig(currentAccountId, config);
  await loadRepos(config);
  enableSettingsTabs();
  markConnected(config);
  refreshAccountSelector();

  if (config.repoId && repoSelect.querySelector(`option[value="${config.repoId}"]`)) {
    repoSelect.value = config.repoId;
  } else if (repoSelect.options.length > 1) {
    repoSelect.value = repoSelect.options[1].value;
  }
  if (config.saveRepoId) {
    saveRepoSelect.value = config.saveRepoId;
  }

  if (repoSelect.value) {
    autoSave(repoSelect);
  }

  switchTab("filelink");
  uploadCurrentPath = config.uploadPath || "/";
  saveCurrentPath = config.savePath || "/";
  navigateUploadFolder(uploadCurrentPath);
  navigateSaveFolder(saveCurrentPath);
}

// Re-enable connect button on input change
for (const input of [serverUrlInput, usernameInput, passwordInput, otpInput]) {
  input.addEventListener("input", () => {
    connectBtn.textContent = I18N.get("connect") || "Connect";
    connectBtn.disabled = false;
    connectStatus.className = "status";
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !connectBtn.disabled) connectBtn.click();
  });
}

connectBtn.addEventListener("click", async () => {
  const serverUrl = serverUrlInput.value.trim().replace(/\/+$/, "");
  const username = usernameInput.value.trim();
  const password = passwordInput.value;
  const otp = otpInput.value.trim();

  if (!serverUrl || !username || !password) {
    showStatus(connectStatus, I18N.get("fillAllFields") || "Please fill in all fields.", "error");
    return;
  }

  connectBtn.disabled = true;
  connectBtn.textContent = I18N.get("connecting") || "Connecting...";
  connectStatus.className = "status";

  if (serverUrl.startsWith("http://") && !/^http:\/\/(localhost|127\.0\.0\.1)(:|\/|$)/.test(serverUrl)) {
    showStatus(connectStatus, I18N.get("httpWarning") || "Warning: Connecting over HTTP. Your credentials may be transmitted in plaintext.", "info");
  }

  try {
    const token = await seafile.getToken(serverUrl, username, password, otp);
    await onConnected({ serverUrl, username, apiToken: token, authMethod: "password" });
  } catch (e) {
    showStatus(connectStatus, `Connection failed: ${e.message}`, "error");
    connectBtn.textContent = I18N.get("connect") || "Connect";
    connectBtn.disabled = false;
  }
});

// --- SSO ---
ssoBtn.addEventListener("click", async () => {
  const serverUrl = serverUrlInput.value.trim().replace(/\/+$/, "");
  if (!serverUrl) {
    showStatus(ssoStatus, I18N.get("ssoEnterUrl") || "Please enter the server URL first.", "error");
    return;
  }

  ssoBtn.disabled = true;
  ssoStatus.className = "status";

  if (serverUrl.startsWith("http://") && !/^http:\/\/(localhost|127\.0\.0\.1)(:|\/|$)/.test(serverUrl)) {
    showStatus(ssoStatus, I18N.get("httpWarning") || "Warning: Connecting over HTTP.", "info");
  }

  try {
    const info = await seafile.getServerInfo(serverUrl);
    const features = info.features || [];
    if (!features.includes("client-sso-via-local-browser")) {
      showStatus(ssoStatus, I18N.get("ssoUnavailable") || "SSO via local browser is not enabled on this server.", "info");
      ssoBtn.disabled = false;
      return;
    }

    const result = await seafile.createSSOLink(serverUrl);
    const match = result.link.match(/\/client-sso\/([^/?]+)/);
    if (!match) {
      throw new Error("Failed to parse SSO token from server response.");
    }

    // Open SSO link via Office Dialog API
    Office.context.ui.displayDialogAsync(result.link, { height: 60, width: 40 }, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Fallback: open in new window
        window.open(result.link, "_blank");
      }
    });

    showStatus(ssoStatus, I18N.get("ssoWaiting") || "Waiting for authentication in browser...", "info");

    let elapsed = 0;
    ssoPollingInterval = setInterval(async () => {
      elapsed += 3;
      if (elapsed > 300) {
        clearInterval(ssoPollingInterval);
        ssoPollingInterval = null;
        showStatus(ssoStatus, I18N.get("ssoTimeout") || "SSO login timed out.", "error");
        ssoBtn.disabled = false;
        return;
      }
      try {
        const status = await seafile.checkSSOStatus(serverUrl, match[1]);
        if (status.status === "success" && (status.api_token || status.apiToken || status.api_key)) {
          clearInterval(ssoPollingInterval);
          ssoPollingInterval = null;
          await onConnected({
            serverUrl,
            username: status.username,
            apiToken: status.api_token || status.apiToken || status.api_key,
            authMethod: "sso",
          });
          ssoStatus.className = "status";
        } else if (status.status === "error") {
          clearInterval(ssoPollingInterval);
          ssoPollingInterval = null;
          showStatus(ssoStatus, I18N.get("ssoError") || "SSO login failed.", "error");
          ssoBtn.disabled = false;
        }
      } catch (e) {
        console.error("SSO poll error:", e);
      }
    }, 3000);
  } catch (e) {
    showStatus(ssoStatus, `SSO failed: ${e.message}`, "error");
    ssoBtn.disabled = false;
  }
});

// --- Disconnect ---
disconnectBtn.addEventListener("click", () => {
  const config = getAccountConfig(currentAccountId) || {};
  const serverUrl = config.serverUrl || "";
  saveAccountConfig(currentAccountId, { serverUrl });

  connectedInfo.style.display = "none";
  loginForm.style.display = "block";
  serverUrlInput.value = serverUrl;
  usernameInput.value = "";
  passwordInput.value = "";
  otpInput.value = "";
  connectBtn.textContent = I18N.get("connect") || "Connect";
  connectBtn.disabled = false;
  ssoBtn.disabled = false;
  connectStatus.className = "status";

  disableSettingsTabs();
  switchTab("connection");
  refreshAccountSelector();
});

// --- Repo Loading ---
async function loadRepos(config) {
  const repos = await seafile.listRepos(config.serverUrl, config.apiToken);
  while (repoSelect.options.length > 1) repoSelect.remove(1);
  while (saveRepoSelect.options.length > 1) saveRepoSelect.remove(1);

  const unencrypted = repos.filter(r => !r.encrypted);
  for (const repo of unencrypted) {
    const id = repo.repo_id || repo.id;
    const name = repo.repo_name || repo.name;
    const option = document.createElement("option");
    option.value = id;
    option.textContent = name;
    repoSelect.appendChild(option);
    const saveOption = document.createElement("option");
    saveOption.value = id;
    saveOption.textContent = name;
    saveRepoSelect.appendChild(saveOption);
  }
}

async function refreshRepos() {
  const config = getAccountConfig(currentAccountId);
  if (!config || !config.apiToken) return;
  try {
    const prevRepo = repoSelect.value;
    const prevSaveRepo = saveRepoSelect.value;
    await loadRepos(config);
    if (prevRepo) repoSelect.value = prevRepo;
    if (prevSaveRepo) saveRepoSelect.value = prevSaveRepo;
    if (repoSelect.value) navigateUploadFolder(uploadCurrentPath);
    if (saveRepoSelect.value || repoSelect.value) navigateSaveFolder(saveCurrentPath);
  } catch (e) {
    console.error("Failed to refresh libraries:", e);
  }
}

// --- Folder Pickers ---
async function loadFolderPicker(repoId, path, pathEl, listEl, onNavigate) {
  if (!repoId) return;
  const config = getAccountConfig(currentAccountId);
  if (!config || !config.apiToken) return;

  pathEl.querySelector(".path-text").textContent = path;
  listEl.innerHTML = "";

  try {
    const entries = await seafile.listDir(config.serverUrl, config.apiToken, repoId, path);
    const dirs = entries.filter(e => e.type === "dir");

    if (path !== "/") {
      const parentLi = document.createElement("li");
      const parentPath = path.substring(0, path.lastIndexOf("/")) || "/";
      parentLi.innerHTML = `<span class="folder-icon">${FILE_ICONS.folderUp}</span> ..`;
      parentLi.addEventListener("click", () => onNavigate(parentPath));
      listEl.appendChild(parentLi);
    }
    for (const dir of dirs) {
      const li = document.createElement("li");
      const dirPath = path === "/" ? `/${dir.name}` : `${path}/${dir.name}`;
      li.innerHTML = `<span class="folder-icon">${FILE_ICONS.folder}</span> ${escapeHtml(dir.name)}`;
      li.addEventListener("click", () => onNavigate(dirPath));
      listEl.appendChild(li);
    }
  } catch (e) {
    console.error("Failed to list directory:", e);
  }
}

function navigateUploadFolder(path) {
  uploadCurrentPath = path;
  loadFolderPicker(repoSelect.value, path, uploadPathEl, uploadFolderList, navigateUploadFolder);
}

function navigateSaveFolder(path) {
  saveCurrentPath = path;
  const repoId = saveRepoSelect.value || repoSelect.value;
  loadFolderPicker(repoId, path, savePathEl, saveFolderList, navigateSaveFolder);
}

function toggleFolderPicker(picker) {
  picker.classList.toggle("open");
}

uploadPathEl.addEventListener("click", () => {
  const wasOpen = uploadFolderPicker.classList.contains("open");
  toggleFolderPicker(uploadFolderPicker);
  if (wasOpen) autoSave(uploadFolderPicker);
});

savePathEl.addEventListener("click", () => {
  const wasOpen = saveFolderPicker.classList.contains("open");
  toggleFolderPicker(saveFolderPicker);
  if (wasOpen) autoSave(saveFolderPicker);
});

function closeOpenPickers(e) {
  if (uploadFolderPicker.classList.contains("open") && !uploadFolderPicker.contains(e.target)) {
    uploadFolderPicker.classList.remove("open");
    autoSave(uploadFolderPicker);
  }
  if (saveFolderPicker.classList.contains("open") && !saveFolderPicker.contains(e.target)) {
    saveFolderPicker.classList.remove("open");
    autoSave(saveFolderPicker);
  }
}

document.addEventListener("mousedown", closeOpenPickers);

// --- Auto-Save ---
function flashSaved(el) {
  const formGroup = el.closest(".form-group") || el.parentElement;
  const label = formGroup ? formGroup.querySelector("label") : null;
  if (!label) return;
  let check = label.querySelector(".save-check");
  if (!check) {
    check = document.createElement("span");
    check.className = "save-check";
    check.innerHTML = STATUS_ICONS.success;
    label.appendChild(check);
  }
  check.classList.add("visible");
  clearTimeout(check._timer);
  check._timer = setTimeout(() => check.classList.remove("visible"), 1500);
}

let autoSaveTimer = null;
let autoSaveSource = null;
function autoSave(sourceEl) {
  autoSaveSource = sourceEl || null;
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(() => {
    const config = getAccountConfig(currentAccountId) || {};
    config.repoId = repoSelect.value;
    config.repoName = repoSelect.options[repoSelect.selectedIndex]?.textContent || "";
    config.uploadPath = uploadCurrentPath;
    config.saveRepoId = saveRepoSelect.value || "";
    config.savePath = saveCurrentPath;
    config.sharePasswordMode = sharePasswordModeSelect.value;
    config.sharePasswordLength = Math.max(8, parseInt(sharePasswordLengthInput.value, 10) || 12);
    config.shareCustomPassword = shareCustomPasswordInput.value.trim();
    config.shareExpireDays = Math.max(0, parseInt(shareExpireDaysInput.value, 10) || 0);
    config.showPasswordInEmail = showPasswordInEmailInput.checked;
    config.skipLinkOptions = skipLinkOptionsInput.checked;
    config.saveReplaceExisting = saveReplaceExistingInput.checked;
    config.fileLinkPasswordMode = fileLinkPasswordModeSelect.value;
    config.fileLinkPasswordLength = Math.max(8, parseInt(fileLinkPasswordLengthInput.value, 10) || 12);
    config.fileLinkCustomPassword = fileLinkCustomPasswordInput.value.trim();
    config.fileLinkExpireDays = Math.max(0, parseInt(fileLinkExpireDaysInput.value, 10) || 0);
    saveAccountConfig(currentAccountId, config);
    if (autoSaveSource) flashSaved(autoSaveSource);
  }, 300);
}

// Bind auto-save to all settings inputs
repoSelect.addEventListener("change", () => { uploadCurrentPath = "/"; navigateUploadFolder("/"); autoSave(repoSelect); });
saveRepoSelect.addEventListener("change", () => { saveCurrentPath = "/"; navigateSaveFolder("/"); autoSave(saveRepoSelect); });
shareExpireDaysInput.addEventListener("input", () => { shareExpireDaysInput.value = shareExpireDaysInput.value.replace(/[^0-9]/g, ""); autoSave(shareExpireDaysInput); });
showPasswordInEmailInput.addEventListener("change", () => autoSave(showPasswordInEmailInput.parentElement));
skipLinkOptionsInput.addEventListener("change", () => autoSave(skipLinkOptionsInput.parentElement));
saveReplaceExistingInput.addEventListener("change", () => autoSave(saveReplaceExistingInput.parentElement));

function updateFileLinkPasswordUI() {
  const mode = fileLinkPasswordModeSelect.value;
  document.getElementById("fileLinkPasswordLengthGroup").style.display = mode === "random" ? "block" : "none";
  document.getElementById("fileLinkCustomPasswordGroup").style.display = mode === "custom" ? "block" : "none";
}
fileLinkPasswordModeSelect.addEventListener("change", () => { updateFileLinkPasswordUI(); autoSave(fileLinkPasswordModeSelect); });
fileLinkPasswordLengthInput.addEventListener("input", () => { fileLinkPasswordLengthInput.value = fileLinkPasswordLengthInput.value.replace(/[^0-9]/g, ""); autoSave(fileLinkPasswordLengthInput); });
fileLinkCustomPasswordInput.addEventListener("input", () => autoSave(fileLinkCustomPasswordInput));
fileLinkExpireDaysInput.addEventListener("input", () => { fileLinkExpireDaysInput.value = fileLinkExpireDaysInput.value.replace(/[^0-9]/g, ""); autoSave(fileLinkExpireDaysInput); });

function updateSharePasswordUI() {
  const mode = sharePasswordModeSelect.value;
  document.getElementById("sharePasswordLengthGroup").style.display = mode === "random" ? "block" : "none";
  document.getElementById("shareCustomPasswordGroup").style.display = mode === "custom" ? "block" : "none";
}
sharePasswordModeSelect.addEventListener("change", () => { updateSharePasswordUI(); autoSave(sharePasswordModeSelect); });
sharePasswordLengthInput.addEventListener("input", () => { sharePasswordLengthInput.value = sharePasswordLengthInput.value.replace(/[^0-9]/g, ""); autoSave(sharePasswordLengthInput); });
shareCustomPasswordInput.addEventListener("input", () => autoSave(shareCustomPasswordInput));

// --- Load Account UI ---
async function loadAccountUI() {
  const config = getAccountConfig(currentAccountId);
  connectStatus.className = "status";
  ssoStatus.className = "status";

  if (!config || !config.apiToken) {
    // Not connected - show login form
    connectedInfo.style.display = "none";
    loginForm.style.display = "block";
    serverUrlInput.value = (config && config.serverUrl) || "";
    usernameInput.value = "";
    passwordInput.value = "";
    otpInput.value = "";
    connectBtn.textContent = I18N.get("connect") || "Connect";
    connectBtn.disabled = false;
    ssoBtn.disabled = false;
    disableSettingsTabs();
    switchTab("connection");
    return;
  }

  // Connected - load settings
  serverUrlInput.value = config.serverUrl || "";
  usernameInput.value = config.username || "";
  sharePasswordModeSelect.value = config.sharePasswordMode || "none";
  sharePasswordLengthInput.value = config.sharePasswordLength || 12;
  shareCustomPasswordInput.value = config.shareCustomPassword || "";
  shareExpireDaysInput.value = config.shareExpireDays || 0;
  showPasswordInEmailInput.checked = config.showPasswordInEmail !== false;
  skipLinkOptionsInput.checked = !!config.skipLinkOptions;
  fileLinkPasswordModeSelect.value = config.fileLinkPasswordMode || "none";
  fileLinkPasswordLengthInput.value = config.fileLinkPasswordLength || 12;
  fileLinkCustomPasswordInput.value = config.fileLinkCustomPassword || "";
  fileLinkExpireDaysInput.value = config.fileLinkExpireDays || 0;
  updateFileLinkPasswordUI();
  updateSharePasswordUI();
  saveReplaceExistingInput.checked = !!config.saveReplaceExisting;

  try {
    if (!config.displayName) {
      try {
        const info = await seafile.getAccountInfo(config.serverUrl, config.apiToken);
        config.displayName = info.name || "";
        config.contactEmail = info.contact_email || "";
        saveAccountConfig(currentAccountId, config);
        refreshAccountSelector();
      } catch (e) {
        console.error("Failed to fetch account info:", e);
      }
    }

    await loadRepos(config);
    enableSettingsTabs();
    markConnected(config);

    if (config.repoId && repoSelect.querySelector(`option[value="${config.repoId}"]`)) {
      repoSelect.value = config.repoId;
    } else if (repoSelect.options.length > 1) {
      repoSelect.value = repoSelect.options[1].value;
      autoSave(repoSelect);
    }
    if (config.saveRepoId) {
      saveRepoSelect.value = config.saveRepoId;
    }

    uploadCurrentPath = config.uploadPath || "/";
    saveCurrentPath = config.savePath || "/";
    navigateUploadFolder(uploadCurrentPath);
    navigateSaveFolder(saveCurrentPath);
  } catch (e) {
    showStatus(connectStatus, "Session expired. Please reconnect.", "error");
    connectedInfo.style.display = "none";
    loginForm.style.display = "block";
    disableSettingsTabs();
  }
}

// --- Initialize ---
Office.onReady(async () => {
  await I18N.init();
  I18N.applyToDocument();

  // Ensure at least one account exists
  const ids = getAccountIds();
  if (ids.length === 0) {
    const newId = generateAccountId();
    saveAccountConfig(newId, {});
    currentAccountId = newId;
  } else {
    currentAccountId = ids[0];
  }

  refreshAccountSelector();
  accountSelectEl.value = currentAccountId;
  loadAccountUI();
});
