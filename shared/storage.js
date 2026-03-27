/**
 * Account storage manager for the Seafile Outlook Add-in.
 *
 * Uses localStorage for account data. Each account is stored under
 * key "seafile_account_{id}". The account list is stored under "seafile_accounts".
 *
 * This replaces Thunderbird's browser.storage.local API with a compatible
 * pattern for Office.js add-ins.
 */

const STORAGE_PREFIX = "seafile_";
const ACCOUNTS_KEY = STORAGE_PREFIX + "accounts";

/**
 * Generate a simple unique ID for new accounts.
 * @returns {string}
 */
function generateAccountId() {
  return "acc_" + Date.now().toString(36) + "_" + Math.random().toString(36).substring(2, 8);
}

/**
 * Get the list of all account IDs.
 * @returns {string[]}
 */
function getAccountIds() {
  try {
    return JSON.parse(localStorage.getItem(ACCOUNTS_KEY)) || [];
  } catch {
    return [];
  }
}

/**
 * Save the list of account IDs.
 * @param {string[]} ids
 */
function saveAccountIds(ids) {
  localStorage.setItem(ACCOUNTS_KEY, JSON.stringify(ids));
}

/**
 * Get account configuration by ID.
 * @param {string} accountId
 * @returns {Object|null}
 */
function getAccountConfig(accountId) {
  try {
    const data = localStorage.getItem(STORAGE_PREFIX + accountId);
    return data ? JSON.parse(data) : null;
  } catch {
    return null;
  }
}

/**
 * Save account configuration.
 * @param {string} accountId
 * @param {Object} config
 */
function saveAccountConfig(accountId, config) {
  localStorage.setItem(STORAGE_PREFIX + accountId, JSON.stringify(config));

  // Ensure account is in the accounts list
  const ids = getAccountIds();
  if (!ids.includes(accountId)) {
    ids.push(accountId);
    saveAccountIds(ids);
  }
}

/**
 * Delete an account and its data.
 * @param {string} accountId
 */
function deleteAccount(accountId) {
  localStorage.removeItem(STORAGE_PREFIX + accountId);
  localStorage.removeItem(STORAGE_PREFIX + accountId + "_files");
  const ids = getAccountIds().filter(id => id !== accountId);
  saveAccountIds(ids);
}

/**
 * Get all configured accounts (with apiToken and repoId).
 * @returns {Array<Object>} Array of { accountId, serverUrl, username, displayName }
 */
function getAllConfiguredAccounts() {
  const ids = getAccountIds();
  const result = [];
  for (const id of ids) {
    const config = getAccountConfig(id);
    if (config && config.apiToken && config.repoId) {
      result.push({
        accountId: id,
        serverUrl: config.serverUrl,
        username: config.username,
        displayName: config.displayName || "",
      });
    }
  }
  return result;
}

/**
 * Get the first configured account (fallback).
 * @returns {Object|null} { accountId, ...config } or null
 */
function getFirstConfiguredAccount() {
  const accounts = getAllConfiguredAccounts();
  if (accounts.length === 0) return null;
  const config = getAccountConfig(accounts[0].accountId);
  return config ? { accountId: accounts[0].accountId, ...config } : null;
}

/**
 * Resolve an account: use explicit ID if provided, otherwise first configured.
 * @param {string} [accountId]
 * @returns {Object} { accountId, ...config }
 * @throws {Error} If no account is configured
 */
function resolveAccount(accountId) {
  if (accountId) {
    const config = getAccountConfig(accountId);
    if (config && config.apiToken) {
      return { accountId, ...config };
    }
  }
  const first = getFirstConfiguredAccount();
  if (!first) {
    throw new Error("No Seafile account configured.");
  }
  return first;
}

/**
 * Save file metadata for tracking uploaded files (for deletion support).
 * @param {string} accountId
 * @param {string} fileKey - Unique key for the file
 * @param {Object} metadata - { path, shareLinkToken }
 */
function saveFileMetadata(accountId, fileKey, metadata) {
  const key = STORAGE_PREFIX + accountId + "_files";
  let files = {};
  try {
    files = JSON.parse(localStorage.getItem(key)) || {};
  } catch { /* ignore */ }
  files[fileKey] = metadata;
  localStorage.setItem(key, JSON.stringify(files));
}

/**
 * Get and remove file metadata.
 * @param {string} accountId
 * @param {string} fileKey
 * @returns {Object|null}
 */
function popFileMetadata(accountId, fileKey) {
  const key = STORAGE_PREFIX + accountId + "_files";
  let files = {};
  try {
    files = JSON.parse(localStorage.getItem(key)) || {};
  } catch { /* ignore */ }
  const metadata = files[fileKey] || null;
  if (metadata) {
    delete files[fileKey];
    localStorage.setItem(key, JSON.stringify(files));
  }
  return metadata;
}

/**
 * Get/set the last used account ID for a feature.
 * @param {string} feature - "compose" or "read"
 * @param {string} [accountId] - If provided, sets the value
 * @returns {string|null} The last used account ID (when getting)
 */
function lastUsedAccount(feature, accountId) {
  const key = STORAGE_PREFIX + "lastAccount_" + feature;
  if (accountId !== undefined) {
    localStorage.setItem(key, accountId);
    return accountId;
  }
  return localStorage.getItem(key);
}
