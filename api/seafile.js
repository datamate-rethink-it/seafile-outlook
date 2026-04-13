/**
 * Seafile API client for Outlook Web Add-in.
 * Ported from the Seafile Thunderbird extension.
 *
 * In dev mode (localhost), API calls are routed through a local proxy
 * to avoid CORS issues. In production, calls go directly to the Seafile server
 * (which must have CORS configured, or be on the same domain).
 */

/**
 * Check if we're running on the dev server (localhost).
 * If so, prefix API URLs with the proxy path.
 */
function proxyUrl(url) {
  if (location.hostname === "localhost" || location.hostname === "127.0.0.1") {
    return `${location.origin}/seafile-proxy/${url}`;
  }
  return url;
}

class SeafileAPI {

  /**
   * Authenticate with username/password and get an API token.
   */
  async getToken(server, username, password, otp) {
    const headers = { "Content-Type": "application/x-www-form-urlencoded" };
    if (otp) {
      headers["X-SEAFILE-OTP"] = otp;
    }
    const resp = await fetch(proxyUrl(`${server}/api2/auth-token/`), {
      method: "POST",
      headers,
      body: new URLSearchParams({ username, password }),
    });
    if (!resp.ok) {
      const text = await resp.text();
      console.error(`Auth failed (${resp.status}):`, text);
      throw new Error(`Authentication failed (${resp.status})`);
    }
    const data = await resp.json();
    return data.token;
  }

  /**
   * Get server info (features, version, etc.).
   */
  async getServerInfo(server) {
    const resp = await fetch(proxyUrl(`${server}/api2/server-info/`));
    if (!resp.ok) {
      throw new Error(`Failed to get server info (${resp.status})`);
    }
    return await resp.json();
  }

  /**
   * Request a client SSO login link.
   */
  async createSSOLink(server) {
    const resp = await fetch(proxyUrl(`${server}/api2/client-sso-link/`), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        shib_platform: "outlook-addin",
        shib_device_name: "Outlook",
      }),
    });
    if (!resp.ok) {
      throw new Error(`Failed to create SSO link (${resp.status})`);
    }
    return await resp.json();
  }

  /**
   * Poll SSO login status.
   */
  async checkSSOStatus(server, ssoToken) {
    const resp = await fetch(proxyUrl(`${server}/api2/client-sso-link/${ssoToken}/`));
    if (!resp.ok) {
      throw new Error(`Failed to check SSO status (${resp.status})`);
    }
    return await resp.json();
  }

  /**
   * Get account info (display name, contact email, usage, etc.).
   */
  async getAccountInfo(server, token) {
    const resp = await fetch(proxyUrl(`${server}/api2/account/info/`), {
      headers: { Authorization: `Token ${token}` },
    });
    if (!resp.ok) {
      throw new Error(`Failed to get account info (${resp.status})`);
    }
    return await resp.json();
  }

  /**
   * List all accessible libraries/repos.
   */
  async listRepos(server, token) {
    const resp = await fetch(proxyUrl(`${server}/api/v2.1/repos/`), {
      headers: { Authorization: `Token ${token}` },
    });
    if (!resp.ok) {
      throw new Error(`Failed to list libraries (${resp.status})`);
    }
    const data = await resp.json();
    return data.repos || data;
  }

  /**
   * Get an upload link for a given repo and target directory.
   */
  async getUploadLink(server, token, repoId, parentDir = "/") {
    const params = new URLSearchParams({ p: parentDir });
    const resp = await fetch(proxyUrl(`${server}/api2/repos/${repoId}/upload-link/?${params}`), {
      headers: { Authorization: `Token ${token}` },
    });
    if (!resp.ok) {
      throw new Error(`Failed to get upload link (${resp.status})`);
    }
    const link = await resp.json();
    return typeof link === "string" ? link : link.upload_link || link;
  }

  /**
   * Upload a file to Seafile.
   */
  async uploadFile(uploadLink, token, file, fileName, parentDir, signal, replace = true) {
    const formData = new FormData();
    formData.append("file", file, fileName);
    formData.append("parent_dir", parentDir);
    formData.append("replace", replace ? "1" : "0");

    const url = uploadLink.endsWith("?ret-json=1")
      ? uploadLink
      : `${uploadLink}?ret-json=1`;

    const resp = await fetch(proxyUrl(url), {
      method: "POST",
      headers: { Authorization: `Token ${token}` },
      body: formData,
      signal,
    });
    if (!resp.ok) {
      const text = await resp.text();
      console.error(`Upload failed (${resp.status}):`, text);
      throw new Error(`File upload failed (${resp.status})`);
    }
    const data = await resp.json();
    return Array.isArray(data) ? data[0] : data;
  }

  /**
   * Get existing share links for a file.
   */
  async getShareLinks(server, token, repoId, path) {
    const params = new URLSearchParams({ repo_id: repoId, path });
    const resp = await fetch(proxyUrl(`${server}/api/v2.1/share-links/?${params}`), {
      headers: { Authorization: `Token ${token}` },
    });
    if (!resp.ok) {
      return [];
    }
    return await resp.json();
  }

  /**
   * Create a share/download link for a file.
   */
  async createShareLink(server, token, repoId, path, options = {}) {
    const body = { repo_id: repoId, path };
    if (options.password) {
      body.password = options.password;
    }
    if (options.expireDays) {
      body.expire_days = options.expireDays;
    }

    const resp = await fetch(proxyUrl(`${server}/api/v2.1/share-links/`), {
      method: "POST",
      headers: {
        Authorization: `Token ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      const text = await resp.text();
      console.error(`Share link failed (${resp.status}):`, text);
      let detail = "";
      try {
        const json = JSON.parse(text);
        detail = json.error_msg || json.detail || json.password || "";
        if (Array.isArray(detail)) detail = detail.join(". ");
      } catch { /* not JSON */ }
      throw new Error(detail || `Failed to create share link (${resp.status})`);
    }
    return await resp.json();
  }

  /**
   * Delete a share link.
   */
  async deleteShareLink(server, token, linkToken) {
    const resp = await fetch(proxyUrl(`${server}/api/v2.1/share-links/${linkToken}/`), {
      method: "DELETE",
      headers: { Authorization: `Token ${token}` },
    });
    if (!resp.ok && resp.status !== 404) {
      throw new Error(`Failed to delete share link (${resp.status})`);
    }
  }

  /**
   * Delete a file from a repo.
   */
  async deleteFile(server, token, repoId, path) {
    const resp = await fetch(
      proxyUrl(`${server}/api2/repos/${repoId}/file/?p=${encodeURIComponent(path)}`),
      {
        method: "DELETE",
        headers: { Authorization: `Token ${token}` },
      }
    );
    if (!resp.ok && resp.status !== 404) {
      throw new Error(`Failed to delete file (${resp.status})`);
    }
  }

  /**
   * Create a directory in a repo.
   */
  async createDir(server, token, repoId, path) {
    const resp = await fetch(
      proxyUrl(`${server}/api2/repos/${repoId}/dir/?p=${encodeURIComponent(path)}`),
      {
        method: "POST",
        headers: {
          Authorization: `Token ${token}`,
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({ operation: "mkdir" }),
      }
    );
    if (!resp.ok && resp.status !== 409) {
      throw new Error(`Failed to create directory (${resp.status})`);
    }
  }

  /**
   * List directory contents.
   */
  async listDir(server, token, repoId, path = "/") {
    const resp = await fetch(
      proxyUrl(`${server}/api2/repos/${repoId}/dir/?p=${encodeURIComponent(path)}`),
      {
        headers: { Authorization: `Token ${token}` },
      }
    );
    if (!resp.ok) {
      throw new Error(`Failed to list directory (${resp.status})`);
    }
    return await resp.json();
  }

  /**
   * Revoke an API token server-side (best-effort).
   */
  async logout(server, token) {
    const resp = await fetch(proxyUrl(`${server}/api2/logout-device/`), {
      method: "POST",
      headers: { Authorization: `Token ${token}` },
    });
    if (!resp.ok && resp.status !== 401) {
      throw new Error(`Failed to logout (${resp.status})`);
    }
  }

  /**
   * Check if a directory exists in a repo.
   */
  async dirExists(server, token, repoId, path) {
    const resp = await fetch(
      proxyUrl(`${server}/api2/repos/${repoId}/dir/?p=${encodeURIComponent(path)}`),
      {
        headers: { Authorization: `Token ${token}` },
      }
    );
    return resp.ok;
  }
}

/**
 * Extract share link token from a Seafile share URL.
 */
function extractTokenFromUrl(url) {
  if (!url) return null;
  const match = url.match(/\/[fd]\/([a-zA-Z0-9]+)\/?/);
  return match ? match[1] : null;
}
