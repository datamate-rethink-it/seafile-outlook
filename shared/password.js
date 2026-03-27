/**
 * Cryptographically secure password generator.
 * Ported from the Seafile Thunderbird extension.
 */

/**
 * Generate a cryptographically secure random integer in [0, max).
 */
function secureRandomInt(max) {
  const array = new Uint32Array(1);
  crypto.getRandomValues(array);
  return array[0] % max;
}

/**
 * Generate a random password (mixed case + digits + special chars).
 * First and last characters are always alphanumeric (for double-click selection).
 * @param {number} [length=12] - Password length
 * @returns {string}
 */
function generateRandomPassword(length = 12) {
  const lower = "abcdefghijkmnpqrstuvwxyz";
  const upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";
  const digits = "23456789";
  const special = "!@#$%&*?";
  const all = lower + upper + digits + special;

  // Ensure at least one of each type
  const required = [
    lower[secureRandomInt(lower.length)],
    upper[secureRandomInt(upper.length)],
    digits[secureRandomInt(digits.length)],
    special[secureRandomInt(special.length)],
  ];
  const rest = [];
  for (let i = required.length; i < length; i++) {
    rest.push(all[secureRandomInt(all.length)]);
  }

  // Combine and shuffle (Fisher-Yates)
  const result = [...required, ...rest];
  for (let i = result.length - 1; i > 0; i--) {
    const j = secureRandomInt(i + 1);
    [result[i], result[j]] = [result[j], result[i]];
  }

  // Alphanumeric at start and end for double-click selection
  const alnum = lower + upper + digits;
  result[0] = alnum[secureRandomInt(alnum.length)];
  result[result.length - 1] = alnum[secureRandomInt(alnum.length)];

  // Ensure at least one special char remains in the middle
  const hasSpecial = result.some(c => special.includes(c));
  if (!hasSpecial) {
    const pos = 1 + secureRandomInt(result.length - 2);
    result[pos] = special[secureRandomInt(special.length)];
  }

  return result.join("");
}

/**
 * Resolve the effective password based on a password mode config.
 * @param {string} mode - "none", "random", or "custom"
 * @param {number} length - Password length (for random mode)
 * @param {string} customPassword - Custom password (for custom mode)
 * @returns {string} Password or empty string
 */
function resolvePassword(mode, length, customPassword) {
  if (mode === "random") return generateRandomPassword(length || 12);
  if (mode === "custom") return customPassword || "";
  return "";
}
