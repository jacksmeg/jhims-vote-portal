function normalizeStaffId(value) {
  return String(value || "").trim().toUpperCase();
}

function normalizePhoneNumber(value) {
  const digitsOnly = String(value || "").replace(/\D/g, "");

  if (!digitsOnly) {
    return "";
  }

  if (digitsOnly.startsWith("233") && digitsOnly.length === 12) {
    return `0${digitsOnly.slice(3)}`;
  }

  if (digitsOnly.length === 9) {
    return `0${digitsOnly}`;
  }

  return digitsOnly;
}

function isLikelyPhoneNumber(value) {
  const normalized = normalizePhoneNumber(value);
  return normalized.length >= 10 && normalized.length <= 15;
}

function toSmsPhoneNumber(value) {
  const normalized = normalizePhoneNumber(value);
  const digitsOnly = normalized.replace(/\D/g, "");

  if (!digitsOnly) {
    return "";
  }

  if (digitsOnly.startsWith("0") && digitsOnly.length === 10) {
    return `+233${digitsOnly.slice(1)}`;
  }

  if (digitsOnly.startsWith("233") && digitsOnly.length === 12) {
    return `+${digitsOnly}`;
  }

  if (digitsOnly.length === 9) {
    return `+233${digitsOnly}`;
  }

  if (digitsOnly.length >= 10 && digitsOnly.length <= 15) {
    return `+${digitsOnly}`;
  }

  return "";
}

module.exports = {
  isLikelyPhoneNumber,
  normalizePhoneNumber,
  normalizeStaffId,
  toSmsPhoneNumber,
};
