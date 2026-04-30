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

module.exports = {
  isLikelyPhoneNumber,
  normalizePhoneNumber,
  normalizeStaffId,
};
