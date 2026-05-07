require("dotenv").config();

const bcrypt = require("bcryptjs");
const crypto = require("node:crypto");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const path = require("node:path");
const dayjs = require("dayjs");
const express = require("express");
const session = require("express-session");
const helmet = require("helmet");
const multer = require("multer");
const PDFDocument = require("pdfkit");

const {
  databasePath,
  db,
  getAllSettings,
  initDatabase,
  nowIso,
  runTransaction,
  setSetting,
} = require("./db");
const { ensureVoterTemplate, parseVoterWorkbook } = require("./helpers/excel");
const {
  isLikelyPhoneNumber,
  normalizePhoneNumber,
  normalizeStaffId,
  toSmsPhoneNumber,
} = require("./helpers/phone");
const { requireAdmin, requireVoter } = require("./middleware/auth");

const app = express();
app.set("trust proxy", 1);

const host = process.env.HOST || "0.0.0.0";
const port = Number.parseInt(process.env.PORT || "3000", 10);
const sessionSecret =
  process.env.SESSION_SECRET || "development-session-secret-change-me";
const adminUsername = String(process.env.ADMIN_USERNAME || "admin").trim();
const adminPasswordHash = bcrypt.hashSync(
  String(process.env.ADMIN_PASSWORD || "ChangeMe123!"),
  10,
);
const defaultElectionName =
  process.env.ELECTION_NAME || "Organization Election Portal";
const isProduction = process.env.NODE_ENV === "production";
const twilioAccountSid = String(process.env.TWILIO_ACCOUNT_SID || "").trim();
const twilioAuthToken = String(process.env.TWILIO_AUTH_TOKEN || "").trim();
const twilioVerifyServiceSid = String(process.env.TWILIO_VERIFY_SERVICE_SID || "").trim();
const twilioOtpConfigured = Boolean(
  twilioAccountSid && twilioAuthToken && twilioVerifyServiceSid,
);
const configuredOtpProvider = String(process.env.OTP_PROVIDER || "")
  .trim()
  .toLowerCase();
const otpProvider =
  configuredOtpProvider === "twilio" ||
  configuredOtpProvider === "dev" ||
  configuredOtpProvider === "disabled"
    ? configuredOtpProvider
    : twilioOtpConfigured
      ? "twilio"
      : isProduction
        ? "disabled"
        : "dev";
const otpTtlMinutes = Math.min(
  Math.max(Number.parseInt(process.env.OTP_TTL_MINUTES || "10", 10) || 10, 1),
  30,
);
const otpResendCooldownSeconds = Math.min(
  Math.max(
    Number.parseInt(process.env.OTP_RESEND_COOLDOWN_SECONDS || "30", 10) || 30,
    0,
  ),
  300,
);
const devOtpCodeLength = 6;
const sessionSecureCookie = String(
  process.env.SESSION_SECURE_COOKIE || (isProduction ? "true" : "false"),
)
  .trim()
  .toLowerCase() === "true";

const publicDirectory = path.join(process.cwd(), "public");
const templatesDirectory = path.join(publicDirectory, "templates");
const legacyUploadsDirectory = path.join(publicDirectory, "uploads");
const storageRoot = process.env.STORAGE_ROOT
  ? path.isAbsolute(process.env.STORAGE_ROOT)
    ? process.env.STORAGE_ROOT
    : path.join(process.cwd(), process.env.STORAGE_ROOT)
  : path.join(process.cwd(), "data", "storage");
const uploadsRootDirectory = path.join(storageRoot, "uploads");
const candidateUploadsDirectory = path.join(uploadsRootDirectory, "candidates");
const brandingUploadsDirectory = path.join(uploadsRootDirectory, "branding");
const nominationUploadsDirectory = path.join(uploadsRootDirectory, "nominations");
const importUploadsDirectory = path.join(storageRoot, "imports");
const backupsDirectory = path.join(storageRoot, "backups");
const templatePath = path.join(
  templatesDirectory,
  "voter-import-template.xlsx",
);
const staffLoginTemplatePath = path.join(
  templatesDirectory,
  "staff-login-template.xlsx",
);
const themeOptions = [
  {
    value: "heritage",
    label: "Heritage Gold",
    description: "Warm gold and navy tones for the current election style.",
  },
  {
    value: "emerald",
    label: "Emerald Pulse",
    description: "Fresh green branding with a lighter civic dashboard look.",
  },
  {
    value: "midnight",
    label: "Midnight Blue",
    description: "Deep blue contrast with brighter highlights for readability.",
  },
];

function ensureDirectories() {
  [
    publicDirectory,
    templatesDirectory,
    storageRoot,
    uploadsRootDirectory,
    candidateUploadsDirectory,
    brandingUploadsDirectory,
    nominationUploadsDirectory,
    importUploadsDirectory,
    backupsDirectory,
  ].forEach((directoryPath) => {
    fs.mkdirSync(directoryPath, { recursive: true });
  });
}

function setFlash(req, type, message) {
  req.session.flash = { type, message };
}

function parseInteger(value, fallback = 0) {
  const parsed = Number.parseInt(String(value || ""), 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function formatDateTime(value) {
  if (!value) {
    return "Not set";
  }

  const parsed = dayjs(value);
  return parsed.isValid() ? parsed.format("DD MMM YYYY, h:mm A") : "Invalid date";
}

function formatPercent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function toSafeFilename(value) {
  return (
    String(value || "election-results")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "") || "election-results"
  );
}

function getInitials(name) {
  return String(name || "")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() || "")
    .join("");
}

function safeJsonParse(value, fallback) {
  try {
    return value ? JSON.parse(value) : fallback;
  } catch (_error) {
    return fallback;
  }
}

function getThemeOptions() {
  return themeOptions;
}

function isWithinDirectory(targetPath, baseDirectory) {
  const relativePath = path.relative(baseDirectory, targetPath);
  return (
    relativePath === "" ||
    (!relativePath.startsWith("..") && !path.isAbsolute(relativePath))
  );
}

function resolvePathWithin(baseDirectory, relativePath) {
  const basePath = path.resolve(baseDirectory);
  const absolutePath = path.resolve(baseDirectory, relativePath);
  return isWithinDirectory(absolutePath, basePath) ? absolutePath : "";
}

function normalizeAssetPath(filePath) {
  const absolutePath = path.resolve(filePath);
  const uploadsPath = path.resolve(uploadsRootDirectory);
  const publicPath = path.resolve(publicDirectory);

  if (isWithinDirectory(absolutePath, uploadsPath)) {
    const relativePath = path.relative(uploadsPath, absolutePath).replace(/\\/g, "/");
    return `/uploads/${relativePath}`;
  }

  if (isWithinDirectory(absolutePath, publicPath)) {
    const relativePath = path.relative(publicPath, absolutePath).replace(/\\/g, "/");
    return `/${relativePath}`;
  }

  throw new Error("Asset path is outside the directories served by the app.");
}

function resolveAssetPath(assetPath) {
  if (!assetPath) {
    return "";
  }

  const normalizedPath = String(assetPath).trim().replace(/\\/g, "/");
  const relativePath = normalizedPath.replace(/^[/\\]+/, "");

  if (
    normalizedPath.startsWith("/uploads/") ||
    relativePath === "uploads" ||
    relativePath.startsWith("uploads/")
  ) {
    const uploadRelativePath =
      relativePath === "uploads" ? "" : relativePath.replace(/^uploads\//, "");
    const uploadPath = resolvePathWithin(uploadsRootDirectory, uploadRelativePath);
    if (uploadPath && fs.existsSync(uploadPath)) {
      return uploadPath;
    }

    const legacyUploadPath = resolvePathWithin(legacyUploadsDirectory, uploadRelativePath);
    return legacyUploadPath || uploadPath;
  }

  return resolvePathWithin(publicDirectory, relativePath);
}

function getElectionSettings() {
  const settings = getAllSettings();
  const electionName = settings.election_name || defaultElectionName;
  return {
    electionName,
    organizationLogoPath: settings.organization_logo_path || "",
    phase: settings.election_phase || "setup",
    opensAt: settings.opens_at || "",
    closesAt: settings.closes_at || "",
    resultsVisibility: settings.results_visibility || "after_close",
    themeName: settings.theme_name || "heritage",
    declarationTitle: settings.declaration_title || "Official Declaration Block",
    committeeName: settings.committee_name || `${electionName} Committee`,
    chairmanName: settings.chairman_name || "",
    secretaryName: settings.secretary_name || "",
    chairmanSignaturePath: settings.chairman_signature_path || "",
    secretarySignaturePath: settings.secretary_signature_path || "",
    nominationPhase: settings.nomination_phase || "setup",
    nominationOpensAt: settings.nomination_opens_at || "",
    nominationClosesAt: settings.nomination_closes_at || "",
  };
}

function logSystemAudit(action, details = {}) {
  db.prepare(`
    INSERT INTO audit_logs (
      actor_type,
      actor_identifier,
      action,
      details_json,
      ip_address,
      user_agent,
      created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(
    "system",
    "scheduler",
    action,
    JSON.stringify(details),
    "",
    "",
    nowIso(),
  );
}

function getPositionsWithoutCandidates() {
  return db.prepare(`
    SELECT p.name
    FROM positions p
    LEFT JOIN candidates c
      ON c.position_id = p.id
      AND c.is_active = 1
    WHERE p.is_active = 1
    GROUP BY p.id
    HAVING COUNT(c.id) = 0
  `).all();
}

function getElectionReadiness(settings = getElectionSettings()) {
  const metrics = getDashboardMetrics();
  const positionsWithoutCandidates = getPositionsWithoutCandidates();
  const issues = [];

  if (metrics.totalVoters === 0) {
    issues.push("Import at least one voter before voting can open.");
  }

  if (metrics.totalPositions === 0 || metrics.totalCandidates === 0) {
    issues.push("Add positions and candidates before voting can open.");
  }

  if (!settings.opensAt || !settings.closesAt) {
    issues.push("Set both the voting start time and closing time.");
  } else if (!dayjs(settings.opensAt).isBefore(dayjs(settings.closesAt))) {
    issues.push("The closing time must be later than the opening time.");
  }

  if (positionsWithoutCandidates.length > 0) {
    issues.push(
      `Every position needs at least one candidate. Missing: ${positionsWithoutCandidates
        .map((position) => position.name)
        .join(", ")}.`,
    );
  }

  return {
    isReady: issues.length === 0,
    issues,
    metrics,
    positionsWithoutCandidates,
  };
}

function getNominationReadiness(settings = getElectionSettings()) {
  const issues = [];
  const positions = getPositions();

  if (positions.length === 0) {
    issues.push("Add at least one position before opening nominations.");
  }

  if (!settings.nominationOpensAt || !settings.nominationClosesAt) {
    issues.push("Set both the nomination opening and closing time.");
  } else if (!dayjs(settings.nominationOpensAt).isBefore(dayjs(settings.nominationClosesAt))) {
    issues.push("The nomination closing time must be later than the opening time.");
  }

  return {
    isReady: issues.length === 0,
    issues,
  };
}

function computeElectionState(settings) {
  const now = dayjs();
  const opensAt = settings.opensAt ? dayjs(settings.opensAt) : null;
  const closesAt = settings.closesAt ? dayjs(settings.closesAt) : null;
  const readiness = getElectionReadiness(settings);

  let status = "setup";
  let message =
    "The election is still in setup. Add voters, positions, and candidates before opening voting.";

  if (settings.phase === "setup") {
    if (readiness.isReady && opensAt?.isValid() && closesAt?.isValid()) {
      if (now.isBefore(opensAt)) {
        status = "scheduled";
        message = `Voting is scheduled to open automatically on ${formatDateTime(
          settings.opensAt,
        )}.`;
      } else if (!now.isBefore(closesAt)) {
        message = `The scheduled voting window ended on ${formatDateTime(
          settings.closesAt,
        )}. Update the election dates to schedule a new vote.`;
      }
    }
  } else if (settings.phase === "open") {
    if (opensAt && now.isBefore(opensAt)) {
      status = "scheduled";
      message = `Voting has been opened by the committee and will start on ${formatDateTime(
        settings.opensAt,
      )}.`;
    } else if (closesAt && !now.isBefore(closesAt)) {
      status = "closed";
      message = `Voting closed on ${formatDateTime(settings.closesAt)}.`;
    } else {
      status = "open";
      message = "Voting is currently open. Each registered staff member can vote once.";
    }
  } else if (settings.phase === "closed") {
    status = "closed";
    message = "Voting has been closed. Ballots are locked and results are final.";
  }

  return {
    status,
    message,
    isOpen: status === "open",
    isScheduled: status === "scheduled",
    isClosed: status === "closed",
    isSetup: settings.phase === "setup",
    canEditSetup: settings.phase === "setup",
    badgeLabel:
      status === "open"
        ? "Voting Open"
        : status === "scheduled"
          ? "Scheduled"
          : status === "closed"
            ? "Closed"
            : "Setup",
    badgeClass:
      status === "open"
        ? "status-open"
        : status === "scheduled"
          ? "status-scheduled"
          : status === "closed"
            ? "status-closed"
            : "status-setup",
  };
}

function computeNominationState(settings) {
  const now = dayjs();
  const opensAt = settings.nominationOpensAt ? dayjs(settings.nominationOpensAt) : null;
  const closesAt = settings.nominationClosesAt ? dayjs(settings.nominationClosesAt) : null;
  const readiness = getNominationReadiness(settings);

  let status = "setup";
  let message =
    "Nominations are in setup. Set the nomination window before candidates can apply.";

  if (settings.nominationPhase === "setup") {
    if (readiness.isReady && opensAt?.isValid() && closesAt?.isValid()) {
      if (now.isBefore(opensAt)) {
        status = "scheduled";
        message = `Nominations are scheduled to open on ${formatDateTime(settings.nominationOpensAt)}.`;
      } else if (!now.isBefore(closesAt)) {
        message = `The nomination window ended on ${formatDateTime(settings.nominationClosesAt)}. Update the dates to reopen nominations.`;
      }
    }
  } else if (settings.nominationPhase === "open") {
    if (opensAt?.isValid() && now.isBefore(opensAt)) {
      status = "scheduled";
      message = `Nominations have been opened by the committee and will start on ${formatDateTime(
        settings.nominationOpensAt,
      )}.`;
    } else if (closesAt?.isValid() && !now.isBefore(closesAt)) {
      status = "closed";
      message = `Nominations closed on ${formatDateTime(settings.nominationClosesAt)}.`;
    } else {
      status = "open";
      message = "Nominations are currently open for registered staff members.";
    }
  } else if (settings.nominationPhase === "closed") {
    status = "closed";
    message = "Nominations have been closed for this election cycle.";
  }

  return {
    status,
    message,
    isOpen: status === "open",
    isScheduled: status === "scheduled",
    isClosed: status === "closed",
    canSubmit: status === "open",
    badgeLabel:
      status === "open"
        ? "Nominations Open"
        : status === "scheduled"
          ? "Scheduled"
          : status === "closed"
            ? "Closed"
            : "Setup",
    badgeClass:
      status === "open"
        ? "status-open"
        : status === "scheduled"
          ? "status-scheduled"
          : status === "closed"
            ? "status-closed"
            : "status-setup",
  };
}

function syncAutomaticClosure() {
  const settings = getElectionSettings();
  const readiness = getElectionReadiness(settings);
  const now = dayjs();
  const opensAt = settings.opensAt ? dayjs(settings.opensAt) : null;
  const closesAt = settings.closesAt ? dayjs(settings.closesAt) : null;

  if (
    settings.phase === "setup" &&
    readiness.isReady &&
    opensAt?.isValid() &&
    closesAt?.isValid() &&
    !now.isBefore(opensAt) &&
    now.isBefore(closesAt)
  ) {
    setSetting("election_phase", "open");
    logSystemAudit("election_auto_opened", {
      opensAt: settings.opensAt,
      closesAt: settings.closesAt,
    });
    return;
  }

  if (settings.phase === "open" && closesAt?.isValid() && !now.isBefore(closesAt)) {
    setSetting("election_phase", "closed");
    logSystemAudit("election_auto_closed", {
      closesAt: settings.closesAt,
    });
  }
}

function syncAutomaticNominationLifecycle() {
  const settings = getElectionSettings();
  const readiness = getNominationReadiness(settings);
  const now = dayjs();
  const opensAt = settings.nominationOpensAt ? dayjs(settings.nominationOpensAt) : null;
  const closesAt = settings.nominationClosesAt ? dayjs(settings.nominationClosesAt) : null;

  if (
    settings.nominationPhase === "setup" &&
    readiness.isReady &&
    opensAt?.isValid() &&
    closesAt?.isValid() &&
    !now.isBefore(opensAt) &&
    now.isBefore(closesAt)
  ) {
    setSetting("nomination_phase", "open");
    logSystemAudit("nomination_auto_opened", {
      opensAt: settings.nominationOpensAt,
      closesAt: settings.nominationClosesAt,
    });
    return;
  }

  if (
    settings.nominationPhase === "open" &&
    closesAt?.isValid() &&
    !now.isBefore(closesAt)
  ) {
    setSetting("nomination_phase", "closed");
    logSystemAudit("nomination_auto_closed", {
      closesAt: settings.nominationClosesAt,
    });
  }
}

function logAudit(req, actorType, actorIdentifier, action, details = {}) {
  db.prepare(`
    INSERT INTO audit_logs (
      actor_type,
      actor_identifier,
      action,
      details_json,
      ip_address,
      user_agent,
      created_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(
    actorType,
    actorIdentifier,
    action,
    JSON.stringify(details),
    req.ip || "",
    req.get("user-agent") || "",
    nowIso(),
  );
}

async function safeRemoveFile(filePath) {
  if (!filePath) {
    return;
  }

  await fsp.rm(filePath, { force: true });
}

async function safeRemoveUploadedRequestFiles(filesMap) {
  const files = Object.values(filesMap || {}).flat();

  for (const file of files) {
    await safeRemoveFile(file?.path);
  }
}

function isOtpVerificationEnabled() {
  return otpProvider === "twilio" || otpProvider === "dev";
}

function clearVoterSession(req) {
  req.session.voter = null;
  req.session.ballotSelections = null;
  req.session.pendingBallot = null;
  req.session.pendingVoterVerification = null;
}

function clearNominationSession(req) {
  req.session.nominationApplicant = null;
}

function maskPhoneNumber(value) {
  const normalized = normalizePhoneNumber(value);
  if (!normalized) {
    return "";
  }

  if (normalized.length <= 4) {
    return normalized;
  }

  const prefixLength = Math.min(3, Math.max(normalized.length - 4, 1));
  const visiblePrefix = normalized.slice(0, prefixLength);
  const visibleSuffix = normalized.slice(-4);
  const hiddenLength = Math.max(normalized.length - visiblePrefix.length - visibleSuffix.length, 0);

  return `${visiblePrefix}${"*".repeat(hiddenLength)}${visibleSuffix}`;
}

function generateDevOtpCode(length = devOtpCodeLength) {
  let code = "";

  while (code.length < length) {
    code += crypto.randomInt(0, 10).toString();
  }

  return code;
}

function hashOtpCode(code) {
  return crypto.createHash("sha256").update(String(code || "")).digest("hex");
}

function getOtpExpiryIso() {
  return dayjs().add(otpTtlMinutes, "minute").toISOString();
}

function getOtpResendAvailableIso() {
  return dayjs().add(otpResendCooldownSeconds, "second").toISOString();
}

function isPendingOtpExpired(pendingVerification) {
  if (!pendingVerification?.expiresAt) {
    return true;
  }

  const expiresAt = dayjs(pendingVerification.expiresAt);
  return !expiresAt.isValid() || !dayjs().isBefore(expiresAt);
}

function buildPendingVoterVerification(voterRecord, phoneNumber, smsPhoneNumber, challenge) {
  return {
    voterId: voterRecord.id,
    staffId: voterRecord.staffId,
    fullName: voterRecord.fullName,
    phoneNumber,
    maskedPhoneNumber: maskPhoneNumber(phoneNumber),
    smsPhoneNumber,
    provider: otpProvider,
    verificationSid: challenge.verificationSid || "",
    devCodeHash: challenge.devCodeHash || "",
    devCodePreview: challenge.devCodePreview || "",
    sentAt: nowIso(),
    expiresAt: getOtpExpiryIso(),
    resendAvailableAt: getOtpResendAvailableIso(),
    attempts: 0,
  };
}

function beginAuthenticatedVoterSession(req, voterRecord) {
  req.session.voter = {
    voterId: voterRecord.id,
    staffId: voterRecord.staffId,
    fullName: voterRecord.fullName,
  };
  req.session.ballotSelections = {};
  req.session.pendingBallot = null;
  req.session.pendingVoterVerification = null;
  req.session.voteComplete = null;
}

function beginNominationApplicantSession(req, voterRecord) {
  req.session.nominationApplicant = {
    voterId: voterRecord.id,
    staffId: voterRecord.staffId,
    fullName: voterRecord.fullName,
    phoneNumber: voterRecord.phoneNumber,
    department: voterRecord.department || "",
  };
}

async function parseOtpApiResponse(response) {
  const responseText = await response.text();

  try {
    return responseText ? JSON.parse(responseText) : {};
  } catch (_error) {
    return { message: responseText };
  }
}

async function sendTwilioOtpCode(smsPhoneNumber) {
  if (!twilioOtpConfigured) {
    throw new Error("The OTP SMS service is not configured yet. Add the Twilio Verify credentials first.");
  }

  const response = await fetch(
    `https://verify.twilio.com/v2/Services/${encodeURIComponent(twilioVerifyServiceSid)}/Verifications`,
    {
      method: "POST",
      headers: {
        authorization: `Basic ${Buffer.from(
          `${twilioAccountSid}:${twilioAuthToken}`,
        ).toString("base64")}`,
        "content-type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        To: smsPhoneNumber,
        Channel: "sms",
      }),
    },
  );
  const payload = await parseOtpApiResponse(response);

  if (!response.ok) {
    throw new Error(
      payload?.message ||
        "The OTP SMS could not be sent right now. Please try again in a moment.",
    );
  }

  return {
    verificationSid: payload?.sid || "",
  };
}

async function sendOtpChallenge(smsPhoneNumber) {
  if (otpProvider === "twilio") {
    return sendTwilioOtpCode(smsPhoneNumber);
  }

  if (otpProvider === "dev") {
    if (isProduction) {
      throw new Error(
        "Development OTP mode is not allowed in production. Configure Twilio Verify before using OTP on the live site.",
      );
    }

    const devCode = generateDevOtpCode();
    return {
      verificationSid: `DEV-${crypto.randomUUID()}`,
      devCodeHash: hashOtpCode(devCode),
      devCodePreview: devCode,
    };
  }

  return null;
}

async function verifyTwilioOtpCode(pendingVerification, code) {
  if (!twilioOtpConfigured) {
    throw new Error("The OTP SMS service is not configured yet. Add the Twilio Verify credentials first.");
  }

  const requestBody = new URLSearchParams({
    Code: code,
  });

  if (pendingVerification.verificationSid) {
    requestBody.set("VerificationSid", pendingVerification.verificationSid);
  } else {
    requestBody.set("To", pendingVerification.smsPhoneNumber);
  }

  const response = await fetch(
    `https://verify.twilio.com/v2/Services/${encodeURIComponent(
      twilioVerifyServiceSid,
    )}/VerificationCheck`,
    {
      method: "POST",
      headers: {
        authorization: `Basic ${Buffer.from(
          `${twilioAccountSid}:${twilioAuthToken}`,
        ).toString("base64")}`,
        "content-type": "application/x-www-form-urlencoded",
      },
      body: requestBody,
    },
  );
  const payload = await parseOtpApiResponse(response);

  if (response.status === 404) {
    return {
      approved: false,
      errorMessage: "This OTP has expired. Request a new code and try again.",
    };
  }

  if (!response.ok) {
    throw new Error(
      payload?.message || "The OTP could not be verified right now. Please try again.",
    );
  }

  return {
    approved: payload?.status === "approved" || payload?.valid === true,
    errorMessage:
      payload?.status === "max_attempts_reached"
        ? "Too many incorrect OTP attempts. Request a new code and try again."
        : payload?.status === "expired"
          ? "This OTP has expired. Request a new code and try again."
          : "The OTP code is incorrect. Please try again.",
  };
}

async function verifyOtpChallenge(pendingVerification, code) {
  if (otpProvider === "twilio") {
    return verifyTwilioOtpCode(pendingVerification, code);
  }

  if (otpProvider === "dev") {
    return {
      approved: hashOtpCode(code) === pendingVerification.devCodeHash,
      errorMessage: "The OTP code is incorrect. Please try again.",
    };
  }

  return {
    approved: false,
    errorMessage: "OTP verification is not enabled for this election.",
  };
}

function getDashboardMetrics() {
  return db.prepare(`
    SELECT
      (SELECT COUNT(*) FROM voters) AS totalVoters,
      (SELECT COUNT(*) FROM voters WHERE has_voted = 1) AS votedCount,
      (SELECT COUNT(*) FROM positions WHERE is_active = 1) AS totalPositions,
      (SELECT COUNT(*) FROM candidates WHERE is_active = 1) AS totalCandidates,
      (SELECT COUNT(*) FROM ballots) AS totalBallots
  `).get();
}

function getPositionsWithCandidates() {
  return db.prepare(`
    SELECT
      p.id AS positionId,
      p.name AS positionName,
      p.sort_order AS positionOrder,
      c.id AS candidateId,
      c.name AS candidateName,
      c.photo_path AS photoPath,
      c.bio AS bio,
      c.sort_order AS candidateOrder
    FROM positions p
    LEFT JOIN candidates c
      ON c.position_id = p.id
      AND c.is_active = 1
    WHERE p.is_active = 1
    ORDER BY p.sort_order ASC, p.name ASC, c.sort_order ASC, c.name ASC
  `).all();
}

function getBallotData() {
  const rows = getPositionsWithCandidates();
  const positionsMap = new Map();

  for (const row of rows) {
    if (!positionsMap.has(row.positionId)) {
      positionsMap.set(row.positionId, {
        id: row.positionId,
        name: row.positionName,
        sortOrder: row.positionOrder,
        candidates: [],
      });
    }

    if (row.candidateId) {
      positionsMap.get(row.positionId).candidates.push({
        id: row.candidateId,
        name: row.candidateName,
        photoPath: row.photoPath,
        bio: row.bio,
        sortOrder: row.candidateOrder,
      });
    }
  }

  return Array.from(positionsMap.values());
}

function getBallotSelectionMap(req) {
  return req.session.ballotSelections || {};
}

function buildSelectionsFromMap(ballotData, selectionMap) {
  const selections = [];
  const incompleteSteps = [];

  ballotData.forEach((position, index) => {
    const savedSelection =
      selectionMap[String(position.id)] || selectionMap[position.id] || null;
    const isSkipped = Boolean(
      savedSelection?.isSkipped || savedSelection?.skipped || savedSelection?.abstained,
    );

    if (isSkipped) {
      selections.push({
        positionId: position.id,
        positionName: position.name,
        candidateId: null,
        candidateName: "",
        isSkipped: true,
      });
      return;
    }

    const candidateId = parseInteger(savedSelection?.candidateId ?? savedSelection, 0);
    const candidate = position.candidates.find((entry) => entry.id === candidateId);

    if (!candidate) {
      incompleteSteps.push({
        stepNumber: index + 1,
        positionId: position.id,
        positionName: position.name,
      });
      return;
    }

    selections.push({
      positionId: position.id,
      positionName: position.name,
      candidateId: candidate.id,
      candidateName: candidate.name,
      isSkipped: false,
    });
  });

  return {
    selections,
    incompleteSteps,
  };
}

function getResultsExportPayload() {
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const metrics = getDashboardMetrics();
  const results = getResultsSummary();
  const generatedAt = nowIso();
  return {
    settings,
    electionState,
    metrics,
    results,
    generatedAt,
    nonVoters: Math.max(metrics.totalVoters - metrics.votedCount, 0),
    resultsStatusLabel: electionState.isClosed ? "Final" : "Provisional",
    declarationLabel: electionState.isClosed
      ? "Official declaration ready for sign-off."
      : "Provisional monitoring only until voting closes.",
  };
}

function getElectionArchives() {
  const rows = db.prepare(`
    SELECT
      id,
      election_name AS electionName,
      phase,
      opens_at AS opensAt,
      closes_at AS closesAt,
      archived_at AS archivedAt,
      total_voters AS totalVoters,
      votes_cast AS votesCast,
      turnout_rate AS turnoutRate,
      results_json AS resultsJson
    FROM election_archives
    ORDER BY archived_at DESC
  `).all();

  return rows.map((row) => {
    const parsedResults = safeJsonParse(row.resultsJson, []);
    return {
      ...row,
      resultsCount: parsedResults.length,
    };
  });
}

function getElectionArchiveById(archiveId) {
  const row = db.prepare(`
    SELECT
      id,
      election_name AS electionName,
      phase,
      opens_at AS opensAt,
      closes_at AS closesAt,
      archived_at AS archivedAt,
      total_voters AS totalVoters,
      votes_cast AS votesCast,
      turnout_rate AS turnoutRate,
      settings_json AS settingsJson,
      metrics_json AS metricsJson,
      results_json AS resultsJson
    FROM election_archives
    WHERE id = ?
  `).get(archiveId);

  if (!row) {
    return null;
  }

  return {
    ...row,
    settingsSnapshot: safeJsonParse(row.settingsJson, {}),
    metricsSnapshot: safeJsonParse(row.metricsJson, {}),
    results: safeJsonParse(row.resultsJson, []),
  };
}

function ensurePdfSpace(document, neededHeight = 80) {
  const bottomLimit = document.page.height - document.page.margins.bottom;
  if (document.y + neededHeight > bottomLimit) {
    document.addPage();
  }
}

function renderResultsPdf(document, payload) {
  const {
    settings,
    metrics,
    results,
    generatedAt,
    nonVoters,
    resultsStatusLabel,
  } = payload;
  const logoPath = settings.organizationLogoPath
    ? resolveAssetPath(settings.organizationLogoPath)
    : "";

  if (logoPath && fs.existsSync(logoPath)) {
    try {
      document.image(logoPath, 50, 42, { fit: [56, 56], align: "left" });
    } catch (_error) {
      // Ignore invalid image parsing and continue generating the PDF.
    }
  }

  document
    .font("Helvetica-Bold")
    .fontSize(22)
    .fillColor("#102338")
    .text(settings.electionName, 120, 48, { align: "left" });

  document
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#5d6d80")
    .text(`${resultsStatusLabel} election results`, 120, 78)
    .text(`Generated: ${formatDateTime(generatedAt)}`, 120, 94)
    .text(`Voting opened: ${formatDateTime(settings.opensAt)}`, 120, 110)
    .text(`Voting closed: ${formatDateTime(settings.closesAt)}`, 120, 126);

  document
    .moveTo(50, 148)
    .lineTo(545, 148)
    .strokeColor("#d8c197")
    .lineWidth(1)
    .stroke();

  document.y = 168;
  document
    .font("Helvetica-Bold")
    .fontSize(12)
    .fillColor("#102338")
    .text("Official Results Summary");
  document
    .moveDown(0.4)
    .font("Helvetica")
    .fontSize(11)
    .fillColor("#102338")
    .text(`Total voters: ${metrics.totalVoters}`)
    .text(`Votes cast: ${metrics.votedCount}`)
    .text(`Non-voters: ${nonVoters}`)
    .text(
      `Turnout: ${
        metrics.totalVoters
          ? formatPercent(metrics.votedCount / metrics.totalVoters)
          : "0.0%"
      }`,
    )
    .text(`Status: ${resultsStatusLabel}`);

  results.forEach((result) => {
    ensurePdfSpace(document, 120 + result.candidates.length * 28);

    document
      .moveDown(1.1)
      .font("Helvetica-Bold")
      .fontSize(15)
      .fillColor("#102338")
      .text(result.name);

    document
      .moveDown(0.2)
      .font("Helvetica")
      .fontSize(10)
      .fillColor("#5d6d80")
      .text(`Winner: ${result.winnerLabel}`)
      .text(`Margin of victory: ${result.marginLabel}`)
      .text(`Valid votes: ${result.totalVotes}`);

    const headerY = document.y + 10;
    document
      .font("Helvetica-Bold")
      .fontSize(10)
      .fillColor("#102338")
      .text("Candidate", 50, headerY, { width: 270 })
      .text("Votes", 345, headerY, { width: 50, align: "right" })
      .text("Share", 405, headerY, { width: 60, align: "right" })
      .text("Status", 475, headerY, { width: 60, align: "right" });

    document
      .moveTo(50, headerY + 16)
      .lineTo(545, headerY + 16)
      .strokeColor("#e3d7c0")
      .lineWidth(1)
      .stroke();

    let rowY = headerY + 26;

    result.candidates.forEach((candidate) => {
      ensurePdfSpace(document, 38);
      const candidatePhotoPath = candidate.photoPath
        ? resolveAssetPath(candidate.photoPath)
        : "";
      const hasCandidatePhoto = candidatePhotoPath && fs.existsSync(candidatePhotoPath);

      if (hasCandidatePhoto) {
        try {
          document.image(candidatePhotoPath, 50, rowY - 2, { fit: [18, 18] });
        } catch (_error) {
          // Ignore invalid image parsing and continue with the textual row.
        }
      }

      document
        .font("Helvetica")
        .fontSize(10)
        .fillColor("#102338")
        .text(candidate.name, hasCandidatePhoto ? 74 : 50, rowY, { width: 246 })
        .text(String(candidate.voteCount), 345, rowY, { width: 50, align: "right" })
        .text(
          result.totalVotes
            ? formatPercent(candidate.voteCount / result.totalVotes)
            : "0.0%",
          405,
          rowY,
          { width: 60, align: "right" },
        );
      document
        .font(candidate.isLeading && result.totalVotes > 0 ? "Helvetica-Bold" : "Helvetica")
        .text(
          candidate.isLeading && result.totalVotes > 0 ? "Winner" : "Candidate",
          475,
          rowY,
          { width: 60, align: "right" },
        );

      rowY += 24;
    });

    document.y = rowY + 4;
  });

  ensurePdfSpace(document, 120);
  document
    .moveDown(1)
    .font("Helvetica-Bold")
    .fontSize(12)
    .fillColor("#102338")
    .text(settings.declarationTitle);

  document
    .moveDown(0.4)
    .font("Helvetica")
    .fontSize(11)
    .fillColor("#102338")
    .text(`Declared by: ${settings.committeeName}`)
    .text(`Committee: ${settings.committeeName}`)
    .text(`Chairman: ${settings.chairmanName || "Chairman"}`)
    .text(`Secretary: ${settings.secretaryName || "Secretary"}`)
    .text(`Date declared: ${formatDateTime(generatedAt)}`);

  const chairmanSignaturePath = settings.chairmanSignaturePath
    ? resolveAssetPath(settings.chairmanSignaturePath)
    : "";
  const secretarySignaturePath = settings.secretarySignaturePath
    ? resolveAssetPath(settings.secretarySignaturePath)
    : "";
  const signatureY = document.y + 36;
  const signatureImageY = Math.max(signatureY - 46, document.y + 8);

  if (chairmanSignaturePath && fs.existsSync(chairmanSignaturePath)) {
    try {
      document.image(chairmanSignaturePath, 70, signatureImageY, {
        fit: [140, 40],
        align: "left",
      });
    } catch (_error) {
      // Ignore invalid image parsing and continue with the signature line.
    }
  }

  if (secretarySignaturePath && fs.existsSync(secretarySignaturePath)) {
    try {
      document.image(secretarySignaturePath, 340, signatureImageY, {
        fit: [140, 40],
        align: "left",
      });
    } catch (_error) {
      // Ignore invalid image parsing and continue with the signature line.
    }
  }

  document
    .moveTo(60, signatureY)
    .lineTo(240, signatureY)
    .strokeColor("#102338")
    .lineWidth(1)
    .stroke();
  document
    .moveTo(330, signatureY)
    .lineTo(510, signatureY)
    .strokeColor("#102338")
    .lineWidth(1)
    .stroke();
  document
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#5d6d80")
    .text(settings.chairmanName || "Chairman", 85, signatureY + 8)
    .text("Chairman Signature", 75, signatureY + 22)
    .text(settings.secretaryName || "Secretary", 355, signatureY + 8)
    .text("Secretary Signature", 350, signatureY + 22);
}

function renderNominationFormPdf(document, payload) {
  const { settings, positions, generatedAt } = payload;
  const logoPath = settings.organizationLogoPath
    ? resolveAssetPath(settings.organizationLogoPath)
    : "";

  if (logoPath && fs.existsSync(logoPath)) {
    try {
      document.image(logoPath, 50, 42, { fit: [56, 56], align: "left" });
    } catch (_error) {
      // Ignore image parsing errors and continue with the PDF.
    }
  }

  document
    .font("Helvetica-Bold")
    .fontSize(21)
    .fillColor("#102338")
    .text(settings.electionName, 120, 48, { align: "left" })
    .fontSize(12)
    .text("Nomination Form", 120, 78);

  document
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#5d6d80")
    .text(`Generated: ${formatDateTime(generatedAt)}`, 120, 96)
    .text(`Nomination opens: ${formatDateTime(settings.nominationOpensAt)}`, 120, 112)
    .text(`Nomination closes: ${formatDateTime(settings.nominationClosesAt)}`, 120, 128);

  document
    .moveTo(50, 148)
    .lineTo(545, 148)
    .strokeColor("#d8c197")
    .lineWidth(1)
    .stroke();

  document.y = 172;
  document
    .font("Helvetica-Bold")
    .fontSize(12)
    .fillColor("#102338")
    .text("Applicant Information");

  const sections = [
    "Full Name: ________________________________________________",
    "Staff ID: ______________________",
    "Phone Number: ______________________",
    "Department: ______________________________________________",
    "",
    "Position Applying For: ____________________________________",
    "Proposer Name: ____________________________________________",
    "Seconder Name: ____________________________________________",
    "",
    "Short Profile / Bio:",
    "____________________________________________________________",
    "____________________________________________________________",
    "____________________________________________________________",
    "",
    "Manifesto / Message:",
    "____________________________________________________________",
    "____________________________________________________________",
    "____________________________________________________________",
    "",
    "Declaration:",
    "I confirm that the information provided for this nomination is accurate.",
    "Applicant Signature: _______________________________________",
  ];

  document
    .moveDown(0.6)
    .font("Helvetica")
    .fontSize(11)
    .fillColor("#102338");

  sections.forEach((line) => {
    ensurePdfSpace(document, 24);
    document.text(line);
  });

  document.moveDown(1);
  document.font("Helvetica-Bold").text("Available Positions");
  document.font("Helvetica");
  positions.forEach((position, index) => {
    ensurePdfSpace(document, 20);
    document.text(`${index + 1}. ${position.name}`);
  });
}

function renderNominationReportPdf(document, payload) {
  const { settings, nominations, generatedAt } = payload;
  const logoPath = settings.organizationLogoPath
    ? resolveAssetPath(settings.organizationLogoPath)
    : "";

  if (logoPath && fs.existsSync(logoPath)) {
    try {
      document.image(logoPath, 50, 42, { fit: [56, 56], align: "left" });
    } catch (_error) {
      // Ignore image parsing errors and continue with the PDF.
    }
  }

  document
    .font("Helvetica-Bold")
    .fontSize(20)
    .fillColor("#102338")
    .text(settings.electionName, 120, 48, { align: "left" })
    .fontSize(12)
    .text("Nomination Applications Report", 120, 78);

  document
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#5d6d80")
    .text(`Generated: ${formatDateTime(generatedAt)}`, 120, 96)
    .text(
      `Nomination window: ${formatDateTime(settings.nominationOpensAt)} to ${formatDateTime(
        settings.nominationClosesAt,
      )}`,
      120,
      112,
    );

  document
    .moveTo(50, 142)
    .lineTo(545, 142)
    .strokeColor("#d8c197")
    .lineWidth(1)
    .stroke();

  document.y = 160;

  if (nominations.length === 0) {
    document
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor("#102338")
      .text("No nomination applications have been submitted yet.");
    return;
  }

  nominations.forEach((nomination, index) => {
    ensurePdfSpace(document, 132);
    document
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor("#102338")
      .text(`${index + 1}. ${nomination.fullName}`);

    document
      .moveDown(0.25)
      .font("Helvetica")
      .fontSize(10)
      .fillColor("#42586f")
      .text(`Staff ID: ${nomination.staffId}`)
      .text(`Position: ${nomination.positionName}`)
      .text(`Department: ${nomination.department || "Not provided"}`)
      .text(`Status: ${nomination.statusMeta.label}`)
      .text(`Submitted: ${formatDateTime(nomination.submittedAt)}`)
      .text(`Proposer: ${nomination.proposerName}`)
      .text(`Seconder: ${nomination.seconderName}`)
      .text(`Manifesto: ${nomination.manifesto || "Not provided"}`)
      .text(`Admin Notes: ${nomination.adminNotes || "None"}`);

    document
      .moveDown(0.55)
      .moveTo(50, document.y)
      .lineTo(545, document.y)
      .strokeColor("#e3d7c0")
      .lineWidth(1)
      .stroke()
      .moveDown(0.7);
  });
}

function getVoters() {
  return db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      department,
      has_voted AS hasVoted,
      voted_at AS votedAt,
      created_at AS createdAt
    FROM voters
    ORDER BY staff_id ASC
  `).all();
}

function getPositions() {
  return db.prepare(`
    SELECT
      id,
      name,
      sort_order AS sortOrder,
      created_at AS createdAt
    FROM positions
    WHERE is_active = 1
    ORDER BY sort_order ASC, name ASC
  `).all();
}

function getCandidates() {
  return db.prepare(`
    SELECT
      c.id,
      c.name,
      c.photo_path AS photoPath,
      c.bio,
      c.sort_order AS sortOrder,
      p.name AS positionName,
      p.id AS positionId
    FROM candidates c
    INNER JOIN positions p ON p.id = c.position_id
    WHERE c.is_active = 1
    ORDER BY p.sort_order ASC, p.name ASC, c.sort_order ASC, c.name ASC
  `).all();
}

function getRegisteredStaffRecord(staffId, phoneNumber) {
  const normalizedStaffId = normalizeStaffId(staffId);
  const normalizedPhoneNumber = normalizePhoneNumber(phoneNumber);

  if (!normalizedStaffId || !normalizedPhoneNumber) {
    return null;
  }

  const voterRecord = db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      department
    FROM voters
    WHERE staff_id = ?
  `).get(normalizedStaffId);

  if (!voterRecord || voterRecord.phoneNumber !== normalizedPhoneNumber) {
    return null;
  }

  return voterRecord;
}

function getNominationStatusMeta(status) {
  const normalizedStatus = String(status || "pending").trim().toLowerCase();

  switch (normalizedStatus) {
    case "approved":
      return {
        value: "approved",
        label: "Approved",
        className: "inline-badge--success",
        canApplicantEdit: false,
      };
    case "rejected":
      return {
        value: "rejected",
        label: "Rejected",
        className: "inline-badge--danger",
        canApplicantEdit: false,
      };
    case "correction_requested":
      return {
        value: "correction_requested",
        label: "Correction Requested",
        className: "inline-badge--warning",
        canApplicantEdit: true,
      };
    case "published":
      return {
        value: "published",
        label: "Published as Candidate",
        className: "inline-badge--info",
        canApplicantEdit: false,
      };
    default:
      return {
        value: "pending",
        label: "Pending Review",
        className: "inline-badge--muted",
        canApplicantEdit: false,
      };
  }
}

function mapNominationRow(row) {
  const statusMeta = getNominationStatusMeta(row.status);
  return {
    ...row,
    statusMeta,
  };
}

function getNominationMetrics() {
  const totals = db.prepare(`
    SELECT
      COUNT(*) AS total,
      SUM(CASE WHEN status = 'pending' THEN 1 ELSE 0 END) AS pendingCount,
      SUM(CASE WHEN status = 'approved' THEN 1 ELSE 0 END) AS approvedCount,
      SUM(CASE WHEN status = 'rejected' THEN 1 ELSE 0 END) AS rejectedCount,
      SUM(CASE WHEN status = 'correction_requested' THEN 1 ELSE 0 END) AS correctionCount,
      SUM(CASE WHEN status = 'published' THEN 1 ELSE 0 END) AS publishedCount
    FROM nominations
  `).get();

  return {
    total: Number(totals.total || 0),
    pendingCount: Number(totals.pendingCount || 0),
    approvedCount: Number(totals.approvedCount || 0),
    rejectedCount: Number(totals.rejectedCount || 0),
    correctionCount: Number(totals.correctionCount || 0),
    publishedCount: Number(totals.publishedCount || 0),
  };
}

function getNominationList() {
  const rows = db.prepare(`
    SELECT
      n.id,
      n.voter_id AS voterId,
      n.position_id AS positionId,
      n.staff_id AS staffId,
      n.full_name AS fullName,
      n.phone_number AS phoneNumber,
      n.department,
      n.photo_path AS photoPath,
      n.bio,
      n.manifesto,
      n.proposer_name AS proposerName,
      n.seconder_name AS seconderName,
      n.declaration_accepted AS declarationAccepted,
      n.status,
      n.admin_notes AS adminNotes,
      n.reviewed_at AS reviewedAt,
      n.reviewed_by AS reviewedBy,
      n.published_candidate_id AS publishedCandidateId,
      n.submitted_at AS submittedAt,
      n.created_at AS createdAt,
      n.updated_at AS updatedAt,
      p.name AS positionName
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    ORDER BY n.submitted_at DESC, n.id DESC
  `).all();

  return rows.map(mapNominationRow);
}

function getNominationsForVoter(voterId) {
  const rows = db.prepare(`
    SELECT
      n.id,
      n.voter_id AS voterId,
      n.position_id AS positionId,
      n.staff_id AS staffId,
      n.full_name AS fullName,
      n.phone_number AS phoneNumber,
      n.department,
      n.photo_path AS photoPath,
      n.bio,
      n.manifesto,
      n.proposer_name AS proposerName,
      n.seconder_name AS seconderName,
      n.declaration_accepted AS declarationAccepted,
      n.status,
      n.admin_notes AS adminNotes,
      n.reviewed_at AS reviewedAt,
      n.reviewed_by AS reviewedBy,
      n.published_candidate_id AS publishedCandidateId,
      n.submitted_at AS submittedAt,
      n.created_at AS createdAt,
      n.updated_at AS updatedAt,
      p.name AS positionName
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    WHERE n.voter_id = ?
    ORDER BY n.submitted_at DESC, n.id DESC
  `).all(voterId);

  return rows.map(mapNominationRow);
}

function getNominationById(nominationId) {
  const row = db.prepare(`
    SELECT
      n.id,
      n.voter_id AS voterId,
      n.position_id AS positionId,
      n.staff_id AS staffId,
      n.full_name AS fullName,
      n.phone_number AS phoneNumber,
      n.department,
      n.photo_path AS photoPath,
      n.bio,
      n.manifesto,
      n.proposer_name AS proposerName,
      n.seconder_name AS seconderName,
      n.declaration_accepted AS declarationAccepted,
      n.status,
      n.admin_notes AS adminNotes,
      n.reviewed_at AS reviewedAt,
      n.reviewed_by AS reviewedBy,
      n.published_candidate_id AS publishedCandidateId,
      n.submitted_at AS submittedAt,
      n.created_at AS createdAt,
      n.updated_at AS updatedAt,
      p.name AS positionName
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    WHERE n.id = ?
  `).get(nominationId);

  return row ? mapNominationRow(row) : null;
}

function getAuditLogs(limit = 150) {
  return db.prepare(`
    SELECT
      id,
      actor_type AS actorType,
      actor_identifier AS actorIdentifier,
      action,
      details_json AS detailsJson,
      ip_address AS ipAddress,
      created_at AS createdAt
    FROM audit_logs
    ORDER BY created_at DESC
    LIMIT ?
  `).all(limit);
}

function getResultsSummary() {
  const rows = db.prepare(`
    SELECT
      p.id AS positionId,
      p.name AS positionName,
      p.sort_order AS positionOrder,
      c.id AS candidateId,
      c.name AS candidateName,
      c.photo_path AS photoPath,
      COUNT(be.id) AS voteCount
    FROM positions p
    LEFT JOIN candidates c
      ON c.position_id = p.id
      AND c.is_active = 1
    LEFT JOIN ballot_entries be ON be.candidate_id = c.id
    WHERE p.is_active = 1
    GROUP BY p.id, c.id
    ORDER BY p.sort_order ASC, p.name ASC, c.sort_order ASC, c.name ASC
  `).all();

  const positionsMap = new Map();

  for (const row of rows) {
    if (!positionsMap.has(row.positionId)) {
      positionsMap.set(row.positionId, {
        id: row.positionId,
        name: row.positionName,
        sortOrder: row.positionOrder,
        totalVotes: 0,
        candidates: [],
        winnerLabel: "No votes yet",
      });
    }

    if (row.candidateId) {
      const voteCount = Number(row.voteCount || 0);
      positionsMap.get(row.positionId).candidates.push({
        id: row.candidateId,
        name: row.candidateName,
        photoPath: row.photoPath,
        voteCount,
      });
      positionsMap.get(row.positionId).totalVotes += voteCount;
    }
  }

  const summaries = Array.from(positionsMap.values());

  for (const summary of summaries) {
    if (summary.candidates.length === 0) {
      summary.winnerLabel = "No candidates added";
      summary.marginVotes = 0;
      summary.marginLabel = "No candidates available";
      continue;
    }

    const sortedByVotes = [...summary.candidates].sort((left, right) => {
      if (right.voteCount !== left.voteCount) {
        return right.voteCount - left.voteCount;
      }

      return left.name.localeCompare(right.name);
    });
    const highestVoteCount = sortedByVotes[0]?.voteCount || 0;
    const secondHighestVoteCount = sortedByVotes[1]?.voteCount || 0;

    const winners = summary.candidates.filter(
      (candidate) => candidate.voteCount === highestVoteCount,
    );

    summary.winnerLabel =
      highestVoteCount === 0
        ? "No votes recorded"
        : winners.length > 1
          ? `Tie: ${winners.map((candidate) => candidate.name).join(", ")}`
          : winners[0].name;
    summary.marginVotes =
      highestVoteCount > 0 && winners.length === 1
        ? Math.max(highestVoteCount - secondHighestVoteCount, 0)
        : 0;
    summary.marginLabel =
      highestVoteCount === 0
        ? "No votes recorded"
        : winners.length > 1
          ? "Tie"
          : sortedByVotes.length === 1
            ? "Unopposed"
            : `${summary.marginVotes} vote lead`;

    summary.candidates = summary.candidates.map((candidate) => ({
      ...candidate,
      shareRatio: summary.totalVotes ? candidate.voteCount / summary.totalVotes : 0,
      isLeading: highestVoteCount > 0 && candidate.voteCount === highestVoteCount,
    }));
  }

  return summaries;
}

function formatDashboardAuditAction(action) {
  return String(action || "activity")
    .split("_")
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function formatDashboardRelativeTime(value) {
  const target = dayjs(value);

  if (!target.isValid()) {
    return "Just now";
  }

  const seconds = Math.max(dayjs().diff(target, "second"), 0);

  if (seconds < 60) {
    return `${seconds}s ago`;
  }

  const minutes = Math.floor(seconds / 60);
  if (minutes < 60) {
    return `${minutes}m ago`;
  }

  const hours = Math.floor(minutes / 60);
  if (hours < 24) {
    return `${hours}h ago`;
  }

  const days = Math.floor(hours / 24);
  if (days < 7) {
    return `${days}d ago`;
  }

  return target.format("DD MMM");
}

function summarizeAuditDetails(details) {
  if (!details || typeof details !== "object") {
    return "";
  }

  return (
    details.candidateName ||
    details.positionName ||
    details.staffId ||
    details.electionName ||
    details.reason ||
    details.position ||
    ""
  );
}

function getDashboardActivity(limit = 5) {
  return getAuditLogs(limit).map((log) => {
    const details = safeJsonParse(log.detailsJson, {});
    return {
      ...log,
      actionLabel: formatDashboardAuditAction(log.action),
      detailLabel: summarizeAuditDetails(details),
      timeAgo: formatDashboardRelativeTime(log.createdAt),
      accentClass:
        log.actorType === "voter"
          ? "dashboard-activity__dot--cyan"
          : log.actorType === "system"
            ? "dashboard-activity__dot--amber"
            : "dashboard-activity__dot--green",
    };
  });
}

function getDashboardTimeline(settings, pointCount = 6) {
  const safePointCount = Math.max(pointCount, 2);
  const now = dayjs();
  const configuredOpen = settings.opensAt ? dayjs(settings.opensAt) : null;
  const configuredClose = settings.closesAt ? dayjs(settings.closesAt) : null;

  let end = configuredClose?.isValid() && configuredClose.isBefore(now) ? configuredClose : now;
  if (!end?.isValid()) {
    end = now;
  }

  let start = configuredOpen?.isValid() ? configuredOpen : end.subtract(safePointCount - 1, "hour");
  if (!start?.isValid() || !start.isBefore(end)) {
    start = end.subtract(safePointCount - 1, "hour");
  }

  const totalSpanMs = Math.max(end.valueOf() - start.valueOf(), 1);
  const ballotTimes = db
    .prepare(`
      SELECT submitted_at AS submittedAt
      FROM ballots
      ORDER BY submitted_at ASC
    `)
    .all()
    .map((row) => dayjs(row.submittedAt))
    .filter((value) => value.isValid())
    .map((value) => value.valueOf());

  const series = [];
  let cursor = 0;
  let runningTotal = 0;

  for (let index = 0; index < safePointCount; index += 1) {
    const ratio = safePointCount === 1 ? 1 : index / (safePointCount - 1);
    const pointMs = start.valueOf() + totalSpanMs * ratio;

    while (cursor < ballotTimes.length && ballotTimes[cursor] <= pointMs) {
      runningTotal += 1;
      cursor += 1;
    }

    series.push({
      label: dayjs(pointMs).format(totalSpanMs >= 86400000 ? "DD MMM" : "h A"),
      value: runningTotal,
    });
  }

  return series;
}

function buildLineChartShape(series, width = 340, height = 168) {
  const values = series.map((point) => Number(point.value || 0));
  const maxValue = Math.max(...values, 1);
  const inset = 14;
  const plotWidth = Math.max(width - inset * 2, 1);
  const plotHeight = Math.max(height - inset * 2, 1);

  const points = values.map((value, index) => {
    const x =
      values.length === 1
        ? width / 2
        : inset + (plotWidth * index) / Math.max(values.length - 1, 1);
    const y = inset + plotHeight - (value / maxValue) * plotHeight;
    return `${x.toFixed(2)},${y.toFixed(2)}`;
  });

  return {
    width,
    height,
    maxValue,
    points: points.join(" "),
    areaPoints: `${inset},${height - inset} ${points.join(" ")} ${width - inset},${height - inset}`,
  };
}

function getDashboardTopCandidates(results, limit = 4) {
  return results
    .flatMap((position) =>
      position.candidates.map((candidate) => ({
        ...candidate,
        positionName: position.name,
        totalVotesForPosition: position.totalVotes,
      })),
    )
    .sort((left, right) => {
      if (right.voteCount !== left.voteCount) {
        return right.voteCount - left.voteCount;
      }

      if (right.shareRatio !== left.shareRatio) {
        return right.shareRatio - left.shareRatio;
      }

      return left.name.localeCompare(right.name);
    })
    .slice(0, limit);
}

function getDashboardPositionBars(results, limit = 5) {
  const positions = results.slice(0, limit);
  const maxVotes = Math.max(...positions.map((position) => position.totalVotes), 1);

  return positions.map((position, index) => ({
    ...position,
    fillWidth: `${Math.max((position.totalVotes / maxVotes) * 100, position.totalVotes > 0 ? 12 : 0)}%`,
    toneClass: `dashboard-bar--tone-${(index % 5) + 1}`,
  }));
}

function ensureSetupMode(req, res) {
  const settings = getElectionSettings();

  if (settings.phase !== "setup") {
    setFlash(
      req,
      "error",
      "Setup is locked once voting has been opened. Create a new election cycle to make structural changes.",
    );
    res.redirect("/admin");
    return false;
  }

  return true;
}

function createImageUpload(destinationDirectory, errorMessage) {
  return multer({
    storage: multer.diskStorage({
      destination(_req, _file, callback) {
        callback(null, destinationDirectory);
      },
      filename(_req, file, callback) {
        const extension = path.extname(file.originalname || "").toLowerCase() || ".jpg";
        callback(null, `${Date.now()}-${crypto.randomUUID()}${extension}`);
      },
    }),
    limits: { fileSize: 5 * 1024 * 1024 },
    fileFilter(_req, file, callback) {
      if (!file.mimetype.startsWith("image/")) {
        callback(new Error(errorMessage));
        return;
      }

      callback(null, true);
    },
  });
}

const candidateUpload = createImageUpload(
  candidateUploadsDirectory,
  "Candidate photos must be image files.",
);

const nominationUpload = createImageUpload(
  nominationUploadsDirectory,
  "Nomination photos must be image files.",
);

const brandingUpload = createImageUpload(
  brandingUploadsDirectory,
  "Organization logos must be image files.",
);

const declarationUpload = createImageUpload(
  brandingUploadsDirectory,
  "Declaration signatures must be image files.",
);

const voterImportUpload = multer({
  storage: multer.diskStorage({
    destination(_req, _file, callback) {
      callback(null, importUploadsDirectory);
    },
    filename(_req, file, callback) {
      const extension = path.extname(file.originalname || "").toLowerCase() || ".xlsx";
      callback(null, `${Date.now()}-${crypto.randomUUID()}${extension}`);
    },
  }),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter(_req, file, callback) {
    const extension = path.extname(file.originalname || "").toLowerCase();
    if (![".xlsx", ".xls", ".csv"].includes(extension)) {
      callback(new Error("Voter imports must be Excel or CSV files."));
      return;
    }

    callback(null, true);
  },
});

app.set("view engine", "ejs");
app.set("views", path.join(process.cwd(), "views"));

app.use(
  helmet({
    contentSecurityPolicy: false,
    crossOriginResourcePolicy: false,
  }),
);
app.use(express.urlencoded({ extended: true }));
app.use("/uploads", express.static(uploadsRootDirectory));
app.use(express.static(publicDirectory));
app.use(
  session({
    secret: sessionSecret,
    proxy: sessionSecureCookie,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: sessionSecureCookie,
      maxAge: 1000 * 60 * 60 * 4,
    },
  }),
);

app.use((req, res, next) => {
  syncAutomaticClosure();
  syncAutomaticNominationLifecycle();

  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const nominationState = computeNominationState(settings);

  res.locals.settings = settings;
  res.locals.electionState = electionState;
  res.locals.nominationState = nominationState;
  res.locals.currentPath = req.path;
  res.locals.admin = req.session.admin || null;
  res.locals.voter = req.session.voter || null;
  res.locals.nominationApplicant = req.session.nominationApplicant || null;
  res.locals.currentYear = new Date().getFullYear();
  res.locals.formatDateTime = formatDateTime;
  res.locals.formatPercent = formatPercent;
  res.locals.voteClosesAtMs = settings.closesAt ? dayjs(settings.closesAt).valueOf() : null;
  res.locals.getInitials = getInitials;
  res.locals.flash = req.session.flash || null;
  delete req.session.flash;

  next();
});

app.get("/", (req, res) => {
  const metrics = getDashboardMetrics();
  res.render("home", { pageTitle: "Election Portal", metrics });
});

app.get("/health", (_req, res) => {
  res.status(200).json({ ok: true });
});

app.get("/nomination/form/download", (_req, res) => {
  const settings = getElectionSettings();
  const positions = getPositions();
  const filename = `${toSafeFilename(settings.electionName)}-nomination-form.pdf`;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: "A4",
    margin: 50,
    info: {
      Title: `${settings.electionName} Nomination Form`,
      Author: "Organization Vote Portal",
    },
  });

  document.pipe(res);
  renderNominationFormPdf(document, {
    settings,
    positions,
    generatedAt: nowIso(),
  });
  document.end();
});

app.get("/nomination/login", (req, res) => {
  if (req.session.nominationApplicant) {
    return res.redirect("/nomination/form");
  }

  return res.render("nomination-login", {
    pageTitle: "Nomination Login",
    mode: "apply",
  });
});

app.post("/nomination/login", (req, res) => {
  const staffId = normalizeStaffId(req.body.staffId);
  const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
  const settings = getElectionSettings();
  const nominationState = computeNominationState(settings);

  if (!staffId || !phoneNumber) {
    setFlash(req, "error", "Enter both your staff ID and phone number.");
    return res.redirect("/nomination/login");
  }

  if (!nominationState.isOpen) {
    setFlash(req, "error", nominationState.message);
    return res.redirect("/nomination/login");
  }

  const voterRecord = getRegisteredStaffRecord(staffId, phoneNumber);

  if (!voterRecord) {
    logAudit(req, "nomination", staffId || "unknown", "nomination_login_failed", {
      reason: "invalid_credentials",
    });
    setFlash(
      req,
      "error",
      "Your staff ID and phone number do not match a registered staff record.",
    );
    return res.redirect("/nomination/login");
  }

  beginNominationApplicantSession(req, voterRecord);
  logAudit(req, "nomination", voterRecord.staffId, "nomination_login_success");
  return res.redirect("/nomination/form");
});

app.get("/nomination/status/login", (req, res) => {
  if (req.session.nominationApplicant) {
    return res.redirect("/nomination/status");
  }

  return res.render("nomination-login", {
    pageTitle: "Nomination Status",
    mode: "status",
  });
});

app.post("/nomination/status/login", (req, res) => {
  const staffId = normalizeStaffId(req.body.staffId);
  const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);

  if (!staffId || !phoneNumber) {
    setFlash(req, "error", "Enter both your staff ID and phone number.");
    return res.redirect("/nomination/status/login");
  }

  const voterRecord = getRegisteredStaffRecord(staffId, phoneNumber);

  if (!voterRecord) {
    logAudit(req, "nomination", staffId || "unknown", "nomination_status_login_failed", {
      reason: "invalid_credentials",
    });
    setFlash(
      req,
      "error",
      "Your staff ID and phone number do not match a registered staff record.",
    );
    return res.redirect("/nomination/status/login");
  }

  beginNominationApplicantSession(req, voterRecord);
  logAudit(req, "nomination", voterRecord.staffId, "nomination_status_login_success");
  return res.redirect("/nomination/status");
});

app.get("/nomination/form", (req, res) => {
  const applicant = req.session.nominationApplicant;

  if (!applicant) {
    setFlash(req, "error", "Sign in with your staff ID and phone number to continue.");
    return res.redirect("/nomination/login");
  }

  const settings = getElectionSettings();
  const nominationState = computeNominationState(settings);
  const positions = getPositions();
  const nominations = getNominationsForVoter(applicant.voterId);
  const editNominationId = parseInteger(req.query.edit, 0);
  const editableNomination =
    editNominationId > 0
      ? nominations.find(
          (nomination) =>
            nomination.id === editNominationId && nomination.statusMeta.canApplicantEdit,
        ) || null
      : null;

  if (!nominationState.isOpen && !editableNomination) {
    setFlash(req, "error", nominationState.message);
    return res.redirect("/nomination/status");
  }

  return res.render("nomination-form", {
    pageTitle: "Apply for Nomination",
    applicant,
    positions,
    nominations,
    editableNomination,
  });
});

app.post(
  "/nomination/form",
  nominationUpload.single("photo"),
  async (req, res) => {
    const applicant = req.session.nominationApplicant;

    if (!applicant) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Sign in with your staff ID and phone number to continue.");
      return res.redirect("/nomination/login");
    }

    const settings = getElectionSettings();
    const nominationState = computeNominationState(settings);
    const nominationId = parseInteger(req.body.nominationId, 0);
    const fullName = String(req.body.fullName || "").trim();
    const department = String(req.body.department || "").trim();
    const positionId = parseInteger(req.body.positionId, 0);
    const bio = String(req.body.bio || "").trim();
    const manifesto = String(req.body.manifesto || "").trim();
    const proposerName = String(req.body.proposerName || "").trim();
    const seconderName = String(req.body.seconderName || "").trim();
    const declarationAccepted = req.body.declarationAccepted === "on";

    const existingEditableNomination =
      nominationId > 0 ? getNominationById(nominationId) : null;
    const isCorrectionResubmission =
      existingEditableNomination &&
      existingEditableNomination.voterId === applicant.voterId &&
      existingEditableNomination.statusMeta.canApplicantEdit;

    if (!nominationState.isOpen && !isCorrectionResubmission) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", nominationState.message);
      return res.redirect("/nomination/status");
    }

    if (
      !fullName ||
      !department ||
      !positionId ||
      !bio ||
      !manifesto ||
      !proposerName ||
      !seconderName ||
      !declarationAccepted
    ) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "Complete every nomination field and confirm the declaration before submitting.",
      );
      return res.redirect(
        isCorrectionResubmission ? `/nomination/form?edit=${existingEditableNomination.id}` : "/nomination/form",
      );
    }

    const position = db.prepare(`
      SELECT id, name
      FROM positions
      WHERE id = ?
        AND is_active = 1
    `).get(positionId);

    if (!position) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Choose a valid position before submitting the nomination.");
      return res.redirect(
        isCorrectionResubmission ? `/nomination/form?edit=${existingEditableNomination.id}` : "/nomination/form",
      );
    }

    const duplicateNomination = db.prepare(`
      SELECT id, status
      FROM nominations
      WHERE voter_id = ?
        AND position_id = ?
        AND id <> ?
    `).get(applicant.voterId, positionId, isCorrectionResubmission ? existingEditableNomination.id : 0);

    if (duplicateNomination) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "You already have a nomination record for this position. Use the nomination status page to review it.",
      );
      return res.redirect("/nomination/status");
    }

    if (
      !req.file &&
      (!isCorrectionResubmission || !existingEditableNomination.photoPath)
    ) {
      setFlash(req, "error", "Upload a candidate photo before submitting the nomination.");
      return res.redirect(
        isCorrectionResubmission ? `/nomination/form?edit=${existingEditableNomination.id}` : "/nomination/form",
      );
    }

    const submittedAt = nowIso();
    const nextPhotoPath = req.file
      ? normalizeAssetPath(req.file.path)
      : existingEditableNomination?.photoPath || "";

    try {
      if (isCorrectionResubmission) {
        db.prepare(`
          UPDATE nominations
          SET
            position_id = ?,
            full_name = ?,
            phone_number = ?,
            department = ?,
            photo_path = ?,
            bio = ?,
            manifesto = ?,
            proposer_name = ?,
            seconder_name = ?,
            declaration_accepted = 1,
            status = 'pending',
            reviewed_at = NULL,
            reviewed_by = '',
            submitted_at = ?,
            updated_at = ?
          WHERE id = ?
        `).run(
          positionId,
          fullName,
          applicant.phoneNumber,
          department,
          nextPhotoPath,
          bio,
          manifesto,
          proposerName,
          seconderName,
          submittedAt,
          submittedAt,
          existingEditableNomination.id,
        );

        if (req.file && existingEditableNomination.photoPath) {
          await safeRemoveFile(resolveAssetPath(existingEditableNomination.photoPath));
        }

        logAudit(req, "nomination", applicant.staffId, "nomination_resubmitted", {
          nominationId: existingEditableNomination.id,
          positionName: position.name,
        });
      } else {
        db.prepare(`
          INSERT INTO nominations (
            voter_id,
            position_id,
            staff_id,
            full_name,
            phone_number,
            department,
            photo_path,
            bio,
            manifesto,
            proposer_name,
            seconder_name,
            declaration_accepted,
            status,
            admin_notes,
            submitted_at,
            created_at,
            updated_at
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'pending', '', ?, ?, ?)
        `).run(
          applicant.voterId,
          positionId,
          applicant.staffId,
          fullName,
          applicant.phoneNumber,
          department,
          nextPhotoPath,
          bio,
          manifesto,
          proposerName,
          seconderName,
          submittedAt,
          submittedAt,
          submittedAt,
        );

        logAudit(req, "nomination", applicant.staffId, "nomination_submitted", {
          positionName: position.name,
        });
      }

      req.session.nominationApplicant.fullName = fullName;
      req.session.nominationApplicant.department = department;

      setFlash(
        req,
        "success",
        isCorrectionResubmission
          ? "Your nomination has been resubmitted for review."
          : "Your nomination has been submitted successfully.",
      );
      return res.redirect("/nomination/status");
    } catch (error) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", `Nomination submission failed: ${error.message}`);
      return res.redirect(
        isCorrectionResubmission ? `/nomination/form?edit=${existingEditableNomination.id}` : "/nomination/form",
      );
    }
  },
);

app.get("/nomination/status", (req, res) => {
  const applicant = req.session.nominationApplicant;

  if (!applicant) {
    setFlash(req, "error", "Sign in with your staff ID and phone number to check nomination status.");
    return res.redirect("/nomination/status/login");
  }

  const nominations = getNominationsForVoter(applicant.voterId);
  return res.render("nomination-status", {
    pageTitle: "Nomination Status",
    applicant,
    nominations,
  });
});

app.post("/nomination/logout", (req, res) => {
  clearNominationSession(req);
  return res.redirect("/nomination/login");
});

app.get("/vote/login", (req, res) => {
  if (req.session.voter) {
    return res.redirect("/vote");
  }

  if (req.session.pendingVoterVerification && isOtpVerificationEnabled()) {
    return res.redirect("/vote/verify-otp");
  }

  return res.render("vote-login", {
    pageTitle: "Voter Login",
    otpEnabled: isOtpVerificationEnabled(),
  });
});

app.post("/vote/login", async (req, res) => {
  const staffId = normalizeStaffId(req.body.staffId);
  const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!staffId || !phoneNumber) {
    setFlash(req, "error", "Enter both your staff ID and phone number.");
    return res.redirect("/vote/login");
  }

  if (!electionState.isOpen) {
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const voterRecord = db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      has_voted AS hasVoted
    FROM voters
    WHERE staff_id = ?
  `).get(staffId);

  if (!voterRecord) {
    logAudit(req, "voter", staffId, "voter_login_failed", {
      reason: "staff_id_not_found",
    });
    setFlash(req, "error", "Your staff ID is not registered for this election.");
    return res.redirect("/vote/login");
  }

  if (voterRecord.phoneNumber !== phoneNumber) {
    logAudit(req, "voter", staffId, "voter_login_failed", {
      reason: "phone_number_mismatch",
    });
    setFlash(
      req,
      "error",
      "The phone number does not match the registered record for this staff ID.",
    );
    return res.redirect("/vote/login");
  }

  if (voterRecord.hasVoted) {
    logAudit(req, "voter", staffId, "voter_login_rejected", {
      reason: "already_voted",
    });
    setFlash(req, "error", "You have already voted. You cannot vote again.");
    return res.redirect("/vote/login");
  }

  if (isOtpVerificationEnabled()) {
    const smsPhoneNumber = toSmsPhoneNumber(phoneNumber);

    if (!smsPhoneNumber) {
      logAudit(req, "voter", staffId, "voter_otp_send_failed", {
        reason: "invalid_sms_phone_format",
      });
      setFlash(
        req,
        "error",
        "Your phone number is registered, but it is not in a valid SMS format for OTP delivery. Please contact the election committee.",
      );
      return res.redirect("/vote/login");
    }

    try {
      const challenge = await sendOtpChallenge(smsPhoneNumber);
      clearVoterSession(req);
      req.session.voteComplete = null;
      req.session.pendingVoterVerification = buildPendingVoterVerification(
        voterRecord,
        phoneNumber,
        smsPhoneNumber,
        challenge,
      );

      logAudit(req, "voter", staffId, "voter_otp_sent", {
        provider: otpProvider,
        phoneNumber: maskPhoneNumber(phoneNumber),
      });
      setFlash(
        req,
        "success",
        `A one-time OTP code has been sent to ${maskPhoneNumber(phoneNumber)}.`,
      );
      return res.redirect("/vote/verify-otp");
    } catch (error) {
      logAudit(req, "voter", staffId, "voter_otp_send_failed", {
        reason: "provider_error",
        message: error.message,
      });
      setFlash(req, "error", error.message);
      return res.redirect("/vote/login");
    }
  }

  beginAuthenticatedVoterSession(req, voterRecord);
  logAudit(req, "voter", staffId, "voter_login_success", {
    otpProvider: "disabled",
  });
  return res.redirect("/vote");
});

app.get("/vote/verify-otp", (req, res) => {
  if (req.session.voter) {
    return res.redirect("/vote");
  }

  const pendingVerification = req.session.pendingVoterVerification || null;
  if (!pendingVerification) {
    setFlash(req, "error", "Sign in first so we can send your OTP code.");
    return res.redirect("/vote/login");
  }

  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    clearVoterSession(req);
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const resendAvailableAt = pendingVerification.resendAvailableAt
    ? dayjs(pendingVerification.resendAvailableAt)
    : null;
  const expiresAt = pendingVerification.expiresAt ? dayjs(pendingVerification.expiresAt) : null;

  return res.render("vote-verify-otp", {
    pageTitle: "Verify OTP",
    maskedPhoneNumber: pendingVerification.maskedPhoneNumber,
    otpTtlMinutes,
    resendAvailableAtMs: resendAvailableAt?.isValid() ? resendAvailableAt.valueOf() : null,
    expiresAtMs: expiresAt?.isValid() ? expiresAt.valueOf() : null,
    isExpired: isPendingOtpExpired(pendingVerification),
    devOtpPreview:
      pendingVerification.provider === "dev" && !isProduction
        ? pendingVerification.devCodePreview
        : "",
  });
});

app.post("/vote/verify-otp", async (req, res) => {
  const pendingVerification = req.session.pendingVoterVerification || null;
  const code = String(req.body.code || "").replace(/\s+/g, "");

  if (!pendingVerification) {
    setFlash(req, "error", "Sign in first so we can send your OTP code.");
    return res.redirect("/vote/login");
  }

  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    clearVoterSession(req);
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  if (!code) {
    setFlash(req, "error", "Enter the OTP code that was sent to your phone.");
    return res.redirect("/vote/verify-otp");
  }

  if (isPendingOtpExpired(pendingVerification)) {
    setFlash(req, "error", "This OTP has expired. Request a new code and try again.");
    return res.redirect("/vote/verify-otp");
  }

  const voterRecord = db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      has_voted AS hasVoted
    FROM voters
    WHERE id = ?
  `).get(pendingVerification.voterId);

  if (!voterRecord) {
    clearVoterSession(req);
    setFlash(req, "error", "Your voter record is no longer available. Please sign in again.");
    return res.redirect("/vote/login");
  }

  if (voterRecord.phoneNumber !== pendingVerification.phoneNumber) {
    clearVoterSession(req);
    setFlash(
      req,
      "error",
      "Your registered phone number was updated. Please sign in again to receive a new OTP.",
    );
    return res.redirect("/vote/login");
  }

  if (voterRecord.hasVoted) {
    clearVoterSession(req);
    logAudit(req, "voter", voterRecord.staffId, "voter_login_rejected", {
      reason: "already_voted",
    });
    setFlash(req, "error", "You have already voted. You cannot vote again.");
    return res.redirect("/vote/login");
  }

  try {
    const verification = await verifyOtpChallenge(pendingVerification, code);

    if (!verification.approved) {
      pendingVerification.attempts = parseInteger(pendingVerification.attempts, 0) + 1;
      req.session.pendingVoterVerification = pendingVerification;
      logAudit(req, "voter", voterRecord.staffId, "voter_otp_failed", {
        attempts: pendingVerification.attempts,
      });
      setFlash(req, "error", verification.errorMessage);
      return res.redirect("/vote/verify-otp");
    }
  } catch (error) {
    setFlash(req, "error", error.message);
    return res.redirect("/vote/verify-otp");
  }

  beginAuthenticatedVoterSession(req, voterRecord);
  logAudit(req, "voter", voterRecord.staffId, "voter_otp_verified", {
    provider: pendingVerification.provider || otpProvider,
  });
  logAudit(req, "voter", voterRecord.staffId, "voter_login_success", {
    otpProvider: pendingVerification.provider || otpProvider,
  });
  return res.redirect("/vote");
});

app.post("/vote/resend-otp", async (req, res) => {
  const pendingVerification = req.session.pendingVoterVerification || null;

  if (!pendingVerification) {
    setFlash(req, "error", "Sign in first so we can send your OTP code.");
    return res.redirect("/vote/login");
  }

  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    clearVoterSession(req);
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const resendAvailableAt = pendingVerification.resendAvailableAt
    ? dayjs(pendingVerification.resendAvailableAt)
    : null;
  if (resendAvailableAt?.isValid() && dayjs().isBefore(resendAvailableAt)) {
    const secondsRemaining = Math.max(resendAvailableAt.diff(dayjs(), "second"), 1);
    setFlash(req, "error", `Please wait ${secondsRemaining} more seconds before requesting a new OTP.`);
    return res.redirect("/vote/verify-otp");
  }

  const voterRecord = db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      has_voted AS hasVoted
    FROM voters
    WHERE id = ?
  `).get(pendingVerification.voterId);

  if (!voterRecord) {
    clearVoterSession(req);
    setFlash(req, "error", "Your voter record is no longer available. Please sign in again.");
    return res.redirect("/vote/login");
  }

  if (voterRecord.hasVoted) {
    clearVoterSession(req);
    setFlash(req, "error", "You have already voted. You cannot vote again.");
    return res.redirect("/vote/login");
  }

  const smsPhoneNumber = toSmsPhoneNumber(voterRecord.phoneNumber);
  if (!smsPhoneNumber) {
    clearVoterSession(req);
    setFlash(
      req,
      "error",
      "Your phone number is registered, but it is not in a valid SMS format for OTP delivery. Please contact the election committee.",
    );
    return res.redirect("/vote/login");
  }

  try {
    const challenge = await sendOtpChallenge(smsPhoneNumber);
    req.session.pendingVoterVerification = buildPendingVoterVerification(
      voterRecord,
      voterRecord.phoneNumber,
      smsPhoneNumber,
      challenge,
    );

    logAudit(req, "voter", voterRecord.staffId, "voter_otp_resent", {
      provider: otpProvider,
      phoneNumber: maskPhoneNumber(voterRecord.phoneNumber),
    });
    setFlash(
      req,
      "success",
      `A new OTP code has been sent to ${maskPhoneNumber(voterRecord.phoneNumber)}.`,
    );
  } catch (error) {
    logAudit(req, "voter", voterRecord.staffId, "voter_otp_send_failed", {
      reason: "resend_provider_error",
      message: error.message,
    });
    setFlash(req, "error", error.message);
  }

  return res.redirect("/vote/verify-otp");
});

app.post("/vote/cancel-otp", (req, res) => {
  req.session.pendingVoterVerification = null;
  req.session.pendingBallot = null;
  req.session.ballotSelections = null;
  res.redirect("/vote/login");
});

app.get("/vote", requireVoter, (req, res) => {
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const voterRecord = db.prepare(`
    SELECT has_voted AS hasVoted
    FROM voters
    WHERE id = ?
  `).get(req.session.voter.voterId);

  if (!voterRecord || voterRecord.hasVoted) {
    clearVoterSession(req);
    setFlash(req, "error", "You have already voted. You cannot vote again.");
    return res.redirect("/vote/login");
  }

  const ballotData = getBallotData();
  if (ballotData.length === 0) {
    setFlash(
      req,
      "error",
      "No voting positions are available yet. Please contact the election committee.",
    );
    return res.redirect("/vote/login");
  }

  const selectionMap = getBallotSelectionMap(req);
  const { incompleteSteps } = buildSelectionsFromMap(ballotData, selectionMap);
  if (incompleteSteps.length === 0) {
    return res.redirect("/vote/confirm");
  }

  return res.redirect(`/vote/step/${incompleteSteps[0].stepNumber}`);
});

app.get("/vote/step/:stepNumber", requireVoter, (req, res) => {
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const voterRecord = db.prepare(`
    SELECT has_voted AS hasVoted
    FROM voters
    WHERE id = ?
  `).get(req.session.voter.voterId);

  if (!voterRecord || voterRecord.hasVoted) {
    clearVoterSession(req);
    setFlash(req, "error", "You have already voted. You cannot vote again.");
    return res.redirect("/vote/login");
  }

  const ballotData = getBallotData();

  if (ballotData.length === 0) {
    setFlash(
      req,
      "error",
      "No voting positions are available yet. Please contact the election committee.",
    );
    return res.redirect("/vote/login");
  }

  const currentStep = Math.min(
    Math.max(parseInteger(req.params.stepNumber, 1), 1),
    ballotData.length,
  );
  const currentPosition = ballotData[currentStep - 1];
  const selectionMap = getBallotSelectionMap(req);
  const savedSelection =
    selectionMap[String(currentPosition.id)] || selectionMap[currentPosition.id] || null;
  const isSkippedSelection = Boolean(
    savedSelection?.isSkipped || savedSelection?.skipped || savedSelection?.abstained,
  );
  const selectedCandidateId = parseInteger(savedSelection?.candidateId ?? savedSelection, 0);
  const { selections } = buildSelectionsFromMap(ballotData, selectionMap);

  return res.render("vote-ballot", {
    pageTitle: "Cast Your Vote",
    currentPosition,
    currentStep,
    totalSteps: ballotData.length,
    completedSteps: selections.length,
    selectedCandidateId,
    isSkippedSelection,
    isLastStep: currentStep === ballotData.length,
    progressPercent: (currentStep / ballotData.length) * 100,
  });
});

app.post("/vote/step/:stepNumber", requireVoter, (req, res) => {
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const ballotData = getBallotData();

  if (ballotData.length === 0) {
    setFlash(
      req,
      "error",
      "No voting positions are available yet. Please contact the election committee.",
    );
    return res.redirect("/vote/login");
  }

  const currentStep = Math.min(
    Math.max(parseInteger(req.params.stepNumber, 1), 1),
    ballotData.length,
  );
  const currentPosition = ballotData[currentStep - 1];
  const action = String(req.body.action || "select").trim().toLowerCase();
  const selectionMap = getBallotSelectionMap(req);

  if (action === "skip") {
    selectionMap[String(currentPosition.id)] = {
      positionId: currentPosition.id,
      positionName: currentPosition.name,
      candidateId: null,
      candidateName: "",
      isSkipped: true,
    };
    req.session.ballotSelections = selectionMap;
    req.session.pendingBallot = null;

    if (currentStep >= ballotData.length) {
      return res.redirect("/vote/confirm");
    }

    return res.redirect(`/vote/step/${currentStep + 1}`);
  }

  const candidateId = parseInteger(req.body.candidateId, 0);
  const selectedCandidate = currentPosition.candidates.find(
    (candidate) => candidate.id === candidateId,
  );

  if (!selectedCandidate) {
    setFlash(
      req,
      "error",
      `Choose a candidate for ${currentPosition.name} or use Skip Position.`,
    );
    return res.redirect(`/vote/step/${currentStep}`);
  }

  selectionMap[String(currentPosition.id)] = {
    positionId: currentPosition.id,
    positionName: currentPosition.name,
    candidateId: selectedCandidate.id,
    candidateName: selectedCandidate.name,
    isSkipped: false,
  };
  req.session.ballotSelections = selectionMap;
  req.session.pendingBallot = null;

  if (currentStep >= ballotData.length) {
    return res.redirect("/vote/confirm");
  }

  return res.redirect(`/vote/step/${currentStep + 1}`);
});

app.get("/vote/confirm", requireVoter, (req, res) => {
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const ballotData = getBallotData();
  const { selections, incompleteSteps } = buildSelectionsFromMap(
    ballotData,
    getBallotSelectionMap(req),
  );

  if (incompleteSteps.length > 0) {
    setFlash(
      req,
      "error",
      `Choose a candidate for ${incompleteSteps[0].positionName} or skip that position.`,
    );
    return res.redirect(`/vote/step/${incompleteSteps[0].stepNumber}`);
  }

  req.session.pendingBallot = selections;
  return res.render("vote-confirm", {
    pageTitle: "Confirm Your Vote",
    selections,
    lastStep: ballotData.length,
  });
});

app.post("/vote/submit", requireVoter, (req, res) => {
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  if (!electionState.isOpen) {
    setFlash(req, "error", electionState.message);
    return res.redirect("/vote/login");
  }

  const ballotData = getBallotData();
  const { selections, incompleteSteps } = buildSelectionsFromMap(
    ballotData,
    getBallotSelectionMap(req),
  );

  if (incompleteSteps.length > 0 || ballotData.length === 0) {
    setFlash(
      req,
      "error",
      "Your ballot confirmation expired. Please continue from the next uncompleted position.",
    );
    return res.redirect(
      incompleteSteps.length > 0 ? `/vote/step/${incompleteSteps[0].stepNumber}` : "/vote",
    );
  }

  try {
    runTransaction(() => {
      const voterRecord = db.prepare(`
        SELECT
          id,
          staff_id AS staffId,
          has_voted AS hasVoted
        FROM voters
        WHERE id = ?
      `).get(req.session.voter.voterId);

      if (!voterRecord) {
        throw new Error("Your voter record is no longer available.");
      }

      if (voterRecord.hasVoted) {
        throw new Error("You have already voted. You cannot vote again.");
      }

      for (const selection of selections) {
        if (selection.isSkipped || !selection.candidateId) {
          continue;
        }

        const candidateRecord = db.prepare(`
          SELECT id
          FROM candidates
          WHERE id = ?
            AND position_id = ?
            AND is_active = 1
        `).get(selection.candidateId, selection.positionId);

        if (!candidateRecord) {
          throw new Error(
            "One of your selected candidates is no longer available. Please review your ballot again.",
          );
        }
      }

      const submittedAt = nowIso();
      const ballotInsert = db.prepare(`
        INSERT INTO ballots (
          voter_id,
          submitted_at,
          ip_address,
          user_agent,
          created_at
        )
        VALUES (?, ?, ?, ?, ?)
      `).run(
        voterRecord.id,
        submittedAt,
        req.ip || "",
        req.get("user-agent") || "",
        submittedAt,
      );

      for (const selection of selections) {
        if (selection.isSkipped || !selection.candidateId) {
          continue;
        }

        db.prepare(`
          INSERT INTO ballot_entries (
            ballot_id,
            position_id,
            candidate_id,
            created_at
          )
          VALUES (?, ?, ?, ?)
        `).run(
          ballotInsert.lastInsertRowid,
          selection.positionId,
          selection.candidateId,
          submittedAt,
        );
      }

      db.prepare(`
        UPDATE voters
        SET
          has_voted = 1,
          voted_at = ?,
          updated_at = ?
        WHERE id = ?
      `).run(submittedAt, submittedAt, voterRecord.id);
    });
  } catch (error) {
    setFlash(req, "error", error.message);
    return res.redirect("/vote");
  }

  const skippedCount = selections.filter((selection) => selection.isSkipped).length;
  const submittedChoices = selections.length - skippedCount;

  logAudit(req, "voter", req.session.voter.staffId, "vote_submitted", {
    positionsReviewed: selections.length,
    submittedChoices,
    skippedCount,
  });

  req.session.voteComplete = {
    voterName: req.session.voter.fullName,
    submittedAt: nowIso(),
  };
  clearVoterSession(req);

  return res.redirect("/vote/complete");
});

app.get("/vote/complete", (req, res) => {
  if (!req.session.voteComplete) {
    return res.redirect("/vote/login");
  }

  return res.render("vote-complete", {
    pageTitle: "Vote Submitted",
    receipt: req.session.voteComplete,
  });
});

app.post("/vote/logout", (req, res) => {
  clearVoterSession(req);
  req.session.voteComplete = null;
  res.redirect("/vote/login");
});

app.get("/admin/login", (req, res) => {
  if (req.session.admin) {
    return res.redirect("/admin");
  }

  return res.render("admin-login", { pageTitle: "Admin Login" });
});

app.post("/admin/login", (req, res) => {
  const username = String(req.body.username || "").trim();
  const password = String(req.body.password || "");

  if (username !== adminUsername || !bcrypt.compareSync(password, adminPasswordHash)) {
    logAudit(req, "admin", username || "unknown", "admin_login_failed");
    setFlash(req, "error", "Invalid administrator username or password.");
    return res.redirect("/admin/login");
  }

  req.session.admin = { username };
  logAudit(req, "admin", username, "admin_login_success");
  return res.redirect("/admin");
});

app.post("/admin/logout", requireAdmin, (req, res) => {
  logAudit(req, "admin", req.session.admin.username, "admin_logout");
  req.session.admin = null;
  res.redirect("/admin/login");
});

app.get("/admin", requireAdmin, (req, res) => {
  const metrics = getDashboardMetrics();
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const archives = getElectionArchives().slice(0, 5);
  const turnoutRate = metrics.totalVoters ? metrics.votedCount / metrics.totalVoters : 0;
  const notVotedCount = Math.max(metrics.totalVoters - metrics.votedCount, 0);
  const resultsPreview = getResultsSummary();
  const activityFeed = getDashboardActivity(5);
  const topCandidates = getDashboardTopCandidates(resultsPreview);
  const positionBars = getDashboardPositionBars(resultsPreview);
  const timelineSeries = getDashboardTimeline(settings);
  const timelineChart = buildLineChartShape(timelineSeries);
  const positionReadiness = db.prepare(`
    SELECT
      p.name AS positionName,
      COUNT(c.id) AS candidateCount
    FROM positions p
    LEFT JOIN candidates c
      ON c.position_id = p.id
      AND c.is_active = 1
    WHERE p.is_active = 1
    GROUP BY p.id
    ORDER BY p.sort_order ASC, p.name ASC
  `).all();

  return res.render("admin-dashboard", {
    pageTitle: "Admin Dashboard",
    metrics,
    settings,
    electionState,
    positionReadiness,
    archives,
    turnoutRate,
    notVotedCount,
    activityFeed,
    topCandidates,
    positionBars,
    timelineSeries,
    timelineChart,
    themeOptions: getThemeOptions(),
  });
});

app.get("/admin/nominations", requireAdmin, (req, res) => {
  const settings = getElectionSettings();
  const nominationState = computeNominationState(settings);
  const nominations = getNominationList();
  const nominationMetrics = getNominationMetrics();

  return res.render("admin-nominations", {
    pageTitle: "Nominations",
    nominations,
    nominationMetrics,
    nominationState,
    positions: getPositions(),
  });
});

app.get("/admin/nominations/pdf", requireAdmin, (req, res) => {
  const settings = getElectionSettings();
  const nominations = getNominationList();
  const filename = `${toSafeFilename(settings.electionName)}-nominations.pdf`;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: "A4",
    margin: 50,
    info: {
      Title: `${settings.electionName} Nomination Applications`,
      Author: "Organization Vote Portal",
    },
  });

  document.pipe(res);
  renderNominationReportPdf(document, {
    settings,
    nominations,
    generatedAt: nowIso(),
  });
  document.end();
});

app.get("/admin/nominations/:id", requireAdmin, (req, res) => {
  const nomination = getNominationById(parseInteger(req.params.id, 0));

  if (!nomination) {
    setFlash(req, "error", "Nomination application not found.");
    return res.redirect("/admin/nominations");
  }

  return res.render("admin-nomination-detail", {
    pageTitle: "Nomination Review",
    nomination,
    positions: getPositions(),
  });
});

app.post("/admin/nominations/settings", requireAdmin, (req, res) => {
  const opensAt = String(req.body.nominationOpensAt || "").trim();
  const closesAt = String(req.body.nominationClosesAt || "").trim();

  if (!opensAt || !closesAt) {
    setFlash(req, "error", "Set both the nomination opening and closing time.");
    return res.redirect("/admin/nominations");
  }

  if (!dayjs(opensAt).isBefore(dayjs(closesAt))) {
    setFlash(req, "error", "The nomination closing time must be later than the opening time.");
    return res.redirect("/admin/nominations");
  }

  setSetting("nomination_opens_at", opensAt);
  setSetting("nomination_closes_at", closesAt);

  const nextSettings = getElectionSettings();
  const readiness = getNominationReadiness(nextSettings);
  const opensAtValue = dayjs(nextSettings.nominationOpensAt);
  const closesAtValue = dayjs(nextSettings.nominationClosesAt);
  const shouldOpenImmediately =
    readiness.isReady &&
    opensAtValue.isValid() &&
    closesAtValue.isValid() &&
    !dayjs().isBefore(opensAtValue) &&
    dayjs().isBefore(closesAtValue);

  setSetting("nomination_phase", shouldOpenImmediately ? "open" : "setup");

  logAudit(req, "admin", req.session.admin.username, "nomination_settings_updated", {
    opensAt,
    closesAt,
  });

  setFlash(
    req,
    "success",
    shouldOpenImmediately
      ? "Nomination settings saved and nominations are now open."
      : "Nomination settings updated.",
  );
  return res.redirect("/admin/nominations");
});

app.post("/admin/nominations/open", requireAdmin, (req, res) => {
  const settings = getElectionSettings();
  const readiness = getNominationReadiness(settings);

  if (!readiness.isReady) {
    setFlash(req, "error", readiness.issues[0]);
    return res.redirect("/admin/nominations");
  }

  setSetting("nomination_phase", "open");
  logAudit(req, "admin", req.session.admin.username, "nominations_opened", {
    opensAt: settings.nominationOpensAt,
    closesAt: settings.nominationClosesAt,
  });
  setFlash(req, "success", "Nominations have been opened.");
  return res.redirect("/admin/nominations");
});

app.post("/admin/nominations/close", requireAdmin, (req, res) => {
  const settings = getElectionSettings();

  if (settings.nominationPhase !== "open") {
    setFlash(req, "error", "Nominations are not currently open.");
    return res.redirect("/admin/nominations");
  }

  setSetting("nomination_phase", "closed");
  logAudit(req, "admin", req.session.admin.username, "nominations_closed");
  setFlash(req, "success", "Nominations have been closed.");
  return res.redirect("/admin/nominations");
});

app.post(
  "/admin/nominations/:id/update",
  requireAdmin,
  nominationUpload.single("photo"),
  async (req, res) => {
    const nominationId = parseInteger(req.params.id, 0);
    const nomination = getNominationById(nominationId);

    if (!nomination) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Nomination application not found.");
      return res.redirect("/admin/nominations");
    }

    if (nomination.publishedCandidateId) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Published nominations can no longer be edited.");
      return res.redirect(`/admin/nominations/${nominationId}`);
    }

    const fullName = String(req.body.fullName || "").trim();
    const department = String(req.body.department || "").trim();
    const positionId = parseInteger(req.body.positionId, 0);
    const bio = String(req.body.bio || "").trim();
    const manifesto = String(req.body.manifesto || "").trim();
    const proposerName = String(req.body.proposerName || "").trim();
    const seconderName = String(req.body.seconderName || "").trim();

    if (!fullName || !department || !positionId || !bio || !manifesto || !proposerName || !seconderName) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Complete the nomination details before saving.");
      return res.redirect(`/admin/nominations/${nominationId}`);
    }

    const position = db.prepare(`
      SELECT id, name
      FROM positions
      WHERE id = ?
        AND is_active = 1
    `).get(positionId);

    if (!position) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Choose a valid position before saving.");
      return res.redirect(`/admin/nominations/${nominationId}`);
    }

    const duplicateNomination = db.prepare(`
      SELECT id
      FROM nominations
      WHERE voter_id = ?
        AND position_id = ?
        AND id <> ?
    `).get(nomination.voterId, positionId, nominationId);

    if (duplicateNomination) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "This staff member already has another nomination for the selected position.",
      );
      return res.redirect(`/admin/nominations/${nominationId}`);
    }

    const nextPhotoPath = req.file ? normalizeAssetPath(req.file.path) : nomination.photoPath;

    db.prepare(`
      UPDATE nominations
      SET
        position_id = ?,
        full_name = ?,
        department = ?,
        photo_path = ?,
        bio = ?,
        manifesto = ?,
        proposer_name = ?,
        seconder_name = ?,
        updated_at = ?
      WHERE id = ?
    `).run(
      positionId,
      fullName,
      department,
      nextPhotoPath,
      bio,
      manifesto,
      proposerName,
      seconderName,
      nowIso(),
      nominationId,
    );

    if (req.file && nomination.photoPath) {
      await safeRemoveFile(resolveAssetPath(nomination.photoPath));
    }

    logAudit(req, "admin", req.session.admin.username, "nomination_updated_by_admin", {
      nominationId,
      staffId: nomination.staffId,
      positionName: position.name,
    });
    setFlash(req, "success", "Nomination details updated.");
    return res.redirect(`/admin/nominations/${nominationId}`);
  },
);

app.post("/admin/nominations/:id/review", requireAdmin, (req, res) => {
  const nominationId = parseInteger(req.params.id, 0);
  const nomination = getNominationById(nominationId);

  if (!nomination) {
    setFlash(req, "error", "Nomination application not found.");
    return res.redirect("/admin/nominations");
  }

  if (nomination.publishedCandidateId) {
    setFlash(req, "error", "This nomination has already been published as a candidate.");
    return res.redirect(`/admin/nominations/${nominationId}`);
  }

  const nextStatus = String(req.body.status || "").trim().toLowerCase();
  const adminNotes = String(req.body.adminNotes || "").trim();
  const allowedStatuses = new Set(["approved", "rejected", "correction_requested"]);

  if (!allowedStatuses.has(nextStatus)) {
    setFlash(req, "error", "Choose a valid review decision.");
    return res.redirect(`/admin/nominations/${nominationId}`);
  }

  if ((nextStatus === "rejected" || nextStatus === "correction_requested") && !adminNotes) {
    setFlash(req, "error", "Add review notes when rejecting a nomination or requesting correction.");
    return res.redirect(`/admin/nominations/${nominationId}`);
  }

  db.prepare(`
    UPDATE nominations
    SET
      status = ?,
      admin_notes = ?,
      reviewed_at = ?,
      reviewed_by = ?,
      updated_at = ?
    WHERE id = ?
  `).run(nextStatus, adminNotes, nowIso(), req.session.admin.username, nowIso(), nominationId);

  logAudit(req, "admin", req.session.admin.username, "nomination_reviewed", {
    nominationId,
    staffId: nomination.staffId,
    status: nextStatus,
  });
  setFlash(req, "success", `Nomination marked as ${getNominationStatusMeta(nextStatus).label}.`);
  return res.redirect(`/admin/nominations/${nominationId}`);
});

app.post("/admin/nominations/:id/publish", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const nominationId = parseInteger(req.params.id, 0);
  const nomination = getNominationById(nominationId);

  if (!nomination) {
    setFlash(req, "error", "Nomination application not found.");
    return res.redirect("/admin/nominations");
  }

  if (nomination.statusMeta.value !== "approved") {
    setFlash(req, "error", "Only approved nominations can be published as candidates.");
    return res.redirect(`/admin/nominations/${nominationId}`);
  }

  if (nomination.publishedCandidateId) {
    setFlash(req, "error", "This nomination has already been published as a candidate.");
    return res.redirect(`/admin/nominations/${nominationId}`);
  }

  const timestamp = nowIso();
  const candidateBio = nomination.manifesto
    ? `${nomination.bio}\n\nManifesto:\n${nomination.manifesto}`
    : nomination.bio;

  try {
    const candidateInsert = db.prepare(`
      INSERT INTO candidates (
        position_id,
        name,
        photo_path,
        bio,
        sort_order,
        is_active,
        created_at,
        updated_at
      )
      VALUES (?, ?, ?, ?, 0, 1, ?, ?)
    `).run(
      nomination.positionId,
      nomination.fullName,
      nomination.photoPath,
      candidateBio,
      timestamp,
      timestamp,
    );

    db.prepare(`
      UPDATE nominations
      SET
        status = 'published',
        published_candidate_id = ?,
        reviewed_at = ?,
        reviewed_by = ?,
        updated_at = ?
      WHERE id = ?
    `).run(
      Number(candidateInsert.lastInsertRowid),
      timestamp,
      req.session.admin.username,
      timestamp,
      nominationId,
    );

    logAudit(req, "admin", req.session.admin.username, "nomination_published_as_candidate", {
      nominationId,
      staffId: nomination.staffId,
      candidateId: Number(candidateInsert.lastInsertRowid),
      positionName: nomination.positionName,
    });
    setFlash(req, "success", "Approved nomination has been published to the candidate list.");
  } catch (error) {
    setFlash(
      req,
      "error",
      `The nomination could not be published as a candidate: ${error.message}`,
    );
  }

  return res.redirect(`/admin/nominations/${nominationId}`);
});

app.post("/admin/settings", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const electionName = String(req.body.electionName || "").trim();
  const opensAt = String(req.body.opensAt || "").trim();
  const closesAt = String(req.body.closesAt || "").trim();

  if (!electionName) {
    setFlash(req, "error", "Enter an election name before saving settings.");
    return res.redirect("/admin");
  }

  if (opensAt && closesAt && !dayjs(opensAt).isBefore(dayjs(closesAt))) {
    setFlash(req, "error", "The closing time must be later than the opening time.");
    return res.redirect("/admin");
  }

  setSetting("election_name", electionName);
  setSetting("opens_at", opensAt);
  setSetting("closes_at", closesAt);

  const nextSettings = getElectionSettings();
  const readiness = getElectionReadiness(nextSettings);
  const opensAtValue = nextSettings.opensAt ? dayjs(nextSettings.opensAt) : null;
  const closesAtValue = nextSettings.closesAt ? dayjs(nextSettings.closesAt) : null;
  const shouldOpenImmediately =
    readiness.isReady &&
    opensAtValue?.isValid() &&
    closesAtValue?.isValid() &&
    !dayjs().isBefore(opensAtValue) &&
    dayjs().isBefore(closesAtValue);

  if (shouldOpenImmediately) {
    setSetting("election_phase", "open");
    logAudit(req, "admin", req.session.admin.username, "election_auto_open_triggered", {
      opensAt: nextSettings.opensAt,
      closesAt: nextSettings.closesAt,
    });
  }

  logAudit(req, "admin", req.session.admin.username, "election_settings_updated", {
    electionName,
    opensAt,
    closesAt,
  });

  if (shouldOpenImmediately) {
    setFlash(req, "success", "Election settings saved and voting is now open automatically.");
  } else if (readiness.isReady && opensAtValue?.isValid() && dayjs().isBefore(opensAtValue)) {
    setFlash(
      req,
      "success",
      `Election settings saved. Voting will open automatically on ${formatDateTime(
        nextSettings.opensAt,
      )}.`,
    );
  } else {
    setFlash(req, "success", "Election settings updated.");
  }

  return res.redirect("/admin");
});

app.post(
  "/admin/declaration",
  requireAdmin,
  declarationUpload.fields([
    { name: "chairmanSignature", maxCount: 1 },
    { name: "secretarySignature", maxCount: 1 },
  ]),
  async (req, res) => {
    const declarationTitle = String(req.body.declarationTitle || "").trim();
    const committeeName = String(req.body.committeeName || "").trim();
    const chairmanName = String(req.body.chairmanName || "").trim();
    const secretaryName = String(req.body.secretaryName || "").trim();
    const chairmanSignatureFile = req.files?.chairmanSignature?.[0] || null;
    const secretarySignatureFile = req.files?.secretarySignature?.[0] || null;

    if (!declarationTitle || !committeeName) {
      await safeRemoveUploadedRequestFiles(req.files);
      setFlash(req, "error", "Enter the declaration title and committee name before saving.");
      return res.redirect("/admin");
    }

    const currentSettings = getElectionSettings();
    const nextChairmanSignaturePath = chairmanSignatureFile
      ? normalizeAssetPath(chairmanSignatureFile.path)
      : currentSettings.chairmanSignaturePath;
    const nextSecretarySignaturePath = secretarySignatureFile
      ? normalizeAssetPath(secretarySignatureFile.path)
      : currentSettings.secretarySignaturePath;

    try {
      runTransaction(() => {
        setSetting("declaration_title", declarationTitle);
        setSetting("committee_name", committeeName);
        setSetting("chairman_name", chairmanName);
        setSetting("secretary_name", secretaryName);
        setSetting("chairman_signature_path", nextChairmanSignaturePath);
        setSetting("secretary_signature_path", nextSecretarySignaturePath);
      });

      if (
        chairmanSignatureFile &&
        currentSettings.chairmanSignaturePath &&
        currentSettings.chairmanSignaturePath !== nextChairmanSignaturePath
      ) {
        await safeRemoveFile(resolveAssetPath(currentSettings.chairmanSignaturePath));
      }

      if (
        secretarySignatureFile &&
        currentSettings.secretarySignaturePath &&
        currentSettings.secretarySignaturePath !== nextSecretarySignaturePath
      ) {
        await safeRemoveFile(resolveAssetPath(currentSettings.secretarySignaturePath));
      }

      logAudit(req, "admin", req.session.admin.username, "declaration_settings_updated", {
        declarationTitle,
        committeeName,
        chairmanName,
        secretaryName,
        updatedChairmanSignature: Boolean(chairmanSignatureFile),
        updatedSecretarySignature: Boolean(secretarySignatureFile),
      });

      setFlash(req, "success", "Declaration settings updated.");
    } catch (error) {
      await safeRemoveUploadedRequestFiles(req.files);
      setFlash(req, "error", `Declaration update failed: ${error.message}`);
    }

    return res.redirect("/admin");
  },
);

app.post("/admin/theme", requireAdmin, (req, res) => {
  const themeName = String(req.body.themeName || "").trim();
  const selectedTheme = getThemeOptions().find((theme) => theme.value === themeName);

  if (!selectedTheme) {
    setFlash(req, "error", "Choose one of the available software themes.");
    return res.redirect("/admin");
  }

  setSetting("theme_name", selectedTheme.value);
  logAudit(req, "admin", req.session.admin.username, "theme_updated", {
    themeName: selectedTheme.value,
  });
  setFlash(req, "success", `${selectedTheme.label} has been applied across the portal.`);
  return res.redirect("/admin");
});

app.post(
  "/admin/logo",
  requireAdmin,
  brandingUpload.single("logo"),
  async (req, res) => {
    if (!req.file) {
      setFlash(req, "error", "Choose a logo image to upload.");
      return res.redirect("/admin");
    }

    const nextLogoPath = normalizeAssetPath(req.file.path);
    const previousLogoPath = getElectionSettings().organizationLogoPath;

    try {
      setSetting("organization_logo_path", nextLogoPath);

      if (previousLogoPath && previousLogoPath !== nextLogoPath) {
        await safeRemoveFile(resolveAssetPath(previousLogoPath));
      }

      logAudit(req, "admin", req.session.admin.username, "organization_logo_updated", {
        logoPath: nextLogoPath,
      });

      setFlash(req, "success", "Organization logo updated.");
    } catch (error) {
      await safeRemoveFile(req.file.path);
      setFlash(req, "error", `Logo upload failed: ${error.message}`);
    }

    return res.redirect("/admin");
  },
);

app.post("/admin/election/open", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const settings = getElectionSettings();
  const readiness = getElectionReadiness(settings);

  if (!readiness.isReady) {
    setFlash(req, "error", readiness.issues[0]);
    return res.redirect("/admin");
  }

  setSetting("election_phase", "open");
  logAudit(req, "admin", req.session.admin.username, "election_opened", {
    opensAt: settings.opensAt,
    closesAt: settings.closesAt,
  });

  setFlash(req, "success", "Voting has been opened and the election setup is now locked.");
  return res.redirect("/admin");
});

app.post("/admin/election/close", requireAdmin, (req, res) => {
  const settings = getElectionSettings();

  if (settings.phase !== "open") {
    setFlash(req, "error", "Voting is not currently open.");
    return res.redirect("/admin");
  }

  setSetting("election_phase", "closed");
  logAudit(req, "admin", req.session.admin.username, "election_closed");
  setFlash(req, "success", "Voting has been closed. Results are now final.");
  return res.redirect("/admin/results");
});

app.post("/admin/election/archive-reset", requireAdmin, async (req, res) => {
  const settings = getElectionSettings();

  if (settings.phase !== "closed") {
    setFlash(req, "error", "Close the election before archiving and resetting the system.");
    return res.redirect("/admin/results");
  }

  const metrics = getDashboardMetrics();
  const results = getResultsSummary();

  if (
    metrics.totalVoters === 0 &&
    metrics.totalBallots === 0 &&
    metrics.totalPositions === 0 &&
    metrics.totalCandidates === 0
  ) {
    setFlash(req, "error", "There is no election data to archive yet.");
    return res.redirect("/admin/results");
  }

  const candidatePhotoRows = db.prepare(`
    SELECT photo_path AS photoPath
    FROM candidates
    WHERE photo_path <> ''
  `).all();
  const nominationPhotoRows = db.prepare(`
    SELECT photo_path AS photoPath
    FROM nominations
    WHERE photo_path <> ''
  `).all();

  let archiveId = 0;

  runTransaction(() => {
    const archivedAt = nowIso();
    const archiveInsert = db.prepare(`
      INSERT INTO election_archives (
        election_name,
        phase,
        opens_at,
        closes_at,
        archived_at,
        total_voters,
        votes_cast,
        turnout_rate,
        settings_json,
        metrics_json,
        results_json
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      settings.electionName,
      settings.phase,
      settings.opensAt,
      settings.closesAt,
      archivedAt,
      metrics.totalVoters,
      metrics.votedCount,
      metrics.totalVoters ? metrics.votedCount / metrics.totalVoters : 0,
      JSON.stringify(settings),
      JSON.stringify(metrics),
      JSON.stringify(results),
    );

    archiveId = Number(archiveInsert.lastInsertRowid);

    db.prepare("DELETE FROM ballot_entries").run();
    db.prepare("DELETE FROM ballots").run();
    db.prepare("DELETE FROM voters").run();
    db.prepare("DELETE FROM nominations").run();
    db.prepare("DELETE FROM candidates").run();
    db.prepare("DELETE FROM positions").run();

    setSetting("election_phase", "setup");
    setSetting("opens_at", "");
    setSetting("closes_at", "");
    setSetting("nomination_phase", "setup");
    setSetting("nomination_opens_at", "");
    setSetting("nomination_closes_at", "");
  });

  for (const photoRow of candidatePhotoRows) {
    await safeRemoveFile(resolveAssetPath(photoRow.photoPath));
  }

  for (const photoRow of nominationPhotoRows) {
    await safeRemoveFile(resolveAssetPath(photoRow.photoPath));
  }

  logAudit(req, "admin", req.session.admin.username, "election_archived_and_reset", {
    archiveId,
    electionName: settings.electionName,
    totalVoters: metrics.totalVoters,
    votesCast: metrics.votedCount,
  });

  setFlash(
    req,
    "success",
    `Election archived successfully. The system has been reset and is ready for the next election.`,
  );
  return res.redirect(`/admin/archives/${archiveId}`);
});

app.post("/admin/backup", requireAdmin, async (req, res) => {
  const backupName = `vote-portal-backup-${dayjs().format("YYYYMMDD-HHmmss")}.sqlite`;
  const backupPath = path.join(backupsDirectory, backupName);

  try {
    await fsp.copyFile(databasePath, backupPath);
    logAudit(req, "admin", req.session.admin.username, "database_backup_created", {
      backupFile: backupName,
    });
    setFlash(req, "success", `Backup created: ${backupName}`);
  } catch (error) {
    setFlash(req, "error", `Backup failed: ${error.message}`);
  }

  return res.redirect("/admin");
});

app.get("/admin/voters", requireAdmin, (req, res) => {
  const metrics = getDashboardMetrics();
  const voters = getVoters();
  res.render("admin-voters", {
    pageTitle: "Voters",
    voters,
    metrics,
    templatePath: "/templates/staff-login-template.xlsx",
  });
});

app.post("/admin/voters", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const staffId = normalizeStaffId(req.body.staffId);
  const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
  const fullName = String(req.body.fullName || "").trim();
  const department = String(req.body.department || "").trim();

  if (!staffId || !isLikelyPhoneNumber(phoneNumber)) {
    setFlash(
      req,
      "error",
      "Enter a unique staff ID and a valid phone number before saving the voter.",
    );
    return res.redirect("/admin/voters");
  }

  try {
    const timestamp = nowIso();
    db.prepare(`
      INSERT INTO voters (
        staff_id,
        phone_number,
        full_name,
        department,
        has_voted,
        created_at,
        updated_at
      )
      VALUES (?, ?, ?, ?, 0, ?, ?)
    `).run(staffId, phoneNumber, fullName, department, timestamp, timestamp);

    logAudit(req, "admin", req.session.admin.username, "voter_added_manually", {
      staffId,
    });
    setFlash(req, "success", `${staffId} has been added to the voter list.`);
  } catch (_error) {
    setFlash(req, "error", "That staff ID already exists or could not be added.");
  }

  return res.redirect("/admin/voters");
});

app.post(
  "/admin/voters/import",
  requireAdmin,
  voterImportUpload.single("voterFile"),
  async (req, res) => {
    if (!ensureSetupMode(req, res)) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      return;
    }

    if (!req.file) {
      setFlash(req, "error", "Choose an Excel or CSV file to import voters.");
      return res.redirect("/admin/voters");
    }

    try {
      const rows = await parseVoterWorkbook(req.file.path);

      if (rows.length === 0) {
        throw new Error("The uploaded file does not contain any voter rows.");
      }

      const seenStaffIds = new Set();
      const preparedRows = rows.map((row) => {
        const staffId = normalizeStaffId(row.staff_id);
        const phoneNumber = normalizePhoneNumber(row.phone_number);
        const fullName = String(row.full_name || "").trim();
        const department = String(row.department || "").trim();

        if (!staffId) {
          throw new Error(`Row ${row.__rowNumber}: staff_id is required.`);
        }

        if (!isLikelyPhoneNumber(phoneNumber)) {
          throw new Error(
            `Row ${row.__rowNumber}: phone_number must be a valid registered phone number.`,
          );
        }

        if (seenStaffIds.has(staffId)) {
          throw new Error(`Row ${row.__rowNumber}: duplicate staff_id ${staffId} found.`);
        }

        seenStaffIds.add(staffId);

        return {
          staffId,
          phoneNumber,
          fullName,
          department,
        };
      });

      const existingStaffIds = new Set(
        db
          .prepare("SELECT staff_id AS staffId FROM voters")
          .all()
          .map((row) => row.staffId),
      );

      let createdCount = 0;
      let updatedCount = 0;

      runTransaction(() => {
        for (const row of preparedRows) {
          const timestamp = nowIso();
          db.prepare(`
            INSERT INTO voters (
              staff_id,
              phone_number,
              full_name,
              department,
              has_voted,
              created_at,
              updated_at
            )
            VALUES (?, ?, ?, ?, 0, ?, ?)
            ON CONFLICT(staff_id) DO UPDATE SET
              phone_number = excluded.phone_number,
              full_name = excluded.full_name,
              department = excluded.department,
              updated_at = excluded.updated_at
          `).run(
            row.staffId,
            row.phoneNumber,
            row.fullName,
            row.department,
            timestamp,
            timestamp,
          );

          if (existingStaffIds.has(row.staffId)) {
            updatedCount += 1;
          } else {
            createdCount += 1;
          }
        }
      });

      logAudit(req, "admin", req.session.admin.username, "voters_imported", {
        createdCount,
        updatedCount,
      });

      setFlash(
        req,
        "success",
        `Voter import completed. Added ${createdCount} new voters and updated ${updatedCount} existing records.`,
      );
    } catch (error) {
      setFlash(req, "error", error.message);
    } finally {
      await safeRemoveFile(req.file.path);
    }

    return res.redirect("/admin/voters");
  },
);

app.get("/admin/voters/:id(\\d+)/edit", requireAdmin, (req, res) => {
  const voterId = parseInteger(req.params.id, 0);
  const voter = db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      department,
      has_voted AS hasVoted
    FROM voters
    WHERE id = ?
  `).get(voterId);

  if (!voter) {
    setFlash(req, "error", "Voter not found.");
    return res.redirect("/admin/voters");
  }

  return res.render("admin-voter-edit", {
    pageTitle: "Edit Voter",
    voter,
  });
});

app.post("/admin/voters/:id(\\d+)", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const voterId = parseInteger(req.params.id, 0);
  const existingVoter = db.prepare(`
    SELECT
      id,
      staff_id AS staffId
    FROM voters
    WHERE id = ?
  `).get(voterId);

  if (!existingVoter) {
    setFlash(req, "error", "Voter not found.");
    return res.redirect("/admin/voters");
  }

  const staffId = normalizeStaffId(req.body.staffId);
  const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
  const fullName = String(req.body.fullName || "").trim();
  const department = String(req.body.department || "").trim();

  if (!staffId || !isLikelyPhoneNumber(phoneNumber)) {
    setFlash(
      req,
      "error",
      "Enter a unique staff ID and a valid phone number before saving the voter.",
    );
    return res.redirect(`/admin/voters/${voterId}/edit`);
  }

  const duplicateVoter = db.prepare(`
    SELECT id
    FROM voters
    WHERE staff_id = ?
      AND id <> ?
  `).get(staffId, voterId);

  if (duplicateVoter) {
    setFlash(req, "error", "Another voter is already using that staff ID.");
    return res.redirect(`/admin/voters/${voterId}/edit`);
  }

  db.prepare(`
    UPDATE voters
    SET
      staff_id = ?,
      phone_number = ?,
      full_name = ?,
      department = ?,
      updated_at = ?
    WHERE id = ?
  `).run(staffId, phoneNumber, fullName, department, nowIso(), voterId);

  logAudit(req, "admin", req.session.admin.username, "voter_updated", {
    voterId,
    staffId,
  });
  setFlash(req, "success", `${staffId} has been updated.`);
  return res.redirect("/admin/voters");
});

app.post("/admin/voters/clear", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const voterCount = db.prepare("SELECT COUNT(*) AS total FROM voters").get().total;

  if (voterCount === 0) {
    setFlash(req, "error", "There are no imported voters to remove.");
    return res.redirect("/admin/voters");
  }

  const ballotCount = db.prepare("SELECT COUNT(*) AS total FROM ballots").get().total;

  if (ballotCount > 0) {
    setFlash(
      req,
      "error",
      "Imported voters cannot be cleared after ballot activity exists. Start a new election cycle before removing voter records.",
    );
    return res.redirect("/admin/voters");
  }

  runTransaction(() => {
    db.prepare("DELETE FROM voters").run();
  });

  logAudit(req, "admin", req.session.admin.username, "voters_cleared", {
    removedCount: voterCount,
  });

  setFlash(req, "success", `Removed ${voterCount} imported voter records.`);
  return res.redirect("/admin/voters");
});

app.get("/admin/setup", requireAdmin, (req, res) => {
  const positions = getPositions();
  const candidates = getCandidates();
  res.render("admin-setup", {
    pageTitle: "Election Setup",
    positions,
    candidates,
  });
});

app.get("/admin/archives", requireAdmin, (req, res) => {
  const archives = getElectionArchives();
  res.render("admin-archives", {
    pageTitle: "Election Archives",
    archives,
  });
});

app.get("/admin/archives/:id", requireAdmin, (req, res) => {
  const archiveId = parseInteger(req.params.id, 0);
  const archive = getElectionArchiveById(archiveId);

  if (!archive) {
    setFlash(req, "error", "Archived election not found.");
    return res.redirect("/admin/archives");
  }

  return res.render("admin-archive-detail", {
    pageTitle: "Archived Election",
    archive,
  });
});

app.post("/admin/archives/:id/delete", requireAdmin, (req, res) => {
  const archiveId = parseInteger(req.params.id, 0);
  const archive = db.prepare(`
    SELECT
      id,
      election_name AS electionName
    FROM election_archives
    WHERE id = ?
  `).get(archiveId);

  if (!archive) {
    setFlash(req, "error", "Archived election not found.");
    return res.redirect("/admin/archives");
  }

  db.prepare("DELETE FROM election_archives WHERE id = ?").run(archiveId);

  logAudit(req, "admin", req.session.admin.username, "archive_deleted", {
    archiveId,
    electionName: archive.electionName,
  });

  setFlash(req, "success", `Archived election "${archive.electionName}" has been deleted.`);
  return res.redirect("/admin/archives");
});

app.post("/admin/positions", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const positionName = String(req.body.positionName || "").trim();
  const sortOrder = parseInteger(req.body.sortOrder, 0);

  if (!positionName) {
    setFlash(req, "error", "Enter a position name.");
    return res.redirect("/admin/setup");
  }

  try {
    const timestamp = nowIso();
    db.prepare(`
      INSERT INTO positions (name, sort_order, is_active, created_at, updated_at)
      VALUES (?, ?, 1, ?, ?)
    `).run(positionName, sortOrder, timestamp, timestamp);

    logAudit(req, "admin", req.session.admin.username, "position_added", {
      positionName,
    });
    setFlash(req, "success", `${positionName} added to the election ballot.`);
  } catch (_error) {
    setFlash(req, "error", "That position already exists or could not be created.");
  }

  return res.redirect("/admin/setup");
});

app.post("/admin/positions/:id/delete", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const positionId = parseInteger(req.params.id, 0);
  const nominationCount = db.prepare(`
    SELECT COUNT(*) AS total
    FROM nominations
    WHERE position_id = ?
  `).get(positionId);

  if (nominationCount.total > 0) {
    setFlash(
      req,
      "error",
      "This position has nomination applications. Review or clear those nominations before deleting the position.",
    );
    return res.redirect("/admin/setup");
  }

  const candidateCount = db.prepare(`
    SELECT COUNT(*) AS total
    FROM candidates
    WHERE position_id = ?
      AND is_active = 1
  `).get(positionId);

  if (candidateCount.total > 0) {
    setFlash(
      req,
      "error",
      "Remove candidates from this position before deleting the position.",
    );
    return res.redirect("/admin/setup");
  }

  db.prepare("DELETE FROM positions WHERE id = ?").run(positionId);
  logAudit(req, "admin", req.session.admin.username, "position_deleted", {
    positionId,
  });
  setFlash(req, "success", "Position removed.");
  return res.redirect("/admin/setup");
});

app.post(
  "/admin/candidates",
  requireAdmin,
  candidateUpload.single("photo"),
  async (req, res) => {
    if (!ensureSetupMode(req, res)) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      return;
    }

    const name = String(req.body.name || "").trim();
    const bio = String(req.body.bio || "").trim();
    const sortOrder = parseInteger(req.body.sortOrder, 0);
    const positionId = parseInteger(req.body.positionId, 0);

    if (!name || !positionId) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Enter the candidate name and select a position.");
      return res.redirect("/admin/setup");
    }

    const position = db.prepare(`
      SELECT name
      FROM positions
      WHERE id = ?
        AND is_active = 1
    `).get(positionId);

    if (!position) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Choose a valid position for the candidate.");
      return res.redirect("/admin/setup");
    }

    try {
      const timestamp = nowIso();
      db.prepare(`
        INSERT INTO candidates (
          position_id,
          name,
          photo_path,
          bio,
          sort_order,
          is_active,
          created_at,
          updated_at
        )
        VALUES (?, ?, ?, ?, ?, 1, ?, ?)
      `).run(
        positionId,
        name,
        req.file ? normalizeAssetPath(req.file.path) : "",
        bio,
        sortOrder,
        timestamp,
        timestamp,
      );

      logAudit(req, "admin", req.session.admin.username, "candidate_added", {
        candidateName: name,
        positionName: position.name,
      });

      setFlash(req, "success", `${name} added under ${position.name}.`);
    } catch (_error) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "That candidate already exists for the selected position or could not be created.",
      );
    }

    return res.redirect("/admin/setup");
  },
);

app.get("/admin/candidates/:id/edit", requireAdmin, (req, res) => {
  const candidateId = parseInteger(req.params.id, 0);
  const positions = getPositions();
  const candidate = db.prepare(`
    SELECT
      id,
      name,
      position_id AS positionId,
      photo_path AS photoPath,
      bio,
      sort_order AS sortOrder
    FROM candidates
    WHERE id = ?
      AND is_active = 1
  `).get(candidateId);

  if (!candidate) {
    setFlash(req, "error", "Candidate not found.");
    return res.redirect("/admin/setup");
  }

  return res.render("admin-candidate-edit", {
    pageTitle: "Edit Candidate",
    positions,
    candidate,
  });
});

app.post(
  "/admin/candidates/:id",
  requireAdmin,
  candidateUpload.single("photo"),
  async (req, res) => {
    if (!ensureSetupMode(req, res)) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      return;
    }

    const candidateId = parseInteger(req.params.id, 0);
    const existingCandidate = db.prepare(`
      SELECT
        id,
        name,
        position_id AS positionId,
        photo_path AS photoPath
      FROM candidates
      WHERE id = ?
        AND is_active = 1
    `).get(candidateId);

    if (!existingCandidate) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Candidate not found.");
      return res.redirect("/admin/setup");
    }

    const name = String(req.body.name || "").trim();
    const bio = String(req.body.bio || "").trim();
    const sortOrder = parseInteger(req.body.sortOrder, 0);
    const positionId = parseInteger(req.body.positionId, 0);

    if (!name || !positionId) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Enter the candidate name and select a position.");
      return res.redirect(`/admin/candidates/${candidateId}/edit`);
    }

    const position = db.prepare(`
      SELECT name
      FROM positions
      WHERE id = ?
        AND is_active = 1
    `).get(positionId);

    if (!position) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", "Choose a valid position for the candidate.");
      return res.redirect(`/admin/candidates/${candidateId}/edit`);
    }

    const nextPhotoPath = req.file
      ? normalizeAssetPath(req.file.path)
      : existingCandidate.photoPath;

    try {
      const timestamp = nowIso();
      db.prepare(`
        UPDATE candidates
        SET
          position_id = ?,
          name = ?,
          photo_path = ?,
          bio = ?,
          sort_order = ?,
          updated_at = ?
        WHERE id = ?
      `).run(positionId, name, nextPhotoPath, bio, sortOrder, timestamp, candidateId);

      if (req.file && existingCandidate.photoPath) {
        await safeRemoveFile(resolveAssetPath(existingCandidate.photoPath));
      }

      logAudit(req, "admin", req.session.admin.username, "candidate_updated", {
        candidateId,
        candidateName: name,
        positionName: position.name,
      });

      setFlash(req, "success", `${name} updated successfully.`);
      return res.redirect("/admin/setup");
    } catch (_error) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "That candidate already exists for the selected position or could not be updated.",
      );
      return res.redirect(`/admin/candidates/${candidateId}/edit`);
    }
  },
);

app.post("/admin/candidates/:id/delete", requireAdmin, async (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const candidateId = parseInteger(req.params.id, 0);
  const candidate = db.prepare(`
    SELECT
      name,
      photo_path AS photoPath
    FROM candidates
    WHERE id = ?
  `).get(candidateId);

  if (!candidate) {
    setFlash(req, "error", "Candidate not found.");
    return res.redirect("/admin/setup");
  }

  db.prepare("DELETE FROM candidates WHERE id = ?").run(candidateId);

  if (candidate.photoPath) {
    await safeRemoveFile(resolveAssetPath(candidate.photoPath));
  }

  logAudit(req, "admin", req.session.admin.username, "candidate_deleted", {
    candidateId,
    candidateName: candidate.name,
  });
  setFlash(req, "success", "Candidate removed.");
  return res.redirect("/admin/setup");
});

app.get("/admin/results/print", requireAdmin, (req, res) => {
  const payload = getResultsExportPayload();

  if (!payload.electionState.isClosed) {
    setFlash(
      req,
      "error",
      "Results are available for printing only after the election has been closed.",
    );
    return res.redirect("/admin/results");
  }

  return res.render("admin-results-print", {
    pageTitle: "Print Results",
    ...payload,
  });
});

app.get("/admin/results/pdf", requireAdmin, (req, res) => {
  const payload = getResultsExportPayload();

  if (!payload.electionState.isClosed) {
    setFlash(
      req,
      "error",
      "Results are available for export only after the election has been closed.",
    );
    return res.redirect("/admin/results");
  }

  const filename = `${toSafeFilename(payload.settings.electionName)}-results.pdf`;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: "A4",
    margin: 50,
    info: {
      Title: `${payload.settings.electionName} Results`,
      Author: "Organization Vote Portal",
    },
  });

  document.pipe(res);
  renderResultsPdf(document, payload);
  document.end();
});

app.get("/admin/results", requireAdmin, (req, res) => {
  const payload = getResultsExportPayload();
  const showResults = payload.electionState.isOpen || payload.electionState.isClosed;
  const pageIntro = payload.electionState.isOpen
    ? "Monitor live provisional statistics and candidate performance while voting is in progress."
    : payload.electionState.isClosed
      ? "Review final totals after voting has closed and the ballot is locked."
      : "Results become available for monitoring once voting opens.";

  res.render("admin-results", {
    pageTitle: "Results",
    pageIntro,
    metrics: payload.metrics,
    results: showResults ? payload.results : [],
    generatedAt: payload.generatedAt,
    nonVoters: payload.nonVoters,
    resultsStatusLabel: payload.resultsStatusLabel,
    declarationLabel: payload.declarationLabel,
    resultsLocked: !showResults,
    isLiveResults: payload.electionState.isOpen,
    canArchiveReset: payload.electionState.isClosed,
  });
});

app.get("/admin/audit", requireAdmin, (req, res) => {
  const logs = getAuditLogs();
  res.render("admin-audit", {
    pageTitle: "Audit Log",
    logs,
  });
});

app.use((error, req, res, _next) => {
  console.error(error);

  const redirectTarget = req.path.startsWith("/admin")
    ? "/admin"
    : req.path.startsWith("/nomination")
      ? "/nomination/login"
      : "/";
  setFlash(req, "error", error.message || "Something went wrong.");
  res.redirect(redirectTarget);
});

async function start() {
  ensureDirectories();
  initDatabase(defaultElectionName);
  await ensureVoterTemplate(templatePath);
  await ensureVoterTemplate(staffLoginTemplatePath);

  app.listen(port, host, () => {
    console.log(`Vote portal running on http://localhost:${port}`);
  });
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
