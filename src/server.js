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
  return {
    electionName: settings.election_name || defaultElectionName,
    organizationLogoPath: settings.organization_logo_path || "",
    phase: settings.election_phase || "setup",
    opensAt: settings.opens_at || "",
    closesAt: settings.closes_at || "",
    resultsVisibility: settings.results_visibility || "after_close",
    themeName: settings.theme_name || "heritage",
  };
}

function computeElectionState(settings) {
  const now = dayjs();
  const opensAt = settings.opensAt ? dayjs(settings.opensAt) : null;
  const closesAt = settings.closesAt ? dayjs(settings.closesAt) : null;

  let status = "setup";
  let message =
    "The election is still in setup. Add voters, positions, and candidates before opening voting.";

  if (settings.phase === "open") {
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

function syncAutomaticClosure() {
  const settings = getElectionSettings();
  const state = computeElectionState(settings);

  if (settings.phase === "open" && state.status === "closed") {
    setSetting("election_phase", "closed");
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
  return {
    settings,
    electionState: computeElectionState(settings),
    metrics: getDashboardMetrics(),
    results: getResultsSummary(),
    generatedAt: nowIso(),
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
  const { settings, metrics, results, generatedAt } = payload;
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
    .text(`Final election results`, 120, 78)
    .text(`Generated: ${formatDateTime(generatedAt)}`, 120, 94);

  document
    .moveTo(50, 126)
    .lineTo(545, 126)
    .strokeColor("#d8c197")
    .lineWidth(1)
    .stroke();

  document.y = 148;
  document.font("Helvetica-Bold").fontSize(12).fillColor("#102338").text("Summary");
  document
    .moveDown(0.4)
    .font("Helvetica")
    .fontSize(11)
    .fillColor("#102338")
    .text(`Total voters: ${metrics.totalVoters}`)
    .text(`Votes cast: ${metrics.votedCount}`)
    .text(
      `Turnout: ${
        metrics.totalVoters
          ? formatPercent(metrics.votedCount / metrics.totalVoters)
          : "0.0%"
      }`,
    );

  results.forEach((result) => {
    ensurePdfSpace(document, 110 + result.candidates.length * 24);

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
      .text(`Winner: ${result.winnerLabel}`);

    const headerY = document.y + 10;
    document
      .font("Helvetica-Bold")
      .fontSize(10)
      .fillColor("#102338")
      .text("Candidate", 50, headerY, { width: 300 })
      .text("Votes", 390, headerY, { width: 60, align: "right" })
      .text("Share", 465, headerY, { width: 70, align: "right" });

    document
      .moveTo(50, headerY + 16)
      .lineTo(545, headerY + 16)
      .strokeColor("#e3d7c0")
      .lineWidth(1)
      .stroke();

    let rowY = headerY + 26;

    result.candidates.forEach((candidate) => {
      ensurePdfSpace(document, 34);

      document
        .font("Helvetica")
        .fontSize(10)
        .fillColor("#102338")
        .text(candidate.name, 50, rowY, { width: 300 })
        .text(String(candidate.voteCount), 390, rowY, { width: 60, align: "right" })
        .text(
          result.totalVotes
            ? formatPercent(candidate.voteCount / result.totalVotes)
            : "0.0%",
          465,
          rowY,
          { width: 70, align: "right" },
        );

      rowY += 20;
    });

    document.y = rowY + 4;
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
      continue;
    }

    const highestVoteCount = Math.max(
      ...summary.candidates.map((candidate) => candidate.voteCount),
    );

    const winners = summary.candidates.filter(
      (candidate) => candidate.voteCount === highestVoteCount,
    );

    summary.winnerLabel =
      highestVoteCount === 0
        ? "No votes recorded"
        : winners.length > 1
          ? `Tie: ${winners.map((candidate) => candidate.name).join(", ")}`
          : winners[0].name;

    summary.candidates = summary.candidates.map((candidate) => ({
      ...candidate,
      shareRatio: summary.totalVotes ? candidate.voteCount / summary.totalVotes : 0,
      isLeading: highestVoteCount > 0 && candidate.voteCount === highestVoteCount,
    }));
  }

  return summaries;
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

const brandingUpload = createImageUpload(
  brandingUploadsDirectory,
  "Organization logos must be image files.",
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

  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);

  res.locals.settings = settings;
  res.locals.electionState = electionState;
  res.locals.currentPath = req.path;
  res.locals.admin = req.session.admin || null;
  res.locals.voter = req.session.voter || null;
  res.locals.currentYear = new Date().getFullYear();
  res.locals.formatDateTime = formatDateTime;
  res.locals.formatPercent = formatPercent;
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

app.get("/vote/login", (req, res) => {
  if (req.session.voter) {
    return res.redirect("/vote");
  }

  return res.render("vote-login", { pageTitle: "Voter Login" });
});

app.post("/vote/login", (req, res) => {
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

  req.session.voter = {
    voterId: voterRecord.id,
    staffId: voterRecord.staffId,
    fullName: voterRecord.fullName,
  };
  req.session.ballotSelections = {};
  req.session.pendingBallot = null;

  logAudit(req, "voter", staffId, "voter_login_success");
  return res.redirect("/vote");
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
    req.session.voter = null;
    req.session.ballotSelections = null;
    req.session.pendingBallot = null;
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
    req.session.voter = null;
    req.session.ballotSelections = null;
    req.session.pendingBallot = null;
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
  req.session.voter = null;
  req.session.ballotSelections = null;
  req.session.pendingBallot = null;

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
  req.session.voter = null;
  req.session.ballotSelections = null;
  req.session.pendingBallot = null;
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
    themeOptions: getThemeOptions(),
  });
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

  logAudit(req, "admin", req.session.admin.username, "election_settings_updated", {
    electionName,
    opensAt,
    closesAt,
  });

  setFlash(req, "success", "Election settings updated.");
  return res.redirect("/admin");
});

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
  const metrics = getDashboardMetrics();
  const positionsWithoutCandidates = db.prepare(`
    SELECT p.name
    FROM positions p
    LEFT JOIN candidates c
      ON c.position_id = p.id
      AND c.is_active = 1
    WHERE p.is_active = 1
    GROUP BY p.id
    HAVING COUNT(c.id) = 0
  `).all();

  if (metrics.totalVoters === 0) {
    setFlash(req, "error", "Import at least one voter before opening the election.");
    return res.redirect("/admin");
  }

  if (metrics.totalPositions === 0 || metrics.totalCandidates === 0) {
    setFlash(req, "error", "Add positions and candidates before opening voting.");
    return res.redirect("/admin");
  }

  if (!settings.opensAt || !settings.closesAt) {
    setFlash(
      req,
      "error",
      "Set both the voting start time and closing time before opening the election.",
    );
    return res.redirect("/admin");
  }

  if (!dayjs(settings.opensAt).isBefore(dayjs(settings.closesAt))) {
    setFlash(req, "error", "The closing time must be later than the opening time.");
    return res.redirect("/admin");
  }

  if (positionsWithoutCandidates.length > 0) {
    setFlash(
      req,
      "error",
      `Every position needs at least one candidate. Missing: ${positionsWithoutCandidates
        .map((position) => position.name)
        .join(", ")}.`,
    );
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
    db.prepare("DELETE FROM candidates").run();
    db.prepare("DELETE FROM positions").run();

    setSetting("election_phase", "setup");
    setSetting("opens_at", "");
    setSetting("closes_at", "");
  });

  for (const photoRow of candidatePhotoRows) {
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

  const redirectTarget = req.path.startsWith("/admin") ? "/admin" : "/";
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
