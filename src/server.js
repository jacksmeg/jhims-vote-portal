require("dotenv").config();

const bcrypt = require("bcryptjs");
const crypto = require("node:crypto");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const dayjs = require("dayjs");
const ExcelJS = require("exceljs");
const express = require("express");
const session = require("express-session");
const helmet = require("helmet");
const multer = require("multer");
const PDFDocument = require("pdfkit");
const { PNG } = require("pngjs");
const jpeg = require("jpeg-js");
const QRCode = require("qrcode");

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
const {
  requireAdmin,
  requireObserver,
  requireObserverPasswordReady,
  requireVoter,
} = require("./middleware/auth");

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
const envTwilioAccountSid = String(process.env.TWILIO_ACCOUNT_SID || "").trim();
const envTwilioAuthToken = String(process.env.TWILIO_AUTH_TOKEN || "").trim();
const envTwilioVerifyServiceSid = String(process.env.TWILIO_VERIFY_SERVICE_SID || "").trim();
const envArkeselApiKey = String(process.env.ARKESEL_API_KEY || "").trim();
const envArkeselSenderId = String(process.env.ARKESEL_SENDER_ID || "").trim();
const defaultArkeselOtpMessageTemplate = String(
  process.env.ARKESEL_OTP_MESSAGE ||
    "Your OTP code is %otp_code%. It expires in %expiry% minutes.",
).trim();
const envConfiguredOtpProvider = String(process.env.OTP_PROVIDER || "")
  .trim()
  .toLowerCase();
const envConfiguredOtpTtlMinutes = Math.min(
  Math.max(Number.parseInt(process.env.OTP_TTL_MINUTES || "10", 10) || 10, 1),
  30,
);
const envOtpResendCooldownSeconds = Math.min(
  Math.max(
    Number.parseInt(process.env.OTP_RESEND_COOLDOWN_SECONDS || "30", 10) || 30,
    0,
  ),
  300,
);
const envCaptchaSiteKey = String(process.env.CAPTCHA_SITE_KEY || "").trim();
const envCaptchaSecretKey = String(process.env.CAPTCHA_SECRET_KEY || "").trim();
const turnstileVerifyUrl = "https://challenges.cloudflare.com/turnstile/v0/siteverify";
const devOtpCodeLength = 6;
const observerMaxLoginAttempts = 5;
const observerLockMinutes = 15;
const sessionSecureCookie = String(
  process.env.SESSION_SECURE_COOKIE || (isProduction ? "true" : "false"),
)
  .trim()
  .toLowerCase() === "true";
const adminNotificationClients = new Set();
let resultsSmsAutoSendInProgress = false;

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

function getNetworkUrls(serverHost, serverPort) {
  const urls = [];
  const normalizedHost = String(serverHost || "").trim();

  if (
    !normalizedHost ||
    normalizedHost === "0.0.0.0" ||
    normalizedHost === "::"
  ) {
    urls.push(`http://localhost:${serverPort}`);

    const interfaces = os.networkInterfaces();
    const seenAddresses = new Set();

    for (const networkInterface of Object.values(interfaces)) {
      for (const address of networkInterface || []) {
        if (
          !address ||
          address.internal ||
          address.family !== "IPv4" ||
          seenAddresses.has(address.address)
        ) {
          continue;
        }

        seenAddresses.add(address.address);
        urls.push(`http://${address.address}:${serverPort}`);
      }
    }

    return urls;
  }

  if (
    normalizedHost === "127.0.0.1" ||
    normalizedHost.toLowerCase() === "localhost"
  ) {
    return [`http://localhost:${serverPort}`];
  }

  return [`http://${normalizedHost}:${serverPort}`];
}

const totpTimeStepSeconds = 30;
const totpDigits = 6;
const totpSecretBytes = 20;
const base32Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";

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

function getAdminReturnPath(req, fallback = "/admin") {
  const returnTo = String(req.body?.returnTo || "").trim();

  if (returnTo.startsWith("/admin") && !returnTo.startsWith("//")) {
    return returnTo;
  }

  return fallback;
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

function parseJsonObject(value, fallback = {}) {
  try {
    const parsed = JSON.parse(String(value || ""));
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function humanizeToken(value) {
  return String(value || "")
    .replace(/[_-]+/g, " ")
    .replace(/\b\w/g, (character) => character.toUpperCase())
    .trim();
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

function encodeBase32(buffer) {
  let bits = 0;
  let value = 0;
  let output = "";

  for (const byte of buffer) {
    value = (value << 8) | byte;
    bits += 8;

    while (bits >= 5) {
      output += base32Alphabet[(value >>> (bits - 5)) & 31];
      bits -= 5;
    }
  }

  if (bits > 0) {
    output += base32Alphabet[(value << (5 - bits)) & 31];
  }

  return output;
}

function decodeBase32(value) {
  const normalized = String(value || "")
    .trim()
    .toUpperCase()
    .replace(/=+$/g, "")
    .replace(/\s+/g, "");

  let bits = 0;
  let output = [];
  let currentValue = 0;

  for (const character of normalized) {
    const characterIndex = base32Alphabet.indexOf(character);

    if (characterIndex < 0) {
      throw new Error("The two-factor secret contains invalid characters.");
    }

    currentValue = (currentValue << 5) | characterIndex;
    bits += 5;

    if (bits >= 8) {
      output.push((currentValue >>> (bits - 8)) & 255);
      bits -= 8;
    }
  }

  return Buffer.from(output);
}

function formatTotpSecret(secret) {
  return String(secret || "")
    .replace(/\s+/g, "")
    .match(/.{1,4}/g)
    ?.join(" ") || "";
}

function normalizeTotpToken(value) {
  return String(value || "").replace(/\D+/g, "");
}

function generateTotpSecret() {
  return encodeBase32(crypto.randomBytes(totpSecretBytes));
}

function getTotpCounter(timestamp = Date.now()) {
  return Math.floor(timestamp / 1000 / totpTimeStepSeconds);
}

function generateTotpToken(secret, timestamp = Date.now()) {
  const secretBuffer = decodeBase32(secret);
  const counter = getTotpCounter(timestamp);
  const counterBuffer = Buffer.alloc(8);
  counterBuffer.writeBigUInt64BE(BigInt(counter));

  const hmac = crypto.createHmac("sha1", secretBuffer).update(counterBuffer).digest();
  const offset = hmac[hmac.length - 1] & 15;
  const binaryCode =
    ((hmac[offset] & 127) << 24) |
    ((hmac[offset + 1] & 255) << 16) |
    ((hmac[offset + 2] & 255) << 8) |
    (hmac[offset + 3] & 255);

  return String(binaryCode % 10 ** totpDigits).padStart(totpDigits, "0");
}

function verifyTotpToken(secret, token, windowSteps = 1) {
  const normalizedToken = normalizeTotpToken(token);

  if (!secret || normalizedToken.length !== totpDigits) {
    return false;
  }

  for (let offset = -windowSteps; offset <= windowSteps; offset += 1) {
    if (
      generateTotpToken(
        secret,
        Date.now() + offset * totpTimeStepSeconds * 1000,
      ) === normalizedToken
    ) {
      return true;
    }
  }

  return false;
}

function getAdminTwoFactorState(settings = getAllSettings()) {
  const secret = String(settings.admin_2fa_secret || "").trim();
  const enabled = String(settings.admin_2fa_enabled || "false").trim().toLowerCase() === "true";

  return {
    enabled: enabled && Boolean(secret),
    secret,
  };
}

function buildAdminTotpUri(secret, issuerName) {
  const issuer = String(issuerName || defaultElectionName || "Election Portal").trim();
  const accountName = adminUsername;
  const label = `${issuer}:${accountName}`;
  const params = new URLSearchParams({
    secret,
    issuer,
    algorithm: "SHA1",
    digits: String(totpDigits),
    period: String(totpTimeStepSeconds),
  });

  return `otpauth://totp/${encodeURIComponent(label)}?${params.toString()}`;
}

async function buildAdminTotpQrCodeDataUrl(otpauthUri) {
  if (!otpauthUri) {
    return "";
  }

  try {
    return await QRCode.toDataURL(otpauthUri, {
      errorCorrectionLevel: "M",
      margin: 1,
      width: 220,
      color: {
        dark: "#102338",
        light: "#FFFFFFFF",
      },
    });
  } catch (_error) {
    return "";
  }
}

function normalizeReferenceCode(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function createNominationReferenceCode() {
  const alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  const segmentLength = 4;
  const segments = [];

  for (let segmentIndex = 0; segmentIndex < 2; segmentIndex += 1) {
    let segment = "";

    for (let characterIndex = 0; characterIndex < segmentLength; characterIndex += 1) {
      segment += alphabet[crypto.randomInt(0, alphabet.length)];
    }

    segments.push(segment);
  }

  return `NOM-${segments.join("-")}`;
}

function normalizeApplicationNumber(value) {
  return normalizeReferenceCode(value);
}

function createNominationApplicationNumber() {
  return createNominationReferenceCode().replace(/^NOM-/, "APP-");
}

function createUniqueNominationApplicationNumber() {
  for (let attempt = 0; attempt < 50; attempt += 1) {
    const applicationNumber = createNominationApplicationNumber();
    const existingRecord = db.prepare(`
      SELECT id
      FROM nominations
      WHERE application_number = ?
    `).get(applicationNumber);

    if (!existingRecord) {
      return applicationNumber;
    }
  }

  throw new Error("Could not generate a unique nomination application number.");
}

function parseReferenceCodeList(rawValue) {
  const candidates = String(rawValue || "")
    .split(/[\r\n,;]+/)
    .map((value) => normalizeReferenceCode(value))
    .filter(Boolean);

  return [...new Set(candidates)];
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

function clampNumber(value, minimum, maximum) {
  return Math.min(Math.max(value, minimum), maximum);
}

function hashNumberFromString(value) {
  return Array.from(String(value || "")).reduce((hash, character) => {
    return (hash * 31 + character.charCodeAt(0)) >>> 0;
  }, 0);
}

function sampleRasterPixel(image, x, y) {
  const pixelX = clampNumber(x, 0, image.width - 1);
  const pixelY = clampNumber(y, 0, image.height - 1);
  const offset = (pixelY * image.width + pixelX) * 4;

  return {
    red: image.data[offset] || 0,
    green: image.data[offset + 1] || 0,
    blue: image.data[offset + 2] || 0,
    alpha: image.data[offset + 3] ?? 255,
  };
}

function resizeRasterImage(image, maxDimension = 1600) {
  const largestSide = Math.max(image.width, image.height);

  if (!Number.isFinite(largestSide) || largestSide <= maxDimension) {
    return {
      width: image.width,
      height: image.height,
      data: Buffer.from(image.data),
    };
  }

  const scale = maxDimension / largestSide;
  const nextWidth = Math.max(1, Math.round(image.width * scale));
  const nextHeight = Math.max(1, Math.round(image.height * scale));
  const resizedData = Buffer.alloc(nextWidth * nextHeight * 4);

  for (let y = 0; y < nextHeight; y += 1) {
    const sourceY = Math.min(image.height - 1, Math.floor(y / scale));

    for (let x = 0; x < nextWidth; x += 1) {
      const sourceX = Math.min(image.width - 1, Math.floor(x / scale));
      const sourceOffset = (sourceY * image.width + sourceX) * 4;
      const targetOffset = (y * nextWidth + x) * 4;

      resizedData[targetOffset] = image.data[sourceOffset];
      resizedData[targetOffset + 1] = image.data[sourceOffset + 1];
      resizedData[targetOffset + 2] = image.data[sourceOffset + 2];
      resizedData[targetOffset + 3] = image.data[sourceOffset + 3];
    }
  }

  return {
    width: nextWidth,
    height: nextHeight,
    data: resizedData,
  };
}

function decodeRasterImage(filePath) {
  const fileBuffer = fs.readFileSync(filePath);
  const extension = path.extname(filePath).toLowerCase();

  if (extension === ".png") {
    const parsedImage = PNG.sync.read(fileBuffer);
    return {
      width: parsedImage.width,
      height: parsedImage.height,
      data: Buffer.from(parsedImage.data),
    };
  }

  if (extension === ".jpg" || extension === ".jpeg") {
    const parsedImage = jpeg.decode(fileBuffer, {
      useTArray: true,
    });
    return {
      width: parsedImage.width,
      height: parsedImage.height,
      data: Buffer.from(parsedImage.data),
    };
  }

  try {
    const parsedPng = PNG.sync.read(fileBuffer);
    return {
      width: parsedPng.width,
      height: parsedPng.height,
      data: Buffer.from(parsedPng.data),
    };
  } catch (_pngError) {
    const parsedJpeg = jpeg.decode(fileBuffer, {
      useTArray: true,
    });

    return {
      width: parsedJpeg.width,
      height: parsedJpeg.height,
      data: Buffer.from(parsedJpeg.data),
    };
  }
}

function getAverageCornerColor(image) {
  const samplePoints = [
    [2, 2],
    [image.width - 3, 2],
    [2, image.height - 3],
    [image.width - 3, image.height - 3],
    [Math.floor(image.width * 0.5), 2],
    [Math.floor(image.width * 0.5), image.height - 3],
  ];
  const totals = samplePoints.reduce(
    (accumulator, [x, y]) => {
      const sample = sampleRasterPixel(image, x, y);
      accumulator.red += sample.red;
      accumulator.green += sample.green;
      accumulator.blue += sample.blue;
      return accumulator;
    },
    { red: 0, green: 0, blue: 0 },
  );

  return {
    red: Math.round(totals.red / samplePoints.length),
    green: Math.round(totals.green / samplePoints.length),
    blue: Math.round(totals.blue / samplePoints.length),
  };
}

function getColorDistance(colorA, colorB) {
  const redDelta = colorA.red - colorB.red;
  const greenDelta = colorA.green - colorB.green;
  const blueDelta = colorA.blue - colorB.blue;
  return Math.sqrt(redDelta ** 2 + greenDelta ** 2 + blueDelta ** 2);
}

function isBackgroundLikePixel(pixel, backgroundColor) {
  const maxChannel = Math.max(pixel.red, pixel.green, pixel.blue);
  const minChannel = Math.min(pixel.red, pixel.green, pixel.blue);
  const brightness = (pixel.red + pixel.green + pixel.blue) / 3;
  const saturation = maxChannel - minChannel;
  const distanceFromBackground = getColorDistance(pixel, backgroundColor);

  if (pixel.alpha < 60) {
    return true;
  }

  if (distanceFromBackground <= 48) {
    return true;
  }

  if (distanceFromBackground <= 72 && saturation <= 28 && brightness >= 190) {
    return true;
  }

  if (brightness >= 242 && saturation <= 18) {
    return true;
  }

  return false;
}

function cropTransparentRasterImage(image, padding = 20) {
  let minX = image.width;
  let minY = image.height;
  let maxX = -1;
  let maxY = -1;

  for (let y = 0; y < image.height; y += 1) {
    for (let x = 0; x < image.width; x += 1) {
      const alpha = image.data[(y * image.width + x) * 4 + 3];
      if (alpha > 18) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x);
        maxY = Math.max(maxY, y);
      }
    }
  }

  if (maxX < minX || maxY < minY) {
    return image;
  }

  const croppedMinX = Math.max(0, minX - padding);
  const croppedMinY = Math.max(0, minY - padding);
  const croppedMaxX = Math.min(image.width - 1, maxX + padding);
  const croppedMaxY = Math.min(image.height - 1, maxY + padding);
  const croppedWidth = croppedMaxX - croppedMinX + 1;
  const croppedHeight = croppedMaxY - croppedMinY + 1;
  const croppedData = Buffer.alloc(croppedWidth * croppedHeight * 4);

  for (let y = 0; y < croppedHeight; y += 1) {
    for (let x = 0; x < croppedWidth; x += 1) {
      const sourceOffset = ((croppedMinY + y) * image.width + (croppedMinX + x)) * 4;
      const targetOffset = (y * croppedWidth + x) * 4;
      image.data.copy(croppedData, targetOffset, sourceOffset, sourceOffset + 4);
    }
  }

  return {
    width: croppedWidth,
    height: croppedHeight,
    data: croppedData,
  };
}

function removePortraitBackground(filePath) {
  const decodedImage = resizeRasterImage(decodeRasterImage(filePath));
  const backgroundColor = getAverageCornerColor(decodedImage);
  const backgroundMask = new Uint8Array(decodedImage.width * decodedImage.height);
  const queue = [];

  const markPixel = (x, y) => {
    if (x < 0 || x >= decodedImage.width || y < 0 || y >= decodedImage.height) {
      return;
    }

    const index = y * decodedImage.width + x;
    if (backgroundMask[index]) {
      return;
    }

    const pixel = sampleRasterPixel(decodedImage, x, y);
    if (!isBackgroundLikePixel(pixel, backgroundColor)) {
      return;
    }

    backgroundMask[index] = 1;
    queue.push(index);
  };

  for (let x = 0; x < decodedImage.width; x += 1) {
    markPixel(x, 0);
    markPixel(x, decodedImage.height - 1);
  }

  for (let y = 0; y < decodedImage.height; y += 1) {
    markPixel(0, y);
    markPixel(decodedImage.width - 1, y);
  }

  let queueIndex = 0;

  while (queueIndex < queue.length) {
    const currentIndex = queue[queueIndex];
    queueIndex += 1;
    const x = currentIndex % decodedImage.width;
    const y = Math.floor(currentIndex / decodedImage.width);

    markPixel(x + 1, y);
    markPixel(x - 1, y);
    markPixel(x, y + 1);
    markPixel(x, y - 1);
  }

  const cleanedData = Buffer.from(decodedImage.data);

  for (let y = 0; y < decodedImage.height; y += 1) {
    for (let x = 0; x < decodedImage.width; x += 1) {
      const pixelIndex = y * decodedImage.width + x;
      const offset = pixelIndex * 4;

      if (backgroundMask[pixelIndex]) {
        cleanedData[offset + 3] = 0;
        continue;
      }

      const pixel = sampleRasterPixel(decodedImage, x, y);
      const distanceFromBackground = getColorDistance(pixel, backgroundColor);
      const hasBackgroundNeighbour =
        (x > 0 && backgroundMask[pixelIndex - 1]) ||
        (x < decodedImage.width - 1 && backgroundMask[pixelIndex + 1]) ||
        (y > 0 && backgroundMask[pixelIndex - decodedImage.width]) ||
        (y < decodedImage.height - 1 && backgroundMask[pixelIndex + decodedImage.width]);

      if (hasBackgroundNeighbour && distanceFromBackground < 95) {
        cleanedData[offset + 3] = Math.min(cleanedData[offset + 3], 208);
      }
    }
  }

  const croppedPortrait = cropTransparentRasterImage(
    {
      width: decodedImage.width,
      height: decodedImage.height,
      data: cleanedData,
    },
    Math.max(18, Math.round(Math.min(decodedImage.width, decodedImage.height) * 0.04)),
  );

  const pngImage = new PNG({
    width: croppedPortrait.width,
    height: croppedPortrait.height,
  });
  pngImage.data = Buffer.from(croppedPortrait.data);
  return PNG.sync.write(pngImage);
}

function getOrganizationDisplayName(electionName) {
  const cleanedName = String(electionName || defaultElectionName)
    .replace(/\bElection Portal\b/gi, "")
    .replace(/\bElection Administration\b/gi, "")
    .replace(/\s{2,}/g, " ")
    .trim();

  return cleanedName || String(electionName || defaultElectionName).trim();
}

function splitFlyerText(value, maxLines = 2, targetLineLength = 16) {
  const words = String(value || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  if (words.length === 0) {
    return [];
  }

  const lines = [];
  let currentLine = "";

  words.forEach((word, index) => {
    const nextLine = currentLine ? `${currentLine} ${word}` : word;
    const isLastLine = lines.length === maxLines - 1;

    if (!isLastLine && nextLine.length > targetLineLength && currentLine) {
      lines.push(currentLine);
      currentLine = word;
      return;
    }

    currentLine = nextLine;

    if (index === words.length - 1) {
      lines.push(currentLine);
    }
  });

  if (lines.length > maxLines) {
    const head = lines.slice(0, maxLines - 1);
    const tail = lines.slice(maxLines - 1).join(" ");
    return [...head, tail];
  }

  return lines;
}

function getNominationFlyerTheme(nomination) {
  const themes = [
    {
      key: "heritage-arc",
      paper: "#fbf6ee",
      primary: "#0d5677",
      secondary: "#f0a848",
      accent: "#0f9cb3",
      dark: "#102338",
      soft: "#ffffff",
      layout: "arc",
    },
    {
      key: "emerald-column",
      paper: "#f7f9f5",
      primary: "#14665a",
      secondary: "#f4b942",
      accent: "#7ed3c2",
      dark: "#16313c",
      soft: "#ffffff",
      layout: "column",
    },
    {
      key: "sunrise-ribbon",
      paper: "#fff7ef",
      primary: "#b7512a",
      secondary: "#183b6b",
      accent: "#f09d5f",
      dark: "#18263a",
      soft: "#fffdf8",
      layout: "ribbon",
    },
    {
      key: "royal-spotlight",
      paper: "#f6f7fb",
      primary: "#29407f",
      secondary: "#0fa3b1",
      accent: "#f2b441",
      dark: "#16223a",
      soft: "#ffffff",
      layout: "spotlight",
    },
  ];
  const seed = hashNumberFromString(
    `${nomination.applicationNumber}|${nomination.fullName}|${nomination.positionName}`,
  );
  const theme = themes[seed % themes.length];

  return {
    ...theme,
    seed,
  };
}

function safeDrawImage(document, imageSource, x, y, options = {}) {
  if (!imageSource) {
    return false;
  }

  try {
    document.image(imageSource, x, y, options);
    return true;
  } catch (_error) {
    return false;
  }
}

function drawFlyerPortrait(document, imageSource, options = {}) {
  const {
    x = 0,
    y = 0,
    width = 200,
    height = 260,
    shape = "rounded",
    radius = 28,
    borderWidth = 6,
    borderColor = "#ffffff",
    panelColor = "#ffffff",
    imageInset = 0,
    fallbackLabel = "",
    fallbackTextColor = "#102338",
  } = options;
  const innerX = x + imageInset;
  const innerY = y + imageInset;
  const innerWidth = Math.max(width - imageInset * 2, 1);
  const innerHeight = Math.max(height - imageInset * 2, 1);

  const drawShapePath = () => {
    if (shape === "circle") {
      const radiusValue = Math.min(width, height) / 2;
      document.circle(x + width / 2, y + height / 2, radiusValue);
      return;
    }

    document.roundedRect(x, y, width, height, radius);
  };

  document.save();
  drawShapePath();
  document.fillColor(panelColor).fill();
  document.restore();

  if (imageSource) {
    document.save();
    drawShapePath();
    document.clip();
    safeDrawImage(document, imageSource, innerX, innerY, {
      fit: [innerWidth, innerHeight],
      align: "center",
      valign: "center",
    });
    document.restore();
  } else if (fallbackLabel) {
    document
      .fillColor(fallbackTextColor)
      .font("Helvetica-Bold")
      .fontSize(Math.min(width, height) * 0.24)
      .text(fallbackLabel, x, y + height / 2 - 26, {
        width,
        align: "center",
      });
  }

  document.save();
  drawShapePath();
  document.lineWidth(borderWidth).strokeColor(borderColor).stroke();
  document.restore();
}

function getNominationPortraitSource(photoPath) {
  const resolvedPhotoPath = photoPath ? resolveAssetPath(photoPath) : "";
  if (!resolvedPhotoPath || !fs.existsSync(resolvedPhotoPath)) {
    return "";
  }

  return resolvedPhotoPath;
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

function clampInteger(value, fallback, minimum, maximum) {
  const parsedValue = Number.parseInt(String(value ?? ""), 10);
  const safeValue = Number.isFinite(parsedValue) ? parsedValue : fallback;
  return Math.min(Math.max(safeValue, minimum), maximum);
}

function getOtpProviderLabel(provider) {
  switch (provider) {
    case "arkesel":
      return "Arkesel";
    case "twilio":
      return "Twilio";
    case "dev":
      return "Development";
    default:
      return "Disabled";
  }
}

function getOtpConfig() {
  const settings = getAllSettings();
  const twilioAccountSid = String(settings.twilio_account_sid || envTwilioAccountSid || "").trim();
  const twilioAuthToken = String(settings.twilio_auth_token || envTwilioAuthToken || "").trim();
  const twilioVerifyServiceSid = String(
    settings.twilio_verify_service_sid || envTwilioVerifyServiceSid || "",
  ).trim();
  const arkeselApiKey = String(settings.arkesel_api_key || envArkeselApiKey || "").trim();
  const arkeselSenderId = String(settings.arkesel_sender_id || envArkeselSenderId || "").trim();
  const arkeselOtpMessageTemplate = String(
    settings.arkesel_otp_message || defaultArkeselOtpMessageTemplate,
  ).trim();
  const twilioConfigured = Boolean(
    twilioAccountSid && twilioAuthToken && twilioVerifyServiceSid,
  );
  const arkeselConfigured = Boolean(arkeselApiKey && arkeselSenderId);
  const configuredProvider = String(settings.otp_provider || envConfiguredOtpProvider || "")
    .trim()
    .toLowerCase();
  const provider =
    configuredProvider === "twilio" ||
    configuredProvider === "arkesel" ||
    configuredProvider === "dev" ||
    configuredProvider === "disabled"
      ? configuredProvider
      : twilioConfigured
        ? "twilio"
        : arkeselConfigured
          ? "arkesel"
          : isProduction
            ? "disabled"
            : "dev";
  const configuredTtlMinutes = clampInteger(
    settings.otp_ttl_minutes || envConfiguredOtpTtlMinutes,
    envConfiguredOtpTtlMinutes,
    1,
    30,
  );
  const ttlMinutes = provider === "arkesel"
    ? Math.min(configuredTtlMinutes, 10)
    : configuredTtlMinutes;
  const resendCooldownSeconds = clampInteger(
    settings.otp_resend_cooldown_seconds || envOtpResendCooldownSeconds,
    envOtpResendCooldownSeconds,
    0,
    300,
  );

  return {
    provider,
    providerLabel: getOtpProviderLabel(provider),
    ttlMinutes,
    resendCooldownSeconds,
    twilioAccountSid,
    twilioAuthToken,
    twilioVerifyServiceSid,
    twilioConfigured,
    arkeselApiKey,
    arkeselSenderId,
    arkeselOtpMessageTemplate,
    arkeselConfigured,
  };
}

function getAdminOtpSettingsView() {
  const settings = getAllSettings();
  const otpConfig = getOtpConfig();
  const preferredProvider = String(settings.otp_provider || envConfiguredOtpProvider || otpConfig.provider)
    .trim()
    .toLowerCase();

  return {
    provider:
      preferredProvider === "twilio" ||
      preferredProvider === "arkesel" ||
      preferredProvider === "dev" ||
      preferredProvider === "disabled"
        ? preferredProvider
        : otpConfig.provider,
    effectiveProvider: otpConfig.provider,
    effectiveProviderLabel: otpConfig.providerLabel,
    ttlMinutes: otpConfig.ttlMinutes,
    resendCooldownSeconds: otpConfig.resendCooldownSeconds,
    arkeselApiKey: String(settings.arkesel_api_key || envArkeselApiKey || "").trim(),
    arkeselSenderId: String(settings.arkesel_sender_id || envArkeselSenderId || "").trim(),
    arkeselOtpMessage: String(
      settings.arkesel_otp_message || defaultArkeselOtpMessageTemplate,
    ).trim(),
    isEnabled: otpConfig.provider === "twilio" || otpConfig.provider === "arkesel" || otpConfig.provider === "dev",
    arkeselConfigured: otpConfig.arkeselConfigured,
    twilioConfigured: otpConfig.twilioConfigured,
  };
}

function isEnabledSetting(value) {
  return String(value || "").trim().toLowerCase() === "true";
}

function getCaptchaConfig(settings = getAllSettings()) {
  const siteKey = String(settings.captcha_site_key || envCaptchaSiteKey || "").trim();
  const secretKey = String(settings.captcha_secret_key || envCaptchaSecretKey || "").trim();
  const isConfigured = Boolean(siteKey && secretKey);
  const isEnabled = isEnabledSetting(settings.captcha_enabled) && isConfigured;

  return {
    isConfigured,
    isEnabled,
    siteKey,
    secretKey,
    protectVoterLogin: isEnabledSetting(settings.captcha_protect_voter_login || "true"),
    protectAdminLogin: isEnabledSetting(settings.captcha_protect_admin_login || "true"),
    protectObserverLogin: isEnabledSetting(settings.captcha_protect_observer_login || "true"),
    protectNomination: isEnabledSetting(settings.captcha_protect_nomination || "true"),
  };
}

function getCaptchaPublicConfig(settings = getAllSettings()) {
  const config = getCaptchaConfig(settings);
  return {
    isConfigured: config.isConfigured,
    isEnabled: config.isEnabled,
    siteKey: config.siteKey,
  };
}

function getAdminCaptchaSettingsView() {
  const settings = getAllSettings();
  const config = getCaptchaConfig(settings);
  return {
    isConfigured: config.isConfigured,
    isEnabled: config.isEnabled,
    requestedEnabled: isEnabledSetting(settings.captcha_enabled),
    siteKey: config.siteKey,
    secretConfigured: Boolean(String(settings.captcha_secret_key || envCaptchaSecretKey || "").trim()),
    protectVoterLogin: config.protectVoterLogin,
    protectAdminLogin: config.protectAdminLogin,
    protectObserverLogin: config.protectObserverLogin,
    protectNomination: config.protectNomination,
  };
}

function isCaptchaRequiredForContext(context) {
  const config = getCaptchaConfig();
  if (!config.isEnabled) {
    return false;
  }

  switch (context) {
    case "voter_login":
      return config.protectVoterLogin;
    case "admin_login":
      return config.protectAdminLogin;
    case "observer_login":
      return config.protectObserverLogin;
    case "nomination":
      return config.protectNomination;
    default:
      return false;
  }
}

function getCaptchaFailureRedirect(context) {
  switch (context) {
    case "admin_login":
      return "/admin/login";
    case "observer_login":
      return "/observer/login";
    case "nomination":
      return "/nomination/status/login";
    case "voter_login":
    default:
      return "/vote/login";
  }
}

async function verifyCaptchaSubmission(req, context) {
  if (!isCaptchaRequiredForContext(context)) {
    return { ok: true, skipped: true };
  }

  const config = getCaptchaConfig();
  const token = String(req.body["cf-turnstile-response"] || "").trim();
  if (!token) {
    return { ok: false, reason: "missing_token" };
  }

  const formData = new URLSearchParams();
  formData.set("secret", config.secretKey);
  formData.set("response", token);
  if (req.ip) {
    formData.set("remoteip", req.ip);
  }

  try {
    const response = await fetch(turnstileVerifyUrl, {
      method: "POST",
      body: formData,
    });
    const payload = await response.json();

    if (payload.success) {
      return { ok: true };
    }

    return {
      ok: false,
      reason: "verification_failed",
      errorCodes: Array.isArray(payload["error-codes"]) ? payload["error-codes"] : [],
    };
  } catch (error) {
    return {
      ok: false,
      reason: "verification_error",
      message: error.message,
    };
  }
}

function requireCaptcha(context) {
  return async (req, res, next) => {
    const result = await verifyCaptchaSubmission(req, context);
    if (result.ok) {
      return next();
    }

    logAudit(req, "security", context, "captcha_verification_failed", {
      reason: result.reason,
      errorCodes: result.errorCodes || [],
    });
    setFlash(req, "error", "Please complete the security check before continuing.");
    return res.redirect(getCaptchaFailureRedirect(context));
  };
}

function logSystemAudit(action, details = {}) {
  const insertResult = db.prepare(`
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

  createAdminNotificationFromAudit(
    "system",
    "scheduler",
    action,
    details,
    Number(insertResult.lastInsertRowid),
  );
}

function normalizeAdminNotification(row) {
  return {
    id: Number(row.id),
    category: row.category || "system",
    priority: row.priority || "normal",
    title: row.title || "Notification",
    body: row.body || "",
    linkUrl: row.linkUrl || "",
    sourceType: row.sourceType || "",
    sourceId: row.sourceId || "",
    createdBy: row.createdBy || "",
    readAt: row.readAt || "",
    createdAt: row.createdAt || "",
    updatedAt: row.updatedAt || "",
    isUnread: !row.readAt,
    timeLabel: row.createdAt ? formatDateTime(row.createdAt) : "",
  };
}

function getAdminNotifications(limit = 12) {
  const safeLimit = clampInteger(limit, 12, 1, 50);
  return db.prepare(`
    SELECT
      id,
      category,
      priority,
      title,
      body,
      link_url AS linkUrl,
      source_type AS sourceType,
      source_id AS sourceId,
      created_by AS createdBy,
      read_at AS readAt,
      created_at AS createdAt,
      updated_at AS updatedAt
    FROM admin_notifications
    ORDER BY created_at DESC, id DESC
    LIMIT ?
  `).all(safeLimit).map(normalizeAdminNotification);
}

function getUnreadAdminNotificationCount() {
  return Number(
    db.prepare(`
      SELECT COUNT(*) AS total
      FROM admin_notifications
      WHERE read_at IS NULL
    `).get()?.total || 0,
  );
}

function getAdminNotificationSnapshot(limit = 12) {
  return {
    unreadCount: getUnreadAdminNotificationCount(),
    notifications: getAdminNotifications(limit),
    generatedAt: nowIso(),
  };
}

function writeAdminNotificationEvent(response, eventName, payload) {
  response.write(`event: ${eventName}\n`);
  response.write(`data: ${JSON.stringify(payload)}\n\n`);
}

function broadcastAdminNotificationUpdate() {
  if (adminNotificationClients.size === 0) {
    return;
  }

  const payload = getAdminNotificationSnapshot();
  for (const response of adminNotificationClients) {
    try {
      writeAdminNotificationEvent(response, "notifications:update", payload);
    } catch (_error) {
      adminNotificationClients.delete(response);
    }
  }
}

function createAdminNotification({
  category = "system",
  priority = "normal",
  title,
  body = "",
  linkUrl = "",
  sourceType = "",
  sourceId = "",
  createdBy = "system",
} = {}) {
  const safeTitle = String(title || "").trim().slice(0, 120);
  if (!safeTitle) {
    return null;
  }

  const timestamp = nowIso();
  const result = db.prepare(`
    INSERT INTO admin_notifications (
      category,
      priority,
      title,
      body,
      link_url,
      source_type,
      source_id,
      created_by,
      created_at,
      updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    String(category || "system").trim().slice(0, 40) || "system",
    String(priority || "normal").trim().slice(0, 20) || "normal",
    safeTitle,
    String(body || "").trim().slice(0, 500),
    String(linkUrl || "").trim().slice(0, 240),
    String(sourceType || "").trim().slice(0, 60),
    String(sourceId || "").trim().slice(0, 80),
    String(createdBy || "system").trim().slice(0, 80),
    timestamp,
    timestamp,
  );

  broadcastAdminNotificationUpdate();
  return Number(result.lastInsertRowid);
}

function markAdminNotificationsRead({ ids = [], all = false } = {}) {
  const timestamp = nowIso();
  if (all) {
    db.prepare(`
      UPDATE admin_notifications
      SET read_at = COALESCE(read_at, ?), updated_at = ?
      WHERE read_at IS NULL
    `).run(timestamp, timestamp);
    broadcastAdminNotificationUpdate();
    return;
  }

  const safeIds = [...new Set(ids.map((id) => parseInteger(id, 0)).filter((id) => id > 0))];
  if (safeIds.length === 0) {
    return;
  }

  const placeholders = safeIds.map(() => "?").join(", ");
  db.prepare(`
    UPDATE admin_notifications
    SET read_at = COALESCE(read_at, ?), updated_at = ?
    WHERE id IN (${placeholders})
  `).run(timestamp, timestamp, ...safeIds);
  broadcastAdminNotificationUpdate();
}

function getAdminNotificationForAudit(actorType, actorIdentifier, action, details = {}, auditId = "") {
  switch (action) {
    case "vote_submitted":
      return {
        category: "vote",
        priority: "normal",
        title: "Vote recorded",
        body: `${actorIdentifier} submitted ${details.submittedChoices || 0} choice${Number(details.submittedChoices || 0) === 1 ? "" : "s"} and skipped ${details.skippedCount || 0}.`,
        linkUrl: "/admin/results",
      };
    case "nomination_submitted":
      return {
        category: "nomination",
        priority: "high",
        title: "New nomination submitted",
        body: `${details.applicationNumber || actorIdentifier} applied for ${details.positionName || "a position"} (${details.staffId || "staff ID not listed"}).`,
        linkUrl: details.nominationId ? `/admin/nominations/${details.nominationId}` : "/admin/nominations",
      };
    case "nomination_resubmitted":
      return {
        category: "nomination",
        priority: "normal",
        title: "Nomination correction resubmitted",
        body: `${details.applicationNumber || actorIdentifier} resubmitted corrections for ${details.positionName || "a position"}.`,
        linkUrl: details.nominationId ? `/admin/nominations/${details.nominationId}` : "/admin/nominations",
      };
    case "observer_incident_submitted":
      return {
        category: "observer",
        priority: "high",
        title: "Observer incident submitted",
        body: `${actorIdentifier} submitted a ${details.category || "general"} incident report.`,
        linkUrl: "/admin/observers#observer-incidents",
      };
    case "results_sms_sent":
      return {
        category: "results",
        priority: details.failureCount > 0 ? "high" : "normal",
        title: "Provisional results SMS sent",
        body: `${Number(details.successCount || 0).toLocaleString()} sent, ${Number(details.failureCount || 0).toLocaleString()} failed.`,
        linkUrl: "/admin/results",
      };
    case "voter_otp_send_failed":
    case "voter_otp_failed":
    case "otp_test_send_failed":
    case "captcha_verification_failed":
    case "admin_login_failed":
      return {
        category: "security",
        priority: "urgent",
        title: "Security attention needed",
        body: `${humanizeToken(action)} for ${actorIdentifier || actorType}.`,
        linkUrl: action.includes("otp") ? "/admin/otp-logs" : "/admin/audit",
      };
    default:
      return null;
  }
}

function createAdminNotificationFromAudit(actorType, actorIdentifier, action, details = {}, auditId = "") {
  const notification = getAdminNotificationForAudit(
    actorType,
    actorIdentifier,
    action,
    details,
    auditId,
  );

  if (!notification) {
    return null;
  }

  return createAdminNotification({
    ...notification,
    sourceType: "audit",
    sourceId: String(auditId || ""),
    createdBy: actorIdentifier || actorType || "system",
  });
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
  const electionState = computeElectionState(settings);

  let status = "setup";
  let message =
    "Nominations are in setup. Set the nomination window before candidates can apply.";

  if (electionState.isOpen || electionState.isClosed) {
    return {
      status: "closed",
      message: electionState.isOpen
        ? "Nominations close automatically once voting starts."
        : "Nominations are closed for this election cycle.",
      isOpen: false,
      isScheduled: false,
      isClosed: true,
      canSubmit: false,
      badgeLabel: "Closed",
      badgeClass: "status-closed",
    };
  }

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
      message = "Nominations are currently open for applicants.";
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
    if (settings.nominationPhase !== "closed") {
      setSetting("nomination_phase", "closed");
      logSystemAudit("nomination_auto_closed_for_voting", {
        trigger: "election_auto_opened",
        opensAt: settings.opensAt,
      });
    }
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
    void triggerAutomaticResultsSms("election_auto_closed");
  }
}

function syncAutomaticNominationLifecycle() {
  const settings = getElectionSettings();
  const readiness = getNominationReadiness(settings);
  const electionState = computeElectionState(settings);
  const now = dayjs();
  const opensAt = settings.nominationOpensAt ? dayjs(settings.nominationOpensAt) : null;
  const closesAt = settings.nominationClosesAt ? dayjs(settings.nominationClosesAt) : null;

  if ((electionState.isOpen || electionState.isClosed) && settings.nominationPhase !== "closed") {
    setSetting("nomination_phase", "closed");
    logSystemAudit("nomination_auto_closed_for_voting", {
      electionPhase: settings.phase,
      nominationClosesAt: settings.nominationClosesAt,
    });
    return;
  }

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
  const insertResult = db.prepare(`
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

  createAdminNotificationFromAudit(
    actorType,
    actorIdentifier,
    action,
    details,
    Number(insertResult.lastInsertRowid),
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

function isOtpVerificationEnabled(otpConfig = getOtpConfig()) {
  return (
    otpConfig.provider === "twilio" ||
    otpConfig.provider === "arkesel" ||
    otpConfig.provider === "dev"
  );
}

function clearVoterSession(req) {
  req.session.voter = null;
  req.session.ballotSelections = null;
  req.session.pendingBallot = null;
  req.session.pendingVoterVerification = null;
}

function clearAdminAccess(req) {
  req.session.admin = null;
  req.session.pendingAdminTwoFactor = null;
}

function clearNominationSession(req) {
  req.session.nominationApplicant = null;
}

function clearObserverSession(req) {
  req.session.observer = null;
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

function toArkeselOtpNumber(value) {
  const smsPhoneNumber = toSmsPhoneNumber(value);
  return smsPhoneNumber ? smsPhoneNumber.replace(/^\+/, "") : "";
}

function getOtpExpiryIso(otpConfig = getOtpConfig()) {
  return dayjs().add(otpConfig.ttlMinutes, "minute").toISOString();
}

function getOtpResendAvailableIso(otpConfig = getOtpConfig()) {
  return dayjs().add(otpConfig.resendCooldownSeconds, "second").toISOString();
}

function isPendingOtpExpired(pendingVerification) {
  if (!pendingVerification?.expiresAt) {
    return true;
  }

  const expiresAt = dayjs(pendingVerification.expiresAt);
  return !expiresAt.isValid() || !dayjs().isBefore(expiresAt);
}

function buildPendingVoterVerification(
  voterRecord,
  phoneNumber,
  smsPhoneNumber,
  challenge,
  otpConfig = getOtpConfig(),
) {
  return {
    voterId: voterRecord.id,
    staffId: voterRecord.staffId,
    fullName: voterRecord.fullName,
    phoneNumber,
    maskedPhoneNumber: maskPhoneNumber(phoneNumber),
    smsPhoneNumber,
    provider: challenge.provider || otpConfig.provider,
    verificationSid: challenge.verificationSid || "",
    devCodeHash: challenge.devCodeHash || "",
    devCodePreview: challenge.devCodePreview || "",
    sentAt: nowIso(),
    expiresAt: getOtpExpiryIso(otpConfig),
    resendAvailableAt: getOtpResendAvailableIso(otpConfig),
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

function beginNominationApplicantSession(req, nomination) {
  req.session.nominationApplicant = {
    nominationId: nomination.id,
    applicationNumber: nomination.applicationNumber,
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

async function sendTwilioOtpCode(smsPhoneNumber, otpConfig = getOtpConfig()) {
  if (!otpConfig.twilioConfigured) {
    throw new Error("The OTP SMS service is not configured yet. Add the Twilio Verify credentials first.");
  }

  const response = await fetch(
    `https://verify.twilio.com/v2/Services/${encodeURIComponent(otpConfig.twilioVerifyServiceSid)}/Verifications`,
    {
      method: "POST",
      headers: {
        authorization: `Basic ${Buffer.from(
          `${otpConfig.twilioAccountSid}:${otpConfig.twilioAuthToken}`,
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
    provider: "twilio",
    verificationSid: payload?.sid || "",
  };
}

async function sendArkeselOtpCode(smsPhoneNumber, otpConfig = getOtpConfig()) {
  if (!otpConfig.arkeselConfigured) {
    throw new Error(
      "The OTP SMS service is not configured yet. Add your Arkesel API key and sender ID first.",
    );
  }

  if (otpConfig.arkeselSenderId.length > 11) {
    throw new Error("Your Arkesel sender ID must be 11 characters or fewer.");
  }

  if (!otpConfig.arkeselOtpMessageTemplate.includes("%otp_code%")) {
    throw new Error(
      "Your Arkesel OTP message must include %otp_code% so the verification code can be inserted.",
    );
  }

  const arkeselNumber = toArkeselOtpNumber(smsPhoneNumber);
  if (!arkeselNumber) {
    throw new Error("The phone number is not in a valid format for Arkesel OTP delivery.");
  }

  const response = await fetch("https://sms.arkesel.com/api/otp/generate", {
    method: "POST",
    headers: {
      "api-key": otpConfig.arkeselApiKey,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      expiry: otpConfig.ttlMinutes,
      length: devOtpCodeLength,
      medium: "sms",
      message: otpConfig.arkeselOtpMessageTemplate,
      number: arkeselNumber,
      sender_id: otpConfig.arkeselSenderId,
      type: "numeric",
    }),
  });
  const payload = await parseOtpApiResponse(response);

  if (!response.ok || String(payload?.code || "") !== "1000") {
    throw new Error(
      payload?.message ||
        "The OTP SMS could not be sent right now. Please try again in a moment.",
    );
  }

  return { provider: "arkesel" };
}

function isLikelyArkeselSmsFailure(payload) {
  const status = String(payload?.status || payload?.code || "").toLowerCase();
  const message = String(
    payload?.message || payload?.msg || payload?.error || "",
  ).toLowerCase();

  return (
    ["error", "failed", "fail", "false"].includes(status) ||
    message.includes("invalid") ||
    message.includes("insufficient") ||
    message.includes("failed") ||
    message.includes("error")
  );
}

async function sendArkeselTextMessage(phoneNumber, message, otpConfig = getOtpConfig()) {
  if (!otpConfig.arkeselConfigured) {
    throw new Error("Add your Arkesel API key and sender ID before sending SMS.");
  }

  if (otpConfig.arkeselSenderId.length > 11) {
    throw new Error("Your Arkesel sender ID must be 11 characters or fewer.");
  }

  const arkeselNumber = toArkeselOtpNumber(phoneNumber);
  if (!arkeselNumber) {
    throw new Error("The phone number is not in a valid format for Arkesel SMS delivery.");
  }

  const smsText = normalizeSmsMessageText(message);
  if (!smsText) {
    throw new Error("SMS message cannot be empty.");
  }

  const query = new URLSearchParams({
    action: "send-sms",
    api_key: otpConfig.arkeselApiKey,
    to: arkeselNumber,
    from: otpConfig.arkeselSenderId,
    sms: smsText,
  });
  const response = await fetch(`https://sms.arkesel.com/sms/api?${query.toString()}`);
  const payload = await parseOtpApiResponse(response);

  if (!response.ok || isLikelyArkeselSmsFailure(payload)) {
    throw new Error(
      payload?.message ||
        payload?.msg ||
        payload?.error ||
        "The SMS could not be sent right now. Please check your Arkesel balance and try again.",
    );
  }

  return {
    provider: "arkesel",
    payload,
  };
}

async function sendOtpChallenge(smsPhoneNumber, otpConfig = getOtpConfig()) {
  if (otpConfig.provider === "twilio") {
    return sendTwilioOtpCode(smsPhoneNumber, otpConfig);
  }

  if (otpConfig.provider === "arkesel") {
    return sendArkeselOtpCode(smsPhoneNumber, otpConfig);
  }

  if (otpConfig.provider === "dev") {
    if (isProduction) {
      throw new Error(
        "Development OTP mode is not allowed in production. Configure Twilio Verify or Arkesel before using OTP on the live site.",
      );
    }

    const devCode = generateDevOtpCode();
    return {
      provider: "dev",
      verificationSid: `DEV-${crypto.randomUUID()}`,
      devCodeHash: hashOtpCode(devCode),
      devCodePreview: devCode,
    };
  }

  return null;
}

async function verifyTwilioOtpCode(pendingVerification, code, otpConfig = getOtpConfig()) {
  if (!otpConfig.twilioConfigured) {
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
      otpConfig.twilioVerifyServiceSid,
    )}/VerificationCheck`,
    {
      method: "POST",
      headers: {
        authorization: `Basic ${Buffer.from(
          `${otpConfig.twilioAccountSid}:${otpConfig.twilioAuthToken}`,
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

async function verifyArkeselOtpCode(
  pendingVerification,
  code,
  otpConfig = getOtpConfig(),
) {
  if (!otpConfig.arkeselConfigured) {
    throw new Error(
      "The OTP SMS service is not configured yet. Add your Arkesel API key and sender ID first.",
    );
  }

  const arkeselNumber = toArkeselOtpNumber(pendingVerification.smsPhoneNumber);
  if (!arkeselNumber) {
    throw new Error("The phone number is not in a valid format for Arkesel OTP verification.");
  }

  const response = await fetch("https://sms.arkesel.com/api/otp/verify", {
    method: "POST",
    headers: {
      "api-key": otpConfig.arkeselApiKey,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      code,
      number: arkeselNumber,
    }),
  });
  const payload = await parseOtpApiResponse(response);
  const responseCode = String(payload?.code || payload?.status || "");

  if (!response.ok && response.status !== 422) {
    throw new Error(
      payload?.message || "The OTP could not be verified right now. Please try again.",
    );
  }

  return {
    approved: responseCode === "1100",
    errorMessage:
      responseCode === "1105"
        ? "This OTP has expired. Request a new code and try again."
        : responseCode === "1104"
          ? "The OTP code is incorrect. Please try again."
          : payload?.message || "The OTP code could not be verified. Please try again.",
  };
}

async function verifyOtpChallenge(pendingVerification, code) {
  const otpConfig = getOtpConfig();
  const provider = pendingVerification?.provider || otpConfig.provider;

  if (provider === "twilio") {
    return verifyTwilioOtpCode(pendingVerification, code, otpConfig);
  }

  if (provider === "arkesel") {
    return verifyArkeselOtpCode(pendingVerification, code, otpConfig);
  }

  if (provider === "dev") {
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
        stepNumber: index + 1,
        positionId: position.id,
        positionName: position.name,
        candidateId: null,
        candidateName: "",
        candidatePhotoPath: "",
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
      stepNumber: index + 1,
      positionId: position.id,
      positionName: position.name,
      candidateId: candidate.id,
      candidateName: candidate.name,
      candidatePhotoPath: candidate.photoPath || "",
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

function sanitizeSmsText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeSmsMessageText(value) {
  return String(value || "")
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

function truncateSmsText(value, maxLength = 40) {
  const text = sanitizeSmsText(value);
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, Math.max(maxLength - 3, 1)).trim()}...`;
}

function buildResultsSmsFingerprint(payload) {
  const compactResults = payload.results.map((position) => ({
    id: position.id,
    name: position.name,
    totalVotes: position.totalVotes,
    winnerLabel: position.winnerLabel,
    candidates: position.candidates.map((candidate) => ({
      id: candidate.id,
      name: candidate.name,
      voteCount: candidate.voteCount,
    })),
  }));

  return crypto
    .createHash("sha256")
    .update(
      JSON.stringify({
        electionName: payload.settings.electionName,
        opensAt: payload.settings.opensAt,
        closesAt: payload.settings.closesAt,
        totalVoters: payload.metrics.totalVoters,
        votedCount: payload.metrics.votedCount,
        results: compactResults,
      }),
    )
    .digest("hex");
}

function buildResultsSmsMessage(payload) {
  const electionName = truncateSmsText(payload.settings.electionName, 52);
  const turnout = payload.metrics.totalVoters
    ? formatPercent(payload.metrics.votedCount / payload.metrics.totalVoters)
    : "0.0%";
  const winnerItems = payload.results.map((position) => {
    const leadingCandidate = position.candidates.find((candidate) => candidate.isLeading);
    const winnerName = truncateSmsText(position.winnerLabel || "No votes recorded", 30);
    const winnerVotes = leadingCandidate ? leadingCandidate.voteCount : 0;
    const voteLabel = Number(winnerVotes) === 1 ? "vote" : "votes";
    return `${truncateSmsText(position.name, 20)}: ${winnerName} - ${winnerVotes} ${voteLabel}`;
  });

  let includedCount = winnerItems.length;
  let message = "";

  do {
    const shownWinners = winnerItems
      .slice(0, includedCount)
      .map((winner, index) => `${index + 1}. ${winner}`);
    const hiddenCount = winnerItems.length - includedCount;
    const hiddenText = hiddenCount > 0 ? `+${hiddenCount} more position${hiddenCount === 1 ? "" : "s"}` : "";
    const winnerLines = shownWinners.length > 0
      ? ["Winners:", ...shownWinners, hiddenText].filter(Boolean)
      : ["Winners: No position results recorded yet."];

    message = normalizeSmsMessageText([
      "PROVISIONAL RESULTS",
      electionName,
      `Turnout: ${turnout} (${payload.metrics.votedCount}/${payload.metrics.totalVoters})`,
      ...winnerLines,
      "Status: Subject to official declaration.",
    ].join("\n"));
    includedCount -= 1;
  } while (message.length > 620 && includedCount > 0);

  return message;
}

function getResultsSmsRecipients() {
  const rows = db.prepare(`
    SELECT
      id,
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName
    FROM voters
    WHERE phone_number IS NOT NULL
      AND TRIM(phone_number) <> ''
    ORDER BY staff_id COLLATE NOCASE ASC
  `).all();
  const seenNumbers = new Set();
  const recipients = [];
  let duplicateCount = 0;
  let invalidCount = 0;

  for (const row of rows) {
    const smsNumber = toArkeselOtpNumber(row.phoneNumber);

    if (!smsNumber) {
      invalidCount += 1;
      continue;
    }

    if (seenNumbers.has(smsNumber)) {
      duplicateCount += 1;
      continue;
    }

    seenNumbers.add(smsNumber);
    recipients.push({
      id: row.id,
      staffId: row.staffId,
      fullName: row.fullName,
      phoneNumber: row.phoneNumber,
      smsNumber,
      maskedPhoneNumber: maskPhoneNumber(row.phoneNumber),
    });
  }

  return {
    recipients,
    duplicateCount,
    invalidCount,
    totalWithPhone: rows.length,
  };
}

function getResultsSmsStatus(payload) {
  const settings = getAllSettings();
  const otpConfig = getOtpConfig();
  const recipientSnapshot = getResultsSmsRecipients();
  const currentFingerprint = buildResultsSmsFingerprint(payload);
  const lastFingerprint = String(settings.results_sms_last_fingerprint || "");
  const autoLastAttemptFingerprint = String(
    settings.results_sms_auto_last_attempt_fingerprint || "",
  );
  const alreadySentForCurrentResults = Boolean(
    lastFingerprint && lastFingerprint === currentFingerprint,
  );
  const autoAlreadyAttemptedForCurrentResults = Boolean(
    autoLastAttemptFingerprint && autoLastAttemptFingerprint === currentFingerprint,
  );

  return {
    providerLabel: "Arkesel",
    autoEnabled: String(settings.results_sms_auto_enabled || "false") === "true",
    autoLastAttemptAt: String(settings.results_sms_auto_last_attempt_at || ""),
    autoAlreadyAttemptedForCurrentResults,
    configured: otpConfig.arkeselConfigured,
    senderId: otpConfig.arkeselSenderId || "",
    recipientCount: recipientSnapshot.recipients.length,
    duplicateCount: recipientSnapshot.duplicateCount,
    invalidCount: recipientSnapshot.invalidCount,
    totalWithPhone: recipientSnapshot.totalWithPhone,
    messagePreview: buildResultsSmsMessage(payload),
    currentFingerprint,
    alreadySentForCurrentResults,
    canSend:
      payload.electionState.isClosed &&
      otpConfig.arkeselConfigured &&
      recipientSnapshot.recipients.length > 0,
    lastSentAt: String(settings.results_sms_last_sent_at || ""),
    lastSuccessCount: Number(settings.results_sms_last_success_count || 0),
    lastFailureCount: Number(settings.results_sms_last_failure_count || 0),
  };
}

async function sendResultsSmsBatch(recipients, message, otpConfig = getOtpConfig()) {
  const failedRecipients = [];
  let successCount = 0;

  for (const recipient of recipients) {
    try {
      await sendArkeselTextMessage(recipient.smsNumber, message, otpConfig);
      successCount += 1;
    } catch (error) {
      failedRecipients.push({
        staffId: recipient.staffId,
        phoneNumber: recipient.maskedPhoneNumber,
        error: error.message || "SMS failed",
      });
    }
  }

  return {
    successCount,
    failureCount: failedRecipients.length,
    failedRecipients,
  };
}

async function triggerAutomaticResultsSms(reason = "election_closed") {
  if (resultsSmsAutoSendInProgress) {
    return {
      skipped: true,
      reason: "already_running",
    };
  }

  const payload = getResultsExportPayload();
  const resultsSms = getResultsSmsStatus(payload);

  if (
    !payload.electionState.isClosed ||
    !resultsSms.autoEnabled ||
    !resultsSms.configured ||
    resultsSms.recipientCount === 0 ||
    resultsSms.alreadySentForCurrentResults ||
    resultsSms.autoAlreadyAttemptedForCurrentResults
  ) {
    return {
      skipped: true,
      reason: "not_ready",
    };
  }

  resultsSmsAutoSendInProgress = true;

  try {
    const recipientSnapshot = getResultsSmsRecipients();
    const message = buildResultsSmsMessage(payload);
    const attemptedAt = nowIso();

    setSetting("results_sms_auto_last_attempt_at", attemptedAt);
    setSetting("results_sms_auto_last_attempt_fingerprint", resultsSms.currentFingerprint);

    const sendResult = await sendResultsSmsBatch(
      recipientSnapshot.recipients,
      message,
      getOtpConfig(),
    );

    setSetting("results_sms_last_sent_at", attemptedAt);
    setSetting("results_sms_last_success_count", String(sendResult.successCount));
    setSetting("results_sms_last_failure_count", String(sendResult.failureCount));

    if (sendResult.successCount > 0) {
      setSetting("results_sms_last_fingerprint", resultsSms.currentFingerprint);
    }

    logSystemAudit("results_sms_sent", {
      automatic: true,
      reason,
      electionName: payload.settings.electionName,
      resultsFingerprint: resultsSms.currentFingerprint,
      recipientCount: recipientSnapshot.recipients.length,
      successCount: sendResult.successCount,
      failureCount: sendResult.failureCount,
      duplicateCount: recipientSnapshot.duplicateCount,
      invalidCount: recipientSnapshot.invalidCount,
      failedRecipients: sendResult.failedRecipients.slice(0, 5),
    });

    return sendResult;
  } catch (error) {
    logSystemAudit("results_sms_sent", {
      automatic: true,
      reason,
      successCount: 0,
      failureCount: resultsSms.recipientCount,
      error: error.message || "Automatic results SMS failed.",
    });
    return {
      successCount: 0,
      failureCount: resultsSms.recipientCount,
      failedRecipients: [],
    };
  } finally {
    resultsSmsAutoSendInProgress = false;
  }
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

function measurePdfBlockHeight(document, options) {
  const {
    label = "",
    value = "",
    width = 120,
    minHeight = 66,
    valueFontSize = 10.5,
  } = options || {};
  const innerWidth = Math.max(width - 28, 44);
  const normalizedValue = String(value || "-").trim() || "-";

  document.font("Helvetica-Bold").fontSize(8.5);
  const labelHeight = document.heightOfString(String(label || "").toUpperCase(), {
    width: innerWidth,
  });

  document.font("Helvetica").fontSize(valueFontSize);
  const valueHeight = document.heightOfString(normalizedValue, {
    width: innerWidth,
  });

  return Math.max(minHeight, 18 + labelHeight + 10 + valueHeight + 18);
}

function drawPdfBlockField(document, options) {
  const {
    x,
    y,
    width,
    label = "",
    value = "",
    height = null,
    minHeight = 66,
    valueFontSize = 10.5,
    fillColor = "#f8fbff",
    strokeColor = "#d7e4f3",
  } = options;
  const blockHeight =
    height ||
    measurePdfBlockHeight(document, {
      label,
      value,
      width,
      minHeight,
      valueFontSize,
    });
  const innerX = x + 14;
  const innerY = y + 12;
  const innerWidth = width - 28;
  const normalizedValue = String(value || "-").trim() || "-";

  document
    .roundedRect(x, y, width, blockHeight, 14)
    .fillAndStroke(fillColor, strokeColor);

  document
    .fillColor("#5d6d80")
    .font("Helvetica-Bold")
    .fontSize(8.5)
    .text(String(label || "").toUpperCase(), innerX, innerY, {
      width: innerWidth,
    });

  document
    .fillColor("#102338")
    .font("Helvetica")
    .fontSize(valueFontSize)
    .text(normalizedValue, innerX, innerY + 18, {
      width: innerWidth,
    });

  return blockHeight;
}

function drawPdfBlockRow(document, fields, options = {}) {
  const startX = document.page.margins.left;
  const totalWidth =
    document.page.width - document.page.margins.left - document.page.margins.right;
  const gap = options.gap || 16;
  const afterGap = options.afterGap || 12;
  const fieldWidth =
    fields.length > 1
      ? (totalWidth - gap * (fields.length - 1)) / fields.length
      : totalWidth;
  const rowHeight = Math.max(
    ...fields.map((field) =>
      measurePdfBlockHeight(document, {
        ...field,
        width: field.width || fieldWidth,
      }),
    ),
  );

  ensurePdfSpace(document, rowHeight + afterGap);
  const rowY = document.y;

  fields.forEach((field, index) => {
    const width = field.width || fieldWidth;
    const x = startX + index * (fieldWidth + gap);
    drawPdfBlockField(document, {
      ...field,
      x,
      y: rowY,
      width,
      height: rowHeight,
    });
  });

  document.y = rowY + rowHeight + afterGap;
}

function drawPdfWideTextBlock(document, options) {
  const {
    label = "",
    value = "",
    minHeight = 92,
    valueFontSize = 10.5,
  } = options || {};
  const startX = document.page.margins.left;
  const width =
    document.page.width - document.page.margins.left - document.page.margins.right;
  const height = measurePdfBlockHeight(document, {
    label,
    value,
    width,
    minHeight,
    valueFontSize,
  });

  ensurePdfSpace(document, height + 12);
  const startY = document.y;
  drawPdfBlockField(document, {
    x: startX,
    y: startY,
    width,
    label,
    value,
    height,
    minHeight,
    valueFontSize,
  });
  document.y = startY + height + 12;
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
  const { settings, positions, generatedAt, nomination = null } = payload;
  const logoPath = settings.organizationLogoPath
    ? resolveAssetPath(settings.organizationLogoPath)
    : "";
  const candidatePhotoPath = nomination?.photoPath
    ? resolveAssetPath(nomination.photoPath)
    : "";
  const positionName =
    nomination?.positionName ||
    positions.find((position) => position.id === nomination?.positionId)?.name ||
    "";

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
    .text(nomination ? "Submitted Nomination Form" : "Nomination Form", 120, 78);

  document
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#5d6d80")
    .text(`Generated: ${formatDateTime(generatedAt)}`, 120, 96)
    .text(`Nomination opens: ${formatDateTime(settings.nominationOpensAt)}`, 120, 112)
    .text(`Nomination closes: ${formatDateTime(settings.nominationClosesAt)}`, 120, 128);

  if (candidatePhotoPath && fs.existsSync(candidatePhotoPath)) {
    try {
      document.image(candidatePhotoPath, 430, 44, {
        fit: [82, 82],
        align: "right",
        valign: "top",
      });
    } catch (_error) {
      // Ignore image parsing errors and continue with the PDF.
    }
  }

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

  if (nomination) {
    document.y += 10;

    drawPdfBlockRow(document, [
      {
        label: "Application Number",
        value: nomination.applicationNumber || "-",
      },
      {
        label: "Current Status",
        value: nomination.statusMeta?.label || nomination.status || "Pending",
      },
    ]);

    drawPdfBlockRow(document, [
      {
        label: "Full Name",
        value: nomination.fullName || "-",
      },
      {
        label: "Position Applying For",
        value: positionName || "-",
      },
    ]);

    drawPdfBlockRow(document, [
      {
        label: "Staff ID",
        value: nomination.staffId || "-",
      },
      {
        label: "Phone Number",
        value: nomination.phoneNumber || "-",
      },
    ]);

    drawPdfBlockRow(document, [
      {
        label: "Department",
        value: nomination.department || "-",
      },
      {
        label: "Submitted At",
        value: formatDateTime(nomination.submittedAt),
      },
    ]);

    drawPdfBlockRow(document, [
      {
        label: "Proposer Name",
        value: nomination.proposerName || "-",
      },
      {
        label: "Seconder Name",
        value: nomination.seconderName || "-",
      },
    ]);

    drawPdfWideTextBlock(document, {
      label: "Short Profile / Bio",
      value: nomination.bio || "-",
      minHeight: 110,
    });

    drawPdfWideTextBlock(document, {
      label: "Manifesto / Message",
      value: nomination.manifesto || "-",
      minHeight: 140,
    });

    drawPdfWideTextBlock(document, {
      label: "Declaration",
      value: nomination.declarationAccepted
        ? "The applicant confirmed that all submitted details are accurate and personally submitted."
        : "Declaration was not confirmed at submission time.",
      minHeight: 78,
    });

    if (nomination.adminNotes) {
      drawPdfWideTextBlock(document, {
        label: "Committee Note",
        value: nomination.adminNotes,
        minHeight: 78,
      });
    }

    return;
  }

  document
    .moveDown(0.6)
    .font("Helvetica")
    .fontSize(11)
    .fillColor("#102338");

  const sections = [
    "Full Name: ________________________________________________",
    "Staff ID: ______________________",
    "Phone Number: ______________________",
    "Department: ______________________________________________",
    "Application Number (assigned after submission): __________________",
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

function renderNominationFlyerPdf(document, payload) {
  const { settings, nomination, generatedAt } = payload;
  const theme = getNominationFlyerTheme(nomination);
  const organizationName = getOrganizationDisplayName(settings.electionName);
  const logoPath = settings.organizationLogoPath
    ? resolveAssetPath(settings.organizationLogoPath)
    : "";
  const portraitSource = getNominationPortraitSource(nomination.photoPath);
  const fullNameLines = splitFlyerText(
    String(nomination.fullName || "").toUpperCase(),
    2,
    16,
  );
  const positionLines = splitFlyerText(
    String(nomination.positionName || "").toUpperCase(),
    3,
    13,
  );
  const canvasWidth = document.page.width;
  const canvasHeight = document.page.height;
  const accentShiftX = theme.seed % 46;
  const accentShiftY = Math.floor(theme.seed / 13) % 40;

  document.rect(0, 0, canvasWidth, canvasHeight).fill(theme.paper);
  document.save();
  document.opacity(0.18);
  document.circle(88 + accentShiftX, 84 + accentShiftY, 22 + (theme.seed % 10)).fill(
    theme.secondary,
  );
  document.circle(
    canvasWidth - 102 - Math.round(accentShiftX * 0.35),
    canvasHeight - 92,
    16 + (theme.seed % 8),
  ).fill(theme.accent);
  document.roundedRect(
    canvasWidth - 188,
    84 + Math.round(accentShiftY * 0.5),
    84 + (theme.seed % 24),
    12,
    6,
  ).fill(theme.secondary);
  document.restore();

  if (theme.layout === "arc") {
    document.save();
    document.circle(180, 170, 168).fill(theme.primary);
    document.circle(180, 170, 126).fill("#ffffff");
    document.restore();

    document.polygon([0, 0], [180, 0], [0, 190]).fill(theme.accent);
    document.polygon([canvasWidth, canvasHeight], [canvasWidth - 120, canvasHeight], [
      canvasWidth,
      canvasHeight - 150,
    ]).fill(theme.secondary);

    document.roundedRect(420, 250, 240, 220, 28).fill(theme.primary);
    document.roundedRect(128, 610, 210, 86, 20).fill(theme.soft);

    document.roundedRect(416, 122, 126, 34, 17).fill(theme.secondary);
    document
      .fillColor("#ffffff")
      .font("Helvetica-Bold")
      .fontSize(12)
      .text("NOMINATION FLYER", 416, 132, {
        width: 126,
        align: "center",
      });

    drawFlyerPortrait(document, portraitSource, {
      x: 28,
      y: 116,
      width: 344,
      height: 548,
      shape: "rounded",
      radius: 36,
      borderWidth: 8,
      borderColor: "#ffffff",
      panelColor: "#ffffff",
      imageInset: 4,
      fallbackLabel: getInitials(nomination.fullName),
      fallbackTextColor: theme.primary,
    });

    safeDrawImage(document, logoPath, 432, 60, {
      fit: [72, 72],
      align: "left",
    });

    document
      .fillColor(theme.dark)
      .font("Helvetica-Bold")
      .fontSize(18)
      .text(organizationName, 516, 66, {
        width: 180,
        align: "left",
      })
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#5d6d80")
      .text("Election Portal Nomination Poster", 516, 108, {
        width: 180,
        align: "left",
      });

    document
      .fillColor("#ffffff")
      .font("Helvetica")
      .fontSize(18)
      .text("ASPIRANT FOR", 452, 286, { width: 170 })
      .font("Helvetica-Bold")
      .fontSize(33)
      .text(positionLines.join("\n"), 452, 326, {
        width: 180,
        lineGap: -6,
      });

    document
      .fillColor(theme.primary)
      .font("Helvetica-Bold")
      .fontSize(fullNameLines.join("").length > 18 ? 28 : 34)
      .text(fullNameLines.join("\n"), 146, 625, {
        width: 174,
        align: "center",
        lineGap: -6,
      });
  } else if (theme.layout === "column") {
    document.rect(0, 0, canvasWidth, canvasHeight).fill(theme.primary);
    document.roundedRect(40, 36, 430, 688, 44).fill(theme.soft);
    document.roundedRect(492, 36, 228, 688, 36).fill(theme.paper);
    document.polygon([470, 36], [585, 36], [470, 160]).fill(theme.secondary);
    document.polygon([40, canvasHeight - 68], [216, canvasHeight - 68], [40, canvasHeight]).fill(
      theme.accent,
    );

    drawFlyerPortrait(document, portraitSource, {
      x: 335,
      y: 140,
      width: 318,
      height: 500,
      shape: "rounded",
      radius: 38,
      borderWidth: 8,
      borderColor: "#ffffff",
      panelColor: "#eff4ef",
      imageInset: 4,
      fallbackLabel: getInitials(nomination.fullName),
      fallbackTextColor: theme.primary,
    });
    safeDrawImage(document, logoPath, 78, 78, {
      fit: [76, 76],
      align: "left",
    });

    document
      .fillColor(theme.dark)
      .font("Helvetica-Bold")
      .fontSize(20)
      .text(organizationName, 166, 86, {
        width: 250,
      })
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#5f6974")
      .text("Official nomination poster", 166, 126, {
        width: 220,
      });

    document.roundedRect(78, 210, 286, 182, 30).fill(theme.primary);
    document
      .fillColor("#ffffff")
      .font("Helvetica")
      .fontSize(16)
      .text("CONTESTING FOR", 104, 244, { width: 236 })
      .font("Helvetica-Bold")
      .fontSize(31)
      .text(positionLines.join("\n"), 104, 282, {
        width: 236,
        lineGap: -6,
      });

    document.roundedRect(78, 430, 286, 164, 26).fill("#ffffff");
    document
      .fillColor(theme.primary)
      .font("Helvetica-Bold")
      .fontSize(fullNameLines.join("").length > 18 ? 30 : 36)
      .text(fullNameLines.join("\n"), 100, 468, {
        width: 240,
        align: "left",
        lineGap: -6,
      })
      .font("Helvetica")
      .fontSize(12)
      .fillColor("#5f6974")
      .text("Nomination application approved for review", 100, 552, {
        width: 220,
      });
  } else if (theme.layout === "ribbon") {
    document.rect(0, 0, canvasWidth, canvasHeight).fill(theme.paper);
    document.polygon([0, 0], [canvasWidth, 0], [canvasWidth, 122], [0, 188]).fill(
      theme.secondary,
    );
    document.roundedRect(52, 96, 656, 600, 48).fill(theme.soft);
    document.roundedRect(426, 214, 236, 264, 30).fill(theme.primary);
    document.roundedRect(104, 592, 320, 88, 22).fill(theme.secondary);
    document.polygon([0, canvasHeight - 150], [180, canvasHeight], [0, canvasHeight]).fill(
      theme.accent,
    );
    document.polygon(
      [canvasWidth, canvasHeight - 170],
      [canvasWidth - 150, canvasHeight],
      [canvasWidth, canvasHeight],
    ).fill(theme.secondary);

    drawFlyerPortrait(document, portraitSource, {
      x: 84,
      y: 160,
      width: 312,
      height: 438,
      shape: "rounded",
      radius: 34,
      borderWidth: 8,
      borderColor: "#ffffff",
      panelColor: "#fff7ef",
      imageInset: 4,
      fallbackLabel: getInitials(nomination.fullName),
      fallbackTextColor: theme.secondary,
    });
    safeDrawImage(document, logoPath, 94, 120, {
      fit: [66, 66],
      align: "left",
    });

    document
      .fillColor("#ffffff")
      .font("Helvetica-Bold")
      .fontSize(18)
      .text(organizationName, 176, 124, {
        width: 280,
      })
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#e7eef7")
      .text("Election nomination campaign sheet", 176, 160, {
        width: 250,
      });

    document
      .fillColor("#ffffff")
      .font("Helvetica")
      .fontSize(16)
      .text("NOMINEE FOR", 454, 252, { width: 180 })
      .font("Helvetica-Bold")
      .fontSize(31)
      .text(positionLines.join("\n"), 454, 292, {
        width: 170,
        lineGap: -6,
      });

    document
      .fillColor("#ffffff")
      .font("Helvetica-Bold")
      .fontSize(fullNameLines.join("").length > 18 ? 28 : 34)
      .text(fullNameLines.join("\n"), 124, 610, {
        width: 280,
        align: "center",
        lineGap: -6,
      });
  } else {
    document.rect(0, 0, canvasWidth, canvasHeight).fill(theme.paper);
    document.roundedRect(36, 34, 688, 692, 42).fill("#ffffff");
    document.circle(534, 194, 150).fill(theme.primary);
    document.circle(534, 194, 108).fill(theme.soft);
    document.roundedRect(78, 174, 286, 204, 32).fill(theme.primary);
    document.roundedRect(78, 454, 286, 166, 26).fill(theme.paper);
    document.polygon([566, 0], [canvasWidth, 0], [canvasWidth, 160]).fill(theme.accent);
    document.polygon([0, 548], [126, canvasHeight], [0, canvasHeight]).fill(theme.secondary);

    drawFlyerPortrait(document, portraitSource, {
      x: 380,
      y: 64,
      width: 306,
      height: 306,
      shape: "circle",
      borderWidth: 10,
      borderColor: "#ffffff",
      panelColor: "#f6f7fb",
      imageInset: 8,
      fallbackLabel: getInitials(nomination.fullName),
      fallbackTextColor: theme.primary,
    });
    safeDrawImage(document, logoPath, 88, 74, {
      fit: [68, 68],
      align: "left",
    });

    document
      .fillColor(theme.dark)
      .font("Helvetica-Bold")
      .fontSize(19)
      .text(organizationName, 168, 82, {
        width: 270,
      })
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#607086")
      .text("Nomination spotlight", 168, 122, {
        width: 180,
      });

    document
      .fillColor("#ffffff")
      .font("Helvetica")
      .fontSize(16)
      .text("POSITION", 102, 214, { width: 220 })
      .font("Helvetica-Bold")
      .fontSize(32)
      .text(positionLines.join("\n"), 102, 250, {
        width: 220,
        lineGap: -6,
      });

    document
      .fillColor(theme.dark)
      .font("Helvetica-Bold")
      .fontSize(fullNameLines.join("").length > 18 ? 30 : 36)
      .text(fullNameLines.join("\n"), 102, 492, {
        width: 236,
        lineGap: -6,
      })
      .font("Helvetica")
      .fontSize(12)
      .fillColor("#607086")
      .text("Presented through the official election portal", 102, 578, {
        width: 210,
      });
  }

  document
    .fillColor(theme.dark)
    .font("Helvetica")
    .fontSize(10)
    .text(`Application No. ${nomination.applicationNumber}`, 72, 714, {
      width: 210,
      align: "left",
    })
    .text(`Generated ${formatDateTime(generatedAt)}`, 470, 714, {
      width: 220,
      align: "right",
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
      .text(`Application Number: ${nomination.applicationNumber || "Not assigned"}`)
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

function renderPositionBallotPaperPdf(document, payload) {
  const { settings, position, generatedAt } = payload;
  const logoPath = settings.organizationLogoPath
    ? resolveAssetPath(settings.organizationLogoPath)
    : "";
  const pageWidth = document.page.width;
  const pageHeight = document.page.height;
  const marginX = 42;
  const marginY = 42;
  const contentWidth = pageWidth - marginX * 2;
  const tableGap = 3;
  const photoColumnWidth = 96;
  const markColumnWidth = 84;
  const nameColumnWidth = contentWidth - photoColumnWidth - markColumnWidth - tableGap * 2;
  const rowHeight = 90;
  const tableHeaderHeight = 40;
  const footerMargin = 46;

  const drawMasthead = () => {
    document
      .roundedRect(marginX, marginY, contentWidth, 112, 24)
      .fillColor("#eef2ec")
      .fill()
      .roundedRect(marginX, marginY, contentWidth, 112, 24)
      .lineWidth(1)
      .strokeColor("#b5c2cf")
      .stroke();

    if (logoPath && fs.existsSync(logoPath)) {
      safeDrawImage(document, logoPath, marginX + 16, marginY + 16, {
        fit: [66, 66],
        align: "left",
      });
    } else {
      document
        .roundedRect(marginX + 16, marginY + 16, 66, 66, 18)
        .fillColor("#ffffff")
        .fill()
        .roundedRect(marginX + 16, marginY + 16, 66, 66, 18)
        .lineWidth(1)
        .strokeColor("#d5dde7")
        .stroke()
        .font("Helvetica-Bold")
        .fontSize(20)
        .fillColor("#102338")
        .text("EC", marginX + 16, marginY + 38, {
          width: 66,
          align: "center",
        });
    }

    document
      .font("Helvetica-Bold")
      .fontSize(11)
      .fillColor("#5d6d80")
      .text("OFFICIAL BALLOT PAPER", marginX + 98, marginY + 18, {
        width: 220,
      })
      .fontSize(20)
      .fillColor("#102338")
      .text(settings.electionName, marginX + 98, marginY + 34, {
        width: 260,
      })
      .font("Helvetica")
      .fontSize(11)
      .fillColor("#5d6d80")
      .text(`Position: ${position.name}`, marginX + 98, marginY + 64, {
        width: 260,
      });

    document
      .roundedRect(pageWidth - marginX - 156, marginY + 16, 140, 66, 18)
      .fillColor("#ffffff")
      .fill()
      .roundedRect(pageWidth - marginX - 156, marginY + 16, 140, 66, 18)
      .lineWidth(1)
      .strokeColor("#d5dde7")
      .stroke()
      .font("Helvetica-Bold")
      .fontSize(11)
      .fillColor("#102338")
      .text("Vote For One Candidate", pageWidth - marginX - 146, marginY + 28, {
        width: 120,
        align: "center",
      })
      .font("Helvetica")
      .fontSize(9)
      .fillColor("#5d6d80")
      .text(`Generated: ${formatDateTime(generatedAt)}`, pageWidth - marginX - 146, marginY + 56, {
        width: 120,
        align: "center",
      });
  };

  const drawTableHeader = (y) => {
    document
      .rect(marginX, y, contentWidth, tableHeaderHeight)
      .fillColor("#2a3340")
      .fill();

    const nameColumnX = marginX + photoColumnWidth + tableGap;
    const markColumnX = marginX + photoColumnWidth + tableGap + nameColumnWidth + tableGap;

    document
      .font("Helvetica-Bold")
      .fontSize(9)
      .fillColor("#ffffff")
      .text("CANDIDATE PHOTO", marginX + 6, y + 15, {
        width: photoColumnWidth - 12,
        align: "center",
      })
      .text("CANDIDATE NAME", nameColumnX + 6, y + 15, {
        width: nameColumnWidth - 12,
        align: "center",
      })
      .text("MARK BOX", markColumnX + 6, y + 15, {
        width: markColumnWidth - 12,
        align: "center",
      });
  };

  const drawCandidateRow = (candidate, y) => {
    const nameColumnX = marginX + photoColumnWidth + tableGap;
    const markColumnX = marginX + photoColumnWidth + tableGap + nameColumnWidth + tableGap;

    document.rect(marginX, y, photoColumnWidth, rowHeight).fillColor("#ffffff").fill();
    document.rect(nameColumnX, y, nameColumnWidth, rowHeight).fillColor("#ffffff").fill();
    document.rect(markColumnX, y, markColumnWidth, rowHeight).fillColor("#ffffff").fill();

    document
      .rect(marginX, y, contentWidth, rowHeight)
      .lineWidth(1)
      .strokeColor("#2a3340")
      .stroke();
    document
      .moveTo(marginX + photoColumnWidth + 1.5, y)
      .lineTo(marginX + photoColumnWidth + 1.5, y + rowHeight)
      .strokeColor("#2a3340")
      .lineWidth(1)
      .stroke();
    document
      .moveTo(markColumnX - tableGap / 2, y)
      .lineTo(markColumnX - tableGap / 2, y + rowHeight)
      .strokeColor("#2a3340")
      .lineWidth(1)
      .stroke();

    const photoX = marginX + 11;
    const photoY = y + 9;
    if (candidate.photoPath) {
      const candidatePhotoPath = resolveAssetPath(candidate.photoPath);
      if (candidatePhotoPath && fs.existsSync(candidatePhotoPath)) {
        safeDrawImage(document, candidatePhotoPath, photoX, photoY, {
          fit: [74, 72],
          align: "center",
          valign: "center",
        });
      }
    } else {
      document
        .roundedRect(photoX, photoY, 74, 72, 12)
        .fillColor("#f1f5f9")
        .fill()
        .roundedRect(photoX, photoY, 74, 72, 12)
        .lineWidth(1)
        .strokeColor("#c9d4df")
        .stroke()
        .font("Helvetica-Bold")
        .fontSize(20)
        .fillColor("#102338")
        .text(getInitials(candidate.name), photoX, photoY + 22, {
          width: 74,
          align: "center",
        });
    }

    document
      .font("Helvetica-Bold")
      .fontSize(13)
      .fillColor("#102338")
      .text(String(candidate.name || "").toUpperCase(), nameColumnX + 14, y + 22, {
        width: nameColumnWidth - 28,
      })
      .font("Helvetica")
      .fontSize(9)
      .fillColor("#5d6d80")
      .text(position.name.toUpperCase(), nameColumnX + 14, y + 52, {
        width: nameColumnWidth - 28,
      });

    document
      .roundedRect(markColumnX + 14, y + 17, 50, 50, 10)
      .lineWidth(2)
      .strokeColor("#102338")
      .stroke();
  };

  drawMasthead();
  let cursorY = marginY + 130;
  drawTableHeader(cursorY);
  cursorY += tableHeaderHeight + tableGap;

  position.candidates.forEach((candidate, index) => {
    if (cursorY + rowHeight > pageHeight - footerMargin) {
      document.addPage();
      drawMasthead();
      cursorY = marginY + 130;
      drawTableHeader(cursorY);
      cursorY += tableHeaderHeight + tableGap;
    }

    drawCandidateRow(candidate, cursorY);
    cursorY += rowHeight + tableGap;

    if (index === position.candidates.length - 1) {
      document
        .font("Helvetica")
        .fontSize(9)
        .fillColor("#5d6d80")
        .text("Candidate order on this ballot matches the live voter ballot arrangement.", marginX, cursorY + 10, {
          width: contentWidth,
          align: "center",
        });
    }
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

function getVoterDirectoryStatus(voter, electionState = computeElectionState(getElectionSettings())) {
  if (voter.hasVoted) {
    return {
      label: "Voted",
      toneClass: "is-voted",
      filterValue: "voted",
    };
  }

  if (electionState.isClosed) {
    return {
      label: "Not Voted",
      toneClass: "is-not-voted",
      filterValue: "not-voted",
    };
  }

  return {
    label: "Verified",
    toneClass: "is-verified",
    filterValue: "verified",
  };
}

function getVoterManagementSummary(voters, electionState = computeElectionState(getElectionSettings())) {
  const totalVoters = voters.length;
  const votedCount = voters.filter((voter) => Boolean(voter.hasVoted)).length;
  const verifiedCount = voters.filter((voter) => Boolean(voter.phoneNumber)).length;
  const notVotedCount = Math.max(totalVoters - votedCount, 0);
  const pendingCount = voters.filter(
    (voter) => getVoterDirectoryStatus(voter, electionState).filterValue === "pending",
  ).length;

  return {
    totalVoters,
    verifiedCount,
    votedCount,
    notVotedCount,
    pendingCount,
    turnoutRate: totalVoters ? votedCount / totalVoters : 0,
  };
}

function getVoterManagementActivity(limit = 7) {
  const relevantActions = new Set([
    "voter_added_manually",
    "voters_imported",
    "voter_updated",
    "voters_cleared",
    "vote_submitted",
    "voter_login_success",
    "voter_login_rejected",
    "voter_login_failed",
    "voter_otp_sent",
    "voter_otp_verified",
  ]);

  return getAuditLogs(Math.max(limit * 8, 40))
    .filter((log) => relevantActions.has(log.action))
    .map((log) => {
      const details = safeJsonParse(log.detailsJson, {});
      return {
        ...log,
        actionLabel: formatDashboardAuditAction(log.action),
        detailLabel: summarizeAuditDetails(details),
        staffReference:
          details.staffId
          || details.voterId
          || log.actorIdentifier
          || (log.actorType === "admin" ? "ADMIN" : "SYSTEM"),
        timeAgo: formatDashboardRelativeTime(log.createdAt),
      };
    })
    .slice(0, limit);
}

async function buildVotersExportWorkbook(voters, settings, electionState) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Organization Vote Portal";
  workbook.lastModifiedBy = "Organization Vote Portal";
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet("Voters");
  worksheet.columns = [
    { header: "Staff ID", key: "staffId", width: 18 },
    { header: "Phone Number", key: "phoneNumber", width: 18 },
    { header: "Full Name", key: "fullName", width: 28 },
    { header: "Department", key: "department", width: 22 },
    { header: "Status", key: "status", width: 16 },
    { header: "Voted At", key: "votedAt", width: 24 },
    { header: "Created At", key: "createdAt", width: 24 },
  ];

  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  voters.forEach((voter) => {
    const status = getVoterDirectoryStatus(voter, electionState);
    worksheet.addRow({
      staffId: voter.staffId,
      phoneNumber: voter.phoneNumber,
      fullName: voter.fullName || "",
      department: voter.department || "",
      status: status.label,
      votedAt: voter.votedAt ? formatDateTime(voter.votedAt) : "",
      createdAt: voter.createdAt ? formatDateTime(voter.createdAt) : "",
    });
  });

  const summary = workbook.addWorksheet("Summary");
  const voterSummary = getVoterManagementSummary(voters, electionState);
  summary.columns = [
    { header: "Metric", key: "metric", width: 28 },
    { header: "Value", key: "value", width: 18 },
  ];
  summary.getRow(1).font = { bold: true };
  summary.addRows([
    { metric: "Election Name", value: settings.electionName },
    { metric: "Generated At", value: formatDateTime(nowIso()) },
    { metric: "Total Voters", value: voterSummary.totalVoters },
    { metric: "Verified Voters", value: voterSummary.verifiedCount },
    { metric: "Voted", value: voterSummary.votedCount },
    { metric: "Turnout", value: formatPercent(voterSummary.turnoutRate) },
  ]);

  return workbook;
}

function renderVotersPdf(document, { settings, voters, electionState, generatedAt }) {
  const marginX = 42;
  const contentWidth = document.page.width - marginX * 2;
  const pageBottom = document.page.height - 44;
  let cursorY = 46;

  const tableColumns = [
    { label: "Staff ID", width: 92 },
    { label: "Name", width: 164 },
    { label: "Phone", width: 106 },
    { label: "Department", width: 110 },
    { label: "Status", width: 72 },
  ];

  function ensurePageSpace(requiredHeight) {
    if (cursorY + requiredHeight <= pageBottom) {
      return;
    }

    document.addPage();
    cursorY = 46;
    drawTableHeader();
  }

  function drawTableHeader() {
    let x = marginX;
    document
      .font("Helvetica-Bold")
      .fontSize(9)
      .fillColor("#16315c");

    tableColumns.forEach((column) => {
      document.text(column.label, x, cursorY, {
        width: column.width,
      });
      x += column.width;
    });

    cursorY += 18;
    document
      .strokeColor("#d7e4fb")
      .lineWidth(1)
      .moveTo(marginX, cursorY)
      .lineTo(marginX + contentWidth, cursorY)
      .stroke();
    cursorY += 8;
  }

  const summary = getVoterManagementSummary(voters, electionState);

  document
    .font("Helvetica-Bold")
    .fontSize(22)
    .fillColor("#10254d")
    .text(settings.electionName || "Election Portal", marginX, cursorY, {
      width: contentWidth,
    });
  cursorY += 28;

  document
    .font("Helvetica-Bold")
    .fontSize(15)
    .fillColor("#ff3347")
    .text("Voter Management Export", marginX, cursorY, {
      width: contentWidth,
    });
  cursorY += 24;

  document
    .font("Helvetica")
    .fontSize(10)
    .fillColor("#5f7395")
    .text(`Generated: ${formatDateTime(generatedAt)}`, marginX, cursorY)
    .text(`Total Voters: ${summary.totalVoters}`, marginX + 200, cursorY)
    .text(`Turnout: ${formatPercent(summary.turnoutRate)}`, marginX + 340, cursorY);
  cursorY += 22;

  document
    .roundedRect(marginX, cursorY, contentWidth, 64, 18)
    .fillAndStroke("#f7faff", "#e3edff");

  document
    .fillColor("#10254d")
    .font("Helvetica-Bold")
    .fontSize(11)
    .text(`Verified Voters: ${summary.verifiedCount}`, marginX + 18, cursorY + 14)
    .text(`Votes Cast: ${summary.votedCount}`, marginX + 18, cursorY + 34)
    .text(`Not Voted: ${summary.notVotedCount}`, marginX + 214, cursorY + 14)
    .text(`Election State: ${electionState.badgeLabel}`, marginX + 214, cursorY + 34);

  cursorY += 86;
  drawTableHeader();

  voters.forEach((voter) => {
    ensurePageSpace(30);

    const status = getVoterDirectoryStatus(voter, electionState);
    const values = [
      voter.staffId,
      voter.fullName || "—",
      voter.phoneNumber,
      voter.department || "—",
      status.label,
    ];

    let x = marginX;
    document
      .font("Helvetica")
      .fontSize(9)
      .fillColor("#24385e");

    values.forEach((value, index) => {
      document.text(String(value), x, cursorY, {
        width: tableColumns[index].width - 10,
        ellipsis: true,
      });
      x += tableColumns[index].width;
    });

    cursorY += 18;
    document
      .strokeColor("#edf2fb")
      .lineWidth(1)
      .moveTo(marginX, cursorY)
      .lineTo(marginX + contentWidth, cursorY)
      .stroke();
    cursorY += 9;
  });
}

function getPublicVoterStatusByStaffId(staffId) {
  const normalizedStaffId = normalizeStaffId(staffId);
  if (!normalizedStaffId) {
    return null;
  }

  const voter = db.prepare(`
    SELECT
      staff_id AS staffId,
      phone_number AS phoneNumber,
      full_name AS fullName,
      department,
      has_voted AS hasVoted,
      voted_at AS votedAt
    FROM voters
    WHERE staff_id = ?
    LIMIT 1
  `).get(normalizedStaffId);

  if (!voter) {
    return null;
  }

  return {
    ...voter,
    maskedPhoneNumber: maskPhoneNumber(voter.phoneNumber),
    statusLabel: voter.hasVoted ? "Vote already cast" : "Ready to vote",
    statusToneClass: voter.hasVoted ? "is-voted" : "is-ready",
  };
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

function getBallotPositionById(positionId) {
  if (!positionId) {
    return null;
  }

  return getBallotData().find((position) => position.id === positionId) || null;
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

function getNominationAccessCodeStatusMeta(status) {
  const normalizedStatus = String(status || "unused").trim().toLowerCase();

  switch (normalizedStatus) {
    case "used":
      return {
        value: "used",
        label: "Used",
        className: "inline-badge--success",
      };
    case "cancelled":
      return {
        value: "cancelled",
        label: "Cancelled",
        className: "inline-badge--danger",
      };
    case "expired":
      return {
        value: "expired",
        label: "Expired",
        className: "inline-badge--warning",
      };
    default:
      return {
        value: "unused",
        label: "Unused",
        className: "inline-badge--muted",
      };
  }
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

function mapNominationAccessCodeRow(row) {
  return {
    ...row,
    statusMeta: getNominationAccessCodeStatusMeta(row.status),
  };
}

function getUnusedNominationAccessCodeCount() {
  const totals = db.prepare(`
    SELECT COUNT(*) AS total
    FROM nomination_access_codes
    WHERE status = 'unused'
  `).get();

  return Number(totals?.total || 0);
}

function getNominationAccessCodeMetrics() {
  const totals = db.prepare(`
    SELECT
      COUNT(*) AS total,
      SUM(CASE WHEN status = 'unused' THEN 1 ELSE 0 END) AS unusedCount,
      SUM(CASE WHEN status = 'used' THEN 1 ELSE 0 END) AS usedCount,
      SUM(CASE WHEN status = 'cancelled' THEN 1 ELSE 0 END) AS cancelledCount,
      SUM(CASE WHEN status = 'expired' THEN 1 ELSE 0 END) AS expiredCount
    FROM nomination_access_codes
  `).get();

  return {
    total: Number(totals?.total || 0),
    unusedCount: Number(totals?.unusedCount || 0),
    usedCount: Number(totals?.usedCount || 0),
    cancelledCount: Number(totals?.cancelledCount || 0),
    expiredCount: Number(totals?.expiredCount || 0),
  };
}

function getNominationAccessCodeByCode(referenceCode) {
  const normalizedCode = normalizeReferenceCode(referenceCode);

  if (!normalizedCode) {
    return null;
  }

  return db.prepare(`
    SELECT
      id,
      code,
      status,
      linked_nomination_id AS linkedNominationId,
      used_at AS usedAt,
      notes,
      created_at AS createdAt,
      updated_at AS updatedAt
    FROM nomination_access_codes
    WHERE code = ?
  `).get(normalizedCode);
}

function getNominationAccessCodeById(accessCodeId) {
  return db.prepare(`
    SELECT
      id,
      code,
      status,
      linked_nomination_id AS linkedNominationId,
      used_at AS usedAt,
      notes,
      created_at AS createdAt,
      updated_at AS updatedAt
    FROM nomination_access_codes
    WHERE id = ?
  `).get(accessCodeId);
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
      n.access_code_id AS accessCodeId,
      n.access_code AS accessCode,
      n.application_number AS applicationNumber,
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
      p.name AS positionName,
      ac.status AS accessCodeStatus
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    LEFT JOIN nomination_access_codes ac ON ac.id = n.access_code_id
    ORDER BY n.submitted_at DESC, n.id DESC
  `).all();

  return rows.map(mapNominationRow);
}

function getNominationForAccessCode(accessCodeId) {
  const row = db.prepare(`
    SELECT
      n.id,
      n.voter_id AS voterId,
      n.access_code_id AS accessCodeId,
      n.access_code AS accessCode,
      n.application_number AS applicationNumber,
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
      p.name AS positionName,
      ac.status AS accessCodeStatus
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    LEFT JOIN nomination_access_codes ac ON ac.id = n.access_code_id
    WHERE n.access_code_id = ?
  `).get(accessCodeId);

  return row ? mapNominationRow(row) : null;
}

function getNominationById(nominationId) {
  const row = db.prepare(`
    SELECT
      n.id,
      n.voter_id AS voterId,
      n.access_code_id AS accessCodeId,
      n.access_code AS accessCode,
      n.application_number AS applicationNumber,
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
      p.name AS positionName,
      ac.status AS accessCodeStatus
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    LEFT JOIN nomination_access_codes ac ON ac.id = n.access_code_id
    WHERE n.id = ?
  `).get(nominationId);

  return row ? mapNominationRow(row) : null;
}

function getNominationByApplicationNumber(applicationNumber) {
  const normalizedApplicationNumber = normalizeApplicationNumber(applicationNumber);

  if (!normalizedApplicationNumber) {
    return null;
  }

  const row = db.prepare(`
    SELECT
      n.id,
      n.voter_id AS voterId,
      n.access_code_id AS accessCodeId,
      n.access_code AS accessCode,
      n.application_number AS applicationNumber,
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
      p.name AS positionName,
      ac.status AS accessCodeStatus
    FROM nominations n
    INNER JOIN positions p ON p.id = n.position_id
    LEFT JOIN nomination_access_codes ac ON ac.id = n.access_code_id
    WHERE n.application_number = ?
  `).get(normalizedApplicationNumber);

  return row ? mapNominationRow(row) : null;
}

function getNominationAccessCodeList(limit = 250) {
  const rows = db.prepare(`
    SELECT
      ac.id,
      ac.code,
      ac.status,
      ac.linked_nomination_id AS linkedNominationId,
      ac.used_at AS usedAt,
      ac.notes,
      ac.created_at AS createdAt,
      ac.updated_at AS updatedAt,
      n.full_name AS nomineeName,
      n.staff_id AS staffId,
      n.position_id AS positionId,
      p.name AS positionName
    FROM nomination_access_codes ac
    LEFT JOIN nominations n ON n.id = ac.linked_nomination_id
    LEFT JOIN positions p ON p.id = n.position_id
    ORDER BY ac.created_at DESC, ac.id DESC
    LIMIT ?
  `).all(limit);

  return rows.map(mapNominationAccessCodeRow);
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

function getOtpActivityBadge(action) {
  switch (action) {
    case "voter_otp_sent":
    case "voter_otp_verified":
    case "otp_test_sent":
      return {
        label: action === "voter_otp_verified" ? "Verified" : "Sent",
        badgeClass: "status-open",
      };
    case "voter_otp_resent":
      return {
        label: "Resent",
        badgeClass: "status-scheduled",
      };
    case "voter_otp_failed":
      return {
        label: "Code Rejected",
        badgeClass: "status-closed",
      };
    case "voter_otp_send_failed":
    case "otp_test_send_failed":
      return {
        label: "Send Failed",
        badgeClass: "status-closed",
      };
    default:
      return {
        label: humanizeToken(action),
        badgeClass: "status-setup",
      };
  }
}

function buildOtpActivitySummary(action, details) {
  switch (action) {
    case "voter_otp_sent":
      return "Provider accepted a fresh OTP request for voter sign-in.";
    case "voter_otp_resent":
      return "A replacement OTP request was accepted after the resend action.";
    case "voter_otp_verified":
      return "The voter entered a valid OTP and passed verification.";
    case "voter_otp_failed":
      return details.message
        || (details.attempts
          ? `Incorrect or expired code. Failed attempt ${details.attempts}.`
          : "The submitted OTP code could not be verified.");
    case "voter_otp_send_failed":
    case "otp_test_send_failed":
      return details.message
        || (details.reason ? humanizeToken(details.reason) : "The OTP request was rejected before delivery.");
    case "otp_test_sent":
      return "Provider accepted the admin test OTP request.";
    default:
      return details.message || details.reason || humanizeToken(action);
  }
}

function mapOtpActivityLog(row) {
  const details = parseJsonObject(row.detailsJson, {});
  const badge = getOtpActivityBadge(row.action);
  const provider = String(details.provider || "").trim().toLowerCase();

  return {
    ...row,
    details,
    maskedPhoneNumber: String(details.phoneNumber || "").trim() || "Not recorded",
    provider,
    providerLabel: provider ? getOtpProviderLabel(provider) : "Not recorded",
    statusLabel: badge.label,
    badgeClass: badge.badgeClass,
    summary: buildOtpActivitySummary(row.action, details),
  };
}

function getOtpActivityLogs(limit = 150) {
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
    WHERE action IN (
      'voter_otp_sent',
      'voter_otp_resent',
      'voter_otp_send_failed',
      'voter_otp_failed',
      'voter_otp_verified',
      'otp_test_sent',
      'otp_test_send_failed'
    )
    ORDER BY created_at DESC
    LIMIT ?
  `).all(limit).map(mapOtpActivityLog);
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

    summary.candidates = summary.candidates
      .map((candidate) => ({
        ...candidate,
        shareRatio: summary.totalVotes ? candidate.voteCount / summary.totalVotes : 0,
        isLeading: highestVoteCount > 0 && candidate.voteCount === highestVoteCount,
        gapFromLead: Math.max(highestVoteCount - candidate.voteCount, 0),
      }))
      .sort((left, right) => {
        if (right.voteCount !== left.voteCount) {
          return right.voteCount - left.voteCount;
        }

        if (right.shareRatio !== left.shareRatio) {
          return right.shareRatio - left.shareRatio;
        }

        return left.name.localeCompare(right.name);
      })
      .map((candidate, index) => ({
        ...candidate,
        rank: index + 1,
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

function normalizeObserverId(value) {
  const normalized = String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");

  return /^[A-Z0-9-]{4,32}$/.test(normalized) ? normalized : "";
}

function generateObserverId() {
  const row = db.prepare("SELECT COALESCE(MAX(id), 0) + 1 AS nextSequence FROM observer_accounts").get();
  return `OBS-${new Date().getFullYear()}-${String(row.nextSequence || 1).padStart(4, "0")}`;
}

function generateObserverTemporaryPassword() {
  return `Ob!${crypto.randomBytes(8).toString("base64url")}7Z`;
}

function validateObserverPassword(password) {
  const value = String(password || "");
  const issues = [];

  if (value.length < 10) issues.push("at least 10 characters");
  if (!/[A-Z]/.test(value)) issues.push("one uppercase letter");
  if (!/[a-z]/.test(value)) issues.push("one lowercase letter");
  if (!/\d/.test(value)) issues.push("one number");
  if (!/[^A-Za-z0-9]/.test(value)) issues.push("one symbol");

  return {
    isValid: issues.length === 0,
    message: issues.length ? `Use ${issues.join(", ")}.` : "",
  };
}

function mapObserverAccount(row) {
  if (!row) return null;

  const accessExpired = row.accessExpiresAt
    ? dayjs(row.accessExpiresAt).isValid() && !dayjs().isBefore(dayjs(row.accessExpiresAt))
    : false;

  return {
    ...row,
    isActive: Boolean(row.isActive),
    mustChangePassword: Boolean(row.mustChangePassword),
    accessExpired,
    statusLabel: !row.isActive ? "Disabled" : accessExpired ? "Expired" : "Active",
    statusClass: !row.isActive ? "is-disabled" : accessExpired ? "is-expired" : "is-active",
  };
}

function getObserverAccounts() {
  return db.prepare(`
    SELECT
      oa.id,
      oa.observer_id AS observerId,
      oa.full_name AS fullName,
      oa.organization,
      oa.accreditation_number AS accreditationNumber,
      oa.email,
      oa.phone_number AS phoneNumber,
      oa.must_change_password AS mustChangePassword,
      oa.is_active AS isActive,
      oa.access_expires_at AS accessExpiresAt,
      oa.failed_login_attempts AS failedLoginAttempts,
      oa.locked_until AS lockedUntil,
      oa.last_login_at AS lastLoginAt,
      oa.password_changed_at AS passwordChangedAt,
      oa.created_by AS createdBy,
      oa.created_at AS createdAt,
      oa.updated_at AS updatedAt,
      COUNT(oi.id) AS incidentCount
    FROM observer_accounts oa
    LEFT JOIN observer_incidents oi ON oi.observer_account_id = oa.id
    GROUP BY oa.id
    ORDER BY oa.created_at DESC, oa.id DESC
  `).all().map(mapObserverAccount);
}

function getObserverAccountById(accountId) {
  const row = db.prepare(`
    SELECT
      id,
      observer_id AS observerId,
      full_name AS fullName,
      organization,
      accreditation_number AS accreditationNumber,
      email,
      phone_number AS phoneNumber,
      must_change_password AS mustChangePassword,
      is_active AS isActive,
      access_expires_at AS accessExpiresAt,
      failed_login_attempts AS failedLoginAttempts,
      locked_until AS lockedUntil,
      last_login_at AS lastLoginAt,
      password_changed_at AS passwordChangedAt,
      created_by AS createdBy,
      created_at AS createdAt,
      updated_at AS updatedAt
    FROM observer_accounts
    WHERE id = ?
  `).get(accountId);

  return mapObserverAccount(row);
}

function getObserverManagementSummary(accounts, incidents) {
  const todayStart = dayjs().startOf("day");
  return {
    total: accounts.length,
    active: accounts.filter((account) => account.isActive && !account.accessExpired).length,
    loggedInToday: accounts.filter(
      (account) => account.lastLoginAt && dayjs(account.lastLoginAt).isAfter(todayStart),
    ).length,
    reportsSubmitted: incidents.length,
  };
}

function getObserverIncidents({ accountId = null, limit = 100 } = {}) {
  const whereClause = accountId ? "WHERE oi.observer_account_id = ?" : "";
  const statement = db.prepare(`
    SELECT
      oi.id,
      oi.observer_account_id AS observerAccountId,
      oi.category,
      oi.title,
      oi.details,
      oi.status,
      oi.admin_notes AS adminNotes,
      oi.submitted_at AS submittedAt,
      oi.reviewed_at AS reviewedAt,
      oi.reviewed_by AS reviewedBy,
      oa.observer_id AS observerId,
      oa.full_name AS observerName,
      oa.organization
    FROM observer_incidents oi
    INNER JOIN observer_accounts oa ON oa.id = oi.observer_account_id
    ${whereClause}
    ORDER BY oi.submitted_at DESC, oi.id DESC
    LIMIT ?
  `);

  const rows = accountId ? statement.all(accountId, limit) : statement.all(limit);
  return rows.map((incident) => ({
    ...incident,
    statusLabel: humanizeToken(incident.status),
    statusClass: `is-${String(incident.status || "submitted").replace(/[^a-z0-9-]/gi, "-")}`,
  }));
}

function getObserverAccessLogs(limit = 20) {
  return db.prepare(`
    SELECT
      id,
      actor_identifier AS observerId,
      action,
      ip_address AS ipAddress,
      user_agent AS userAgent,
      created_at AS createdAt
    FROM audit_logs
    WHERE actor_type = 'observer'
    ORDER BY created_at DESC
    LIMIT ?
  `).all(limit).map((entry) => ({
    ...entry,
    actionLabel: formatDashboardAuditAction(entry.action),
    timeAgo: formatDashboardRelativeTime(entry.createdAt),
  }));
}

function getAnonymizedObserverActivity(limit = 7) {
  const activityLabels = {
    vote_submitted: ["Ballot recorded", "A verified ballot was securely added to the count."],
    voter_login_success: ["Voter verified", "A registered voter passed the access checks."],
    voter_otp_verified: ["OTP verified", "A voter completed phone verification."],
    election_opened: ["Election opened", "The election committee opened the voting window."],
    election_auto_opened: ["Election opened", "The scheduled voting window opened automatically."],
    election_closed: ["Election closed", "The ballot was locked by the election committee."],
    election_auto_closed: ["Election closed", "The scheduled voting window closed automatically."],
    database_backup_created: ["Backup completed", "A protected election database backup was created."],
  };

  return getAuditLogs(Math.max(limit * 12, 80))
    .filter((entry) => activityLabels[entry.action])
    .slice(0, limit)
    .map((entry, index) => ({
      id: entry.id,
      title: activityLabels[entry.action][0],
      description: activityLabels[entry.action][1],
      createdAt: entry.createdAt,
      timeAgo: formatDashboardRelativeTime(entry.createdAt),
      toneClass: `is-tone-${(index % 5) + 1}`,
    }));
}

function getObserverIntegrityChecks() {
  return [
    { label: "Voting application", status: "Operational" },
    { label: "Ballot storage", status: "Protected" },
    { label: "Audit logging", status: "Active" },
    { label: "Access control", status: "Enforced" },
    { label: "Results protection", status: "Locked while voting" },
  ];
}

function streamObserverReportPdf(res, reportData) {
  const {
    account,
    activity,
    electionState,
    incidents,
    metrics,
    results,
    settings,
  } = reportData;
  const filename = `${toSafeFilename(settings.electionName)}-observer-report-${toSafeFilename(account.observerId)}.pdf`;
  const document = new PDFDocument({ size: "A4", margin: 48, info: { Title: filename } });
  const pageWidth = document.page.width;
  const contentWidth = pageWidth - 96;
  const turnoutRate = metrics.totalVoters ? metrics.votedCount / metrics.totalVoters : 0;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  document.pipe(res);

  document
    .roundedRect(48, 42, contentWidth, 94, 12)
    .fill("#082957")
    .fillColor("#ffffff")
    .font("Helvetica-Bold")
    .fontSize(18)
    .text("ELECTION OBSERVER REPORT", 68, 64)
    .font("Helvetica")
    .fontSize(11)
    .text(settings.electionName, 68, 91)
    .text(`Generated ${formatDateTime(nowIso())}`, 68, 108);

  let cursorY = 158;
  document
    .fillColor("#102b4f")
    .font("Helvetica-Bold")
    .fontSize(13)
    .text("Observer Accreditation", 48, cursorY);
  cursorY += 22;
  document.font("Helvetica").fontSize(10).fillColor("#40556f");
  [
    `Observer: ${account.fullName} (${account.observerId})`,
    `Organization: ${account.organization}`,
    `Accreditation: ${account.accreditationNumber || "Not provided"}`,
    `Election status: ${electionState.badgeLabel}`,
  ].forEach((line) => {
    document.text(line, 56, cursorY);
    cursorY += 17;
  });

  cursorY += 10;
  document.font("Helvetica-Bold").fontSize(13).fillColor("#102b4f").text("Turnout Summary", 48, cursorY);
  cursorY += 24;
  const metricItems = [
    ["Registered voters", metrics.totalVoters],
    ["Votes cast", metrics.votedCount],
    ["Remaining voters", Math.max(metrics.totalVoters - metrics.votedCount, 0)],
    ["Turnout", formatPercent(turnoutRate)],
    ["Active positions", metrics.totalPositions],
  ];
  metricItems.forEach(([label, value], index) => {
    const columnWidth = contentWidth / metricItems.length;
    const x = 48 + columnWidth * index;
    document.roundedRect(x + 2, cursorY, columnWidth - 6, 54, 7).fill(index % 2 ? "#edf4ff" : "#f5f8fb");
    document.font("Helvetica-Bold").fontSize(12).fillColor("#0d2850").text(String(value), x + 10, cursorY + 12, { width: columnWidth - 20, align: "center" });
    document.font("Helvetica").fontSize(7.5).fillColor("#62758c").text(label, x + 7, cursorY + 32, { width: columnWidth - 14, align: "center" });
  });
  cursorY += 75;

  document.font("Helvetica-Bold").fontSize(13).fillColor("#102b4f").text("Observer Incident Register", 48, cursorY);
  cursorY += 22;
  if (!incidents.length) {
    document.font("Helvetica").fontSize(10).fillColor("#62758c").text("No incidents have been submitted by this observer.", 56, cursorY);
    cursorY += 22;
  } else {
    incidents.slice(0, 8).forEach((incident) => {
      document.font("Helvetica-Bold").fontSize(9).fillColor("#173250").text(`${incident.title} - ${incident.statusLabel}`, 56, cursorY, { width: contentWidth - 16 });
      cursorY += 13;
      document.font("Helvetica").fontSize(8).fillColor("#62758c").text(`${formatDateTime(incident.submittedAt)} | ${humanizeToken(incident.category)}`, 56, cursorY);
      cursorY += 17;
    });
  }

  document.font("Helvetica-Bold").fontSize(13).fillColor("#102b4f").text("Anonymized Election Activity", 48, cursorY);
  cursorY += 22;
  activity.slice(0, 8).forEach((item) => {
    document.font("Helvetica-Bold").fontSize(9).fillColor("#173250").text(item.title, 56, cursorY);
    document.font("Helvetica").fillColor("#62758c").text(formatDateTime(item.createdAt), 360, cursorY, { width: 150, align: "right" });
    cursorY += 15;
  });

  if (cursorY > 660) {
    document.addPage();
    cursorY = 52;
  } else {
    cursorY += 14;
  }

  document.font("Helvetica-Bold").fontSize(13).fillColor("#102b4f").text("Election Results", 48, cursorY);
  cursorY += 22;
  if (!electionState.isClosed) {
    document.font("Helvetica-Bold").fontSize(11).fillColor("#d6354d").text("Results locked until polls close.", 56, cursorY);
  } else if (!results.length) {
    document.font("Helvetica").fontSize(10).fillColor("#62758c").text("No result data is available.", 56, cursorY);
  } else {
    results.forEach((position) => {
      if (cursorY > 735) {
        document.addPage();
        cursorY = 52;
      }
      document.font("Helvetica-Bold").fontSize(10).fillColor("#102b4f").text(position.name, 56, cursorY);
      cursorY += 15;
      position.candidates.forEach((candidate) => {
        document.font("Helvetica").fontSize(9).fillColor("#40556f").text(candidate.name, 68, cursorY);
        document.text(`${candidate.voteCount} votes (${formatPercent(candidate.shareRatio)})`, 360, cursorY, { width: 140, align: "right" });
        cursorY += 13;
      });
      cursorY += 8;
    });
  }

  document
    .font("Helvetica")
    .fontSize(8)
    .fillColor("#718096")
    .text("This report contains aggregated election information only. No individual voting choices are disclosed.", 48, document.page.height - 48, { width: contentWidth, align: "center" });
  document.end();
}

function streamObserverManagementPdf(res, { accounts, incidents, settings }) {
  const filename = `${toSafeFilename(settings.electionName)}-observer-register.pdf`;
  const document = new PDFDocument({ size: "A4", layout: "landscape", margin: 42 });
  const pageWidth = document.page.width;
  const contentWidth = pageWidth - 84;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  document.pipe(res);

  document
    .roundedRect(42, 36, contentWidth, 72, 10)
    .fill("#082957")
    .fillColor("#ffffff")
    .font("Helvetica-Bold")
    .fontSize(18)
    .text("OBSERVER ACCESS REGISTER", 60, 54)
    .font("Helvetica")
    .fontSize(10)
    .text(`${settings.electionName} | Generated ${formatDateTime(nowIso())}`, 60, 80);

  let cursorY = 128;
  const headers = ["Observer ID", "Name", "Organization", "Accreditation", "Status", "Last Login", "Incidents"];
  const widths = [92, 126, 142, 104, 72, 126, 58];
  let cursorX = 42;
  document.rect(42, cursorY, contentWidth, 24).fill("#eaf1fb");
  headers.forEach((header, index) => {
    document.font("Helvetica-Bold").fontSize(8).fillColor("#123052").text(header, cursorX + 5, cursorY + 8, { width: widths[index] - 10 });
    cursorX += widths[index];
  });
  cursorY += 24;

  accounts.forEach((account, rowIndex) => {
    if (cursorY > document.page.height - 72) {
      document.addPage();
      cursorY = 48;
    }
    if (rowIndex % 2 === 0) document.rect(42, cursorY, contentWidth, 28).fill("#f8fafc");
    const values = [
      account.observerId,
      account.fullName,
      account.organization,
      account.accreditationNumber || "-",
      account.statusLabel,
      account.lastLoginAt ? formatDateTime(account.lastLoginAt) : "Never",
      String(account.incidentCount || 0),
    ];
    cursorX = 42;
    values.forEach((value, index) => {
      document.font("Helvetica").fontSize(7.5).fillColor("#40556f").text(String(value), cursorX + 5, cursorY + 8, { width: widths[index] - 10, ellipsis: true });
      cursorX += widths[index];
    });
    cursorY += 28;
  });

  cursorY += 20;
  if (cursorY > document.page.height - 170) {
    document.addPage();
    cursorY = 48;
  }
  document.font("Helvetica-Bold").fontSize(13).fillColor("#102b4f").text("Incident Summary", 42, cursorY);
  cursorY += 22;
  if (!incidents.length) {
    document.font("Helvetica").fontSize(9).fillColor("#62758c").text("No observer incidents have been submitted.", 50, cursorY);
  } else {
    incidents.slice(0, 20).forEach((incident) => {
      if (cursorY > document.page.height - 54) {
        document.addPage();
        cursorY = 48;
      }
      document.font("Helvetica-Bold").fontSize(8.5).fillColor("#173250").text(`${incident.observerId} | ${incident.title}`, 50, cursorY, { width: 460 });
      document.font("Helvetica").fontSize(8).fillColor("#62758c").text(`${incident.statusLabel} | ${formatDateTime(incident.submittedAt)}`, 520, cursorY, { width: 230, align: "right" });
      cursorY += 16;
    });
  }

  document.end();
}

function ensureSetupMode(req, res) {
  const settings = getElectionSettings();

  if (settings.phase !== "setup") {
    setFlash(
      req,
      "error",
      "Setup is locked once voting has been opened. Create a new election cycle to make structural changes.",
    );
    res.redirect(getAdminReturnPath(req, "/admin"));
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
app.use(express.json({ limit: "32kb" }));
app.use("/uploads", express.static(uploadsRootDirectory));
app.use(express.static(publicDirectory, { redirect: false }));
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
  void triggerAutomaticResultsSms("closed_election_check");

  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const nominationState = computeNominationState(settings);

  res.locals.settings = settings;
  res.locals.electionState = electionState;
  res.locals.nominationState = nominationState;
  res.locals.currentPath = req.path;
  res.locals.admin = req.session.admin || null;
  res.locals.observer = req.session.observer || null;
  res.locals.voter = req.session.voter || null;
  res.locals.nominationApplicant = req.session.nominationApplicant || null;
  res.locals.adminNotificationUnreadCount = req.session.admin
    ? getUnreadAdminNotificationCount()
    : 0;
  res.locals.captcha = getCaptchaPublicConfig();
  res.locals.currentYear = new Date().getFullYear();
  res.locals.formatDateTime = formatDateTime;
  res.locals.formatPercent = formatPercent;
  res.locals.voteClosesAtMs = settings.closesAt ? dayjs(settings.closesAt).valueOf() : null;
  res.locals.nominationOpensAtMs = settings.nominationOpensAt
    ? dayjs(settings.nominationOpensAt).valueOf()
    : null;
  res.locals.getInitials = getInitials;
  res.locals.flash = req.session.flash || null;
  delete req.session.flash;

  next();
});

app.get("/", (req, res) => {
  const metrics = getDashboardMetrics();
  const turnoutRate = metrics.totalVoters ? metrics.votedCount / metrics.totalVoters : 0;
  const requestedStaffId = String(req.query.voterStatusStaffId || "").trim();
  const normalizedStaffId = normalizeStaffId(requestedStaffId);
  let voterStatusLookup = null;

  if (requestedStaffId) {
    if (!normalizedStaffId) {
      voterStatusLookup = {
        staffId: requestedStaffId,
        found: false,
        invalid: true,
        message: "Enter a valid staff ID to check voter details.",
      };
    } else {
      const voterRecord = getPublicVoterStatusByStaffId(normalizedStaffId);
      voterStatusLookup = voterRecord
        ? {
            found: true,
            invalid: false,
            ...voterRecord,
          }
        : {
            staffId: normalizedStaffId,
            found: false,
            invalid: false,
            message: "No registered voter record was found for that staff ID.",
          };
    }
  }

  res.render("home", {
    pageTitle: "Election Portal",
    metrics,
    turnoutRate,
    voterStatusLookup,
  });
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

app.get("/nomination/status/pdf", (req, res) => {
  const applicant = req.session.nominationApplicant;

  if (!applicant) {
    setFlash(req, "error", "Enter your application number to download your submitted nomination form.");
    return res.redirect("/nomination/status/login");
  }

  const nomination =
    getNominationById(applicant.nominationId || 0) ||
    getNominationByApplicationNumber(applicant.applicationNumber);

  if (!nomination) {
    clearNominationSession(req);
    setFlash(req, "error", "That nomination application could not be found.");
    return res.redirect("/nomination/status/login");
  }

  const settings = getElectionSettings();
  const positions = getPositions();
  const filename = `${toSafeFilename(settings.electionName)}-${toSafeFilename(
    nomination.applicationNumber || nomination.fullName || "nomination",
  )}.pdf`;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: "A4",
    margin: 50,
    info: {
      Title: `${settings.electionName} Submitted Nomination Form`,
      Author: "Organization Vote Portal",
      Subject: nomination.applicationNumber || "Nomination Application",
    },
  });

  logAudit(req, "nomination", nomination.applicationNumber, "nomination_form_downloaded", {
    nominationId: nomination.id,
  });

  document.pipe(res);
  renderNominationFormPdf(document, {
    settings,
    positions,
    nomination,
    generatedAt: nowIso(),
  });
  document.end();
});

app.get("/nomination/status/flyer", (req, res) => {
  const applicant = req.session.nominationApplicant;

  if (!applicant) {
    setFlash(req, "error", "Enter your application number to download your campaign flyer.");
    return res.redirect("/nomination/status/login");
  }

  const nomination =
    getNominationById(applicant.nominationId || 0) ||
    getNominationByApplicationNumber(applicant.applicationNumber);

  if (!nomination) {
    clearNominationSession(req);
    setFlash(req, "error", "That nomination application could not be found.");
    return res.redirect("/nomination/status/login");
  }

  const settings = getElectionSettings();
  const theme = getNominationFlyerTheme(nomination);
  const filename = `${toSafeFilename(settings.electionName)}-${toSafeFilename(
    nomination.applicationNumber || nomination.fullName || "nomination",
  )}-flyer.pdf`;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: [760, 760],
    margin: 0,
    info: {
      Title: `${settings.electionName} Campaign Flyer`,
      Author: "Organization Vote Portal",
      Subject: `${nomination.fullName} campaign flyer`,
      Keywords: `nomination flyer,${nomination.positionName},${nomination.applicationNumber}`,
    },
  });

  logAudit(req, "nomination", nomination.applicationNumber, "nomination_flyer_downloaded", {
    nominationId: nomination.id,
    theme: theme.key,
  });

  document.pipe(res);
  renderNominationFlyerPdf(document, {
    settings,
    nomination,
    generatedAt: nowIso(),
  });
  document.end();
});

app.get("/nomination/login", (req, res) => {
  return res.redirect("/nomination/form");
});

app.post("/nomination/login", (req, res) => {
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

app.post("/nomination/status/login", requireCaptcha("nomination"), (req, res) => {
  const applicationNumber = normalizeApplicationNumber(req.body.applicationNumber);

  if (!applicationNumber) {
    setFlash(req, "error", "Enter your application number to check status.");
    return res.redirect("/nomination/status/login");
  }

  const nomination = getNominationByApplicationNumber(applicationNumber);

  if (!nomination) {
    logAudit(req, "nomination", applicationNumber || "unknown", "nomination_status_login_failed", {
      reason: "application_not_found",
    });
    setFlash(
      req,
      "error",
      "That application number was not found. Check it and try again.",
    );
    return res.redirect("/nomination/status/login");
  }

  beginNominationApplicantSession(req, nomination);
  logAudit(req, "nomination", nomination.applicationNumber, "nomination_status_login_success", {
    nominationId: nomination.id,
  });
  return res.redirect("/nomination/status");
});

app.get("/nomination/form", (req, res) => {
  const settings = getElectionSettings();
  const nominationState = computeNominationState(settings);
  const positions = getPositions();
  const applicant = req.session.nominationApplicant || null;
  const editNominationId = parseInteger(req.query.edit, 0);
  const editableNomination =
    editNominationId > 0 &&
    applicant?.nominationId === editNominationId
      ? getNominationById(editNominationId)
      : null;

  if (editNominationId > 0 && (!editableNomination || !editableNomination.statusMeta.canApplicantEdit)) {
    setFlash(
      req,
      "error",
      "Enter your application number first before continuing a correction request.",
    );
    return res.redirect("/nomination/status/login");
  }

  if (editNominationId > 0 && !nominationState.isOpen) {
    setFlash(req, "error", nominationState.message);
    return res.redirect("/nomination/status/login");
  }

  if (!nominationState.isOpen && !nominationState.isScheduled) {
    setFlash(req, "error", nominationState.message);
    return res.redirect("/nomination/status/login");
  }

  return res.render("nomination-form", {
    pageTitle: "Apply for Nomination",
    applicant,
    positions,
    currentNomination: editableNomination,
    editableNomination,
    formLocked: !editableNomination && nominationState.isScheduled,
  });
});

app.post(
  "/nomination/form",
  nominationUpload.single("photo"),
  async (req, res) => {
    const applicant = req.session.nominationApplicant || null;
    const settings = getElectionSettings();
    const nominationState = computeNominationState(settings);
    const nominationId = parseInteger(req.body.nominationId, 0);
    const fullName = String(req.body.fullName || "").trim();
    const staffId = normalizeStaffId(req.body.staffId);
    const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
    const department = String(req.body.department || "").trim();
    const positionId = parseInteger(req.body.positionId, 0);
    const bio = String(req.body.bio || "").trim();
    const manifesto = String(req.body.manifesto || "").trim();
    const proposerName = String(req.body.proposerName || "").trim();
    const seconderName = String(req.body.seconderName || "").trim();
    const declarationAccepted = req.body.declarationAccepted === "on";
    const existingEditableNomination =
      nominationId > 0 && applicant?.nominationId === nominationId
        ? getNominationById(nominationId)
        : null;
    const isCorrectionResubmission =
      existingEditableNomination &&
      existingEditableNomination.statusMeta.canApplicantEdit;

    const captchaResult = await verifyCaptchaSubmission(req, "nomination");
    if (!captchaResult.ok) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      logAudit(req, "security", "nomination", "captcha_verification_failed", {
        reason: captchaResult.reason,
        errorCodes: captchaResult.errorCodes || [],
      });
      setFlash(req, "error", "Please complete the security check before submitting the nomination.");
      return res.redirect(
        isCorrectionResubmission
          ? `/nomination/form?edit=${existingEditableNomination.id}`
          : "/nomination/form",
      );
    }

    if (nominationId > 0 && !isCorrectionResubmission) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "Enter your application number first before continuing a correction request.",
      );
      return res.redirect("/nomination/status/login");
    }

    if (!nominationState.isOpen) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(req, "error", nominationState.message);
      return res.redirect(
        isCorrectionResubmission ? `/nomination/form?edit=${existingEditableNomination.id}` : "/nomination/form",
      );
    }

    if (
      !fullName ||
      !staffId ||
      !phoneNumber ||
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
        "Complete every nomination field, including your staff details, before submitting.",
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
      WHERE staff_id = ?
        AND position_id = ?
        AND id <> ?
    `).get(staffId, positionId, isCorrectionResubmission ? existingEditableNomination.id : 0);

    if (duplicateNomination) {
      if (req.file) {
        await safeRemoveFile(req.file.path);
      }
      setFlash(
        req,
        "error",
        "You already have a nomination record for this position. Use the nomination status page to review it.",
      );
      return res.redirect("/nomination/status/login");
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
      const applicationNumber =
        existingEditableNomination?.applicationNumber || createUniqueNominationApplicationNumber();

      if (isCorrectionResubmission) {
        db.prepare(`
          UPDATE nominations
          SET
            access_code_id = NULL,
            access_code = '',
            application_number = ?,
            position_id = ?,
            staff_id = ?,
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
          applicationNumber,
          positionId,
          staffId,
          fullName,
          phoneNumber,
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

        logAudit(req, "nomination", applicationNumber, "nomination_resubmitted", {
          applicationNumber,
          staffId,
          nominationId: existingEditableNomination.id,
          positionName: position.name,
        });
      } else {
        const insertResult = db.prepare(`
          INSERT INTO nominations (
            voter_id,
            access_code_id,
            access_code,
            application_number,
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
          VALUES (NULL, NULL, '', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'pending', '', ?, ?, ?)
        `).run(
          applicationNumber,
          positionId,
          staffId,
          fullName,
          phoneNumber,
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

        const createdNominationId = Number(insertResult.lastInsertRowid);

        req.session.nominationApplicant = {
          nominationId: createdNominationId,
          applicationNumber,
        };

        logAudit(req, "nomination", applicationNumber, "nomination_submitted", {
          applicationNumber,
          staffId,
          nominationId: createdNominationId,
          positionName: position.name,
        });
      }

      req.session.nominationApplicant = {
        nominationId: isCorrectionResubmission
          ? existingEditableNomination.id
          : req.session.nominationApplicant.nominationId,
        applicationNumber,
      };

      setFlash(
        req,
        "success",
        isCorrectionResubmission
          ? "Your nomination has been resubmitted for review."
          : `Your nomination has been submitted successfully. Save application number ${applicationNumber} to check your status and download your campaign flyer.`,
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
    setFlash(req, "error", "Enter your application number to check nomination status.");
    return res.redirect("/nomination/status/login");
  }

  const nomination =
    getNominationById(applicant.nominationId || 0) ||
    getNominationByApplicationNumber(applicant.applicationNumber);

  if (!nomination) {
    clearNominationSession(req);
    setFlash(req, "error", "That nomination application could not be found.");
    return res.redirect("/nomination/status/login");
  }

  req.session.nominationApplicant = {
    nominationId: nomination.id,
    applicationNumber: nomination.applicationNumber,
  };

  return res.render("nomination-status", {
    pageTitle: "Nomination Status",
    applicant: req.session.nominationApplicant,
    nomination,
  });
});

app.post("/nomination/logout", (req, res) => {
  clearNominationSession(req);
  return res.redirect("/nomination/status/login");
});

app.get("/demo/how-to-vote", (req, res) => {
  const liveBallot = getBallotData().filter((position) => position.candidates.length > 0);
  const demoBallot = liveBallot.length > 0
    ? liveBallot
    : [
        {
          id: "demo-president",
          name: "President",
          candidates: [
            { id: "demo-ama", name: "Ama Boateng", photoPath: "" },
            { id: "demo-john", name: "John Mensah", photoPath: "" },
          ],
        },
        {
          id: "demo-secretary",
          name: "Secretary",
          candidates: [
            { id: "demo-mary", name: "Mary Owusu", photoPath: "" },
            { id: "demo-daniel", name: "Daniel Appiah", photoPath: "" },
          ],
        },
        {
          id: "demo-treasurer",
          name: "Treasurer",
          candidates: [
            { id: "demo-kofi", name: "Kofi Asare", photoPath: "" },
            { id: "demo-akosua", name: "Akosua Mensah", photoPath: "" },
          ],
        },
      ];

  return res.render("vote-demo", {
    pageTitle: "How To Vote Demo",
    demoBallot,
  });
});

app.get("/vote/login", (req, res) => {
  if (req.session.voter) {
    return res.redirect("/vote");
  }

  if (req.session.pendingVoterVerification && isOtpVerificationEnabled()) {
    return res.redirect("/vote/verify-otp");
  }

  const otpEnabled = isOtpVerificationEnabled();
  return res.render("vote-login", {
    pageTitle: "Voter Login",
    otpEnabled,
  });
});

app.post("/vote/login", requireCaptcha("voter_login"), async (req, res) => {
  const staffId = normalizeStaffId(req.body.staffId);
  const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const otpConfig = getOtpConfig();

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

  if (isOtpVerificationEnabled(otpConfig)) {
    const smsPhoneNumber = toSmsPhoneNumber(phoneNumber);

    if (!smsPhoneNumber) {
      logAudit(req, "voter", staffId, "voter_otp_send_failed", {
        provider: otpConfig.provider,
        phoneNumber: maskPhoneNumber(phoneNumber),
        phase: "initial",
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
      const challenge = await sendOtpChallenge(smsPhoneNumber, otpConfig);
      clearVoterSession(req);
      req.session.voteComplete = null;
      req.session.pendingVoterVerification = buildPendingVoterVerification(
        voterRecord,
        phoneNumber,
        smsPhoneNumber,
        challenge,
        otpConfig,
      );

      logAudit(req, "voter", staffId, "voter_otp_sent", {
        provider: otpConfig.provider,
        phoneNumber: maskPhoneNumber(phoneNumber),
        phase: "initial",
      });
      setFlash(
        req,
        "success",
        `A one-time OTP code has been sent to ${maskPhoneNumber(phoneNumber)}.`,
      );
      return res.redirect("/vote/verify-otp");
    } catch (error) {
      logAudit(req, "voter", staffId, "voter_otp_send_failed", {
        provider: otpConfig.provider,
        phoneNumber: maskPhoneNumber(phoneNumber),
        phase: "initial",
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
  const otpLifetimeMinutes = Math.max(
    Math.ceil(expiresAt?.diff(pendingVerification.sentAt ? dayjs(pendingVerification.sentAt) : dayjs(), "minute", true) || 0),
    1,
  );

  return res.render("vote-verify-otp", {
    pageTitle: "Verify OTP",
    maskedPhoneNumber: pendingVerification.maskedPhoneNumber,
    otpTtlMinutes: otpLifetimeMinutes,
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
        provider: pendingVerification.provider || getOtpConfig().provider,
        phoneNumber: maskPhoneNumber(voterRecord.phoneNumber),
        message: verification.errorMessage,
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
    provider: pendingVerification.provider || getOtpConfig().provider,
  });
  logAudit(req, "voter", voterRecord.staffId, "voter_login_success", {
    otpProvider: pendingVerification.provider || getOtpConfig().provider,
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
  const otpConfig = getOtpConfig();

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
    logAudit(req, "voter", voterRecord.staffId, "voter_otp_send_failed", {
      provider: otpConfig.provider,
      phoneNumber: maskPhoneNumber(voterRecord.phoneNumber),
      phase: "resend",
      reason: "invalid_sms_phone_format",
    });
    clearVoterSession(req);
    setFlash(
      req,
      "error",
      "Your phone number is registered, but it is not in a valid SMS format for OTP delivery. Please contact the election committee.",
    );
    return res.redirect("/vote/login");
  }

  try {
    const challenge = await sendOtpChallenge(smsPhoneNumber, otpConfig);
    req.session.pendingVoterVerification = buildPendingVoterVerification(
      voterRecord,
      voterRecord.phoneNumber,
      smsPhoneNumber,
      challenge,
      otpConfig,
    );

    logAudit(req, "voter", voterRecord.staffId, "voter_otp_resent", {
      provider: otpConfig.provider,
      phoneNumber: maskPhoneNumber(voterRecord.phoneNumber),
      phase: "resend",
    });
    setFlash(
      req,
      "success",
      `A new OTP code has been sent to ${maskPhoneNumber(voterRecord.phoneNumber)}.`,
    );
  } catch (error) {
    logAudit(req, "voter", voterRecord.staffId, "voter_otp_send_failed", {
      provider: otpConfig.provider,
      phoneNumber: maskPhoneNumber(voterRecord.phoneNumber),
      phase: "resend",
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
    positionsReviewed: selections.length,
    submittedChoices,
    skippedCount,
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

app.get("/observer/login", (req, res) => {
  if (req.session.observer) {
    return res.redirect("/observer");
  }

  return res.render("observer-login", {
    pageTitle: "Observer Sign In",
  });
});

app.post("/observer/login", requireCaptcha("observer_login"), (req, res) => {
  const observerId = normalizeObserverId(req.body.observerId);
  const password = String(req.body.password || "");
  const row = observerId
    ? db.prepare(`
        SELECT
          id,
          observer_id AS observerId,
          full_name AS fullName,
          organization,
          accreditation_number AS accreditationNumber,
          password_hash AS passwordHash,
          must_change_password AS mustChangePassword,
          is_active AS isActive,
          access_expires_at AS accessExpiresAt,
          failed_login_attempts AS failedLoginAttempts,
          locked_until AS lockedUntil
        FROM observer_accounts
        WHERE observer_id = ?
      `).get(observerId)
    : null;
  const lockedUntil = row?.lockedUntil ? dayjs(row.lockedUntil) : null;
  const isLocked = lockedUntil?.isValid() && dayjs().isBefore(lockedUntil);
  const accessExpired = row?.accessExpiresAt
    ? dayjs(row.accessExpiresAt).isValid() && !dayjs().isBefore(dayjs(row.accessExpiresAt))
    : false;
  const passwordAccepted = row ? bcrypt.compareSync(password, row.passwordHash) : false;

  if (!row || !row.isActive || accessExpired || isLocked || !passwordAccepted) {
    if (row && !isLocked && row.isActive && !accessExpired && !passwordAccepted) {
      const attempts = Number(row.failedLoginAttempts || 0) + 1;
      const shouldLock = attempts >= observerMaxLoginAttempts;
      db.prepare(`
        UPDATE observer_accounts
        SET failed_login_attempts = ?, locked_until = ?, updated_at = ?
        WHERE id = ?
      `).run(
        attempts,
        shouldLock ? dayjs().add(observerLockMinutes, "minute").toISOString() : "",
        nowIso(),
        row.id,
      );
    }

    logAudit(req, "observer", observerId || "unknown", "observer_login_failed", {
      reason: isLocked
        ? "account_locked"
        : accessExpired
          ? "access_expired"
          : row && !row.isActive
            ? "account_disabled"
            : "invalid_credentials",
    });
    setFlash(
      req,
      "error",
      isLocked
        ? `Too many unsuccessful attempts. Try again after ${formatDateTime(lockedUntil.toISOString())}.`
        : "Invalid observer ID or password. Check the credentials issued by the election committee.",
    );
    return res.redirect("/observer/login");
  }

  const timestamp = nowIso();
  db.prepare(`
    UPDATE observer_accounts
    SET failed_login_attempts = 0, locked_until = '', last_login_at = ?, updated_at = ?
    WHERE id = ?
  `).run(timestamp, timestamp, row.id);

  clearVoterSession(req);
  clearNominationSession(req);
  req.session.observer = {
    accountId: row.id,
    observerId: row.observerId,
    fullName: row.fullName,
  };
  logAudit(req, "observer", row.observerId, "observer_login_success", {
    organization: row.organization,
  });

  if (row.mustChangePassword) {
    setFlash(req, "info", "Create a private password before opening the observer dashboard.");
    return res.redirect("/observer/change-password");
  }

  return res.redirect("/observer");
});

app.get("/observer/change-password", requireObserver, (req, res) => {
  return res.render("observer-change-password", {
    pageTitle: "Create Observer Password",
    account: req.observerAccount,
  });
});

app.post("/observer/change-password", requireObserver, (req, res) => {
  const password = String(req.body.password || "");
  const passwordConfirmation = String(req.body.passwordConfirmation || "");
  const validation = validateObserverPassword(password);

  if (!validation.isValid) {
    setFlash(req, "error", validation.message);
    return res.redirect("/observer/change-password");
  }

  if (password !== passwordConfirmation) {
    setFlash(req, "error", "The password confirmation does not match.");
    return res.redirect("/observer/change-password");
  }

  const timestamp = nowIso();
  db.prepare(`
    UPDATE observer_accounts
    SET password_hash = ?, must_change_password = 0, password_changed_at = ?, updated_at = ?
    WHERE id = ?
  `).run(bcrypt.hashSync(password, 12), timestamp, timestamp, req.observerAccount.id);
  logAudit(req, "observer", req.observerAccount.observerId, "observer_password_changed");
  setFlash(req, "success", "Your private observer password has been saved.");
  return res.redirect("/observer");
});

app.get("/observer", requireObserverPasswordReady, (req, res) => {
  const metrics = getDashboardMetrics();
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const turnoutRate = metrics.totalVoters ? metrics.votedCount / metrics.totalVoters : 0;
  const remainingVoters = Math.max(metrics.totalVoters - metrics.votedCount, 0);
  const timelineSeries = getDashboardTimeline(settings, 8);
  const timelineChart = buildLineChartShape(timelineSeries, 680, 238);
  const resultsUnlocked = electionState.isClosed;

  return res.render("observer-dashboard", {
    pageTitle: "Observer Dashboard",
    account: req.observerAccount,
    activity: getAnonymizedObserverActivity(7),
    candidates: getBallotData(),
    incidents: getObserverIncidents({ accountId: req.observerAccount.id, limit: 8 }),
    integrityChecks: getObserverIntegrityChecks(),
    metrics,
    remainingVoters,
    results: resultsUnlocked ? getResultsSummary() : [],
    resultsUnlocked,
    settings,
    timelineChart,
    timelineSeries,
    turnoutRate,
  });
});

app.post("/observer/incidents", requireObserverPasswordReady, (req, res) => {
  const category = String(req.body.category || "general").trim().toLowerCase();
  const title = String(req.body.title || "").trim();
  const details = String(req.body.details || "").trim();
  const allowedCategories = new Set(["general", "access", "process", "technical", "security"]);

  if (!title || title.length > 120 || details.length < 10 || details.length > 2000) {
    setFlash(req, "error", "Enter a short title and incident details between 10 and 2,000 characters.");
    return res.redirect("/observer#report-incident");
  }

  const timestamp = nowIso();
  const result = db.prepare(`
    INSERT INTO observer_incidents (
      observer_account_id, category, title, details, status,
      submitted_at, created_at, updated_at
    ) VALUES (?, ?, ?, ?, 'submitted', ?, ?, ?)
  `).run(
    req.observerAccount.id,
    allowedCategories.has(category) ? category : "general",
    title,
    details,
    timestamp,
    timestamp,
    timestamp,
  );
  logAudit(req, "observer", req.observerAccount.observerId, "observer_incident_submitted", {
    incidentId: Number(result.lastInsertRowid),
    category: allowedCategories.has(category) ? category : "general",
  });
  setFlash(req, "success", "Your incident report has been recorded for the election committee.");
  return res.redirect("/observer#my-incidents");
});

app.get("/observer/report.pdf", requireObserverPasswordReady, (req, res) => {
  const metrics = getDashboardMetrics();
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  logAudit(req, "observer", req.observerAccount.observerId, "observer_report_exported");
  return streamObserverReportPdf(res, {
    account: req.observerAccount,
    activity: getAnonymizedObserverActivity(12),
    electionState,
    incidents: getObserverIncidents({ accountId: req.observerAccount.id, limit: 50 }),
    metrics,
    results: electionState.isClosed ? getResultsSummary() : [],
    settings,
  });
});

app.post("/observer/logout", requireObserver, (req, res) => {
  logAudit(req, "observer", req.observerAccount.observerId, "observer_logout");
  clearObserverSession(req);
  return res.redirect("/observer/login");
});

app.get("/admin/login", (req, res) => {
  if (req.session.admin) {
    return res.redirect("/admin");
  }

  if (req.session.pendingAdminTwoFactor) {
    return res.redirect("/admin/verify-2fa");
  }

  return res.render("admin-login", {
    pageTitle: "Admin Login",
    adminTwoFactorEnabled: getAdminTwoFactorState().enabled,
  });
});

app.post("/admin/login", requireCaptcha("admin_login"), (req, res) => {
  const username = String(req.body.username || "").trim();
  const password = String(req.body.password || "");
  const twoFactorState = getAdminTwoFactorState();

  if (username !== adminUsername || !bcrypt.compareSync(password, adminPasswordHash)) {
    logAudit(req, "admin", username || "unknown", "admin_login_failed");
    setFlash(req, "error", "Invalid administrator username or password.");
    return res.redirect("/admin/login");
  }

  if (twoFactorState.enabled) {
    req.session.pendingAdminTwoFactor = {
      username,
      issuedAt: nowIso(),
    };
    logAudit(req, "admin", username, "admin_login_password_verified");
    setFlash(req, "success", "Password accepted. Enter your authenticator code to finish signing in.");
    return res.redirect("/admin/verify-2fa");
  }

  req.session.admin = { username };
  logAudit(req, "admin", username, "admin_login_success");
  return res.redirect("/admin");
});

app.get("/admin/verify-2fa", (req, res) => {
  if (req.session.admin) {
    return res.redirect("/admin");
  }

  if (!req.session.pendingAdminTwoFactor) {
    setFlash(req, "error", "Sign in with your administrator credentials first.");
    return res.redirect("/admin/login");
  }

  return res.render("admin-verify-2fa", {
    pageTitle: "Admin Two-Factor Verification",
  });
});

app.post("/admin/verify-2fa", (req, res) => {
  const pendingAdminTwoFactor = req.session.pendingAdminTwoFactor;
  const verificationCode = normalizeTotpToken(req.body.verificationCode);
  const twoFactorState = getAdminTwoFactorState();

  if (!pendingAdminTwoFactor || !twoFactorState.enabled) {
    setFlash(req, "error", "Sign in with your administrator credentials first.");
    return res.redirect("/admin/login");
  }

  if (!verifyTotpToken(twoFactorState.secret, verificationCode)) {
    logAudit(
      req,
      "admin",
      pendingAdminTwoFactor.username || "unknown",
      "admin_2fa_verification_failed",
    );
    setFlash(req, "error", "Invalid two-factor code. Enter the latest code from your authenticator app.");
    return res.redirect("/admin/verify-2fa");
  }

  req.session.admin = { username: pendingAdminTwoFactor.username };
  req.session.pendingAdminTwoFactor = null;
  logAudit(req, "admin", pendingAdminTwoFactor.username, "admin_login_success");
  setFlash(req, "success", "Two-factor verification complete.");
  return res.redirect("/admin");
});

app.post("/admin/logout", requireAdmin, (req, res) => {
  logAudit(req, "admin", req.session.admin.username, "admin_logout");
  clearAdminAccess(req);
  res.redirect("/admin/login");
});

app.get("/admin/notifications", requireAdmin, (req, res) => {
  const limit = clampInteger(req.query.limit, 12, 1, 50);
  return res.json(getAdminNotificationSnapshot(limit));
});

app.post("/admin/notifications", requireAdmin, (req, res) => {
  const title = String(req.body.title || "").trim();
  const body = String(req.body.body || "").trim();
  const linkUrl = String(req.body.linkUrl || "").trim();

  if (!title || title.length > 120 || !body || body.length > 500) {
    return res.status(400).json({
      error: "Enter a title and message. Title must be 120 characters or fewer, and message must be 500 characters or fewer.",
    });
  }

  if (linkUrl && (!linkUrl.startsWith("/admin") || linkUrl.startsWith("//"))) {
    return res.status(400).json({
      error: "Message links must point to an admin page.",
    });
  }

  const notificationId = createAdminNotification({
    category: "message",
    priority: "normal",
    title,
    body,
    linkUrl,
    sourceType: "admin_message",
    sourceId: crypto.randomUUID(),
    createdBy: req.session.admin.username,
  });

  logAudit(req, "admin", req.session.admin.username, "admin_message_sent", {
    notificationId,
    title,
  });

  return res.status(201).json(getAdminNotificationSnapshot());
});

app.post("/admin/notifications/read", requireAdmin, (req, res) => {
  const all = req.body.all === true || req.body.all === "true" || req.body.all === "on";
  const ids = Array.isArray(req.body.ids)
    ? req.body.ids
    : req.body.id
      ? [req.body.id]
      : [];

  markAdminNotificationsRead({ ids, all });
  return res.json(getAdminNotificationSnapshot());
});

app.get("/admin/notifications/stream", requireAdmin, (req, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();

  adminNotificationClients.add(res);
  writeAdminNotificationEvent(res, "notifications:snapshot", getAdminNotificationSnapshot());

  const heartbeat = setInterval(() => {
    try {
      writeAdminNotificationEvent(res, "notifications:heartbeat", { at: nowIso() });
    } catch (_error) {
      clearInterval(heartbeat);
      adminNotificationClients.delete(res);
    }
  }, 25000);

  req.on("close", () => {
    clearInterval(heartbeat);
    adminNotificationClients.delete(res);
  });
});

app.get("/admin", requireAdmin, async (req, res) => {
  const metrics = getDashboardMetrics();
  const settings = getElectionSettings();
  const adminTwoFactorState = getAdminTwoFactorState();
  let pendingAdminTwoFactorSetup = req.session.adminTwoFactorSetup || null;
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

  if (pendingAdminTwoFactorSetup?.otpauthUri) {
    pendingAdminTwoFactorSetup = {
      ...pendingAdminTwoFactorSetup,
      qrCodeDataUrl: await buildAdminTotpQrCodeDataUrl(
        pendingAdminTwoFactorSetup.otpauthUri,
      ),
    };
  }

  return res.render("admin-dashboard", {
    pageTitle: "Admin Dashboard",
    metrics,
    settings,
    electionState,
    positionReadiness,
    archives,
    turnoutRate,
    notVotedCount,
    resultsPreview,
    activityFeed,
    topCandidates,
    positionBars,
    timelineSeries,
    timelineChart,
    themeOptions: getThemeOptions(),
    adminTwoFactorState,
    pendingAdminTwoFactorSetup,
  });
});

app.get("/admin/positions", requireAdmin, (req, res) => {
  const settings = getElectionSettings();
  const positions = getPositions();
  const candidates = getCandidates();
  const metrics = getDashboardMetrics();
  const readiness = getElectionReadiness(settings);

  return res.render("admin-positions", {
    pageTitle: "Positions & Election Schedule",
    settings,
    positions,
    candidates,
    metrics,
    readiness,
  });
});

app.get("/admin/settings", requireAdmin, async (req, res) => {
  const settings = getElectionSettings();
  const metrics = getDashboardMetrics();
  const adminTwoFactorState = getAdminTwoFactorState();
  let pendingAdminTwoFactorSetup = req.session.adminTwoFactorSetup || null;

  if (pendingAdminTwoFactorSetup?.otpauthUri) {
    pendingAdminTwoFactorSetup = {
      ...pendingAdminTwoFactorSetup,
      qrCodeDataUrl: await buildAdminTotpQrCodeDataUrl(
        pendingAdminTwoFactorSetup.otpauthUri,
      ),
    };
  }

  return res.render("admin-settings", {
    pageTitle: "System Settings",
    settings,
    metrics,
    readiness: getElectionReadiness(settings),
    otpSettings: getAdminOtpSettingsView(),
    captchaSettings: getAdminCaptchaSettingsView(),
    resultsSms: getResultsSmsStatus(getResultsExportPayload()),
    otpActivity: getOtpActivityLogs(7),
    themeOptions: getThemeOptions(),
    adminTwoFactorState,
    pendingAdminTwoFactorSetup,
    archives: getElectionArchives().slice(0, 3),
    isProduction,
  });
});

app.get("/admin/observers", requireAdmin, (req, res) => {
  const accounts = getObserverAccounts();
  const incidents = getObserverIncidents({ limit: 100 });
  const createdCredential = req.session.observerCredential || null;
  delete req.session.observerCredential;

  return res.render("admin-observers", {
    pageTitle: "Observer Management",
    accounts,
    accessLogs: getObserverAccessLogs(12),
    createdCredential,
    incidents,
    observerSummary: getObserverManagementSummary(accounts, incidents),
  });
});

app.post("/admin/observers", requireAdmin, (req, res) => {
  const fullName = String(req.body.fullName || "").trim();
  const organization = String(req.body.organization || "").trim();
  const accreditationNumber = String(req.body.accreditationNumber || "").trim();
  const email = String(req.body.email || "").trim();
  const phoneNumber = String(req.body.phoneNumber || "").trim();
  const accessExpiresAtValue = String(req.body.accessExpiresAt || "").trim();
  const accessExpiresAt = dayjs(accessExpiresAtValue);

  if (!fullName || !organization || !accreditationNumber) {
    setFlash(req, "error", "Full name, organization and accreditation number are required.");
    return res.redirect("/admin/observers#create-observer");
  }

  if (!accessExpiresAt.isValid() || !dayjs().isBefore(accessExpiresAt)) {
    setFlash(req, "error", "Choose a valid observer access expiry date in the future.");
    return res.redirect("/admin/observers#create-observer");
  }

  const observerId = generateObserverId();
  const temporaryPassword = generateObserverTemporaryPassword();
  const timestamp = nowIso();

  db.prepare(`
    INSERT INTO observer_accounts (
      observer_id, full_name, organization, accreditation_number, email,
      phone_number, password_hash, must_change_password, is_active,
      access_expires_at, created_by, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, 1, 1, ?, ?, ?, ?)
  `).run(
    observerId,
    fullName,
    organization,
    accreditationNumber,
    email,
    phoneNumber,
    bcrypt.hashSync(temporaryPassword, 12),
    accessExpiresAt.toISOString(),
    req.session.admin.username,
    timestamp,
    timestamp,
  );

  req.session.observerCredential = {
    observerId,
    temporaryPassword,
    fullName,
    action: "created",
  };
  logAudit(req, "admin", req.session.admin.username, "observer_account_created", {
    observerId,
    organization,
  });
  setFlash(req, "success", "Observer account created. Copy the temporary credentials from the secure window.");
  return res.redirect("/admin/observers");
});

app.post("/admin/observers/:id(\\d+)", requireAdmin, (req, res) => {
  const accountId = Number.parseInt(req.params.id, 10);
  const account = getObserverAccountById(accountId);
  const fullName = String(req.body.fullName || "").trim();
  const organization = String(req.body.organization || "").trim();
  const accreditationNumber = String(req.body.accreditationNumber || "").trim();
  const email = String(req.body.email || "").trim();
  const phoneNumber = String(req.body.phoneNumber || "").trim();
  const accessExpiresAt = dayjs(String(req.body.accessExpiresAt || "").trim());

  if (!account) {
    setFlash(req, "error", "Observer account not found.");
    return res.redirect("/admin/observers");
  }

  if (!fullName || !organization || !accreditationNumber || !accessExpiresAt.isValid()) {
    setFlash(req, "error", "Complete all required observer account fields.");
    return res.redirect(`/admin/observers#observer-${accountId}`);
  }

  db.prepare(`
    UPDATE observer_accounts
    SET full_name = ?, organization = ?, accreditation_number = ?, email = ?,
        phone_number = ?, access_expires_at = ?, updated_at = ?
    WHERE id = ?
  `).run(
    fullName,
    organization,
    accreditationNumber,
    email,
    phoneNumber,
    accessExpiresAt.toISOString(),
    nowIso(),
    accountId,
  );
  logAudit(req, "admin", req.session.admin.username, "observer_account_updated", {
    observerId: account.observerId,
  });
  setFlash(req, "success", `${account.observerId} was updated.`);
  return res.redirect(`/admin/observers#observer-${accountId}`);
});

app.post("/admin/observers/:id(\\d+)/status", requireAdmin, (req, res) => {
  const accountId = Number.parseInt(req.params.id, 10);
  const account = getObserverAccountById(accountId);
  const makeActive = String(req.body.action || "") === "enable";

  if (!account) {
    setFlash(req, "error", "Observer account not found.");
    return res.redirect("/admin/observers");
  }

  db.prepare(`
    UPDATE observer_accounts
    SET is_active = ?, failed_login_attempts = 0, locked_until = '', updated_at = ?
    WHERE id = ?
  `).run(makeActive ? 1 : 0, nowIso(), accountId);
  logAudit(req, "admin", req.session.admin.username, makeActive ? "observer_account_enabled" : "observer_account_disabled", {
    observerId: account.observerId,
  });
  setFlash(req, "success", `${account.observerId} is now ${makeActive ? "active" : "disabled"}.`);
  return res.redirect("/admin/observers");
});

app.post("/admin/observers/:id(\\d+)/reset-password", requireAdmin, (req, res) => {
  const accountId = Number.parseInt(req.params.id, 10);
  const account = getObserverAccountById(accountId);

  if (!account) {
    setFlash(req, "error", "Observer account not found.");
    return res.redirect("/admin/observers");
  }

  const temporaryPassword = generateObserverTemporaryPassword();
  db.prepare(`
    UPDATE observer_accounts
    SET password_hash = ?, must_change_password = 1, failed_login_attempts = 0,
        locked_until = '', password_changed_at = NULL, updated_at = ?
    WHERE id = ?
  `).run(bcrypt.hashSync(temporaryPassword, 12), nowIso(), accountId);
  req.session.observerCredential = {
    observerId: account.observerId,
    temporaryPassword,
    fullName: account.fullName,
    action: "reset",
  };
  logAudit(req, "admin", req.session.admin.username, "observer_password_reset", {
    observerId: account.observerId,
  });
  setFlash(req, "success", "A new temporary observer password was generated.");
  return res.redirect("/admin/observers");
});

app.post("/admin/observers/:id(\\d+)/delete", requireAdmin, (req, res) => {
  const accountId = Number.parseInt(req.params.id, 10);
  const account = getObserverAccountById(accountId);

  if (!account) {
    setFlash(req, "error", "Observer account not found.");
    return res.redirect("/admin/observers");
  }

  db.prepare("DELETE FROM observer_accounts WHERE id = ?").run(accountId);
  logAudit(req, "admin", req.session.admin.username, "observer_account_deleted", {
    observerId: account.observerId,
    organization: account.organization,
  });
  setFlash(req, "success", `${account.observerId} was deleted.`);
  return res.redirect("/admin/observers");
});

app.post("/admin/observer-incidents/:id(\\d+)", requireAdmin, (req, res) => {
  const incidentId = Number.parseInt(req.params.id, 10);
  const status = String(req.body.status || "reviewing").trim().toLowerCase();
  const adminNotes = String(req.body.adminNotes || "").trim().slice(0, 2000);
  const allowedStatuses = new Set(["submitted", "reviewing", "resolved", "dismissed"]);
  const incident = db.prepare("SELECT id FROM observer_incidents WHERE id = ?").get(incidentId);

  if (!incident || !allowedStatuses.has(status)) {
    setFlash(req, "error", "Observer incident could not be updated.");
    return res.redirect("/admin/observers#observer-incidents");
  }

  const timestamp = nowIso();
  db.prepare(`
    UPDATE observer_incidents
    SET status = ?, admin_notes = ?, reviewed_at = ?, reviewed_by = ?, updated_at = ?
    WHERE id = ?
  `).run(status, adminNotes, timestamp, req.session.admin.username, timestamp, incidentId);
  logAudit(req, "admin", req.session.admin.username, "observer_incident_reviewed", {
    incidentId,
    status,
  });
  setFlash(req, "success", "Observer incident status updated.");
  return res.redirect("/admin/observers#observer-incidents");
});

app.get("/admin/observers/report.pdf", requireAdmin, (req, res) => {
  const accounts = getObserverAccounts();
  const incidents = getObserverIncidents({ limit: 250 });
  logAudit(req, "admin", req.session.admin.username, "observer_register_exported", {
    observerCount: accounts.length,
    incidentCount: incidents.length,
  });
  return streamObserverManagementPdf(res, {
    accounts,
    incidents,
    settings: getElectionSettings(),
  });
});

app.get("/admin/otp-settings", requireAdmin, (req, res) => {
  return res.render("admin-otp-settings", {
    pageTitle: "OTP Settings",
    otpSettings: getAdminOtpSettingsView(),
    otpActivity: getOtpActivityLogs(7),
    isProduction,
  });
});

app.post("/admin/otp-settings", requireAdmin, (req, res) => {
  const provider = String(req.body.provider || "disabled").trim().toLowerCase();
  const ttlMinutes = clampInteger(req.body.ttlMinutes, 10, 1, 30);
  const resendCooldownSeconds = clampInteger(req.body.resendCooldownSeconds, 30, 0, 300);
  const submittedArkeselApiKey = String(req.body.arkeselApiKey || "").trim();
  const arkeselApiKey = submittedArkeselApiKey || getAdminOtpSettingsView().arkeselApiKey;
  const arkeselSenderId = String(req.body.arkeselSenderId || "").trim();
  const arkeselOtpMessage = String(
    req.body.arkeselOtpMessage || defaultArkeselOtpMessageTemplate,
  ).trim();
  const validProviders = new Set(["disabled", "arkesel", "twilio", "dev"]);

  if (!validProviders.has(provider)) {
    setFlash(req, "error", "Choose a valid OTP provider before saving.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  if (provider === "dev" && isProduction) {
    setFlash(req, "error", "Development OTP mode cannot be enabled on the live site.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  if (arkeselSenderId && arkeselSenderId.length > 11) {
    setFlash(req, "error", "Arkesel sender ID must be 11 characters or fewer.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  if (arkeselOtpMessage && !arkeselOtpMessage.includes("%otp_code%")) {
    setFlash(req, "error", "Arkesel OTP message must include %otp_code%.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  if (provider === "arkesel" && (!arkeselApiKey || !arkeselSenderId)) {
    setFlash(
      req,
      "error",
      "Enter both the Arkesel API key and sender ID before enabling Arkesel OTP.",
    );
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  runTransaction(() => {
    setSetting("otp_provider", provider);
    setSetting("otp_ttl_minutes", ttlMinutes);
    setSetting("otp_resend_cooldown_seconds", resendCooldownSeconds);
    setSetting("arkesel_api_key", arkeselApiKey);
    setSetting("arkesel_sender_id", arkeselSenderId);
    setSetting("arkesel_otp_message", arkeselOtpMessage);
  });

  logAudit(req, "admin", req.session.admin.username, "otp_settings_updated", {
    provider,
    ttlMinutes,
    resendCooldownSeconds,
    hasArkeselApiKey: Boolean(arkeselApiKey),
    hasArkeselSenderId: Boolean(arkeselSenderId),
  });

  const nextOtpConfig = getOtpConfig();
  const providerLabel = getOtpProviderLabel(provider);
  const configSuffix =
    provider === "arkesel" && !nextOtpConfig.arkeselConfigured
      ? " Arkesel is selected, but the credentials are still incomplete."
      : provider === "twilio" && !nextOtpConfig.twilioConfigured
        ? " Twilio is selected, but its environment credentials are not configured yet."
        : "";

  setFlash(req, "success", `${providerLabel} OTP settings saved.${configSuffix}`);
  return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
});

app.post("/admin/captcha-settings", requireAdmin, (req, res) => {
  const captchaEnabled = req.body.captchaEnabled === "on";
  const siteKey = String(req.body.captchaSiteKey || "").trim();
  const submittedSecretKey = String(req.body.captchaSecretKey || "").trim();
  const existingCaptchaConfig = getCaptchaConfig();
  const secretKey = submittedSecretKey || existingCaptchaConfig.secretKey;

  runTransaction(() => {
    setSetting("captcha_enabled", captchaEnabled ? "true" : "false");
    setSetting("captcha_site_key", siteKey);
    if (submittedSecretKey) {
      setSetting("captcha_secret_key", submittedSecretKey);
    }
    setSetting(
      "captcha_protect_voter_login",
      req.body.protectVoterLogin === "on" ? "true" : "false",
    );
    setSetting(
      "captcha_protect_admin_login",
      req.body.protectAdminLogin === "on" ? "true" : "false",
    );
    setSetting(
      "captcha_protect_observer_login",
      req.body.protectObserverLogin === "on" ? "true" : "false",
    );
    setSetting(
      "captcha_protect_nomination",
      req.body.protectNomination === "on" ? "true" : "false",
    );
  });

  const nextCaptchaConfig = getCaptchaConfig();
  logAudit(req, "admin", req.session.admin.username, "captcha_settings_updated", {
    enabled: captchaEnabled,
    active: nextCaptchaConfig.isEnabled,
    hasSiteKey: Boolean(siteKey || envCaptchaSiteKey),
    hasSecretKey: Boolean(secretKey || envCaptchaSecretKey),
    protectedForms: {
      voterLogin: nextCaptchaConfig.protectVoterLogin,
      adminLogin: nextCaptchaConfig.protectAdminLogin,
      observerLogin: nextCaptchaConfig.protectObserverLogin,
      nomination: nextCaptchaConfig.protectNomination,
    },
  });

  const statusMessage = nextCaptchaConfig.isEnabled
    ? "CAPTCHA protection is now active."
    : captchaEnabled
      ? "CAPTCHA settings saved, but protection is inactive until both keys are configured."
      : "CAPTCHA protection has been disabled.";

  setFlash(req, "success", statusMessage);
  return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
});

app.post("/admin/otp-settings/test", requireAdmin, async (req, res) => {
  const otpConfig = getOtpConfig();
  const rawPhoneNumber = String(req.body.testPhoneNumber || "").trim();
  const maskedPhoneNumber = maskPhoneNumber(rawPhoneNumber);

  if (!rawPhoneNumber) {
    setFlash(req, "error", "Enter a phone number before sending a test OTP.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  if (!isOtpVerificationEnabled(otpConfig)) {
    setFlash(
      req,
      "error",
      "OTP is currently disabled. Save and enable an OTP provider first before sending a test.",
    );
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  const smsPhoneNumber = toSmsPhoneNumber(rawPhoneNumber);
  if (!smsPhoneNumber) {
    logAudit(req, "admin", req.session.admin.username, "otp_test_send_failed", {
      provider: otpConfig.provider,
      phoneNumber: maskedPhoneNumber,
      reason: "invalid_sms_phone_format",
    });
    setFlash(req, "error", "Enter a valid phone number in SMS format before testing OTP.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
  }

  try {
    const challenge = await sendOtpChallenge(smsPhoneNumber, otpConfig);
    logAudit(req, "admin", req.session.admin.username, "otp_test_sent", {
      provider: otpConfig.provider,
      phoneNumber: maskedPhoneNumber,
    });

    const devSuffix =
      otpConfig.provider === "dev" && challenge?.devCodePreview && !isProduction
        ? ` Development code: ${challenge.devCodePreview}.`
        : "";
    setFlash(
      req,
      "success",
      `Test OTP request accepted for ${maskedPhoneNumber || rawPhoneNumber} using ${otpConfig.providerLabel}.${devSuffix}`,
    );
  } catch (error) {
    logAudit(req, "admin", req.session.admin.username, "otp_test_send_failed", {
      provider: otpConfig.provider,
      phoneNumber: maskedPhoneNumber,
      reason: "provider_error",
      message: error.message,
    });
    setFlash(req, "error", error.message);
  }

  return res.redirect(getAdminReturnPath(req, "/admin/settings#security"));
});

app.get("/admin/otp-logs", requireAdmin, (req, res) => {
  return res.render("admin-otp-logs", {
    pageTitle: "OTP Activity",
    logs: getOtpActivityLogs(7),
  });
});

app.post("/admin/2fa/setup", requireAdmin, (req, res) => {
  const adminTwoFactorState = getAdminTwoFactorState();

  if (adminTwoFactorState.enabled) {
    setFlash(req, "error", "Admin two-factor authentication is already enabled.");
    return res.redirect("/admin/settings#security");
  }

  const secret = generateTotpSecret();
  const settings = getElectionSettings();

  req.session.adminTwoFactorSetup = {
    secret,
    formattedSecret: formatTotpSecret(secret),
    issuer: settings.electionName,
    accountName: adminUsername,
    otpauthUri: buildAdminTotpUri(secret, settings.electionName),
    generatedAt: nowIso(),
  };

  logAudit(req, "admin", req.session.admin.username, "admin_2fa_setup_started");
  setFlash(
    req,
    "success",
    "Two-factor setup is ready. Scan the QR code with your authenticator app, then enter the 6-digit code to activate it.",
  );
  return res.redirect("/admin/settings#security");
});

app.post("/admin/2fa/cancel-setup", requireAdmin, (req, res) => {
  req.session.adminTwoFactorSetup = null;
  setFlash(req, "success", "Pending two-factor setup was cancelled.");
  return res.redirect("/admin/settings#security");
});

app.post("/admin/2fa/enable", requireAdmin, (req, res) => {
  const pendingAdminTwoFactorSetup = req.session.adminTwoFactorSetup;
  const verificationCode = normalizeTotpToken(req.body.verificationCode);

  if (!pendingAdminTwoFactorSetup?.secret) {
    setFlash(req, "error", "Start two-factor setup first before trying to enable it.");
    return res.redirect("/admin/settings#security");
  }

  if (!verifyTotpToken(pendingAdminTwoFactorSetup.secret, verificationCode)) {
    logAudit(req, "admin", req.session.admin.username, "admin_2fa_enable_failed");
    setFlash(req, "error", "Invalid authenticator code. Enter the latest 6-digit code and try again.");
    return res.redirect("/admin/settings#security");
  }

  setSetting("admin_2fa_secret", pendingAdminTwoFactorSetup.secret);
  setSetting("admin_2fa_enabled", "true");
  req.session.adminTwoFactorSetup = null;
  logAudit(req, "admin", req.session.admin.username, "admin_2fa_enabled");
  setFlash(req, "success", "Admin two-factor authentication is now enabled.");
  return res.redirect("/admin/settings#security");
});

app.post("/admin/2fa/disable", requireAdmin, (req, res) => {
  const adminTwoFactorState = getAdminTwoFactorState();
  const verificationCode = normalizeTotpToken(req.body.verificationCode);

  if (!adminTwoFactorState.enabled) {
    setFlash(req, "error", "Admin two-factor authentication is not enabled right now.");
    return res.redirect("/admin/settings#security");
  }

  if (!verifyTotpToken(adminTwoFactorState.secret, verificationCode)) {
    logAudit(req, "admin", req.session.admin.username, "admin_2fa_disable_failed");
    setFlash(req, "error", "Invalid authenticator code. Enter the latest 6-digit code to disable 2FA.");
    return res.redirect("/admin/settings#security");
  }

  setSetting("admin_2fa_secret", "");
  setSetting("admin_2fa_enabled", "false");
  req.session.adminTwoFactorSetup = null;
  clearAdminAccess(req);
  logAudit(req, "admin", adminUsername, "admin_2fa_disabled");
  setFlash(req, "success", "Admin two-factor authentication has been disabled. Sign in again to continue.");
  return res.redirect("/admin/login");
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

app.post("/admin/nominations/codes/generate", requireAdmin, (req, res) => {
  setFlash(req, "error", "This action is no longer available in the current nomination workflow.");
  return res.redirect("/admin/nominations");
});

app.post("/admin/nominations/codes/import", requireAdmin, (req, res) => {
  setFlash(req, "error", "This action is no longer available in the current nomination workflow.");
  return res.redirect("/admin/nominations");
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
  const electionState = computeElectionState(nextSettings);
  const opensAtValue = dayjs(nextSettings.nominationOpensAt);
  const closesAtValue = dayjs(nextSettings.nominationClosesAt);
  const shouldOpenImmediately =
    readiness.isReady &&
    !electionState.isOpen &&
    !electionState.isClosed &&
    opensAtValue.isValid() &&
    closesAtValue.isValid() &&
    !dayjs().isBefore(opensAtValue) &&
    dayjs().isBefore(closesAtValue);

  setSetting(
    "nomination_phase",
    electionState.isOpen || electionState.isClosed
      ? "closed"
      : shouldOpenImmediately
        ? "open"
        : "setup",
  );

  logAudit(req, "admin", req.session.admin.username, "nomination_settings_updated", {
    opensAt,
    closesAt,
  });

  setFlash(
    req,
    "success",
    electionState.isOpen || electionState.isClosed
      ? "Nomination settings updated. Nominations remain closed once voting has started."
      : shouldOpenImmediately
      ? "Nomination settings saved and nominations are now open."
      : "Nomination settings updated.",
  );
  return res.redirect("/admin/nominations");
});

app.post("/admin/nominations/open", requireAdmin, (req, res) => {
  const settings = getElectionSettings();
  const readiness = getNominationReadiness(settings);
  const electionState = computeElectionState(settings);

  if (electionState.isOpen || electionState.isClosed) {
    setFlash(req, "error", "Nominations cannot be opened once voting has started.");
    return res.redirect("/admin/nominations");
  }

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
    const staffId = normalizeStaffId(req.body.staffId);
    const phoneNumber = normalizePhoneNumber(req.body.phoneNumber);
    const department = String(req.body.department || "").trim();
    const positionId = parseInteger(req.body.positionId, 0);
    const bio = String(req.body.bio || "").trim();
    const manifesto = String(req.body.manifesto || "").trim();
    const proposerName = String(req.body.proposerName || "").trim();
    const seconderName = String(req.body.seconderName || "").trim();

    if (
      !fullName ||
      !staffId ||
      !phoneNumber ||
      !department ||
      !positionId ||
      !bio ||
      !manifesto ||
      !proposerName ||
      !seconderName
    ) {
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
      WHERE staff_id = ?
        AND position_id = ?
        AND id <> ?
    `).get(staffId, positionId, nominationId);

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
        staff_id = ?,
        full_name = ?,
        phone_number = ?,
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
      staffId,
      fullName,
      phoneNumber,
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
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
  }

  if (opensAt && closesAt && !dayjs(opensAt).isBefore(dayjs(closesAt))) {
    setFlash(req, "error", "The closing time must be later than the opening time.");
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
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
    if (nextSettings.nominationPhase !== "closed") {
      setSetting("nomination_phase", "closed");
      logAudit(req, "admin", req.session.admin.username, "nominations_closed_for_voting", {
        trigger: "election_settings_auto_open",
      });
    }
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

  return res.redirect(getAdminReturnPath(req, "/admin/positions"));
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
      return res.redirect(getAdminReturnPath(req, "/admin/settings#declaration"));
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

    return res.redirect(getAdminReturnPath(req, "/admin/settings#declaration"));
  },
);

app.post("/admin/theme", requireAdmin, (req, res) => {
  const themeName = String(req.body.themeName || "").trim();
  const selectedTheme = getThemeOptions().find((theme) => theme.value === themeName);

  if (!selectedTheme) {
    setFlash(req, "error", "Choose one of the available software themes.");
    return res.redirect(getAdminReturnPath(req, "/admin/settings#appearance"));
  }

  setSetting("theme_name", selectedTheme.value);
  logAudit(req, "admin", req.session.admin.username, "theme_updated", {
    themeName: selectedTheme.value,
  });
  setFlash(req, "success", `${selectedTheme.label} has been applied across the portal.`);
  return res.redirect(getAdminReturnPath(req, "/admin/settings#appearance"));
});

app.post(
  "/admin/logo",
  requireAdmin,
  brandingUpload.single("logo"),
  async (req, res) => {
    if (!req.file) {
      setFlash(req, "error", "Choose a logo image to upload.");
      return res.redirect(getAdminReturnPath(req, "/admin/settings#appearance"));
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

    return res.redirect(getAdminReturnPath(req, "/admin/settings#appearance"));
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
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
  }

  setSetting("election_phase", "open");
  if (settings.nominationPhase !== "closed") {
    setSetting("nomination_phase", "closed");
    logAudit(req, "admin", req.session.admin.username, "nominations_closed_for_voting", {
      trigger: "manual_election_open",
    });
  }
  logAudit(req, "admin", req.session.admin.username, "election_opened", {
    opensAt: settings.opensAt,
    closesAt: settings.closesAt,
  });

  setFlash(req, "success", "Voting has been opened and the election setup is now locked.");
  return res.redirect(getAdminReturnPath(req, "/admin/positions"));
});

app.post("/admin/election/close", requireAdmin, (req, res) => {
  const settings = getElectionSettings();

  if (settings.phase !== "open") {
    setFlash(req, "error", "Voting is not currently open.");
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
  }

  setSetting("election_phase", "closed");
  logAudit(req, "admin", req.session.admin.username, "election_closed");
  void triggerAutomaticResultsSms("manual_election_closed");
  setFlash(req, "success", "Voting has been closed. Results are now final. Automatic results SMS will send if enabled.");
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
    db.prepare("DELETE FROM nominations").run();
    db.prepare("DELETE FROM nomination_access_codes").run();
    db.prepare("DELETE FROM observer_incidents").run();
    db.prepare("DELETE FROM observer_accounts").run();
    db.prepare("DELETE FROM candidates").run();
    db.prepare("DELETE FROM positions").run();
    db.prepare("DELETE FROM voters").run();
    db.prepare("DELETE FROM admin_notifications").run();

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

app.post("/admin/system/fresh-reset", requireAdmin, async (req, res) => {
  const settings = getElectionSettings();
  const confirmationText = String(req.body.confirmationText || "").trim().toUpperCase();

  if (settings.phase === "open") {
    setFlash(req, "error", "Close voting before running a full fresh reset.");
    return res.redirect("/admin/settings#maintenance");
  }

  if (confirmationText !== "RESET ENTIRE SYSTEM") {
    setFlash(req, "error", "Type RESET ENTIRE SYSTEM exactly before running the fresh reset.");
    return res.redirect("/admin/settings#maintenance");
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
  const backupName = `vote-portal-pre-fresh-reset-${dayjs().format("YYYYMMDD-HHmmss")}.sqlite`;
  const backupPath = path.join(backupsDirectory, backupName);

  try {
    await fsp.copyFile(databasePath, backupPath);
  } catch (error) {
    setFlash(req, "error", `Fresh reset backup failed: ${error.message}`);
    return res.redirect("/admin/settings#maintenance");
  }

  runTransaction(() => {
    db.prepare("DELETE FROM ballot_entries").run();
    db.prepare("DELETE FROM ballots").run();
    db.prepare("DELETE FROM nominations").run();
    db.prepare("DELETE FROM nomination_access_codes").run();
    db.prepare("DELETE FROM observer_incidents").run();
    db.prepare("DELETE FROM observer_accounts").run();
    db.prepare("DELETE FROM candidates").run();
    db.prepare("DELETE FROM positions").run();
    db.prepare("DELETE FROM voters").run();
    db.prepare("DELETE FROM election_archives").run();
    db.prepare("DELETE FROM audit_logs").run();
    db.prepare("DELETE FROM admin_notifications").run();

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

  clearVoterSession(req);
  clearNominationSession(req);
  clearObserverSession(req);
  req.session.pendingVoterVerification = null;
  req.session.pendingAdminTwoFactor = null;

  logAudit(req, "admin", req.session.admin.username, "system_fresh_reset", {
    backupFile: backupName,
    preservedSettings: [
      "election_name",
      "organization_logo_path",
      "theme_name",
      "declaration_settings",
      "admin_2fa",
    ],
  });

  setFlash(
    req,
    "success",
    `Fresh reset complete. Test voters, nominations, candidates, results history, and audit logs were cleared. Safety backup saved as ${backupName}.`,
  );
  return res.redirect("/admin/settings#maintenance");
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

  return res.redirect(getAdminReturnPath(req, "/admin/settings#maintenance"));
});

app.get("/admin/voters", requireAdmin, (req, res) => {
  const metrics = getDashboardMetrics();
  const voters = getVoters();
  const voterActivity = getVoterManagementActivity(7);
  res.render("admin-voters", {
    pageTitle: "Voters",
    voters,
    metrics,
    voterActivity,
    templatePath: "/templates/staff-login-template.xlsx",
  });
});

app.get("/admin/voters/export/excel", requireAdmin, async (req, res) => {
  const voters = getVoters();
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const filename = `${toSafeFilename(settings.electionName)}-voters.xlsx`;
  const workbook = await buildVotersExportWorkbook(voters, settings, electionState);

  logAudit(req, "admin", req.session.admin.username, "voters_exported_excel", {
    totalVoters: voters.length,
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  await workbook.xlsx.write(res);
  res.end();
});

app.get("/admin/voters/export/pdf", requireAdmin, (req, res) => {
  const voters = getVoters();
  const settings = getElectionSettings();
  const electionState = computeElectionState(settings);
  const filename = `${toSafeFilename(settings.electionName)}-voters.pdf`;

  logAudit(req, "admin", req.session.admin.username, "voters_exported_pdf", {
    totalVoters: voters.length,
  });

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: "A4",
    margin: 0,
    info: {
      Title: `${settings.electionName} Voters Export`,
      Author: "Organization Vote Portal",
      Subject: "Voter management export",
    },
  });

  document.pipe(res);
  renderVotersPdf(document, {
    settings,
    voters,
    electionState,
    generatedAt: nowIso(),
  });
  document.end();
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

app.post("/admin/voters/:id(\\d+)/delete", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const voterId = parseInteger(req.params.id, 0);
  const voter = db.prepare(`
    SELECT
      id,
      staff_id AS staffId
    FROM voters
    WHERE id = ?
  `).get(voterId);

  if (!voter) {
    setFlash(req, "error", "Voter not found.");
    return res.redirect("/admin/voters");
  }

  const ballotCount = db.prepare("SELECT COUNT(*) AS total FROM ballots").get().total;

  if (ballotCount > 0) {
    setFlash(
      req,
      "error",
      "Voter records cannot be deleted after ballot activity exists. Start a new election cycle before removing voter records.",
    );
    return res.redirect("/admin/voters");
  }

  db.prepare("DELETE FROM voters WHERE id = ?").run(voterId);

  logAudit(req, "admin", req.session.admin.username, "voter_deleted", {
    voterId,
    staffId: voter.staffId,
  });

  setFlash(req, "success", `${voter.staffId} has been removed from the voter list.`);
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

app.get("/admin/ballot-layout", requireAdmin, (req, res) => {
  const ballotPositions = getBallotData();
  res.render("admin-ballot-layout", {
    pageTitle: "Ballot Layout",
    ballotPositions,
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
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
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

  return res.redirect(getAdminReturnPath(req, "/admin/positions"));
});

app.post("/admin/positions/:id", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const positionId = parseInteger(req.params.id, 0);
  const sortOrder = parseInteger(req.body.sortOrder, 0);
  const position = db.prepare(`
    SELECT
      id,
      name,
      sort_order AS sortOrder
    FROM positions
    WHERE id = ?
      AND is_active = 1
  `).get(positionId);

  if (!position) {
    setFlash(req, "error", "Position not found.");
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
  }

  db.prepare(`
    UPDATE positions
    SET
      sort_order = ?,
      updated_at = ?
    WHERE id = ?
  `).run(sortOrder, nowIso(), positionId);

  logAudit(req, "admin", req.session.admin.username, "position_sort_order_updated", {
    positionId,
    positionName: position.name,
    previousSortOrder: position.sortOrder,
    nextSortOrder: sortOrder,
  });
  setFlash(
    req,
    "success",
    `${position.name} display order updated to ${sortOrder}. Smaller numbers appear first.`,
  );
  return res.redirect(getAdminReturnPath(req, "/admin/positions"));
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
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
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
    return res.redirect(getAdminReturnPath(req, "/admin/positions"));
  }

  db.prepare("DELETE FROM positions WHERE id = ?").run(positionId);
  logAudit(req, "admin", req.session.admin.username, "position_deleted", {
    positionId,
  });
  setFlash(req, "success", "Position removed.");
  return res.redirect(getAdminReturnPath(req, "/admin/positions"));
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

app.post("/admin/candidates/:id/ballot-order", requireAdmin, (req, res) => {
  if (!ensureSetupMode(req, res)) {
    return;
  }

  const candidateId = parseInteger(req.params.id, 0);
  const sortOrder = parseInteger(req.body.sortOrder, 0);
  const redirectPositionId = parseInteger(req.body.positionId, 0);
  const redirectTo = String(req.body.redirectTo || "").trim();
  const candidate = db.prepare(`
    SELECT
      c.id,
      c.name,
      c.sort_order AS sortOrder,
      p.id AS positionId,
      p.name AS positionName
    FROM candidates c
    INNER JOIN positions p ON p.id = c.position_id
    WHERE c.id = ?
      AND c.is_active = 1
      AND p.is_active = 1
  `).get(candidateId);

  if (!candidate) {
    setFlash(req, "error", "Candidate not found.");
    return res.redirect("/admin/ballot-layout");
  }

  db.prepare(`
    UPDATE candidates
    SET
      sort_order = ?,
      updated_at = ?
    WHERE id = ?
  `).run(sortOrder, nowIso(), candidateId);

  logAudit(req, "admin", req.session.admin.username, "candidate_ballot_order_updated", {
    candidateId,
    candidateName: candidate.name,
    positionId: candidate.positionId,
    positionName: candidate.positionName,
    previousSortOrder: candidate.sortOrder,
    nextSortOrder: sortOrder,
  });
  setFlash(
    req,
    "success",
    `${candidate.name} will now appear with ballot order ${sortOrder} under ${candidate.positionName}.`,
  );

  if (redirectTo.startsWith("/admin/setup") || redirectTo.startsWith("/admin/ballot-layout")) {
    return res.redirect(redirectTo);
  }

  return res.redirect(
    `/admin/ballot-layout${redirectPositionId ? `#position-${redirectPositionId}` : ""}`,
  );
});

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

app.get("/admin/positions/:id/ballot-paper/print", requireAdmin, (req, res) => {
  const positionId = parseInteger(req.params.id, 0);
  const position = getBallotPositionById(positionId);

  if (!position) {
    setFlash(req, "error", "Ballot position not found.");
    return res.redirect("/admin/ballot-layout");
  }

  return res.render("admin-ballot-paper-print", {
    pageTitle: `${position.name} Ballot Paper`,
    position,
    generatedAt: nowIso(),
  });
});

app.get("/admin/positions/:id/ballot-paper/pdf", requireAdmin, (req, res) => {
  const positionId = parseInteger(req.params.id, 0);
  const position = getBallotPositionById(positionId);

  if (!position) {
    setFlash(req, "error", "Ballot position not found.");
    return res.redirect("/admin/ballot-layout");
  }

  const settings = getElectionSettings();
  const filename = `${toSafeFilename(settings.electionName)}-${toSafeFilename(position.name)}-ballot-paper.pdf`;

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

  const document = new PDFDocument({
    size: "A4",
    margin: 46,
    info: {
      Title: `${settings.electionName} ${position.name} Ballot Paper`,
      Author: "Organization Vote Portal",
      Subject: `${position.name} ballot paper`,
    },
  });

  logAudit(req, "admin", req.session.admin.username, "position_ballot_paper_downloaded", {
    positionId: position.id,
    positionName: position.name,
    candidateCount: position.candidates.length,
  });

  document.pipe(res);
  renderPositionBallotPaperPdf(document, {
    settings,
    position,
    generatedAt: nowIso(),
  });
  document.end();
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

app.post("/admin/results/sms-settings", requireAdmin, (req, res) => {
  const autoEnabled = req.body.autoEnabled === "on";

  setSetting("results_sms_auto_enabled", autoEnabled ? "true" : "false");
  logAudit(req, "admin", req.session.admin.username, "results_sms_settings_updated", {
    autoEnabled,
  });
  setFlash(
    req,
    "success",
    autoEnabled
      ? "Automatic provisional results SMS is enabled. It will send once voting closes."
      : "Automatic provisional results SMS has been disabled.",
  );

  return res.redirect(getAdminReturnPath(req, "/admin/results"));
});

app.post("/admin/results/sms", requireAdmin, async (req, res) => {
  const payload = getResultsExportPayload();

  if (resultsSmsAutoSendInProgress) {
    setFlash(
      req,
      "warning",
      "Automatic results SMS is already sending. Please wait a moment before sending manually.",
    );
    return res.redirect("/admin/results");
  }

  if (!payload.electionState.isClosed) {
    setFlash(
      req,
      "error",
      "Provisional results SMS can only be sent after voting has closed.",
    );
    return res.redirect("/admin/results");
  }

  const resultsSms = getResultsSmsStatus(payload);

  if (!resultsSms.configured) {
    setFlash(
      req,
      "error",
      "Add your Arkesel API key and sender ID before sending results by SMS.",
    );
    return res.redirect("/admin/results");
  }

  if (resultsSms.recipientCount === 0) {
    setFlash(req, "error", "No voter phone numbers are available for results SMS.");
    return res.redirect("/admin/results");
  }

  const confirmedResend = req.body.confirmResend === "on";
  if (resultsSms.alreadySentForCurrentResults && !confirmedResend) {
    setFlash(
      req,
      "error",
      "These exact results have already been sent. Tick the resend confirmation box if you want to send them again.",
    );
    return res.redirect("/admin/results");
  }

  const recipientSnapshot = getResultsSmsRecipients();
  const message = buildResultsSmsMessage(payload);
  const sentAt = nowIso();
  const sendResult = await sendResultsSmsBatch(
    recipientSnapshot.recipients,
    message,
    getOtpConfig(),
  );

  setSetting("results_sms_last_sent_at", sentAt);
  setSetting("results_sms_last_success_count", String(sendResult.successCount));
  setSetting("results_sms_last_failure_count", String(sendResult.failureCount));

  if (sendResult.successCount > 0) {
    setSetting("results_sms_last_fingerprint", resultsSms.currentFingerprint);
  }

  logAudit(req, "admin", req.session.admin.username, "results_sms_sent", {
    electionName: payload.settings.electionName,
    resultsFingerprint: resultsSms.currentFingerprint,
    recipientCount: recipientSnapshot.recipients.length,
    successCount: sendResult.successCount,
    failureCount: sendResult.failureCount,
    duplicateCount: recipientSnapshot.duplicateCount,
    invalidCount: recipientSnapshot.invalidCount,
    failedRecipients: sendResult.failedRecipients.slice(0, 5),
  });

  if (sendResult.successCount > 0 && sendResult.failureCount === 0) {
    setFlash(
      req,
      "success",
      `Provisional results SMS sent successfully to ${sendResult.successCount.toLocaleString()} voters.`,
    );
  } else if (sendResult.successCount > 0) {
    setFlash(
      req,
      "warning",
      `Results SMS sent to ${sendResult.successCount.toLocaleString()} voters, but ${sendResult.failureCount.toLocaleString()} failed. Check your Arkesel balance and phone numbers.`,
    );
  } else {
    setFlash(
      req,
      "error",
      "Results SMS could not be sent to any voter. Check your Arkesel balance, sender ID, and API key.",
    );
  }

  return res.redirect("/admin/results");
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
    resultsSms: getResultsSmsStatus(payload),
  });
});

app.get("/admin/audit", requireAdmin, (req, res) => {
  const logs = getAuditLogs(7);
  res.render("admin-audit", {
    pageTitle: "Audit Log",
    logs,
  });
});

app.use((error, req, res, _next) => {
  console.error(error);

  const statusCode = error.status || error.statusCode || 500;
  if (req.path.startsWith("/admin/notifications") || req.accepts("json") === "json") {
    return res.status(statusCode >= 400 && statusCode < 600 ? statusCode : 500).json({
      error: error.message || "Something went wrong.",
    });
  }

  const redirectTarget = req.path.startsWith("/admin")
    ? "/admin"
    : req.path.startsWith("/nomination")
      ? "/nomination/login"
      : "/";
  if (req.session) {
    setFlash(req, "error", error.message || "Something went wrong.");
  }
  res.redirect(redirectTarget);
});

async function start() {
  ensureDirectories();
  initDatabase(defaultElectionName);
  await ensureVoterTemplate(templatePath);
  await ensureVoterTemplate(staffLoginTemplatePath);

  app.listen(port, host, () => {
    const urls = getNetworkUrls(host, port);

    console.log("Vote portal is running.");
    console.log("");
    console.log("Open on this PC:");
    console.log(`- ${urls[0]}`);

    if (urls.length > 1) {
      console.log("");
      console.log("Open from other PCs on the same network:");

      for (const url of urls.slice(1)) {
        console.log(`- ${url}`);
      }
    }

    console.log("");
    console.log("Useful links:");
    console.log(`- Voter login: ${urls[0]}/vote/login`);
    console.log(`- Admin login: ${urls[0]}/admin/login`);
    console.log("");
    console.log(`Database file: ${databasePath}`);
    console.log(`Storage folder: ${storageRoot}`);
  });
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
