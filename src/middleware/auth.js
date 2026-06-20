const { db } = require("../db");

function requireAdmin(req, res, next) {
  if (!req.session.admin) {
    req.session.flash = {
      type: "error",
      message: "Please sign in as an administrator to continue.",
    };
    return res.redirect("/admin/login");
  }

  return next();
}

function requireObserver(req, res, next) {
  const observerSession = req.session.observer;

  if (!observerSession?.accountId) {
    req.session.flash = {
      type: "error",
      message: "Please sign in with your observer credentials to continue.",
    };
    return res.redirect("/observer/login");
  }

  const account = db.prepare(`
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
      last_login_at AS lastLoginAt
    FROM observer_accounts
    WHERE id = ?
  `).get(observerSession.accountId);

  const accessExpired = account?.accessExpiresAt
    ? new Date(account.accessExpiresAt).getTime() <= Date.now()
    : false;

  if (!account || !account.isActive || accessExpired) {
    req.session.observer = null;
    req.session.flash = {
      type: "error",
      message: accessExpired
        ? "Your observer access period has expired. Contact the election committee."
        : "This observer account is not active. Contact the election committee.",
    };
    return res.redirect("/observer/login");
  }

  req.observerAccount = account;
  req.session.observer = {
    accountId: account.id,
    observerId: account.observerId,
    fullName: account.fullName,
  };
  return next();
}

function requireObserverPasswordReady(req, res, next) {
  return requireObserver(req, res, () => {
    if (req.observerAccount.mustChangePassword) {
      req.session.flash = {
        type: "info",
        message: "Create a private password before opening the observer dashboard.",
      };
      return res.redirect("/observer/change-password");
    }

    return next();
  });
}

function requireVoter(req, res, next) {
  if (!req.session.voter) {
    req.session.flash = {
      type: "error",
      message: "Please sign in with your staff ID and phone number to continue.",
    };
    return res.redirect("/vote/login");
  }

  return next();
}

module.exports = {
  requireAdmin,
  requireObserver,
  requireObserverPasswordReady,
  requireVoter,
};
