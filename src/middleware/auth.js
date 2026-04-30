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
  requireVoter,
};
