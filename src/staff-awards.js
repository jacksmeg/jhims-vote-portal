const { createHash, randomUUID } = require("node:crypto");

const categories = ["Best Nurse", "Best Midwife", "Best Clinical Staff", "Best Paramedical Staff"];
const staffFields = [
  ["name", "Full name"], ["grade", "Grade"], ["employeeNo", "Employee number"],
  ["department", "Department or unit"],
];
const rubric = [
  ["Professionalism", [
    ["standards", "Adherence to professional standards", [
      "Performs tasks competently with occasional errors.",
      "Consistently follows professional standards and guidelines.",
      "Goes beyond minimum requirements and sets an example for others.",
    ]],
    ["crisis", "Handling challenging situations", [
      "Responds with some composure but may need time or support.",
      "Handles crises calmly and decisively with minimal disruption.",
      "Responds quickly and strategically while communicating clearly.",
    ]],
  ]],
  ["Teamwork", [
    ["collaboration", "Collaboration with team members", [
      "Contributes to team discussions and activities.",
      "Actively fosters communication and teamwork.",
      "Leads and motivates colleagues to achieve team goals.",
    ]],
    ["assistance", "Willingness to assist colleagues", [
      "Offers help when workload and availability allow.",
      "Readily offers support and looks for opportunities to help.",
      "Goes above and beyond to support colleagues and team success.",
    ]],
    ["leadership", "Leadership in performance", [
      "Provides direction and support and motivates colleagues when needed.",
      "Provides clear direction, support, and recognition that help the team succeed.",
      "Sets a clear vision and builds commitment and enthusiasm across the team.",
    ]],
  ]],
  ["Innovation", [
    ["originality", "Originality and implementation of ideas", [
      "Is willing to implement new ideas with support.",
      "Takes initiative to develop and implement new ideas.",
      "Leads innovations that produce meaningful improvements in service quality.",
    ]],
    ["replication", "Ease of replication", [
      "The idea can be replicated with some clarification or adaptation.",
      "Communicates the idea clearly enough for others to replicate it.",
      "Anticipates barriers and makes the idea practical and easy to adopt.",
    ]],
  ]],
  ["Customer Care", [
    ["patients", "Positive interaction with patients or clients", [
      "Is generally polite and provides some information and support.",
      "Listens, acknowledges concerns, and provides clear information and support.",
      "Builds trust and proactively provides resources and support.",
    ]],
    ["colleagues", "Positive interaction with colleagues", [
      "Communicates well at times and considers others’ perspectives.",
      "Maintains respectful relationships, collaborates, and values colleagues’ input.",
      "Listens actively, encourages others, and looks for ways to support them.",
    ]],
  ]],
];
const criteria = rubric.flatMap(([, items]) => items);

const adminPath = "/admin/staff-awards";
const validationError = "Complete the required fields and all nine ratings. If you include an innovation, describe both the idea and how it was implemented.";
const readText = (value, limit = Infinity) => typeof value === "string" ? value.trim().slice(0, limit) : "";
const positiveId = value => /^\d+$/.test(String(value)) && Number.isSafeInteger(Number(value)) && Number(value) > 0 ? Number(value) : null;
const revision = submission => createHash("sha256").update(submission).digest("hex");

function readSubmission(body) {
  const read = key => readText(body[key]);
  const values = {
    nominationType: read("nominationType"), category: read("category"),
    years: read("years"), reasons: read("reasons"), idea: read("idea"),
    implementation: read("implementation"),
  };
  for (const [field] of staffFields) {
    values["nominee_" + field] = read("nominee_" + field);
    if (values.nominationType === "another") values["nominator_" + field] = read("nominator_" + field);
  }
  for (const [key] of criteria) values[key] = read(key);

  return values;
}

function validSubmission(values) {
  const validStaff = prefix => staffFields.every(([field]) => {
    const value = values[prefix + "_" + field];
    return value.length > 0 && value.length <= 120;
  });
  return ["self", "another"].includes(values.nominationType) &&
    categories.includes(values.category) &&
    validStaff("nominee") &&
    (values.nominationType !== "another" || validStaff("nominator")) &&
    /^\d{1,2}$/.test(values.years) &&
    Boolean(values.reasons) &&
    [values.reasons, values.idea, values.implementation].every(value => value.length <= 2000) &&
    Boolean(values.idea) === Boolean(values.implementation) &&
    criteria.every(([key]) => ["1", "2", "3"].includes(values[key]));
}

function registerStaffAwards(app, { db, requireAdmin, verifyCaptchaSubmission }) {
  const hospital = (process.env.STAFF_AWARDS_HOSPITAL || "").trim();
  if (!hospital) return;

  db.exec(`
    CREATE TABLE IF NOT EXISTS staff_award_nominations (
      id INTEGER PRIMARY KEY,
      hospital TEXT NOT NULL,
      submitted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      request_id TEXT NOT NULL UNIQUE,
      submission TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS staff_award_deleted_nominations (
      nomination_id INTEGER PRIMARY KEY,
      deleted_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    )
  `);
  const insert = db.prepare("INSERT INTO staff_award_nominations (hospital, request_id, submission) VALUES (?, ?, ?) ON CONFLICT(request_id) DO NOTHING");
  const active = "hospital = ? AND id NOT IN (SELECT nomination_id FROM staff_award_deleted_nominations)";
  const matching = `${active}
    AND (? = '' OR json_extract(submission, '$.category') = ?)
    AND (? = '' OR instr(lower(json_extract(submission, '$.nominee_name')), ?) > 0
      OR instr(lower(json_extract(submission, '$.nominee_employeeNo')), ?) > 0)`;
  const list = db.prepare(`SELECT id, submitted_at,
    json_extract(submission, '$.nominee_name') AS nominee,
    json_extract(submission, '$.category') AS category
    FROM staff_award_nominations WHERE ${matching} AND (? IS NULL OR id < ?)
    ORDER BY id DESC LIMIT 21`);
  const count = db.prepare(`SELECT count(*) AS total FROM staff_award_nominations WHERE ${matching}`);
  const find = db.prepare(`SELECT * FROM staff_award_nominations WHERE ${active} AND id = ?`);
  const update = db.prepare(`UPDATE staff_award_nominations SET submission = ?
    WHERE ${active} AND id = ? AND submission = ?`);
  const remove = db.prepare(`INSERT INTO staff_award_deleted_nominations (nomination_id)
    SELECT id FROM staff_award_nominations WHERE ${active} AND id = ? AND submission = ?`);
  const model = { hospital, categories, staffFields, rubric };

  function renderForm(req, res, { status = 200, values = {}, error = "", submitted = false } = {}) {
    req.session.staffAwardsToken ||= randomUUID();
    return res.status(status).render("staff-awards", {
      ...model, pageTitle: "2026 Staff Awards", token: req.session.staffAwardsToken,
      values, error, submitted, requestId: values.requestId || randomUUID(),
    });
  }

  app.get("/staff-awards", (req, res) => {
    res.set("Cache-Control", "no-store");
    const submitted = req.session.staffAwardsSubmitted === true;
    delete req.session.staffAwardsSubmitted;
    return renderForm(req, res, { submitted });
  });

  app.post("/staff-awards", async (req, res, next) => {
    res.set("Cache-Control", "no-store");
    const read = key => typeof req.body[key] === "string" ? req.body[key].trim() : "";
    if (!req.session.staffAwardsToken || read("token") !== req.session.staffAwardsToken) {
      return renderForm(req, res, { status: 403, error: "Your form expired. Please fill it in again." });
    }

    const requestId = read("requestId");
    const values = readSubmission(req.body);
    if (!/^[0-9a-f-]{36}$/.test(requestId) || !validSubmission(values)) {
      return renderForm(req, res, { status: 422, values, error: validationError });
    }

    try {
      const captcha = await verifyCaptchaSubmission(req, "nomination");
      if (!captcha.ok) {
        return renderForm(req, res, { status: 422, values: { ...values, requestId }, error: "Complete the security check and try again." });
      }
      insert.run(hospital, requestId, JSON.stringify(values));
      req.session.staffAwardsSubmitted = true;
      return res.redirect(303, "/staff-awards");
    } catch (error) {
      return next(error);
    }
  });

  function renderAdmin(req, res, { status = 200, error = "", draft } = {}) {
    res.set("Cache-Control", "no-store");
    req.session.staffAwardsAdminToken ||= randomUUID();
    const q = readText(req.query.q, 120);
    const category = categories.includes(req.query.category) ? req.query.category : "";
    const before = positiveId(req.query.before);
    const id = positiveId(req.query.id);
    const query = q.toLowerCase();
    const params = [hospital, category, category, query, query, query];
    const rows = list.all(...params, before, before);
    const nominations = rows.slice(0, 20);
    const selected = find.get(hospital, id || nominations[0]?.id || 0);
    const mode = ["edit", "delete"].includes(req.query.action) ? req.query.action : "";
    const feedback = req.session.staffAwardsAdminFeedback;
    delete req.session.staffAwardsAdminFeedback;

    function pageUrl(overrides = {}) {
      const pairs = Object.entries({ q, category, before, ...overrides }).filter(([, value]) => value);
      return pairs.length ? adminPath + "?" + new URLSearchParams(pairs).toString() : adminPath;
    }

    return res.status(status).render("admin-staff-awards", {
      ...model, pageTitle: "Staff award nominations", nominations, q, category, before, pageUrl,
      total: count.get(...params).total, nextId: rows.length > 20 ? nominations.at(-1).id : null,
      selected, values: draft || (selected ? JSON.parse(selected.submission) : {}), mode,
      token: req.session.staffAwardsAdminToken, version: selected ? revision(selected.submission) : "",
      error: error || (id && !selected ? "This nomination is no longer available." : ""), feedback,
    });
  }

  app.get(adminPath, requireAdmin, (req, res) => renderAdmin(req, res));

  app.post(adminPath, requireAdmin, (req, res) => {
    res.set("Cache-Control", "no-store");
    const body = req.body || {};
    if (!req.session.staffAwardsAdminToken || body.token !== req.session.staffAwardsAdminToken) {
      return res.status(403).send("Your form expired. Reopen Staff Awards and try again.");
    }
    const id = positiveId(req.query.id);
    const action = req.query.action;
    if (!id || !["edit", "delete"].includes(action)) return res.sendStatus(400);
    const current = find.get(hospital, id);
    if (!current) return renderAdmin(req, res, { status: 404 });
    if (body.version !== revision(current.submission)) {
      return renderAdmin(req, res, { status: 409, error: "This nomination changed. Review the latest version before continuing." });
    }

    let result;
    if (action === "edit") {
      const previous = JSON.parse(current.submission);
      const draft = readSubmission({ ...body, nominationType: previous.nominationType });
      if (!validSubmission(draft)) return renderAdmin(req, res, { status: 422, error: validationError, draft });
      result = update.run(JSON.stringify({ ...previous, ...draft }), hospital, id, current.submission);
    } else {
      result = remove.run(hospital, id, current.submission);
    }
    if (!result.changes) return renderAdmin(req, res, { status: 409, error: "This nomination changed. Refresh the list and try again." });

    req.session.staffAwardsAdminFeedback = action === "edit" ? "Nomination updated." : "Nomination deleted.";
    const redirect = new URLSearchParams();
    const q = readText(req.query.q, 120);
    if (q) redirect.set("q", q);
    if (categories.includes(req.query.category)) redirect.set("category", req.query.category);
    if (positiveId(req.query.before)) redirect.set("before", req.query.before);
    if (action === "edit") redirect.set("id", id);
    return res.redirect(303, redirect.size ? adminPath + "?" + redirect.toString() : adminPath);
  });
}

module.exports = { registerStaffAwards };
