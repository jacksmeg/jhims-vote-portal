const { randomUUID } = require("node:crypto");

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
    )
  `);
  const insert = db.prepare("INSERT INTO staff_award_nominations (hospital, request_id, submission) VALUES (?, ?, ?) ON CONFLICT(request_id) DO NOTHING");
  const list = db.prepare("SELECT * FROM staff_award_nominations WHERE hospital = ? AND id < ? ORDER BY id DESC LIMIT 21");
  const latest = db.prepare("SELECT * FROM staff_award_nominations WHERE hospital = ? ORDER BY id DESC LIMIT 21");
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

    const validStaff = prefix => staffFields.every(([field]) => {
      const value = values[prefix + "_" + field];
      return value.length > 0 && value.length <= 120;
    });
    if (!["self", "another"].includes(values.nominationType) ||
        !/^[0-9a-f-]{36}$/.test(requestId) ||
        !categories.includes(values.category) ||
        !validStaff("nominee") ||
        (values.nominationType === "another" && !validStaff("nominator")) ||
        !/^\d{1,2}$/.test(values.years) ||
        !values.reasons || [values.reasons, values.idea, values.implementation].some(value => value.length > 2000) ||
        Boolean(values.idea) !== Boolean(values.implementation) ||
        !criteria.every(([key]) => ["1", "2", "3"].includes(values[key]))) {
      return renderForm(req, res, { status: 422, values, error: "Complete the required fields and all nine ratings. If you include an innovation, describe both the idea and how it was implemented." });
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

  app.get("/admin/staff-awards", requireAdmin, (req, res) => {
    res.set("Cache-Control", "no-store");
    const before = Number(req.query.before);
    const rows = Number.isSafeInteger(before) && before > 0 ? list.all(hospital, before) : latest.all(hospital);
    const nominations = rows.slice(0, 20).map(row => ({ ...row, values: JSON.parse(row.submission) }));
    return res.render("admin-staff-awards", {
      ...model, pageTitle: "Staff Award Nominations", nominations, nextId: rows.length > 20 ? rows[19].id : null,
    });
  });
}

module.exports = { registerStaffAwards };
