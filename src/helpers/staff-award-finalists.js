const { createHash, randomUUID } = require("node:crypto");

function createFinalistBridge({ db, hospital, getShortlist, isClosed, canPrepare }) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS staff_award_election_links (
      hospital TEXT NOT NULL,
      category TEXT NOT NULL,
      employee_number TEXT NOT NULL,
      nomination_id INTEGER NOT NULL UNIQUE REFERENCES nominations(id),
      source_signature TEXT NOT NULL,
      PRIMARY KEY (hospital, category, employee_number)
    )
  `);
  const links = db.prepare("SELECT l.*, n.published_candidate_id, n.status AS approval_status FROM staff_award_election_links l JOIN nominations n ON n.id=l.nomination_id WHERE l.hospital=?");
  const identity = values => [values.category, String(values.nominee_employeeNo || "").trim().toLowerCase()].join("\n");
  function shortlist() {
    const result = new Map();
    for (const row of getShortlist()) {
      const values = JSON.parse(row.submission);
      const key = identity(values);
      if (!result.has(key)) result.set(key, {
        values, signature: createHash("sha256").update(row.submission + ":" + row.review_id).digest("hex"),
      });
    }
    return result;
  }
  function sync() {
    if (!isClosed() || !canPrepare()) return;
    const current = shortlist();
    db.exec("BEGIN IMMEDIATE");
    try {
      const existing = new Map(links.all(hospital).map(link => [link.category + "\n" + link.employee_number, link]));
      const now = new Date().toISOString();
      for (const [key, { values, signature }] of current) {
        const link = existing.get(key);
        if (link?.published_candidate_id || link?.source_signature === signature) continue;
        db.prepare("INSERT OR IGNORE INTO positions(name,sort_order,is_active,created_at,updated_at) VALUES (?,0,1,?,?)").run(values.category, now, now);
        const position = db.prepare("SELECT id FROM positions WHERE name=?").get(values.category);
        const bio = "Staff Awards committee shortlist.\n\n" + values.reasons;
        if (link) {
          db.prepare("UPDATE nominations SET position_id=?,full_name=?,staff_id=?,department=?,bio=?,status='pending',admin_notes='',reviewed_at=NULL,reviewed_by='',updated_at=? WHERE id=? AND published_candidate_id IS NULL")
            .run(position.id, values.nominee_name, values.nominee_employeeNo, values.nominee_department, bio, now, link.nomination_id);
          db.prepare("UPDATE staff_award_election_links SET source_signature=? WHERE nomination_id=?").run(signature, link.nomination_id);
        } else {
          const inserted = db.prepare(`INSERT INTO nominations(application_number,position_id,staff_id,full_name,phone_number,department,bio,submitted_at,created_at,updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)`).run("AWARD-" + randomUUID().slice(0, 12).toUpperCase(),
              position.id, values.nominee_employeeNo, values.nominee_name, "", values.nominee_department, bio, now, now, now);
          db.prepare("INSERT INTO staff_award_election_links(hospital,category,employee_number,nomination_id,source_signature) VALUES (?,?,?,?,?)")
            .run(hospital, values.category, String(values.nominee_employeeNo).trim().toLowerCase(), Number(inserted.lastInsertRowid), signature);
        }
      }
      for (const [key, link] of existing) {
        if (current.has(key) || link.published_candidate_id || link.source_signature === "") continue;
        db.prepare("UPDATE nominations SET status='rejected',admin_notes='Removed from the committee shortlist.',updated_at=? WHERE id=? AND published_candidate_id IS NULL").run(now, link.nomination_id);
        db.prepare("UPDATE staff_award_election_links SET source_signature='' WHERE nomination_id=?").run(link.nomination_id);
      }
      db.exec("COMMIT");
    } catch (error) { db.exec("ROLLBACK"); throw error; }
  }
  function linked(values) {
    return links.all(hospital).find(link => link.category === values.category &&
      link.employee_number === String(values.nominee_employeeNo || "").trim().toLowerCase());
  }
  function isReady(nominationId) {
    const link = links.all(hospital).find(item => item.nomination_id === nominationId);
    if (!link) return true;
    const source = shortlist().get(link.category + "\n" + link.employee_number);
    return isClosed() && Boolean(source && source.signature === link.source_signature);
  }
  return { sync, linked, isReady };
}

module.exports = { createFinalistBridge };
