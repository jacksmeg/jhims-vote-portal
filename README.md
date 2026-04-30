# Organization Vote Portal

An internal election web app for organizations that need:

- voter login with `staff ID + phone number`
- one-position-at-a-time voting wizard with next/back navigation
- one staff member to one ballot
- Excel-based voter import
- candidate and position management with photos
- candidate editing after save
- organization logo upload from admin
- open/close voting controls
- turnout tracking, results, audit logs, and backups
- printable results page and PDF export after voting closes

## Features

- Voter verification against imported staff records
- Duplicate-vote blocking after submission
- Ballot confirmation page before final vote submission
- Step-by-step position voting flow with next position navigation
- Admin login for election committee access
- Position and candidate setup pages
- Candidate edit screen for saved records
- Organization logo upload shown on admin and voter pages
- Excel/CSV voter import
- Final results page unlocked after the election is closed
- Results print view and downloadable PDF export
- Audit log for key admin and voter actions
- Database backup button from the admin dashboard

## Tech Stack

- Node.js
- Express
- EJS templates
- SQLite via Node's built-in `node:sqlite`
- Excel import/export with `exceljs`

## Quick Start

1. Install dependencies:

```bash
npm install
```

2. Copy the environment example if you want custom settings:

```bash
copy .env.example .env
```

3. Start the app:

```bash
npm start
```

4. Open:

```text
http://localhost:3000
```

## Default Admin Login

If you do not create a `.env` file, the app falls back to:

- Username: `admin`
- Password: `ChangeMe123!`

Change these before using the system for a real election.

## Environment Variables

See [.env.example](/C:/Users/El/Desktop/VOTE%20PORTAL/.env.example).

- `PORT` - local server port
- `HOST` - host interface to bind to
- `DATABASE_PATH` - SQLite file path
- `STORAGE_ROOT` - folder for uploads, imports, and backups
- `SESSION_SECRET` - session signing secret
- `SESSION_SECURE_COOKIE` - set to `true` behind HTTPS in production
- `ADMIN_USERNAME` - admin login username
- `ADMIN_PASSWORD` - admin login password
- `ELECTION_NAME` - name shown in the app

## Voter Import Template

The Excel template is generated at:

- [public/templates/voter-import-template.xlsx](/C:/Users/El/Desktop/VOTE%20PORTAL/public/templates/voter-import-template.xlsx)
- [public/templates/staff-login-template.xlsx](/C:/Users/El/Desktop/VOTE%20PORTAL/public/templates/staff-login-template.xlsx)

Required columns:

- `staff_id`
- `phone_number`

Optional columns:

- `full_name`
- `department`

## Candidate Setup

Use the admin dashboard to:

- add positions
- add candidates
- upload candidate photos
- set candidate ballot order

## Security Notes

- The current phone verification checks whether the entered phone number matches the registered voter record.
- Results stay hidden until the election is officially closed.
- Setup changes are locked once voting is opened.
- Every successful ballot marks the voter as already voted.

If you later want SMS OTP verification, we can add that as a second factor on top of the current staff record match.

## Put It Online

This app is ready to be hosted publicly so you can share separate links with voters and admins.

Recommended live links:

- Voters: `https://vote.yourdomain.com/vote/login`
- Admin: `https://vote.yourdomain.com/admin/login`

### Recommended Hosting

Because this system stores:

- the SQLite database
- imported voter files
- candidate photos
- the organization logo
- backup files

it should be deployed on hosting that supports persistent disk storage. A Render web service with a persistent disk is a good fit, and a starter or higher paid plan is required for the disk.

### Deployment Files Included

This project now includes [render.yaml](/C:/Users/El/Desktop/VOTE%20PORTAL/render.yaml), which is set up for:

- Node.js web hosting
- `/health` health checks
- persistent disk storage at `/var/data`
- secure cookies in production
- secret prompting for admin credentials

### Deploy Steps

1. Create a GitHub, GitLab, or Bitbucket repository for this project.
2. Push this project to that repository.
3. In Render, create a new Blueprint using the repository.
4. When prompted, set:
   - `ADMIN_USERNAME`
   - `ADMIN_PASSWORD`
   - `ELECTION_NAME`
5. After deploy finishes, open your live URL.
6. In Render, add your custom domain or subdomain such as `vote.yourdomain.com`.
7. Share these links with your team:
   - Voters: `https://vote.yourdomain.com/vote/login`
   - Admin: `https://vote.yourdomain.com/admin/login`

### Important Production Note

Do not deploy this app to hosting that has only temporary filesystem storage unless you also change it to use an external database and file storage service. Without persistent storage, uploads and election data can be lost after restart or redeploy.
