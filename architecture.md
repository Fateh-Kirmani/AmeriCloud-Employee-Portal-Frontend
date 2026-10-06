# AmeriCloud Employee Portal — Architecture

## Overview

A single-page web application (SPA) for AmeriCloud Telecom employees to submit timesheets, manage PTO, sign documents, and access HR tools. No build step, no framework — served as static files from RunCloud with a FastAPI backend on Vercel and Supabase as the database.

**Live URL:** `https://americloud-employee-portal.vercel.app`

---

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | Vanilla HTML / CSS / ES5 JavaScript — single `index.html` |
| Auth | MSAL.js v3 (`@azure/msal-browser@3`) — CDN via jsdelivr |
| Fonts | Google Fonts: Archivo 700/800 (headings), Manrope 400/500/600 (body) |
| Backend | FastAPI on Vercel (`https://americloud-employee-portal-backend.vercel.app`) |
| Database | Supabase (PostgreSQL + file storage) |
| Auth provider | Microsoft Entra ID (Azure AD) |
| Hosting | Vercel — auto-deploys from `master` branch |
| Repo | `Fateh-Kirmani/AmeriCloud-Employee-Portal-Frontend` (GitHub) |

---

## File Structure

```
americloud-employee-portal/
├── index.html              — SPA shell: all views, sections, panels, modals
├── portal.js               — All client logic (~3900 lines)
├── portal-config.js        — Static data: news, directory, staff roster
├── portal.css              — All styles + design tokens
├── architecture.md         — This document
├── README.md               — Setup and deploy guide
├── .gitignore
├── deploy/
│   └── nginx-runcloud.conf — Nginx hardening config for RunCloud
└── assets/
    ├── brand/
    │   ├── americloud-logo.png         — Sign-in card logo
    │   ├── americloud-logo-white.png   — Sidebar + topbar logo
    │   └── favicon.png
    ├── photos/
    │   └── home-hero-network-dusk.png  — Home section hero background
    └── docs/
        ├── AmeriCloud Telecom Handbook- 2026.pdf
        ├── AmeriCloud Telecom Weekly Timesheet - Standard 2026 6.xlsx
        ├── AmeriCloud Telecom Weekly Timesheet - California 2026 5.xlsx
        └── AmeriCloud Telecom Weekly Timesheet - Offshore 2026 1.xlsx
```

---

## Authentication & Authorization

### Flow

1. Page loads → shows loading screen
2. `msal.PublicClientApplication` initializes
3. `handleRedirectPromise()` resolves any in-progress redirect
4. If accounts exist in `sessionStorage` → go to dashboard; else → show sign-in screen
5. Sign-in button calls `msalInstance.loginRedirect(loginRequest)` — redirects to Microsoft's login page
6. On return, MSAL handles the redirect, stores the account, calls `initDashboard(account)`
7. Sign-out calls `msalInstance.logoutRedirect()`

### MSAL Config

```
Client ID:    64cfff00-2deb-4a7a-91fc-74431baac966
Tenant ID:    f8e7955f-b183-4786-b8ca-7a51ecf001e8
Authority:    https://login.microsoftonline.com/f8e7955f-b183-4786-b8ca-7a51ecf001e8
Redirect URI: window.location.origin
Cache:        sessionStorage
Scopes:       User.Read, openid, profile
```

### Role Detection

Roles are read from the `groups` array in the ID token's `idTokenClaims`. Azure AD must be configured to emit group claims.

```
GROUP_MANAGERS = 5ef854d4-0938-4d06-87a6-3643b9a92796
GROUP_HR       = 6f5e8021-9bcd-44c8-8a66-cff4014d5e07
```

| Role | Condition | Extra nav item unlocked |
|---|---|---|
| Employee | Everyone authenticated | — |
| Manager | Member of `GROUP_MANAGERS` | Manager Approvals section |
| HR | Member of `GROUP_HR` | HR Dashboard section |

A user can be both Manager and HR simultaneously.

### API Authentication

Every `apiCall()` attaches the MSAL ID token as a Bearer token:
```
Authorization: Bearer {idToken}
```
The backend validates this token against Azure AD.

---

## Navigation & Sections

All sections live in `index.html` and are toggled by `hidden` attribute. Navigation links carry `data-section` attributes; clicking them calls `activateSection(id)`.

| Section ID | Nav label | Visible to |
|---|---|---|
| `home` | Home | All |
| `pay` | Pay & Benefits | All |
| `timesheets` | Timesheets | All |
| `forms` | Forms & Policies | All |
| `resources` | Resources & Tools | All |
| `trainings` | Trainings & Certs | All |
| `manager` | Manager Approvals | Managers only |
| `hr` | HR Dashboard | HR only |

`nav-manager` and `nav-hr` are `hidden` by default and revealed by JS after role detection.

---

## Panel System

Panels (drawers/modals) are managed by a module-level panel system:

- `openPanelModal(panel)` — closes any open panel, shows backdrop, opens new panel
- `closePanelModal()` — hides active panel and backdrop
- `portal-panel-backdrop` — semi-transparent overlay; clicking it closes the active panel
- All `.portal-ts-panel__close` buttons are wired at dashboard init to call `closePanelModal()`

Panels are full-height right-side drawers on desktop, full-screen overlays on mobile.

---

## Features by Role

### All Employees

#### Home Greeting (`setupHomeGreeting`)
- Time-of-day greeting with display name
- At-a-glance cards: next payday, policies to sign (from `/policies/my`), timesheets to review (managers only, from `/timesheets/form/team`)
- Hero ticker strip: employee count + countries from `PORTAL_CONFIG.directory` (static)
- Company news rendered from `PORTAL_CONFIG.news` (static — first item as featured, rest in grid)

#### Timesheet Submission (`setupTimesheetForm`)
See [Timesheet System](#timesheet-system) section for full detail.

#### PTO Request (`setupPTORequest`)
- Panel: `panel-pto-request`
- Fields: Start date, End date (max +13 days), Reason
- Submit: `POST /pto/` → `{ start_date, end_date, reason }`

#### My PTO Status (`setupMyPTORequests`)
- Panel: `panel-my-pto`
- Loads: `GET /pto/my`
- Shows table with Manager Status and HR Status columns

#### Incident Report (`setupIncidentReport`)
- Panel: `panel-incident-report`
- Fields: Incident date, Location, Incident type, Description
- Submit: `POST /incidents/` → `{ incident_date, location, incident_type, description }`

#### Employee Handbook (`setupHandbook`)
- "View" link opens `assets/docs/AmeriCloud Telecom Handbook- 2026.pdf` in a new tab
- "Sign Handbook" button: `POST /handbook/sign` — records acknowledgment tied to the user's token identity

#### My Policies (`setupEmployeePolicies`)
- Opens modal listing policies issued to the signed-in employee
- Loads: `GET /policies/my`
- "Acknowledge" button: `POST /policies/acknowledge/{id}`
- Policy name link: `GET /policies/issued/{issuanceId}/url` → opens PDF in browser

#### My EFS (`setupEmployeeEfs`)
- EFS = Essential Functions Statement (HR-issued document)
- On init: `GET /efs/my` — if items exist, shows the EFS list item in the UI
- "View" button: `GET /efs/view/{id}` → signed URL, opens in new tab
- "Sign EFS" button (Pending only): `POST /efs/sign/{id}`

#### Company Directory (`setupDirectory`)
- Panel: `panel-directory`
- Data: `PORTAL_CONFIG.directory` — pure client-side, no API
- Search by name or email (live filter)
- Filter chips by country (US, India, Pakistan) and department
- Each card: avatar initial (color-coded by country), name, email, Teams deep-link
- Managers also have a "Team Directory" button that opens the same panel filtered to their team

#### Company Calendar (`setupCalendar`)
- Panel: `panel-calendar`
- Pure client-side render: payday dates (15th of each month, adjusted for weekends), 2026/2027 holiday schedule
- No API call

#### Trainings & Certifications (`setupTrainings`)
- View: `GET /trainings` — lists all uploaded training docs for the signed-in user
  - Click doc name: `GET /trainings/{id}/file` → opens blob URL in new tab
  - Edit issuance/renewal dates inline: `PATCH /trainings/{id}`
  - Delete (admin only — `fkirmani@americloudtelecom.com`): `DELETE /trainings/{id}`
- Upload: `POST /trainings/upload` (multipart: `file`, `issuance_date`, `renewal_date`)
  - Admin can additionally supply `employee_email` and `employee_name` to upload on behalf of any employee

---

### Manager Only

#### Team Timesheet Viewer (`setupFormTimesheetViewer('manager', teamEmails)`)
- Loads: `GET /timesheets/form/team?emails={comma-separated}`
- Table: Employee, Week, Type, Submitted, Status, Approve/Reject buttons
- Quick approve/reject: `POST /timesheets/form/{id}/approve` or `/reject`
- Detail view: `GET /timesheets/form/{id}` → full timesheet with editable supervisor signature fields
- Approve from detail: `POST /timesheets/form/{id}/approve` with `{ manager_sig, manager_sig_date }`

#### Team PTO Approvals (`setupPTOApprovals(teamEmails, 'manager')`)
- Loads: `GET /pto/team?emails={comma-separated}`
- Approve: `PATCH /pto/{id}/manager?status=approved`
- Deny: `PATCH /pto/{id}/manager?status=denied`
- Can see HR Status column (read-only) — two-stage approval model

---

### HR Only

#### All-Employee Timesheet Viewer (`setupFormTimesheetViewer('hr', null)`)
- Loads: `GET /timesheets/form/all`
- Read-only — no approve/reject
- "Download Excel" link per submission: `GET /timesheets/form/{id}/excel` → xlsx blob download

#### Handbook Sign-Off Tracker (`setupHandbookSignoffViewer`)
- Loads: `GET /handbook/signoffs`
- Progress bar (% of staff signed) + table: Employee, Status badge, Signed Date

#### Policy Issuance & Tracking (`setupHRPolicies`)
- Issue: `POST /policies/issue` (multipart: `policy_name`, `issued_to` JSON array of emails, `file` PDF ≤4 MB)
- View issued: `GET /policies/issued` → each row has expandable "Track Acknowledgments" with per-employee status
- Policy PDF link: `GET /policies/issued/{issuanceId}/url`

#### Incident Reports (`setupHRIncidentReports`)
- Loads: `GET /incidents/`
- Read-only table: Employee, Submitted, Incident Date, Type badge, Location, Description

#### EFS Issuance & Tracking (`setupHREfs`)
- Issue: `POST /efs/issue` (multipart: `file` PDF/Word ≤10 MB, `issued_to` JSON array of emails)
- View issued: `GET /efs/issued` → each row expandable with per-employee signature status

#### All-Employee PTO Approvals (`setupPTOApprovals(null, 'hr')`)
- Loads: `GET /pto/all`
- Approve: `PATCH /pto/{id}/hr?status=approved`
- Deny: `PATCH /pto/{id}/hr?status=denied`
- Can see Manager Status column (read-only)

---

## Timesheet System

### Form Routing

`setupTimesheetForm` determines which timesheet type(s) to show each employee:

| Condition | Forms shown |
|---|---|
| `fkirmani@americloudtelecom.com` (IT admin) | All three |
| `directory[user].country` is India or Pakistan | Offshore only |
| `staff[user].state` is CA | California only |
| Any other state | Standard only |
| No state and not offshore | Both CA and Standard |

> Note: `state` field is not currently populated in `portal-config.js` staff entries, so US employees always see both CA and Standard forms.

### Standard Timesheet

**Template:** `AmeriCloud Telecom Weekly Timesheet - Standard 2026 6.xlsx`

**Header fields:** Employee Name (auto), Employee ID (auto), Department (auto from directory), Week Start (Sun picker), Week End (Sat — auto-calculated, read-only), Manager Name

**Daily columns:** Day, Date (auto), Project Code, Shift Start, Rest 1 Start, Rest 1 End, Meal Start, Meal End, Rest 2 Start, Rest 2 End, Shift End, Hours (calculated)

**Summary:** Total Hours Worked, Regular Hours (max 40/wk), Overtime 1.5× Hours (>40/wk), PTO / Sick / Holiday Hours (manual input)

**Hours formula:** `(Shift End − Shift Start) − Meal Duration`. Rest breaks are paid (not deducted). Meal is unpaid (deducted if both meal start and end are entered).

**Weekly OT:** Regular = min(total, 40). OT 1.5× = max(0, total − 40).

**Attestation:** "I certify that this timesheet accurately reflects all hours I worked during the week shown, including overtime, and that meal periods recorded were taken as shown."

**Waiver section:** Present — for voluntarily skipped meal/rest breaks.

---

### California Timesheet

**Template:** `AmeriCloud Telecom Weekly Timesheet - California 2026 5.xlsx`

**Header fields:** Same as Standard.

**Daily columns:** Same as Standard, plus **Meal Waiver** (Y/N dropdown after Meal End).

**Summary (OT breakdown table per day):** Day, Hours, Regular, OT 1.5×, OT 2.0×, Missed Meal, Missed Rest, Premium Hrs — plus weekly totals row.

**Additional summary section:** Total Hours Worked, Regular Hours (max 40/wk), OT 1.5× (daily 8–12 hrs), Weekly OT 1.5× (reg >40), **Total OT 1.5×** (subtotal), OT 2.0× (>12/day; 7th-day >8), PTO / Sick / Holiday Hours, Missed-Break Premium Hours (§226.7).

**Special control:** "Saturday is my 7th consecutive workday" checkbox — changes that day's OT rate.

**California OT rules (per day):**

| Hours worked | Rate |
|---|---|
| 0–8 | Regular |
| 8–12 | OT 1.5× |
| >12 | OT 2.0× |
| 7th consecutive day, 0–8h | OT 1.5× |
| 7th consecutive day, >8h | OT 2.0× |

Weekly OT 1.5× = max(0, sum(regular hours) − 40)

**Missed Meal premium (§226.7):** Flagged when hours > 5 AND meal waiver ≠ Y AND meal start or end is blank. One premium hour per occurrence.

**Missed Rest premium (§226.7):** Flagged when hours > 3.5 and rest 1 start/end is blank, OR hours > 6 and rest 2 start/end is blank. Max one premium hour per day.

**Attestation:** "I certify that this timesheet accurately reflects all hours I worked and all meal and rest periods I took, waived, or missed. I took every rest and meal break shown, free of duty and uninterrupted. On any day I marked Meal Waiver = Y, my shift was 6 hours or less and I voluntarily chose to waive my meal period. I was not pressured to under-report hours or over-report breaks."

**Waiver section:** California-specific language covering waiver of first meal period (≤6h shift) and second meal period (≤12h shift).

---

### Offshore Timesheet

**Template:** `AmeriCloud Telecom Weekly Timesheet - Offshore 2026 1.xlsx`

**For:** India and Pakistan employees only.

**Header fields:** Same as Standard (Manager field labelled "Manager Name (approval)").

**Daily columns:** Day, Date (auto), Code 1, Hrs 1, Code 2, Hrs 2, Code 3, Hrs 3, Start (reference), End (reference), Day Total (auto-calculated)

**Hours entry:** Entered directly as decimals in 0.5 increments. Start/End times are for manager reference only and do **not** drive calculations.

**Summary:** Total Hours Worked, Hours on Non-Billable Codes, Hours on Project Codes, Over 40 hours this week (Yes/No)

**Non-billable detection:** Any code beginning with `LBR` or `BDV`. Non-billable hours = sum of Hrs where the paired code starts with LBR or BDV. Project hours = Total − Non-billable.

**Attestation:** "I certify that this timesheet accurately reflects all hours I worked during the week shown, on the codes shown."

**No waiver section.**

---

### Timesheet Submission

All three types submit to `POST /timesheets/form/submit`.

**Common fields:**
```
ts_type            "standard" | "ca" | "offshore"
week_start         "YYYY-MM-DD"
employee_id        string
department         string
manager_name       string
pto_hours          number
employee_sig       string
employee_sig_date  "YYYY-MM-DD"
waiver_sig         string (empty for offshore)
waiver_sig_date    "YYYY-MM-DD" (empty for offshore)
days               array (7 items, one per day Sun–Sat)
```

**Standard/CA day object:**
```
date, day, projectCode, shiftStart, rest1Start, rest1End,
mealStart, mealEnd, mealWaiver, rest2Start, rest2End,
shiftEnd, hoursWorked, is7thDay
```

**Offshore day object:**
```
date, day, code1, hours1, code2, hours2, code3, hours3,
startTime, endTime, dayTotal
```

### Timesheet Interactive Tour

Both Standard/CA forms and the Offshore form have a "Quick tour" button. The tour uses a spotlight + tooltip overlay system (`ts-tour-overlay`, `ts-tour-spotlight`, `ts-tour-tooltip`).

- **Standard/CA tour:** 6 steps — Week Start, Project Code, Time Entry, Calculated Totals, Attestation, Waiver
- **Offshore tour:** 7 steps — Week Start, 3 Codes per day, Hours/increments, Non-billable codes, Start/End reference, Summary, Attestation

Each form type sets `_activeTourSteps` before calling `startTsTour()` so the correct sequence runs.

---

## API Reference

All requests: `Authorization: Bearer {MSAL idToken}`, `Content-Type: application/json` (unless multipart).

| Endpoint | Method | Description |
|---|---|---|
| `/handbook/sign` | POST | Record handbook acknowledgment |
| `/handbook/signoffs` | GET | All employee sign-off statuses (HR) |
| `/timesheets/form/submit` | POST | Submit a filled-out timesheet form |
| `/timesheets/form/all` | GET | All form submissions (HR) |
| `/timesheets/form/team?emails=` | GET | Team submissions (Manager) |
| `/timesheets/form/{id}` | GET | Single submission detail |
| `/timesheets/form/{id}/approve` | POST | Approve with `{ manager_sig, manager_sig_date }` |
| `/timesheets/form/{id}/reject` | POST | Reject submission |
| `/timesheets/form/{id}/excel` | GET | Download submission as xlsx |
| `/policies/my` | GET | Policies issued to signed-in employee |
| `/policies/issue` | POST | Issue policy PDF to employees (HR, multipart) |
| `/policies/issued` | GET | All issued policies with acknowledgment status (HR) |
| `/policies/issued/{id}/url` | GET | Signed URL to open policy PDF |
| `/policies/acknowledge/{id}` | POST | Acknowledge a policy |
| `/pto/` | POST | Submit PTO request `{ start_date, end_date, reason }` |
| `/pto/my` | GET | Signed-in employee's PTO requests |
| `/pto/team?emails=` | GET | Team's PTO requests (Manager) |
| `/pto/all` | GET | All PTO requests (HR) |
| `/pto/{id}/manager?status=` | PATCH | Manager approve/deny |
| `/pto/{id}/hr?status=` | PATCH | HR approve/deny |
| `/incidents/` | POST | Submit incident report |
| `/incidents/` | GET | All incident reports (HR) |
| `/efs/my` | GET | EFS docs issued to signed-in employee |
| `/efs/view/{id}` | GET | Signed URL to view EFS document |
| `/efs/sign/{id}` | POST | Sign an EFS |
| `/efs/issue` | POST | Issue EFS to employees (HR, multipart) |
| `/efs/issued` | GET | All issued EFS with signature status (HR) |
| `/trainings` | GET | Training docs for signed-in employee |
| `/trainings/upload` | POST | Upload training doc (multipart) |
| `/trainings/{id}` | PATCH | Update issuance/renewal dates |
| `/trainings/{id}` | DELETE | Delete training doc (admin only) |
| `/trainings/{id}/file` | GET | Download/view training doc file |

---

## Data Configuration (`portal-config.js`)

`PORTAL_CONFIG` is a global object with three keys. It is loaded before `portal.js` and used entirely client-side — no API calls read from it.

### `news` — Array
| Field | Type | Notes |
|---|---|---|
| `date` | string | ISO date (YYYY-MM-DD) |
| `category` | string | `'Company'` or `'Portal'` |
| `title` | string | |
| `body` | string | |
| `pinned` | boolean | Optional; pinned items sort first |

### `directory` — Array (~100 entries)
| Field | Type | Notes |
|---|---|---|
| `name` | string | Full name |
| `email` | string | Company email |
| `country` | string | `'US'`, `'India'`, or `'Pakistan'` |
| `department` | string | One of: BBU Swaps, Construction, Design, Integration, Managed Services, Optimization, Overhead, RF Data Collection, Other |

Used for: directory browser, timesheet type routing (offshore detection), training upload admin selector.

### `staff` — Object keyed by email (~110 entries)
| Field | Type | Notes |
|---|---|---|
| `name` | string | Full name |
| `role` | string | `'employee'`, `'manager'`, or `'hr'` |
| `team` | array | Emails of direct reports (managers only) |

Used for: policy/EFS employee pickers, manager team lookups, timesheet routing, display names.

> `state` and `manager` fields are referenced in timesheet routing and header pre-fill but are **not currently populated** in any staff entry. US employees will see both CA and Standard forms until `state` is added.

---

## Deployment

### Frontend
- **Host:** Vercel
- **URL:** `https://americloud-employee-portal.vercel.app`
- **Auto-deploy:** Every push to `master` branch triggers deployment
- **Build:** None — static files served directly

### Backend
- **Host:** Vercel
- **Framework:** FastAPI
- **URL:** `https://americloud-employee-portal-backend.vercel.app`
- **Database:** Supabase (PostgreSQL + storage)

### Local Development
- Frontend: `python3 -m http.server 8000` (serves from `americloud-employee-portal/`)
- Backend: `http://localhost:3000` (auto-detected by `API_BASE` check)

---

## Known Issues & TODOs

### Active gaps
- **`state` field missing from `portal-config.js` staff entries** — US employees always see both CA and Standard timesheet forms instead of being routed to the correct one. Add `state: "CA"` (or another state abbreviation) to each staff entry to enable proper routing.
- **Open enrollment banner** — `enrollment-banner` div in the Home section is hidden (`hidden` attribute) and contains a `[DATE]` placeholder. Activate it and fill in the date when enrollment opens.

### Dead code (safe to leave or remove)
- `setupTimesheetUpload()` and `setupTimesheetViewer()` — legacy file-upload timesheet functions, not called from `initDashboard`. Superseded by `setupFormTimesheetViewer()`.
- `getGraphToken()` — defined but never called. Was for Microsoft Graph `User.ReadBasic.All` scope.
- `IS_DEMO` variable and all `if (IS_DEMO)` branches — always `false` now that the dev bypass was removed. Branches are harmless dead code.

### README / docs drift
- `README.md` still references a SharePoint-based architecture (SharePoint list IDs, Power Automate, DocuSign PowerForm). These were from an earlier design and do not reflect the current Supabase + FastAPI implementation.

### Commented-out feature
- **Expense Approvals** — a list item in the Manager section is commented out in `index.html`. Not yet implemented.
