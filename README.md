# AmeriCloud Employee Portal

Internal employee self-service portal for AmeriCloud Telecom Solutions.
Live at: https://americloud-employee-portal.vercel.app

## Local preview

No build step required. Serve the repo root with Python:

    python3 -m http.server 8000

Then open http://localhost:8000 — you will see the sign-in screen.
To preview the dashboard: http://localhost:8000#demo

The `#demo` bypass shows all nav sections (including Manager Approvals and HR Dashboard) so the full layout is previewable. It does **not** show role-restricted content to real employees — roles are enforced by Azure AD group membership after real auth.

## Deployment

Vercel auto-deploys from this repo on every push to `master`.

## Handbook Sign-off Setup (DocuSign + SharePoint + Power Automate)

### Step A — Upload the handbook to SharePoint

1. In your SharePoint site, go to the **Documents** library (not Timesheets)
2. Upload the handbook PDF
3. Click the three dots (…) next to the file → **Copy link** → set it to "People in your organization can view"
4. Paste that URL into `portal-config.js` as the value of `handbookUrl`

### Step B — Create the DocuSign PowerForm

1. Sign in to [docusign.com](https://www.docusign.com) with your admin account
2. Go to **Templates** → create a new template → upload the handbook PDF
3. Add a signature field for **Signer 1** (the employee)
4. Save the template, then go to **PowerForms** → **Create PowerForm** → select your template
5. Under **Signer 1**, set the role name to `Signer` and check "Allow PowerForm senders to preset recipient information"
6. Save and copy the **PowerForm URL**
7. Paste that URL into `portal-config.js` as `docuSignPowerFormUrl`

The portal automatically appends `?Name={employee name}&Email={employee email}` to that URL so employees don't have to type their info.

### Step C — Create the sign-off tracking list in SharePoint

1. In your SharePoint site, click **New → List** → name it **Handbook Sign-offs**
2. Add these columns (exact names, case-sensitive):
   - `EmployeeEmail` — type: Single line of text
   - `Status` — type: Choice, options: **Signed**, **Pending**
   - `SignedDate` — type: Date and time
3. After creating the list, look at the URL in your browser — it will contain `lists/{id}`. Copy that ID.
4. Paste it into `portal-config.js` as `handbookSignoffListId`

### Step D — Set up Power Automate to auto-update the list

Power Automate is included in your M365 Business plan at no extra cost.

1. Go to [make.powerautomate.com](https://make.powerautomate.com) and sign in
2. Click **Create → Automated cloud flow**
3. Search for trigger: **DocuSign — When an envelope status changes**
   - Connect your DocuSign account when prompted
   - Set Status Filter to: `completed`
4. Add action: **SharePoint — Get items** → select your site → select **Handbook Sign-offs** list
   - Filter query: `EmployeeEmail eq '{triggerBody()?['signerEmail']}'`
5. Add a **Condition**: if the above returns any items → **Yes**: Update item (set Status = Signed, SignedDate = now) → **No**: Create item (EmployeeEmail = signer email, Status = Signed, SignedDate = now)
6. Save the flow and turn it on

After this is live: every time an employee completes signing in DocuSign, Power Automate writes their sign-off to the SharePoint list. The portal HR dashboard reads that list and shows the live status table.

---

## Timesheet Storage Setup (do this before launch)

Timesheets upload to a SharePoint document library via Microsoft Graph API.
Complete these steps in order once the portal is live on its subdomain.

### Step 1 — Create the SharePoint document library

1. Go to `https://americloudtelecom.sharepoint.com` → click **Create site** → choose **Team site**
2. Name it something like **AmeriCloud Portal** (internal, not public)
3. Inside the site, click **New → Document library** → name it **Timesheets**

### Step 2 — Add the Graph API permission in Entra ID

1. Go to [https://entra.microsoft.com](https://entra.microsoft.com) → **App registrations** → find your portal app
2. Click **API permissions → Add a permission → Microsoft Graph → Delegated permissions**
3. Search for `Sites.ReadWrite.All` → check it → click **Add permissions**
4. Click **Grant admin consent for AmeriCloud Telecom** → confirm Yes

### Step 3 — Find your SharePoint Site ID and Drive ID

Open **Graph Explorer** at [https://developer.microsoft.com/en-us/graph/graph-explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and sign in with your admin account.

Run this query (replace `EmployeePortal` with your actual site name):
```
GET https://graph.microsoft.com/v1.0/sites/americloudtelecom.sharepoint.com:/sites/EmployeePortal
```
Copy the `id` field from the response — that is your **Site ID**.

Then run:
```
GET https://graph.microsoft.com/v1.0/sites/{YOUR-SITE-ID}/drives
```
Find the drive named **Timesheets** in the response, copy its `id` — that is your **Drive ID**.

### Step 4 — Fill in portal-config.js

Open `portal-config.js` and replace:
- `YOUR-SHAREPOINT-SITE-ID` with the Site ID from Step 3
- `YOUR-SHAREPOINT-DRIVE-ID` with the Drive ID from Step 3
- Add every employee to the `staff` object (see examples in the file)

---

---

## Policy Issuance & Attestation Setup (SharePoint + Power Automate)

This system lets HR upload a policy PDF, select employees, and send DocuSign signing requests — all tracked in SharePoint.

### Step A — Create the SharePoint lists

Go to your SharePoint site and create two lists:

**List 1: Policy Issuances**
1. Click **New → List** → name it **Policy Issuances**
2. Add these columns (exact names, case-sensitive):
   - `PolicyName` — type: Single line of text
   - `PolicyUrl` — type: Single line of text
   - `IssuedTo` — type: Multiple lines of text (stores a JSON array of email strings)
   - `IssuedDate` — type: Date and time
3. Note the list ID from the URL: `…/lists/{id}/…`
4. Paste it into `portal-config.js` as `policyIssuancesListId`

**List 2: Policy Acknowledgments**
1. Click **New → List** → name it **Policy Acknowledgments**
2. Add these columns (exact names, case-sensitive):
   - `PolicyItemId` — type: Number (stores the ID from the Policy Issuances list)
   - `PolicyName` — type: Single line of text
   - `EmployeeEmail` — type: Single line of text
   - `EmployeeName` — type: Single line of text
   - `Status` — type: Choice, options: **Pending**, **Signed**
   - `SignedDate` — type: Date and time
3. Note the list ID from the URL
4. Paste it into `portal-config.js` as `policyAcknowledgmentsListId`

### Step B — Create a DocuSign template for policy documents

Each policy document uploaded by HR needs a DocuSign template to send for signature. The simplest approach is a general-purpose "sign this document" template.

1. In DocuSign admin, go to **Templates** → **New Template**
2. Name it something like "Policy Acknowledgment" — this template will be used for all policies
3. Leave the document blank (Power Automate will attach the actual policy PDF)
4. Add a signature field for **Signer 1**
5. Save the template and note the **Template ID** (shown in the URL)

### Step C — Create the Power Automate flow (DocuSign sending + status update)

This flow fires when HR issues a policy and sends DocuSign envelopes to each selected employee.

**Flow 1: Send DocuSign envelopes when a policy is issued**
1. Go to [make.powerautomate.com](https://make.powerautomate.com) and create **Automated cloud flow**
2. Trigger: **SharePoint — When an item is created** → select your site → select **Policy Issuances** list
3. Add action: **Initialize variable** → name: `emails`, type: Array → value: `json(triggerBody()?['IssuedTo'])`
4. Add action: **Apply to each** on the `emails` variable
5. Inside the loop, add: **DocuSign — Create and send envelope** → attach the policy PDF (use `PolicyUrl` from the trigger), set signer email = current item, name = look up from your staff list or just use email, use your template
6. Save and turn on the flow

**Flow 2: Update signing status when DocuSign completes**
1. Create another **Automated cloud flow**
2. Trigger: **DocuSign — When an envelope status changes** → Status Filter: `completed`
3. Add action: **SharePoint — Get items** from **Policy Acknowledgments** list
   - Filter query: `EmployeeEmail eq '${triggerBody()?['signerEmail']}' and Status eq 'Pending'`
4. Add action: **Apply to each** on the returned items
5. Inside: **SharePoint — Update item** → set `Status = Signed`, `SignedDate = utcNow()`
6. Save and turn on the flow

After this is live: the portal's HR "View Issued Policies" table will show green "Signatures Completed" once all employees have signed, and the employee's "My Policies" modal will update to show "✓ Signed" for each completed policy.

### Step D — Fill in portal-config.js

Add the two list IDs (from Step A) to `portal-config.js`:
```
policyIssuancesListId:       'paste-list-id-here',
policyAcknowledgmentsListId: 'paste-list-id-here',
```

---

## PTO Request Setup (SharePoint list)

PTO requests submitted by employees are stored in a SharePoint list. Managers read that list to approve or deny requests.

### Step A — Create the PTO Requests SharePoint list

1. In your SharePoint site, click **New → List** → name it **PTO Requests**
2. Add these columns (exact names, case-sensitive):
   - `EmployeeEmail` — type: Single line of text
   - `EmployeeName` — type: Single line of text
   - `StartDate` — type: Date and time *(date only, no time needed)*
   - `EndDate` — type: Date and time *(date only, no time needed)*
   - `Reason` — type: Multiple lines of text
   - `Status` — type: Choice, options: **Pending**, **Approved**, **Denied** *(set Pending as default)*
3. After creating the list, look at the URL in your browser — it will contain `lists/{id}`. Copy that ID.
4. Paste it into `portal-config.js` as `ptoRequestsListId`

### Step B — Grant the portal app permission

The portal uses the same `Sites.ReadWrite.All` Microsoft Graph permission already required for timesheets — no extra permissions needed if you completed the Timesheet Storage Setup steps.

### Step C — (Optional) Power Automate notification to manager

To notify a manager by email when one of their team members submits a PTO request:

1. Go to [make.powerautomate.com](https://make.powerautomate.com) and create an **Automated cloud flow**
2. Trigger: **SharePoint — When an item is created** → select your site → select **PTO Requests** list
3. Add action: **Office 365 Outlook — Send an email (V2)**
   - To: the manager's email address (you can hard-code it or look it up from a staff list)
   - Subject: `PTO Request from [EmployeeName]` — use the `EmployeeName` field from the trigger
   - Body: include `EmployeeName`, `StartDate`, `EndDate`, and `Reason` from the trigger fields
4. Save and turn on the flow

After this is live, managers will receive an email the moment an employee submits a request, in addition to seeing it in the portal's PTO Approvals table.

---

## Incident Report Setup (SharePoint list)

Safety and incident reports submitted by employees are stored in a SharePoint list. HR can view all submissions from the HR Dashboard.

### Step A — Create the Incident Reports SharePoint list

1. In your SharePoint site, click **New → List** → name it **Incident Reports**
2. Add these columns (exact names, case-sensitive):
   - `EmployeeEmail` — type: Single line of text
   - `EmployeeName` — type: Single line of text
   - `IncidentDate` — type: Date and time *(date only)*
   - `Location` — type: Single line of text
   - `IncidentType` — type: Choice, options: **Injury**, **Near Miss**, **Property Damage**, **Other**
   - `Description` — type: Multiple lines of text *(plain text)*
3. After creating the list, copy the list ID from the URL (`…/lists/{id}/…`)
4. Paste it into `portal-config.js` as `incidentReportsListId`

### Step B — (Optional) Email HR when a report is submitted

1. Go to [make.powerautomate.com](https://make.powerautomate.com) and create an **Automated cloud flow**
2. Trigger: **SharePoint — When an item is created** → select your site → select **Incident Reports** list
3. Add action: **Office 365 Outlook — Send an email (V2)**
   - To: HR's email address
   - Subject: `Incident Report — [IncidentType] — [EmployeeName]`
   - Body: include all fields (EmployeeName, IncidentDate, Location, IncidentType, Description)
4. Save and turn on the flow

---

## Timesheet → SharePoint Inbox Upload Setup

When an employee submits a timesheet, the backend generates the filled Excel and uploads it to:

    Timesheets/{Department}/{Country}/Inbox/{EmpID} - {Name} - Week Ending mm-dd-yyyy.xlsx

Department comes from the form's Department field; Country comes from the employee's `country` in the `portal-config.js` directory. The upload only happens if both match an existing folder (BBU Swaps, Construction, Design, Exempt, Integration, Optimization, Overhead, RF Data Collection / US, Pakistan, India). Otherwise it is skipped, so stray folders never get created. A failed upload never blocks the submission. Resubmitting the same week replaces the file in Inbox.

Until all the environment variables below are set, the upload is turned off and nothing else changes.

### Step 1 — Add the Graph application permission

1. Go to [https://entra.microsoft.com](https://entra.microsoft.com) → **App registrations** → open the portal app (the one whose Client ID is in `AZURE_CLIENT_ID`)
2. **API permissions → Add a permission → Microsoft Graph → Application permissions** (not Delegated)
3. Search `Files.ReadWrite.All` → check it → **Add permissions**
4. Click **Grant admin consent for AmeriCloud Telecom** → Yes. The status column should show a green check.

### Step 2 — Create a client secret

1. Same app registration → **Certificates & secrets → Client secrets → New client secret**
2. Description: `portal-api-sharepoint`, Expires: 24 months → **Add**
3. Copy the **Value** column right away. It is only shown once. This is `AZURE_CLIENT_SECRET`.
4. Set a calendar reminder to rotate it before it expires. Uploads will stop silently when it does.

### Step 3 — Get the Site ID and Drive ID

1. Open [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and sign in with your admin account
2. Open the Timesheets library in SharePoint and note the site part of the URL, e.g. `https://americloudtelecom.sharepoint.com/sites/ACTelecomSolutions/...`
3. Run (replace the site path with yours):
   ```
   GET https://graph.microsoft.com/v1.0/sites/americloudtelecom.sharepoint.com:/sites/ACTelecomSolutions
   ```
   Copy the `id` field. This is `SP_SITE_ID`.
4. Run:
   ```
   GET https://graph.microsoft.com/v1.0/sites/{SP_SITE_ID}/drives
   ```
   Find the drive whose `webUrl` ends in the library that contains the department folders (it may be named "Timesheets" or "Documents", depending on whether Timesheets is its own library or a folder inside HR). Copy its `id`. This is `SP_DRIVE_ID`.
5. Run `GET https://graph.microsoft.com/v1.0/drives/{SP_DRIVE_ID}/root/children`.
   - If the department folders (BBU Swaps, Construction, …) are listed directly, leave `SP_BASE_PATH` unset.
   - If you see a parent folder instead (e.g. `Timesheets`, or `HR`), open it with `GET https://graph.microsoft.com/v1.0/drives/{SP_DRIVE_ID}/root:/Timesheets:/children` and keep going until you reach the folder that contains the department folders. That path (e.g. `Timesheets` or `HR/Timesheets`) is your `SP_BASE_PATH`.

### Step 4 — Add environment variables in Vercel (backend project)

Vercel → **americloud-portal-api** project → **Settings → Environment Variables** → add for Production:

| Name | Value |
|---|---|
| `AZURE_CLIENT_SECRET` | secret value from Step 2 |
| `SP_SITE_ID` | from Step 3.3 |
| `SP_DRIVE_ID` | from Step 3.4 |
| `SP_BASE_PATH` | only if needed, from Step 3.5 (e.g. `Timesheets`) |

(`AZURE_TENANT_ID` and `AZURE_CLIENT_ID` are already set.) Then go to **Deployments → ⋯ → Redeploy** so the new variables take effect.

### Step 5 — Test

Submit a timesheet as an employee whose department has a SharePoint folder (e.g. a BBU Swaps employee). The file should appear in `BBU Swaps/{Country}/Inbox` within a few seconds. If it doesn't, check Vercel → backend project → **Logs** for a line starting `SharePoint upload`.

---

## Before launch checklist

- [ ] Complete handbook sign-off setup (see "Handbook Sign-off Setup" above)
- [ ] Set `handbookUrl`, `docuSignPowerFormUrl`, `handbookSignoffListId` in `portal-config.js`
- [ ] Complete SharePoint setup (see "Timesheet Storage Setup" above)
- [ ] Complete policy issuance setup (see "Policy Issuance & Attestation Setup" above)
- [ ] Set `policyIssuancesListId`, `policyAcknowledgmentsListId` in `portal-config.js`
- [ ] Complete PTO Request setup (see "PTO Request Setup" above)
- [ ] Set `ptoRequestsListId` in `portal-config.js`
- [ ] Complete Incident Report setup (see "Incident Report Setup" above)
- [ ] Set `incidentReportsListId` in `portal-config.js`
- [ ] Fill in `portal-config.js`: SharePoint IDs + full staff roster
- [ ] Replace all `href="#"` placeholder links with real URLs
- [ ] Fill in contact info in the footer (HR, payroll, timesheet, IT emails)
- [ ] Set `MSAL_CLIENT_ID` and `MSAL_TENANT_ID` in `portal.js`
- [ ] Set `GROUP_MANAGERS` and `GROUP_HR` object IDs in `portal.js`
- [ ] Remove `#demo` bypass block from `portal.js` (the top `if` block)
- [ ] Confirm SSL / HTTPS is working on Vercel
- [ ] Test sign-in with a real M365 account on the live subdomain
- [ ] Test role-gated nav: Manager sees "Manager Approvals", HR sees "HR Dashboard"
- [ ] Test timesheet upload as a regular employee
- [ ] Test timesheet viewer as a manager (should see only their team)
- [ ] Test timesheet viewer as HR (should see everyone)
