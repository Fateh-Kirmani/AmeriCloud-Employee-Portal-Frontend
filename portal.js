/* global msal, PORTAL_CONFIG */
(function () {
  'use strict';

  // ── Module-level state ────────────────────────────────────────────────────
  var IS_DEMO      = false;
  var msalInstance;           // set after MSAL init (not in demo mode)
  var currentAccount;         // set after successful sign-in

  // ── Demo mock handbook sign-offs ─────────────────────────────────────────
  var DEMO_SIGNOFFS = [
    { email: 'alex.rivera@americloudtelecom.com',    name: 'Alex Rivera',    status: 'Signed',  signedDate: '2026-07-08T14:22:00Z' },
    { email: 'maria.santos@americloudtelecom.com',   name: 'Maria Santos',   status: 'Signed',  signedDate: '2026-07-09T09:05:00Z' },
    { email: 'jordan.wallace@americloudtelecom.com', name: 'Jordan Wallace', status: 'Pending', signedDate: null },
    { email: 'chris.nguyen@americloudtelecom.com',   name: 'Chris Nguyen',   status: 'Pending', signedDate: null },
    { email: 'sam.patel@americloudtelecom.com',      name: 'Sam Patel',      status: 'Signed',  signedDate: '2026-07-10T11:30:00Z' },
  ];

  // ── Demo mock timesheets ──────────────────────────────────────────────────
  var DEMO_TIMESHEETS = [
    { name: 'Alex Rivera',    email: 'alex.rivera@americloudtelecom.com',    filename: '2026-07-10 - Alex Rivera - Standard Timesheet.xlsx',   date: '2026-07-10T12:34:00Z', downloadUrl: '#' },
    { name: 'Maria Santos',   email: 'maria.santos@americloudtelecom.com',   filename: '2026-07-10 - Maria Santos - CA Timesheet.xlsx',         date: '2026-07-10T09:15:00Z', downloadUrl: '#' },
    { name: 'Jordan Wallace', email: 'jordan.wallace@americloudtelecom.com', filename: '2026-07-10 - Jordan Wallace - Standard Timesheet.xlsx', date: '2026-07-10T08:50:00Z', downloadUrl: '#' },
    { name: 'Alex Rivera',    email: 'alex.rivera@americloudtelecom.com',    filename: '2026-07-03 - Alex Rivera - Standard Timesheet.xlsx',    date: '2026-07-03T12:00:00Z', downloadUrl: '#' },
    { name: 'Maria Santos',   email: 'maria.santos@americloudtelecom.com',   filename: '2026-07-03 - Maria Santos - CA Timesheet.xlsx',         date: '2026-07-03T09:00:00Z', downloadUrl: '#' },
  ];

  // ── Demo mock staff (employee picker) ─────────────────────────────────────
  var DEMO_STAFF = [
    { email: 'alex.rivera@americloudtelecom.com',    name: 'Alex Rivera' },
    { email: 'chris.nguyen@americloudtelecom.com',   name: 'Chris Nguyen' },
    { email: 'jordan.wallace@americloudtelecom.com', name: 'Jordan Wallace' },
    { email: 'maria.santos@americloudtelecom.com',   name: 'Maria Santos' },
    { email: 'sam.patel@americloudtelecom.com',      name: 'Sam Patel' },
  ];

  // ── Demo mock issued policies ─────────────────────────────────────────────
  var DEMO_ISSUED_POLICIES = [
    {
      id: '1', name: 'Remote Work Policy 2026',
      issuedDate: '2026-07-01T00:00:00Z',
      issuedTo: ['alex.rivera@americloudtelecom.com', 'chris.nguyen@americloudtelecom.com',
                 'jordan.wallace@americloudtelecom.com', 'maria.santos@americloudtelecom.com',
                 'sam.patel@americloudtelecom.com']
    },
    {
      id: '2', name: 'Safety & PPE Guidelines',
      issuedDate: '2026-07-05T00:00:00Z',
      issuedTo: ['alex.rivera@americloudtelecom.com', 'jordan.wallace@americloudtelecom.com',
                 'sam.patel@americloudtelecom.com']
    }
  ];

  // ── Demo mock policy acknowledgments ─────────────────────────────────────
  var DEMO_POLICY_ACKS = [
    { policyId: '1', email: 'alex.rivera@americloudtelecom.com',    name: 'Alex Rivera',    status: 'Signed',  signedDate: '2026-07-02T14:00:00Z' },
    { policyId: '1', email: 'chris.nguyen@americloudtelecom.com',   name: 'Chris Nguyen',   status: 'Pending', signedDate: null },
    { policyId: '1', email: 'jordan.wallace@americloudtelecom.com', name: 'Jordan Wallace', status: 'Signed',  signedDate: '2026-07-03T09:00:00Z' },
    { policyId: '1', email: 'maria.santos@americloudtelecom.com',   name: 'Maria Santos',   status: 'Pending', signedDate: null },
    { policyId: '1', email: 'sam.patel@americloudtelecom.com',      name: 'Sam Patel',      status: 'Signed',  signedDate: '2026-07-02T16:30:00Z' },
    { policyId: '2', email: 'alex.rivera@americloudtelecom.com',    name: 'Alex Rivera',    status: 'Signed',  signedDate: '2026-07-06T11:00:00Z' },
    { policyId: '2', email: 'jordan.wallace@americloudtelecom.com', name: 'Jordan Wallace', status: 'Pending', signedDate: null },
    { policyId: '2', email: 'sam.patel@americloudtelecom.com',      name: 'Sam Patel',      status: 'Signed',  signedDate: '2026-07-07T10:00:00Z' }
  ];

  // ── Demo mock incident reports ────────────────────────────────────────────
  var DEMO_INCIDENT_REPORTS = [
    { id: '1', email: 'alex.rivera@americloudtelecom.com',  name: 'Alex Rivera',  incidentDate: '2026-08-18', location: 'Warehouse B',  incidentType: 'Near Miss',        description: 'Forklift passed within 2 feet of me while I was walking in the aisle. No injury, but the area was poorly lit and the forklift operator did not see me.', submittedDate: '2026-08-18T15:42:00Z' },
    { id: '2', email: 'sam.patel@americloudtelecom.com',    name: 'Sam Patel',    incidentDate: '2026-08-15', location: 'Main Office',  incidentType: 'Property Damage',  description: 'Company laptop fell from desk and screen cracked. Occurred during a Teams call when the power adapter cable was accidentally pulled.', submittedDate: '2026-08-15T11:20:00Z' },
    { id: '3', email: 'maria.santos@americloudtelecom.com', name: 'Maria Santos', incidentDate: '2026-08-10', location: 'Job Site 7',   incidentType: 'Injury',           description: 'Slipped on wet concrete while exiting the site trailer. Minor sprain to left ankle. First aid applied on site. Did not require hospital visit.', submittedDate: '2026-08-10T17:05:00Z' },
  ];

  // ── Demo mock PTO requests ────────────────────────────────────────────────
  var DEMO_PTO_REQUESTS = [
    { id: '1', email: 'alex.rivera@americloudtelecom.com',    name: 'Alex Rivera',    startDate: '2026-08-04', endDate: '2026-08-08', reason: 'Family vacation',                status: 'Pending',  submittedDate: '2026-07-25T10:00:00Z' },
    { id: '2', email: 'jordan.wallace@americloudtelecom.com', name: 'Jordan Wallace', startDate: '2026-07-31', endDate: '2026-08-01', reason: 'Moving to a new apartment',       status: 'Pending',  submittedDate: '2026-07-26T14:30:00Z' },
    { id: '3', email: 'chris.nguyen@americloudtelecom.com',   name: 'Chris Nguyen',   startDate: '2026-08-11', endDate: '2026-08-14', reason: 'Medical appointment and recovery', status: 'Approved', submittedDate: '2026-07-20T09:00:00Z' },
  ];

  // ── Dev bypass — visit index.html#demo to preview without real auth ───────
  // REMOVE this entire block before announcing the portal to employees.
  if (window.location.hash === '#demo') {
    IS_DEMO = true;
    document.getElementById('signin-view').hidden    = true;
    document.getElementById('dashboard-view').hidden = false;
    initDashboard(null);
    return;
  }
  // ── End dev bypass ────────────────────────────────────────────────────────

  // ── MSAL config ───────────────────────────────────────────────────────────
  var MSAL_CLIENT_ID = '64cfff00-2deb-4a7a-91fc-74431baac966';
  var MSAL_TENANT_ID = 'f8e7955f-b183-4786-b8ca-7a51ecf001e8';
  var GROUP_MANAGERS = '5ef854d4-0938-4d06-87a6-3643b9a92796';
  var GROUP_HR       = '6f5e8021-9bcd-44c8-8a66-cff4014d5e07';

  var msalConfig = {
    auth: {
      clientId:    MSAL_CLIENT_ID,
      authority:   'https://login.microsoftonline.com/' + MSAL_TENANT_ID,
      redirectUri: window.location.origin
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
  };

  var loginRequest = { scopes: ['User.Read', 'openid', 'profile'] };
  var API_BASE = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1')
    ? 'http://localhost:3000'
    : 'https://americloud-employee-portal-backend.vercel.app';
  function isConfigured() { return true; }
  msalInstance = new msal.PublicClientApplication(msalConfig);

  // ── View helpers ──────────────────────────────────────────────────────────
  function showDashboard(account) {
    currentAccount = account;
    document.getElementById('signin-view').hidden    = true;
    document.getElementById('dashboard-view').hidden = false;
    initDashboard(account);
  }

  function showSignIn() {
    document.getElementById('signin-view').hidden    = false;
    document.getElementById('dashboard-view').hidden = true;
  }

  // ── Boot: MSAL v3 requires initialize() before any other call ─────────────
  msalInstance.initialize().then(function () {
    // Auth buttons (wired after initialize so loginRedirect works)
    document.getElementById('btn-signin').addEventListener('click', function () {
      msalInstance.loginRedirect(loginRequest);
    });

    document.getElementById('btn-signout').addEventListener('click', function () {
      msalInstance.logoutRedirect({ postLogoutRedirectUri: window.location.origin });
    });

    return msalInstance.handleRedirectPromise();
  }).then(function () {
    var accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) { showDashboard(accounts[0]); } else { showSignIn(); }
  }).catch(function (err) { console.error('MSAL error:', err); showSignIn(); });


  // ══════════════════════════════════════════════════════════════════════════
  // Dashboard init — called once after auth (or in demo mode)
  // ══════════════════════════════════════════════════════════════════════════
  function initDashboard(account) {
    var userEmail       = account ? account.username.toLowerCase() : 'demo@americloudtelecom.com';
    var userDisplayName = account ? (account.name || account.username) : 'Demo User';

    // ── Role detection ────────────────────────────────────────────────────
    var isManager = false;
    var isHR      = false;

    if (account) {
      var groups = (account.idTokenClaims && account.idTokenClaims.groups) || [];
      if (groups.indexOf(GROUP_HR) !== -1) {
        isHR = true;
        document.getElementById('nav-hr').hidden = false;
      }
      if (groups.indexOf(GROUP_MANAGERS) !== -1) {
        isManager = true;
        document.getElementById('nav-manager').hidden = false;
      }
    } else {
      // Demo: show all role nav items
      isManager = true;
      isHR      = true;
      document.getElementById('nav-manager').hidden = false;
      document.getElementById('nav-hr').hidden      = false;
    }

    // ── Section switching ─────────────────────────────────────────────────
    var navLinks    = document.querySelectorAll('.portal-nav__link[data-section]');
    var allSections = document.querySelectorAll('.portal-section[id^="section-"]');

    function activateSection(id) {
      allSections.forEach(function (s) { s.hidden = s.id !== 'section-' + id; });
      navLinks.forEach(function (l) { l.classList.toggle('is-active', l.dataset.section === id); });
      document.querySelector('.portal-content').scrollTo(0, 0);
    }

    navLinks.forEach(function (link) {
      link.addEventListener('click', function (e) {
        e.preventDefault();
        activateSection(link.dataset.section);
        closeSidebar();
      });
    });

    document.querySelectorAll('[data-goto]').forEach(function (el) {
      el.addEventListener('click', function (e) { e.preventDefault(); activateSection(el.dataset.goto); });
    });

    activateSection('home');

    // ── Mobile sidebar ────────────────────────────────────────────────────
    var sidebar = document.getElementById('portal-sidebar');
    var overlay = document.getElementById('portal-overlay');

    document.getElementById('burger').addEventListener('click', function () {
      sidebar.classList.add('is-open');
      overlay.classList.add('is-visible');
      document.body.style.overflow = 'hidden';
    });
    overlay.addEventListener('click', closeSidebar);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeSidebar(); });

    function closeSidebar() {
      sidebar.classList.remove('is-open');
      overlay.classList.remove('is-visible');
      document.body.style.overflow = '';
    }

    // ── Handbook ──────────────────────────────────────────────────────────
    setupHandbook(userEmail, userDisplayName);

    // ── Timesheet upload ──────────────────────────────────────────────────
    setupTimesheetUpload(userEmail, userDisplayName);

    // ── Timesheet viewer ──────────────────────────────────────────────────
    var staffEntry = PORTAL_CONFIG.staff[userEmail] || null;

    if (isManager) {
      var team = (staffEntry && staffEntry.team) ? staffEntry.team : [];
      // In demo mode use null so mock data shows for any team
      setupTimesheetViewer('manager', IS_DEMO ? null : team);
      setupPTOApprovals(IS_DEMO ? null : team, 'manager');
    }

    if (isHR) {
      var allEmails = Object.keys(PORTAL_CONFIG.staff);
      setupTimesheetViewer('hr', IS_DEMO ? null : allEmails);
      setupHandbookSignoffViewer();
      setupHRPolicies();
      setupHRIncidentReports();
      setupPTOApprovals(null, 'hr');
    }

    // Available to all roles — shows policies issued to the signed-in employee
    setupEmployeePolicies(userEmail);

    // Available to all roles — PTO request form
    setupPTORequest(userEmail, userDisplayName);

    // Available to all roles — my submitted PTO status
    setupMyPTORequests();

    // Available to all roles — incident report form
    setupIncidentReport(userEmail, userDisplayName);

    // Available to all roles — company directory
    setupDirectory();
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Handbook — view link + DocuSign sign button
  // ══════════════════════════════════════════════════════════════════════════
  function setupHandbook(userEmail, userDisplayName) {
    var signBtn = document.getElementById('btn-sign-handbook');
    if (!signBtn) return;

    signBtn.addEventListener('click', function () {
      if (IS_DEMO) {
        alert('Demo mode: clicking this would record your handbook acknowledgment.');
        return;
      }

      signBtn.disabled = true;
      signBtn.textContent = 'Signing…';

      apiCall('/handbook/sign', { method: 'POST' })
        .then(function () {
          signBtn.textContent = '✓ Signed';
          signBtn.classList.add('portal-btn--success');
        })
        .catch(function (err) {
          console.error('Handbook sign error:', err);
          signBtn.disabled = false;
          signBtn.textContent = 'Sign Handbook';
          alert('Could not record your signature. Please try again or contact IT.');
        });
    });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // HR — Handbook sign-off tracker
  // ══════════════════════════════════════════════════════════════════════════
  function setupHandbookSignoffViewer() {
    var btn    = document.getElementById('btn-view-handbook-signoffs');
    var panel  = document.getElementById('panel-handbook-signoffs');
    var bodyEl = document.getElementById('body-handbook-signoffs');
    var refresh = document.getElementById('btn-refresh-handbook-signoffs');
    if (!btn || !panel || !bodyEl) return;

    var loaded = false;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Hide Sign-offs';
        if (!loaded) { loadAndRender(); loaded = true; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View Sign-offs';
      }
    });

    if (refresh) {
      refresh.addEventListener('click', function () { loaded = false; loadAndRender(); loaded = true; });
    }

    function loadAndRender() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading sign-off status…</p>';
      loadHandbookSignoffs()
        .then(function (rows) { renderSignoffTable(rows, bodyEl); })
        .catch(function (err) {
          console.error('Signoff load error:', err);
          bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load sign-off data. Click Refresh to try again.</p>';
        });
    }
  }

  function loadHandbookSignoffs() {
    if (IS_DEMO) {
      return Promise.resolve(DEMO_SIGNOFFS.slice().sort(function (a, b) {
        if (a.status === b.status) return a.name.localeCompare(b.name);
        return a.status === 'Signed' ? -1 : 1;
      }));
    }

    return apiCall('/handbook/signoffs').then(function (rows) {
      return rows.map(function (r) {
        return { name: r.employee_name, email: r.employee_email, status: r.status, signedDate: r.signed_date };
      }).sort(function (a, b) {
        if (a.status === b.status) return a.name.localeCompare(b.name);
        return a.status === 'Signed' ? -1 : 1;
      });
    });
  }

  function renderSignoffTable(rows, bodyEl) {
    if (!rows || rows.length === 0) {
      var notSetup = !isConfigured()
        || !PORTAL_CONFIG.handbookSignoffListId
        || PORTAL_CONFIG.handbookSignoffListId === 'YOUR-HANDBOOK-SIGNOFF-LIST-ID';
      bodyEl.innerHTML = notSetup
        ? '<p class="portal-ts-empty">Sign-off tracking is not configured yet. See README for setup instructions.</p>'
        : '<p class="portal-ts-empty">No employees found. Fill in the staff roster in portal-config.js.</p>';
      return;
    }

    var signed  = rows.filter(function (r) { return r.status === 'Signed'; }).length;
    var total   = rows.length;
    var pct     = Math.round((signed / total) * 100);

    var html = '<div class="portal-signoff-summary">'
      + '<span class="portal-signoff-count">' + signed + ' of ' + total + ' signed</span>'
      + '<div class="portal-signoff-bar"><div class="portal-signoff-bar__fill" style="width:' + pct + '%"></div></div>'
      + '</div>'
      + '<div class="portal-ts-wrap"><table class="portal-ts-table">'
      + '<thead><tr><th>Employee</th><th>Status</th><th>Signed</th></tr></thead><tbody>';

    rows.forEach(function (row) {
      var badge = row.status === 'Signed'
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">&#10003; Signed</span>'
        : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';
      var date = row.signedDate ? formatDate(row.signedDate) : '—';

      html += '<tr>'
        + '<td class="portal-ts-col-name">' + escapeHtml(row.name) + '<br><span class="portal-signoff-email">' + escapeHtml(row.email) + '</span></td>'
        + '<td>' + badge + '</td>'
        + '<td class="portal-ts-col-date">' + date + '</td>'
        + '</tr>';
    });

    html += '</tbody></table></div>';
    bodyEl.innerHTML = html;
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Employee — Policies & Acknowledgments modal
  // ══════════════════════════════════════════════════════════════════════════
  function setupEmployeePolicies(userEmail) {
    var openBtn  = document.getElementById('btn-open-policies-modal');
    var modal    = document.getElementById('modal-policies');
    var overlay  = document.getElementById('modal-policies-overlay');
    var closeBtn = document.getElementById('btn-close-policies-modal');
    var bodyEl   = document.getElementById('body-policies-modal');
    if (!openBtn || !modal || !bodyEl) return;

    var loaded = false;

    function openModal() {
      modal.hidden = false;
      document.body.style.overflow = 'hidden';
      if (!loaded) { loadAndRender(); loaded = true; }
    }

    function closeModal() {
      modal.hidden = true;
      document.body.style.overflow = '';
    }

    openBtn.addEventListener('click', openModal);
    if (closeBtn) closeBtn.addEventListener('click', closeModal);
    if (overlay)  overlay.addEventListener('click', closeModal);
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && !modal.hidden) closeModal();
    });

    function loadAndRender() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading your policies…</p>';
      loadPoliciesForEmployee(userEmail)
        .then(function (rows) { renderEmployeePoliciesModal(rows, bodyEl); })
        .catch(function (err) {
          console.error('Policy load error:', err);
          bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load policies. Try again later.</p>';
        });
    }
  }


  function loadPoliciesForEmployee(userEmail) {
    if (IS_DEMO) {
      var demoAcks = DEMO_POLICY_ACKS.filter(function (a) { return a.email === 'alex.rivera@americloudtelecom.com'; });
      return Promise.resolve(DEMO_ISSUED_POLICIES.map(function (p) {
        var ack = null;
        for (var i = 0; i < demoAcks.length; i++) { if (demoAcks[i].policyId === p.id) { ack = demoAcks[i]; break; } }
        return { id: p.id, name: p.name, issuedDate: p.issuedDate, status: ack ? ack.status : 'Pending', signedDate: ack ? ack.signedDate : null };
      }));
    }

    return apiCall('/policies/my').then(function (rows) {
      return rows.map(function (r) {
        return { id: r.id, name: r.policy_name, issuedDate: r.issued_date, status: r.status, signedDate: r.signed_date };
      });
    });
  }


  function renderEmployeePoliciesModal(rows, bodyEl) {
    if (!rows || rows.length === 0) {
      bodyEl.innerHTML = '<p class="portal-ts-empty">No policies have been issued to you yet.</p>';
      return;
    }

    var html = '<div class="portal-ts-wrap"><table class="portal-ts-table">'
      + '<thead><tr><th>Policy</th><th>Issued</th><th>Status</th><th></th></tr></thead><tbody>';

    rows.forEach(function (row) {
      var badge = row.status === 'Signed'
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">&#10003; Signed</span>'
        : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';

      var action = row.status !== 'Signed'
        ? '<button class="portal-ts-link portal-pol-ack-btn" data-ack-id="' + escapeHtml(String(row.id)) + '" type="button">Acknowledge &#10003;</button>'
        : '';

      html += '<tr>'
        + '<td class="portal-ts-col-name">' + escapeHtml(row.name) + '</td>'
        + '<td class="portal-ts-col-date">' + (row.issuedDate ? formatDate(row.issuedDate) : '—') + '</td>'
        + '<td id="pol-status-' + escapeHtml(String(row.id)) + '">' + badge + '</td>'
        + '<td class="portal-pol-action">' + action + '</td>'
        + '</tr>';
    });

    html += '</tbody></table></div>';
    bodyEl.innerHTML = html;

    bodyEl.querySelectorAll('.portal-pol-ack-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        btn.disabled = true;
        apiCall('/policies/acknowledge/' + btn.dataset.ackId, { method: 'POST' })
          .then(function () {
            var cell = document.getElementById('pol-status-' + btn.dataset.ackId);
            if (cell) cell.innerHTML = '<span class="portal-signoff-badge portal-signoff-badge--signed">&#10003; Signed</span>';
            btn.outerHTML = '';
          })
          .catch(function (err) {
            console.error('Acknowledge error:', err);
            btn.disabled = false;
          });
      });
    });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // HR — Policy Issuance & Attestation Tracking
  // ══════════════════════════════════════════════════════════════════════════

  // Tracks checked emails across search re-renders (module-scoped for this session)
  var _pickerSelected = [];

  function setupHRPolicies() {
    var issueBtn  = document.getElementById('btn-issue-policy');
    var viewBtn   = document.getElementById('btn-view-issued-policies');
    var issuePanel = document.getElementById('panel-issue-policy');
    var viewPanel  = document.getElementById('panel-view-policies');
    if (!issueBtn || !viewBtn || !issuePanel || !viewPanel) return;

    var viewLoaded = false;

    // Wire file label text update once
    var policyFileInput = document.getElementById('policy-file');
    if (policyFileInput) {
      policyFileInput.addEventListener('change', function () {
        var labelText = document.getElementById('policy-file-label-text');
        if (labelText && this.files.length) labelText.textContent = this.files[0].name;
      });
    }

    // Wire search input once — re-renders list with filter while preserving selections
    var searchInput = document.getElementById('policy-emp-search');
    if (searchInput) {
      searchInput.addEventListener('input', function () { buildEmployeePicker(this.value); });
    }

    issueBtn.addEventListener('click', function () {
      if (issuePanel.hidden) {
        issuePanel.hidden = false;
        viewPanel.hidden  = true;
        viewBtn.textContent = 'View Issued Policies';
        // Reset picker state when opening
        _pickerSelected = [];
        if (searchInput) searchInput.value = '';
        buildEmployeePicker('');
      } else {
        issuePanel.hidden = true;
      }
    });

    viewBtn.addEventListener('click', function () {
      if (viewPanel.hidden) {
        viewPanel.hidden  = false;
        issuePanel.hidden = true;
        issueBtn.textContent = 'Issue a Policy';
        if (!viewLoaded) { loadAndRenderIssuedPolicies(); viewLoaded = true; }
      } else {
        viewPanel.hidden = true;
      }
    });

    var refreshBtn = document.getElementById('btn-refresh-issued-policies');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', function () { viewLoaded = false; loadAndRenderIssuedPolicies(); viewLoaded = true; });
    }

    var form = document.getElementById('form-issue-policy');
    if (form) {
      form.addEventListener('submit', function (e) { e.preventDefault(); handlePolicyIssue(); });
    }

    function loadAndRenderIssuedPolicies() {
      var bodyEl = document.getElementById('body-issued-policies');
      if (!bodyEl) return;
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading issued policies…</p>';
      loadIssuedPolicies()
        .then(function (data) { renderIssuedPoliciesTable(data, bodyEl); })
        .catch(function (err) {
          console.error('Issued policies load error:', err);
          bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load policies. Click Refresh to try again.</p>';
        });
    }
  }


  function buildEmployeePicker(filter) {
    var listEl = document.getElementById('policy-emp-list');
    if (!listEl) return;

    var employees;
    if (IS_DEMO) {
      employees = DEMO_STAFF;
    } else {
      employees = Object.keys(PORTAL_CONFIG.staff).map(function (email) {
        return { email: email, name: PORTAL_CONFIG.staff[email].name || email };
      }).sort(function (a, b) { return a.name.localeCompare(b.name); });
    }

    if (!employees.length) {
      listEl.innerHTML = '<p class="portal-picker-empty">No employees configured. Fill in the staff roster in portal-config.js.</p>';
      return;
    }

    var term = (filter || '').toLowerCase();
    var filtered = term
      ? employees.filter(function (e) {
          return e.name.toLowerCase().indexOf(term) !== -1 || e.email.toLowerCase().indexOf(term) !== -1;
        })
      : employees;

    var allChecked = filtered.length > 0 && filtered.every(function (e) { return _pickerSelected.indexOf(e.email) !== -1; });

    var html = '<label class="portal-picker-item portal-picker-selectall">'
      + '<input type="checkbox" id="picker-select-all" class="portal-picker-cb"' + (allChecked ? ' checked' : '') + '> '
      + '<span>Select All (' + filtered.length + ')</span></label>';

    filtered.forEach(function (emp) {
      var checked = _pickerSelected.indexOf(emp.email) !== -1 ? ' checked' : '';
      html += '<label class="portal-picker-item">'
        + '<input type="checkbox" class="picker-emp-checkbox portal-picker-cb" value="' + escapeHtml(emp.email) + '"' + checked + '> '
        + '<span class="portal-picker-name">' + escapeHtml(emp.name) + '</span>'
        + '<span class="portal-picker-email">' + escapeHtml(emp.email) + '</span>'
        + '</label>';
    });
    listEl.innerHTML = html;

    var selectAll = document.getElementById('picker-select-all');
    if (selectAll) {
      selectAll.addEventListener('change', function () {
        filtered.forEach(function (emp) {
          var idx = _pickerSelected.indexOf(emp.email);
          if (selectAll.checked && idx === -1) _pickerSelected.push(emp.email);
          else if (!selectAll.checked && idx !== -1) _pickerSelected.splice(idx, 1);
        });
        listEl.querySelectorAll('.picker-emp-checkbox').forEach(function (cb) { cb.checked = selectAll.checked; });
      });
    }

    listEl.querySelectorAll('.picker-emp-checkbox').forEach(function (cb) {
      cb.addEventListener('change', function () {
        var idx = _pickerSelected.indexOf(cb.value);
        if (cb.checked && idx === -1) _pickerSelected.push(cb.value);
        else if (!cb.checked && idx !== -1) _pickerSelected.splice(idx, 1);
        var nowAll = filtered.every(function (e) { return _pickerSelected.indexOf(e.email) !== -1; });
        if (selectAll) selectAll.checked = nowAll;
      });
    });
  }


  function handlePolicyIssue() {
    var nameInput = document.getElementById('policy-name');
    var fileInput = document.getElementById('policy-file');
    var statusEl  = document.getElementById('issue-policy-status');
    var submitBtn = document.getElementById('btn-issue-policy-submit');
    if (!nameInput || !fileInput || !statusEl || !submitBtn) return;

    var policyName     = nameInput.value.trim();
    var file           = fileInput.files[0];
    var selectedEmails = _pickerSelected.slice(); // snapshot

    if (!policyName)              { showIssueStatus('error', 'Please enter a policy name.'); return; }
    if (!file)                    { showIssueStatus('error', 'Please select a policy document (PDF).'); return; }
    if (!selectedEmails.length)   { showIssueStatus('error', 'Please select at least one employee.'); return; }
    if (file.size > 4 * 1024 * 1024) { showIssueStatus('error', 'File exceeds 4 MB. Please reduce file size and try again.'); return; }

    if (IS_DEMO) {
      showIssueStatus('success', 'Demo: “' + policyName + '” would be issued to ' + selectedEmails.length + ' employee(s) via DocuSign.');
      return;
    }

    if (!isConfigured()
        || !PORTAL_CONFIG.policyIssuancesListId
        || PORTAL_CONFIG.policyIssuancesListId === 'YOUR-POLICY-ISSUANCES-LIST-ID') {
      showIssueStatus('error', 'Policy system is not configured yet. See README for setup steps.');
      return;
    }

    submitBtn.disabled = true;
    showIssueStatus('uploading', 'Uploading policy and creating records…');

    issuePolicyDocument(file, policyName, selectedEmails)
      .then(function () {
        showIssueStatus('success', '✓ Policy issued to ' + selectedEmails.length + ' employee(s). DocuSign signing requests will be sent via email (Power Automate).');
        nameInput.value = '';
        fileInput.value = '';
        var labelText = document.getElementById('policy-file-label-text');
        if (labelText) labelText.textContent = 'Choose a PDF file…';
        _pickerSelected = [];
        buildEmployeePicker('');
        var searchInput = document.getElementById('policy-emp-search');
        if (searchInput) searchInput.value = '';
      })
      .catch(function (err) {
        console.error('Policy issue error:', err);
        showIssueStatus('error', 'Failed to issue policy. Please try again or contact IT.');
      })
      .then(function () { submitBtn.disabled = false; });

    function showIssueStatus(type, msg) {
      statusEl.hidden    = false;
      statusEl.className = 'portal-ts-status portal-ts-status--' + type;
      statusEl.textContent = msg;
      if (type === 'success') setTimeout(function () { statusEl.hidden = true; }, 8000);
    }
  }


  function issuePolicyDocument(file, policyName, emails) {
    var fd = new FormData();
    fd.append('policy_name', policyName);
    fd.append('issued_to',   JSON.stringify(emails));
    fd.append('file',        file);
    return getApiToken().then(function (token) {
      return fetch(API_BASE + '/policies/issue', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token },
        body: fd
      });
    }).then(function (r) {
      if (!r.ok) throw new Error('Policy issue HTTP ' + r.status);
    });
  }


  function loadIssuedPolicies() {
    if (IS_DEMO) {
      return Promise.resolve({ policies: DEMO_ISSUED_POLICIES, acks: DEMO_POLICY_ACKS });
    }

    return apiCall('/policies/issued').then(function (rows) {
      var policies = rows.map(function (p) {
        return { id: String(p.id), name: p.policy_name, issuedDate: p.issued_date, issuedTo: p.issued_to };
      });
      var acks = [];
      rows.forEach(function (p) {
        (p.acknowledgments || []).forEach(function (a) {
          acks.push({ policyId: String(p.id), email: a.employee_email, name: a.employee_name, status: a.status, signedDate: a.signed_date });
        });
      });
      return { policies: policies, acks: acks };
    });
  }


  function renderIssuedPoliciesTable(data, bodyEl) {
    var policies = (data && data.policies) || [];
    var acks     = (data && data.acks)     || [];

    if (!policies.length) {
      var notSetup = !isConfigured() || !PORTAL_CONFIG.policyIssuancesListId || PORTAL_CONFIG.policyIssuancesListId === 'YOUR-POLICY-ISSUANCES-LIST-ID';
      bodyEl.innerHTML = notSetup
        ? '<p class="portal-ts-empty">Policy system is not configured yet. See README for setup steps.</p>'
        : '<p class="portal-ts-empty">No policies have been issued yet. Click <strong>Issue a Policy</strong> to get started.</p>';
      return;
    }

    // Group acks by policyId
    var acksByPolicy = {};
    acks.forEach(function (ack) {
      var pid = String(ack.policyId);
      if (!acksByPolicy[pid]) acksByPolicy[pid] = [];
      acksByPolicy[pid].push(ack);
    });

    var html = '<div class="portal-ts-wrap"><table class="portal-ts-table portal-pol-table">'
      + '<thead><tr><th>Policy</th><th>Issued</th><th>Status</th><th>Track</th></tr></thead><tbody>';

    policies.forEach(function (pol) {
      var pAcks   = acksByPolicy[String(pol.id)] || [];
      var total   = pol.issuedTo.length;
      var signed  = pAcks.filter(function (a) { return a.status === 'Signed'; }).length;
      var allDone = total > 0 && signed >= total;

      var statusBadge = allDone
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Signatures Completed</span>'
        : '<span class="portal-signoff-badge portal-signoff-badge--pending">Signatures Pending</span>';

      var safeId = escapeHtml(String(pol.id));
      html += '<tr>'
        + '<td class="portal-ts-col-name">' + escapeHtml(pol.name) + '</td>'
        + '<td class="portal-ts-col-date">' + formatDate(pol.issuedDate) + '</td>'
        + '<td>' + statusBadge + '</td>'
        + '<td class="portal-pol-track-cell">'
        + '<button class="portal-pol-track-btn" data-pol-id="' + safeId + '" type="button">Track Acknowledgments</button>'
        + '</td></tr>'
        + '<tr class="portal-pol-ack-row" id="pol-ack-' + safeId + '" hidden>'
        + '<td colspan="4" class="portal-pol-ack-cell">'
        + renderAckSubTable(pol, pAcks)
        + '</td></tr>';
    });

    html += '</tbody></table></div>';
    bodyEl.innerHTML = html;

    bodyEl.querySelectorAll('.portal-pol-track-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var pid    = btn.dataset.polId;
        var ackRow = document.getElementById('pol-ack-' + pid);
        if (!ackRow) return;
        ackRow.hidden   = !ackRow.hidden;
        btn.textContent = ackRow.hidden ? 'Track Acknowledgments' : 'Hide Acknowledgments';
      });
    });
  }


  function renderAckSubTable(policy, pAcks) {
    var emailMap = {};
    pAcks.forEach(function (a) { emailMap[a.email] = a; });

    var rows = policy.issuedTo.map(function (email) {
      var ack   = emailMap[email];
      var entry = PORTAL_CONFIG.staff[email];
      return {
        name:       (ack && ack.name) || (entry && entry.name) || email,
        email:      email,
        status:     ack ? ack.status : 'Pending',
        signedDate: ack ? ack.signedDate : null
      };
    }).sort(function (a, b) {
      if (a.status === b.status) return a.name.localeCompare(b.name);
      return a.status === 'Signed' ? -1 : 1;
    });

    if (!rows.length) return '<p class="portal-pol-ack-empty">No employees assigned to this policy.</p>';

    var html = '<table class="portal-pol-ack-table"><tbody>';
    rows.forEach(function (row) {
      var icon = row.status === 'Signed'
        ? '<span class="portal-pol-ack-icon portal-pol-ack-icon--signed">&#10003;</span>'
        : '<span class="portal-pol-ack-icon portal-pol-ack-icon--pending">&#10007;</span>';
      html += '<tr>'
        + '<td class="portal-pol-ack-icon-col">' + icon + '</td>'
        + '<td class="portal-pol-ack-name">' + escapeHtml(row.name) + '<br><span class="portal-signoff-email">' + escapeHtml(row.email) + '</span></td>'
        + '<td class="portal-pol-ack-date">' + (row.signedDate ? escapeHtml(formatDate(row.signedDate)) : '') + '</td>'
        + '</tr>';
    });
    html += '</tbody></table>';
    return html;
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Timesheet upload
  // ══════════════════════════════════════════════════════════════════════════
  function setupTimesheetUpload(userEmail, userDisplayName) {
    var pairs = [
      { btnId: 'btn-submit-ca',  inputId: 'ts-file-ca'  },
      { btnId: 'btn-submit-std', inputId: 'ts-file-std' }
    ];
    var statusEl = document.getElementById('ts-upload-status');

    pairs.forEach(function (p) {
      var btn   = document.getElementById(p.btnId);
      var input = document.getElementById(p.inputId);
      if (!btn || !input) return;

      btn.addEventListener('click', function () {
        input.value = '';
        input.click();
      });

      input.addEventListener('change', function () {
        var file = this.files[0];
        if (!file) return;

        if (file.size > 10 * 1024 * 1024) {
          showStatus('error', 'File exceeds 10 MB. Please contact IT if you need help compressing it.');
          return;
        }

        if (IS_DEMO) {
          showStatus('success', 'Demo mode: "' + file.name + '" would be uploaded.');
          return;
        }

        setBtnState(btn, 'uploading');
        showStatus('uploading', 'Uploading ' + file.name + '…');

        uploadTimesheetFile(file, userEmail, userDisplayName)
          .then(function () {
            setBtnState(btn, 'success');
            showStatus('success', '✓ ' + file.name + ' submitted successfully.');
            setTimeout(function () { setBtnState(btn, 'idle'); }, 3000);
          })
          .catch(function (err) {
            console.error('Upload error:', err);
            setBtnState(btn, 'error');
            showStatus('error', 'Upload failed. Please try again or contact IT.');
            setTimeout(function () { setBtnState(btn, 'idle'); }, 4000);
          });
      });
    });

    function setBtnState(btn, state) {
      btn.classList.remove('is-uploading', 'is-success', 'is-error');
      btn.disabled = (state === 'uploading');
      if (state !== 'idle') btn.classList.add('is-' + state);
    }

    function showStatus(type, msg) {
      statusEl.hidden    = false;
      statusEl.className = 'portal-ts-status portal-ts-status--' + type;
      statusEl.textContent = msg;
      if (type === 'success') setTimeout(function () { statusEl.hidden = true; }, 5000);
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Timesheet viewer (manager + HR)
  // emails: array of M365 emails to load, or null for demo/all
  // ══════════════════════════════════════════════════════════════════════════
  function setupTimesheetViewer(role, emails) {
    var btn     = document.getElementById('btn-view-' + role + '-ts');
    var panel   = document.getElementById('panel-' + role + '-ts');
    var bodyEl  = document.getElementById('body-' + role + '-ts');
    var refresh = document.getElementById('btn-refresh-' + role + '-ts');
    if (!btn || !panel || !bodyEl) return;

    var loaded = false;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Hide Timesheets';
        if (!loaded) { loadAndRender(); loaded = true; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View Timesheets';
      }
    });

    if (refresh) {
      refresh.addEventListener('click', function () { loaded = false; loadAndRender(); loaded = true; });
    }

    function loadAndRender() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading timesheets…</p>';
      loadTimesheetsForEmails(emails)
        .then(function (rows) { renderTimesheetTable(rows, bodyEl, role); })
        .catch(function (err) {
          console.error('Timesheet load error:', err);
          bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load timesheets. Check your connection and click Refresh to try again.</p>';
        });
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Timesheet table renderer
  // ══════════════════════════════════════════════════════════════════════════
  function renderTimesheetTable(rows, bodyEl, role) {
    var isManager = (role === 'manager');
    if (!rows || rows.length === 0) {
      bodyEl.innerHTML = '<p class="portal-ts-empty">No timesheets have been submitted yet.</p>';
      return;
    }

    var html = '<div class="portal-ts-wrap"><table class="portal-ts-table">'
      + '<thead><tr><th>Employee</th><th>Submitted</th><th>File</th><th>Status</th>'
      + (isManager ? '<th></th>' : '')
      + '</tr></thead><tbody>';

    rows.forEach(function (row) {
      var fileCell = row.downloadId
        ? '<a href="#" class="portal-ts-link" data-ts-id="' + row.downloadId + '" data-ts-name="' + escapeHtml(row.filename) + '">' + escapeHtml(row.filename) + '</a>'
        : '<span class="portal-ts-nolink">' + escapeHtml(row.filename) + ' <em>(demo)</em></span>';

      var st = row.status || 'pending';
      var statusBadge = st === 'approved'
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
        : st === 'rejected'
          ? '<span class="portal-pto-badge--denied">Rejected</span>'
          : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';

      var actionCell = '';
      if (isManager) {
        actionCell = st === 'pending'
          ? '<div class="portal-pto-actions" id="ts-actions-' + row.downloadId + '">'
            + '<button class="portal-pto-approve" data-tsaction-id="' + row.downloadId + '" data-action="approved" type="button" title="Approve">&#10003;</button>'
            + '<button class="portal-pto-deny"    data-tsaction-id="' + row.downloadId + '" data-action="rejected" type="button" title="Reject">&#10007;</button>'
            + '</div>'
          : '<span class="portal-pto-actioned">—</span>';
      }

      html += '<tr>'
        + '<td class="portal-ts-col-name">' + escapeHtml(row.name) + '</td>'
        + '<td class="portal-ts-col-date">' + formatDate(row.date) + '</td>'
        + '<td class="portal-ts-col-file">' + fileCell + '</td>'
        + '<td id="ts-status-' + row.downloadId + '">' + statusBadge + '</td>'
        + (isManager ? '<td>' + actionCell + '</td>' : '')
        + '</tr>';
    });

    html += '</tbody></table></div>';
    bodyEl.innerHTML = html;

    bodyEl.querySelectorAll('[data-ts-id][data-ts-name]').forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        if (!IS_DEMO) downloadFile('/timesheets/download/' + a.dataset.tsId, a.dataset.tsName);
      });
    });

    if (isManager) {
      bodyEl.querySelectorAll('[data-tsaction-id]').forEach(function (b) {
        b.addEventListener('click', function () {
          handleTimesheetDecision(b.dataset.tsactionId, b.dataset.action);
        });
      });
    }
  }

  function handleTimesheetDecision(id, action) {
    var actionsEl  = document.getElementById('ts-actions-' + id);
    var statusCell = document.getElementById('ts-status-' + id);
    if (actionsEl) actionsEl.querySelectorAll('button').forEach(function (b) { b.disabled = true; });
    if (IS_DEMO) {
      if (actionsEl)  actionsEl.outerHTML = '<span class="portal-pto-actioned">—</span>';
      if (statusCell) statusCell.innerHTML = action === 'approved'
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
        : '<span class="portal-pto-badge--denied">Rejected</span>';
      return;
    }
    apiCall('/timesheets/' + id + '?status=' + action, { method: 'PATCH' })
      .then(function () {
        if (actionsEl)  actionsEl.outerHTML = '<span class="portal-pto-actioned">—</span>';
        if (statusCell) statusCell.innerHTML = action === 'approved'
          ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
          : '<span class="portal-pto-badge--denied">Rejected</span>';
      })
      .catch(function (err) {
        console.error('Timesheet decision error:', err);
        if (actionsEl) actionsEl.querySelectorAll('button').forEach(function (b) { b.disabled = false; });
      });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Backend: load timesheets
  // ══════════════════════════════════════════════════════════════════════════
  function loadTimesheetsForEmails(emails) {
    if (IS_DEMO) {
      return Promise.resolve(
        DEMO_TIMESHEETS
          .filter(function (ts) { return !emails || emails.indexOf(ts.email) !== -1; })
          .sort(function (a, b) { return new Date(b.date) - new Date(a.date); })
      );
    }

    var path = emails === null
      ? '/timesheets/all'
      : '/timesheets/team?emails=' + encodeURIComponent(emails.join(','));

    return apiCall(path).then(function (rows) {
      return rows.map(function (r) {
        return {
          name:       r.employee_name,
          email:      r.employee_email,
          filename:   r.filename,
          date:       r.uploaded_at,
          downloadId: r.id,
          status:     r.status || 'pending'
        };
      });
    });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Backend: upload a timesheet file
  // ══════════════════════════════════════════════════════════════════════════
  function uploadTimesheetFile(file, userEmail, userDisplayName) {
    var date     = new Date().toISOString().slice(0, 10);
    var newName  = date + ' - ' + userDisplayName + ' - ' + file.name;
    var renamed  = new File([file], newName, { type: file.type });
    var fd       = new FormData();
    fd.append('file', renamed);
    return getApiToken().then(function (token) {
      return fetch(API_BASE + '/timesheets/', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token },
        body: fd
      });
    }).then(function (r) {
      if (!r.ok) throw new Error('Upload HTTP ' + r.status);
    });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Backend API helpers
  // ══════════════════════════════════════════════════════════════════════════
  function getApiToken() {
    var req = { scopes: ['openid', 'profile'], account: currentAccount };
    return msalInstance.acquireTokenSilent(req)
      .catch(function () { return msalInstance.acquireTokenPopup(req); })
      .then(function (result) { return result.idToken; });
  }

  function apiCall(path, options) {
    return getApiToken().then(function (token) {
      var opts    = options || {};
      var headers = Object.assign({}, opts.headers || {}, { 'Authorization': 'Bearer ' + token });
      if (!(opts.body instanceof FormData)) {
        headers['Content-Type'] = headers['Content-Type'] || 'application/json';
      }
      return fetch(API_BASE + path, Object.assign({}, opts, { headers: headers }));
    }).then(function (r) {
      if (!r.ok) return r.json().then(function (e) { throw new Error(e.detail || 'API error ' + r.status); }).catch(function () { throw new Error('API error ' + r.status); });
      return r.json();
    });
  }

  function getGraphToken() {
    var req = { scopes: ['https://graph.microsoft.com/User.ReadBasic.All'], account: currentAccount };
    return msalInstance.acquireTokenSilent(req)
      .catch(function () { return msalInstance.acquireTokenPopup(req); })
      .then(function (result) { return result.accessToken; });
  }

  function apiDownload(path) {
    return getApiToken().then(function (token) {
      return fetch(API_BASE + path, { headers: { 'Authorization': 'Bearer ' + token } });
    });
  }

  function downloadFile(path, filename) {
    apiDownload(path).then(function (r) {
      if (!r.ok) throw new Error('Download failed');
      return r.blob();
    }).then(function (blob) {
      var url = URL.createObjectURL(blob);
      var a   = document.createElement('a');
      a.href  = url; a.download = filename;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a); URL.revokeObjectURL(url);
    }).catch(function (err) {
      console.error('Download error:', err);
      alert('Download failed. Please try again.');
    });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Utilities
  // ══════════════════════════════════════════════════════════════════════════

  function formatDate(iso) {
    var d = new Date(iso);
    return d.toLocaleString('en-US', {
      month: 'short', day: 'numeric', year: 'numeric',
      hour: 'numeric', minute: '2-digit', hour12: true
    });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Format "Aug 4–8, 2026" or "Jul 31–Aug 1, 2026"
  function formatDateRange(startStr, endStr) {
    if (!startStr) return '';
    var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var sp = startStr.split('-');
    var start = new Date(+sp[0], +sp[1] - 1, +sp[2]);
    var end   = start;
    if (endStr && endStr !== startStr) {
      var ep = endStr.split('-'); end = new Date(+ep[0], +ep[1] - 1, +ep[2]);
    }
    var endYear = end.getFullYear();
    var sLabel  = MONTHS[start.getMonth()] + ' ' + start.getDate();
    if (startStr === endStr || !endStr) return sLabel + ', ' + endYear;
    if (start.getMonth() === end.getMonth() && start.getFullYear() === end.getFullYear()) {
      return sLabel + '–' + end.getDate() + ', ' + endYear;
    }
    return sLabel + '–' + MONTHS[end.getMonth()] + ' ' + end.getDate() + ', ' + endYear;
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Company Directory — searchable employee list with Teams chat links
  // ══════════════════════════════════════════════════════════════════════════
  function setupDirectory() {
    var btn     = document.getElementById('btn-open-directory');
    var btn2    = document.getElementById('btn-open-team-directory');
    var panel   = document.getElementById('panel-directory');
    var search  = document.getElementById('dir-search');
    var grid    = document.getElementById('dir-grid');
    var count   = document.getElementById('dir-count');
    if (!panel || !grid) return;

    var roster = [];
    var rosterLoaded = false;
    var COLORS = ['#0f1e42','#1e3a5f','#1e40af','#065f46','#9f1239','#854d0e','#5b21b6','#48566f'];

    function togglePanel() {
      var opening = panel.hidden;
      panel.hidden = !opening;
      var label = opening ? 'Close' : 'Browse';
      if (btn)  btn.textContent  = label;
      if (btn2) btn2.textContent = label;
      if (opening) {
        if (search) { search.value = ''; search.focus(); }
        if (!rosterLoaded) loadDirectory();
      }
    }

    if (btn)  btn.addEventListener('click',  togglePanel);
    if (btn2) btn2.addEventListener('click', togglePanel);

    if (search) {
      search.addEventListener('input', function () {
        var q = search.value.toLowerCase().trim();
        renderGrid(q ? roster.filter(function (e) {
          return e.name.toLowerCase().indexOf(q) !== -1 || (e.email || '').toLowerCase().indexOf(q) !== -1;
        }) : roster);
      });
    }

    function loadDirectory() {
      roster = (PORTAL_CONFIG.directory || []).slice();
      rosterLoaded = true;
      renderGrid(roster);
    }

    function renderGrid(entries) {
      if (count) count.textContent = entries.length + ' of ' + roster.length + ' people';
      if (!entries.length) {
        grid.innerHTML = '<p class="portal-dir-empty">No results match your search.</p>';
        return;
      }
      grid.innerHTML = entries.map(function (emp) {
        var initial  = emp.name.charAt(0).toUpperCase();
        var color    = COLORS[emp.name.charCodeAt(0) % COLORS.length];
        var teamsUrl = 'https://teams.microsoft.com/l/chat/0/0?users=' + encodeURIComponent(emp.email);
        return '<div class="portal-dir-card">'
          + '<div class="portal-dir-avatar" style="background:' + color + '">' + escapeHtml(initial) + '</div>'
          + '<div class="portal-dir-info">'
          + '<span class="portal-dir-name">' + escapeHtml(emp.name) + '</span>'
          + '<span class="portal-dir-email">' + escapeHtml(emp.email) + '</span>'
          + (emp.title ? '<span class="portal-dir-title">' + escapeHtml(emp.title) + (emp.dept ? ' — ' + escapeHtml(emp.dept) : '') + '</span>' : '')
          + '</div>'
          + '<a href="' + teamsUrl + '" target="_blank" rel="noopener" class="portal-dir-teams" title="Message ' + escapeHtml(emp.name) + ' on Teams">'
          + '<svg width="13" height="13" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M20 2H4c-1.1 0-2 .9-2 2v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2z"/></svg>'
          + ' Chat</a>'
          + '</div>';
      }).join('');
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Incident Report — employee-facing form
  // ══════════════════════════════════════════════════════════════════════════
  function setupIncidentReport(userEmail, userDisplayName) {
    var openBtn   = document.getElementById('btn-open-incident-form');
    var panel     = document.getElementById('panel-incident-report');
    var form      = document.getElementById('form-incident-report');
    var dateEl    = document.getElementById('inc-date');
    var typeEl    = document.getElementById('inc-type');
    var locEl     = document.getElementById('inc-location');
    var descEl    = document.getElementById('inc-desc');
    var statusEl  = document.getElementById('incident-report-status');
    var submitBtn = document.getElementById('btn-incident-submit');
    if (!openBtn || !panel || !form || !dateEl || !descEl) return;

    dateEl.max = new Date().toISOString().slice(0, 10); // can't report a future incident

    openBtn.addEventListener('click', function () {
      var opening = panel.hidden;
      panel.hidden = !opening;
      openBtn.textContent = opening ? 'Close Form' : 'Report Incident';
      if (opening) {
        form.reset();
        dateEl.max = new Date().toISOString().slice(0, 10);
        if (statusEl) statusEl.hidden = true;
      }
    });

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var incidentDate = dateEl.value;
      var incidentType = typeEl ? typeEl.value : '';
      var location     = locEl  ? locEl.value.trim() : '';
      var description  = descEl.value.trim();

      if (!incidentDate)  { showStatus('error', 'Please select the date of the incident.'); return; }
      if (!incidentType)  { showStatus('error', 'Please select an incident type.'); return; }
      if (!description)   { showStatus('error', 'Please describe what happened.'); return; }

      if (IS_DEMO) {
        showStatus('success', '✓ Demo: Incident report submitted. HR has been notified.');
        return;
      }

      submitBtn.disabled = true;
      showStatus('uploading', 'Submitting report…');

      apiCall('/incidents/', {
        method: 'POST',
        body: JSON.stringify({
          incident_date: incidentDate,
          location:      location,
          incident_type: incidentType,
          description:   description
        })
      })
      .then(function () {
        showStatus('success', '✓ Report submitted. HR will follow up with you.');
        form.reset();
        dateEl.max = new Date().toISOString().slice(0, 10);
      })
      .catch(function (err) {
        console.error('Incident submit error:', err);
        showStatus('error', 'Failed to submit. Please try again or contact IT directly.');
      })
      .then(function () { submitBtn.disabled = false; });
    });

    function showStatus(type, msg) {
      statusEl.hidden    = false;
      statusEl.className = 'portal-ts-status portal-ts-status--' + type;
      statusEl.textContent = msg;
      if (type === 'success') setTimeout(function () { statusEl.hidden = true; }, 9000);
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // HR — view all incident reports
  // ══════════════════════════════════════════════════════════════════════════
  function setupHRIncidentReports() {
    var btn     = document.getElementById('btn-view-incident-reports');
    var panel   = document.getElementById('panel-incident-reports-hr');
    var bodyEl  = document.getElementById('body-incident-reports-hr');
    var refresh = document.getElementById('btn-refresh-incident-reports');
    if (!btn || !panel || !bodyEl) return;

    var loaded = false;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Hide Reports';
        if (!loaded) { loadAndRender(); loaded = true; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View Reports';
      }
    });

    if (refresh) {
      refresh.addEventListener('click', function () { loaded = false; loadAndRender(); loaded = true; });
    }

    function loadAndRender() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading incident reports…</p>';

      if (IS_DEMO) {
        renderTable(DEMO_INCIDENT_REPORTS.slice().sort(function (a, b) {
          return new Date(b.submittedDate) - new Date(a.submittedDate);
        }));
        return;
      }

      apiCall('/incidents/')
        .then(function (data) {
          var rows = data.map(function (item) {
            return {
              id:            String(item.id),
              email:         item.employee_email || '',
              name:          item.employee_name  || item.employee_email || '',
              incidentDate:  item.incident_date  || '',
              location:      item.location       || '',
              incidentType:  item.incident_type  || '',
              description:   item.description    || '',
              submittedDate: item.created_at
            };
          });
          renderTable(rows);
        })
      .catch(function (err) {
        console.error('Incident reports load error:', err);
        bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load reports. Click Refresh to try again.</p>';
      });
    }

    function renderTable(rows) {
      if (!rows || rows.length === 0) {
        var notSetup = !IS_DEMO && (!isConfigured() || !PORTAL_CONFIG.incidentReportsListId || PORTAL_CONFIG.incidentReportsListId === 'YOUR-INCIDENT-REPORTS-LIST-ID');
        bodyEl.innerHTML = notSetup
          ? '<p class="portal-ts-empty">Incident reporting is not yet configured. See README for setup steps.</p>'
          : '<p class="portal-ts-empty">No incident reports have been submitted yet.</p>';
        return;
      }

      var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      function fmtDate(iso) {
        if (!iso) return '';
        var d = new Date(iso);
        return MONTHS[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
      }

      var typeBadge = {
        'Injury':          'portal-inc-badge portal-inc-badge--injury',
        'Near Miss':       'portal-inc-badge portal-inc-badge--nearmiss',
        'Property Damage': 'portal-inc-badge portal-inc-badge--damage',
        'Other':           'portal-inc-badge portal-inc-badge--other'
      };

      var html = '<div class="portal-ts-wrap"><table class="portal-ts-table portal-inc-table">'
        + '<thead><tr><th>Employee</th><th>Submitted</th><th>Incident Date</th><th>Type</th><th>Location</th><th>Description</th></tr></thead><tbody>';

      rows.forEach(function (row) {
        var cls = typeBadge[row.incidentType] || 'portal-inc-badge';
        html += '<tr>'
          + '<td class="portal-ts-col-name">' + escapeHtml(row.name) + '<br><span class="portal-signoff-email">' + escapeHtml(row.email) + '</span></td>'
          + '<td class="portal-pto-dates-col">' + escapeHtml(fmtDate(row.submittedDate)) + '</td>'
          + '<td class="portal-pto-dates-col">' + escapeHtml(fmtDate(row.incidentDate)) + '</td>'
          + '<td><span class="' + cls + '">' + escapeHtml(row.incidentType) + '</span></td>'
          + '<td class="portal-inc-location">' + escapeHtml(row.location || '—') + '</td>'
          + '<td class="portal-inc-description">' + escapeHtml(row.description) + '</td>'
          + '</tr>';
      });

      html += '</tbody></table></div>';
      bodyEl.innerHTML = html;
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // PTO Request — employee-facing form
  // ══════════════════════════════════════════════════════════════════════════
  function setupPTORequest(userEmail, userDisplayName) {
    var openBtn   = document.getElementById('btn-open-pto-form');
    var panel     = document.getElementById('panel-pto-request');
    var form      = document.getElementById('form-pto-request');
    var startEl   = document.getElementById('pto-start');
    var endEl     = document.getElementById('pto-end');
    var reasonEl  = document.getElementById('pto-reason');
    var statusEl  = document.getElementById('pto-request-status');
    var submitBtn = document.getElementById('btn-pto-submit');
    if (!openBtn || !panel || !form || !startEl || !endEl || !reasonEl) return;

    openBtn.addEventListener('click', function () {
      var opening = panel.hidden;
      panel.hidden = !opening;
      openBtn.textContent = opening ? 'Close Form' : 'Submit Request';
      if (opening) {
        form.reset();
        var todayStr = new Date().toISOString().slice(0, 10);
        startEl.min = todayStr;
        endEl.min   = '';
        endEl.max   = '';
        statusEl.hidden = true;
      }
    });

    startEl.addEventListener('change', function () {
      if (!startEl.value) { endEl.min = ''; endEl.max = ''; return; }
      var s   = startEl.value.split('-');
      var max = new Date(+s[0], +s[1] - 1, +s[2] + 13).toISOString().slice(0, 10);
      endEl.min = startEl.value;
      endEl.max = max;
      if (endEl.value && (endEl.value < startEl.value || endEl.value > max)) endEl.value = '';
    });

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var start  = startEl.value;
      var end    = endEl.value;
      var reason = reasonEl.value.trim();
      if (!start)  { showPTOStatus('error', 'Please select a start date.'); return; }
      if (!end)    { showPTOStatus('error', 'Please select an end date.');   return; }
      if (!reason) { showPTOStatus('error', 'Please enter a reason for your request.'); return; }

      if (IS_DEMO) {
        showPTOStatus('success', '✓ Demo: PTO request for ' + formatDateRange(start, end) + ' submitted. Your manager will be notified.');
        return;
      }

      submitBtn.disabled = true;
      showPTOStatus('uploading', 'Submitting…');

      apiCall('/pto/', {
        method: 'POST',
        body: JSON.stringify({ start_date: start, end_date: end, reason: reason })
      })
      .then(function () {
        showPTOStatus('success', '✓ PTO request for ' + formatDateRange(start, end) + ' submitted. Your manager will review it shortly.');
        form.reset();
        endEl.min = ''; endEl.max = '';
      })
      .catch(function (err) {
        console.error('PTO submit error:', err);
        showPTOStatus('error', 'Failed to submit. Please try again or contact IT.');
      })
      .then(function () { submitBtn.disabled = false; });
    });

    function showPTOStatus(type, msg) {
      statusEl.hidden    = false;
      statusEl.className = 'portal-ts-status portal-ts-status--' + type;
      statusEl.textContent = msg;
      if (type === 'success') setTimeout(function () { statusEl.hidden = true; }, 8000);
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // PTO Approvals — manager and HR table (role = 'manager' | 'hr')
  // ══════════════════════════════════════════════════════════════════════════
  function setupPTOApprovals(teamEmails, role) {
    var isHRRole = (role === 'hr');
    var pfx     = isHRRole ? 'hr-pto' : 'pto-approvals';
    var btn     = document.getElementById('btn-view-' + pfx);
    var panel   = document.getElementById('panel-' + pfx);
    var bodyEl  = document.getElementById('body-' + pfx);
    var refresh = document.getElementById('btn-refresh-' + pfx);
    if (!btn || !panel || !bodyEl) return;

    var loaded = false;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Hide Requests';
        if (!loaded) { loadAndRender(); loaded = true; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View Requests';
      }
    });

    if (refresh) {
      refresh.addEventListener('click', function () { loaded = false; loadAndRender(); loaded = true; });
    }

    function mapRow(item) {
      return {
        id:            String(item.id),
        email:         item.employee_email  || '',
        name:          item.employee_name   || item.employee_email || '',
        startDate:     item.start_date      || '',
        endDate:       item.end_date        || '',
        reason:        item.reason          || '',
        managerStatus: item.manager_status  || 'pending',
        hrStatus:      item.hr_status       || 'pending',
        submittedDate: item.created_at
      };
    }

    function loadAndRender() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading PTO requests…</p>';
      if (IS_DEMO) {
        var rows = DEMO_PTO_REQUESTS
          .filter(function (r) { return !teamEmails || teamEmails.indexOf(r.email) !== -1; })
          .map(function (r) { return { id: r.id, email: r.email, name: r.name, startDate: r.startDate, endDate: r.endDate, reason: r.reason, managerStatus: 'pending', hrStatus: 'pending', submittedDate: r.submittedDate }; })
          .sort(function (a, b) { return new Date(b.submittedDate) - new Date(a.submittedDate); });
        renderPTOTable(rows);
        return;
      }
      var path = isHRRole ? '/pto/all'
        : (teamEmails && teamEmails.length ? '/pto/team?emails=' + encodeURIComponent(teamEmails.join(',')) : '/pto/team?emails=');
      apiCall(path)
        .then(function (data) { renderPTOTable(data.map(mapRow)); })
        .catch(function (err) {
          console.error('PTO load error:', err);
          bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load requests. Click Refresh to try again.</p>';
        });
    }

    function ptoBadge(status) {
      return status === 'approved'
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
        : status === 'denied'
          ? '<span class="portal-pto-badge--denied">Denied</span>'
          : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';
    }

    function renderPTOTable(rows) {
      if (!rows || rows.length === 0) {
        bodyEl.innerHTML = '<p class="portal-ts-empty">No PTO requests yet.</p>';
        return;
      }

      // Column headers: who this role acts on + the other role's status for visibility
      var myCol    = isHRRole ? 'HR Decision'      : 'Manager Decision';
      var otherCol = isHRRole ? 'Manager Status'   : 'HR Status';

      var html = '<div class="portal-ts-wrap"><table class="portal-ts-table portal-pto-table">'
        + '<thead><tr><th>Employee</th><th>Dates</th><th>Reason</th><th>' + otherCol + '</th><th>' + myCol + '</th><th></th></tr></thead><tbody>';

      rows.forEach(function (row) {
        var myStatus    = isHRRole ? row.hrStatus      : row.managerStatus;
        var otherStatus = isHRRole ? row.managerStatus : row.hrStatus;
        var myId        = 'pto-my-status-' + escapeHtml(row.id);
        var actId       = 'pto-actions-' + escapeHtml(row.id);

        var actions = myStatus === 'pending'
          ? '<div class="portal-pto-actions" id="' + actId + '">'
            + '<button class="portal-pto-approve" data-pto-id="' + escapeHtml(row.id) + '" data-action="approved" type="button" title="Approve">&#10003;</button>'
            + '<button class="portal-pto-deny"    data-pto-id="' + escapeHtml(row.id) + '" data-action="denied"   type="button" title="Deny">&#10007;</button>'
            + '</div>'
          : '<span class="portal-pto-actioned">—</span>';

        html += '<tr>'
          + '<td class="portal-ts-col-name">' + escapeHtml(row.name) + '<br><span class="portal-signoff-email">' + escapeHtml(row.email) + '</span></td>'
          + '<td class="portal-pto-dates-col">' + escapeHtml(formatDateRange(row.startDate, row.endDate)) + '</td>'
          + '<td class="portal-pto-reason-col">' + escapeHtml(row.reason) + '</td>'
          + '<td>' + ptoBadge(otherStatus) + '</td>'
          + '<td id="' + myId + '">' + ptoBadge(myStatus) + '</td>'
          + '<td>' + actions + '</td>'
          + '</tr>';
      });

      html += '</tbody></table></div>';
      bodyEl.innerHTML = html;

      bodyEl.querySelectorAll('[data-pto-id]').forEach(function (b) {
        b.addEventListener('click', function () {
          handlePTODecision(b.dataset.ptoId, b.dataset.action);
        });
      });
    }

    function handlePTODecision(id, action) {
      var actionsEl  = document.getElementById('pto-actions-' + id);
      var statusCell = document.getElementById('pto-my-status-' + id);
      if (actionsEl) actionsEl.querySelectorAll('button').forEach(function (b) { b.disabled = true; });

      function applyUpdate() {
        if (actionsEl)  actionsEl.outerHTML = '<span class="portal-pto-actioned">—</span>';
        if (statusCell) statusCell.innerHTML = action === 'approved'
          ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
          : '<span class="portal-pto-badge--denied">Denied</span>';
      }

      if (IS_DEMO) { applyUpdate(); return; }

      var endpoint = isHRRole ? '/pto/' + id + '/hr' : '/pto/' + id + '/manager';
      apiCall(endpoint + '?status=' + action, { method: 'PATCH' })
        .then(applyUpdate)
        .catch(function (err) {
          console.error('PTO decision error:', err);
          if (actionsEl) actionsEl.querySelectorAll('button').forEach(function (b) { b.disabled = false; });
        });
    }
  }


  // ══════════════════════════════════════════════════════════════════════════
  // My PTO Requests — employee view of their own submissions
  // ══════════════════════════════════════════════════════════════════════════
  function setupMyPTORequests() {
    var btn    = document.getElementById('btn-view-my-pto');
    var panel  = document.getElementById('panel-my-pto');
    var bodyEl = document.getElementById('body-my-pto');
    if (!btn || !panel || !bodyEl) return;

    var loaded = false;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Hide My Requests';
        if (!loaded) { loadAndRender(); loaded = true; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View My Requests';
      }
    });

    function loadAndRender() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading your PTO requests…</p>';
      if (IS_DEMO) {
        renderMyPTO([]);
        return;
      }
      apiCall('/pto/my')
        .then(function (data) {
          var rows = data.map(function (item) {
            return {
              startDate:     item.start_date     || '',
              endDate:       item.end_date       || '',
              reason:        item.reason         || '',
              managerStatus: item.manager_status || 'pending',
              hrStatus:      item.hr_status      || 'pending',
              submittedDate: item.created_at
            };
          });
          renderMyPTO(rows);
        })
        .catch(function (err) {
          console.error('My PTO load error:', err);
          bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load your requests. Try again.</p>';
        });
    }

    function renderMyPTO(rows) {
      if (!rows || rows.length === 0) {
        bodyEl.innerHTML = '<p class="portal-ts-empty">You have not submitted any PTO requests yet.</p>';
        return;
      }
      function badge(status) {
        return status === 'approved'
          ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
          : status === 'denied'
            ? '<span class="portal-pto-badge--denied">Denied</span>'
            : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';
      }
      var html = '<div class="portal-ts-wrap"><table class="portal-ts-table portal-pto-table">'
        + '<thead><tr><th>Dates</th><th>Reason</th><th>Manager</th><th>HR</th></tr></thead><tbody>';
      rows.forEach(function (row) {
        html += '<tr>'
          + '<td class="portal-pto-dates-col">' + escapeHtml(formatDateRange(row.startDate, row.endDate)) + '</td>'
          + '<td class="portal-pto-reason-col">' + escapeHtml(row.reason) + '</td>'
          + '<td>' + badge(row.managerStatus) + '</td>'
          + '<td>' + badge(row.hrStatus) + '</td>'
          + '</tr>';
      });
      html += '</tbody></table></div>';
      bodyEl.innerHTML = html;
    }
  }


}());
