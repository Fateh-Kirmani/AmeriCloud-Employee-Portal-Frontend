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

    // ── Show signed-in user ───────────────────────────────────────────────
    var sidebarUser = document.getElementById('sidebar-user');
    if (sidebarUser) {
      document.getElementById('sidebar-user-name').textContent  = userDisplayName;
      document.getElementById('sidebar-user-email').textContent = userEmail;
      sidebarUser.hidden = false;
    }
    var topbarName = document.getElementById('topbar-user-name');
    if (topbarName) {
      topbarName.textContent = userDisplayName;
      topbarName.hidden = false;
    }

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

    // ── Timesheet form ────────────────────────────────────────────────────
    var staffEntry = PORTAL_CONFIG.staff[userEmail] || null;
    var team = (staffEntry && staffEntry.team) ? staffEntry.team : [];

    // Timesheet form — available to all employees
    setupTimesheetForm(userEmail, userDisplayName, staffEntry);

    // Available to all roles — greeting + at-a-glance strip
    setupHomeGreeting(userDisplayName, isManager, isHR, team);

    if (isManager) {
      setupFormTimesheetViewer('manager', IS_DEMO ? null : team);
      setupPTOApprovals(IS_DEMO ? null : team, 'manager');
    }

    if (isHR) {
      var allEmails = Object.keys(PORTAL_CONFIG.staff);
      setupFormTimesheetViewer('hr', null);
      setupHandbookSignoffViewer();
      setupHRPolicies();
      setupHRIncidentReports();
      setupHREfs();
      setupPTOApprovals(null, 'hr');
    }

    // Available to all roles — shows policies issued to the signed-in employee
    setupEmployeePolicies(userEmail);

    // Available to all roles — shows EFS if issued to this employee
    setupEmployeeEfs(userEmail);

    // Available to all roles — PTO request form
    setupPTORequest(userEmail, userDisplayName);

    // Available to all roles — my submitted PTO status
    setupMyPTORequests();

    // Available to all roles — incident report form
    setupIncidentReport(userEmail, userDisplayName);

    // Available to all roles — company news
    setupNews();

    // Available to all roles — company directory
    setupDirectory(team);

    // Available to all roles — company calendar
    setupCalendar();

    // Available to all roles — trainings & certifications
    setupTrainings(userEmail, userDisplayName);
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Company Calendar
  // ══════════════════════════════════════════════════════════════════════════
  function setupCalendar() {
    var btn    = document.getElementById('btn-open-calendar');
    var panel  = document.getElementById('panel-calendar');
    var bodyEl = document.getElementById('body-calendar');
    if (!btn || !panel || !bodyEl) return;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Close ×';
        if (!bodyEl.dataset.rendered) { renderCalendar(bodyEl); bodyEl.dataset.rendered = '1'; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View →';
      }
    });
  }

  // ══════════════════════════════════════════════════════════════════════════
  // Trainings & Certifications
  // ══════════════════════════════════════════════════════════════════════════
  var TRAININGS_ADMIN = 'fkirmani@americloudtelecom.com';

  function setupTrainings(userEmail, userDisplayName) {
    var modalView        = document.getElementById('modal-trainings-view');
    var modalViewOverlay = document.getElementById('modal-trainings-view-overlay');
    var bodyView         = document.getElementById('body-trainings-view');
    var closeView        = document.getElementById('btn-close-trainings-view');
    var btnView          = document.getElementById('btn-view-my-trainings');

    var modalUpload        = document.getElementById('modal-trainings-upload');
    var modalUploadOverlay = document.getElementById('modal-trainings-upload-overlay');
    var closeUpload        = document.getElementById('btn-close-trainings-upload');
    var btnUpload          = document.getElementById('btn-upload-training');
    var formUpload         = document.getElementById('form-training-upload');
    var uploadStatus       = document.getElementById('training-upload-status');
    var adminRow           = document.getElementById('training-upload-admin-row');
    var adminSelect        = document.getElementById('training-upload-emp-email');

    function closeViewModal()   { if (modalView)   modalView.hidden   = true; }
    function closeUploadModal() { if (modalUpload) modalUpload.hidden = true; }

    if (closeView)        closeView.addEventListener('click', closeViewModal);
    if (modalViewOverlay) modalViewOverlay.addEventListener('click', closeViewModal);
    if (closeUpload)        closeUpload.addEventListener('click', closeUploadModal);
    if (modalUploadOverlay) modalUploadOverlay.addEventListener('click', closeUploadModal);

    // Admin: show employee selector and populate from directory
    var isAdmin = userEmail && userEmail.toLowerCase() === TRAININGS_ADMIN;
    if (isAdmin && adminRow) {
      adminRow.hidden = false;
      if (adminSelect && window.PORTAL_CONFIG && window.PORTAL_CONFIG.directory) {
        PORTAL_CONFIG.directory.forEach(function(emp) {
          var opt = document.createElement('option');
          opt.value = emp.email;
          opt.textContent = emp.name + ' (' + emp.email + ')';
          adminSelect.appendChild(opt);
        });
      }
    }

    // ── View modal ──────────────────────────────────────────────────────────
    if (btnView) {
      btnView.addEventListener('click', function() {
        if (modalView) modalView.hidden = false;
        loadTrainingDocs();
      });
    }

    function loadTrainingDocs() {
      if (!bodyView) return;
      bodyView.innerHTML = '<p class="portal-ts-empty">Loading…</p>';
      getApiToken().then(function(token) {
        return fetch(API_BASE + '/trainings', { headers: { 'Authorization': 'Bearer ' + token } });
      }).then(function(r) {
        return r.json();
      }).then(function(docs) {
        if (!docs || !docs.length) {
          bodyView.innerHTML = '<p class="portal-ts-empty">No documents on file.</p>';
          return;
        }
        var showEmp = isAdmin;
        var html = '<table class="portal-ts-table">'
          + '<thead><tr>'
          + (showEmp ? '<th>Employee</th>' : '')
          + '<th>Document</th>'
          + '<th>Issuance Date</th>'
          + '<th>Renewal Date</th>'
          + '<th></th>'
          + '</tr></thead><tbody>';
        docs.forEach(function(doc) {
          html += '<tr data-doc-id="' + doc.id + '">'
            + (showEmp ? '<td>' + esc(doc.employee_name) + '</td>' : '')
            + '<td><a href="#" class="portal-training-link" data-doc-id="' + doc.id + '">' + esc(doc.filename) + '</a></td>'
            + '<td class="td-iss">' + (doc.issuance_date ? esc(doc.issuance_date) : '<span style="color:var(--slate-2)">—</span>') + '</td>'
            + '<td class="td-ren">' + (doc.renewal_date  ? esc(doc.renewal_date)  : '<span style="color:var(--slate-2)">—</span>') + '</td>'
            + '<td><button class="portal-training-edit-btn" data-doc-id="' + doc.id
              + '" data-iss="' + esc(doc.issuance_date) + '" data-ren="' + esc(doc.renewal_date) + '">Edit</button></td>'
            + '</tr>';
        });
        html += '</tbody></table>';
        bodyView.innerHTML = html;

        bodyView.querySelectorAll('.portal-training-link').forEach(function(a) {
          a.addEventListener('click', function(e) {
            e.preventDefault();
            openTrainingFile(parseInt(a.dataset.docId, 10));
          });
        });

        bodyView.querySelectorAll('.portal-training-edit-btn').forEach(function(btn) {
          btn.addEventListener('click', function() {
            var docId = parseInt(btn.dataset.docId, 10);
            var row   = bodyView.querySelector('tr[data-doc-id="' + docId + '"]');
            var tdIss = row.querySelector('.td-iss');
            var tdRen = row.querySelector('.td-ren');
            var curIss = btn.dataset.iss;
            var curRen = btn.dataset.ren;

            // Switch row to edit mode
            tdIss.innerHTML = '<input class="portal-date-input" placeholder="MM/DD/YYYY" value="' + esc(curIss) + '">';
            tdRen.innerHTML = '<input class="portal-date-input" placeholder="MM/DD/YYYY (optional)" value="' + esc(curRen) + '">';
            btn.textContent = 'Save';
            btn.classList.add('portal-training-save-btn');
            btn.classList.remove('portal-training-edit-btn');

            btn.addEventListener('click', function onSave() {
              btn.removeEventListener('click', onSave);
              var newIss = tdIss.querySelector('input').value.trim();
              var newRen = tdRen.querySelector('input').value.trim();
              btn.disabled = true;
              btn.textContent = '…';

              var fd = new FormData();
              fd.append('issuance_date', newIss);
              fd.append('renewal_date',  newRen);

              getApiToken().then(function(token) {
                return fetch(API_BASE + '/trainings/' + docId, {
                  method: 'PATCH',
                  headers: { 'Authorization': 'Bearer ' + token },
                  body: fd,
                });
              }).then(function(r) {
                if (!r.ok) throw new Error('Save failed (' + r.status + ')');
                return r.json();
              }).then(function(updated) {
                tdIss.innerHTML = updated.issuance_date ? esc(updated.issuance_date) : '<span style="color:var(--slate-2)">—</span>';
                tdRen.innerHTML = updated.renewal_date  ? esc(updated.renewal_date)  : '<span style="color:var(--slate-2)">—</span>';
                btn.dataset.iss = updated.issuance_date;
                btn.dataset.ren = updated.renewal_date;
                btn.textContent = 'Edit';
                btn.disabled = false;
                btn.classList.add('portal-training-edit-btn');
                btn.classList.remove('portal-training-save-btn');
              }).catch(function(err) {
                alert('Could not save dates: ' + err.message);
                btn.textContent = 'Save';
                btn.disabled = false;
              });
            }, { once: true });
          });
        });
      }).catch(function(err) {
        bodyView.innerHTML = '<p class="portal-ts-empty">Error loading documents.</p>';
        console.error('Training docs load error:', err);
      });
    }

    function openTrainingFile(docId) {
      getApiToken().then(function(token) {
        return fetch(API_BASE + '/trainings/' + docId + '/file', { headers: { 'Authorization': 'Bearer ' + token } });
      }).then(function(r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        var ct = r.headers.get('Content-Type') || '';
        return r.blob().then(function(blob) {
          var blobUrl = URL.createObjectURL(new Blob([blob], { type: ct }));
          window.open(blobUrl, '_blank');
        });
      }).catch(function(err) {
        alert('Could not open document. Please try again.');
        console.error('Training file open error:', err);
      });
    }

    // ── Upload modal ─────────────────────────────────────────────────────────
    if (btnUpload) {
      btnUpload.addEventListener('click', function() {
        if (formUpload) formUpload.reset();
        if (uploadStatus) { uploadStatus.hidden = true; uploadStatus.textContent = ''; }
        if (modalUpload) modalUpload.hidden = false;
      });
    }

    if (formUpload) {
      formUpload.addEventListener('submit', function(e) {
        e.preventDefault();
        var fileInput    = document.getElementById('training-upload-file');
        var issuedInput  = document.getElementById('training-upload-issued');
        var renewalInput = document.getElementById('training-upload-renewal');
        var submitBtn    = document.getElementById('btn-training-upload-submit');

        if (!fileInput || !fileInput.files.length) {
          showUploadStatus('error', 'Please select a file to upload.');
          return;
        }
        var issuedVal = (issuedInput ? issuedInput.value.trim() : '');
        if (!issuedVal) {
          showUploadStatus('error', 'Date of issuance is required.');
          return;
        }

        if (submitBtn) submitBtn.disabled = true;
        showUploadStatus('info', 'Uploading…');

        var fd = new FormData();
        fd.append('file',          fileInput.files[0]);
        fd.append('issuance_date', issuedVal);
        fd.append('renewal_date',  renewalInput ? renewalInput.value.trim() : '');
        if (isAdmin && adminSelect && adminSelect.value) {
          fd.append('employee_email', adminSelect.value);
          var selOpt = adminSelect.options[adminSelect.selectedIndex];
          var empName = selOpt ? selOpt.textContent.replace(/\s*\(.*\)$/, '').trim() : '';
          fd.append('employee_name', empName);
        }

        getApiToken().then(function(token) {
          return fetch(API_BASE + '/trainings/upload', {
            method: 'POST',
            headers: { 'Authorization': 'Bearer ' + token },
            body: fd,
          });
        }).then(function(r) {
          if (!r.ok) return r.json().then(function(d) { throw new Error(d.detail || 'Upload failed'); });
          return r.json();
        }).then(function() {
          showUploadStatus('ok', 'Document uploaded successfully.');
          if (formUpload) formUpload.reset();
          if (submitBtn) submitBtn.disabled = false;
        }).catch(function(err) {
          showUploadStatus('error', 'Upload failed: ' + err.message);
          if (submitBtn) submitBtn.disabled = false;
          console.error('Training upload error:', err);
        });
      });
    }

    function showUploadStatus(type, msg) {
      if (!uploadStatus) return;
      uploadStatus.hidden = false;
      uploadStatus.className = 'portal-ts-status portal-ts-status--' + (type === 'ok' ? 'success' : type === 'error' ? 'error' : 'uploading');
      uploadStatus.textContent = msg;
    }

    function esc(s) {
      return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }
  }


  function renderCalendar(bodyEl) {
    var today = new Date(); today.setHours(0, 0, 0, 0);

    var MONTHS_LONG  = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    var MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var DAYS_LONG    = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];

    function fmt(d) {
      return DAYS_LONG[d.getDay()] + ', ' + MONTHS_LONG[d.getMonth()] + ' ' + d.getDate();
    }
    function fmtShort(d) {
      return MONTHS_SHORT[d.getMonth()] + ' ' + d.getDate();
    }

    // Payday: 15th, Saturday → Monday 17th, Sunday → Monday 16th
    function payday(year, month) {
      var d = new Date(year, month, 15);
      var dow = d.getDay();
      if (dow === 6) return new Date(year, month, 17);
      if (dow === 0) return new Date(year, month, 16);
      return d;
    }

    // Next upcoming payday
    var npYear = today.getFullYear(), npMonth = today.getMonth();
    var np = payday(npYear, npMonth);
    if (np < today) {
      npMonth++;
      if (npMonth > 11) { npMonth = 0; npYear++; }
      np = payday(npYear, npMonth);
    }
    var diffDays  = Math.round((np - today) / (1000 * 60 * 60 * 24));
    var diffLabel = diffDays === 0 ? 'Today!' : diffDays === 1 ? 'Tomorrow' : 'In ' + diffDays + ' days';

    // Holidays
    var HOLIDAYS = {
      2026: [
        { name: 'New Year\'s Day',  actual: new Date(2026, 0,  1),  observed: null },
        { name: 'Memorial Day',     actual: new Date(2026, 4,  25), observed: null },
        { name: 'Independence Day', actual: new Date(2026, 6,  4),  observed: new Date(2026, 6, 3)  },
        { name: 'Labor Day',        actual: new Date(2026, 8,  7),  observed: null },
        { name: 'Thanksgiving Day', actual: new Date(2026, 10, 26), observed: null },
        { name: 'Christmas Day',    actual: new Date(2026, 11, 25), observed: null },
      ],
      2027: [
        { name: 'New Year\'s Day',  actual: new Date(2027, 0,  1),  observed: null },
        { name: 'Memorial Day',     actual: new Date(2027, 4,  31), observed: null },
        { name: 'Independence Day', actual: new Date(2027, 6,  4),  observed: new Date(2027, 6, 5)  },
        { name: 'Labor Day',        actual: new Date(2027, 8,  6),  observed: null },
        { name: 'Thanksgiving Day', actual: new Date(2027, 10, 25), observed: null },
        { name: 'Christmas Day',    actual: new Date(2027, 11, 25), observed: new Date(2027, 11, 24) },
      ]
    };

    function buildHolidayTable(holidays) {
      var html = '<table class="portal-ts-table portal-cal-holiday-table"><thead><tr><th>Holiday</th><th>Date</th></tr></thead><tbody>';
      holidays.forEach(function (h) {
        var effectiveDate = h.observed || h.actual;
        var isPast = effectiveDate < today;
        var observedNote = h.observed
          ? '<br><span class="portal-cal-observed-note">Observed: ' + fmt(h.observed) + '</span>'
          : '';
        html += '<tr' + (isPast ? ' class="portal-cal-past-row"' : '') + '>'
          + '<td class="portal-cal-holiday-name">' + escapeHtml(h.name) + '</td>'
          + '<td class="portal-cal-holiday-date">' + fmt(h.actual) + observedNote + '</td>'
          + '</tr>';
      });
      return html + '</tbody></table>';
    }

    var defaultYear = today.getFullYear() >= 2027 ? 2027 : 2026;

    // ── Render ──────────────────────────────────────────────────────────────
    var html = '';

    // Next Payday card
    html += '<div class="portal-cal-next-payday">'
      + '<div>'
        + '<div class="portal-cal-payday-label">Next Payday</div>'
        + '<div class="portal-cal-payday-date">' + fmt(np) + '</div>'
      + '</div>'
      + '<div class="portal-cal-payday-countdown">' + diffLabel + '</div>'
      + '</div>';

    // Holidays section
    html += '<div class="portal-cal-section">'
      + '<div class="portal-cal-section-head">'
        + '<span class="portal-cal-section-title">Holidays</span>'
        + '<div class="portal-cal-year-tabs">'
          + '<button class="portal-cal-year-tab' + (defaultYear === 2026 ? ' portal-cal-year-tab--active' : '') + '" data-cal-year="2026" type="button">2026</button>'
          + '<button class="portal-cal-year-tab' + (defaultYear === 2027 ? ' portal-cal-year-tab--active' : '') + '" data-cal-year="2027" type="button">2027</button>'
        + '</div>'
      + '</div>'
      + '<div id="cal-holidays-2026"' + (defaultYear !== 2026 ? ' hidden' : '') + '>' + buildHolidayTable(HOLIDAYS[2026]) + '</div>'
      + '<div id="cal-holidays-2027"' + (defaultYear !== 2027 ? ' hidden' : '') + '>' + buildHolidayTable(HOLIDAYS[2027]) + '</div>'
      + '</div>';

    // All Paydays collapsible
    html += '<div class="portal-cal-section">'
      + '<button class="portal-cal-toggle" id="cal-toggle-paydays" type="button" aria-expanded="false">&#9660; All Paydays</button>'
      + '<div id="cal-paydays-body" class="portal-cal-paydays-grid" hidden>';

    [2026, 2027].forEach(function (yr) {
      html += '<div class="portal-cal-pd-year"><div class="portal-cal-pd-yr-label">' + yr + '</div>';
      for (var m = 0; m < 12; m++) {
        var pd = payday(yr, m);
        var isPast = pd < today;
        var isObserved = pd.getDate() !== 15;
        html += '<div class="portal-cal-pd-row' + (isPast ? ' portal-cal-past-row' : '') + '">'
          + '<span class="portal-cal-pd-month">' + MONTHS_LONG[m] + '</span>'
          + '<span class="portal-cal-pd-date">' + fmtShort(pd) + (isObserved ? ' <span class="portal-cal-observed-note">obs.</span>' : '') + '</span>'
          + '</div>';
      }
      html += '</div>';
    });

    html += '</div></div>';

    bodyEl.innerHTML = html;

    // Year tab toggle
    bodyEl.querySelectorAll('[data-cal-year]').forEach(function (tab) {
      tab.addEventListener('click', function () {
        bodyEl.querySelectorAll('[data-cal-year]').forEach(function (t) { t.classList.remove('portal-cal-year-tab--active'); });
        tab.classList.add('portal-cal-year-tab--active');
        var yr = tab.dataset.calYear;
        document.getElementById('cal-holidays-2026').hidden = (yr !== '2026');
        document.getElementById('cal-holidays-2027').hidden = (yr !== '2027');
      });
    });

    // Paydays toggle
    var toggleBtn   = document.getElementById('cal-toggle-paydays');
    var paydaysBody = document.getElementById('cal-paydays-body');
    toggleBtn.addEventListener('click', function () {
      var open = !paydaysBody.hidden;
      paydaysBody.hidden = open;
      toggleBtn.setAttribute('aria-expanded', String(!open));
      toggleBtn.innerHTML = (open ? '&#9660;' : '&#9650;') + ' All Paydays';
    });
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
  // ══════════════════════════════════════════════════════════════════════════
  // Employee — Essential Functions Statement (EFS)
  // ══════════════════════════════════════════════════════════════════════════
  function setupEmployeeEfs() {
    if (IS_DEMO) return;
    var item    = document.getElementById('efs-item');
    var descEl  = document.getElementById('efs-item-desc');
    var actionsEl = document.getElementById('efs-item-actions');
    var viewBtn = document.getElementById('btn-view-my-efs');
    var signBtn = document.getElementById('btn-sign-my-efs');
    if (!item) return;

    apiCall('/efs/my').then(function (docs) {
      if (!docs || !docs.length) return;

      // Prioritise most-recent Pending; fall back to most-recent overall
      var pending = docs.filter(function (d) { return d.status === 'Pending'; });
      var current = pending.length ? pending[0] : docs[0];

      item.hidden = false;

      if (current.status === 'Signed') {
        var signedDate = current.signed_date
          ? new Date(current.signed_date).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
          : '';
        if (descEl) descEl.textContent = 'Your Essential Functions Statement is on file.' + (signedDate ? ' Signed on ' + signedDate + '.' : '');
        if (signBtn) signBtn.hidden = true;
      }

      if (viewBtn) {
        viewBtn.addEventListener('click', function () {
          viewBtn.disabled = true;
          viewBtn.textContent = 'Loading…';
          apiCall('/efs/view/' + current.id).then(function (data) {
            window.open(data.url, '_blank', 'noopener');
          }).catch(function () {
            alert('Could not load the document. Please try again.');
          }).finally(function () {
            viewBtn.disabled = false;
            viewBtn.textContent = 'View';
          });
        });
      }

      if (signBtn && current.status === 'Pending') {
        signBtn.addEventListener('click', function () {
          if (!confirm('By clicking OK you are electronically signing your Essential Functions Statement. Continue?')) return;
          signBtn.disabled = true;
          signBtn.textContent = 'Signing…';
          apiCall('/efs/sign/' + current.id, { method: 'POST' }).then(function () {
            signBtn.hidden = true;
            var badge = document.createElement('span');
            badge.className = 'portal-signoff-badge portal-signoff-badge--signed';
            badge.textContent = '✓ Signed';
            if (actionsEl) actionsEl.appendChild(badge);
            if (descEl) descEl.textContent = 'Your Essential Functions Statement is on file. Signed today.';
          }).catch(function () {
            signBtn.disabled = false;
            signBtn.textContent = 'Sign EFS';
            alert('Could not record your signature. Please try again.');
          });
        });
      }
    }).catch(function () {});
  }


  // ══════════════════════════════════════════════════════════════════════════
  // HR — EFS Issuance & Acknowledgement Tracking
  // ══════════════════════════════════════════════════════════════════════════
  function setupHREfs() {
    var issueBtn  = document.getElementById('btn-issue-efs');
    var viewBtn   = document.getElementById('btn-view-efs-records');
    var issuePanel = document.getElementById('panel-issue-efs');
    var viewPanel  = document.getElementById('panel-view-efs');
    if (!issueBtn || !viewBtn || !issuePanel || !viewPanel) return;

    var viewLoaded = false;
    var _efsPickerSelected = [];

    issueBtn.addEventListener('click', function () {
      if (issuePanel.hidden) {
        issuePanel.hidden = false;
        viewPanel.hidden  = true;
        viewBtn.textContent = 'View Acknowledgements';
        _efsPickerSelected = [];
        var efsSearch = document.getElementById('efs-emp-search');
        if (efsSearch) efsSearch.value = '';
        buildEfsPicker('');
      } else {
        issuePanel.hidden = true;
      }
    });

    viewBtn.addEventListener('click', function () {
      if (viewPanel.hidden) {
        viewPanel.hidden  = false;
        issuePanel.hidden = true;
        issueBtn.textContent = 'Issue EFS';
        if (!viewLoaded) { loadEfsRecords(); viewLoaded = true; }
      } else {
        viewPanel.hidden = true;
      }
    });

    var refreshEfsBtn = document.getElementById('btn-refresh-efs');
    if (refreshEfsBtn) {
      refreshEfsBtn.addEventListener('click', function () { viewLoaded = false; loadEfsRecords(); viewLoaded = true; });
    }

    var efsFileInput = document.getElementById('efs-file');
    if (efsFileInput) {
      efsFileInput.addEventListener('change', function () {
        var lbl = document.getElementById('efs-file-label-text');
        if (lbl && this.files.length) lbl.textContent = this.files[0].name;
      });
    }

    var efsSearchInput = document.getElementById('efs-emp-search');
    if (efsSearchInput) {
      efsSearchInput.addEventListener('input', function () { buildEfsPicker(this.value); });
    }

    var efsForm = document.getElementById('form-issue-efs');
    if (efsForm) {
      efsForm.addEventListener('submit', function (e) { e.preventDefault(); handleEfsIssue(); });
    }

    function buildEfsPicker(filter) {
      var listEl = document.getElementById('efs-emp-list');
      if (!listEl) return;

      var employees = IS_DEMO ? DEMO_STAFF : Object.keys(PORTAL_CONFIG.staff).map(function (email) {
        return { email: email, name: PORTAL_CONFIG.staff[email].name || email };
      }).sort(function (a, b) { return a.name.localeCompare(b.name); });

      var term     = (filter || '').toLowerCase();
      var filtered = term
        ? employees.filter(function (e) {
            return e.name.toLowerCase().indexOf(term) !== -1 || e.email.toLowerCase().indexOf(term) !== -1;
          })
        : employees;

      var allChecked = filtered.length > 0 && filtered.every(function (e) { return _efsPickerSelected.indexOf(e.email) !== -1; });

      var html = '<label class="portal-picker-item portal-picker-selectall">'
        + '<input type="checkbox" id="efs-picker-all" class="portal-picker-cb"' + (allChecked ? ' checked' : '') + '> '
        + '<span>Select All (' + filtered.length + ')</span></label>';

      filtered.forEach(function (emp) {
        var checked = _efsPickerSelected.indexOf(emp.email) !== -1 ? ' checked' : '';
        html += '<label class="portal-picker-item">'
          + '<input type="checkbox" class="efs-picker-emp-cb portal-picker-cb" value="' + escapeHtml(emp.email) + '"' + checked + '> '
          + '<span class="portal-picker-name">' + escapeHtml(emp.name) + '</span>'
          + '<span class="portal-picker-email">' + escapeHtml(emp.email) + '</span>'
          + '</label>';
      });
      listEl.innerHTML = html;

      var selectAll = document.getElementById('efs-picker-all');
      if (selectAll) {
        selectAll.addEventListener('change', function () {
          filtered.forEach(function (emp) {
            var idx = _efsPickerSelected.indexOf(emp.email);
            if (selectAll.checked && idx === -1) _efsPickerSelected.push(emp.email);
            else if (!selectAll.checked && idx !== -1) _efsPickerSelected.splice(idx, 1);
          });
          listEl.querySelectorAll('.efs-picker-emp-cb').forEach(function (cb) { cb.checked = selectAll.checked; });
        });
      }

      listEl.querySelectorAll('.efs-picker-emp-cb').forEach(function (cb) {
        cb.addEventListener('change', function () {
          var idx = _efsPickerSelected.indexOf(cb.value);
          if (cb.checked && idx === -1) _efsPickerSelected.push(cb.value);
          else if (!cb.checked && idx !== -1) _efsPickerSelected.splice(idx, 1);
          var nowAll = filtered.every(function (e) { return _efsPickerSelected.indexOf(e.email) !== -1; });
          if (selectAll) selectAll.checked = nowAll;
        });
      });
    }

    function handleEfsIssue() {
      var fileInput = document.getElementById('efs-file');
      var submitBtn = document.getElementById('btn-issue-efs-submit');
      if (!fileInput || !submitBtn) return;

      var file           = fileInput.files[0];
      var selectedEmails = _efsPickerSelected.slice();

      if (!file)                       { showEfsStatus('error', 'Please select an EFS document.'); return; }
      if (!selectedEmails.length)      { showEfsStatus('error', 'Please select at least one employee.'); return; }
      if (file.size > 10 * 1024 * 1024) { showEfsStatus('error', 'File exceeds 10 MB. Please use a smaller file.'); return; }

      if (IS_DEMO) {
        showEfsStatus('success', 'Demo: EFS would be issued to ' + selectedEmails.length + ' employee(s).');
        return;
      }

      submitBtn.disabled = true;
      showEfsStatus('uploading', 'Uploading EFS document…');

      var fd = new FormData();
      fd.append('file', file);
      fd.append('issued_to', JSON.stringify(selectedEmails));

      getApiToken().then(function (token) {
        return fetch(API_BASE + '/efs/issue', {
          method:  'POST',
          headers: { 'Authorization': 'Bearer ' + token },
          body:    fd,
        });
      }).then(function (r) {
        if (!r.ok) throw new Error('HTTP ' + r.status);
        showEfsStatus('success', '✓ EFS issued to ' + selectedEmails.length + ' employee(s).');
        fileInput.value = '';
        var lbl = document.getElementById('efs-file-label-text');
        if (lbl) lbl.textContent = 'Choose a PDF or Word file…';
        _efsPickerSelected = [];
        buildEfsPicker('');
        var efsSearch = document.getElementById('efs-emp-search');
        if (efsSearch) efsSearch.value = '';
        submitBtn.disabled = false;
        viewLoaded = false;
      }).catch(function () {
        showEfsStatus('error', 'Upload failed. Please try again.');
        submitBtn.disabled = false;
      });
    }

    function showEfsStatus(type, msg) {
      var el = document.getElementById('issue-efs-status');
      if (!el) return;
      el.hidden    = false;
      el.className = 'portal-ts-status portal-ts-status--' + type;
      el.textContent = msg;
      if (type === 'success') setTimeout(function () { el.hidden = true; }, 8000);
    }

    function loadEfsRecords() {
      var bodyEl = document.getElementById('body-efs-records');
      if (!bodyEl) return;
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading EFS records…</p>';
      apiCall('/efs/issued').then(function (data) {
        renderEfsRecords(data, bodyEl);
      }).catch(function () {
        bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load records. Click Refresh to try again.</p>';
      });
    }

    function renderEfsRecords(data, bodyEl) {
      if (!data || !data.length) {
        bodyEl.innerHTML = '<p class="portal-ts-empty">No EFS documents have been issued yet.</p>';
        return;
      }

      var html = '<div class="portal-ts-table-wrap"><table class="portal-ts-table"><thead><tr>'
        + '<th>Issued Date</th><th>Recipients</th><th>Signatures</th><th>Track</th></tr></thead><tbody>';

      data.forEach(function (doc) {
        var issuedDate   = doc.issued_date ? new Date(doc.issued_date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }) : '—';
        var total        = doc.acknowledgments.length;
        var signedCount  = doc.acknowledgments.filter(function (a) { return a.status === 'Signed'; }).length;
        var allSigned    = signedCount === total && total > 0;
        var statusBadge  = allSigned
          ? '<span class="portal-signoff-badge portal-signoff-badge--signed">All Signed (' + total + '/' + total + ')</span>'
          : '<span class="portal-signoff-badge portal-signoff-badge--pending">' + signedCount + '/' + total + ' Signed</span>';

        html += '<tr>'
          + '<td class="portal-ts-col-date">' + issuedDate + '</td>'
          + '<td>' + total + ' employee' + (total === 1 ? '' : 's') + '</td>'
          + '<td>' + statusBadge + '</td>'
          + '<td><button class="portal-pol-track-btn" data-efs-doc-id="' + doc.id + '" type="button">Track Signatures</button></td>'
          + '</tr>'
          + '<tr class="portal-pol-ack-row" id="efs-ack-' + doc.id + '" hidden>'
          + '<td colspan="4" class="portal-pol-ack-cell">'
          + renderEfsAckTable(doc.acknowledgments)
          + '</td></tr>';
      });

      html += '</tbody></table></div>';
      bodyEl.innerHTML = html;

      bodyEl.querySelectorAll('.portal-pol-track-btn[data-efs-doc-id]').forEach(function (btn) {
        btn.addEventListener('click', function () {
          var ackRow = document.getElementById('efs-ack-' + btn.dataset.efsDocId);
          if (!ackRow) return;
          ackRow.hidden   = !ackRow.hidden;
          btn.textContent = ackRow.hidden ? 'Track Signatures' : 'Hide Signatures';
        });
      });
    }

    function renderEfsAckTable(acks) {
      if (!acks || !acks.length) return '<p class="portal-pol-ack-empty">No recipients.</p>';
      var html = '<table class="portal-efs-ack-table"><thead><tr><th>Employee</th><th>Status</th><th>Signed Date</th></tr></thead><tbody>';
      acks.forEach(function (a) {
        var signed = a.status === 'Signed';
        var badge  = signed
          ? '<span class="portal-signoff-badge portal-signoff-badge--signed">✓ Signed</span>'
          : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';
        var date = a.signed_date
          ? new Date(a.signed_date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
          : '—';
        html += '<tr>'
          + '<td>' + escapeHtml(a.employee_name || a.employee_email) + '</td>'
          + '<td>' + badge + '</td>'
          + '<td>' + date + '</td>'
          + '</tr>';
      });
      html += '</tbody></table>';
      return html;
    }
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
        return { id: r.id, issuanceId: r.issuance_id, name: r.policy_name, issuedDate: r.issued_date, status: r.status, signedDate: r.signed_date };
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

      var nameCell = (row.issuanceId && !IS_DEMO)
        ? '<a href="#" class="portal-ts-link" data-pol-issuance-id="' + escapeHtml(String(row.issuanceId)) + '" data-pol-name="' + escapeHtml(row.name) + '">' + escapeHtml(row.name) + '</a>'
        : escapeHtml(row.name);

      html += '<tr>'
        + '<td class="portal-ts-col-name">' + nameCell + '</td>'
        + '<td class="portal-ts-col-date">' + (row.issuedDate ? formatDate(row.issuedDate) : '—') + '</td>'
        + '<td id="pol-status-' + escapeHtml(String(row.id)) + '">' + badge + '</td>'
        + '<td class="portal-pol-action">' + action + '</td>'
        + '</tr>';
    });

    html += '</tbody></table></div>';
    bodyEl.innerHTML = html;

    bodyEl.querySelectorAll('[data-pol-issuance-id]').forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        openFileInBrowser('/policies/issued/' + a.dataset.polIssuanceId + '/url', a.dataset.polName);
      });
    });

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
      bodyEl.innerHTML = '<p class="portal-ts-empty">No policies have been issued yet. Click <strong>Issue a Policy</strong> to get started.</p>';
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

      var safeId   = escapeHtml(String(pol.id));
      var polLink  = IS_DEMO
        ? escapeHtml(pol.name)
        : '<a href="#" class="portal-ts-link" data-pol-issuance-id="' + safeId + '" data-pol-name="' + escapeHtml(pol.name) + '">' + escapeHtml(pol.name) + '</a>';
      html += '<tr>'
        + '<td class="portal-ts-col-name">' + polLink + '</td>'
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

    bodyEl.querySelectorAll('[data-pol-issuance-id]').forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        openFileInBrowser('/policies/issued/' + a.dataset.polIssuanceId + '/url', a.dataset.polName);
      });
    });

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
  // Timesheet Form — employee inline form (Standard + CA)
  // ══════════════════════════════════════════════════════════════════════════
  function setupTimesheetForm(userEmail, userDisplayName, staffEntry) {
    var tsCA  = document.getElementById('ts-item-ca');
    var tsStd = document.getElementById('ts-item-std');

    var workState = (staffEntry && staffEntry.state) ? staffEntry.state.toUpperCase() : '';
    if (workState === 'CA') {
      if (tsCA) tsCA.hidden = false;
    } else if (workState) {
      if (tsStd) tsStd.hidden = false;
    } else {
      if (tsCA) tsCA.hidden = false;
      if (tsStd) tsStd.hidden = false;
    }

    var btnStd = document.getElementById('btn-open-ts-form-std');
    var btnCA  = document.getElementById('btn-open-ts-form-ca');

    var tsFormModal  = document.getElementById('modal-ts-form');
    var tsFormBody   = document.getElementById('body-ts-form-modal');
    var tsFormTitle  = document.getElementById('modal-ts-form-title');
    var tsFormClose  = document.getElementById('btn-close-ts-form-modal');
    var tsFormOverlay = document.getElementById('modal-ts-form-overlay');

    function closeTsFormModal() { if (tsFormModal) tsFormModal.hidden = true; }
    if (tsFormClose)   tsFormClose.addEventListener('click', closeTsFormModal);
    if (tsFormOverlay) tsFormOverlay.addEventListener('click', closeTsFormModal);

    // Wire Non-Billable Codes modal close (once)
    var codesModal   = document.getElementById('modal-ts-codes');
    var codesClose   = document.getElementById('btn-close-ts-codes-modal');
    var codesOverlay = document.getElementById('modal-ts-codes-overlay');
    if (codesClose)   codesClose.addEventListener('click',   function () { if (codesModal) codesModal.hidden = true; });
    if (codesOverlay) codesOverlay.addEventListener('click', function () { if (codesModal) codesModal.hidden = true; });

    function wireForm(btn, tsType) {
      if (!btn || !tsFormModal || !tsFormBody) return;
      btn.addEventListener('click', function () {
        var label = tsType === 'ca' ? 'California' : 'Standard';
        if (tsFormTitle) tsFormTitle.textContent = 'Weekly Timesheet — ' + label;
        if (tsFormBody.dataset.builtType !== tsType) {
          buildTimesheetForm(tsFormBody, tsType, userEmail, userDisplayName, staffEntry);
          tsFormBody.dataset.builtType = tsType;
        }
        tsFormModal.hidden = false;
      });
    }

    wireForm(btnStd, 'standard');
    wireForm(btnCA,  'ca');
  }


  function buildTimesheetForm(panelEl, tsType, userEmail, userDisplayName, staffEntry) {
    var isCA = tsType === 'ca';

    function lastSunday() {
      var d = new Date();
      d.setDate(d.getDate() - d.getDay());
      return d.toISOString().slice(0, 10);
    }

    var staffState   = (staffEntry && staffEntry.state)      ? staffEntry.state      : (isCA ? 'CA' : '');
    var staffManager = (staffEntry && staffEntry.manager)     ? staffEntry.manager    : '';
    var staffId      = (staffEntry && staffEntry.employeeId)  ? staffEntry.employeeId : '';
    var DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    var todayIso = new Date().toISOString().slice(0, 10);

    var html = '<div class="portal-ts-form-header">'
      + '<div class="portal-ts-form-row-inline"><label class="portal-form-label">Week Starting (Sunday)</label>'
      + '<input type="date" id="tsf-week-start" class="portal-form-input portal-ts-date-input" value="' + lastSunday() + '"></div>'
      + '<div class="portal-ts-form-row-inline"><label class="portal-form-label">Employee ID</label>'
      + '<input type="text" id="tsf-emp-id" class="portal-form-input portal-ts-short-input" value="' + escapeHtml(staffId) + '" placeholder="ID" maxlength="20"></div>'
      + (isCA ? '' : '<div class="portal-ts-form-row-inline"><label class="portal-form-label">Work State</label>'
        + '<input type="text" id="tsf-state" class="portal-form-input portal-ts-short-input" value="' + escapeHtml(staffState) + '" placeholder="e.g. TX" maxlength="2"></div>')
      + '<div class="portal-ts-form-row-inline"><label class="portal-form-label">Manager Name</label>'
      + '<input type="text" id="tsf-manager" class="portal-form-input" value="' + escapeHtml(staffManager) + '" placeholder="Manager full name" maxlength="100"></div>'
      + '</div>';

    html += '<div class="portal-ts-form-note">'
      + '<p>ONE LINE PER DAY. In <strong>Project Code</strong> put the code from the tracker — three capital letters then nine digits, e.g. <strong>ERI202607135</strong>. For a non-billable day put the LBR code instead (see Codes List). '
      + 'Enter times in 24-hour format: <strong>0900 = 9:00 AM, 1300 = 1:00 PM, 1730 = 5:30 PM, 2200 = 10:00 PM.</strong> '
      + 'Meal time is unpaid; rest breaks are paid up to 10 minutes each and any extra is deducted. Submit by <strong>Monday 12:00 PM</strong> for the prior week.</p>'
      + '<button type="button" class="portal-ts-codes-btn" id="btn-ts-codes-' + tsType + '">Non-Billable Codes</button>'
      + '</div>';

    html += '<div class="portal-ts-form-table-wrap"><table class="portal-ts-form-table" id="tsf-table">'
      + '<thead><tr><th>Day</th><th>Date</th><th>Project Code</th>'
      + '<th>Shift Start</th><th>Rest 1 Start</th><th>Rest 1 End</th>'
      + '<th>Meal Start</th><th>Meal End</th>'
      + (isCA ? '<th>Meal Waiver</th>' : '')
      + '<th>Rest 2 Start</th><th>Rest 2 End</th><th>Shift End</th>'
      + '<th>Hours</th><th>Notes</th>'
      + '</tr></thead><tbody>';

    for (var i = 0; i < 7; i++) {
      html += '<tr class="tsf-day-row" data-day-index="' + i + '">'
        + '<td class="tsf-day-label">' + DAYS[i] + '</td>'
        + '<td class="tsf-date-cell" id="tsf-date-' + i + '">—</td>'
        + '<td><input type="text" class="portal-form-input tsf-proj" data-day="' + i + '" placeholder="Code" maxlength="30"></td>'
        + '<td><input type="text" class="tsf-time tsf-shift-start portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td><input type="text" class="tsf-time tsf-rest1s      portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td><input type="text" class="tsf-time tsf-rest1e      portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td><input type="text" class="tsf-time tsf-meals       portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td><input type="text" class="tsf-time tsf-meale       portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + (isCA ? '<td><select class="tsf-meal-waiver" data-day="' + i + '"><option value=""></option><option value="Y">Y</option><option value="N">N</option></select></td>' : '')
        + '<td><input type="text" class="tsf-time tsf-rest2s      portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td><input type="text" class="tsf-time tsf-rest2e      portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td><input type="text" class="tsf-time tsf-shift-end   portal-form-input tsf-time-input" data-day="' + i + '" placeholder="HH:MM" maxlength="5"></td>'
        + '<td class="tsf-hours-cell" id="tsf-hours-' + i + '">—</td>'
        + '<td><input type="text" class="portal-form-input tsf-notes" data-day="' + i + '" placeholder="Notes" maxlength="200"></td>'
        + '</tr>';
    }
    html += '</tbody></table></div>';

    if (isCA) {
      html += '<div class="portal-ts-form-ot-summary">'
        + '<table class="portal-ts-form-ot-table">'
        + '<thead><tr><th>Day</th><th>Hours</th><th>Regular</th><th>OT 1.5×</th><th>OT 2.0×</th><th>Missed Meal</th><th>Missed Rest</th><th>Premium Hrs</th></tr></thead>'
        + '<tbody id="tsf-ot-body"></tbody><tfoot id="tsf-ot-foot"></tfoot>'
        + '</table></div>';
    } else {
      html += '<div class="portal-ts-form-summary">'
        + '<div class="tsf-summary-row"><span>Total Hours</span><span id="tsf-total-hours">0.00</span></div>'
        + '<div class="tsf-summary-row"><span>Regular (≤40h)</span><span id="tsf-reg-hours">0.00</span></div>'
        + '<div class="tsf-summary-row"><span>Overtime (&gt;40h)</span><span id="tsf-ot-hours">0.00</span></div>'
        + '</div>';
    }

    html += '<div class="portal-ts-pto-row"><label class="portal-form-label">PTO Hours this week</label>'
      + '<input type="number" id="tsf-pto" class="portal-form-input portal-ts-short-input" min="0" max="40" step="0.5" value="0"></div>';

    if (isCA) {
      html += '<div class="portal-ts-7thday-row">'
        + '<label><input type="checkbox" id="tsf-7th-day"> Saturday is my 7th consecutive workday (all hours at OT rate)</label>'
        + '</div>';
    }

    html += '<div class="portal-ts-attestation">'
      + '<div class="portal-ts-attestation-title">EMPLOYEE ATTESTATION</div>'
      + '<p class="portal-ts-attestation-text">I certify that this timesheet accurately reflects all hours I worked and all meal and rest periods I took, waived, or missed. I took every rest and meal break shown, free of duty and uninterrupted; I noted in the Notes column any meal that was interrupted or worked through. On any day I marked Meal Waiver = Y, my shift was 6 hours or less and I voluntarily chose to waive my meal period. I was not pressured to under-report hours or over-report breaks.</p>'
      + '<div class="portal-ts-sig-grid">'
      + '<div><label>Employee Signature</label><input type="text" id="tsf-emp-sig" class="portal-form-input" placeholder="Type your full name" maxlength="100"></div>'
      + '<div><label>Date</label><input type="date" id="tsf-emp-sig-date" class="portal-form-input portal-ts-date-input" value="' + todayIso + '"></div>'
      + '<div><label>Supervisor Signature <em>(filled by manager when approving)</em></label><input type="text" class="portal-form-input" placeholder="—" disabled></div>'
      + '<div><label>Date</label><input type="date" class="portal-form-input portal-ts-date-input" disabled></div>'
      + '</div></div>'
      + '<div class="portal-ts-attestation" style="margin-top:10px">'
      + '<div class="portal-ts-attestation-title">VOLUNTARY MEAL / REST BREAK WAIVER &nbsp;&mdash;&nbsp; complete ONLY if a break was voluntarily skipped</div>'
      + '<p class="portal-ts-attestation-text">I confirm that I was provided the opportunity to take all required meal and rest breaks but voluntarily chose not to take one or more of them on the dates above. In California, this also documents waiver of the first meal period on a shift of 6 hours or less, or the second meal period on a shift of 12 hours or less. This form is not to be used if a break was not provided by the company.</p>'
      + '<div class="portal-ts-sig-grid portal-ts-sig-grid--half">'
      + '<div><label>Employee Signature <em>(leave blank if no waiver applies)</em></label><input type="text" id="tsf-waiver-sig" class="portal-form-input" placeholder="Type your full name if applicable" maxlength="100"></div>'
      + '<div><label>Date</label><input type="date" id="tsf-waiver-sig-date" class="portal-form-input portal-ts-date-input"></div>'
      + '</div></div>';

    html += '<div class="portal-ts-form-submit-row">'
      + '<button type="button" id="tsf-clear-' + tsType + '" class="portal-ts-form-clear-btn">Clear Form</button>'
      + '<button type="button" id="tsf-submit-' + tsType + '" class="portal-ts-form-submit-btn">Submit Timesheet</button>'
      + '</div>'
      + '<div id="tsf-status-' + tsType + '" class="portal-ts-status" hidden></div>';

    panelEl.innerHTML = html;

    // Date label wiring
    function updateDateLabels() {
      var wsEl = panelEl.querySelector('#tsf-week-start');
      if (!wsEl || !wsEl.value) return;
      var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var p = wsEl.value.split('-');
      var sun = new Date(+p[0], +p[1] - 1, +p[2]);
      for (var d = 0; d < 7; d++) {
        var el = panelEl.querySelector('#tsf-date-' + d);
        if (!el) continue;
        var curr = new Date(sun); curr.setDate(curr.getDate() + d);
        el.textContent = MONTHS[curr.getMonth()] + ' ' + curr.getDate();
      }
    }
    var wsInput = panelEl.querySelector('#tsf-week-start');
    if (wsInput) { wsInput.addEventListener('change', updateDateLabels); updateDateLabels(); }

    function toMins(t) {
      if (!t) return null;
      var p = t.split(':');
      return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
    }

    function calcDayHours(row) {
      var s = toMins(row.querySelector('.tsf-shift-start').value);
      var e = toMins(row.querySelector('.tsf-shift-end').value);
      if (s === null || e === null) return 0;
      if (e < s) e += 1440;
      var tot = e - s;
      var ms = toMins(row.querySelector('.tsf-meals').value);
      var me = toMins(row.querySelector('.tsf-meale').value);
      if (ms !== null && me !== null) { if (me < ms) me += 1440; tot -= (me - ms); }
      return Math.max(0, tot / 60);
    }

    function recalcAll() {
      var rows   = panelEl.querySelectorAll('.tsf-day-row');
      var DLABELS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
      var cb7     = panelEl.querySelector('#tsf-7th-day');

      if (isCA) {
        var totH = 0, totReg = 0, tot15 = 0, tot20 = 0, totMM = 0, totMR = 0, totPrem = 0;
        var bodyHtml = '';

        rows.forEach(function (row, i) {
          var h = calcDayHours(row);
          var hoursEl = panelEl.querySelector('#tsf-hours-' + i);
          if (hoursEl) hoursEl.textContent = h > 0 ? h.toFixed(2) : '—';

          if (h === 0) {
            bodyHtml += '<tr><td>' + DLABELS[i] + '</td>' + '<td>—</td>'.repeat(7) + '</tr>';
            return;
          }

          var dayIs7th = cb7 && cb7.checked && (i === 6);
          var reg = 0, ot15 = 0, ot20 = 0;
          if (dayIs7th) {
            ot15 = Math.min(h, 8); ot20 = Math.max(0, h - 8);
          } else {
            reg  = Math.min(h, 8);
            ot15 = Math.max(0, Math.min(h - 8, 4));
            ot20 = Math.max(0, h - 12);
          }

          var waiver = (row.querySelector('.tsf-meal-waiver') || {}).value || '';
          var mealS  = row.querySelector('.tsf-meals').value;
          var mealE  = row.querySelector('.tsf-meale').value;
          var r1s    = row.querySelector('.tsf-rest1s').value;
          var r1e    = row.querySelector('.tsf-rest1e').value;
          var r2s    = row.querySelector('.tsf-rest2s').value;
          var r2e    = row.querySelector('.tsf-rest2e').value;

          var mm = 0, mr = 0;
          if (h > 5 && waiver !== 'Y') {
            if (!mealS || !mealE) { mm = 1; }
            else { var ms2 = toMins(mealS), me2 = toMins(mealE); if (me2 < ms2) me2 += 1440; if ((me2 - ms2) < 30) mm = 1; }
          }
          function restOk(rs, re) {
            if (!rs || !re) return false;
            var a = toMins(rs), b = toMins(re); if (b < a) b += 1440; return (b - a) >= 10;
          }
          if (h >= 3.5 && !restOk(r1s, r1e)) mr++;
          if (h >= 7   && !restOk(r2s, r2e)) mr++;

          var prem = mm + mr;
          totH += h; totReg += reg; tot15 += ot15; tot20 += ot20; totMM += mm; totMR += mr; totPrem += prem;

          function f(n) { return n > 0 ? n.toFixed(2) : '—'; }
          bodyHtml += '<tr><td>' + DLABELS[i] + '</td><td>' + h.toFixed(2) + '</td>'
            + '<td>' + f(reg) + '</td><td>' + f(ot15) + '</td><td>' + f(ot20) + '</td>'
            + '<td>' + (mm || '—') + '</td><td>' + (mr || '—') + '</td><td>' + (prem || '—') + '</td></tr>';
        });

        var bodyEl2 = panelEl.querySelector('#tsf-ot-body');
        var footEl  = panelEl.querySelector('#tsf-ot-foot');
        if (bodyEl2) bodyEl2.innerHTML = bodyHtml;
        if (footEl) footEl.innerHTML = '<tr class="tsf-ot-total"><td><strong>Total</strong></td>'
          + '<td><strong>' + totH.toFixed(2) + '</strong></td>'
          + '<td><strong>' + totReg.toFixed(2) + '</strong></td>'
          + '<td><strong>' + tot15.toFixed(2) + '</strong></td>'
          + '<td><strong>' + tot20.toFixed(2) + '</strong></td>'
          + '<td><strong>' + (totMM || '—') + '</strong></td>'
          + '<td><strong>' + (totMR || '—') + '</strong></td>'
          + '<td><strong>' + (totPrem || '—') + '</strong></td></tr>';

      } else {
        var totalH = 0;
        rows.forEach(function (row, i) {
          var h = calcDayHours(row);
          var hoursEl = panelEl.querySelector('#tsf-hours-' + i);
          if (hoursEl) hoursEl.textContent = h > 0 ? h.toFixed(2) : '—';
          totalH += h;
        });
        var regH = Math.min(totalH, 40), otH = Math.max(0, totalH - 40);
        var t = panelEl.querySelector('#tsf-total-hours');
        var r = panelEl.querySelector('#tsf-reg-hours');
        var o = panelEl.querySelector('#tsf-ot-hours');
        if (t) t.textContent = totalH.toFixed(2);
        if (r) r.textContent = regH.toFixed(2);
        if (o) o.textContent = otH.toFixed(2);
      }
    }

    panelEl.querySelectorAll('.tsf-time, .tsf-meal-waiver').forEach(function (el) {
      el.addEventListener('change', recalcAll);
      el.addEventListener('input',  recalcAll);
    });
    var cb7El = panelEl.querySelector('#tsf-7th-day');
    if (cb7El) cb7El.addEventListener('change', recalcAll);
    recalcAll();

    var codesBtn = panelEl.querySelector('#btn-ts-codes-' + tsType);
    if (codesBtn) {
      codesBtn.addEventListener('click', function () {
        var m = document.getElementById('modal-ts-codes');
        if (m) m.hidden = false;
      });
    }

    var submitBtn = panelEl.querySelector('#tsf-submit-' + tsType);
    if (submitBtn) {
      submitBtn.addEventListener('click', function () {
        submitTimesheetForm(panelEl, tsType, userEmail, userDisplayName, staffEntry, isCA);
      });
    }

    var clearBtn = panelEl.querySelector('#tsf-clear-' + tsType);
    if (clearBtn) {
      clearBtn.addEventListener('click', function () {
        var wsEl2 = panelEl.querySelector('#tsf-week-start');
        if (wsEl2) { wsEl2.value = lastSunday(); updateDateLabels(); }
        panelEl.querySelectorAll('.tsf-time').forEach(function (el) { el.value = ''; });
        panelEl.querySelectorAll('.tsf-proj, .tsf-notes').forEach(function (el) { el.value = ''; });
        panelEl.querySelectorAll('.tsf-meal-waiver').forEach(function (el) { el.value = ''; });
        var ptoEl2 = panelEl.querySelector('#tsf-pto');
        if (ptoEl2) ptoEl2.value = '0';
        var cb7El2 = panelEl.querySelector('#tsf-7th-day');
        if (cb7El2) cb7El2.checked = false;
        recalcAll();
      });
    }
  }


  function submitTimesheetForm(panelEl, tsType, userEmail, userDisplayName, staffEntry, isCA) {
    var weekStartEl    = panelEl.querySelector('#tsf-week-start');
    var empIdEl        = panelEl.querySelector('#tsf-emp-id');
    var stateEl        = panelEl.querySelector('#tsf-state');
    var managerEl      = panelEl.querySelector('#tsf-manager');
    var ptoEl          = panelEl.querySelector('#tsf-pto');
    var empSigEl       = panelEl.querySelector('#tsf-emp-sig');
    var empSigDateEl   = panelEl.querySelector('#tsf-emp-sig-date');
    var waiverSigEl    = panelEl.querySelector('#tsf-waiver-sig');
    var waiverDateEl   = panelEl.querySelector('#tsf-waiver-sig-date');
    var statusEl       = panelEl.querySelector('#tsf-status-' + tsType);
    var submitBtn      = panelEl.querySelector('#tsf-submit-' + tsType);

    function showSt(type, msg) {
      if (!statusEl) return;
      statusEl.hidden = false;
      statusEl.className = 'portal-ts-status portal-ts-status--' + type;
      statusEl.textContent = msg;
      if (type === 'success') setTimeout(function () { statusEl.hidden = true; }, 9000);
    }

    if (!weekStartEl || !weekStartEl.value) { showSt('error', 'Please select the week starting date.'); return; }

    function toMins(t) {
      if (!t) return null;
      var p = t.split(':'); return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
    }

    var cb7 = panelEl.querySelector('#tsf-7th-day');
    var is7thChecked = cb7 && cb7.checked;
    var DAYS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    var parts = weekStartEl.value.split('-');
    var sunday = new Date(+parts[0], +parts[1] - 1, +parts[2]);
    var days = [];

    panelEl.querySelectorAll('.tsf-day-row').forEach(function (row, i) {
      var curr = new Date(sunday); curr.setDate(curr.getDate() + i);
      var dateStr = curr.getFullYear() + '-' + String(curr.getMonth()+1).padStart(2,'0') + '-' + String(curr.getDate()).padStart(2,'0');
      var ss = row.querySelector('.tsf-shift-start').value;
      var se = row.querySelector('.tsf-shift-end').value;
      var ms = row.querySelector('.tsf-meals').value;
      var me = row.querySelector('.tsf-meale').value;
      var h = 0;
      if (ss && se) {
        var sv = toMins(ss), ev = toMins(se);
        if (ev < sv) ev += 1440;
        var tot = ev - sv;
        if (ms && me) { var msv = toMins(ms), mev = toMins(me); if (mev < msv) mev += 1440; tot -= (mev - msv); }
        h = Math.max(0, tot / 60);
      }
      var waiverEl = row.querySelector('.tsf-meal-waiver');
      days.push({
        date:        dateStr,
        day:         DAYS[i],
        projectCode: (row.querySelector('.tsf-proj')   || {}).value || '',
        shiftStart:  ss,
        rest1Start:  row.querySelector('.tsf-rest1s').value,
        rest1End:    row.querySelector('.tsf-rest1e').value,
        mealStart:   ms,
        mealEnd:     me,
        mealWaiver:  waiverEl ? waiverEl.value : '',
        rest2Start:  row.querySelector('.tsf-rest2s').value,
        rest2End:    row.querySelector('.tsf-rest2e').value,
        shiftEnd:    se,
        hoursWorked: Math.round(h * 100) / 100,
        is7thDay:    isCA && is7thChecked && (i === 6),
        notes:       (row.querySelector('.tsf-notes')  || {}).value || ''
      });
    });

    if (!days.some(function (d) { return d.shiftStart || d.shiftEnd; })) {
      showSt('error', 'Please fill in at least one day\'s shift times.');
      return;
    }

    if (IS_DEMO) {
      showSt('success', '✓ Demo: Timesheet for week of ' + weekStartEl.value + ' would be submitted.');
      return;
    }

    submitBtn.disabled = true;
    showSt('uploading', 'Submitting timesheet…');

    apiCall('/timesheets/form/submit', {
      method: 'POST',
      body: JSON.stringify({
        ts_type:           tsType,
        week_start:        weekStartEl.value,
        employee_id:       empIdEl      ? empIdEl.value.trim()      : '',
        work_state:        stateEl      ? stateEl.value.trim().toUpperCase() : (isCA ? 'CA' : ''),
        manager_name:      managerEl    ? managerEl.value.trim()    : '',
        pto_hours:         ptoEl        ? parseFloat(ptoEl.value || '0') : 0,
        employee_sig:      empSigEl     ? empSigEl.value.trim()     : '',
        employee_sig_date: empSigDateEl ? empSigDateEl.value        : '',
        waiver_sig:        waiverSigEl  ? waiverSigEl.value.trim()  : '',
        waiver_sig_date:   waiverDateEl ? waiverDateEl.value        : '',
        days:              days
      })
    }).then(function () {
      showSt('success', '✓ Timesheet submitted for week of ' + weekStartEl.value + '. Your manager and HR have been notified.');
      submitBtn.disabled = false;
    }).catch(function (err) {
      console.error('TS form submit error:', err);
      showSt('error', 'Submission failed. Please try again or contact IT.');
      submitBtn.disabled = false;
    });
  }


  // ══════════════════════════════════════════════════════════════════════════
  // Form Timesheet Viewer — manager & HR
  // ══════════════════════════════════════════════════════════════════════════
  function setupFormTimesheetViewer(role, emails) {
    var isHRRole   = (role === 'hr');
    var btnId      = isHRRole ? 'btn-view-hr-ts-form'       : 'btn-view-manager-ts-form';
    var panelId    = isHRRole ? 'panel-hr-ts-form'           : 'panel-manager-ts-form';
    var bodyId     = isHRRole ? 'body-hr-ts-form'            : 'body-manager-ts-form';
    var refreshId  = isHRRole ? 'btn-refresh-hr-ts-form'     : 'btn-refresh-manager-ts-form';

    var btn     = document.getElementById(btnId);
    var panel   = document.getElementById(panelId);
    var bodyEl  = document.getElementById(bodyId);
    var refresh = document.getElementById(refreshId);
    if (!btn || !panel || !bodyEl) return;

    var loaded = false;

    btn.addEventListener('click', function () {
      if (panel.hidden) {
        panel.hidden = false;
        btn.textContent = 'Hide Timesheets';
        if (!loaded) { doLoad(); loaded = true; }
      } else {
        panel.hidden = true;
        btn.textContent = 'View Timesheets';
      }
    });

    if (refresh) refresh.addEventListener('click', function () { loaded = false; doLoad(); loaded = true; });

    function doLoad() {
      bodyEl.innerHTML = '<p class="portal-ts-loading">Loading timesheets…</p>';

      if (IS_DEMO) {
        renderFormTsList([], bodyEl, isHRRole);
        return;
      }

      var path = isHRRole
        ? '/timesheets/form/all'
        : (emails && emails.length ? '/timesheets/form/team?emails=' + encodeURIComponent(emails.join(',')) : '/timesheets/form/team?emails=');

      apiCall(path).then(function (data) {
        renderFormTsList(data, bodyEl, isHRRole);
      }).catch(function (err) {
        console.error('Form TS load error:', err);
        bodyEl.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load. Click Refresh to try again.</p>';
      });
    }
  }


  function renderFormTsList(rows, bodyEl, isHR) {
    if (!rows || !rows.length) {
      bodyEl.innerHTML = '<p class="portal-ts-empty">No timesheet submissions yet.</p>';
      return;
    }

    var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    function fmtWeek(ds) {
      if (!ds) return '—';
      var p = ds.split('-'); var d = new Date(+p[0], +p[1]-1, +p[2]);
      return 'Week of ' + MONTHS[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
    }

    var html = '<div class="portal-ts-wrap"><table class="portal-ts-table">'
      + '<thead><tr><th>Employee</th><th>Week</th><th>Type</th><th>Submitted</th><th>Status</th>'
      + (isHR ? '<th>Excel</th>' : '<th>Action</th>')
      + '</tr></thead><tbody>';

    rows.forEach(function (row) {
      var st = row.status || 'pending';
      var badge = st === 'approved'
        ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
        : st === 'rejected'
          ? '<span class="portal-pto-badge--denied">Rejected</span>'
          : '<span class="portal-signoff-badge portal-signoff-badge--pending">Pending</span>';

      var weekLink = '<a href="#" class="portal-ts-link portal-tsf-detail-link" data-tsf-id="' + row.id + '" data-tsf-status="' + st + '" data-tsf-role="' + (isHR ? 'hr' : 'manager') + '">' + escapeHtml(fmtWeek(row.week_start)) + '</a>';

      var actionCell = '';
      if (isHR) {
        actionCell = '<span class="portal-ts-nolink">—</span>';
      } else {
        actionCell = st === 'pending'
          ? '<div class="portal-pto-actions" id="tsf-act-' + row.id + '">'
            + '<button class="portal-pto-approve portal-tsf-act-btn" data-tsf-id="' + row.id + '" data-action="approve" type="button" title="Approve">&#10003;</button>'
            + '<button class="portal-pto-deny    portal-tsf-act-btn" data-tsf-id="' + row.id + '" data-action="reject"  type="button" title="Reject">&#10007;</button>'
            + '</div>'
          : '<span class="portal-pto-actioned">—</span>';
      }

      html += '<tr>'
        + '<td class="portal-ts-col-name">' + escapeHtml(row.employee_name || row.employee_email) + '</td>'
        + '<td class="portal-ts-col-file">' + weekLink + '</td>'
        + '<td>' + escapeHtml((row.ts_type || 'standard').toUpperCase()) + '</td>'
        + '<td class="portal-ts-col-date">' + formatDate(row.submitted_at) + '</td>'
        + '<td id="tsf-status-row-' + row.id + '">' + badge + '</td>'
        + '<td>' + actionCell + '</td>'
        + '</tr>';
    });

    html += '</tbody></table></div>';
    bodyEl.innerHTML = html;

    var tsDetailModal   = document.getElementById('modal-ts-detail');
    var tsDetailBody    = document.getElementById('body-ts-detail-modal');
    var tsDetailTitle   = document.getElementById('modal-ts-detail-title');
    var tsDetailClose   = document.getElementById('btn-close-ts-detail-modal');
    var tsDetailOverlay = document.getElementById('modal-ts-detail-overlay');
    if (!tsDetailModal._wired) {
      tsDetailModal._wired = true;
      if (tsDetailClose)   tsDetailClose.addEventListener('click',   function () { tsDetailModal.hidden = true; });
      if (tsDetailOverlay) tsDetailOverlay.addEventListener('click', function () { tsDetailModal.hidden = true; });
    }

    bodyEl.querySelectorAll('.portal-tsf-detail-link').forEach(function (a) {
      a.addEventListener('click', function (e) {
        e.preventDefault();
        var id        = a.dataset.tsfId;
        var role      = a.dataset.tsfRole || 'manager';
        var status    = a.dataset.tsfStatus || 'pending';
        var weekLabel = a.textContent;
        if (tsDetailTitle) tsDetailTitle.textContent = weekLabel;
        if (tsDetailBody)  tsDetailBody.innerHTML = '<p class="portal-ts-loading">Loading…</p>';
        if (tsDetailModal) tsDetailModal.hidden = false;
        apiCall('/timesheets/form/' + id).then(function (data) {
          if (!tsDetailBody) return;
          tsDetailBody.innerHTML = renderFormTsDetail(data, { role: role, submissionId: id, status: status });

          // Wire manager approval actions in detail modal
          var appBtn = tsDetailBody.querySelector('.tsfd-approve-btn');
          var rejBtn = tsDetailBody.querySelector('.tsfd-reject-btn');
          var xlsBtn = tsDetailBody.querySelector('.tsfd-excel-btn');

          if (appBtn) {
            appBtn.addEventListener('click', function () {
              var mgrSigEl  = tsDetailBody.querySelector('#tsfd-mgr-sig');
              var mgrDateEl = tsDetailBody.querySelector('#tsfd-mgr-sig-date');
              var payload = {
                manager_sig:      mgrSigEl  ? mgrSigEl.value.trim()  : '',
                manager_sig_date: mgrDateEl ? mgrDateEl.value        : ''
              };
              appBtn.disabled = true; appBtn.textContent = 'Approving…';
              apiCall('/timesheets/form/' + id + '/approve', { method: 'POST', body: JSON.stringify(payload) })
                .then(function () {
                  if (tsDetailModal) tsDetailModal.hidden = true;
                  var stCell = document.getElementById('tsf-status-row-' + id);
                  var actEl  = document.getElementById('tsf-act-' + id);
                  if (stCell) stCell.innerHTML = '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>';
                  if (actEl)  actEl.outerHTML = '<span class="portal-pto-actioned">—</span>';
                  a.dataset.tsfStatus = 'approved';
                })
                .catch(function (err) { alert('Approval failed: ' + err.message); appBtn.disabled = false; appBtn.textContent = 'Approve'; });
            });
          }

          if (rejBtn) {
            rejBtn.addEventListener('click', function () {
              rejBtn.disabled = true; rejBtn.textContent = 'Rejecting…';
              apiCall('/timesheets/form/' + id + '/reject', { method: 'POST' })
                .then(function () {
                  if (tsDetailModal) tsDetailModal.hidden = true;
                  var stCell = document.getElementById('tsf-status-row-' + id);
                  var actEl  = document.getElementById('tsf-act-' + id);
                  if (stCell) stCell.innerHTML = '<span class="portal-pto-badge--denied">Rejected</span>';
                  if (actEl)  actEl.outerHTML = '<span class="portal-pto-actioned">—</span>';
                  a.dataset.tsfStatus = 'rejected';
                })
                .catch(function (err) { alert('Rejection failed: ' + err.message); rejBtn.disabled = false; rejBtn.textContent = 'Reject'; });
            });
          }

          if (xlsBtn) {
            xlsBtn.addEventListener('click', function () {
              xlsBtn.disabled = true; xlsBtn.textContent = 'Preparing…';
              getApiToken().then(function (token) {
                return fetch(API_BASE + '/timesheets/form/' + id + '/excel', { headers: { 'Authorization': 'Bearer ' + token } });
              }).then(function (r) {
                if (!r.ok) throw new Error('HTTP ' + r.status);
                return r.blob();
              }).then(function (blob) {
                var url = URL.createObjectURL(blob);
                var dl = document.createElement('a'); dl.href = url; dl.download = 'timesheet-' + id + '.xlsx';
                document.body.appendChild(dl); dl.click(); document.body.removeChild(dl);
                setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
              }).catch(function () { alert('Could not generate Excel. Please try again.'); })
                .then(function () { xlsBtn.disabled = false; xlsBtn.textContent = 'Download Excel'; });
            });
          }
        }).catch(function () {
          if (tsDetailBody) tsDetailBody.innerHTML = '<p class="portal-ts-empty portal-ts-empty--error">Could not load detail.</p>';
        });
      });
    });

    bodyEl.querySelectorAll('.portal-tsf-act-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var id     = btn.dataset.tsfId;
        var action = btn.dataset.action;
        var actEl  = document.getElementById('tsf-act-' + id);
        var stCell = document.getElementById('tsf-status-row-' + id);
        if (actEl) actEl.querySelectorAll('button').forEach(function (b) { b.disabled = true; });
        apiCall('/timesheets/form/' + id + '/' + action, { method: 'POST' })
          .then(function () {
            if (actEl)  actEl.outerHTML = '<span class="portal-pto-actioned">—</span>';
            if (stCell) stCell.innerHTML = action === 'approve'
              ? '<span class="portal-signoff-badge portal-signoff-badge--signed">Approved</span>'
              : '<span class="portal-pto-badge--denied">Rejected</span>';
          })
          .catch(function (err) {
            console.error('TSF action error:', err);
            if (actEl) actEl.querySelectorAll('button').forEach(function (b) { b.disabled = false; });
          });
      });
    });

  }


  function renderFormTsDetail(data, opts) {
    opts = opts || {};
    var viewRole  = opts.role || 'manager';
    var subId     = opts.submissionId || '';
    var subStatus = opts.status || data.status || 'pending';
    var isCA = data.ts_type === 'ca';
    var days = [];
    if (Array.isArray(data.form_data)) {
      days = data.form_data;
    } else {
      try { days = JSON.parse(data.form_data); } catch (e) { days = []; }
    }

    var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    function fmtDate(ds) {
      if (!ds) return '—';
      var p = ds.split('-'); var d = new Date(+p[0], +p[1]-1, +p[2]);
      return MONTHS[d.getMonth()] + ' ' + d.getDate();
    }

    var html = '<div class="portal-tsf-detail">';
    html += '<div class="portal-tsf-detail-meta">'
      + '<span><strong>Employee:</strong> ' + escapeHtml(data.employee_name || data.employee_email) + '</span>'
      + '<span><strong>Week of:</strong> ' + escapeHtml(data.week_start || '') + '</span>'
      + (data.manager_name ? '<span><strong>Manager:</strong> ' + escapeHtml(data.manager_name) + '</span>' : '')
      + (data.work_state   ? '<span><strong>State:</strong> '   + escapeHtml(data.work_state)   + '</span>' : '')
      + '</div>';

    html += '<div class="portal-ts-form-table-wrap"><table class="portal-ts-form-table">'
      + '<thead><tr><th>Day</th><th>Date</th><th>Project</th><th>Shift Start</th>'
      + '<th>Rest 1</th><th>Meal</th>'
      + (isCA ? '<th>Waiver</th>' : '')
      + '<th>Rest 2</th><th>Shift End</th><th>Hours</th><th>Notes</th>'
      + '</tr></thead><tbody>';

    var totalH = 0;
    days.forEach(function (d) {
      var r1  = (d.rest1Start && d.rest1End) ? d.rest1Start + '–' + d.rest1End : '—';
      var meal = (d.mealStart  && d.mealEnd)  ? d.mealStart  + '–' + d.mealEnd  : '—';
      var r2  = (d.rest2Start && d.rest2End) ? d.rest2Start + '–' + d.rest2End : '—';
      totalH += (d.hoursWorked || 0);
      html += '<tr>'
        + '<td class="tsf-day-label">'  + escapeHtml(d.day   || '') + '</td>'
        + '<td class="tsf-date-cell">'  + escapeHtml(fmtDate(d.date)) + '</td>'
        + '<td>' + escapeHtml(d.projectCode || '—') + '</td>'
        + '<td>' + escapeHtml(d.shiftStart  || '—') + '</td>'
        + '<td>' + escapeHtml(r1) + '</td>'
        + '<td>' + escapeHtml(meal) + '</td>'
        + (isCA ? '<td>' + escapeHtml(d.mealWaiver || '—') + '</td>' : '')
        + '<td>' + escapeHtml(r2) + '</td>'
        + '<td>' + escapeHtml(d.shiftEnd   || '—') + '</td>'
        + '<td class="tsf-hours-cell">' + (d.hoursWorked > 0 ? d.hoursWorked.toFixed(2) : '—') + '</td>'
        + '<td>' + escapeHtml(d.notes || '') + '</td>'
        + '</tr>';
    });
    html += '<tr class="tsf-ot-total"><td colspan="' + (isCA ? '9' : '8') + '" style="text-align:right;padding-right:8px"><strong>Total Hours</strong></td>'
      + '<td class="tsf-hours-cell"><strong>' + totalH.toFixed(2) + '</strong></td><td></td></tr>';
    html += '</tbody></table></div>';

    if (isCA) {
      function toM(t) { if (!t) return null; var p = t.split(':'); return parseInt(p[0],10)*60+parseInt(p[1],10); }
      function rOk(rs, re) { if (!rs||!re) return false; var a=toM(rs),b=toM(re); if(b<a)b+=1440; return (b-a)>=10; }
      function fv(n) { return n > 0 ? n.toFixed(2) : '—'; }
      var tH=0,tReg=0,t15=0,t20=0,tMM=0,tMR=0,tPr=0;
      var otRows = '';
      days.forEach(function (d) {
        var h = d.hoursWorked || 0;
        if (h === 0) { otRows += '<tr><td>' + escapeHtml(d.day||'') + '</td>' + '<td>—</td>'.repeat(7) + '</tr>'; return; }
        var is7 = d.is7thDay || false;
        var reg=0, ot15=0, ot20=0;
        if (is7) { ot15=Math.min(h,8); ot20=Math.max(0,h-8); }
        else { reg=Math.min(h,8); ot15=Math.max(0,Math.min(h-8,4)); ot20=Math.max(0,h-12); }
        var mm=0, mr=0;
        if (h>5 && d.mealWaiver!=='Y') {
          if (!d.mealStart||!d.mealEnd) { mm=1; }
          else { var ms2=toM(d.mealStart),me2=toM(d.mealEnd); if(me2<ms2)me2+=1440; if((me2-ms2)<30)mm=1; }
        }
        if (h>=3.5 && !rOk(d.rest1Start,d.rest1End)) mr++;
        if (h>=7   && !rOk(d.rest2Start,d.rest2End)) mr++;
        var pr=mm+mr;
        tH+=h; tReg+=reg; t15+=ot15; t20+=ot20; tMM+=mm; tMR+=mr; tPr+=pr;
        otRows += '<tr><td>' + escapeHtml(d.day||'') + '</td><td>' + h.toFixed(2) + '</td>'
          + '<td>' + fv(reg) + '</td><td>' + fv(ot15) + '</td><td>' + fv(ot20) + '</td>'
          + '<td>' + (mm||'—') + '</td><td>' + (mr||'—') + '</td><td>' + (pr||'—') + '</td></tr>';
      });
      html += '<div class="portal-ts-form-ot-summary" style="margin-top:16px">'
        + '<table class="portal-ts-form-ot-table"><thead><tr>'
        + '<th>Day</th><th>Hours</th><th>Regular</th><th>OT 1.5\xd7</th><th>OT 2.0\xd7</th><th>Missed Meal</th><th>Missed Rest</th><th>Premium Hrs</th>'
        + '</tr></thead><tbody>' + otRows + '</tbody>'
        + '<tfoot><tr class="tsf-ot-total">'
        + '<td><strong>Total</strong></td><td><strong>' + tH.toFixed(2) + '</strong></td>'
        + '<td><strong>' + tReg.toFixed(2) + '</strong></td><td><strong>' + t15.toFixed(2) + '</strong></td><td><strong>' + t20.toFixed(2) + '</strong></td>'
        + '<td><strong>' + (tMM||'—') + '</strong></td><td><strong>' + (tMR||'—') + '</strong></td><td><strong>' + (tPr||'—') + '</strong></td>'
        + '</tr></tfoot></table></div>';
    }

    if (data.pto_hours > 0) {
      html += '<p style="margin:8px 0;font-size:.83rem;color:var(--slate)">PTO Hours: <strong>' + data.pto_hours + '</strong></p>';
    }

    // Attestation / signature section
    var todayIso2 = new Date().toISOString().slice(0, 10);
    html += '<div class="portal-ts-attestation" style="margin-top:16px">'
      + '<div class="portal-ts-attestation-title">EMPLOYEE ATTESTATION</div>'
      + '<p class="portal-ts-attestation-text">I certify that this timesheet accurately reflects all hours I worked and all meal and rest periods I took, waived, or missed. I took every rest and meal break shown, free of duty and uninterrupted; I noted in the Notes column any meal that was interrupted or worked through. On any day I marked Meal Waiver = Y, my shift was 6 hours or less and I voluntarily chose to waive my meal period. I was not pressured to under-report hours or over-report breaks.</p>'
      + '<div class="portal-ts-sig-grid">'
      + '<div><label>Employee Signature</label><input type="text" class="portal-form-input" value="' + escapeHtml(data.employee_sig || '') + '" readonly></div>'
      + '<div><label>Date</label><input type="date" class="portal-form-input portal-ts-date-input" value="' + escapeHtml(data.employee_sig_date || '') + '" readonly></div>';

    if (viewRole === 'manager' && subStatus === 'pending') {
      html += '<div><label>Supervisor Signature</label><input type="text" id="tsfd-mgr-sig" class="portal-form-input" placeholder="Your full name" maxlength="100" value="' + escapeHtml(data.manager_sig || '') + '"></div>'
        + '<div><label>Date</label><input type="date" id="tsfd-mgr-sig-date" class="portal-form-input portal-ts-date-input" value="' + escapeHtml(data.manager_sig_date || todayIso2) + '"></div>';
    } else {
      html += '<div><label>Supervisor Signature</label><input type="text" class="portal-form-input" value="' + escapeHtml(data.manager_sig || '') + '" readonly></div>'
        + '<div><label>Date</label><input type="date" class="portal-form-input portal-ts-date-input" value="' + escapeHtml(data.manager_sig_date || '') + '" readonly></div>';
    }
    html += '</div></div>'
      + '<div class="portal-ts-attestation" style="margin-top:10px">'
      + '<div class="portal-ts-attestation-title">VOLUNTARY MEAL / REST BREAK WAIVER &nbsp;&mdash;&nbsp; complete ONLY if a break was voluntarily skipped</div>'
      + '<p class="portal-ts-attestation-text">I confirm that I was provided the opportunity to take all required meal and rest breaks but voluntarily chose not to take one or more of them on the dates above. In California, this also documents waiver of the first meal period on a shift of 6 hours or less, or the second meal period on a shift of 12 hours or less. This form is not to be used if a break was not provided by the company.</p>'
      + '<div class="portal-ts-sig-grid portal-ts-sig-grid--half">'
      + '<div><label>Employee Signature</label><input type="text" class="portal-form-input" value="' + escapeHtml(data.waiver_sig || '') + '" readonly></div>'
      + '<div><label>Date</label><input type="date" class="portal-form-input portal-ts-date-input" value="' + escapeHtml(data.waiver_sig_date || '') + '" readonly></div>'
      + '</div></div>';

    // Role-specific action buttons
    if (viewRole === 'manager' && subStatus === 'pending') {
      html += '<div class="portal-ts-form-submit-row" style="margin-top:16px">'
        + '<button type="button" class="tsfd-reject-btn portal-ts-form-clear-btn">Reject</button>'
        + '<button type="button" class="tsfd-approve-btn portal-ts-form-submit-btn">Approve &amp; Sign</button>'
        + '</div>';
    } else if (viewRole === 'hr') {
      html += '<div class="portal-ts-form-submit-row" style="margin-top:16px">'
        + '<button type="button" class="tsfd-excel-btn portal-ts-form-submit-btn">Download Excel</button>'
        + '</div>';
    }

    html += '</div>';
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
        if (!IS_DEMO) openFileInBrowser('/timesheets/url/' + a.dataset.tsId, a.dataset.tsName);
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

  function openFileInBrowser(apiPath, filename) {
    apiCall(apiPath).then(function (data) {
      var ext = (filename || '').split('.').pop().toLowerCase();
      var url = data.url;
      if (ext === 'pdf') {
        window.open(url, '_blank');
      } else {
        window.open('https://view.officeapps.live.com/op/view.aspx?src=' + encodeURIComponent(url), '_blank');
      }
    }).catch(function (err) {
      console.error('Could not open file:', err);
      alert('Could not open file. Please try again.');
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
  // Home greeting + at-a-glance strip
  // ══════════════════════════════════════════════════════════════════════════
  function setupHomeGreeting(userDisplayName, isManager, isHR, team) {
    var welcomeEl  = document.getElementById('home-welcome');
    var greetingEl = document.getElementById('home-greeting');
    var glanceEl   = document.getElementById('home-glance');
    if (!welcomeEl || !greetingEl || !glanceEl) return;

    // Time-of-day greeting
    var hour      = new Date().getHours();
    var tod       = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
    var firstName = userDisplayName.split(' ')[0];
    greetingEl.textContent = tod + ', ' + firstName + '.';
    welcomeEl.hidden = false;

    // Next payday
    function paydayFor(yr, mo) {
      var d = new Date(yr, mo, 15);
      if (d.getDay() === 6) d.setDate(17);
      if (d.getDay() === 0) d.setDate(16);
      return d;
    }
    var now   = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var pd    = paydayFor(today.getFullYear(), today.getMonth());
    if (pd < today) {
      var nm = today.getMonth() === 11 ? 0 : today.getMonth() + 1;
      var ny = today.getMonth() === 11 ? today.getFullYear() + 1 : today.getFullYear();
      pd = paydayFor(ny, nm);
    }
    var diffDays = Math.round((pd - today) / 86400000);
    var pdStr    = pd.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    var pdSub    = diffDays === 0 ? 'Today — payday! 🎉' : 'in ' + diffDays + ' day' + (diffDays === 1 ? '' : 's');

    // Populate hero ticker
    var tickerEl    = document.getElementById('hero-ticker');
    var tickerInner = document.getElementById('hero-ticker-inner');
    if (tickerEl && tickerInner) {
      var dir       = PORTAL_CONFIG.directory || [];
      var empCount  = dir.length;
      var countries = dir.reduce(function (acc, e) {
        if (e.country && acc.indexOf(e.country) === -1) acc.push(e.country);
        return acc;
      }, []);
      tickerInner.textContent = empCount + ' employees  •  ' + countries.join(' · ');
      tickerEl.hidden = false;
    }

    glanceEl.innerHTML = '<div class="portal-glance-card portal-glance-card--highlight">'
      + '<span class="portal-glance-card__label">Next Payday</span>'
      + '<span class="portal-glance-card__value">' + escapeHtml(pdStr) + '</span>'
      + '<span class="portal-glance-card__sub">' + escapeHtml(pdSub) + '</span>'
      + '</div>';

    // Pending policy signatures — shown for all roles
    if (!IS_DEMO) {
      apiCall('/policies/my').then(function (data) {
        var pending = (data || []).filter(function (p) { return p.status === 'Pending'; }).length;
        glanceEl.innerHTML += '<div class="portal-glance-card' + (pending > 0 ? ' portal-glance-card--alert' : '') + '">'
          + '<span class="portal-glance-card__label">Policies to Sign</span>'
          + '<span class="portal-glance-card__value">' + pending + '</span>'
          + '<span class="portal-glance-card__sub">' + (pending === 0 ? 'All up to date' : 'Awaiting your signature') + '</span>'
          + '</div>';
      }).catch(function () {});
    }

    // Pending timesheet reviews — managers with a team
    if (isManager && !IS_DEMO && team && team.length) {
      apiCall('/timesheets/team?emails=' + encodeURIComponent(team.join(','))).then(function (data) {
        var pending = (data || []).filter(function (t) { return t.status === 'pending'; }).length;
        glanceEl.innerHTML += '<div class="portal-glance-card' + (pending > 0 ? ' portal-glance-card--alert' : '') + '">'
          + '<span class="portal-glance-card__label">Timesheets to Review</span>'
          + '<span class="portal-glance-card__value">' + pending + '</span>'
          + '<span class="portal-glance-card__sub">' + (pending === 0 ? 'Nothing pending' : 'Awaiting your review') + '</span>'
          + '</div>';
      }).catch(function () {});
    }
  }

  // ══════════════════════════════════════════════════════════════════════════
  // Company News
  // ══════════════════════════════════════════════════════════════════════════
  function setupNews() {
    var listEl = document.getElementById('news-list');
    if (!listEl) return;

    var BADGE_COLORS = { Company: '#0f1e42', Portal: '#2563eb' };

    var items = (PORTAL_CONFIG.news || []).slice();
    items.sort(function (a, b) {
      if (a.pinned && !b.pinned) return -1;
      if (!a.pinned && b.pinned) return  1;
      return (b.date || '').localeCompare(a.date || '');
    });

    if (!items.length) return;

    function cardHtml(item, featured) {
      var color   = BADGE_COLORS[item.category] || BADGE_COLORS.Portal;
      var dateStr = '';
      if (item.date) {
        var d = new Date(item.date + 'T00:00:00');
        dateStr = d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
      }
      var cls = 'portal-news-card'
        + (item.pinned ? ' portal-news-card--pinned'  : '')
        + (featured    ? ' portal-news-card--featured' : '');
      var meta = '<div class="portal-news-meta">'
        + (item.category ? '<span class="portal-news-badge" style="background:' + color + '">' + escapeHtml(item.category) + '</span>' : '')
        + (dateStr       ? '<span class="portal-news-date">' + dateStr + '</span>' : '')
        + (item.pinned   ? '<span class="portal-news-pin" title="Pinned">&#128204;</span>' : '')
        + '</div>';
      return '<div class="' + cls + '">'
        + '<div class="portal-news-card-head">' + meta
        + '<p class="portal-news-title">' + escapeHtml(item.title) + '</p>'
        + '</div>'
        + '<div class="portal-news-card-body">'
        + '<p class="portal-news-body">' + escapeHtml(item.body) + '</p>'
        + '</div>'
        + '</div>';
    }

    var html = cardHtml(items[0], true);
    if (items.length > 1) {
      html += '<div class="portal-news-grid">'
        + items.slice(1).map(function (item) { return cardHtml(item, false); }).join('')
        + '</div>';
    }
    listEl.innerHTML = html;
  }

  // ══════════════════════════════════════════════════════════════════════════
  // Company Directory — searchable employee list with Teams chat links
  // ══════════════════════════════════════════════════════════════════════════
  function setupDirectory(teamEmails) {
    var btn     = document.getElementById('btn-open-directory');
    var btn2    = document.getElementById('btn-open-team-directory');
    var panel   = document.getElementById('panel-directory');
    var search  = document.getElementById('dir-search');
    var grid    = document.getElementById('dir-grid');
    var count   = document.getElementById('dir-count');
    var panelLabel = panel ? panel.querySelector('.portal-ts-panel__label') : null;
    if (!panel || !grid) return;

    var fullRoster = [];
    var roster     = [];  // active roster (full or team-filtered)
    var rosterLoaded = false;
    var COLORS = ['#0f1e42','#1e3a5f','#1e40af','#065f46','#9f1239','#854d0e','#5b21b6','#48566f'];

    function openPanel(isTeamView) {
      if (!rosterLoaded) {
        fullRoster = (PORTAL_CONFIG.directory || []).slice();
        rosterLoaded = true;
      }
      if (isTeamView && teamEmails && teamEmails.length) {
        roster = fullRoster.filter(function (e) { return teamEmails.indexOf((e.email || '').toLowerCase()) !== -1; });
        if (panelLabel) panelLabel.textContent = 'Team Directory';
      } else {
        roster = fullRoster;
        if (panelLabel) panelLabel.textContent = 'Company Directory';
      }
      if (search) { search.value = ''; search.focus(); }
      renderGrid(roster);
      panel.hidden = false;
      if (btn)  btn.textContent  = 'Close';
      if (btn2) btn2.textContent = 'Close';
    }

    function closePanel() {
      panel.hidden = true;
      if (btn)  btn.textContent  = 'Browse';
      if (btn2) btn2.textContent = 'Browse';
    }

    if (btn) {
      btn.addEventListener('click', function () {
        panel.hidden ? openPanel(false) : closePanel();
      });
    }
    if (btn2) {
      btn2.addEventListener('click', function () {
        panel.hidden ? openPanel(true) : closePanel();
      });
    }

    if (search) {
      search.addEventListener('input', function () {
        var q = search.value.toLowerCase().trim();
        renderGrid(q ? roster.filter(function (e) {
          return e.name.toLowerCase().indexOf(q) !== -1 || (e.email || '').toLowerCase().indexOf(q) !== -1;
        }) : roster);
      });
    }

    function renderCard(emp) {
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
    }

    var COUNTRY_ORDER  = ['US', 'India', 'Pakistan'];
    var COUNTRY_LABELS = { US: 'United States', India: 'India', Pakistan: 'Pakistan' };

    function renderGrid(entries) {
      if (count) count.textContent = entries.length + ' of ' + roster.length + (roster.length === 1 ? ' person' : ' people');
      if (!entries.length) {
        grid.innerHTML = '<p class="portal-dir-empty">No results match your search.</p>';
        return;
      }

      // Group by country
      var groups = {};
      COUNTRY_ORDER.forEach(function (c) { groups[c] = []; });
      entries.forEach(function (emp) {
        var c = emp.country || 'US';
        if (!groups[c]) groups[c] = [];
        groups[c].push(emp);
      });

      var html = '';
      COUNTRY_ORDER.forEach(function (c) {
        var grp = groups[c];
        if (!grp || !grp.length) return;
        html += '<div class="portal-dir-country">'
          + '<div class="portal-dir-country-head portal-dir-country-head--' + c.toLowerCase() + '">'
          + escapeHtml(COUNTRY_LABELS[c])
          + ' <span class="portal-dir-country-count">(' + grp.length + ')</span>'
          + '</div>'
          + '<div class="portal-dir-country-grid">' + grp.map(renderCard).join('') + '</div>'
          + '</div>';
      });
      grid.innerHTML = html;
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
        bodyEl.innerHTML = '<p class="portal-ts-empty">No incident reports have been submitted yet.</p>';
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
