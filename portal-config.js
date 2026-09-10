/* ─────────────────────────────────────────────────────────────────────────────
   portal-config.js  —  Staff roster & SharePoint storage configuration
   See README.md "Timesheet Storage Setup" for step-by-step setup instructions.
───────────────────────────────────────────────────────────────────────────── */
var PORTAL_CONFIG = {

  // ── SharePoint document library ───────────────────────────────────────────
  // Instructions to find these IDs are in README.md "Timesheet Storage Setup".
  sharePointSiteId:  'YOUR-SHAREPOINT-SITE-ID',   // e.g. "americloudtelecom.sharepoint.com,abc123..."
  sharePointDriveId: 'YOUR-SHAREPOINT-DRIVE-ID',  // e.g. "b!abc123..."
  timesheetFolder:   'Timesheets',

  // ── Policies ──────────────────────────────────────────────────────────────
  // SharePoint folder where policy PDFs are stored (inside the same drive as timesheets)
  policyFolder:                'Policies',
  // SharePoint list tracking each policy issuance (PolicyName, PolicyUrl, IssuedTo JSON, IssuedDate)
  policyIssuancesListId:       'YOUR-POLICY-ISSUANCES-LIST-ID',
  // SharePoint list tracking per-employee signing status (PolicyItemId, EmployeeEmail, EmployeeName, Status, SignedDate)
  policyAcknowledgmentsListId: 'YOUR-POLICY-ACKNOWLEDGMENTS-LIST-ID',

  // ── PTO Requests ──────────────────────────────────────────────────────────
  // SharePoint list for employee PTO requests (see README "PTO Request Setup").
  // Columns: EmployeeEmail, EmployeeName, StartDate, EndDate, Reason, Status
  ptoRequestsListId: 'YOUR-PTO-REQUESTS-LIST-ID',

  // ── Incident Reports ──────────────────────────────────────────────────────
  // SharePoint list for safety & incident reports (see README "Incident Report Setup").
  // Columns: EmployeeEmail, EmployeeName, IncidentDate, Location, IncidentType, Description
  incidentReportsListId: 'YOUR-INCIDENT-REPORTS-LIST-ID',

  // ── Employee Handbook ─────────────────────────────────────────────────────
  // handbookUrl: direct link to the handbook PDF. Use a SharePoint URL (recommended)
  // or 'assets/docs/handbook.pdf' if you place the file in the repo.
  handbookUrl:           '#',  // replace with SharePoint PDF URL or local path

  // DocuSign PowerForm URL — created once in DocuSign admin (see README).
  // Portal appends ?Name={employee name}&Email={employee email} automatically.
  docuSignPowerFormUrl:  'YOUR-DOCUSIGN-POWERFORM-URL',

  // SharePoint list ID for handbook sign-off tracking (see README).
  // List must have columns: EmployeeEmail (text), Status (Signed/Pending), SignedDate (date).
  handbookSignoffListId: 'YOUR-HANDBOOK-SIGNOFF-LIST-ID',

  // ── Company Directory ─────────────────────────────────────────────────────
  // Sorted A–Z. Only names and emails — no PII. Update when headcount changes.
  directory: [
    { name: 'Abdul Fateh',               email: 'afateh@americloudtelecom.com'          },
    { name: 'Abelino Bautista',           email: 'abautista@americloudtelecom.com'       },
    { name: 'Afzal Karim',               email: 'akarim@americloudtelecom.com'           },
    { name: 'Ahmed Ghani',               email: 'aghani@americloudtelecom.com'           },
    { name: 'Ahmed Syed',                email: 'ahsyed@americloudtelecom.com'           },
    { name: 'Ajay Kumar Andapally',      email: 'ajkumar@americloudtelecom.com'          },
    { name: 'Alexis Pena',              email: 'apena@americloudtelecom.com'             },
    { name: 'Ali Ahmed Zaidi Syed',      email: 'szaidi@americloudtelecom.com'           },
    { name: 'Amaechi Nwankwo',           email: 'anwankwo@americloudtelecom.com'         },
    { name: 'Ammar Khan',                email: 'akhan@americloudtelecom.com'            },
    { name: 'Anmol Ahsan',               email: 'aahsan@americloudtelecom.com'           },
    { name: 'Arnell Avecilla',           email: 'aavecilla@americloudtelecom.com'        },
    { name: 'Arslan Arif',               email: 'marslan@americloudtelecom.com'          },
    { name: 'Arvin Mariano Fernandez',   email: 'amariano@americloudtelecom.com'         },
    { name: 'Ayesha Fouad',              email: 'afouad@americloudtelecom.com'           },
    { name: 'Ayesha Syed',              email: 'accounts@americloudtelecom.com'          },
    { name: 'Bakhtawar Jabeen',          email: 'bjabeen@americloudtelecom.com'          },
    { name: 'Chintan Kumar Patel',       email: 'cpatel@americloudtelecom.com'           },
    { name: 'Daniel Lapada',             email: 'dlapada@americloudtelecom.com'          },
    { name: 'Dinesh Pottunuri',          email: 'dpottunuri@americloudtelecom.com'       },
    { name: 'Esteban Fernandez',         email: 'efernandez@americloudtelecom.com'       },
    { name: 'Fahad Ghulam',              email: 'fghulam@americloudtelecom.com'          },
    { name: 'Ferdinand Nicasio',         email: 'fnicasio@americloudtelecom.com'         },
    { name: 'Gary Polloso',              email: 'gpolloso@americloudtelecom.com'         },
    { name: 'Hajira Bibi',               email: 'hbibi@americloudtelecom.com'            },
    { name: 'Hamza Shahzad',             email: 'hshahzad@americloudtelecom.com'         },
    { name: 'Humzah Shoaib',             email: 'hshoaib@americloudtelecom.com'          },
    { name: 'Isaiah Nicasio',            email: 'inicasio@americloudtelecom.com'         },
    { name: 'Izza Karim',                email: 'ikarim@americloudtelecom.com'           },
    { name: 'James Faux',                email: 'jfaux@americloudtelecom.com'            },
    { name: 'Jihad Mounir',              email: 'jmounir@americloudtelecom.com'          },
    { name: 'John Gladhill',             email: 'jgladhill@americloudtelecom.com'        },
    { name: 'Joselito Araneta',          email: 'jaraneta@americloudtelecom.com'         },
    { name: 'Kamran Khan',               email: 'kakhan@americloudtelecom.com'           },
    { name: 'Kaushik Gohel',             email: 'kgohel@americloudtelecom.com'           },
    { name: 'Maham Ashraf',              email: 'mashraf@americloudtelecom.com'          },
    { name: 'Maira Waheed',              email: 'mwaheed@americloudtelecom.com'          },
    { name: 'Manish Bommakanti',         email: 'mbommakanti@americloudtelecom.com'      },
    { name: 'Mateen Hussain',            email: 'mhussain@americloudtelecom.com'         },
    { name: 'Mati ul Haq',               email: 'mhaq@americloudtelecom.com'             },
    { name: 'Mohammed Faizan Khan',      email: 'fakhan@americloudtelecom.com'           },
    { name: 'Mohammed Furquan Ahmed',    email: 'mfahmed@americloudtelecom.com'          },
    { name: 'Mohammed Hussain',          email: 'mohammed.hussain@americloudtelecom.com' },
    { name: 'Mohammed Rayyan Ansari',    email: 'ransari@americloudtelecom.com'          },
    { name: 'Muhammad Bilal Sheykh',     email: 'msheykh@americloudtelecom.com'          },
    { name: 'Muhammad Zargham Baig',     email: 'mzbaig@americloudtelecom.com'           },
    { name: 'Mustafa Mohammed',          email: 'mmohammed@americloud.net'               },
    { name: 'Naaman Bakshi',             email: 'nbakhshi@americloudtelecom.com'         },
    { name: 'Nadeem Rasool',             email: 'nrasool@americloudtelecom.com'          },
    { name: 'Narender Bihan',            email: 'nbihan@americloudtelecom.com'           },
    { name: 'Nasrulla Ataulla',          email: 'nkhan@americloudtelecom.com'            },
    { name: 'Nishant Jawale',            email: 'njawale@americloudtelecom.com'          },
    { name: 'Nitin Nanagiri',            email: 'nnanagiri@americloudtelecom.com'        },
    { name: 'Orlando Medina',            email: 'omedina@americloudtelecom.com'          },
    { name: 'Peter Tijing',              email: 'ptijing@americloudtelecom.com'          },
    { name: 'Petterson Eden',            email: 'peden@americloudtelecom.com'            },
    { name: 'Pushmeena Saeed',           email: 'psaeed@americloudtelecom.com'           },
    { name: 'Qais Khadher',              email: 'qkhadher@americloudtelecom.com'         },
    { name: 'Qazi Saud Bin Saeed',       email: 'qsaud@americloudtelecom.com'            },
    { name: 'Reem Kirmani',              email: 'rkirmani@americloudtelecom.com'         },
    { name: 'Rosneidy Galindo',          email: 'rgalindo@americloudtelecom.com'         },
    { name: 'Sai Manikanta Kolla',       email: 'smkolla@americloudtelecom.com'          },
    { name: 'Saleem Waheed',             email: 'swaheed@americloudtelecom.com'          },
    { name: 'Sarmad Shahzad',            email: 'sashahzad@americloudtelecom.com'        },
    { name: 'Sathiamurthy Laxman',       email: 'lsathiamurthi@americloudtelecom.com'   },
    { name: 'Soban Yousaf',              email: 'syousaf@americloudtelecom.com'          },
    { name: 'Sri Vishnu Yadala',         email: 'svyadala@americloudtelecom.com'         },
    { name: 'Syed Muhammad Fateh Kirmani', email: 'fkirmani@americloudtelecom.com'      },
    { name: 'Syeda Maria Naseer',        email: 'mnaseer@americloudtelecom.com'          },
    { name: 'Tarab Naveed',              email: 'tnaveed@americloudtelecom.com'          },
    { name: 'Tayyab Ali',                email: 'tali@americloudtelecom.com'             },
    { name: 'Thirumalesh Naggari',       email: 'tnaggari@americloudtelecom.com'         },
    { name: 'Yamen Al Zeidi',            email: 'yamenkilani@americloudtelecom.com'      },
  ],

  // ── Staff roster ──────────────────────────────────────────────────────────
  // Keyed by M365 email (lowercase).
  // role:  'employee' | 'manager' | 'hr'
  // team:  managers only — array of employee emails this manager supervises.
  //
  // Add one entry per person. Remove the leading // from each example line.
  staff: {
    // 'john.doe@americloudtelecom.com':    { name: 'John Doe',    role: 'employee' },
    // 'jane.smith@americloudtelecom.com':  { name: 'Jane Smith',  role: 'manager',
    //                                        team: ['john.doe@americloudtelecom.com'] },
    // 'hr@americloudtelecom.com':          { name: 'HR Admin',    role: 'hr' },
  }

};
