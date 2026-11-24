LEARNER GATE PASS - SETUP

1. Create a Google Sheet. Open Extensions > Apps Script and paste Code.gs.
2. Change ADMIN_PASSWORD at the top, then run setup() once and approve permissions.
3. Project Settings: set time zone to Asia/Manila.
4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
   Copy the /exec URL (use the one WITHOUT /a/macros/<domain>).
5. Open gatepass-sheets.html (double-click to test, or host on HTTPS such as Netlify Drop
   or GitHub Pages for camera use). Paste the /exec URL in the Backend URL box.
6. Sign in as superadmin with your ADMIN_PASSWORD.

Optional SMS: Project Settings > Script properties > add SMS_API_KEY (and SMS_SENDER).
To update code later: Deploy > Manage deployments > pencil > New version (keeps the same URL).
