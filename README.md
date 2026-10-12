# বরিশাল ক্যাডেট কলেজ মূল্যায়ন ফর্ম

The form in `index.html` sends college evaluation submissions to the Apps Script web app whose `/exec` URL is configured in `COLLEGE_GOOGLE_SHEETS_WEB_APP_URL`. The Apps Script source in `Code.gs` writes to the spreadsheet ID configured in `COLLEGE_SPREADSHEET_ID`.

## Configure Google Sheets saving

1. Open the Apps Script project that owns the web app URL in `index.html`.
2. Replace its script source with the current contents of `Code.gs` and save. The local file is not automatically synchronized with Apps Script.
3. In the Apps Script editor, run `setupCollegeEvaluationSheet` once and approve spreadsheet access. Check the execution log for the spreadsheet and worksheet details.
4. Open **Deploy → Manage deployments**, edit the web app deployment, select **New version**, and deploy it. The web app must execute as an account with access to the spreadsheet, and its access setting must allow the form's respondents to submit.
5. Confirm that `COLLEGE_GOOGLE_SHEETS_WEB_APP_URL` in `index.html` contains the `/exec` URL for this deployment. If deploying a different Apps Script project, update the URL and publish the updated form.

Do not treat a successful response from an older deployment as proof that the current `Code.gs` is running. The current handler validates required class, house, form, ratings, and goal/objective response (`হ্যাঁ`, `না`, or `নিশ্চিত নই`) before it writes a row. A description is required only for `হ্যাঁ`. Rating comments are required for the three lowest ratings; the final management comment is optional.

The `College Evaluation` worksheet is created or migrated by `setupCollegeEvaluationSheet`. Existing rating columns and responses are retained, including legacy columns no longer shown in the current form. New submissions add the current rating values, their comments, the goal/objective response, and the optional management comment on one row.

The public web app URL is not authentication. Anyone who obtains it may be able to submit data as the deploying account. Share the form and endpoint carefully and monitor the spreadsheet.
