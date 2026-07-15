# Casa de las Estrellas - Admissions final

This folder contains the public bilingual admissions form and the Google Apps Script backend for Casa de las Estrellas.

## Files

- `index.html`: public form page for GitHub Pages.
- `styles.css`: responsive visual design.
- `app.js`: bilingual form logic, validation, draft saving, file base64 encoding, and submission.
- `google-apps-script.gs`: Google Apps Script web app that creates the student folder, saves files in Drive, records Google Sheets rows, and emails admissions.
- `README.md`: setup instructions.

## Important security rule

Do not place private data, Google Drive folder IDs, Sheet IDs, credentials, or submitted family information in GitHub. The public files only contain the form. Private IDs are stored in Apps Script Properties.

## 1. Publish the public form with GitHub Pages

1. Upload only the five files in this folder to the GitHub repository or Pages branch.
2. In GitHub, open **Settings > Pages**.
3. Select the branch and folder that contain `index.html`.
4. Save and copy the published GitHub Pages URL.

## 2. Create Google Sheet

Upload `Historial_Estudiantes_Consolidado.xlsx` to Google Drive and open it as a Google Sheet. Use that converted spreadsheet as the admissions Sheet so the consolidated profiles, new internal codes and available HubSpot information are available to the form. The script creates these missing administrative tabs automatically:

1. `Base_Admisiones`
2. `Documentos`
3. `Salud`
4. `Cocina`
5. `Autorizaciones`
6. `Vacunas`
7. `Pendientes`
8. `Drive_Folders`
9. `Historial_Estudiantes`
10. `Log`

Copy the Sheet ID from the URL.

## 3. Create Google Drive root folder

Create one private root folder in Google Drive for student admission files. Copy the folder ID from the Drive URL.

Each new student creates a folder named like:

`Sofia Ramirez Gonzalez - RG06S17`

Existing students reuse their profile and the same Drive folder.

Inside it, Apps Script creates:

- `01_Identificación`
- `02_Documentos académicos`
- `03_Documentos médicos`
- `04_Vacunas`
- `05_Autorizaciones`
- `06_Contrato y matrícula`
- `07_Buseta`
- `08_Otros documentos`

## 4. Configure Apps Script

1. Open [script.google.com](https://script.google.com/).
2. Create a new project.
3. Paste the full content of `google-apps-script.gs`.
4. Open **Project Settings > Script Properties**.
5. Add these properties:

| Property | Value |
| --- | --- |
| `SPREADSHEET_ID` | Google Sheet ID |
| `DRIVE_ROOT_FOLDER_ID` | Drive admissions root folder ID |
| `ADMISSIONS_EMAIL` | Email or comma-separated emails for notifications |
| `TEST_NOTIFICATION_EMAILS` | Optional comma-separated emails for testing, for example Kelly and Dayana |
| `FINANCE_EMAIL` | Email or comma-separated emails for finance/CxC monthly payment requests |
| `MONTHLY_PAYMENT_EMAIL` | Optional alias for monthly payment request notifications |
| `SCHOOL_NAME` | Casa de las Estrellas |
| `DECLARATION_TEMPLATE_ID` | ID del Google Doc creado desde `Template_declaracion_final_CDE.docx` |

For the legal declaration template:

1. Upload `Template_declaracion_final_CDE.docx` to Google Drive.
2. Open it with Google Docs so Drive creates a native Google Doc copy.
3. Keep that master document unchanged and copy only its ID from the URL.
4. Save that ID as the `DECLARATION_TEMPLATE_ID` Script Property.

Apps Script creates a new declaration copy for each student, fills the template markers, inserts the captured signature, and stores the completed Google Doc and PDF in `05_Autorizaciones`.

Apps Script creates new Google Sheets columns automatically when needed. The latest version adds: `submissionId`, `student_id_expiration_date`, `guardian1_id_expiration_date`, `guardian2_id_expiration_date`, `kitchen_general_note`, `guardian1_id_file_url`, `guardian2_id_file_url`, `guardian2_id_not_applicable`, `vaccination_document_url`, `vaccination_document_not_applicable`, `vaccination_notes_or_declaration`, `paymentPlan`, `paymentMethod`, `monthlyPaymentReason`, `final_declaration_url`, `parent_confirmation_email_sent`, `internal_notification_email_sent`, and `finance_notification_email_sent`.

## 5. Deploy the Apps Script Web App

1. Click **Deploy > New deployment**.
2. Select **Web app**.
3. Execute as: **Me**.
4. Who has access: choose the setting appropriate for your school. For a public GitHub Pages form, use access that allows external submissions.
5. Deploy and authorize the required permissions.
6. Copy the Web App URL.

Required permissions include access to Google Sheets, Google Drive, sending email, and script properties.

## 6. Connect the form to Apps Script

Open `app.js` and replace:

```js
PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE
```

with the Web App URL from Apps Script.

## 7. Test

1. Open the GitHub Pages form.
2. Complete all required fields.
3. Attach small test files.
4. Submit.
5. Confirm that Apps Script returns a success message.
6. In Google Drive, check the created student folder and subfolders.
7. In Google Sheets, check the tabs and new rows.
8. Confirm that the admissions notification email arrived.

## Existing student profiles and reports

At the beginning of the first section, families enter the new student code and birth year. Apps Script verifies both values before returning existing information from `Base_Admisiones` or `Historial_Estudiantes`. The student's name, surnames and current grade appear first so the family can confirm the correct profile before continuing. All dates are selected with separate day, month and year dropdowns to avoid regional date-format ambiguity.

New student codes use this format: first letter of first surname + first letter of second surname, birth day, first letter of first name, and last two digits of birth year. Example: `RG06L17` for Luciana Rojas Gonzalez, born on 06/08/2017. When a student does not have a second surname in the source data, the temporary letter `X` is used and the row is marked for review.

The worksheet containing the profiles must be named exactly `Historial_Estudiantes`. Do not upload this private workbook to GitHub.

Each submission generates a PDF authorization summary for the family. When the student is not vaccinated, it also generates a signed no-vaccination declaration. The PDFs are stored in the student's folder and emailed to the parent email addresses entered in the form.

## Notes

- The form saves a local draft only in the user's browser. Drafts are not uploaded until submission.
- Files are encoded as base64 in the browser and sent to Apps Script.
- Large files can fail if they exceed Apps Script request limits. The frontend default limit is 12 MB per file and can be adjusted in `app.js`.
- Apps Script stores document links and the student Drive folder link in the Sheet.
