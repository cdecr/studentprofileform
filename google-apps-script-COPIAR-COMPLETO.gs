const SUBFOLDERS = [
  '01_Identificación',
  '02_Documentos académicos',
  '03_Documentos médicos',
  '04_Vacunas',
  '05_Autorizaciones',
  '06_Contrato y matrícula',
  '07_Buseta',
  '08_Otros documentos'
];

const TABS = ['Base_Admisiones','Documentos','Salud','Cocina','Autorizaciones','Vacunas','Pendientes','Drive_Folders','Dashboard_Admisiones','Historial_Estudiantes','Log'];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'lookupStudent') return jsonOutput(lookupStudent(payload.studentId, payload.verificationBirthYear || payload.verificationBirthDate));
    if (payload.action === 'syncDashboardData' || payload.action === 'getDashboardData') return jsonOutput(syncDashboardData());
    if (payload.action === 'sendFormLink') return jsonOutput(sendFormLink(payload.student_id || payload.studentId, payload.form_type || payload.formType, payload.email, payload.public_base_url || payload.publicBaseUrl));
    if (payload.action === 'sendFormReminder') return jsonOutput(sendFormReminder(payload.student_id || payload.studentId, payload.form_type || payload.formType, payload.email, payload.public_base_url || payload.publicBaseUrl));
    if (payload.action === 'updatePaymentStatus') return jsonOutput(updatePaymentStatus(payload.student_id || payload.studentId, payload.payment_type || payload.paymentType, payload.status, payload));
    return jsonOutput(handleSubmission(payload));
  } catch (error) {
    logError(error);
    return jsonOutput({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function doGet() {
  return jsonOutput({ ok: true, service: 'Casa de las Estrellas Admissions API' });
}

function handleSubmission(payload) {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = requireProp(props, 'SPREADSHEET_ID');
  const rootFolderId = requireProp(props, 'DRIVE_ROOT_FOLDER_ID');
  const admissionsEmail = buildInternalRecipients(props);
  const financeEmail = buildFinanceRecipients(props);
  const ss = SpreadsheetApp.openById(spreadsheetId);
  ensureTabs(ss);

  const data = payload.data || {};
  const submissionId = sanitizeSubmissionId(payload.submissionId || data.submissionId || '');
  data.submissionId = submissionId;
  const studentId = data.studentId || nextStudentId(ss, data.studentSince, data.birthDate);
  data.studentId = studentId;
  assertNoDuplicateProfile(ss.getSheetByName('Base_Admisiones'), data, studentId);
  const fullName = data.legalFullName || data.fullName || [data.firstName, data.secondName, data.lastName1, data.lastName2].filter(Boolean).join(' ') || 'Estudiante';
  data.fullName = fullName;
  applySubmissionAliases(data);

  const root = DriveApp.getFolderById(rootFolderId);
  const folderName = sanitizeFileName(fullName + ' - ' + studentId);
  const studentFolder = getOrCreateStudentFolder(ss, root, studentId, folderName);
  const subfolders = getOrCreateSubfolders(studentFolder);
  const savedFiles = saveFiles(payload.files || [], subfolders, submissionId);
  const reportFiles = createFamilyReports(data, studentId, subfolders, savedFiles);
  reportFiles.forEach(file => savedFiles.push(file));
  const finalDeclaration = reportFiles.filter(file => file.field === 'finalDeclaration')[0];
  data.final_declaration_url = finalDeclaration ? finalDeclaration.url : '';
  const pendingDocs = getPendingDocuments(data, savedFiles, ss, studentId);

  const baseRow = buildBaseRow(payload, data, studentFolder, savedFiles, pendingDocs);
  upsertObject(ss.getSheetByName('Base_Admisiones'), baseRow, 'studentId');
  appendDocuments(ss.getSheetByName('Documentos'), studentId, savedFiles);
  appendObject(ss.getSheetByName('Salud'), pickHealth(data));
  appendObject(ss.getSheetByName('Cocina'), pickDiet(data));
  appendObject(ss.getSheetByName('Autorizaciones'), pick(data, ['studentId','fullName','confirmTruth','confirmDataUse','confirmContact','imageChoice','completedBy','completedByRelation','submissionDate','digitalSignature','reviewConfirmed']));
  appendObject(ss.getSheetByName('Vacunas'), pickVaccines(data));
  appendPending(ss.getSheetByName('Pendientes'), studentId, pendingDocs);
  upsertObject(ss.getSheetByName('Drive_Folders'), { studentId, fullName, folderName: studentFolder.getName(), folderUrl: studentFolder.getUrl(), updatedAt: new Date() }, 'studentId');
  upsertObject(ss.getSheetByName('Dashboard_Admisiones'), { studentId, fullName, schoolPeriod: data.schoolPeriod, entryType: data.entryType, status: 'Solicitud recibida', missingDocuments: pendingDocs.join(', '), studentFolderUrl: studentFolder.getUrl(), contractSent: 'No', contractSigned: 'No', enrollmentPayment: 'Pendiente', enrollmentCompleted: 'No', internalNotes: '', followUp: '', updatedAt: new Date() }, 'studentId');
  configureAdmissionsDashboard(ss.getSheetByName('Dashboard_Admisiones'));
  appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: data.entryType === 'existing' ? 'profile_updated' : 'submission_created', studentId, fullName, files: savedFiles.length });

  try {
    sendNotification(admissionsEmail, data, studentFolder, savedFiles, pendingDocs, ss.getUrl(), reportFiles);
    data.internal_notification_email_sent = new Date();
  } catch (internalEmailError) {
    data.internal_notification_email_sent = 'Error: ' + String(internalEmailError);
    appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: 'internal_notification_email_error', studentId, submissionId, message: String(internalEmailError) });
  }
  try {
    sendFinanceNotificationIfNeeded(financeEmail, data, studentFolder, ss.getUrl());
    if (data.paymentPlan === 'monthly_request') data.finance_notification_email_sent = new Date();
  } catch (financeEmailError) {
    data.finance_notification_email_sent = 'Error: ' + String(financeEmailError);
    appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: 'finance_notification_email_error', studentId, submissionId, message: String(financeEmailError) });
  }
  try {
    sendFamilyReports(data, reportFiles, savedFiles);
    data.parent_confirmation_email_sent = new Date();
  } catch (familyEmailError) {
    appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: 'family_report_email_error', studentId, message: String(familyEmailError) });
  }
  upsertObject(ss.getSheetByName('Base_Admisiones'), pick(data, ['studentId','submissionId','final_declaration_url','parent_confirmation_email_sent','internal_notification_email_sent','finance_notification_email_sent']), 'studentId');
  return { ok: true, studentId, folderUrl: studentFolder.getUrl(), files: savedFiles.length, pending: pendingDocs, submissionId };
}

function lookupStudent(studentId, verificationBirthYear) {
  if (!studentId) return { ok: false, data: null };
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(requireProp(props, 'SPREADSHEET_ID'));
  const normalizedId = String(studentId).trim().toUpperCase();
  const birthYear = String(verificationBirthYear || '').match(/\d{4}/);
  let foundCode = false;
  const base = ss.getSheetByName('Base_Admisiones');
  const existing = findRowObject(base, 'studentId', normalizedId);
  if (existing) {
    foundCode = true;
    if (birthYearMatches(existing.birthDate, birthYear ? birthYear[0] : '')) {
      existing.studentLookupCode = normalizedId;
      existing.entryType = 'existing';
      existing.existingDocumentDetails = getExistingDocumentDetails(ss, normalizedId);
      existing.existingDocumentFields = uniqueValues(existing.existingDocumentDetails.map(item => item.field));
      return { ok: true, data: existing };
    }
  }

  for (const sheet of ss.getSheets()) {
    if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) continue;
    const range = sheet.getDataRange();
    const values = range.getDisplayValues();
    const rawValues = range.getValues();
    const headers = values[0].map(String);
    const idIndex = findHeaderIndex(headers, ['codigo_unico_nuevo','dashboard_key','Código nuevo','Nuevo código','Número interno estudiante','studentId']);
    if (idIndex < 0) continue;
    for (let r = 1; r < values.length; r++) {
      const currentCode = idIndex >= 0 ? String(values[r][idIndex]).trim().toUpperCase() : '';
      if (currentCode !== normalizedId) continue;
      foundCode = true;
      const row = values[r];
      const rawRow = rawValues[r];
      const firstName = legacyValue(headers, row, ['nombre','Nombre','Nombre(s)','firstName']);
      const explicitFirstSurname = legacyValue(headers, row, ['apellido1','Primer apellido','lastName1']);
      const explicitSecondSurname = legacyValue(headers, row, ['apellido2','Segundo apellido','lastName2']);
      const legacyFamily = explicitFirstSurname || legacyValue(headers, row, ['apellidos_original','Apellido','Apellidos','lastName']);
      const hasSplitSurnames = findHeaderIndex(headers, ['apellido1','Primer apellido','lastName1']) >= 0;
      const isConsolidatedHistory = findHeaderIndex(headers, ['Código interno']) >= 0;
      const lastNames = hasSplitSurnames
        ? { first: explicitFirstSurname, second: explicitSecondSurname }
        : (isConsolidatedHistory ? { first: legacyFamily, second: '' } : splitLegacyLastNames(legacyFamily));
      if (explicitSecondSurname) lastNames.second = explicitSecondSurname;
      const fullName = legacyValue(headers, row, ['nombre_completo','Nombre completo','fullName']) || [firstName, lastNames.first, lastNames.second].filter(Boolean).join(' ').trim();
      const history = findStudentHistory(ss, firstName, legacyFamily);
      const motherPhone = splitPhone(legacyValue(headers, row, ['tutor1_telefono','Tutor 1 teléfono','Phone Number','Teléfono','motherPhone']));
      const fatherPhone = splitPhone(legacyValue(headers, row, ['tutor2_telefono','Tutor 2 teléfono','2nd Phone Number','fatherPhone']));
      const idText = legacyValue(headers, row, ['identificacion_numerica','identificacion_original','ID Card Number / Passport','Número de identificación','Identificación estudiante','idNumber']);
      const birthIndex = findHeaderIndex(headers, ['fecha_nacimiento_dd/mm/yyyy','fecha_nacimiento','Fecha de nacimiento del estudiante','Fecha de nacimiento','birthDate']);
      const storedBirthDate = birthIndex >= 0 ? formatDateForForm(rawRow[birthIndex]) : '';
      if (!birthYearMatches(storedBirthDate, birthYear ? birthYear[0] : '')) continue;
      const grade = legacyValue(headers, row, ['clase_grado_2026_27','Grado actual','requestedGrade','currentGrade']) || sheet.getName();
      const statusText = legacyValue(headers, row, ['estado_estudiante','studentStatus']);
      const allergyOrDiet = legacyValue(headers, row, ['alergias_o_dieta','Alergias o dieta especial','otherAllergies']);
      const dietType = inferDietType(allergyOrDiet);
      const hasAllergies = allergyOrDiet && !/^(no|none|n\/a|sin restricciones)$/i.test(String(allergyOrDiet).trim()) ? 'yes' : 'no';
      const record = {
        studentLookupCode: normalizedId,
        studentId: currentCode || normalizedId,
        entryType: 'existing',
        admissionYear: currentSchoolYear(),
        schoolPeriod: legacyValue(headers, row, ['school_year','Año lectivo','schoolPeriod']) || currentSchoolYear(),
        currentGrade: grade,
        requestedGrade: grade,
        studentStatus: /inactivo/i.test(statusText) ? 'inactive' : 'active',
        studentSince: legacyValue(headers, row, ['Estudiante desde','Año de ingreso','studentSince']) || history.studentSince,
        startingGrade: legacyValue(headers, row, ['Grado de ingreso','startingGrade']) || history.startingGrade || '',
        firstName,
        lastName1: lastNames.first,
        lastName2: lastNames.second,
        fullName,
        birthDate: storedBirthDate,
        idType: inferIdType(idText),
        idNumber: cleanLegacyId(idText),
        citizenship: normalizeLegacyCountry(legacyValue(headers, row, ['nacionalidad_pais','PAIS','Pais','País','Nacionalidad estudiante','citizenship'])),
        gender: normalizeLegacyGender(legacyValue(headers, row, ['Género','Gender/Género','gender'])),
        firstLanguage: legacyValue(headers, row, ['idioma_principal','Idioma principal','firstLanguage']),
        homeLanguage: legacyValue(headers, row, ['idioma_principal','Idioma principal','homeLanguage']),
        hasAllergies,
        allergyDetail: allergyOrDiet,
        has_food_restrictions: hasAllergies,
        food_allergy_or_restriction: hasAllergies === 'yes' ? allergyOrDiet : '',
        foods_to_avoid: hasAllergies === 'yes' ? allergyOrDiet : '',
        dietType,
        dietDetail: allergyOrDiet,
        eatsEggs: normalizeYesNoNa(legacyValue(headers, row, ['vegetariano_consume_huevo','eatsEggs'])),
        otherAllergies: allergyOrDiet,
        currentAddressDetails: legacyValue(headers, row, ['Dirección en Nosara','Nosara home address | Dirección de casa en Nosara','currentAddressDetails']),
        motherName: legacyValue(headers, row, ['tutor1_nombre','Nombre del Padre o Tutor 1','Tutor 1 nombre','motherName']),
        motherEmail: legacyValue(headers, row, ['tutor1_email','Correo electrónico','Tutor 1 email','motherEmail']),
        motherPhoneCode: motherPhone.code,
        motherPhone: motherPhone.number,
        motherCitizenship: normalizeLegacyCountry(legacyValue(headers, row, ['Tutor 1 nacionalidad','motherNationality','motherCitizenship'])),
        motherId: cleanLegacyId(legacyValue(headers, row, ['Tutor 1 identificación','motherIdNumber','motherId'])),
        fatherName: legacyValue(headers, row, ['tutor2_nombre','Nombre del Padre o Tutor 2','Tutor 2 nombre','fatherName']),
        fatherEmail: legacyValue(headers, row, ['tutor2_email','2nd Email | 2do correo electrónico','Tutor 2 email','fatherEmail']),
        fatherPhoneCode: fatherPhone.code,
        fatherPhone: fatherPhone.number,
        fatherCitizenship: normalizeLegacyCountry(legacyValue(headers, row, ['Tutor 2 nacionalidad','fatherNationality','fatherCitizenship'])),
        fatherId: cleanLegacyId(legacyValue(headers, row, ['Tutor 2 identificación','fatherIdNumber','fatherId'])),
        additionalInfo: legacyValue(headers, row, ['notas','Notas','additionalInfo'])
      };
      const tutor1 = splitPersonName(record.motherName);
      const tutor2 = splitPersonName(record.fatherName);
      Object.assign(record, {
        'legalGuardians.0.firstName': tutor1.firstName,
        'legalGuardians.0.lastName1': tutor1.lastName1,
        'legalGuardians.0.lastName2': tutor1.lastName2,
        'legalGuardians.0.relationship': legacyValue(headers, row, ['tutor1_relacion','Tutor 1 relación']) || 'Madre',
        'legalGuardians.0.citizenship': record.motherCitizenship || '',
        'legalGuardians.0.idType': inferIdType(record.motherId),
        'legalGuardians.0.idNumber': record.motherId || '',
        'legalGuardians.0.phoneCode': record.motherPhoneCode || '+506',
        'legalGuardians.0.phone': record.motherPhone || '',
        'legalGuardians.0.email': record.motherEmail || '',
        'legalGuardians.0.sameAddress': record.currentAddressDetails ? 'yes' : '',
        'legalGuardians.0.livesWithStudent': normalizeYesNo(legacyValue(headers, row, ['tutor1_vive_con_estudiante','Tutor 1 vive con estudiante'])) || '',
        'legalGuardians.0.isEmergencyContact': record.motherPhone ? 'yes' : '',
        'legalGuardians.0.emergencyPriority': record.motherPhone ? 'Principal' : '',
        'legalGuardians.1.firstName': tutor2.firstName,
        'legalGuardians.1.lastName1': tutor2.lastName1,
        'legalGuardians.1.lastName2': tutor2.lastName2,
        'legalGuardians.1.relationship': legacyValue(headers, row, ['tutor2_relacion','Tutor 2 relación']) || (record.fatherName ? 'Padre' : ''),
        'legalGuardians.1.citizenship': record.fatherCitizenship || '',
        'legalGuardians.1.idType': inferIdType(record.fatherId),
        'legalGuardians.1.idNumber': record.fatherId || '',
        'legalGuardians.1.phoneCode': record.fatherPhoneCode || '+506',
        'legalGuardians.1.phone': record.fatherPhone || '',
        'legalGuardians.1.email': record.fatherEmail || '',
        'legalGuardians.1.sameAddress': record.currentAddressDetails && record.fatherName ? 'yes' : '',
        'legalGuardians.1.livesWithStudent': normalizeYesNo(legacyValue(headers, row, ['tutor2_vive_con_estudiante','Tutor 2 vive con estudiante'])) || '',
        'legalGuardians.1.isEmergencyContact': record.fatherPhone ? 'yes' : '',
        'legalGuardians.1.emergencyPriority': record.fatherPhone ? 'Secundario' : ''
      });
      record.existingDocumentDetails = getExistingDocumentDetails(ss, record.studentId);
      record.existingDocumentFields = uniqueValues(record.existingDocumentDetails.map(item => item.field));
      return { ok: true, data: record };
    }
  }
  return { ok: false, data: null, reason: foundCode ? 'verification_failed' : 'not_found' };
}

function syncDashboardData() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(requireProp(props, 'SPREADSHEET_ID'));
  ensureTabs(ss);
  const base = sheetObjects(ss.getSheetByName('Base_Admisiones'));
  const dashboard = sheetObjects(ss.getSheetByName('Dashboard_Admisiones'));
  const documents = sheetObjects(ss.getSheetByName('Documentos'));
  const pending = sheetObjects(ss.getSheetByName('Pendientes'));
  const driveFolders = sheetObjects(ss.getSheetByName('Drive_Folders'));
  const health = latestByStudent(sheetObjects(ss.getSheetByName('Salud')));
  const kitchen = latestByStudent(sheetObjects(ss.getSheetByName('Cocina')));
  const byId = {};
  base.forEach(row => {
    const id = dashboardStudentId(row);
    if (id) byId[id] = Object.assign(byId[id] || {}, row);
  });
  dashboard.forEach(row => {
    const id = dashboardStudentId(row);
    if (id) byId[id] = Object.assign(byId[id] || {}, row);
  });
  driveFolders.forEach(row => {
    const id = dashboardStudentId(row);
    if (id) byId[id] = Object.assign(byId[id] || {}, row);
  });
  const documentMap = groupByStudent(documents);
  const pendingMap = groupByStudent(pending);
  const records = Object.keys(byId).sort().map(id => dashboardRecordFromRows(id, byId[id], documentMap[id] || [], pendingMap[id] || [], health[id] || {}, kitchen[id] || {}));
  return { ok: true, records, generatedAt: new Date() };
}

function sendFormLink(studentId, formType, email, publicBaseUrl) {
  return sendDashboardFormEmail('link', studentId, formType, email, publicBaseUrl);
}

function sendFormReminder(studentId, formType, email, publicBaseUrl) {
  return sendDashboardFormEmail('reminder', studentId, formType, email, publicBaseUrl);
}

function sendDashboardFormEmail(kind, studentId, formType, email, publicBaseUrl) {
  if (!studentId || !formType) return { ok: false, error: 'Missing student_id or form_type' };
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(requireProp(props, 'SPREADSHEET_ID'));
  ensureTabs(ss);
  const record = findRowObject(ss.getSheetByName('Base_Admisiones'), 'studentId', studentId) || findRowObject(ss.getSheetByName('Dashboard_Admisiones'), 'studentId', studentId) || {};
  const to = email || record['legalGuardians.0.email'] || record.guardian_email || record.email || props.getProperty('ADMISSIONS_EMAIL') || Session.getActiveUser().getEmail();
  if (!to) return { ok: false, error: 'Missing recipient email' };
  const formUrl = buildPublicFormUrl(formType, publicBaseUrl, props);
  const studentName = record.fullName || record.student_name || record.studentName || studentId;
  const subject = kind === 'reminder' ? `Recordatorio de formulario - Casa de las Estrellas` : `Formulario Casa de las Estrellas`;
  const body = [
    'Hola,',
    '',
    kind === 'reminder' ? 'Le recordamos completar el siguiente formulario:' : 'Por favor complete el siguiente formulario:',
    '',
    formUrl,
    '',
    `Estudiante: ${studentName}`,
    `Código: ${studentId}`,
    '',
    'Gracias,',
    'Casa de las Estrellas'
  ].join('\n');
  MailApp.sendEmail({ to, subject, body });
  const update = kind === 'reminder'
    ? { studentId, last_reminder_sent_at: new Date(), next_action: 'Recordatorio enviado' }
    : { studentId, last_form_link_sent: formType, last_form_link_sent_at: new Date(), last_form_link_url: formUrl, next_action: 'Formulario enviado' };
  upsertObject(ss.getSheetByName('Dashboard_Admisiones'), update, 'studentId');
  appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: kind === 'reminder' ? 'dashboard_form_reminder_sent' : 'dashboard_form_link_sent', studentId, formType, emailTo: to });
  return { ok: true, url: formUrl, email_to: to };
}

function updatePaymentStatus(studentId, paymentType, status, data) {
  if (!studentId) return { ok: false, error: 'Missing student_id' };
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(requireProp(props, 'SPREADSHEET_ID'));
  ensureTabs(ss);
  const update = {
    studentId,
    enrollmentPayment: status || 'Actualizado',
    paymentType: paymentType || '',
    paymentUpdatedAt: new Date(),
    paymentNotes: data && data.notes ? data.notes : ''
  };
  upsertObject(ss.getSheetByName('Dashboard_Admisiones'), update, 'studentId');
  appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: 'payment_status_updated', studentId, paymentType, status });
  return { ok: true };
}

function buildPublicFormUrl(formType, publicBaseUrl, props) {
  const configured = props.getProperty('FORM_PUBLIC_URL') || props.getProperty('STUDENT_FORM_URL') || 'https://cdecr.github.io/studentprofileform/';
  if (formType === 'application' || formType === 'reenrollment' || formType === 'documents' || formType === 'legal') return configured;
  const base = publicBaseUrl || configured;
  if (/studentprofileform\/?$/i.test(base)) return configured;
  return String(base).replace(/admin\.html.*$/,'') + 'forms.html?form=' + encodeURIComponent(formType || 'inquiry');
}

function legacyValue(headers, row, names) {
  const index = findHeaderIndex(headers, names);
  return index >= 0 ? String(row[index] || '').trim() : '';
}

function splitLegacyLastNames(value) {
  const parts = String(value || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length < 2) return { first: parts[0] || '', second: 'No registrado' };
  return { first: parts.slice(0, -1).join(' '), second: parts[parts.length - 1] };
}

function splitPersonName(value) {
  const parts = String(value || '').trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return { firstName: '', lastName1: '', lastName2: '' };
  if (parts.length === 1) return { firstName: parts[0], lastName1: '', lastName2: '' };
  if (parts.length === 2) return { firstName: parts[0], lastName1: parts[1], lastName2: '' };
  return { firstName: parts.slice(0, -2).join(' '), lastName1: parts[parts.length - 2], lastName2: parts[parts.length - 1] };
}

function formatDateForForm(value) {
  if (!value) return '';
  const iso = String(value).match(/^\d{4}-\d{2}-\d{2}/);
  if (iso) return iso[0];
  const slashDate = String(value).trim().match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (slashDate) {
    const day = slashDate[1].padStart(2, '0');
    const month = slashDate[2].padStart(2, '0');
    const year = slashDate[3];
    return `${year}-${month}-${day}`;
  }
  const numericSerial = typeof value === 'number' ? value : (/^\d{5}(\.\d+)?$/.test(String(value).trim()) ? Number(value) : 0);
  if (numericSerial > 20000) {
    const date = new Date(Date.UTC(1899, 11, 30) + numericSerial * 24 * 60 * 60 * 1000);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function datesMatch(stored, provided) {
  const storedKey = formatDateForForm(stored);
  const providedKey = formatDateForForm(provided);
  return Boolean(storedKey && providedKey && storedKey === providedKey);
}

function birthYearMatches(stored, providedYear) {
  const storedKey = formatDateForForm(stored);
  const storedYear = storedKey ? storedKey.slice(0, 4) : String(stored || '').match(/\d{4}/);
  return Boolean(storedYear && providedYear && String(storedYear) === String(providedYear));
}

function cleanLegacyId(value) {
  return String(value || '').replace(/\D/g, '');
}

function inferIdType(value) {
  const text = String(value || '').toLowerCase();
  if (text.indexOf('pasaporte') >= 0 || text.indexOf('passport') >= 0) return 'Pasaporte';
  if (text.indexOf('dimex') >= 0 || text.indexOf('residencia') >= 0) return 'DIMEX';
  if (text.indexOf('cedula') >= 0 || text.indexOf('cédula') >= 0) return 'Cédula';
  return value ? 'Otro' : '';
}

function normalizeLegacyCountry(value) {
  const text = String(value || '').trim();
  const key = text.toLowerCase().replace(/\./g, '').replace(/\s+/g, ' ');
  const map = {
    'costa rican': 'Costa Rica', 'costa rica': 'Costa Rica', 'costarricense': 'Costa Rica', 'costaricense': 'Costa Rica', 'costarricanse': 'Costa Rica',
    'ee uu': 'United States', 'eeuu': 'United States', 'usa': 'United States', 'united states': 'United States', 'american': 'United States', 'estadounidense': 'United States',
    'canadian': 'Canada', 'canada': 'Canada', 'german': 'Germany', 'alemana': 'Germany', 'france': 'France',
    'argentina': 'Argentina', 'argentino': 'Argentina', 'polish': 'Poland', 'poland': 'Poland',
    'nicaragua': 'Nicaragua', 'israel': 'Israel', 'columbian': 'Colombia', 'colombian': 'Colombia'
  };
  return map[key] || text;
}

function normalizeLegacyGender(value) {
  const key = normalizeName(value);
  if (key === 'girl nina' || key === 'girl' || key === 'nina' || key === 'female' || key === 'femenino') return 'Femenino';
  if (key === 'boy nino' || key === 'boy' || key === 'nino' || key === 'male' || key === 'masculino') return 'Masculino';
  return value ? 'Sin especificar' : '';
}

function inferDietType(value) {
  const text = normalizeName(value);
  if (!text || text === 'none' || text === 'no' || text === 'n a') return 'Sin restricciones';
  if (text.indexOf('vegan') >= 0 || text.indexOf('vegana') >= 0) return 'Vegana';
  if (text.indexOf('vegetarian') >= 0 || text.indexOf('vegetariana') >= 0) return 'Vegetariana';
  if (text.indexOf('gluten') >= 0) return 'Sin gluten';
  if (text.indexOf('lacteo') >= 0 || text.indexOf('dairy') >= 0 || text.indexOf('milk') >= 0) return 'Sin lácteos';
  return 'Otra';
}

function normalizeYesNoNa(value) {
  const text = normalizeName(value);
  if (!text) return '';
  if (['si','yes','true','x'].indexOf(text) >= 0) return 'yes';
  if (['no','false'].indexOf(text) >= 0) return 'no';
  return 'na';
}

function normalizeYesNo(value) {
  const text = normalizeName(value);
  if (!text) return '';
  if (['si','yes','true','x'].indexOf(text) >= 0) return 'yes';
  if (['no','false'].indexOf(text) >= 0) return 'no';
  return '';
}

function splitPhone(value) {
  const digits = String(value || '').replace(/\D/g, '');
  const codes = ['972','506','505','507','49','44','34','57','55','54','52','1'];
  for (const code of codes) {
    if (digits.indexOf(code) === 0 && digits.length > code.length + 6) return { code: '+' + code, number: digits.slice(code.length) };
  }
  return { code: '+506', number: digits };
}

function findHeaderIndex(headers, names) {
  const normalized = headers.map(h => String(h).trim().toLowerCase());
  for (const name of names) {
    const index = normalized.indexOf(String(name).trim().toLowerCase());
    if (index >= 0) return index;
  }
  return -1;
}

function findStudentHistory(ss, studentName, familyName) {
  const sheet = ss.getSheetByName('Historial_Estudiantes');
  if (!sheet || sheet.getLastRow() < 2) return {};
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(String);
  const yearIndex = findHeaderIndex(headers, ['Ano comienzo curso','Año comienzo curso','Año de ingreso','studentSince']);
  const gradeIndex = findHeaderIndex(headers, ['Grado en el que comenzo','Grado en el que comenzó','Grado de ingreso','startingGrade']);
  const familyIndex = findHeaderIndex(headers, ['Family','Apellido','Apellidos','Familia']);
  const nameIndex = findHeaderIndex(headers, ["Student´s name","Student's name",'Nombre','Nombre(s)','Student name']);
  if ([yearIndex, familyIndex, nameIndex].some(index => index < 0)) return {};
  const wantedName = normalizeName(studentName);
  const wantedFamily = normalizeName(familyName);
  for (let r = 1; r < values.length; r++) {
    const rowName = normalizeName(values[r][nameIndex]);
    const rowFamily = normalizeName(values[r][familyIndex]);
    const sameName = rowName === wantedName || rowName.indexOf(wantedName) >= 0 || wantedName.indexOf(rowName) >= 0;
    const sameFamily = rowFamily === wantedFamily || rowFamily.indexOf(wantedFamily) >= 0 || wantedFamily.indexOf(rowFamily) >= 0;
    if (sameName && sameFamily) return { studentSince: values[r][yearIndex], startingGrade: gradeIndex >= 0 ? values[r][gradeIndex] : '' };
  }
  return {};
}

function normalizeName(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function surnamesMatch(stored, provided) {
  const storedName = normalizeName(stored);
  const providedName = normalizeName(provided);
  if (!storedName || !providedName) return false;
  const storedParts = storedName.split(' ');
  const providedParts = providedName.split(' ');
  return providedParts.every(part => storedParts.indexOf(part) >= 0) || storedName.indexOf(providedName) >= 0;
}

function findRowObject(sheet, key, value) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const index = headers.indexOf(key);
  if (index < 0) return null;
  for (let r = values.length - 1; r >= 1; r--) {
    if (String(values[r][index]).trim().toUpperCase() === String(value).trim().toUpperCase()) {
      const obj = {};
      headers.forEach((header, i) => obj[header] = values[r][i]);
      return obj;
    }
  }
  return null;
}

function sheetObjects(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) obj[header] = row[index];
    });
    return obj;
  });
}

function latestByStudent(rows) {
  const map = {};
  rows.forEach(row => {
    const id = dashboardStudentId(row);
    if (id) map[id] = Object.assign(map[id] || {}, row);
  });
  return map;
}

function groupByStudent(rows) {
  const map = {};
  rows.forEach(row => {
    const id = dashboardStudentId(row);
    if (!id) return;
    if (!map[id]) map[id] = [];
    map[id].push(row);
  });
  return map;
}

function dashboardStudentId(row) {
  return String(row.studentId || row.student_id || row.codigo_unico_nuevo || row.dashboard_key || '').trim().toUpperCase();
}

function dashboardRecordFromRows(studentId, row, documents, pending, health, kitchen) {
  const guardian = primaryGuardianSummary(row);
  const fullName = row.fullName || row.legalFullName || [row.firstName, row.lastName1, row.lastName2].filter(Boolean).join(' ') || row.student_name || '';
  const pendingText = pending.map(item => item.pendingDocument || item.status || '').filter(Boolean).join(', ') || row.missingDocuments || row.documentsPending || '';
  const documentText = documents.map(item => item.name || item.file_name || item.field || '').filter(Boolean).join(', ');
  const status = normalizeDashboardStatus(row.status || row.current_status || row.entryType || 'En proceso');
  return Object.assign({}, row, {
    student_id: studentId,
    codigo_unico_nuevo: studentId,
    dashboard_key: studentId + '_' + (row.schoolPeriod || row.school_year || ''),
    student_name: fullName,
    student_names: fullName,
    student_initials: initialsForName(fullName),
    grade: row.currentGrade || row.grade || row.lastGrade || '',
    current_status: status,
    record_type: row.entryType === 'new' ? 'Interesado' : 'Estudiante',
    guardian_name: guardian.name || row.completedBy || '',
    guardian_email: guardian.email || '',
    guardian_phone: guardian.phone || '',
    photo_preview_url: firstDocumentUrl(documents, ['studentPhoto']),
    folder_url: row.folderUrl || row.studentFolderUrl || '',
    studentFolderUrl: row.studentFolderUrl || row.folderUrl || '',
    documents_summary: documentText,
    pending_documents: pendingText,
    needs_review: pendingText,
    alergias_o_dieta: kitchen.food_allergy_or_restriction || kitchen.foods_to_avoid || health.allergyDetail || '',
    idioma_principal: row.firstLanguage || row.idioma_principal || '',
    nacionalidad_pais: row.citizenship || row.nacionalidad_pais || '',
    identificacion_numerica: row.idNumber || row.identificacion_numerica || '',
    fecha_nacimiento_dd_mm_yyyy: formatDashboardDate(row.birthDate),
    'fecha_nacimiento_dd/mm/yyyy': formatDashboardDate(row.birthDate),
    dia: datePart(row.birthDate, 'day'),
    mes: datePart(row.birthDate, 'month'),
    yyyy: datePart(row.birthDate, 'year'),
    last_form_link_sent: row.last_form_link_sent || '',
    last_form_link_sent_at: row.last_form_link_sent_at || '',
    last_reminder_sent_at: row.last_reminder_sent_at || '',
    next_action: row.next_action || row.followUp || ''
  });
}

function normalizeDashboardStatus(value) {
  const text = String(value || '').toLowerCase();
  if (/matr[ií]cula completada|matriculado|confirmado|aprobado/.test(text)) return 'Matriculado';
  if (/pendiente|documento|confirmar/.test(text)) return 'Pendiente de confirmar';
  if (/inactivo|retirado/.test(text)) return 'Inactivo';
  return 'En proceso';
}

function initialsForName(name) {
  return String(name || '').split(/\s+/).filter(Boolean).slice(0, 2).map(part => part[0]).join('').toUpperCase();
}

function firstDocumentUrl(documents, fields) {
  const found = (documents || []).find(item => fields.indexOf(item.field) >= 0);
  return found ? (found.url || found.file_url || '') : '';
}

function formatDashboardDate(value) {
  const iso = normalizeDateValue(value);
  if (!iso) return '';
  const parts = iso.split('-');
  return [parts[2], parts[1], parts[0]].join('/');
}

function datePart(value, part) {
  const iso = normalizeDateValue(value);
  if (!iso) return '';
  const parts = iso.split('-');
  if (part === 'day') return parts[2];
  if (part === 'month') return parts[1];
  return parts[0];
}

function ensureTabs(ss) {
  TABS.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
}

function configureAdmissionsDashboard(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const statusColumn = headers.indexOf('status') + 1;
  if (statusColumn) {
    const statuses = ['Solicitud recibida','En revisión','Pendiente de documentos','Entrevista pendiente','Entrevista programada','Aprobado','Lista de espera','No admitido','Contrato enviado','Contrato firmado','Pago de matrícula pendiente','Pago de matrícula recibido','Matrícula completada'];
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(false).build();
    sheet.getRange(2, statusColumn, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#247f91').setFontColor('#ffffff');
  if (!sheet.getFilter() && sheet.getLastRow() >= 1) sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), sheet.getLastColumn()).createFilter();
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function getOrCreateSubfolders(studentFolder) {
  const map = {};
  SUBFOLDERS.forEach(name => {
    const existing = studentFolder.getFoldersByName(name);
    map[name] = existing.hasNext() ? existing.next() : studentFolder.createFolder(name);
  });
  return map;
}

function getOrCreateStudentFolder(ss, root, studentId, folderName) {
  const record = findRowObject(ss.getSheetByName('Drive_Folders'), 'studentId', studentId);
  if (record && record.folderUrl) {
    const match = String(record.folderUrl).match(/[-\w]{20,}/);
    if (match) {
      try { return DriveApp.getFolderById(match[0]); } catch (ignored) {}
    }
  }
  const folders = root.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().indexOf(studentId) >= 0) return folder;
  }
  return root.createFolder(folderName);
}

function saveFiles(files, subfolders, submissionId) {
  return files.map(file => {
    const category = SUBFOLDERS.indexOf(file.category) >= 0 ? file.category : '08_Otros documentos';
    const prefix = submissionId ? submissionId.slice(-10) + ' - ' : '';
    const fileName = sanitizeFileName(prefix + (file.name || 'archivo'));
    const existingFile = findFileByName(subfolders[category], fileName);
    if (existingFile) {
      return {
        field: file.field,
        category,
        name: existingFile.getName(),
        mimeType: file.mimeType,
        size: existingFile.getSize(),
        url: existingFile.getUrl(),
        id: existingFile.getId(),
        reused: true
      };
    }
    const bytes = Utilities.base64Decode(file.base64);
    const blob = Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', fileName);
    const driveFile = subfolders[category].createFile(blob);
    return {
      field: file.field,
      category,
      name: driveFile.getName(),
      mimeType: file.mimeType,
      size: file.size,
      url: driveFile.getUrl(),
      id: driveFile.getId()
    };
  });
}

function findFileByName(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function buildBaseRow(payload, data, studentFolder, savedFiles, pendingDocs) {
  const row = Object.assign({}, data);
  savedFiles.forEach(function(file) {
    if (file.field === 'guardian1IdFile') row.guardian1_id_file_url = file.url || '';
    if (file.field === 'guardian2IdFile') row.guardian2_id_file_url = file.url || '';
    if (file.field === 'vaccineDocs') row.vaccination_document_url = file.url || '';
    if (file.field === 'finalDeclaration') row.final_declaration_url = file.url || row.final_declaration_url || '';
  });
  row.submittedAt = payload.submittedAt || new Date();
  row.submissionId = data.submissionId || payload.submissionId || '';
  row.language = payload.language || '';
  row.studentFolderUrl = studentFolder.getUrl();
  row.documentsReceived = savedFiles.map(f => f.field + ': ' + f.name).join('\n');
  row.documentsPending = pendingDocs.join('\n');
  return row;
}

function applySubmissionAliases(data) {
  data.student_id_expiration_date = data.student_id_expiration_date || '';
  data.guardian1_id_expiration_date = data['legalGuardians.0.idExpirationDate'] || '';
  data.guardian2_id_expiration_date = data['legalGuardians.1.idExpirationDate'] || '';
  data.guardian2_id_not_applicable = data.guardian2IdNotApplicable || '';
  data.vaccination_document_not_applicable = data.vaccineDocsNA || '';
  data.vaccination_notes_or_declaration = data.vaccineComments || '';
  data.parent_confirmation_email_sent = data.parent_confirmation_email_sent || '';
  data.internal_notification_email_sent = data.internal_notification_email_sent || '';
}

function appendDocuments(sheet, studentId, files) {
  files.forEach(file => {
    if (!documentRowExists(sheet, studentId, file)) appendObject(sheet, Object.assign({ studentId, uploadedAt: new Date() }, file));
  });
}

function documentRowExists(sheet, studentId, file) {
  if (!sheet || sheet.getLastRow() < 2) return false;
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0];
  const studentIndex = headers.indexOf('studentId');
  const fieldIndex = headers.indexOf('field');
  const idIndex = headers.indexOf('id');
  const nameIndex = headers.indexOf('name');
  if (studentIndex < 0) return false;
  return values.slice(1).some(row => {
    const sameStudent = String(row[studentIndex] || '').trim().toUpperCase() === String(studentId || '').trim().toUpperCase();
    const sameId = idIndex >= 0 && file.id && String(row[idIndex] || '') === String(file.id);
    const sameFieldName = fieldIndex >= 0 && nameIndex >= 0 && String(row[fieldIndex] || '') === String(file.field || '') && String(row[nameIndex] || '') === String(file.name || '');
    return sameStudent && (sameId || sameFieldName);
  });
}

function appendPending(sheet, studentId, pendingDocs) {
  if (!pendingDocs.length) appendObject(sheet, { studentId, status: 'Sin pendientes', createdAt: new Date() });
  pendingDocs.forEach(item => appendObject(sheet, { studentId, pendingDocument: item, status: 'Pendiente', createdAt: new Date() }));
}

function appendObject(sheet, obj) {
  const existingHeaders = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(String) : [];
  const keys = Object.keys(obj);
  const headers = existingHeaders.slice();
  keys.forEach(k => { if (headers.indexOf(k) === -1) headers.push(k); });
  if (headers.length !== existingHeaders.length) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const row = headers.map(h => normalizeCell(obj[h]));
  sheet.appendRow(row);
}

function upsertObject(sheet, obj, key) {
  const existingHeaders = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(String) : [];
  const headers = existingHeaders.slice();
  Object.keys(obj).forEach(name => { if (headers.indexOf(name) === -1) headers.push(name); });
  if (headers.length !== existingHeaders.length) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const keyIndex = headers.indexOf(key);
  let targetRow = -1;
  if (keyIndex >= 0 && sheet.getLastRow() >= 2) {
    const keys = sheet.getRange(2, keyIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (let i = keys.length - 1; i >= 0; i--) {
      if (String(keys[i][0]).trim().toUpperCase() === String(obj[key]).trim().toUpperCase()) { targetRow = i + 2; break; }
    }
  }
  const oldRow = targetRow > 0 ? sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0] : [];
  const row = headers.map((header, index) => Object.prototype.hasOwnProperty.call(obj, header) ? normalizeCell(obj[header]) : (oldRow[index] || ''));
  if (targetRow > 0) sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
  else sheet.appendRow(row);
}

function normalizeCell(value) {
  if (value === null || value === undefined) return '';
  if (Array.isArray(value)) return value.join('\n');
  if (typeof value === 'object' && !(value instanceof Date)) return JSON.stringify(value);
  return value;
}

function assertNoDuplicateProfile(sheet, data, studentId) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(String);
  const idIndex = headers.indexOf('studentId');
  const checks = [
    { key: 'idNumber', value: data.idNumber, label: 'student identification' },
    { key: 'motherEmail', value: data.motherEmail, label: 'primary guardian email' },
    { key: 'fatherEmail', value: data.fatherEmail, label: 'secondary guardian email' }
  ].filter(check => String(check.value || '').trim());
  for (let r = 1; r < values.length; r++) {
    if (idIndex >= 0 && String(values[r][idIndex]).trim().toUpperCase() === String(studentId).trim().toUpperCase()) continue;
    for (const check of checks) {
      const index = headers.indexOf(check.key);
      if (index >= 0 && normalizeName(values[r][index]) === normalizeName(check.value)) throw new Error('Duplicate ' + check.label + ' found in another student profile.');
    }
  }
}

function nextStudentId(ss, studentSince, birthDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const year = String(studentSince || '').match(/\d{4}/);
    const date = formatDateForForm(birthDate);
    if (!year || !date) throw new Error('Student start year and birth date are required to generate the internal code.');
    const prefix = 'CDE' + year[0].slice(-2) + date.slice(5, 7);
    const pattern = new RegExp('^' + prefix + '(\\d{2})$');
    let max = 0;
    ss.getSheets().forEach(sheet => {
      if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
      sheet.getDataRange().getDisplayValues().forEach(row => {
        row.forEach(value => {
          const match = String(value || '').trim().toUpperCase().match(pattern);
          if (match) max = Math.max(max, Number(match[1]));
        });
      });
    });
    if (max >= 99) throw new Error('The internal sequence for ' + prefix + ' has reached 99.');
    return prefix + String(max + 1).padStart(2, '0');
  } finally {
    lock.releaseLock();
  }
}

function currentSchoolYear() {
  const start = new Date().getFullYear();
  return `${start}-${start + 1}`;
}

function getPendingDocuments(data, files, ss, studentId) {
  const present = {};
  files.forEach(f => present[f.field] = true);
  getExistingDocumentFields(ss, studentId).forEach(field => present[field] = true);
  const pending = [];
  if (!present.studentIdFile) pending.push('Identificación o pasaporte del estudiante');
  if (!present.studentPhoto) pending.push('Foto reciente del estudiante');
  if (!present.guardian1IdFile && !present.guardianIdFiles) pending.push('Identificación del tutor legal 1');
  if (hasSecondGuardian(data) && !present.guardian2IdFile && !present.guardianIdFiles && data.guardian2IdNotApplicable !== 'yes') pending.push('Identificación del tutor legal 2');
  if (!present.birthCertificate) pending.push('Certificado de nacimiento');
  if (!present.vaccineDocs && data.vaccineDocsNA !== 'yes') pending.push('Carné de vacunas o consentimiento/declaración firmada');
  return pending;
}

function hasSecondGuardian(data) {
  return ['firstName','lastName1','idNumber','email','phone'].some(function(key) {
    return String(data['legalGuardians.1.' + key] || '').trim();
  });
}

function getExistingDocumentFields(ss, studentId) {
  return uniqueValues(getExistingDocumentDetails(ss, studentId).map(item => item.field));
}

function getExistingDocumentDetails(ss, studentId) {
  const details = [];
  const sheet = ss && ss.getSheetByName('Documentos');
  const id = String(studentId || '').trim().toUpperCase();
  if (sheet && sheet.getLastRow() >= 2) {
    const values = sheet.getDataRange().getDisplayValues();
    const headers = values[0];
    const studentIndex = headers.indexOf('studentId');
    const fieldIndex = headers.indexOf('field');
    const nameIndex = headers.indexOf('name');
    const urlIndex = headers.indexOf('url');
    if (studentIndex >= 0 && fieldIndex >= 0) values.slice(1).forEach(row => {
      if (String(row[studentIndex]).trim().toUpperCase() === id && row[fieldIndex]) details.push({ field: String(row[fieldIndex]).trim(), name: nameIndex >= 0 ? row[nameIndex] : '', url: urlIndex >= 0 ? row[urlIndex] : '' });
    });
  }
  const folder = findStudentFolder(ss, id);
  if (folder) scanDocumentFolder(folder, '', details);
  const seen = {};
  return details.filter(item => item.field && !seen[item.field + '|' + item.name] && (seen[item.field + '|' + item.name] = true));
}

function findStudentFolder(ss, studentId) {
  const record = findRowObject(ss.getSheetByName('Drive_Folders'), 'studentId', studentId);
  if (record && record.folderUrl) {
    const match = String(record.folderUrl).match(/[-\w]{20,}/);
    if (match) try { return DriveApp.getFolderById(match[0]); } catch (ignored) {}
  }
  const rootId = PropertiesService.getScriptProperties().getProperty('DRIVE_ROOT_FOLDER_ID');
  if (!rootId) return null;
  const folders = DriveApp.getFolderById(rootId).getFolders();
  while (folders.hasNext()) { const folder = folders.next(); if (folder.getName().toUpperCase().indexOf(studentId) >= 0) return folder; }
  return null;
}

function scanDocumentFolder(folder, category, details) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const field = inferDocumentField(file.getName(), category || folder.getName());
    if (field) details.push({ field, name: file.getName(), url: file.getUrl() });
  }
  const folders = folder.getFolders();
  while (folders.hasNext()) { const child = folders.next(); scanDocumentFolder(child, child.getName(), details); }
}

function inferDocumentField(fileName, category) {
  const text = normalizeName(fileName + ' ' + category);
  if (/firma digital|declaracion final|autorizacion/.test(text)) return '';
  if (/vacuna|inmunizacion|carnet/.test(text)) return 'vaccineDocs';
  if (/nacimiento|birth/.test(text)) return 'birthCertificate';
  if (/foto|photo|retrato/.test(text)) return 'studentPhoto';
  if (/tutor 1|guardian 1|madre|mama/.test(text) && /cedula|identificacion|pasaporte|id/.test(text)) return 'guardian1IdFile';
  if (/tutor 2|guardian 2|padre/.test(text) && /cedula|identificacion|pasaporte|id/.test(text)) return 'guardian2IdFile';
  if (/mama|madre|padre|tutor|guardian/.test(text) && /cedula|identificacion|pasaporte|id/.test(text)) return 'guardianIdFiles';
  if (/migracion|dimex|visa|residencia/.test(text)) return 'migrationDocs';
  if (/medic|psicolog|terapeut|salud/.test(text) || /03 documentos medicos/.test(text)) return 'medicalDocs';
  if (/reporte|nota|academ|escuela|school/.test(text) || /02 documentos academicos/.test(text)) return 'schoolDocumentFiles';
  if (/cedula|identificacion|pasaporte|student id/.test(text)) return 'studentIdFile';
  return '';
}

function uniqueValues(values) {
  return values.filter(Boolean).filter((value, index, list) => list.indexOf(value) === index);
}

function createFamilyReports(data, studentId, subfolders, savedFiles) {
  const templateId = PropertiesService.getScriptProperties().getProperty('DECLARATION_TEMPLATE_ID');
  if (!templateId) throw new Error('Missing Script Property: DECLARATION_TEMPLATE_ID');
  const signature = savedFiles.find(file => file.field === 'digitalSignatureFile');
  const suffix = data.submissionId ? ' - ' + data.submissionId.slice(-10) : '';
  return [
    createFinalDeclaration(templateId, data, studentId, subfolders['05_Autorizaciones'], signature, savedFiles),
    createPdfReport(`Resumen del formulario - ${data.fullName || studentId}${suffix}`, buildSubmissionPreviewLines(data), subfolders['05_Autorizaciones'], signature)
  ];
}

function createFinalDeclaration(templateId, data, studentId, folder, signature, savedFiles) {
  const suffix = data.submissionId ? ' - ' + data.submissionId.slice(-10) : '';
  const title = `Declaración final - ${data.fullName || studentId}${suffix}`;
  const pdfName = sanitizeFileName(title) + '.pdf';
  const existingPdf = findFileByName(folder, pdfName);
  if (existingPdf) return { field: 'finalDeclaration', category: folder.getName(), name: existingPdf.getName(), mimeType: MimeType.PDF, size: existingPdf.getSize(), url: existingPdf.getUrl(), id: existingPdf.getId(), reused: true };
  const copy = DriveApp.getFileById(templateId).makeCopy(title, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const relation = String(data.completedByRelation || 'Representante legal');
  const guardianIndex = findCompletedByGuardianIndex(data);
  const adultId = data['legalGuardians.' + guardianIndex + '.idNumber'] || data.motherId || data.fatherId || '';
  const adultEmail = data['legalGuardians.' + guardianIndex + '.email'] || data.motherEmail || data.fatherEmail || '';
  const adultPhone = [data['legalGuardians.' + guardianIndex + '.phoneCode'], data['legalGuardians.' + guardianIndex + '.phone']].filter(Boolean).join(' ') || data.motherPhone || data.fatherPhone || '';
  const adultAddress = data['legalGuardians.' + guardianIndex + '.addressDetails'] || data.currentAddressDetails || '';
  const hasVaccineDocument = savedFiles.some(file => file.field === 'vaccineDocs');
  insertSignatureAtMarker(body, signature);
  const values = {
    version: '1.0',
    fecha_emision: formatDeclarationDate(new Date()),
    adulto_nombre: data.completedBy || '',
    adulto_identificacion: adultId || '',
    adulto_relacion: relation,
    adulto_correo: adultEmail || '',
    adulto_telefono: adultPhone || '',
    adulto_domicilio: adultAddress || '',
    estudiante_nombre_completo: data.fullName || '',
    estudiante_codigo: studentId,
    estudiante_identificacion: data.idNumber || '',
    imagen_autoriza_total: data.imageChoice === 'internal_external' ? '[X]' : '[ ]',
    imagen_autoriza_interno: '[ ]',
    imagen_no_autoriza: data.imageChoice === 'no' ? '[X]' : '[ ]',
    vacunas_estado: declarationVaccineStatus(data.vaccinesUpToDate),
    vacunas_documento: hasVaccineDocument ? 'Sí, adjunto al expediente' : 'Pendiente / no adjunto',
    razon_no_vacunacion: data.vaccineComments || 'No aplica',
    grado_nivel: data.currentGrade || data.lastGrade || '',
    lugar_firma: 'Costa Rica',
    fecha_firma: formatDeclarationDate(data.submissionDate || new Date()),
    firma_electronica_simple: data.completedBy || ''
  };
  Object.keys(values).forEach(key => body.replaceText('\\{\\{' + key + '\\}\\}', String(values[key] || '')));
  doc.saveAndClose();
  const pdf = folder.createFile(copy.getAs(MimeType.PDF).setName(pdfName));
  return { field: 'finalDeclaration', category: folder.getName(), name: pdf.getName(), mimeType: MimeType.PDF, size: pdf.getSize(), url: pdf.getUrl(), id: pdf.getId() };
}

function findCompletedByGuardianIndex(data) {
  const signer = normalizeName(data.completedBy || '');
  for (let i = 0; i < 6; i++) {
    const name = normalizeName([data['legalGuardians.' + i + '.firstName'], data['legalGuardians.' + i + '.lastName1'], data['legalGuardians.' + i + '.lastName2']].filter(Boolean).join(' '));
    if (signer && name && (signer === name || name.indexOf(signer) >= 0 || signer.indexOf(name) >= 0)) return i;
  }
  return 0;
}

function insertSignatureAtMarker(body, signature) {
  const found = body.findText('\\{\\{firma_electronica_simple\\}\\}');
  if (!found || !signature) return;
  const text = found.getElement().asText();
  text.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
  const paragraph = text.getParent().asParagraph();
  const image = paragraph.appendInlineImage(DriveApp.getFileById(signature.id).getBlob());
  image.setWidth(160);
}

function declarationVaccineStatus(value) {
  if (value === 'yes') return 'Vacunas al día';
  if (value === 'no') return 'Vacunas no al día';
  return 'Pendiente de confirmar';
}

function formatDeclarationDate(value) {
  const date = value instanceof Date ? value : new Date(String(value) + 'T12:00:00');
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function createPdfReport(title, paragraphs, folder, signature) {
  const pdfName = sanitizeFileName(title) + '.pdf';
  const existingPdf = findFileByName(folder, pdfName);
  if (existingPdf) return { field: 'generatedReport', category: folder.getName(), name: existingPdf.getName(), mimeType: MimeType.PDF, size: existingPdf.getSize(), url: existingPdf.getUrl(), id: existingPdf.getId(), reused: true };
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  body.appendParagraph('Casa de las Estrellas').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  paragraphs.forEach(text => body.appendParagraph(text));
  if (signature && signature.id) {
    body.appendParagraph('Firma electrónica simple:');
    const image = body.appendImage(DriveApp.getFileById(signature.id).getBlob());
    image.setWidth(220);
  }
  doc.saveAndClose();
  const source = DriveApp.getFileById(doc.getId());
  const pdf = folder.createFile(source.getAs(MimeType.PDF).setName(pdfName));
  source.setTrashed(true);
  return { field: 'generatedReport', category: folder.getName(), name: pdf.getName(), mimeType: MimeType.PDF, size: pdf.getSize(), url: pdf.getUrl(), id: pdf.getId() };
}

function sendFamilyReports(data, reports, allFiles) {
  const recipients = getFamilyRecipientEmails(data);
  if (!recipients.length || !reports.length) return;
  const attachments = reports.map(report => DriveApp.getFileById(report.id).getBlob());
  const adultName = data.completedBy || primaryGuardianSummary(data).name || 'familia';
  const studentName = data.legalFullName || data.fullName || 'el/la estudiante';
  MailApp.sendEmail({
    to: recipients.join(','),
    subject: `Confirmación de recepción del formulario — Casa de las Estrellas`,
    body: [
      `Estimado/a ${adultName},`,
      '',
      `Hemos recibido correctamente el formulario de ${studentName}.`,
      '',
      'El equipo de Casa de las Estrellas revisará la información y los documentos enviados. Si se requiere información adicional, nos comunicaremos con usted.',
      data.paymentPlan === 'monthly_request' ? 'Hemos registrado su solicitud de pagos mensuales. El departamento de cuentas por cobrar revisará el caso y enviará la información o el formulario correspondiente.' : '',
      '',
      'Adjuntamos o incluimos a continuación un resumen de la información y autorizaciones registradas para su respaldo.',
      '',
      buildDeclarationSummary(data),
      '',
      'Documentos recibidos:',
      (allFiles || reports).length ? (allFiles || reports).map(report => `- ${report.name}: ${report.url}`).join('\n') : '- Sin documentos registrados',
      '',
      buildNextSteps()
    ].join('\n'),
    attachments
  });
}

function getFamilyRecipientEmails(data) {
  const recipients = [];
  Object.keys(data).forEach(key => {
    if (/^legalGuardians\.\d+\.email$/.test(key) && data[key]) recipients.push(data[key]);
  });
  [data.motherEmail, data.fatherEmail].forEach(value => { if (value) recipients.push(value); });
  return uniqueValues(recipients.map(value => String(value || '').trim().toLowerCase()).filter(Boolean));
}

function buildInternalRecipients(props) {
  const values = [
    props.getProperty('ADMISSIONS_EMAIL') || Session.getActiveUser().getEmail(),
    props.getProperty('TEST_NOTIFICATION_EMAILS') || ''
  ];
  return uniqueValues(values.join(',').split(/[,\n;]+/).map(value => String(value || '').trim().toLowerCase()).filter(Boolean)).join(',');
}

function buildFinanceRecipients(props) {
  const values = [
    props.getProperty('FINANCE_EMAIL') || '',
    props.getProperty('MONTHLY_PAYMENT_EMAIL') || ''
  ];
  return uniqueValues(values.join(',').split(/[,\n;]+/).map(value => String(value || '').trim().toLowerCase()).filter(Boolean)).join(',');
}

function buildDeclarationSummary(data) {
  const answer = value => value === 'yes' ? 'Sí' : value === 'no' ? 'No' : 'No registrado';
  return [
    'Declaraciones registradas:',
    `- Información verdadera y completa: ${answer(data.confirmTruth)}`,
    `- Tratamiento de datos autorizado: ${answer(data.confirmDataUse)}`,
    `- Contacto autorizado: ${answer(data.confirmContact)}`,
    `- Uso de imagen: ${formatImageChoice(data.imageChoice)}`,
    `- Vacunas: ${data.vaccineDocsNA === 'yes' ? 'No aplica / requiere justificación registrada' : declarationVaccineStatus(data.vaccinesUpToDate)}`,
    `- Persona firmante: ${data.completedBy || 'No registrado'} (${data.completedByRelation || 'relación no registrada'})`,
    `- Fecha: ${formatDeclarationDate(data.submissionDate || new Date())}`
  ].join('\n');
}

function formatImageChoice(value) {
  if (value === 'internal_external') return 'Autorizado para fines institucionales, educativos y de comunicación';
  if (value === 'no') return 'No autorizado';
  return 'No registrado';
}

function buildSubmissionPreviewLines(data) {
  var skip = /signature|base64|File$/i;
  var keys = Object.keys(data || {}).filter(function(key) {
    return data[key] !== '' && data[key] !== null && data[key] !== undefined && !skip.test(key);
  }).sort();
  var rows = keys.map(function(key) {
    return formatFieldName(key) + ': ' + formatPreviewValue(data[key]);
  });
  var lines = [
    'Resumen de la información registrada en el formulario.',
    'Estudiante: ' + (data.fullName || [data.firstName, data.lastName1, data.lastName2].filter(Boolean).join(' ') || 'No registrado'),
    'Código: ' + (data.studentId || 'No registrado'),
    ''
  ];
  return lines.concat(rows);
}

function formatPreviewValue(value) {
  if (Array.isArray(value)) return value.join(', ');
  if (value === 'yes') return 'Sí';
  if (value === 'no') return 'No';
  if (value === 'internal_external' || value === 'internal_only') return formatImageChoice(value);
  return String(value);
}

function formatFieldName(key) {
  return String(key).replace(/\.\d+\./g, ' ').replace(/[._]/g, ' ').replace(/\s+/g, ' ').trim();
}

function buildDocumentSectionLinks(files, folder) {
  var byCategory = {};
  (files || []).forEach(function(file) {
    var category = file.category || 'Expediente';
    if (!byCategory[category]) byCategory[category] = [];
    byCategory[category].push(file);
  });
  var lines = ['Expediente digital por sección:', '- Carpeta principal: ' + folder.getUrl()];
  Object.keys(byCategory).sort().forEach(function(category) {
    lines.push('- ' + category + ':');
    byCategory[category].forEach(function(file) {
      lines.push('  - ' + file.name + ': ' + (file.url || ''));
    });
  });
  return lines.join('\n');
}

function sendFinanceNotificationIfNeeded(to, data, folder, sheetUrl) {
  if (data.paymentPlan !== 'monthly_request' || !to) return;
  const guardian = primaryGuardianSummary(data);
  const subject = `Solicitud de pagos mensuales: ${data.fullName || data.studentId || 'Estudiante'}`;
  const body = [
    'Se recibió una solicitud de pagos mensuales desde el formulario de admisión.',
    '',
    `Estudiante: ${data.fullName || ''}`,
    `ID estudiante: ${data.studentId || ''}`,
    `Período escolar: ${data.schoolPeriod || ''}`,
    `Grado actual o solicitado: ${data.currentGrade || ''}`,
    `Tutor responsable: ${guardian.name}`,
    `Correo del tutor: ${guardian.email}`,
    `Teléfono: ${guardian.phone}`,
    `Método de pago preferido: ${formatPaymentMethod(data.paymentMethod)}`,
    `Comentario para CxC: ${data.monthlyPaymentReason || 'Sin comentario adicional'}`,
    '',
    `Carpeta Drive: ${folder.getUrl()}`,
    `Google Sheet: ${sheetUrl}`,
    '',
    'Acción sugerida: contactar a la familia y enviar la información o formulario correspondiente para pagos mensuales.'
  ].join('\n');
  MailApp.sendEmail({ to, subject, body });
}

function formatPaymentMethod(value) {
  const map = { bank_transfer: 'Transferencia bancaria', card: 'Tarjeta', cash: 'Efectivo', other: 'Otro' };
  return map[value] || value || 'No registrado';
}

function sendNotification(to, data, folder, files, pendingDocs, sheetUrl, reports) {
  const subject = `Nueva admisión: ${data.legalFullName || data.fullName || data.firstName || 'Estudiante'} (${data.studentId})`;
  const guardian = primaryGuardianSummary(data);
  const body = [
    'Se recibió un nuevo formulario de admisión.',
    '',
    `Estudiante: ${data.legalFullName || data.fullName || ''}`,
    `ID estudiante: ${data.studentId || ''}`,
    `Período escolar: ${data.schoolPeriod || ''}`,
    `Grado actual o solicitado: ${data.currentGrade || ''}`,
    `Tutor responsable: ${guardian.name}`,
    `Correo del tutor: ${guardian.email}`,
    `Teléfono: ${guardian.phone}`,
    `Autorización de imagen: ${formatImageChoice(data.imageChoice)}`,
    `Estado de vacunas: ${data.vaccineDocsNA === 'yes' ? 'No aplica / requiere justificación registrada' : declarationVaccineStatus(data.vaccinesUpToDate)}`,
    `Carpeta Drive: ${folder.getUrl()}`,
    `Google Sheet: ${sheetUrl}`,
    '',
    buildSubmissionPreviewLines(data).join('\n'),
    '',
    buildDocumentSectionLinks(files, folder),
    '',
    'Documentos recibidos:',
    files.length ? files.map(f => `- ${f.name}: ${f.url}`).join('\n') : '- Ninguno',
    '',
    'Documentos pendientes:',
    pendingDocs.length ? pendingDocs.map(d => `- ${d}`).join('\n') : '- Sin pendientes',
    '',
    buildDeclarationSummary(data),
    '',
    buildNextSteps()
  ].join('\n');
  const attachments = (reports || []).map(report => DriveApp.getFileById(report.id).getBlob());
  MailApp.sendEmail({ to, subject, body, attachments });
}

function primaryGuardianSummary(data) {
  const name = [data['legalGuardians.0.firstName'], data['legalGuardians.0.lastName1'], data['legalGuardians.0.lastName2']].filter(Boolean).join(' ') || data.motherName || data.fatherName || data.completedBy || '';
  const email = data['legalGuardians.0.email'] || data.motherEmail || data.fatherEmail || '';
  const phone = [data['legalGuardians.0.phoneCode'], data['legalGuardians.0.phone']].filter(Boolean).join(' ') || data.motherPhone || data.fatherPhone || '';
  return { name, email, phone };
}

function pick(data, keys) {
  const obj = {};
  keys.forEach(k => obj[k] = data[k] || '');
  return obj;
}

function pickHealth(data) {
  const obj = pick(data, ['studentId','fullName','vaccinesUpToDate','vaccineComments','hasMedicalCondition','medicalConditionDetail','hasAllergies','allergyDetail','takesMedication','medicationName','medicationReason','medicationNotes','hasPhysicalRestriction','physicalRestrictionDetail','hasLearningCondition','learningConditionDetail','hospitalized','hospitalizationReasons','healthNotes']);
  Object.keys(data).forEach(k => { if (k.indexOf('health_') === 0) obj[k] = data[k]; });
  return obj;
}

function pickDiet(data) {
  return pick(data, [
    'studentId',
    'fullName',
    'has_food_restrictions',
    'food_allergy_or_restriction',
    'foods_to_avoid',
    'dairy_alternative_notes',
    'meat_alternative_notes',
    'religious_cultural_food_restrictions',
    'kitchen_safety_instructions',
    'kitchen_general_note'
  ]);
}

function buildNextSteps() {
  return [
    'Próximos pasos:',
    '- El equipo de Casa de las Estrellas revisará la información y los documentos adjuntos.',
    '- Si falta algún documento o aclaración, nos comunicaremos con la familia.',
    '- Una vez revisado el expediente, admisiones indicará los siguientes pasos del proceso.',
    '- El envío del formulario no garantiza admisión ni reserva de cupo hasta completar la revisión correspondiente.'
  ].join('\n');
}

function pickVaccines(data) {
  return pick(data, ['studentId','fullName','vaccinesUpToDate','vaccineComments','vaccineDocsNA','vaccination_document_not_applicable','vaccination_notes_or_declaration','submissionDate','completedBy']);
}

function sanitizeFileName(name) {
  return String(name).replace(/[\\/:*?"<>|#%{}~&]/g, '-').replace(/\s+/g, ' ').trim().slice(0, 180);
}

function sanitizeSubmissionId(value) {
  return String(value || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 80);
}

function requireProp(props, name) {
  const value = props.getProperty(name);
  if (!value) throw new Error(`Missing Script Property: ${name}`);
  return value;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function logError(error) {
  try {
    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = props.getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) return;
    const ss = SpreadsheetApp.openById(spreadsheetId);
    ensureTabs(ss);
    appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: 'error', message: String(error && error.stack ? error.stack : error) });
  } catch (ignored) {}
}
