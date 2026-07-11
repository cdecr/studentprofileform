const SUBFOLDERS = [
  '01_Identificación',
  '02_Documentos académicos',
  '03_Documentos médicos',
  '04_Vacunas',
  '05_Autorizaciones',
  '06_Contrato y matrícula',
  '08_Otros documentos'
];

const TABS = ['Base_Admisiones','Documentos','Salud','Cocina','Autorizaciones','Vacunas','Pendientes','Drive_Folders','Dashboard_Admisiones','Historial_Estudiantes','Log'];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action === 'lookupStudent') return jsonOutput(lookupStudent(payload.studentId, payload.verificationBirthYear || payload.verificationBirthDate));
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
  const admissionsEmail = props.getProperty('ADMISSIONS_EMAIL') || Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(spreadsheetId);
  ensureTabs(ss);

  const data = payload.data || {};
  const studentId = data.studentId || nextStudentId(ss, data.studentSince, data.birthDate);
  data.studentId = studentId;
  assertNoDuplicateProfile(ss.getSheetByName('Base_Admisiones'), data, studentId);
  const fullName = data.legalFullName || data.fullName || [data.firstName, data.secondName, data.lastName1, data.lastName2].filter(Boolean).join(' ') || 'Estudiante';
  data.fullName = fullName;

  const root = DriveApp.getFolderById(rootFolderId);
  const folderName = sanitizeFileName(fullName + ' - ' + studentId);
  const studentFolder = getOrCreateStudentFolder(ss, root, studentId, folderName);
  const subfolders = getOrCreateSubfolders(studentFolder);
  const savedFiles = saveFiles(payload.files || [], subfolders);
  const reportFiles = createFamilyReports(data, studentId, subfolders, savedFiles);
  reportFiles.forEach(file => savedFiles.push(file));
  const pendingDocs = getPendingDocuments(data, savedFiles, ss, studentId);

  const baseRow = buildBaseRow(payload, data, studentFolder, savedFiles, pendingDocs);
  upsertObject(ss.getSheetByName('Base_Admisiones'), baseRow, 'studentId');
  appendDocuments(ss.getSheetByName('Documentos'), studentId, savedFiles);
  appendObject(ss.getSheetByName('Salud'), pick(data, ['studentId','fullName','vaccinesUpToDate','vaccineComments','hasMedicalCondition','medicalConditionDetail','hasAllergies','allergyDetail','takesMedication','medicationName','medicationDose','medicationFrequency','medicationNotes','hasPhysicalRestriction','physicalRestrictionDetail','hasLearningCondition','learningConditionDetail','healthNotes']));
  appendObject(ss.getSheetByName('Cocina'), pickDiet(data));
  appendObject(ss.getSheetByName('Autorizaciones'), pick(data, ['studentId','fullName','confirmTruth','confirmDataUse','confirmContact','imageChoice','confirmRead','confirmNoGuarantee','confirmSubmit','completedBy','completedByRelation','submissionDate','digitalSignature','reviewConfirmed']));
  appendObject(ss.getSheetByName('Vacunas'), pickVaccines(data));
  appendPending(ss.getSheetByName('Pendientes'), studentId, pendingDocs);
  upsertObject(ss.getSheetByName('Drive_Folders'), { studentId, fullName, folderName: studentFolder.getName(), folderUrl: studentFolder.getUrl(), updatedAt: new Date() }, 'studentId');
  upsertObject(ss.getSheetByName('Dashboard_Admisiones'), { studentId, fullName, schoolPeriod: data.schoolPeriod, entryType: data.entryType, status: 'Solicitud recibida', missingDocuments: pendingDocs.join(', '), studentFolderUrl: studentFolder.getUrl(), contractSent: 'No', contractSigned: 'No', enrollmentPayment: 'Pendiente', enrollmentCompleted: 'No', internalNotes: '', followUp: '', updatedAt: new Date() }, 'studentId');
  configureAdmissionsDashboard(ss.getSheetByName('Dashboard_Admisiones'));
  appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: data.entryType === 'existing' ? 'profile_updated' : 'submission_created', studentId, fullName, files: savedFiles.length });

  sendNotification(admissionsEmail, data, studentFolder, savedFiles, pendingDocs, ss.getUrl(), reportFiles);
  try {
    sendFamilyReports(data, reportFiles);
  } catch (familyEmailError) {
    appendObject(ss.getSheetByName('Log'), { timestamp: new Date(), event: 'family_report_email_error', studentId, message: String(familyEmailError) });
  }
  return { ok: true, studentId, folderUrl: studentFolder.getUrl(), files: savedFiles.length, pending: pendingDocs };
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
      record.existingDocumentDetails = getExistingDocumentDetails(ss, record.studentId);
      record.existingDocumentFields = uniqueValues(record.existingDocumentDetails.map(item => item.field));
      return { ok: true, data: record };
    }
  }
  return { ok: false, data: null, reason: foundCode ? 'verification_failed' : 'not_found' };
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

function formatDateForForm(value) {
  if (!value) return '';
  const iso = String(value).match(/^\d{4}-\d{2}-\d{2}/);
  if (iso) return iso[0];
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

function saveFiles(files, subfolders) {
  return files.map(file => {
    const category = SUBFOLDERS.indexOf(file.category) >= 0 ? file.category : '08_Otros documentos';
    const bytes = Utilities.base64Decode(file.base64);
    const blob = Utilities.newBlob(bytes, file.mimeType || 'application/octet-stream', sanitizeFileName(file.name || 'archivo'));
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

function buildBaseRow(payload, data, studentFolder, savedFiles, pendingDocs) {
  const row = Object.assign({}, data);
  row.submittedAt = payload.submittedAt || new Date();
  row.language = payload.language || '';
  row.studentFolderUrl = studentFolder.getUrl();
  row.documentsReceived = savedFiles.map(f => f.field + ': ' + f.name).join('\n');
  row.documentsPending = pendingDocs.join('\n');
  return row;
}

function appendDocuments(sheet, studentId, files) {
  files.forEach(file => appendObject(sheet, Object.assign({ studentId, uploadedAt: new Date() }, file)));
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
  const required = {
    studentIdFile: 'Identificación o pasaporte del estudiante',
    studentPhoto: 'Foto reciente del estudiante',
    guardianIdFiles: 'Identificación de madre, padre o tutor legal',
    birthCertificate: 'Certificado de nacimiento'
  };
  return Object.keys(required).filter(field => !present[field] && data[field] !== 'yes').map(field => required[field]);
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
  return [createFinalDeclaration(templateId, data, studentId, subfolders['05_Autorizaciones'], signature, savedFiles)];
}

function createFinalDeclaration(templateId, data, studentId, folder, signature, savedFiles) {
  const title = `Declaración final - ${data.fullName || studentId}`;
  const copy = DriveApp.getFileById(templateId).makeCopy(title, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  const relation = String(data.completedByRelation || 'Representante legal');
  const useFather = /padre/i.test(relation);
  const adultId = useFather ? data.fatherId : data.motherId;
  const adultEmail = useFather ? data.fatherEmail : data.motherEmail;
  const adultPhone = useFather ? data.fatherPhone : data.motherPhone;
  const adultAddress = useFather ? data.fatherAddressDetails : data.motherAddressDetails;
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
    imagen_autoriza_total: data.imageChoice === 'yes' ? '[X]' : '[ ]',
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
  const pdf = folder.createFile(copy.getAs(MimeType.PDF).setName(sanitizeFileName(title) + '.pdf'));
  return { field: 'finalDeclaration', category: folder.getName(), name: pdf.getName(), mimeType: MimeType.PDF, size: pdf.getSize(), url: pdf.getUrl(), id: pdf.getId() };
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
  const pdf = folder.createFile(source.getAs(MimeType.PDF).setName(sanitizeFileName(title) + '.pdf'));
  source.setTrashed(true);
  return { field: 'generatedReport', category: folder.getName(), name: pdf.getName(), mimeType: MimeType.PDF, size: pdf.getSize(), url: pdf.getUrl(), id: pdf.getId() };
}

function sendFamilyReports(data, reports) {
  const recipients = uniqueValues([data.motherEmail, data.fatherEmail].map(value => String(value || '').trim().toLowerCase()));
  if (!recipients.length || !reports.length) return;
  const attachments = reports.map(report => DriveApp.getFileById(report.id).getBlob());
  MailApp.sendEmail({
    to: recipients.join(','),
    subject: `Resumen de admisión - ${data.legalFullName || data.fullName || 'Estudiante'} - Casa de las Estrellas`,
    body: ['Adjuntamos el resumen de la solicitud y las autorizaciones registradas durante el proceso de admisión. Conserve este documento para sus registros.','',buildDeclarationSummary(data)].join('\n'),
    attachments
  });
}

function buildDeclarationSummary(data) {
  const answer = value => value === 'yes' ? 'Sí' : value === 'no' ? 'No' : 'No registrado';
  return [
    'Declaraciones registradas:',
    `- Información verdadera y completa: ${answer(data.confirmTruth)}`,
    `- Tratamiento de datos autorizado: ${answer(data.confirmDataUse)}`,
    `- Contacto autorizado: ${answer(data.confirmContact)}`,
    `- Uso de imagen: ${data.imageChoice === 'yes' ? 'Autorizado' : data.imageChoice === 'no' ? 'No autorizado' : 'No registrado'}`,
    `- Proceso leído y comprendido: ${answer(data.confirmRead)}`,
    `- Reconoce que el envío no garantiza cupo: ${answer(data.confirmNoGuarantee)}`,
    `- Envío aceptado para revisión: ${answer(data.confirmSubmit)}`,
    `- Persona firmante: ${data.completedBy || 'No registrado'} (${data.completedByRelation || 'relación no registrada'})`,
    `- Fecha: ${formatDeclarationDate(data.submissionDate || new Date())}`
  ].join('\n');
}

function sendNotification(to, data, folder, files, pendingDocs, sheetUrl, reports) {
  const subject = `Nueva admisión: ${data.legalFullName || data.fullName || data.firstName || 'Estudiante'} (${data.studentId})`;
  const body = [
    'Se recibió un nuevo formulario de admisión.',
    '',
    `Estudiante: ${data.legalFullName || data.fullName || ''}`,
    `ID estudiante: ${data.studentId || ''}`,
    `Período escolar: ${data.schoolPeriod || ''}`,
    `Grado actual o solicitado: ${data.currentGrade || ''}`,
    `Encargado principal: ${data.motherName || data.fatherName || data.completedBy || ''}`,
    `Correo del encargado: ${data.motherEmail || data.fatherEmail || ''}`,
    `Teléfono: ${data.motherPhone || data.fatherPhone || ''}`,
    `Carpeta Drive: ${folder.getUrl()}`,
    `Google Sheet: ${sheetUrl}`,
    '',
    'Documentos recibidos:',
    files.length ? files.map(f => `- ${f.name}: ${f.url}`).join('\n') : '- Ninguno',
    '',
    'Documentos pendientes:',
    pendingDocs.length ? pendingDocs.map(d => `- ${d}`).join('\n') : '- Sin pendientes',
    '',
    buildDeclarationSummary(data)
  ].join('\n');
  const attachments = (reports || []).map(report => DriveApp.getFileById(report.id).getBlob());
  MailApp.sendEmail({ to, subject, body, attachments });
}

function pick(data, keys) {
  const obj = {};
  keys.forEach(k => obj[k] = data[k] || '');
  return obj;
}

function pickDiet(data) {
  const obj = pick(data, ['studentId','fullName','dietType','dietOther','foodAllergies','foodRestrictions','dietAdditional','schoolMealNotes']);
  Object.keys(data).forEach(k => { if (k.indexOf('diet_') === 0) obj[k] = data[k]; });
  return obj;
}

function pickVaccines(data) {
  return pick(data, ['studentId','fullName','vaccinesUpToDate','vaccineComments','submissionDate','completedBy']);
}

function sanitizeFileName(name) {
  return String(name).replace(/[\\/:*?"<>|#%{}~&]/g, '-').replace(/\s+/g, ' ').trim().slice(0, 180);
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
