// ============================================================
// PORTOFOLIO DWI SYAFITRI — Google Apps Script
// Mendukung:
// 1) upload file gambar dari komputer/HP
// 2) link gambar Google Drive
//
// Penting:
// - Buat file HTML di Apps Script dengan nama: Index
// - Isi CONFIG di bawah sesuai Spreadsheet dan folder Drive Anda.
// - File Drive yang dipakai sebagai link gambar harus dibagikan:
//   "Siapa saja yang memiliki link" -> "Pelihat".
// ============================================================

const CONFIG = {
  ADMIN_EMAIL: 'dwisyafitri82@gmail.com',
  ADMIN_PASS: 'DWIadmin@288',
  SPREADSHEET_ID: '', // Kosongkan jika script terikat langsung ke Spreadsheet.
  FOLDER_ID: ''       // Kosongkan untuk memakai My Drive.
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portofolio Dwi Syafitri')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkLogin(email, pass) {
  return String(email || '').trim() === CONFIG.ADMIN_EMAIL &&
    String(pass || '') === CONFIG.ADMIN_PASS;
}

function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('Spreadsheet belum terhubung. Isi CONFIG.SPREADSHEET_ID.');
  }
  return active;
}

function getUploadFolder() {
  return CONFIG.FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.FOLDER_ID)
    : DriveApp.getRootFolder();
}

function getSheet(name, headers, defaultRows) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) sheet.appendRow(headers);
    (defaultRows || []).forEach(function(row) { sheet.appendRow(row); });
  }
  return sheet;
}

function getDriveFileId(url) {
  if (!url || String(url).indexOf('data:') === 0) return '';
  const match = String(url).trim().match(
    /(?:[?&]id=|\/d\/|\/file\/d\/)([a-zA-Z0-9_-]{10,})/
  );
  return match ? match[1] : '';
}

function normalizeDriveUrl(url, mode) {
  if (!url) return '';
  const value = String(url).trim();
  if (value.indexOf('data:') === 0) return value;
  const id = getDriveFileId(value);
  if (!id) return value;
  if (mode === 'download') {
    return 'https://drive.google.com/uc?export=download&id=' + id;
  }
  return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1200';
}

function dataUriParts(dataUri) {
  const value = String(dataUri || '');
  const comma = value.indexOf(',');
  if (comma < 0) throw new Error('Format gambar tidak valid.');
  const header = value.slice(0, comma);
  const data = value.slice(comma + 1);
  const mimeMatch = header.match(/^data:([^;]+);base64$/);
  const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
  if (mime.indexOf('image/') !== 0) throw new Error('File harus berupa gambar.');
  if (!data) throw new Error('Data gambar kosong.');
  return { mime: mime, bytes: Utilities.base64Decode(data) };
}

/**
 * Menyimpan upload gambar ke Google Drive.
 * Jika Drive gagal, data URI dikembalikan agar gambar tetap dapat ditampilkan.
 */
function storeImage(base64Data, fileName) {
  if (!base64Data) throw new Error('Data gambar kosong.');
  try {
    const parts = dataUriParts(base64Data);
    const blob = Utilities.newBlob(
      parts.bytes,
      parts.mime,
      fileName || ('image_' + new Date().getTime() + '.jpg')
    );
    const file = getUploadFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return normalizeDriveUrl(file.getUrl(), 'image');
  } catch (error) {
    Logger.log('Upload Drive gagal, memakai fallback data URI: ' + error.message);
    return base64Data;
  }
}

function uploadCV(base64Data, fileName) {
  if (!base64Data) throw new Error('Data CV kosong.');
  const comma = String(base64Data).indexOf(',');
  const raw = comma >= 0 ? String(base64Data).slice(comma + 1) : String(base64Data);
  const bytes = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(
    bytes,
    'application/pdf',
    fileName || ('cv_' + new Date().getTime() + '.pdf')
  );
  const file = getUploadFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return normalizeDriveUrl(file.getUrl(), 'download');
}

function getPortfolioData() {
  const result = {
    profile: {
      name: 'Dwi Syafitri, S.Ak',
      title: 'Finance & Accounting | Business Automation Enthusiast',
      aboutMe: '',
      email: '',
      phone: '',
      location: '',
      cvUrl: '',
      avatarUrl: '',
      skills: []
    },
    projects: [],
    experiences: [],
    educations: []
  };

  const profile = getSheet(
    'profile',
    ['Name', 'Title', 'AboutMe', 'Email', 'Phone', 'Location', 'CVUrl', 'AvatarUrl', 'Skills']
  );
  const profileRows = profile.getDataRange().getValues();
  if (profileRows.length > 1) {
    const row = profileRows[1];
    result.profile = {
      name: String(row[0] || result.profile.name),
      title: String(row[1] || ''),
      aboutMe: String(row[2] || ''),
      email: String(row[3] || ''),
      phone: String(row[4] || ''),
      location: String(row[5] || ''),
      cvUrl: normalizeDriveUrl(String(row[6] || ''), 'download'),
      avatarUrl: normalizeDriveUrl(String(row[7] || ''), 'image'),
      skills: row[8]
        ? String(row[8]).split(',').map(function(s) { return s.trim(); }).filter(Boolean)
        : []
    };
  }

  const projects = getSheet(
    'projects',
    ['ID', 'Nama', 'Kategori', 'Deskripsi', 'Tools', 'Tahun', 'Status', 'Demo', 'GitHub', 'ImageUrl1', 'ImageUrl2', 'ImageUrl3']
  );
  const projectRows = projects.getDataRange().getValues();
  projectRows.shift();
  result.projects = projectRows.filter(function(row) { return row[0] && row[1]; }).map(function(row) {
    return {
      id: String(row[0]), nama: String(row[1]), kategori: String(row[2] || ''),
      deskripsi: String(row[3] || ''), tools: String(row[4] || ''), tahun: String(row[5] || ''),
      status: String(row[6] || ''), demo: String(row[7] || ''), github: String(row[8] || ''),
      imageUrls: [row[9], row[10], row[11]].map(function(url) {
        return normalizeDriveUrl(String(url || ''), 'image');
      }).filter(Boolean)
    };
  });

  const experience = getSheet(
    'experience',
    ['ID', 'Role', 'Company', 'Period', 'Responsibilities']
  );
  const experienceRows = experience.getDataRange().getValues();
  experienceRows.shift();
  result.experiences = experienceRows.filter(function(row) { return row[0] && row[1]; }).map(function(row) {
    return {
      id: String(row[0]), role: String(row[1]), company: String(row[2] || ''),
      period: String(row[3] || ''),
      responsibilities: row[4] ? String(row[4]).split('\n').filter(Boolean) : []
    };
  });

  const education = getSheet(
    'education',
    ['ID', 'Degree', 'Institution', 'Period', 'Description']
  );
  const educationRows = education.getDataRange().getValues();
  educationRows.shift();
  result.educations = educationRows.filter(function(row) { return row[0] && row[1]; }).map(function(row) {
    return {
      id: String(row[0]), degree: String(row[1]), institution: String(row[2] || ''),
      period: String(row[3] || ''), description: String(row[4] || '')
    };
  });
  return result;
}

function saveProfile(obj) {
  const sheet = getSheet(
    'profile',
    ['Name', 'Title', 'AboutMe', 'Email', 'Phone', 'Location', 'CVUrl', 'AvatarUrl', 'Skills']
  );
  let avatarUrl = normalizeDriveUrl(obj.avatarUrl || '', 'image');
  let cvUrl = normalizeDriveUrl(obj.cvUrl || '', 'download');
  if (obj.avatarBase64) avatarUrl = storeImage(obj.avatarBase64, 'avatar_' + new Date().getTime() + '.jpg');
  if (obj.cvBase64) cvUrl = uploadCV(obj.cvBase64, obj.cvFileName || 'cv.pdf');
  const row = [
    obj.name || '', obj.title || '', obj.aboutMe || '', obj.email || '',
    obj.phone || '', obj.location || '', cvUrl, avatarUrl,
    Array.isArray(obj.skills) ? obj.skills.join(', ') : (obj.skills || '')
  ];
  if (sheet.getLastRow() < 2) sheet.appendRow(row);
  else sheet.getRange(2, 1, 1, row.length).setValues([row]);
  return { success: true, avatarUrl: avatarUrl, cvUrl: cvUrl };
}

function saveProject(obj) {
  const sheet = getSheet(
    'projects',
    ['ID', 'Nama', 'Kategori', 'Deskripsi', 'Tools', 'Tahun', 'Status', 'Demo', 'GitHub', 'ImageUrl1', 'ImageUrl2', 'ImageUrl3']
  );
  const imageUrls = [0, 1, 2].map(function(i) {
    return obj.imageBase64s && obj.imageBase64s[i]
      ? storeImage(obj.imageBase64s[i], 'project_' + new Date().getTime() + '_' + i + '.jpg')
      : normalizeDriveUrl((obj.imageUrls || [])[i] || '', 'image');
  });
  const values = [obj.nama || '', obj.kategori || '', obj.deskripsi || '', obj.tools || '',
    obj.tahun || '', obj.status || '', obj.demo || '', obj.github || ''].concat(imageUrls);
  const data = sheet.getDataRange().getValues();
  if (obj.id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(obj.id)) {
        sheet.getRange(i + 1, 2, 1, values.length).setValues([values]);
        return { success: true, id: obj.id };
      }
    }
  }
  const id = 'prj-' + new Date().getTime();
  sheet.appendRow([id].concat(values));
  return { success: true, id: id };
}

function deleteProject(id) {
  return deleteById('projects', id);
}

function saveExperience(obj) {
  const sheet = getSheet('experience', ['ID', 'Role', 'Company', 'Period', 'Responsibilities']);
  const values = [obj.role || '', obj.company || '', obj.period || '',
    Array.isArray(obj.responsibilities) ? obj.responsibilities.join('\n') : (obj.responsibilities || '')];
  return upsertRow(sheet, obj.id, 'exp-', values);
}

function deleteExperience(id) {
  return deleteById('experience', id);
}

function saveEducation(obj) {
  const sheet = getSheet('education', ['ID', 'Degree', 'Institution', 'Period', 'Description']);
  return upsertRow(sheet, obj.id, 'edu-', [
    obj.degree || '', obj.institution || '', obj.period || '', obj.description || ''
  ]);
}

function deleteEducation(id) {
  return deleteById('education', id);
}

function upsertRow(sheet, id, prefix, values) {
  const data = sheet.getDataRange().getValues();
  if (id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 2, 1, values.length).setValues([values]);
        return { success: true, id: id };
      }
    }
  }
  const newId = prefix + new Date().getTime();
  sheet.appendRow([newId].concat(values));
  return { success: true, id: newId };
}

function deleteById(sheetName, id) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}
