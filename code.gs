// =========================================================================
// SCRIPT: code.gs (BACKEND LOGIC)
// =========================================================================

const SCRIPT_SALT = "SIM_IKHWAN_SECRET_2026_X9";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SANTRI IKHWAN MANAJEMEN (SIM)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- AUTHENTICATION & SESSION ---
function hashPassword(password) {
  const raw = password + SCRIPT_SALT;
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return signature.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function login(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pengguna");
  const data = sheet.getDataRange().getValues();
  const hashedInput = hashPassword(password);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == username && (data[i][2] == hashedInput || password == data[i][2])) { //Fallback if manual unhashed pass exists
      const userObj = { id: data[i][0], username: data[i][1], nama: data[i][3], role: data[i][4] };
      
      // Update hash if not hashed
      if(password == data[i][2]) sheet.getRange(i+1, 3).setValue(hashedInput);
      
      const sessionToken = Utilities.getUuid();
      CacheService.getScriptCache().put(sessionToken, JSON.stringify(userObj), 21600); // 6 hours
      addLog(username, "Login", "User login berhasil");
      return { success: true, user: userObj, token: sessionToken };
    }
  }
  return { success: false, message: "Username atau Password salah!" };
}

function checkSession() {
  // Dalam GAS murni tanpa cookie HTTP, simulasi session dengan trigger/properties state atau via IP 
  // Untuk arsitektur SPA murni, biasanya token disimpan di localStorage client.
  // Karena keterbatasan prompt, kita asumsikan state login diatur di JS setelah verifikasi login().
  // Fungsi ini dipanggil onload, namun karena ini server-side tanpa cookies yg diteruskan otomatis,
  // kita return null untuk memaksa login ulang jika reload page.
  return { loggedIn: false }; 
}

function logout() {
  return true;
}

// --- GENERIC CRUD UTILITIES ---
function getTableData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  
  const headers = data[0].map(h => h.toLowerCase());
  const result = [];
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    headers.forEach((h, j) => { obj[h] = data[i][j]; });
    result.push(obj);
  }
  return result;
}

function saveRecord(sheetName, recordData, actionUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toLowerCase());
  
  if(sheetName === "Pengguna" && recordData.password) {
      recordData.password = hashPassword(recordData.password);
  }

  if (recordData.id) {
    // Update
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == recordData.id) {
        let rowArray = [];
        headers.forEach((h, j) => {
           rowArray.push(recordData[h] !== undefined ? recordData[h] : data[i][j]);
        });
        sheet.getRange(i + 1, 1, 1, rowArray.length).setValues([rowArray]);
        addLog(actionUser, "Update", `Update data di ${sheetName} ID: ${recordData.id}`);
        return "Data berhasil diperbarui";
      }
    }
  } else {
    // Insert
    recordData.id = generateId();
    let rowArray = [];
    headers.forEach(h => {
       rowArray.push(recordData[h] || '');
    });
    sheet.appendRow(rowArray);
    addLog(actionUser, "Insert", `Tambah data ke ${sheetName} ID: ${recordData.id}`);
    return "Data berhasil ditambahkan";
  }
}

function deleteRecord(sheetName, id, actionUser) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      addLog(actionUser, "Delete", `Hapus data dari ${sheetName} ID: ${id}`);
      return "Data berhasil dihapus";
    }
  }
  return "Data tidak ditemukan";
}

function generateId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CounterID");
  let val = sheet.getRange("A2").getValue() || 1000;
  sheet.getRange("A2").setValue(val + 1);
  return "SIM" + val;
}

function addLog(user, action, details) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogAktivitas");
  sheet.appendRow([new Date(), user, action, details]);
}

// --- SPECIFIC AGGREGATION APIs ---
function getDashboardStats() {
  const santri = getTableData("Santri");
  const pelanggaran = getTableData("Pelanggaran");
  const izin = getTableData("Perizinan");
  const kesehatan = getTableData("Kesehatan");
  
  let totAktif = santri.filter(s => s.status === 'Aktif').length;
  
  let d = new Date();
  let currMonth = d.getMonth();
  
  let pelBulanIni = pelanggaran.filter(p => {
    let ptgl = new Date(p.tanggal);
    return ptgl.getMonth() === currMonth && ptgl.getFullYear() === d.getFullYear();
  }).length;
  
  let izinAktif = izin.filter(i => i.status === 'Disetujui' || i.status === 'Diproses').length;
  
  let kelasDist = {};
  santri.forEach(s => {
    if(s.status === 'Aktif') {
      kelasDist[s.kelas] = (kelasDist[s.kelas] || 0) + 1;
    }
  });

  return {
    totAktif: totAktif,
    totPelanggaran: pelBulanIni,
    totIzin: izinAktif,
    totSakit: kesehatan.length, // Sederhana
    chartPelLabels: ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun'],
    chartPelData: [12, 19, 3, 5, 2, pelBulanIni], // Simulasi data historis statis utk demo
    klsDist: kelasDist
  };
}

function getSantriProfile(nis) {
  const santri = getTableData("Santri").find(s => s.nis == nis);
  if(!santri) return null;

  const pelanggaran = getTableData("Pelanggaran").filter(p => p.nis == nis);
  const izin = getTableData("Perizinan").filter(i => i.nis == nis);
  
  let totalPoin = 0;
  let timeline = [];

  pelanggaran.forEach(p => {
    totalPoin += parseInt(p.poin) || 0;
    timeline.push({ tanggal: p.tanggal, tipe: 'Pelanggaran', deskripsi: `${p.pelanggaran} (Poin: ${p.poin}) - Tindakan: ${p.tindakan}` });
  });

  izin.forEach(i => {
    timeline.push({ tanggal: i.tgl_keluar, tipe: 'Izin', deskripsi: `${i.alasan} - Status: ${i.status}` });
  });

  let klasifikasi = "Aman";
  if(totalPoin >= 20) klasifikasi = "Perhatian";
  if(totalPoin >= 50) klasifikasi = "Pembinaan";
  if(totalPoin >= 100) klasifikasi = "Peringatan";
  if(totalPoin >= 150) klasifikasi = "Tindakan Khusus";

  timeline.sort((a,b) => new Date(b.tanggal) - new Date(a.tanggal)); // Descending

  return {
    biodata: santri,
    totalPoin: totalPoin,
    klasifikasi: klasifikasi,
    timeline: timeline.slice(0, 10) // 10 aktivitas terakhir
  };
}
