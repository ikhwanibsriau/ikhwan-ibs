// =========================================================================
// SCRIPT: setup.gs (INITIAL INSTALLATION)
// =========================================================================

function runSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = [
    { name: "Pengaturan", cols: ["Key", "Value"] },
    { name: "Pengguna", cols: ["ID", "Username", "Password", "Nama", "Role"] },
    { name: "Santri", cols: ["ID", "NIS", "NISN", "Nama", "Panggilan", "TmpLahir", "TglLahir", "JK", "Kelas", "Jenjang", "Asrama", "Kamar", "TahunMasuk", "Status", "Alamat", "Ayah", "Ibu", "Wali", "HP", "Catatan"] },
    { name: "Pelanggaran", cols: ["ID", "Tanggal", "NIS", "Kategori", "Pelanggaran", "Poin", "Tindakan", "Bukti"] },
    { name: "Perizinan", cols: ["ID", "NIS", "Tgl_Keluar", "Tgl_Kembali", "Alasan", "Status", "Catatan"] },
    { name: "Kesehatan", cols: ["ID", "Tgl", "NIS", "Keluhan", "Diagnosis", "Tindakan", "Petugas"] },
    { name: "CatatanSantri", cols: ["ID", "Tgl", "NIS", "Aspek", "Catatan", "Musyrif"] },
    { name: "KamarMonitoring", cols: ["Kamar", "Wali", "Ketua", "JmlPenghuni", "SkorBersih", "SkorDisiplin"] },
    { name: "MasterKelas", cols: ["ID", "NamaKelas"] },
    { name: "MasterAsrama", cols: ["ID", "NamaAsrama"] },
    { name: "LogAktivitas", cols: ["Timestamp", "User", "Action", "Details"] },
    { name: "CounterID", cols: ["LastID"] }
  ];

  sheets.forEach(sh => {
    let sheet = ss.getSheetByName(sh.name);
    if (!sheet) {
      sheet = ss.insertSheet(sh.name);
    }
    sheet.clear();
    sheet.getRange(1, 1, 1, sh.cols.length).setValues([sh.cols]).setFontWeight("bold").setBackground("#1b5e20").setFontColor("white");
    sheet.setFrozenRows(1);
  });

  // Hapus Sheet default (Sheet1) jika ada
  let sheet1 = ss.getSheetByName("Sheet1");
  if(sheet1) ss.deleteSheet(sheet1);

  // Set Default Super Admin (Password awal: admin123)
  const userSheet = ss.getSheetByName("Pengguna");
  const defaultPass = hashPassword("admin123");
  userSheet.appendRow(["USR001", "admin", defaultPass, "Super Administrator", "Super Admin"]);

  // Set Counter ID start
  const counterSheet = ss.getSheetByName("CounterID");
  counterSheet.getRange("A2").setValue(1000);

  // Set Data Sample untuk Santri agar Dashboard terlihat
  const santriSheet = ss.getSheetByName("Santri");
  santriSheet.appendRow(["SIM1001", "2026001", "-", "Ahmad Fulan", "Ahmad", "Jakarta", "2010-01-01", "L", "10", "SMA", "Umar", "1A", "2026", "Aktif", "Jakarta", "-", "-", "-", "0812345", ""]);

  Browser.msgBox("Setup Berhasil! Struktur Database dan Akun Super Admin telah dibuat.\\n\\nUsername: admin\\nPassword: admin123");
}

function hashPassword(password) {
  // Definisi ulang hash disini khusus untuk setup (harus sama dengan di code.gs)
  const SCRIPT_SALT = "SIM_IKHWAN_SECRET_2026_X9";
  const raw = password + SCRIPT_SALT;
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return signature.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}
