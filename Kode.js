function doGet(e) {
  var halaman = e.parameter.p || e.parameter.page || "Index";
  var scriptUrl = ScriptApp.getService().getUrl();

  // 1. Cek apakah halaman sudah ada di Cache agar super cepat
  var cache = CacheService.getScriptCache();
  var cachedHalaman = cache.get("html_" + halaman);
  
  // Jika ada di cache, kita bisa langsung return (Opsional, tergantung kompleksitas script)
  // Namun cara paling aman adalah mengoptimalkan cara include file:

  try {
    var tmp = HtmlService.createTemplateFromFile(halaman);
    tmp.appUrl = scriptUrl;
    tmp.getParam = e.parameter.get || "";
    
    // Gunakan setSandboxMode untuk performa lebih stabil
    return tmp.evaluate()
        .setTitle('PORTALDATA_RENKEU')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    var fallback = HtmlService.createTemplateFromFile("Index");
    fallback.appUrl = scriptUrl;
    return fallback.evaluate();
  }
}

// Fungsi ini tetap ada tapi sebaiknya jangan sering dipanggil via JS
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getHtmlContent(pageName) {
  try {
    // Membaca file HTML berdasarkan nama (Pos, Upload, Dashboard)
    return HtmlService.createHtmlOutputFromFile(pageName.trim()).getContent();
  } catch (e) {
    throw new Error("File " + pageName + " tidak ditemukan di sidebar.");
  }
}

function clearServerCache() {
  try {
    var cache = CacheService.getScriptCache();
    // Menghapus semua cache yang tersimpan di script
    cache.removeAll(['daftarGlobal', 'lastUpdate']); 
    return "Cache berhasil dibersihkan";
  } catch (e) {
    return "Gagal membersihkan cache: " + e.toString();
  }
}

function getPageContent(page) {
  // Jika Dashboard dipanggil, jalankan fungsi render
  if (page === 'Dashboard' || !page) {
    return renderLandingPage();
  }
  return "<div class='text-white p-10'>Konten halaman " + page + " siap diisi.</div>";
}

// Fungsi render Landing Page dengan pembaruan pada bagian DOKUMEN
function renderLandingPage() {
  const modules = [
    { 
      title: "Pengiriman", 
      id: "Pengiriman", 
      icon: "fa-envelope-open-text", 
      img: "https://images.unsplash.com/photo-1519003722824-194d4455a60c?w=800",
      description: "Pusat Pengiriman Metadata dan Portal."
    },
    { 
      title: "Peraturan", 
      id: "Peraturan", 
      icon: "fa-scale-balanced", 
      img: "https://images.unsplash.com/photo-1566937169390-7be4c63b8a0e?q=60&w=600&auto=format",
      description: "Akses daftar regulasi, kebijakan, dan panduan hukum terbaru."
    },
    { 
      title: "Dokumen Lain & SK", 
      id: "Dokumen", 
      // Ikon diubah menjadi fa-book-bookmark sesuai visual screenshot
      icon: "fa-book-bookmark", 
      // Gambar diubah ke tema perpustakaan/rak buku sesuai Screenshot (235)_2.jpg
      img: "https://images.unsplash.com/photo-1507842217343-583bb7270b66?auto=format&fit=crop&w=800&q=80",
      description: "Pusat penyimpanan berkas digital,dan manajemen dokumen."
    },
    { 
      title: "Informasi", 
      id: "Informasi", 
      icon: "fa-bullhorn", 
      img: "https://images.unsplash.com/photo-1504711434969-e33886168f5c?w=800",
      description: "Pusat Pengumuman internal, dan bantuan teknis."
    }
  ];

  let cards = modules.map(m => `
    <div class="glass-card group cursor-pointer overflow-hidden transform hover:-translate-y-4 transition-all duration-500 shadow-2xl rounded-[2rem] border border-white/5 bg-slate-900/40" onclick="loadPage('${m.id}')">
      <div class="h-80 overflow-hidden relative">
        <!-- Image dengan penanganan error agar tidak kosong jika link bermasalah -->
        <img src="${m.img}" 
             onerror="this.src='https://images.unsplash.com/photo-1516979187457-637abb4f9353?w=800';"
             class="w-full h-full object-cover opacity-30 group-hover:opacity-100 group-hover:scale-110 transition-all duration-700">
        
        <!-- Gradient Overlay agar teks terbaca -->
        <div class="absolute inset-0 bg-gradient-to-t from-[#020617] via-[#020617]/60 to-transparent"></div>
        
        <div class="absolute inset-x-0 bottom-0 p-8">
           <!-- Icon Container -->
           <div class="w-12 h-12 bg-cyan-500/20 backdrop-blur-md rounded-xl flex items-center justify-center mb-4 border border-cyan-500/30 group-hover:bg-cyan-500 group-hover:text-slate-900 transition-all duration-300">
              <i class="fa-solid ${m.icon} text-xl text-cyan-400 group-hover:text-slate-900"></i>
           </div>
           
           <!-- Title -->
           <h3 class="text-2xl font-black text-white tracking-tighter uppercase mb-2 group-hover:text-cyan-400 transition-colors italic">
             ${m.title}
           </h3>

           <!-- Deskripsi -->
           <p class="text-xs text-slate-400 leading-relaxed opacity-0 group-hover:opacity-100 transform translate-y-4 group-hover:translate-y-0 transition-all duration-500 font-medium">
             ${m.description}
           </p>
        </div>
      </div>
    </div>
  `).join('');

  return `
    <div class="w-full max-w-[1400px] mx-auto py-10">
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-8 px-6">
        ${cards}
      </div>
      
      <div class="mt-16 text-center">
        <p class="text-slate-500 text-[10px] uppercase tracking-[0.4em] font-bold opacity-50">
          DINAS KESEHATAN KAB.KUBU RAYA © MANAJEMEN DATA RENCANA KERJA DAN KEUANGAN
        </p>
      </div>
    </div>
  `;
}

// ------------ OTORITAS ADMIN ------------ //
function verifikasiEmailAktif() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAdmin = ss.getSheetByName("Data_Admin"); // Ubah ke nama sheet Admin Anda
    
    if (!sheetAdmin) {
      return { status: "error", message: "Sheet 'Data_Admin' tidak ditemukan!" };
    }
    
    const data = sheetAdmin.getDataRange().getValues();
    let emailDitemukan = false;

    // Loop mulai baris 2 (index 1)
    for (let i = 1; i < data.length; i++) {
      // Kolom A [index 0] = Email
      const dbEmail = data[i][0] ? data[i][0].toString().trim().toLowerCase() : ""; 
      
      if (dbEmail === userEmail) {
        emailDitemukan = true;
        break; 
      }
    }

    return { 
      status: emailDitemukan ? "sukses" : "ditolak", 
      emailAktif: userEmail, 
      message: emailDitemukan ? "AKSES DITERIMA" : "EMAIL TIDAK TERDAFTAR",
      appUrl: ScriptApp.getService().getUrl() 
    };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

/**
 * FUNGSI 2: LOGIN MANUAL (USER & PASS)
 * Elemen: Email[A], Username[B], Password[C], Role[D]
 */
function checkLogin(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetAdmin = ss.getSheetByName("Data_Admin"); 
    const data = sheetAdmin.getDataRange().getValues();
    const emailLoginAktif = Session.getActiveUser().getEmail().toLowerCase().trim();

    for (let i = 1; i < data.length; i++) {
      // Pemetaan kolom sesuai permintaan baru:
      const dbEmail    = data[i][0] ? data[i][0].toString().trim().toLowerCase() : ""; 
      const dbUsername = data[i][1] ? data[i][1].toString().trim() : ""; 
      const dbPassword = data[i][2] ? data[i][2].toString().trim() : ""; 
      const dbRole     = data[i][3] ? data[i][3].toString().trim() : ""; 

      // Validasi Username & Password
      if (dbUsername === username && dbPassword === password) {
        // Validasi tambahan: Apakah email Google yang sedang login sama dengan email di sheet?
        if (dbEmail === emailLoginAktif) {
          return { 
            status: "sukses", 
            role: dbRole,
            emailAktif: emailLoginAktif, 
            message: "AKSES DITERIMA",
            appUrl: ScriptApp.getService().getUrl() 
          };
        } else {
          // Kasus: User/Pass benar tapi email google salah
          return { status: "ditolak", message: "Gunakan Akun Google: " + dbEmail };
        }
      }
    }
    return { status: "error", message: "USERNAME ATAU PASSWORD SALAH" };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}
                // ADMIN POS PENGIRIMAN///
//------------------------------------------------------///
function getNextFormId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Formulir");
    if (!sheet) return "F-0001";

    const lastRow = sheet.getLastRow();
    
    // Jika sheet kosong atau hanya ada header
    if (lastRow < 2) return "F-0001";

    // Ambil data ID di Kolom B (Indeks kolom 2)
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let maxNumber = 0;

    for (let i = 0; i < data.length; i++) {
      let cellValue = data[i][0].toString().trim();
      if (cellValue.startsWith("F-")) {
        let num = parseInt(cellValue.split("-")[1], 10);
        if (!isNaN(num) && num > maxNumber) {
          maxNumber = num;
        }
      }
    }
    // ... kode pencarian maxNumber ...
        const nextId = "F-" + (maxNumber + 1).toString().padStart(4, '0');
        
        Logger.log("ID BARU DIHASILKAN: " + nextId); // Tambahkan baris ini untuk cek manual
        return nextId;
      } catch (e) {
        Logger.log("ERROR SERVER: " + e.message);
        return "F-0001";
      }
    }

function publishProses(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master_Formulir");
  const templateSheet = ss.getSheetByName("Template_Monitoring");
  const id = formData.id;

  let linkInput = formData.url ? formData.url.toString().trim() : "";
  let linkFinal = linkInput;
  let linkSheetMonitoring = "";

  try {
    const jenisOtomatis = (linkInput.toLowerCase().startsWith("http")) ? "PORTAL" : "METADATA";

    // --- ALUR KERJA ASLI: FOLDER & MONITORING ---
    if (jenisOtomatis === 'METADATA' && formData.status === 'PUBLISH') {
      const folderInduk = DriveApp.getFolderById('1rvBMGrGYooNkxFcXwl1y4CsKZhna1Qil');
      const subFolders = folderInduk.getFolders();
      let folderKetemu = false;
      
      while(subFolders.hasNext()){
        const f = subFolders.next();
        if(f.getName().startsWith(id)) {
          linkFinal = f.getUrl();
          folderKetemu = true;
          break;
        }
      }
      
      if(!folderKetemu){
        const newFolder = folderInduk.createFolder(id + " - " + formData.judul);
        newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        linkFinal = newFolder.getUrl(); 
      }

      let sheetTujuan = ss.getSheetByName(id);
      if (!sheetTujuan && templateSheet) {
        sheetTujuan = templateSheet.copyTo(ss).setName(id);
      }
      if(sheetTujuan) linkSheetMonitoring = ss.getUrl() + "#gid=" + sheetTujuan.getSheetId();
    }

    // --- PENULISAN DATA KE MASTER ---
    const valuesB = masterSheet.getRange("B:B").getValues();
    let baris = -1;
    for (let i = 0; i < valuesB.length; i++) {
      if (valuesB[i][0].toString() == id) { baris = i + 1; break; }
    }
    if (baris === -1) baris = masterSheet.getLastRow() + 1;

    const lmp = formData.lampiran_terpisah || ["", "", "", "", "", ""];
    const rumusTotal = `=IFERROR(COUNTA(INDIRECT("'"&B${baris}&"'!A2:A")); 0)`;

    // TULIS DATA PER KOLOM (Sangat Stabil)
    masterSheet.getRange(baris, 1).setValue(new Date());        // A: Tanggal
    masterSheet.getRange(baris, 2).setValue(id);                // B: ID
    masterSheet.getRange(baris, 3).setValue(formData.judul);     // C: Judul
    masterSheet.getRange(baris, 4).setValue(formData.status);    // D: Status
    
    // KOLOM E (FORMAT) - WAJIB MUNCUL SEKARANG
    masterSheet.getRange(baris, 5).setValue(formData.format || "-");
    
    masterSheet.getRange(baris, 6).setFormula(rumusTotal);      // F: Count
    masterSheet.getRange(baris, 7).setValue(linkFinal);         // G: Link Drive
    masterSheet.getRange(baris, 8).setValue(jenisOtomatis);     // H: Jenis
    masterSheet.getRange(baris, 9).setValue(linkSheetMonitoring);// I: Link Sheet

    // TULIS KOLOM J SAMPAI O (JUDUL FILE 1-6)
    masterSheet.getRange(baris, 10, 1, 6).setValues([lmp]);

    SpreadsheetApp.flush(); 

    return { status: "success", message: "Sinkronisasi Berhasil untuk " + id };

  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function otomatisasiJenisFormulir() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_Formulir");
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return;

  // Ambil data Kolom G (Link) dan Kolom H (Jenis)
  // Kolom G adalah index ke-7, Kolom H adalah index ke-8
  const range = sheet.getRange(2, 7, lastRow - 1, 2);
  const values = range.getValues();

  const updatedValues = values.map(row => {
    const link = row[0] ? row[0].toString().trim() : "";
    
    // Logika Penentuan Otomatis
    if (link.toLowerCase().startsWith("http")) {
      return [row[0], "PORTAL"]; // Jika ada link, set PORTAL
    } else {
      return [row[0], "METADATA"]; // Jika tidak ada link, set METADATA
    }
  });

  // Simpan kembali ke spreadsheet
  range.setValues(updatedValues);
}

// ------------------ PORTAL PENGIRIMAN ----------------//
function getDaftarPublish() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_Formulir");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 

  // --- UPDATE: Ambil data dari kolom A sampai O (15 kolom) ---
  // Baris mulai: 2, Kolom mulai: 1, Jumlah baris: lastRow-1, Jumlah kolom: 15
  const range = sheet.getRange(2, 1, lastRow - 1, 15);
  const data = range.getDisplayValues();
  const daftar = [];

  data.forEach((row) => {
    const id = row[1] ? row[1].toString().trim() : "";
    const judul = row[2] ? row[2].toString().trim() : "";
    const status = row[3] ? row[3].toString().trim().toUpperCase() : "";
    const jenis = row[7] ? row[7].toString().trim().toUpperCase() : ""; // Kolom H
    
    if (id !== "" && status === "PUBLISH") {
      
      // --- LOGIKA BARU: Ambil Kolom J sampai O (Index 9 sampai 14) ---
      // Kita filter agar hanya judul yang ada isinya yang masuk ke array
      const lampiranMetadata = [
        row[9],  // J
        row[10], // K
        row[11], // L
        row[12], // M
        row[13], // N
        row[14]  // O
      ].filter(item => item !== "" && item !== "-");

      daftar.push({
        id: id,
        judul: judul,
        status: status,
        format: row[4],       // Kolom E
        totalPengirim: row[5], // Kolom F
        url: row[6],          // Kolom G
        jenis: jenis,         // Kolom H
        // Tambahkan properti baru ini agar bisa dibaca di Portal.html
        lampiran: lampiranMetadata 
      });
    }
  });

  console.log("Data dikirim ke HTML dengan Lampiran: " + JSON.stringify(daftar));
  return daftar;
}

function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

// Mencatat email pengakses PORTAL ke Sheet Kontrol
function catatAksesPortal(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKontrol = ss.getSheetByName("KONTROL_" + id);
  const userEmail = Session.getActiveUser().getEmail();
  
  if (sheetKontrol) {
    sheetKontrol.appendRow([new Date(), userEmail]);
    return true;
  }
  throw new Error("Sistem tidak menemukan Sheet Kontrol untuk ID ini.");
}

function getUnitKerja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_Unit");
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return values.map(r => r[0]).filter(item => item !== "");
}

//-------- PROSES PENGIRIMAN --------//

function simpanDataKeSheets(dataForm, files) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const idForm = dataForm.idForm.toString().trim();
    
    // --- 1. PROSES FOLDER DRIVE ---
    const folderInduk = DriveApp.getFolderById('1rvBMGrGYooNkxFcXwl1y4CsKZhna1Qil');
    let targetFolder = null;
    const subFolders = folderInduk.getFolders();
    while (subFolders.hasNext()) {
      const f = subFolders.next();
      if (f.getName().startsWith(idForm)) { targetFolder = f; break; }
    }
    if (!targetFolder) targetFolder = folderInduk.createFolder(idForm + " - " + dataForm.judulForm);

    // --- 2. PROSES UPLOAD & PEMETAAN KOLOM K-P ---
    let linkLampiran = ["", "", "", "", "", ""]; 
    const sheetMaster = ss.getSheetByName("Master_Formulir");
    const dataMaster = sheetMaster.getDataRange().getValues();
    const barisForm = dataMaster.find(r => r[1].toString() === idForm);

    files.forEach(file => {
      const blob = Utilities.newBlob(Utilities.base64Decode(file.base64), file.mimeType, file.name);
      const newFile = targetFolder.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const linkUrl = newFile.getUrl();

      if (barisForm) {
        // Cari posisi judul di kolom J-O Master (Index 9-14)
        for (let i = 9; i <= 14; i++) {
          if (barisForm[i] === file.type) {
            linkLampiran[i - 9] = linkUrl;
            break;
          }
        }
      }
    });

    // --- 3. SUSUN DATA BARIS (A-P) ---
    const dataBarisLengkap = [
      new Date(),              // A: Tanggal
      idForm,                  // B: ID Formulir
      dataForm.nama,           // C: Nama Pengirim
      "'" + dataForm.nip,      // D: NIP
      dataForm.pangkat,        // E: Pangkat/Gol
      dataForm.jabatan,        // F: Jabatan
      "'" + dataForm.whatsapp, // G: No HP
      dataForm.unit,           // H: Unit Kerja
      Session.getActiveUser().getEmail(), // I: Email
      targetFolder.getUrl(),   // J: LINK FOLDER DRIVE (Sesuai Permintaan)
      linkLampiran[0],         // K: Judul 1
      linkLampiran[1],         // L: Judul 2
      linkLampiran[2],         // M: Judul 3
      linkLampiran[3],         // N: Judul 4
      linkLampiran[4],         // O: Judul 5
      linkLampiran[5]          // P: Link Judul 6
    ];

    // --- 4. LOGIKA PENYIMPANAN PRIORITAS ---
    const sheetIdOtomatis = ss.getSheetByName(idForm);
    const sheetMonitoring = ss.getSheetByName("Template_Monitoring");

    if (sheetIdOtomatis) {
      // JIKA SHEET ID ADA: Hanya simpan ke Sheet ID tersebut
      sheetIdOtomatis.appendRow(dataBarisLengkap);
    } else if (sheetMonitoring) {
      // JIKA SHEET ID TIDAK ADA: Alihkan ke Template_Monitoring
      sheetMonitoring.appendRow(dataBarisLengkap);
    } else {
      throw new Error("Target Sheet (ID atau Monitoring) tidak ditemukan!");
    }

    // --- 5. UPDATE COUNTER (PENTING AGAR TIDAK ERROR) ---
    updateCounterPengirim(idForm);

    SpreadsheetApp.flush();
    return {
      status: "sukses",
      appUrl: ScriptApp.getService().getUrl() 
    };
  } catch (e) {
    throw new Error(e.toString());
  }
}

function rekapTotalPengirim() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master_Formulir");
  const data = masterSheet.getDataRange().getValues();

  // Mulai dari baris ke-2 (indeks 1) untuk melewati header
  for (let i = 1; i < data.length; i++) {
    const formatFormulir = data[i][4]; // Kolom E (indeks 4) adalah FORMAT
    const idJudul = data[i][1];       // Kolom B (indeks 1) adalah ID Judul

    // JIKA FORMAT ADALAH LINK EKSTERNAL, PAKSA NILAI MENJADI 0
    if (formatFormulir === "LINK EKSTERNAL") {
      masterSheet.getRange(i + 1, 6).setValue(0); // Kolom F (indeks 6) adalah Total Pengirim
    } else {
      // LOGIKA ASLI: Hitung jumlah data dari sheet terkait (F-0001, dst)
      try {
        const targetSheet = ss.getSheetByName(idJudul);
        if (targetSheet) {
          // Menghitung jumlah baris berisi data (dikurangi header)
          const total = targetSheet.getLastRow() - 1;
          masterSheet.getRange(i + 1, 6).setValue(total < 0 ? 0 : total);
        }
      } catch (e) {
        masterSheet.getRange(i + 1, 6).setValue(0);
      }
    }
  }
}

// ------------ DASHBOARD OVERVIEW -----------------//

function ambilDataMasterDariSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master_Formulir");
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    // Pastikan ada data di bawah header (baris 2 ke atas)
    if (lastRow < 2) return [];

    // Ambil data baris 2, kolom 1, sebanyak (lastRow-1) baris dan 9 kolom (A sampai I)
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    return data.map(row => {
      return {
        tanggal: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy") : row[0], // [A]
        id: row[1],      // [B]
        judul: row[2],   // [C]
        status: row[3],  // [D]
        total: row[5],   // [F]
        drive: row[6],   // [G]
        sheet: row[8]    // [I]
      };
    });
  } catch (e) {
    return [];
  }
}

function updateStatusFormulir(idJudul, statusBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master_Formulir");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == idJudul) { // Cek Kolom B (ID)
      sheet.getRange(i + 1, 4).setValue(statusBaru); // Update Kolom D (Status)
      return "Sukses";
    }
  }
  return "Gagal";
}

// ------------------------------UPLOAD DOKUMEN ---------------------------//
// Fungsi untuk mendapatkan ID baru secara otomatis
function getNewId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_Dokumen");
    const lastRow = sheet.getLastRow();
    
    // Jika sheet masih kosong (hanya ada header)
    if (lastRow < 2) {
      return "D-0001";
    }
    
    // Ambil ID terakhir dari kolom B (baris terakhir)
    const lastId = sheet.getRange(lastRow, 2).getValue().toString();
    
    // Pastikan format ID benar (D-XXXX)
    if (!lastId.includes("-")) {
      return "D-0001";
    }
    
    const parts = lastId.split('-');
    const nextNum = parseInt(parts[1]) + 1;
    
    // Format menjadi D-0002, D-0003, dst.
    return "D-" + nextNum.toString().padStart(4, '0');
    
  } catch (err) {
    console.log("Error getNewId: " + err.message);
    return "D-0001"; // Fallback jika terjadi error
  }
}

// Fungsi untuk mengambil semua data dokumen untuk tabel rekapan
function getDataRekapan() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_Dokumen");
    
    // Validasi apakah sheet ada
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    // Jika hanya ada header atau kosong, kirim array kosong
    if (lastRow < 2) return [];
    
    // Mengambil data dari baris 2, kolom 1, sebanyak (lastRow - 1) baris, dan 7 kolom
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    return data;
  } catch (e) {
    // Log error di server agar Anda bisa melihatnya di tab 'Executions'
    console.error("Error getDataRekapan: " + e.message);
    return [];
  }
}

function uploadDokumenProses(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_Dokumen");
    const folderId = "1GRLzEGId8vrIcFISLYykAizmduiw-8Kw"; // ID Folder Anda
    
    let fileUrl = "";
    
    // PROSES UPLOAD FILE (Jika ada file yang dipilih)
    if (data.fileBlob && data.fileBlob.bytes) {
      const folder = DriveApp.getFolderById(folderId);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(data.fileBlob.bytes), 
        data.fileBlob.mimeType, 
        data.fileBlob.fileName
      );
      const file = folder.createFile(blob);
      
      // SET IZIN AKSES: Agar file bisa didownload secara publik
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      // FORMAT LINK: Mengubah ke format Direct Download agar Safelink tidak ERROR
      fileUrl = "https://docs.google.com/uc?export=download&id=" + file.getId();
    }

    if (data.mode === 'edit') {
      // --- PROSES UPDATE (EDIT) ---
      const dataSheet = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      // Cari baris berdasarkan ID (Kolom B / Index 1)
      for (let i = 0; i < dataSheet.length; i++) {
        if (dataSheet[i][1] == data.id_doc) {
          rowIndex = i + 1; // Baris di spreadsheet (1-based)
          break;
        }
      }

      if (rowIndex !== -1) {
        // Update baris yang ditemukan
        sheet.getRange(rowIndex, 3).setValue(data.jenis);       // Kolom C
        sheet.getRange(rowIndex, 4).setValue(data.judul);       // Kolom D
        sheet.getRange(rowIndex, 5).setValue(data.nomor);       // Kolom E
        sheet.getRange(rowIndex, 6).setValue(data.tgl_terbit);  // Kolom F
        
        // Update Kolom G hanya jika user mengupload file baru
        if (fileUrl !== "") {
          sheet.getRange(rowIndex, 7).setValue(fileUrl);
        }
        
        return "Berhasil: Data dengan ID " + data.id_doc + " telah diperbarui.";
      } else {
        return "Error: ID " + data.id_doc + " tidak ditemukan di database.";
      }

    } else {
      // --- PROSES SIMPAN BARU ---
      // Jika mode bukan edit, tambahkan baris baru
      sheet.appendRow([
        new Date(),   // Kolom A
        getNewId(),   // Kolom B (Pastikan fungsi getNewId tersedia)
        data.jenis,   // Kolom C
        data.judul,   // Kolom D
        data.nomor,   // Kolom E
        data.tgl_terbit, // Kolom F
        fileUrl       // Kolom G
      ]);
      return "Berhasil: Data baru telah disimpan.";
    }
  } catch (err) {
    return "Error Server: " + err.toString();
  }
}

function updateStatusPenyimpananDokumen(id, statusBaru) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Data_Dokumen"); 
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === id) { // ID di kolom B
        // PINDAH KE KOLOM H: Index 8 (Kolom ke-8)
        // Kolom G (7) tetap untuk Link, Kolom H (8) untuk Status
        sheet.getRange(i + 1, 8).setValue(statusBaru); 
        return "Sukses";
      }
    }
    return "Gagal: ID tidak ditemukan";
  } catch (e) {
    return "Error: " + e.toString();
  }
}