// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    SPREADSHEET_ID: '1KA5ooJzylxucLht9M7zvERMjv3PmFaU4ge_Hi48fz-E',
    SHEETS: {
        PENGAJUAN: 'Pengajuan',
        USERS: 'Users',
        PENUGASAN_PP: 'Penugasan_PP',
        PROSES_PP: 'Proses_PP'
    }
};

// ============================================
// API ENDPOINTS - PUBLIC
// ============================================

/**
 * Get dashboard data for PPK
 */
function apiGetDashboardData(emailPPK) {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(CONFIG.SHEETS.PENGAJUAN);
        const data = sheet.getDataRange().getValues();
        const headers = data[0];

        // Filter by PPK email
        const userRows = data.slice(1).filter(row => row[4] === emailPPK); // Column E: emailPPK

        // Calculate stats
        const stats = {
            total: userRows.length,
            diproses: userRows.filter(r => r[14] === 'DIPROSES').length, // Column O: status
            selesai: userRows.filter(r => r[14] === 'SELESAI').length,
            ditolak: userRows.filter(r => r[14] === 'DITOLAK').length
        };

        // Format data
        const pengajuan = userRows.map(row => {
            const obj = {};
            headers.forEach((h, i) => obj[h] = row[i]);
            return obj;
        });

        return { success: true, stats, pengajuan };

    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * Generate next nomor proses (PR-XXX/2025)
 */
function apiGetNextNomorProses() {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = ss.getSheetByName(CONFIG.SHEETS.PENGAJUAN);
        const data = sheet.getDataRange().getValues();

        const year = new Date().getFullYear();
        const prefix = `PR-`;
        const suffix = `/${year}`;

        // Find max number
        let maxNum = 0;
        data.slice(1).forEach(row => {
            const nomor = row[0]; // Column A: nomorProses
            if (nomor && nomor.startsWith(prefix) && nomor.endsWith(suffix)) {
                const num = parseInt(nomor.replace(prefix, '').replace(suffix, ''));
                if (!isNaN(num) && num > maxNum) maxNum = num;
            }
        });

        const nextNum = String(maxNum + 1).padStart(3, '0');
        return { success: true, nomorProses: `${prefix}${nextNum}${suffix}` };

    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * Submit pengajuan dengan AUTO-ASSIGN PP (Round-robin)
 */
function apiSubmitPengajuan(data) {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

        // 1. AUTO-ASSIGN PP (Round-robin logic)
        const assignedPP = autoAssignPP(ss);
        if (!assignedPP.success) {
            return { success: false, error: 'Tidak ada PP yang tersedia untuk penugasan' };
        }

        // 2. Save to Pengajuan sheet
        const sheetPengajuan = ss.getSheetByName(CONFIG.SHEETS.PENGAJUAN);
        const timestamp = new Date();

        const rowData = [
            data.nomorProses,           // A: nomorProses
            new Date(data.tanggal),     // B: tanggal
            data.namaPaket,             // C: namaPaket
            data.uraianPekerjaan,       // D: uraianPekerjaan
            data.emailPPK,              // E: emailPPK
            data.namaPPK,               // F: namaPPK
            data.satker,                // G: satker
            data.jenisPengadaan,        // H: jenisPengadaan
            data.hpsNominal,            // I: hpsNominal
            data.jangkaWaktu,           // J: jangkaWaktu
            data.notaDinasUrl || '',    // K: notaDinasUrl
            data.hpsUrl || '',          // L: hpsUrl
            data.kontrakUrl || '',      // M: kontrakUrl
            data.spesifikasiUrl || '',  // N: spesifikasiUrl
            'DIPROSES',                 // O: status (default)
            assignedPP.email,          // P: emailPP (AUTO-ASSIGNED)
            assignedPP.nama,           // Q: namaPP (AUTO-ASSIGNED)
            timestamp,                  // R: lastUpdate
            ''                          // S: pdfUrl
        ];

        sheetPengajuan.appendRow(rowData);

        // 3. Save to Penugasan_PP sheet (tracking)
        const sheetPenugasan = ss.getSheetByName(CONFIG.SHEETS.PENUGASAN_PP);
        const penugasanId = Utilities.getUuid();
        sheetPenugasan.appendRow([
            penugasanId,
            data.nomorProses,
            assignedPP.nip,
            assignedPP.nama,
            assignedPP.email,
            timestamp,
            'AKTIF'
        ]);

        // 4. Create empty record in Proses_PP (untuk nanti diisi PP)
        const sheetProses = ss.getSheetByName(CONFIG.SHEETS.PROSES_PP);
        sheetProses.appendRow([
            Utilities.getUuid(),
            data.nomorProses,
            '',  // hasilNego
            '',  // hargaRealisasi
            '',  // namaPerusahaan
            '',  // npwp
            '',  // noKontrak
            'MENUNGGU_PROSES'  // status
        ]);

        // 5. Update PP workload counter (optional: track jumlah penugasan)
        updatePPWorkload(ss, assignedPP.email);

        // 6. Kirim notifikasi email ke PP (placeholder untuk next step)
        // sendEmailNotification(assignedPP.email, 'PENUGASAN_BARU', data);

        return {
            success: true,
            message: 'Pengajuan berhasil disimpan',
            nomorProses: data.nomorProses,
            assignedPP: {
                nama: assignedPP.nama,
                email: assignedPP.email,
                nip: assignedPP.nip
            }
        };

    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * Upload file ke Google Drive
 */
function apiUploadFile(base64Data, filename, mimeType) {
    try {
        // Decode base64
        const decoded = Utilities.base64Decode(base64Data.split(',')[1] || base64Data);
        const blob = Utilities.newBlob(decoded, mimeType, filename);

        // Upload ke folder khusus (buat folder jika belum ada)
        const folderName = 'Sistem_Pengadaan_Files';
        let folder = DriveApp.getFoldersByName(folderName).hasNext()
            ? DriveApp.getFoldersByName(folderName).next()
            : DriveApp.createFolder(folderName);

        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        return {
            success: true,
            fileUrl: file.getUrl(),
            fileId: file.getId()
        };

    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

// ============================================
// INTERNAL FUNCTIONS
// ============================================

/**
 * AUTO-ASSIGN PP dengan algoritma Round-robin + Workload balancing
 */
function autoAssignPP(ss) {
    const sheetUsers = ss.getSheetByName(CONFIG.SHEETS.USERS);
    const usersData = sheetUsers.getDataRange().getValues();
    const headers = usersData[0];

    // Cari semua PP yang aktif
    const ppList = usersData.slice(1)
        .filter(row => row[3] === 'PP' && row[6] === 'AKTIF') // role=PP, status=AKTIF
        .map(row => ({
            email: row[0],
            nama: row[1],
            nip: row[2],
            role: row[3],
            satker: row[4],
            password: row[5],
            status: row[6],
            workload: row[7] || 0 // Kolom H: jumlah penugasan aktif (opsional)
        }));

    if (ppList.length === 0) {
        return { success: false, error: 'No active PP found' };
    }

    // Strategy: Round-robin dengan workload balancing
    // Prioritaskan PP dengan workload terendah
    ppList.sort((a, b) => a.workload - b.workload);

    // Ambil PP dengan workload terendah (jika sama, ambil yang pertama)
    const selectedPP = ppList[0];

    return {
        success: true,
        ...selectedPP
    };
}

/**
 * Update counter workload PP
 */
function updatePPWorkload(ss, emailPP) {
    try {
        const sheetUsers = ss.getSheetByName(CONFIG.SHEETS.USERS);
        const data = sheetUsers.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === emailPP) {
                const currentWorkload = data[i][7] || 0;
                sheetUsers.getRange(i + 1, 8).setValue(currentWorkload + 1); // Kolom H
                break;
            }
        }
    } catch (e) {
        console.error('Error updating workload:', e);
    }
}

/**
 * API untuk mendapatkan daftar PP yang tersedia (untuk Ketua Tim)
 */
function apiGetAvailablePP() {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheetUsers = ss.getSheetByName(CONFIG.SHEETS.USERS);
        const data = sheetUsers.getDataRange().getValues();

        const ppList = data.slice(1)
            .filter(row => row[3] === 'PP' && row[6] === 'AKTIF')
            .map(row => ({
                email: row[0],
                nama: row[1],
                nip: row[2],
                satker: row[4],
                workload: row[7] || 0
            }));

        return { success: true, ppList };
    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

/**
 * API untuk reassign PP (oleh Ketua Tim)
 */
function apiReassignPP(nomorProses, newEmailPP) {
    try {
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

        // 1. Update Pengajuan sheet
        const sheetPengajuan = ss.getSheetByName(CONFIG.SHEETS.PENGAJUAN);
        const dataPengajuan = sheetPengajuan.getDataRange().getValues();

        let rowIndex = -1;
        for (let i = 1; i < dataPengajuan.length; i++) {
            if (dataPengajuan[i][0] === nomorProses) {
                rowIndex = i;
                break;
            }
        }

        if (rowIndex === -1) {
            return { success: false, error: 'Nomor proses tidak ditemukan' };
        }

        // Get new PP data
        const sheetUsers = ss.getSheetByName(CONFIG.SHEETS.USERS);
        const usersData = sheetUsers.getDataRange().getValues();
        const newPP = usersData.find(row => row[0] === newEmailPP && row[3] === 'PP');

        if (!newPP) {
            return { success: false, error: 'PP tidak ditemukan' };
        }

        // Update kolom P dan Q (emailPP dan namaPP)
        sheetPengajuan.getRange(rowIndex + 1, 16).setValue(newEmailPP); // P
        sheetPengajuan.getRange(rowIndex + 1, 17).setValue(newPP[1]);   // Q
        sheetPengajuan.getRange(rowIndex + 1, 18).setValue(new Date()); // R: lastUpdate

        // 2. Update Penugasan_PP sheet (non-aktifkan yang lama, buat yang baru)
        const sheetPenugasan = ss.getSheetByName(CONFIG.SHEETS.PENUGASAN_PP);
        const dataPenugasan = sheetPenugasan.getDataRange().getValues();

        for (let i = 1; i < dataPenugasan.length; i++) {
            if (dataPenugasan[i][1] === nomorProses && dataPenugasan[i][6] === 'AKTIF') {
                sheetPenugasan.getRange(i + 1, 7).setValue('DIGANTI'); // Status lama
                break;
            }
        }

        // Tambah penugasan baru
        sheetPenugasan.appendRow([
            Utilities.getUuid(),
            nomorProses,
            newPP[2], // nip
            newPP[1], // nama
            newEmailPP,
            new Date(),
            'AKTIF'
        ]);

        return {
            success: true,
            message: `Penugasan berhasil diubah ke ${newPP[1]}`,
            newPP: {
                nama: newPP[1],
                email: newEmailPP,
                nip: newPP[2]
            }
        };

    } catch (e) {
        return { success: false, error: e.toString() };
    }
}

// ============================================
// WEB APP ENTRY POINT
// ============================================
function doGet(e) {
    return HtmlService.createHtmlOutputFromFile('Dashboard_PPK_Full')
        .setTitle('Sistem Pengadaan - Dashboard PPK')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
    const action = e.parameter.action;
    const data = JSON.parse(e.postData.contents || '{}');

    switch (action) {
        case 'getDashboardData':
            return ContentService.createTextOutput(JSON.stringify(apiGetDashboardData(data.emailPPK)));
        case 'getNextNomorProses':
            return ContentService.createTextOutput(JSON.stringify(apiGetNextNomorProses()));
        case 'submitPengajuan':
            return ContentService.createTextOutput(JSON.stringify(apiSubmitPengajuan(data)));
        case 'uploadFile':
            return ContentService.createTextOutput(JSON.stringify(apiUploadFile(data.base64Data, data.filename, data.mimeType)));
        case 'getAvailablePP':
            return ContentService.createTextOutput(JSON.stringify(apiGetAvailablePP()));
        case 'reassignPP':
            return ContentService.createTextOutput(JSON.stringify(apiReassignPP(data.nomorProses, data.newEmailPP)));
        default:
            return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }));
    }
}