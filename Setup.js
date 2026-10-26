/**
 * ZettBOT Apps Script Assistant - Setup & Safe Migration
 * Aplikasi: Wedding Planner & Savings Manager (Hatami & Naning)
 * File: setup.gs
 */

function setupDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timezone = "Asia/Jakarta";
    
    var tabunganSheet = ss.getSheetByName("Tabungan");
    var tabunganHeaders = ["ID_Transaksi", "Tanggal", "Pihak", "Tipe", "Kategori", "Nominal", "Catatan"];
    
    if (!tabunganSheet) {
      tabunganSheet = ss.insertSheet("Tabungan");
      tabunganSheet.getRange(1, 1, 1, tabunganHeaders.length).setValues([tabunganHeaders]);
      tabunganSheet.getRange(1, 1, 1, tabunganHeaders.length)
        .setFontWeight("bold")
        .setBackground("#151b26")
        .setFontColor("#00ffaa");
      tabunganSheet.setFrozenRows(1);
    } else {
      var currentHeaders = tabunganSheet.getRange(1, 1, 1, tabunganHeaders.length).getValues()[0];
      for (var i = 0; i < tabunganHeaders.length; i++) {
        if (!currentHeaders[i]) {
          tabunganSheet.getRange(1, i + 1).setValue(tabunganHeaders[i]);
        }
      }
    }

    var plannerSheet = ss.getSheetByName("Planner");
    var plannerHeaders = ["ID_Vendor", "Tahapan", "Kategori", "Nama_Vendor", "Estimasi_Biaya", "Status_Bayar", "Catatan"];
    
    if (!plannerSheet) {
      plannerSheet = ss.insertSheet("Planner");
      plannerSheet.getRange(1, 1, 1, plannerHeaders.length).setValues([plannerHeaders]);
      plannerSheet.getRange(1, 1, 1, plannerHeaders.length)
        .setFontWeight("bold")
        .setBackground("#151b26")
        .setFontColor("#00ffaa");
      plannerSheet.setFrozenRows(1);
    } else {
      var currentPlannerHeaders = plannerSheet.getRange(1, 1, 1, plannerHeaders.length).getValues()[0];
      for (var j = 0; j < plannerHeaders.length; j++) {
        if (!currentPlannerHeaders[j]) {
          plannerSheet.getRange(1, j + 1).setValue(plannerHeaders[j]);
        }
      }
    }

    var seserahanSheet = ss.getSheetByName("Seserahan");
    var seserahanHeaders = ["ID_Barang", "Nama_Barang", "Kategori", "Estimasi_Harga", "Status", "Link_Produk", "Catatan"];
    
    if (!seserahanSheet) {
      seserahanSheet = ss.insertSheet("Seserahan");
      seserahanSheet.getRange(1, 1, 1, seserahanHeaders.length).setValues([seserahanHeaders]);
      seserahanSheet.getRange(1, 1, 1, seserahanHeaders.length)
        .setFontWeight("bold")
        .setBackground("#151b26")
        .setFontColor("#00ffaa");
      seserahanSheet.setFrozenRows(1);
    } else {
      var currentSeserahanHeaders = seserahanSheet.getRange(1, 1, 1, seserahanHeaders.length).getValues()[0];
      for (var k = 0; k < seserahanHeaders.length; k++) {
        if (!currentSeserahanHeaders[k]) {
          seserahanSheet.getRange(1, k + 1).setValue(seserahanHeaders[k]);
        }
      }
    }

    var kuaSheet = ss.getSheetByName("PemberkasanKUA");
    var kuaHeaders = ["ID_Berkas", "Nama_Berkas", "Keterangan", "Status_Pria", "Status_Wanita", "Berlaku_Pria", "Berlaku_Wanita"];
    
    if (!kuaSheet) {
      kuaSheet = ss.insertSheet("PemberkasanKUA");
      kuaSheet.getRange(1, 1, 1, kuaHeaders.length).setValues([kuaHeaders]);
      kuaSheet.getRange(1, 1, 1, kuaHeaders.length)
        .setFontWeight("bold")
        .setBackground("#151b26")
        .setFontColor("#00ffaa");
      kuaSheet.setFrozenRows(1);
    } else {
      var currentKuaHeaders = kuaSheet.getRange(1, 1, 1, kuaHeaders.length).getValues()[0];
      for (var m = 0; m < kuaHeaders.length; m++) {
        if (!currentKuaHeaders[m]) {
          kuaSheet.getRange(1, m + 1).setValue(kuaHeaders[m]);
        }
      }
    }

    var tamuSheet = ss.getSheetByName("Tamu");
    var tamuHeaders = ["ID_Tamu", "Nama_Tamu", "Pihak", "Kategori_Relasi", "Jenis_Undangan", "Jumlah_Pax_Undangan", "Status_Konfirmasi", "No_HP", "Catatan"];
    
    if (!tamuSheet) {
      tamuSheet = ss.insertSheet("Tamu");
      tamuSheet.getRange(1, 1, 1, tamuHeaders.length).setValues([tamuHeaders]);
      tamuSheet.getRange(1, 1, 1, tamuHeaders.length)
        .setFontWeight("bold")
        .setBackground("#151b26")
        .setFontColor("#00ffaa");
      tamuSheet.setFrozenRows(1);
    } else {
      var currentTamuHeaders = tamuSheet.getRange(1, 1, 1, tamuHeaders.length).getValues()[0];
      for (var n = 0; n < tamuHeaders.length; n++) {
        if (!currentTamuHeaders[n]) {
          tamuSheet.getRange(1, n + 1).setValue(tamuHeaders[n]);
        }
      }
    }

    // Tabungan sheet left empty; user will add rows via addTabunganTransaction or manually.
    // No initial data inserted.
    // Ensure sheet has headers only.
    // End of block.


    // Planner sheet left empty; user will add vendor entries via addPlannerVendor or manually.
    // No initial data inserted.
    // Ensure sheet has headers only.
    // End of block.


    if (seserahanSheet.getLastRow() <= 1) {
      var initialSeserahanData = [
        ["SSR-0001", "Skincare & BodyCare Set", "Skincare & Kosmetik", 1250000, "Wishlist", "https://shopee.co.id/", "Paket perawatan seserahan"],
        ["SSR-0002", "Sepatu High Heels Premium", "Sepatu & Tas", 850000, "Sudah Dibeli", "https://vt.tiktok.com/", "Ukuran 38 warna cream Nude"],
        ["SSR-0003", "Mukena Sutra & Sajadah", "Alat Sholat", 650000, "Sudah Dibeli", "https://shopee.co.id/", "Warna putih aksen renda gold"],
        ["SSR-0004", "Set Gamis & Hijab Syari", "Pakaian", 950000, "Wishlist", "https://shopee.co.id/", "Bahan silk premium"]
      ];
      seserahanSheet.getRange(2, 1, initialSeserahanData.length, seserahanHeaders.length).setValues(initialSeserahanData);
    }

    if (kuaSheet.getLastRow() <= 1) {
      var initialKuaData = [
        ["KUA-0001", "KTP Asli dan FC 2 Lembar", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0002", "KK Asli dan FC 2 Lembar", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0003", "Ijazah FC 2 Lembar", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0004", "Akte FC 2 Lembar", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0005", "Pas Foto 3x4 2 Lembar", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0006", "FC KTP Ortu", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0007", "Surat Ket Belum Menikah", "Kasih Pak RT", "N/A", "Belum", "Tidak", "Ya"],
        ["KUA-0008", "Surat Pengantar RT/RW", "Kasih Pak RT", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0009", "Surat Ket Sehat Puskesmas", "Puskesmas", "Belum", "Belum", "Ya", "Ya"],
        ["KUA-0010", "Vaksin Tetanus (TT)", "Puskesmas", "N/A", "Belum", "Tidak", "Ya"],
        ["KUA-0011", "Surat Numpang Nikah", "Kasih Pak RT", "Belum", "N/A", "Ya", "Tidak"],
        ["KUA-0012", "Map Kuning", "Kasih Pak RT", "Belum", "N/A", "Ya", "Tidak"]
      ];
      kuaSheet.getRange(2, 1, initialKuaData.length, kuaHeaders.length).setValues(initialKuaData);
    }

    if (tamuSheet.getLastRow() <= 1) {
      var initialTamuData = [
        ["TMU-0001", "Bpk. Hendra & Istri", "Orang Tua", "Keluarga", "Fisik", 2, "Hadir", "081298765432", "Keluarga besar Banyuasin"],
        ["TMU-0002", "Dimas Anggara", "Hatami", "Teman", "Digital", 2, "Terkirim", "085678901234", "Teman SMA Hatami"],
        ["TMU-0003", "Siti Rahmawati", "Naning", "Teman", "Digital", 1, "Hadir", "087712345678", "Rekan Sejawat Perawat"],
        ["TMU-0004", "Dr. Ahmad Fadhil", "Orang Tua", "VIP/Rekan Kerja", "Fisik", 2, "Belum Dikirim", "081122334455", "Dokter Kepala RS"]
      ];
      tamuSheet.getRange(2, 1, initialTamuData.length, tamuHeaders.length).setValues(initialTamuData);
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Database dan Safe Migrate berhasil dikonfigurasi!" };
  } catch (error) {
    return { success: false, message: "Gagal mendesain database: " + error.toString() };
  }
}