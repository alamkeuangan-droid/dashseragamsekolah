/**
 * ======================================================================
 *  DASHBOARD PPDB — DINAS PENDIDIKAN KOTA MAKASSAR
 *  Google Apps Script Web App
 *
 *  Copyright © 2026 Alamsyah Amin. All rights reserved.
 * ======================================================================
 *
 *  CARA PASANG (SETUP):
 *  1. Buka spreadsheet Anda → Extensions → Apps Script.
 *  2. Buat file "Code.gs" (atau timpa yang ada) → tempel isi file ini.
 *  3. Buat file HTML baru bernama persis "Index" → tempel isi Index.html.
 *  4. Buat file HTML baru bernama persis "LogoScript" → tempel isi
 *     LogoScript.html.
 *  5. Sesuaikan CONFIG di bawah ini (terutama nama sheet) jika berbeda
 *     dengan spreadsheet Anda.
 *  6. Deploy → New deployment → pilih "Web app".
 *     - Execute as: Me
 *     - Who has access: Anyone (atau sesuai kebutuhan Anda)
 *  7. Buka URL Web App yang diberikan.
 *
 *  CATATAN PERFORMA (±28.000 baris data):
 *  - Seluruh data siswa dikirim ke browser HANYA SEKALI saat halaman
 *    dimuat (di dalam getDashboardData). Setelah itu, SEMUA filter di
 *    tab "Data Siswa & Cetak" dikerjakan di sisi browser (JavaScript),
 *    BUKAN dengan membaca ulang spreadsheet — sehingga filter terasa
 *    instan walau datanya besar.
 *  - Hasil agregasi (termasuk daftar siswa) disimpan di CacheService
 *    dalam bentuk "chunk" (potongan <100KB) karena satu nilai cache
 *    Apps Script dibatasi 100KB. Selama cache masih berlaku
 *    (CONFIG.CACHE_SECONDS), pemuatan ulang halaman tidak perlu
 *    membaca ulang seluruh sheet.
 *  - Tombol ↻ "Muat Ulang" di navbar akan memaksa baca ulang dari
 *    spreadsheet (melewati cache).
 * ======================================================================
 */

// ----------------------------------------------------------------------
// KONFIGURASI — SESUAIKAN DI SINI
// ----------------------------------------------------------------------
const CONFIG = {
  // ID spreadsheet sumber data (sudah diisi sesuai link yang diberikan)
  SPREADSHEET_ID: '1LkIceO5gq_Pcs_N3ONxxEGyKM1J2b6R2EpAimpy8Zto',

  // Nama sheet berisi data pendaftar (sheet "hasil potongan").
  // GANTI sesuai nama tab sheet Anda yang sebenarnya.
  SHEET_HASIL: 'Hasil',

  // Nama sheet berisi kuota rombel (sheet "skrombel").
  // GANTI sesuai nama tab sheet Anda yang sebenarnya.
  SHEET_KUOTA: 'skrombel',

  // Nama sheet referensi wilayah (NPSN -> Kelurahan/Kecamatan), kolomnya:
  // NPSN | NAMA SEKOLAH | JENJANG | KELURAHAN | KECAMATAN.
  // GANTI persis sesuai nama tab-nya di spreadsheet Anda (case-sensitive).
  SHEET_DB: 'db',

  // Apakah baris pertama sheet HASIL adalah header? Sheet Anda memakai
  // header (JALUR, STATUS_SEKOLAH, ... waktu), jadi defaultnya true.
  HASIL_HAS_HEADER: true,
  KUOTA_HAS_HEADER: true,
  DB_HAS_HEADER: true,

  APP_TITLE: 'Dashboard SERAGAM — Dinas Pendidikan Kota Makassar',
  APP_SUBTITLE: 'Pemantauan Ukuran seragam jenjang SD SMP',

  // ID file logo di Google Drive (dipakai oleh LogoScript.html)
  LOGO_FILE_ID: '1rEVk72Drtcpiq9zSD1qE06mAWtXJaz2g',

  // Lama cache data (detik). Karena data ±28rb baris cukup berat dibaca
  // dari sheet, nilai yang lebih besar (mis. 600 = 10 menit) akan
  // membuat pemuatan halaman jauh lebih cepat untuk pengguna berikutnya.
  // Tombol ↻ di navbar selalu bisa memaksa baca ulang data terbaru.
  CACHE_SECONDS: 600
};

// Urutan kolom pada sheet HASIL (0 = kolom A), sesuai header:
// JALUR | STATUS_SEKOLAH | JENJANG_PENERAPAN | NPSN | NAMA_SEKOLAH | NAMA
// | NIK | NISN | JENIS_KELAMIN | UK_BAJU | UK_CELANA | waktu
const HASIL_COLS = {
  JALUR: 0,
  STATUS: 1,
  JENJANG: 2,
  NPSN: 3,
  NAMA_SEKOLAH: 4,
  NAMA_SISWA: 5,
  NIK: 6,
  NISN: 7,
  JENIS_KELAMIN: 8,
  UK_BAJU: 9,
  UK_CELANA: 10,
  TIMESTAMP: 11
};

// Urutan kolom pada sheet KUOTA / SKROMBEL (0 = kolom A).
const KUOTA_COLS = {
  NO: 0,
  NPSN: 1,
  NAMA_SEKOLAH: 2,
  KECAMATAN: 3,
  ROMBEL: 4,
  KUOTA: 5
};

// Urutan kolom pada sheet "db" (0 = kolom A), sesuai header:
// NPSN | NAMA SEKOLAH | JENJANG | KELURAHAN | KECAMATAN
// Ini jadi SUMBER UTAMA untuk kecamatan/kelurahan tiap sekolah (dicocokkan
// lewat NPSN), karena lebih lengkap daripada kolom kecamatan di Skrombel.
const DB_COLS = {
  NPSN: 0,
  NAMA_SEKOLAH: 1,
  JENJANG: 2,
  KELURAHAN: 3,
  KECAMATAN: 4
};

// ----------------------------------------------------------------------
// ENTRY POINT WEB APP
// ----------------------------------------------------------------------
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appTitle = CONFIG.APP_TITLE;
  tpl.appSubtitle = CONFIG.APP_SUBTITLE;
  return tpl.evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Menyisipkan file HTML lain (dipakai oleh <?!= include('LogoScript') ?>) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ----------------------------------------------------------------------
// UTIL
// ----------------------------------------------------------------------
function ss_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) {
    throw new Error(
      'Sheet "' + name + '" tidak ditemukan. Periksa CONFIG.SHEET_HASIL / ' +
      'CONFIG.SHEET_KUOTA pada Code.gs agar sesuai nama tab spreadsheet Anda.'
    );
  }
  return sh;
}

function normalize_(v) {
  return String(v == null ? '' : v).trim();
}

/** Kunci "canonical" (huruf besar semua + spasi ganda dirapikan) dipakai
 *  untuk MENYAMAKAN nilai yang sebenarnya sama tapi beda ketikan, mis.
 *  "Zonasi", "ZONASI ", "zonasi" -> semuanya jadi kunci "ZONASI". Tanpa ini,
 *  dropdown filter Sekolah/Jalur bisa "kelihatan benar" tapi tidak match
 *  ke sebagian besar baris karena beda spasi/kapitalisasi antar baris. */
function canonKey_(v) {
  return normalize_(v).toUpperCase().replace(/\s+/g, ' ');
}

/** Dari sekumpulan nilai mentah (boleh banyak variasi ketikan), bikin peta
 *  canonicalKey -> ejaan ASLI yang paling sering muncul (dipakai sbg label
 *  tampilan yang konsisten, sekaligus nilai yang dikirim ke filter). */
function buildRepresentativeMap_(rawValues) {
  const freq = {}; // key -> { rawString: count }
  rawValues.forEach(function (v) {
    const raw = normalize_(v);
    if (!raw) return;
    const key = canonKey_(raw);
    if (!freq[key]) freq[key] = {};
    freq[key][raw] = (freq[key][raw] || 0) + 1;
  });
  const map = {};
  Object.keys(freq).forEach(function (key) {
    const counts = freq[key];
    let best = null, bestCount = -1;
    Object.keys(counts).forEach(function (raw) {
      if (counts[raw] > bestCount) { bestCount = counts[raw]; best = raw; }
    });
    map[key] = best;
  });
  return map;
}

/** Deteksi jenjang dari nama sekolah, dipakai sbg fallback untuk sheet kuota
 *  (sheet Skrombel tidak punya kolom jenjang eksplisit). */
function guessJenjang_(namaSekolah) {
  const s = normalize_(namaSekolah).toUpperCase();
  if (s.indexOf('SMP') !== -1 || s.indexOf('MTS') !== -1) return 'SMP';
  if (s.indexOf('SD') !== -1 || s.indexOf('MI ') !== -1) return 'SD';
  return 'Lainnya';
}

// ----------------------------------------------------------------------
// CACHE BERTINGKAT (CHUNKED) — agar payload besar (>100KB) tetap bisa
// di-cache oleh CacheService, yang membatasi setiap nilai maksimal 100KB.
// ----------------------------------------------------------------------
function cacheStore_(key, obj) {
  try {
    const json = JSON.stringify(obj);
    const chunkSize = 90000; // aman di bawah batas 100KB per nilai cache
    const total = Math.ceil(json.length / chunkSize) || 1;
    const cache = CacheService.getScriptCache();
    const payload = {};
    for (let i = 0; i < total; i++) {
      payload[key + '_c' + i] = json.substring(i * chunkSize, (i + 1) * chunkSize);
    }
    payload[key + '_meta'] = String(total);
    cache.putAll(payload, CONFIG.CACHE_SECONDS);
  } catch (err) {
    // Gagal menyimpan cache tidak fatal — request berikutnya akan
    // membaca ulang langsung dari spreadsheet.
  }
}

function cacheRead_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const metaStr = cache.get(key + '_meta');
    if (!metaStr) return null;
    const total = Number(metaStr);
    const keys = [];
    for (let i = 0; i < total; i++) keys.push(key + '_c' + i);
    const map = cache.getAll(keys);
    let json = '';
    for (let i = 0; i < total; i++) {
      const part = map[key + '_c' + i];
      if (part == null) return null; // ada chunk yang sudah kedaluwarsa
      json += part;
    }
    return JSON.parse(json);
  } catch (err) {
    return null;
  }
}

// ----------------------------------------------------------------------
// PEMBACAAN DATA MENTAH
// ----------------------------------------------------------------------
function readHasilRaw_() {
  const sh = sheet_(CONFIG.SHEET_HASIL);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 12);
  const startRow = CONFIG.HASIL_HAS_HEADER ? 2 : 1;
  if (lastRow < startRow) return [];
  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const namaSiswa = normalize_(row[HASIL_COLS.NAMA_SISWA]);
    const npsn = normalize_(row[HASIL_COLS.NPSN]);
    if (!namaSiswa && !npsn) continue; // lewati baris kosong

    out.push({
      jalur: normalize_(row[HASIL_COLS.JALUR]),
      status: normalize_(row[HASIL_COLS.STATUS]),
      jenjang: normalize_(row[HASIL_COLS.JENJANG]).toUpperCase(),
      npsn: npsn,
      namaSekolah: normalize_(row[HASIL_COLS.NAMA_SEKOLAH]),
      namaSiswa: namaSiswa,
      nik: normalize_(row[HASIL_COLS.NIK]),
      nisn: normalize_(row[HASIL_COLS.NISN]),
      jenisKelamin: normalize_(row[HASIL_COLS.JENIS_KELAMIN]),
      ukBaju: normalize_(row[HASIL_COLS.UK_BAJU]),
      ukCelana: normalize_(row[HASIL_COLS.UK_CELANA])
    });
  }
  return out;
}

function readKuotaRaw_() {
  const sh = sheet_(CONFIG.SHEET_KUOTA);
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 6);
  const startRow = CONFIG.KUOTA_HAS_HEADER ? 2 : 1;
  if (lastRow < startRow) return [];
  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const npsn = normalize_(row[KUOTA_COLS.NPSN]);
    const namaSekolah = normalize_(row[KUOTA_COLS.NAMA_SEKOLAH]);
    if (!npsn && !namaSekolah) continue;

    out.push({
      npsn: npsn,
      namaSekolah: namaSekolah,
      kecamatan: normalize_(row[KUOTA_COLS.KECAMATAN]),
      rombel: Number(row[KUOTA_COLS.ROMBEL]) || 0,
      kuota: Number(row[KUOTA_COLS.KUOTA]) || 0
    });
  }
  return out;
}

/** Baca sheet "db" (referensi NPSN -> Kelurahan/Kecamatan). Sheet ini
 *  opsional: kalau tab-nya belum ada/beda nama, fungsi ini akan
 *  mengembalikan array kosong (tidak bikin seluruh dashboard error),
 *  dan sistem otomatis jatuh ke kecamatan dari sheet Skrombel saja. */
function readDbRaw_() {
  let sh;
  try {
    sh = sheet_(CONFIG.SHEET_DB);
  } catch (err) {
    return [];
  }
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 5);
  const startRow = CONFIG.DB_HAS_HEADER ? 2 : 1;
  if (lastRow < startRow) return [];
  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  const out = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const npsn = normalize_(row[DB_COLS.NPSN]);
    if (!npsn) continue;

    out.push({
      npsn: npsn,
      namaSekolah: normalize_(row[DB_COLS.NAMA_SEKOLAH]),
      jenjang: normalize_(row[DB_COLS.JENJANG]).toUpperCase(),
      kelurahan: normalize_(row[DB_COLS.KELURAHAN]),
      kecamatan: normalize_(row[DB_COLS.KECAMATAN])
    });
  }
  return out;
}

// ----------------------------------------------------------------------
// AGREGASI DATA — DIPANGGIL DARI CLIENT (google.script.run)
// ----------------------------------------------------------------------

/**
 * Mengambil SELURUH data dashboard sekaligus (1 kali round-trip dari
 * client), termasuk daftar siswa mentah (siswaRows) yang akan difilter
 * di sisi browser. Hasilnya di-cache (chunked) sesuai CONFIG.CACHE_SECONDS.
 *
 * Format siswaRows sengaja berupa ARRAY (bukan object) demi ukuran
 * payload yang lebih kecil & lebih cepat untuk ±28.000 baris:
 *   [namaSiswa, jenisKelamin, jenjang, namaSekolah, kecamatan, jalur, ukBaju, ukCelana]
 */
function getDashboardData(forceRefresh) {
  const cacheKey = 'dashboard_data_v4';
  if (CONFIG.CACHE_SECONDS > 0 && !forceRefresh) {
    const cached = cacheRead_(cacheKey);
    if (cached) return cached;
  }

  const hasil = readHasilRaw_();
  const kuota = readKuotaRaw_();
  const db = readDbRaw_();

  // Peta NPSN -> info kuota (kecamatan, kuota, jumlah rombel, nama)
  const kuotaByNpsn = {};
  kuota.forEach(function (k) {
    kuotaByNpsn[k.npsn] = k;
  });

  // Peta NPSN -> info wilayah dari sheet "db" (SUMBER UTAMA kecamatan/
  // kelurahan, karena lebih lengkap daripada kolom kecamatan di Skrombel).
  const dbByNpsn = {};
  db.forEach(function (d) {
    dbByNpsn[d.npsn] = d;
  });

  /** Ambil kecamatan utk sebuah NPSN: utamakan sheet "db", baru Skrombel,
   *  baru 'Tidak Diketahui' kalau dua-duanya tidak punya datanya. */
  function resolveKecamatan_(npsn) {
    const d = dbByNpsn[npsn];
    if (d && d.kecamatan) return d.kecamatan;
    const k = kuotaByNpsn[npsn];
    if (k && k.kecamatan) return k.kecamatan;
    return 'Tidak Diketahui';
  }

  // Peta canonical -> ejaan resmi, dipakai agar filter Sekolah & Jalur di
  // client selalu match (menyatukan variasi ketikan/spasi yang beda-beda).
  const sekolahRepMap = buildRepresentativeMap_(hasil.map(function (h) { return h.namaSekolah; }));
  const jalurRepMap = buildRepresentativeMap_(hasil.map(function (h) { return h.jalur; }));
  function repSekolah_(v) { return sekolahRepMap[canonKey_(v)] || normalize_(v); }
  function repJalur_(v) { return jalurRepMap[canonKey_(v)] || normalize_(v); }

  // -------- 1) Perbandingan jumlah pendaftar vs kuota, per sekolah --------
  const jumlahBySekolah = {}; // npsn -> {count, jenjang, namaSekolah}
  hasil.forEach(function (h) {
    const key = h.npsn || ('NONPSN::' + h.namaSekolah);
    if (!jumlahBySekolah[key]) {
      jumlahBySekolah[key] = { npsn: h.npsn, namaSekolah: h.namaSekolah, jenjang: h.jenjang, count: 0 };
    }
    jumlahBySekolah[key].count++;
  });

  const kuotaSekolah = kuota.map(function (k) {
    const j = jumlahBySekolah[k.npsn];
    const jumlah = j ? j.count : 0;
    const selisih = k.kuota - jumlah;
    let statusLabel, statusKode;
    if (jumlah === 0 && k.kuota > 0) { statusLabel = 'Belum Ada Pendaftar'; statusKode = 'kosong'; }
    else if (jumlah < k.kuota) { statusLabel = 'Kurang (' + selisih + ' slot kosong)'; statusKode = 'kurang'; }
    else if (jumlah === k.kuota) { statusLabel = 'Pas / Penuh'; statusKode = 'pas'; }
    else { statusLabel = 'Lebih (' + Math.abs(selisih) + ' kelebihan)'; statusKode = 'lebih'; }

    return {
      npsn: k.npsn,
      namaSekolah: k.namaSekolah,
      kecamatan: resolveKecamatan_(k.npsn),
      rombel: k.rombel,
      kuota: k.kuota,
      jumlahDaftar: jumlah,
      selisih: selisih,
      persenTerisi: k.kuota > 0 ? Math.round((jumlah / k.kuota) * 1000) / 10 : 0,
      statusKode: statusKode,
      statusLabel: statusLabel,
      jenjang: guessJenjang_(k.namaSekolah)
    };
  });

  // Sekolah yang ada di HASIL tapi tidak ditemukan di data kuota
  const npsnTanpaKuota = [];
  Object.keys(jumlahBySekolah).forEach(function (key) {
    const j = jumlahBySekolah[key];
    if (!kuotaByNpsn[j.npsn]) {
      npsnTanpaKuota.push({ npsn: j.npsn, namaSekolah: j.namaSekolah, jenjang: j.jenjang, jumlahDaftar: j.count });
    }
  });

  // -------- 2) Analisis per kecamatan --------
  const kecHasil = {}; // kecamatan -> count
  hasil.forEach(function (h) {
    const kecamatan = resolveKecamatan_(h.npsn);
    kecHasil[kecamatan] = (kecHasil[kecamatan] || 0) + 1;
  });

  const kecKuota = {}; // kecamatan -> total kuota
  kuota.forEach(function (k) {
    const kecamatan = resolveKecamatan_(k.npsn);
    kecKuota[kecamatan] = (kecKuota[kecamatan] || 0) + k.kuota;
  });

  const semuaKecamatan = Array.from(new Set(Object.keys(kecHasil).concat(Object.keys(kecKuota)))).sort();
  const analisisKecamatan = semuaKecamatan.map(function (kec) {
    return {
      kecamatan: kec,
      jumlahPendaftar: kecHasil[kec] || 0,
      totalKuota: kecKuota[kec] || 0
    };
  });

  // -------- 3) Total SD / SMP --------
  const totalJenjangHasil = {};
  hasil.forEach(function (h) {
    const j = h.jenjang || 'Lainnya';
    totalJenjangHasil[j] = (totalJenjangHasil[j] || 0) + 1;
  });

  const totalJenjangKuota = {};
  kuota.forEach(function (k) {
    const j = guessJenjang_(k.namaSekolah);
    totalJenjangKuota[j] = (totalJenjangKuota[j] || 0) + k.kuota;
  });

  // -------- 4) Ringkasan umum --------
  const totalSekolahKuota = kuota.length;
  const totalKuotaKeseluruhan = kuota.reduce(function (a, k) { return a + k.kuota; }, 0);
  const totalPendaftar = hasil.length;
  const totalSelisih = totalKuotaKeseluruhan - totalPendaftar;

  const jalurCount = {};
  hasil.forEach(function (h) {
    const j = repJalur_(h.jalur) || 'Tidak Diketahui';
    jalurCount[j] = (jalurCount[j] || 0) + 1;
  });

  // -------- 5) Daftar sekolah / kecamatan / jenjang / jalur (utk filter) --------
  // Sekolah & Jalur memakai ejaan "representative" (canonical) yang sama
  // dengan yang dipakai di siswaRows di bawah, supaya filter DIJAMIN match.
  const daftarSekolah = Array.from(new Set(hasil.map(function (h) { return repSekolah_(h.namaSekolah); }).filter(Boolean))).sort();
  const daftarKecamatan = semuaKecamatan;
  const daftarJenjang = Array.from(new Set(hasil.map(function (h) { return h.jenjang; }).filter(Boolean))).sort();
  const daftarJalur = Object.keys(jalurCount).sort();

  // -------- 6) Daftar siswa (format array ringkas, utk difilter di client) --------
  const siswaRows = hasil.map(function (h) {
    return [
      h.namaSiswa,
      h.jenisKelamin,
      h.jenjang,
      repSekolah_(h.namaSekolah),
      resolveKecamatan_(h.npsn),
      repJalur_(h.jalur),
      h.ukBaju,
      h.ukCelana
    ];
  });

  // -------- 7) Deteksi siswa dengan NIK ganda/duplikat pada sheet HASIL --------
  // Baris tanpa NIK (kosong) tidak dianggap "duplikat" satu sama lain.
  const nikGroups_ = {}; // nik -> array baris hasil dengan NIK tsb
  hasil.forEach(function (h) {
    const nik = normalize_(h.nik);
    if (!nik) return;
    if (!nikGroups_[nik]) nikGroups_[nik] = [];
    nikGroups_[nik].push(h);
  });

  const duplikatNik = Object.keys(nikGroups_)
    .filter(function (nik) { return nikGroups_[nik].length > 1; })
    .sort()
    .map(function (nik) {
      const rows = nikGroups_[nik];
      return {
        nik: nik,
        jumlah: rows.length,
        siswa: rows.map(function (h) {
          return {
            namaSiswa: h.namaSiswa,
            jenisKelamin: h.jenisKelamin,
            jenjang: h.jenjang,
            namaSekolah: repSekolah_(h.namaSekolah),
            kecamatan: resolveKecamatan_(h.npsn),
            jalur: repJalur_(h.jalur),
            nisn: h.nisn,
            ukBaju: h.ukBaju,
            ukCelana: h.ukCelana,
            status: h.status
          };
        })
      };
    });

  const totalSiswaDuplikat = duplikatNik.reduce(function (a, g) { return a + g.jumlah; }, 0);

  const result = {
    generatedAt: new Date().toISOString(),
    ringkasan: {
      totalPendaftar: totalPendaftar,
      totalSekolahKuota: totalSekolahKuota,
      totalKuotaKeseluruhan: totalKuotaKeseluruhan,
      totalSelisih: totalSelisih,
      jalurCount: jalurCount
    },
    kuotaSekolah: kuotaSekolah,
    npsnTanpaKuota: npsnTanpaKuota,
    analisisKecamatan: analisisKecamatan,
    totalJenjang: {
      hasil: totalJenjangHasil,
      kuota: totalJenjangKuota
    },
    filterOptions: {
      sekolah: daftarSekolah,
      kecamatan: daftarKecamatan,
      jenjang: daftarJenjang,
      jalur: daftarJalur
    },
    // Urutan field tiap baris: lihat komentar fungsi di atas.
    siswaFields: ['namaSiswa', 'jenisKelamin', 'jenjang', 'namaSekolah', 'kecamatan', 'jalur', 'ukBaju', 'ukCelana'],
    siswaRows: siswaRows,
    // Daftar NIK yang muncul lebih dari sekali pada sheet Hasil, beserta
    // seluruh baris siswa yang memakai NIK tsb (utk tab "Duplikat NIK").
    duplikatNik: duplikatNik,
    ringkasanDuplikat: {
      totalNikDuplikat: duplikatNik.length,
      totalSiswaDuplikat: totalSiswaDuplikat
    }
  };

  if (CONFIG.CACHE_SECONDS > 0) {
    cacheStore_(cacheKey, result);
  }

  return result;
}

/**
 * OPSIONAL / LEGACY: mengembalikan daftar siswa terfilter langsung dari
 * server. Tidak lagi dipakai oleh UI (yang sekarang memfilter di browser
 * dari data yang sudah dikirim oleh getDashboardData, agar tidak perlu
 * membaca ulang ±28rb baris setiap kali filter berubah). Fungsi ini
 * disimpan sebagai referensi/API cadangan bila suatu saat dibutuhkan
 * pemfilteran langsung dari server (mis. untuk integrasi lain).
 */
function getSiswaList(filters) {
  filters = filters || {};
  const hasil = readHasilRaw_();
  const kuota = readKuotaRaw_();
  const db = readDbRaw_();
  const kuotaByNpsn = {};
  kuota.forEach(function (k) { kuotaByNpsn[k.npsn] = k; });
  const dbByNpsn = {};
  db.forEach(function (d) { dbByNpsn[d.npsn] = d; });
  function resolveKecamatanLegacy_(npsn) {
    const d = dbByNpsn[npsn];
    if (d && d.kecamatan) return d.kecamatan;
    const k = kuotaByNpsn[npsn];
    if (k && k.kecamatan) return k.kecamatan;
    return 'Tidak Diketahui';
  }

  let list = hasil.map(function (h) {
    return {
      namaSiswa: h.namaSiswa,
      jenisKelamin: h.jenisKelamin,
      jenjang: h.jenjang,
      namaSekolah: h.namaSekolah,
      npsn: h.npsn,
      kecamatan: resolveKecamatanLegacy_(h.npsn),
      jalur: h.jalur,
      nisn: h.nisn,
      ukBaju: h.ukBaju,
      ukCelana: h.ukCelana
    };
  });

  if (filters.jenjang) list = list.filter(function (s) { return s.jenjang === filters.jenjang; });
  if (filters.kecamatan) list = list.filter(function (s) { return s.kecamatan === filters.kecamatan; });
  if (filters.sekolah) list = list.filter(function (s) { return s.namaSekolah === filters.sekolah; });
  if (filters.jalur) list = list.filter(function (s) { return s.jalur === filters.jalur; });

  list.sort(function (a, b) {
    return a.namaSekolah.localeCompare(b.namaSekolah) || a.namaSiswa.localeCompare(b.namaSiswa);
  });

  return list;
}

// ----------------------------------------------------------------------
// LOGO (dipanggil dari LogoScript.html sbg fallback bila URL publik gagal)
// ----------------------------------------------------------------------
function getLogoDataUri() {
  try {
    const file = DriveApp.getFileById(CONFIG.LOGO_FILE_ID);
    const blob = file.getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (err) {
    return '';
  }
}