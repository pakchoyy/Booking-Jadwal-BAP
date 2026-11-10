/**
 * Pengisi jadwal otomatis untuk tab JADWAL BAP.
 * Tempel sebagai file baru di proyek Apps Script yang sama dengan Code.gs
 * (memakai getSheet_, normTanggal_, HARI, TZ dari Code.gs).
 *
 * Cara pakai:
 *   1. Sesuaikan POLA_LAYANAN dan HARI_LIBUR di bawah.
 *   2. Jalankan isiJadwalOtomatis sekali dari editor untuk mencoba.
 *   3. Jalankan pasangTriggerBulanan sekali supaya terisi otomatis tiap bulan.
 */

// ====== PENGATURAN ======
// hari: 0=Minggu, 1=Senin, ... 6=Sabtu. Nama lokasi harus sama persis dengan LOKASI di halaman.
const POLA_LAYANAN = [
  { lokasi: 'Kantor Imigrasi Tegal',   hari: [1, 2, 3, 4, 5], kuota: 8 },
  { lokasi: 'Beranda Imigrasi Brebes', hari: [1, 4],          kuota: 4 }
];

// Tanggal libur nasional / cuti bersama (yyyy-MM-dd). Tanggal ini dilewati.
const HARI_LIBUR = [
  // '2026-12-25',
];

// Isi jadwal untuk sekian hari ke depan setiap dijalankan
const ISI_HARI_KE_DEPAN = 60;

// ====== FUNGSI UTAMA ======
function isiJadwalOtomatis() {
  const sh = getSheet_(SHEET_JADWAL);
  const last = sh.getLastRow();
  const data = last >= 2 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];

  // Kombinasi tanggal|lokasi yang sudah ada, dan nomor ID terakhir
  const ada = {};
  let maxId = 0;
  data.forEach(r => {
    const tgl = normTanggal_(r[1]);
    if (tgl) ada[tgl + '|' + String(r[3]).trim().toLowerCase()] = true;
    const n = parseInt(String(r[0]).replace(/\D/g, ''), 10);
    if (n > maxId) maxId = n;
  });

  const baru = [];
  const now = new Date();
  for (let i = 0; i <= ISI_HARI_KE_DEPAN; i++) {
    const dt = new Date(now.getFullYear(), now.getMonth(), now.getDate() + i);
    const iso = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    if (HARI_LIBUR.indexOf(iso) !== -1) continue;

    POLA_LAYANAN.forEach(p => {
      if (p.hari.indexOf(dt.getDay()) === -1) return;
      if (ada[iso + '|' + p.lokasi.toLowerCase()]) return;
      maxId++;
      baru.push([('J' + ('00' + maxId).slice(-3)), dt, HARI[dt.getDay()], p.lokasi, p.kuota]);
    });
  }

  if (!baru.length) return 'Tidak ada jadwal baru.';

  const start = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(start, 1, baru.length, 5).setValues(baru);
  sh.getRange(start, 2, baru.length, 1).setNumberFormat('yyyy-mm-dd');
  for (let k = 0; k < baru.length; k++) {
    const r = start + k;
    sh.getRange(r, 6).setFormula(
      "=COUNTIFS('DATA PEMOHON'!$G:$G,B" + r + ",'DATA PEMOHON'!$F:$F,D" + r +
      ",'DATA PEMOHON'!$I:$I,\"<>Dibatalkan\")");
    sh.getRange(r, 7).setFormula('=E' + r + '-F' + r);
  }
  return 'Menambah ' + baru.length + ' jadwal.';
}

// ====== JALANKAN SEKALI: trigger tanggal 1 tiap bulan, sekitar jam 01.00 ======
function pasangTriggerBulanan() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'isiJadwalOtomatis') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('isiJadwalOtomatis')
    .timeBased().onMonthDay(1).atHour(1).create();
}
