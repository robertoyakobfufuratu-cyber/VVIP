/**
 * SISTEM VERIFIKASI KARTU VIP — Google Apps Script
 * Spreadsheet sheets required: Pelanggan, Pemakaian
 * Deploy as Web app: Execute as Me; access: Anyone with the link (or your organization only).
 * Protect the spreadsheet and share the web-app URL only with authorized staff.
 */
const CFG = {
  CUSTOMER_SHEET: 'Pelanggan',
  LOG_SHEET: 'Pemakaian',
  TIMEZONE: 'Asia/Jakarta',
  DISCOUNT_RATE: 0.05
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Verifikasi Kartu VIP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function verifyAndRedeem(cardId, amount, staff) {
  cardId = String(cardId || '').trim().toUpperCase();
  staff = String(staff || '').trim();
  amount = Number(amount);
  if (!/^VIP-[A-Z0-9-]+$/.test(cardId)) return {ok:false, message:'Format ID kartu tidak valid.'};
  if (!Number.isFinite(amount) || amount <= 0) return {ok:false, message:'Masukkan nilai transaksi yang valid.'};
  if (!staff) return {ok:false, message:'Nama petugas wajib diisi.'};

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customers = ss.getSheetByName(CFG.CUSTOMER_SHEET);
    const log = ss.getSheetByName(CFG.LOG_SHEET);
    if (!customers || !log) throw new Error('Sheet Pelanggan atau Pemakaian tidak ditemukan.');

    const rows = customers.getDataRange().getValues();
    let customer = null;
    for (let i=1; i<rows.length; i++) {
      if (String(rows[i][0]).trim().toUpperCase() === cardId) {
        customer = {
          row: i+1, id: cardId, name: String(rows[i][1] || ''),
          activation: rows[i][3], expiry: rows[i][4],
          manualStatus: String(rows[i][5] || 'AKTIF').trim().toUpperCase()
        };
        break;
      }
    }
    if (!customer) return {ok:false, message:'Kartu tidak terdaftar.'};
    if (!customer.name) return {ok:false, message:'Nama pelanggan belum diisi; hubungi admin.'};
    if (customer.manualStatus !== 'AKTIF') return {ok:false, message:'Kartu dinonaktifkan oleh admin.'};
    if (!(customer.activation instanceof Date) || isNaN(customer.activation.getTime())) {
      return {ok:false, message:'Tanggal aktivasi belum diisi.'};
    }
    if (!(customer.expiry instanceof Date) || isNaN(customer.expiry.getTime())) {
      return {ok:false, message:'Tanggal kedaluwarsa belum valid.'};
    }

    const today = Utilities.formatDate(new Date(), CFG.TIMEZONE, 'yyyy-MM-dd');
    const start = Utilities.formatDate(customer.activation, CFG.TIMEZONE, 'yyyy-MM-dd');
    const end = Utilities.formatDate(customer.expiry, CFG.TIMEZONE, 'yyyy-MM-dd');
    if (today < start) return {ok:false, message:'Kartu belum memasuki masa aktif.'};
    if (today > end) return {ok:false, message:'Kartu sudah kedaluwarsa.'};

    const logLast = log.getLastRow();
    if (logLast >= 2) {
      const logRows = log.getRange(2,1,logLast-1,10).getValues();
      for (let i=0; i<logRows.length; i++) {
        const row = logRows[i];
        const existingId = String(row[2] || '').trim().toUpperCase();
        const status = String(row[8] || '').trim().toUpperCase();
        if (existingId === cardId && status === 'BERHASIL' && row[0] instanceof Date) {
          const usedDate = Utilities.formatDate(row[0], CFG.TIMEZONE, 'yyyy-MM-dd');
          if (usedDate === today) return {ok:false, message:'Diskon kartu ini sudah digunakan hari ini.'};
        }
      }
    }

    const discount = Math.round(amount * CFG.DISCOUNT_RATE);
    const now = new Date();
    // Simpan tanggal tanpa komponen jam pada kolom B agar laporan harian bekerja.
    const dateOnly = new Date(
      Number(Utilities.formatDate(now, CFG.TIMEZONE, 'yyyy')),
      Number(Utilities.formatDate(now, CFG.TIMEZONE, 'MM')) - 1,
      Number(Utilities.formatDate(now, CFG.TIMEZONE, 'dd'))
    );
    log.appendRow([now, dateOnly, cardId, customer.name, amount, discount, amount-discount, staff, 'BERHASIL', 'Dicatat oleh verifier']);
    SpreadsheetApp.flush();
    return {
      ok:true, message:'VERIFIKASI BERHASIL — diskon 5% diberikan.',
      name:customer.name, cardId:cardId, amount:amount, discount:discount,
      total:amount-discount, expires:Utilities.formatDate(customer.expiry, CFG.TIMEZONE, 'dd-MM-yyyy')
    };
  } finally {
    lock.releaseLock();
  }
}
