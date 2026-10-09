// ==============================================================
// 3. FUNGSI DASHBOARD PRODUKSI (DARI FILE SPREADSHEET LAIN)
// ==============================================================
function getProduksiData() {
  // Buka file spreadsheet eksternal berdasarkan ID
  const ssId = '1byWyfw8quqLMySGTCjOTXTgbEZecYTgV0mOapt-jRvU';
  try {
    const ssProduksi = SpreadsheetApp.openById(ssId);
    
    // Spesifik mencari Sheet bernama DATA_PROD
    const sheet = ssProduksi.getSheetByName('DATA_PROD');
    if (!sheet) return { error: "Sheet DATA_PROD tidak ditemukan di file target!" };

    const values = sheet.getDataRange().getValues();

    let aggregated = {};
    let availableDates = new Set();

    // Loop dari baris ke-2 (Baris 1 adalah Header)
    for (let i = 1; i < values.length; i++) {
      let r = values[i];
      
      // Catatan Indeks Array: A=0, B=1, F=5, G=6, K=10
      let dateStr = parseDateObj(r[0]); // Kolom A: TGL
      if (!dateStr) continue;

      let cabang = String(r[1] || '').trim().toUpperCase(); // Kolom B: CABANG
      let brand = String(r[10] || '').trim().toUpperCase(); // Kolom K: BRAND
      if (!brand) brand = "UNBRANDED"; 
      
      let counterRaw = r[5]; // Kolom F: COUNTER
      let doffingRaw = r[6]; // Kolom G: DOFFING
      
      // Amankan pembacaan angka desimal/ribuan
      let counter = 0;
      if (typeof counterRaw === 'number') counter = counterRaw;
      else if (counterRaw) counter = parseFloat(counterRaw.toString().replace(/\./g, '').replace(',', '.')) || 0;

      let doffing = 0;
      if (typeof doffingRaw === 'number') doffing = doffingRaw;
      else if (doffingRaw) doffing = parseFloat(doffingRaw.toString().replace(/\./g, '').replace(',', '.')) || 0;

      // --- KONVERSI PCS KE KODI (1 Kodi = 20 Pcs) ---
      let counterKodi = counter / 20;
      let doffingKodi = doffing / 20;

      availableDates.add(dateStr);

      // Kita kelompokkan data per Tanggal + Cabang + Brand
      let key = dateStr + '|' + cabang + '|' + brand;
      if (!aggregated[key]) {
          aggregated[key] = {
              date: dateStr,
              cabang: cabang,
              brand: brand,
              counter: 0,
              doffing: 0
          };
      }
      aggregated[key].counter += counterKodi;
      aggregated[key].doffing += doffingKodi;
    }

    let resultArr = Object.values(aggregated);
    let datesArr = Array.from(availableDates).sort().reverse(); // Terbaru di atas

    return { data: resultArr, availableDates: datesArr };
  } catch(e) {
    // Menangkap error jika izin akses file dibatasi atau ID salah
    return { error: e.toString() };
  }
}