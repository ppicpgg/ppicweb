function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Textile Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sanitizeName(name) {
  if (!name) return "";
  return String(name).trim().toUpperCase().replace(/\s+/g, ' ');
}

// 🛡️ PERBAIKAN 1: Parser Tanggal Super Kebal (Anti-Error Format Excel & String)
function parseDateObj(d) {
  if (d === null || d === undefined || d === '') return "";
  if (d instanceof Date) {
    let m = ('0' + (d.getMonth()+1)).slice(-2);
    let day = ('0' + d.getDate()).slice(-2);
    return `${d.getFullYear()}-${m}-${day}`;
  }
  if (typeof d === 'number') {
    // Tangani format serial date (Excel/Sheets bug)
    let date = new Date(Math.round((d - 25569) * 86400 * 1000));
    let m = ('0' + (date.getUTCMonth()+1)).slice(-2);
    let day = ('0' + date.getUTCDate()).slice(-2);
    return `${date.getUTCFullYear()}-${m}-${day}`;
  }
  if (typeof d === 'string') {
    let str = d.trim();
    if (!str) return "";
    let parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      let m = ('0' + (parsed.getMonth()+1)).slice(-2);
      let day = ('0' + parsed.getDate()).slice(-2);
      return `${parsed.getFullYear()}-${m}-${day}`;
    }
    str = str.replace(/\s/g, '');
    if (str.includes('/')) {
      let p = str.split('/');
      if (p.length === 3) {
         if (p[2].length === 4) return `${p[2]}-${('0'+p[1]).slice(-2)}-${('0'+p[0]).slice(-2)}`;
         if (p[0].length === 4) return `${p[0]}-${('0'+p[1]).slice(-2)}-${('0'+p[2]).slice(-2)}`;
      }
    } else if (str.includes('-')) {
      let p = str.split('-');
      if (p.length === 3) {
         if (p[2].length === 4) return `${p[2]}-${('0'+p[1]).slice(-2)}-${('0'+p[0]).slice(-2)}`;
         if (p[0].length === 4) return `${p[0]}-${('0'+p[1]).slice(-2)}-${('0'+p[2]).slice(-2)}`;
      }
    }
    return str.split('T')[0];
  }
  return "";
}

// ==============================================================
// 1. FUNGSI DASHBOARD STOCK BENANG
// ==============================================================
function getStockData(targetDate = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetKbh = ss.getSheetByName('MASTER_KBH');
  let dataKbh = {};
  if (sheetKbh) {
    const kbhValues = sheetKbh.getDataRange().getValues();
    for (let i = 1; i < kbhValues.length; i++) {
       let rawName = kbhValues[i][0]; 
       let avg = parseFloat(String(kbhValues[i][13])) || 0; 
       let cleanName = sanitizeName(rawName);
       if (cleanName) dataKbh[cleanName] = avg; 
    }
  }

  const sheetBng = ss.getSheetByName('DATA_BNG');
  let dataBng = [];
  let availableDates = []; 
  let latestDate = null;
  let trendData = []; 
  
  if (sheetBng) {
    const values = sheetBng.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    const tglIdx = headers.indexOf('TGL');
    const sAkhirIdx = headers.indexOf('S.AKHIR');
    
    let datesSet = new Set();
    for (let i = 1; i < values.length; i++) {
      let dStr = parseDateObj(values[i][tglIdx]);
      if(dStr) datesSet.add(dStr);
    }
    
    availableDates = Array.from(datesSet).sort().reverse(); 
    latestDate = availableDates[0] || null; 
    let filterDate = targetDate ? targetDate : latestDate;

    let startIndex = availableDates.indexOf(filterDate);
    if (startIndex === -1) startIndex = 0;
    let last7Dates = availableDates.slice(startIndex, startIndex + 7);
    
    let trendObj = {};
    last7Dates.forEach(d => trendObj[d] = 0);

    for (let i = 1; i < values.length; i++) {
      let rowDateStr = parseDateObj(values[i][tglIdx]);
      if (rowDateStr === filterDate) {
        let row = {};
        for (let j = 0; j < headers.length; j++) {
          let cellValue = values[i][j];
          let colName = headers[j];
          if (cellValue instanceof Date) {
             let cm = ('0' + (cellValue.getMonth()+1)).slice(-2);
             let cd = ('0' + cellValue.getDate()).slice(-2);
             cellValue = `${cellValue.getFullYear()}-${cm}-${cd}`;
          }
          if (colName === 'NO BENANG') cellValue = sanitizeName(cellValue);
          row[colName] = cellValue;
        }
        dataBng.push(row);
      }

      if (last7Dates.includes(rowDateStr)) {
        let stockVal = parseFloat(String(values[i][sAkhirIdx])) || 0;
        trendObj[rowDateStr] += stockVal;
      }
    }
    trendData = last7Dates.map(d => ({ date: d, total: trendObj[d] })).reverse();
    latestDate = filterDate; 
  }
  return { stock: dataBng, kbh: dataKbh, activeStockDate: latestDate, availableDates: availableDates, trendData: trendData };
}

// ==============================================================
// 2. FUNGSI ANALISIS DYEING & KANBAN (MAX DURABILITY)
// ==============================================================
function parseSmartQty(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let str = String(val).trim();
  if (str.includes(',') && !str.includes('.')) {
    str = str.replace(',', '.');
  } else if (str.includes('.') && str.includes(',')) {
    str = str.replace(/\./g, '').replace(',', '.');
  }
  let num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function getDyeingData(targetDate = null, modeTampilan = 'Harian') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetMc = ss.getSheetByName('MASTER_MC');
  let dataMc = [];
  let dataMcCaps = {}; 
  if (sheetMc) {
    const mcValues = sheetMc.getDataRange().getValues();
    for (let i = 1; i < mcValues.length; i++) {
       let mcName = mcValues[i][0];
       if (mcName) {
           let mKey = sanitizeName(mcName);
           dataMc.push(mKey);
           dataMcCaps[mKey] = [
             parseSmartQty(mcValues[i][1]), 
             parseSmartQty(mcValues[i][2]), 
             parseSmartQty(mcValues[i][3])
           ];
       }
    }
  }

  let sheetKikc = ss.getSheetByName('MASTER_KIKC');
  let dataKikc = {};
  if (sheetKikc) {
    const lastRow = sheetKikc.getLastRow();
    if (lastRow > 1) {
      const kikcVals = sheetKikc.getRange(1, 1, lastRow, 8).getValues();
      for (let i = 1; i < kikcVals.length; i++) {
         let rawKikc = kikcVals[i][1]; 
         let ket = kikcVals[i][7];     
         let cleanKikc = sanitizeName(rawKikc);
         if (cleanKikc) dataKikc[cleanKikc] = ket ? String(ket).trim() : "-";
      }
    }
  }

  let tHarian = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0, 9:0, 10:0, 11:0, 12:0};
  let tBulanan = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, 8:0, 9:0, 10:0, 11:0, 12:0};
  
  let sheetTarget = ss.getSheetByName('TARGET_DYG');
  if (sheetTarget) {
    const tVals = sheetTarget.getDataRange().getValues();
    for(let r=0; r < Math.min(20, tVals.length); r++) {
       let header = tVals[r][0] ? String(tVals[r][0]).toUpperCase() : "";
       if (header.includes('HARIAN')) {
          for(let c=1; c<=12; c++) {
             tHarian[c] = parseSmartQty(tVals[r][c]);
          }
       }
       if (header.includes('BULANAN')) {
          for(let c=1; c<=12; c++) {
             tBulanan[c] = parseSmartQty(tVals[r][c]);
          }
       }
    }
  }

  const sheetDyg = ss.getSheetByName('ORD_DYG');
  let kanbanOrders = [];
  let dailyStats = {}; 
  let resumeDyeing = {}; 
  let availableDates = [];
  let values = [];

  if (sheetDyg) {
    values = sheetDyg.getDataRange().getValues();
    let dSet = new Set();
    for (let i = 1; i < values.length; i++) {
        let dStr = parseDateObj(values[i][0]);
        if(dStr) dSet.add(dStr);
    }
    availableDates = Array.from(dSet).sort().reverse();
  }

  let latestDataDate = availableDates.length > 0 ? availableDates[0] : parseDateObj(new Date());
  let kanbanFilterStr = targetDate;
  if (!targetDate || !availableDates.includes(targetDate)) {
      kanbanFilterStr = latestDataDate;
  }

  // 7 Hari Terakhir untuk grafik / ringkasan mingguan
  let tDate = new Date(kanbanFilterStr + 'T00:00:00');
  let last7Dates = [];
  for(let i=6; i>=0; i--) {
      let d = new Date(tDate);
      d.setDate(d.getDate() - i);
      last7Dates.push(parseDateObj(d));
  }

  if (values.length > 1) {
      for (let i = 1; i < values.length; i++) {
        let rowDateStr = parseDateObj(values[i][0]); 
        let realClpStr = parseDateObj(values[i][7]); 
        if (!rowDateStr) continue;

        let cleanKikc = sanitizeName(values[i][1]); 
        
        // Parsing QTY yang akurat (menghindari 123.75 menjadi puluhan ribu)
        let qtyVal = parseSmartQty(values[i][5]);
        let cycleVal = parseSmartQty(values[i][6]);
        let bngName = sanitizeName(values[i][2]) || '-';

        // Resume Data (7 Hari)
        if (last7Dates.includes(rowDateStr)) {
            if (!resumeDyeing[bngName]) {
                resumeDyeing[bngName] = {};
                last7Dates.forEach(d => resumeDyeing[bngName][d] = { plan: 0, real: 0 });
            }
            resumeDyeing[bngName][rowDateStr].plan += qtyVal;
            if (realClpStr === rowDateStr) {
                resumeDyeing[bngName][rowDateStr].real += qtyVal;
            }
        }

        // Inisialisasi statistik tanggal input
        if (!dailyStats[rowDateStr]) {
            dailyStats[rowDateStr] = { order: 0, realized: 0, pending: 0, hasNewOrder: false };
        }
        
        // 1. Order HANYA bertambah di tanggal terbitnya (rowDateStr)
        dailyStats[rowDateStr].order += qtyVal;
        dailyStats[rowDateStr].hasNewOrder = true;

        // 2. Realisasi HANYA bertambah di tanggal aktual celup (realClpStr)
        if (realClpStr) {
            if (!dailyStats[realClpStr]) {
                dailyStats[realClpStr] = { order: 0, realized: 0, pending: 0, hasNewOrder: false };
            }
            dailyStats[realClpStr].realized += qtyVal;
        }

        // 3. Pending untuk tanggal input jika belum dicelup
        if (!realClpStr) {
            dailyStats[rowDateStr].pending += qtyVal;
        }

        // 4. KANBAN CARD: Hanya ambil yang diinput pada tanggal yang dipilih
        if (rowDateStr === kanbanFilterStr) {
            kanbanOrders.push({
                'TGL RPH': rowDateStr,
                'REAL CLP': realClpStr,
                'NO KIKC': cleanKikc, 
                'NO BENANG': bngName, 
                'WARNA': values[i][3] ? String(values[i][3]).trim() : '', 
                'NO MC': values[i][4] ? String(values[i][4]).trim().toUpperCase() : '',
                'QTY (Kg)': qtyVal, 
                'CYCLE': cycleVal 
            });
        }
      }
  }

  // Hitung Summary Harian vs Mingguan untuk Widget Atas
  if (!dailyStats[kanbanFilterStr]) {
      dailyStats[kanbanFilterStr] = { order: 0, realized: 0, pending: 0, hasNewOrder: false };
  }

  let summaryHarian = {
      order: dailyStats[kanbanFilterStr].order || 0,
      realized: dailyStats[kanbanFilterStr].realized || 0,
      pending: dailyStats[kanbanFilterStr].pending || 0
  };

  let summaryMingguan = { order: 0, realized: 0, pending: 0 };
  last7Dates.forEach(d => {
      if (dailyStats[d]) {
          summaryMingguan.order += dailyStats[d].order || 0;
          summaryMingguan.realized += dailyStats[d].realized || 0;
          summaryMingguan.pending += dailyStats[d].pending || 0;
      }
  });

  return { 
      kanban: kanbanOrders, 
      stats: dailyStats, 
      summaryHarian: summaryHarian,
      summaryMingguan: summaryMingguan,
      targets: { harian: tHarian, bulanan: tBulanan }, 
      mc: dataMc, 
      mcCaps: dataMcCaps, 
      kikc: dataKikc, 
      activeKanbanDate: kanbanFilterStr,
      resumeData: resumeDyeing,
      last7Dates: last7Dates 
  };
}

