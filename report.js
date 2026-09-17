/* Logika report Monitor Delivery per sprinter.
   Dipakai di browser (window.MDReport) dan Node (module.exports) untuk tes. */
(function (root) {
  const COL = {
    awb: 'No. Waybill', dp: 'DP Delivery', waktuDel: 'Waktu Delivery', kode: 'Kode Sprinter',
    nama: 'Sprinter Delivery', inv: 'Waktu Scan Inventory', bm: 'Waktu Paket Bermasalah',
    alasanBm: 'Alasan Paket Bermasalah', ttd: 'Waktu TTD', ttdRetur: 'TTD Retur', dpTtd: 'DP TTD',
    sprTtd: 'Sprinter TTD', tujuan: 'Tujuan', cod: 'COD'
  };
  const WAJIB = ['awb', 'dp', 'waktuDel', 'kode', 'nama', 'inv', 'bm', 'ttd', 'ttdRetur'];

  function nilai(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
      if ('result' in v) return nilai(v.result);
      if (v.richText) return v.richText.map(t => t.text).join('');
      if ('text' in v) return v.text;
      if (v.error) return null;
    }
    if (typeof v === 'string' && v.trim() === '') return null;
    return v;
  }
  const ada = v => v !== null && v !== undefined;
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  const pad = n => String(n).padStart(2, '0');
  const tglStr = d => d instanceof Date
    ? d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate())
    : String(d).slice(0, 10);

  // Nama DP untuk judul & nama file; lebih dari 2 DP diringkas jadi GABUNGAN
  function labelDp(dp) {
    if (!dp || !dp.length) return 'DP';
    return dp.length <= 2 ? dp.join(' & ') : 'GABUNGAN ' + dp.length + ' DP';
  }

  function bacaBaris(ws) {
    const header = {};
    ws.getRow(1).eachCell((c, i) => { const h = nilai(c.value); if (h !== null) header[String(h).trim()] = i; });
    const kurang = WAJIB.filter(k => !(COL[k] in header)).map(k => COL[k]);
    if (kurang.length) {
      throw new Error('Kolom tidak ditemukan: ' + kurang.join(', ') +
        '. Pastikan file adalah export "Monitor Delivery(Refine)(Detail)".');
    }
    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const o = {};
      for (const k in COL) o[k] = COL[k] in header ? nilai(row.getCell(header[COL[k]]).value) : null;
      if (!ada(o.awb)) continue;
      rows.push(o);
    }
    return rows;
  }

  function hitung(rows) {
    for (const o of rows) {
      o.fTtd = Number(o.ttdRetur) === 0 && ada(o.ttd);   // TTD Retur = 1 adalah TTD retur di pengirim
      o.fBm = ada(o.bm);
      o.fInv = ada(o.inv);
      o.fGagal = !o.fTtd;                                 // wajib scan bermasalah & inventori
      o.fGagalNoInv = o.fGagal && !o.fInv;
      o.fGagalNoBm = o.fGagal && !o.fBm;
    }
    const grup = new Map();
    for (const o of rows) {
      const key = JSON.stringify([String(o.dp), String(o.kode), String(o.nama)]);
      if (!grup.has(key)) grup.set(key, { dp: String(o.dp), kode: String(o.kode), nama: String(o.nama), del: 0, ttd: 0, bm: 0, noBm: 0, inv: 0, noInv: 0 });
      const g = grup.get(key);
      g.del++; g.ttd += o.fTtd; g.bm += o.fGagal && o.fBm; g.noBm += o.fGagalNoBm; g.inv += o.fGagal && o.fInv; g.noInv += o.fGagalNoInv;
    }
    const rep = [...grup.values()]
      .sort((a, b) => cmp(a.dp, b.dp) || cmp(a.kode, b.kode) || cmp(a.nama, b.nama))
      .sort((a, b) => cmp(a.dp, b.dp) || b.del - a.del);
    const kunci = ['del', 'ttd', 'bm', 'noBm', 'inv', 'noInv'];
    const total = {};
    kunci.forEach(k => { total[k] = rep.reduce((n, g) => n + g[k], 0); });
    const urut = (a, b) => cmp(String(a.nama), String(b.nama)) || cmp(+a.waktuDel, +b.waktuDel);
    const tgl = [...new Set(rows.map(o => tglStr(o.waktuDel)))].sort();
    const dp = [...new Set(rows.map(o => String(o.dp)))].sort();   // DP terdeteksi dari file
    return { rep, total, dp, noInv: rows.filter(o => o.fGagalNoInv).sort(urut), noBm: rows.filter(o => o.fGagalNoBm).sort(urut), tgl };
  }

  function buatWorkbook(ExcelJS, hasil, namaSumber) {
    const wb = new ExcelJS.Workbook();
    const tipis = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    const border = { top: tipis, left: tipis, bottom: tipis, right: tipis };

    function sheet(nama, header, data, lebarKhusus) {
      const ws = wb.addWorksheet(nama, { views: [{ state: 'frozen', ySplit: 1 }] });
      ws.addRow(header);
      data.forEach(d => ws.addRow(d));
      const hr = ws.getRow(1);
      hr.height = 32;
      hr.eachCell(c => {
        c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
        c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });
      ws.eachRow(row => {
        for (let i = 1; i <= header.length; i++) {
          const c = row.getCell(i);
          c.border = border;
          if (c.value instanceof Date) c.numFmt = 'yyyy-mm-dd hh:mm:ss';
        }
      });
      header.forEach((h, i) => {
        let ln = String(h).length;
        data.forEach(d => { const v = d[i]; if (ada(v)) ln = Math.max(ln, v instanceof Date ? 19 : String(v).length); });
        ws.getColumn(i + 1).width = (lebarKhusus && lebarKhusus[i]) || Math.min(Math.max(ln + 2, 10), 70);
      });
      return ws;
    }

    const hRep = ['No', 'DP Delivery', 'Kode Sprinter', 'Sprinter Delivery', 'Total Delivery', 'Total TTD',
      'Total Scan Bermasalah', 'Total AWB Tidak Scan Bermasalah', 'Total Scan Inventori', 'Total AWB Tidak Scan Inventori', 'Keterangan'];
    const dRep = hasil.rep.map((g, i) => [i + 1, g.dp, g.kode, g.nama, g.del, g.ttd, g.bm, g.noBm, g.inv, g.noInv, '']);
    const t = hasil.total;
    dRep.push(['', 'TOTAL', '', '', t.del, t.ttd, t.bm, t.noBm, t.inv, t.noInv, '']);
    const ws = sheet('Report Sprinter', hRep, dRep, { 10: 40 });
    ws.eachRow((row, r) => {
      if (r === 1) return;
      row.getCell(11).alignment = { wrapText: true, vertical: 'top' };
      if (r === ws.rowCount) {
        for (let i = 1; i <= hRep.length; i++) {
          const c = row.getCell(i);
          c.font = { bold: true };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
        }
      } else {
        [8, 10].forEach(k => {
          if (row.getCell(k).value > 0) row.getCell(k).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        });
      }
    });

    const hDet = ['No. Waybill', 'DP Delivery', 'Kode Sprinter', 'Sprinter Delivery', 'Waktu Delivery',
      'Waktu Paket Bermasalah', 'Alasan Paket Bermasalah', 'Waktu Scan Inventory', 'Waktu TTD',
      'TTD Retur', 'DP TTD', 'Sprinter TTD', 'Tujuan', 'COD'];
    const det = o => [o.awb, o.dp, o.kode, o.nama, o.waktuDel, o.bm, o.alasanBm, o.inv, o.ttd, o.ttdRetur, o.dpTtd, o.sprTtd, o.tujuan, o.cod];
    sheet('AWB Tidak Scan Bermasalah', hDet, hasil.noBm.map(det));
    sheet('AWB Tidak Scan Inventori', hDet, hasil.noInv.map(det));

    sheet('Definisi', ['Kolom', 'Definisi'], [
      ['Total Delivery', 'Jumlah AWB yang di-scan delivery oleh sprinter'],
      ['Total TTD', 'AWB TTD oleh penerima (Waktu TTD ada & TTD Retur = 0). TTD Retur = 1 tidak dihitung (itu TTD retur di pengirim)'],
      ['Wajib Scan', 'Total Delivery - Total TTD = AWB yang wajib scan bermasalah dan wajib scan inventori'],
      ['Total Scan Bermasalah', 'AWB wajib scan yang punya Waktu Paket Bermasalah'],
      ['Total AWB Tidak Scan Bermasalah', 'AWB wajib scan yang TIDAK punya Waktu Paket Bermasalah (Scan Bermasalah + Tidak Scan Bermasalah = Delivery - TTD)'],
      ['Total Scan Inventori', 'AWB wajib scan yang punya Waktu Scan Inventory'],
      ['Total AWB Tidak Scan Inventori', 'AWB wajib scan yang TIDAK punya Waktu Scan Inventory (Scan Inventori + Tidak Scan Inventori = Delivery - TTD)'],
      ['Sumber', namaSumber + ' | Tanggal delivery: ' + hasil.tgl.join(', ')]
    ], { 1: 110 });

    return wb;
  }

  async function proses(ExcelJS, buffer, namaSumber) {
    const src = new ExcelJS.Workbook();
    await src.xlsx.load(buffer);
    const ws = src.worksheets[0];
    if (!ws) throw new Error('File tidak punya sheet.');
    const rows = bacaBaris(ws);
    if (!rows.length) throw new Error('Tidak ada data AWB di file.');
    const hasil = hitung(rows);
    const wb = buatWorkbook(ExcelJS, hasil, namaSumber);
    const out = await wb.xlsx.writeBuffer();
    const namaFile = 'Report Delivery, Scan Bermasalah dan Inventori ' + labelDp(hasil.dp) + ' ' + hasil.tgl[0] + '.xlsx';
    return { hasil, out, namaFile };
  }

  const api = { proses, hitung, bacaBaris, labelDp };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.MDReport = api;
})(this);
