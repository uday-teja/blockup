// Entry point
document.addEventListener('DOMContentLoaded', () => {
  const popup = createPopup();
  let currentTableId = '';

  const constants = getColumnIndexes();

  // Upload handler
  document.getElementById('upload').addEventListener('change', handleUpload);

  // Column visibility popup
  document.addEventListener('click', handleDocumentClick);

  // Export PDF
  document.getElementById('exportPdfBtn').addEventListener('click', exportToPDF);

  function createPopup() {
    const p = document.createElement('div');
    p.className = 'popup hidden';
    document.body.appendChild(p);
    return p;
  }

  function getColumnIndexes() {
    return {
      DETAIL_NO: 0,
      ITEM_CODE: 1,
      DESCRIPTION: 2,
      L_DIA: 3,
      WIDTH: 4,
      THICKNESS: 5,
      QTY: 6,
      MATERIAL: 7,
      WEIGHT: 8,
      RATE: 9,
      BLOCK_UP_RATE: 10,
      CORNER_CHAMFER: 11,
      Total: 12
    };
  }

  function handleUpload(e) {
    const fr = new FileReader();
    fr.onload = ev => processWorkbook(new Uint8Array(ev.target.result));
    fr.readAsArrayBuffer(e.target.files[0]);
  }

  function processWorkbook(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    const bom = XLSX.utils.sheet_to_json(wb.Sheets['CUTTING BOM'], { header: 1 });
    const assy = XLSX.utils.sheet_to_json(wb.Sheets['ASSY CHECKLIST'], { header: 1 });

    const idx = [
      bom.findIndex(r => r[0] === 'MAIN PLATES'),
      bom.findIndex(r => r[0] === 'VAP ROUND PARTS'),
      bom.findIndex(r => r[0] === 'VAP FLAT PARTS')
    ];


    const mainData = extractMainPlates(bom, assy, idx);
    const roundData = extractRoundParts(bom, idx);
    const flatData = bom.slice(idx[2] + 1).filter(r => parseInt(r[0]));

    const wo = bom.find(r => typeof r[0] === 'string' && r[0].includes('WO.NO:'));
    document.getElementById('woNumber').innerText = `Work Order No: ${wo ? wo[0].split(':')[1].trim() : 'UNKNOWN'}`;

    createTable(popup, 'main', mainData, 'Main Plates', getMainHeaders());
    createTable(popup, 'round', roundData, 'VAP Round Parts', getRoundHeaders(), true, true);
    createTable(popup, 'flat', flatData, 'VAP Flat Parts', getBaseHeaders());

    document.querySelector('.tab-button[data-target="main"]').click();
  }

  function extractMainPlates(bom, assy, idx) {
    return bom.slice(idx[0] + 1, idx[1]).filter(r => parseInt(r[0])).map(row => {
      const match = assy.find(r2 => r2[0] === row[0]);
      row[10] = match ? match[12] : '';
      row[11] = calculateChamfer(row[10]?.toLowerCase() || '', parseFloat(row[constants.THICKNESS]));
      return row;
    });
  }

  function extractRoundParts(bom, idx) {
    return bom.slice(idx[1] + 1, idx[2]).filter(r => parseInt(r[0])).map(row => {
      const item = (row[1] || '').toUpperCase();
      const dia = parseFloat(row[3]);
      const len = parseFloat(row[4]);
      row[10] = (calculateRoundRate(item, dia, len)).toFixed(2);
      return row;
    });
  }

  function calculateRoundRate(code, dia, len) {
    if (code.includes("GB") || code.includes("EGB")) {
      return calculateIDRate(len) + calculateODRate(len);
    }
    return calculateStandardRate(dia, len);
  }

  function calculateIDRate(len) {
    if (len <= 50) return 40;
    if (len <= 80) return 60;
    if (len <= 100) return 70;
    if (len <= 120) return 85;
    if (len <= 150) return 95;
    if (len <= 200) return 115;
    return 130;
  }

  function calculateODRate(len) {
    if (len <= 50) return 45;
    if (len <= 80) return 65;
    if (len <= 100) return 80;
    if (len <= 120) return 110;
    if (len <= 150) return 135;
    return 160;
  }

  function calculateStandardRate(d, l) {
    const table = [
      [200, [30, 35], [50, 40], [80, 50], [120, 60]],
      [350, [30, 50], [50, 55], [80, 60], [120, 70]],
      [500, [30, 70], [50, 75], [80, 90], [120, 100]],
      [600, [30, 100], [50, 105], [80, 140], [120, 160]],
      [700, [30, 130], [50, 140], [80, 160], [120, 210]]
    ];
    for (const [lenLimit, ...brackets] of table) {
      if (l <= lenLimit) {
        for (const [dLimit, rate] of brackets) {
          if (d <= dLimit) return rate;
        }
      }
    }
    return 0;
  }

  function calculateChamfer(chamfer, thickness) {
    const chamferValue = parseInt(chamfer);
    if (chamferValue >= 3 && chamferValue <= 15) return 80;
    if (chamferValue >= 16 && chamferValue <= 25) return 180;
    if (chamferValue >= 26 && chamferValue <= 50) return 280;
    if (chamferValue >= 51 && chamferValue <= 75) return 380;
    if (chamferValue >= 76 && chamferValue <= 90) return thickness <= 60 ? 280 : 350;
    if (chamferValue >= 91 && chamferValue <= 100) return thickness <= 60 ? 300 : 400;
    return '';
  }

  function calculateMainRate(length, thickness, isDT5) {
    if(isDT5)
    {
      return (thickness <= 60 ? 0.30 : thickness <= 100 ? 0.38 : thickness <= 150 ? 0.48 : thickness <= 200 ? 0.58 : thickness <= 250 ? 0.58 : 0.58);
    }
    return length <= 1900
      ? (thickness <= 30 ? 0.14 : thickness <= 60 ? 0.16 : thickness <= 100 ? 0.21 : thickness <= 150 ? 0.26 : thickness <= 200 ? 0.38 : thickness <= 250 ? 0.43 : 0.48)
      : (thickness <= 60 ? 0.23 : thickness <= 100 ? 0.27 : thickness <= 150 ? 0.30 : thickness <= 200 ? 0.38 : thickness <= 250 ? 0.48 : 0.53);
  }

  function calculateMainCost(length, width,thickness, rate, isDt5) {
    if(isDt5)
    {
        return (((length + width) * thickness)/100) * 0.6;
    }
    return ((length * width)/100) * rate;
    // const c = length <= 300 ? 35 : length <= 500 ? 60 : length / 10;
    // return (((length * width) / 100) * rate + c + 80);
  }

  function calculateTotal(data, index) {
    return data.reduce((acc, row) => acc + (parseFloat(row[index]) || 0), 0).toFixed(2);
  }

  function createTable(popup, id, data, title, headers, hideWidth = false, skipCost = false) {

    const visibleColumns = headers.map((_, i) => i).filter(i => !(hideWidth && i === constants.WIDTH) && !(skipCost && i === constants.RATE) );
    const section = document.getElementById(id);
    section.innerHTML = `<h2 class="text-xl font-semibold mb-2">${title}</h2>`;
    const card = document.createElement('div');
    card.className = 'bg-white rounded shadow p-4 mb-6 w-full max-w-6xl mx-auto';
    section.appendChild(card);

    let tableHtml = `<table class="min-w-full divide-y divide-gray-200"><thead class="bg-gray-50 tbl-header"><tr>`;
    visibleColumns.forEach(i => {
      tableHtml += `<th class="px-3 py-2 text-xs font-medium text-gray-500 uppercase">${headers[i]}</th>`;
    });
    tableHtml += `</tr></thead><tbody class="bg-white divide-y divide-gray-200">`;

    data.forEach(row => {
      if (row[constants.WEIGHT] !== undefined && !isNaN(row[constants.WEIGHT])) {
        row[constants.WEIGHT] = parseFloat(row[constants.WEIGHT]).toFixed(2);
      }
      if (hideWidth) row[constants.WIDTH] = '';
      const length = parseFloat(row[constants.L_DIA]);
      const width = parseFloat(row[constants.WIDTH]);
      const thickness = parseFloat(row[constants.THICKNESS]);
      if (!skipCost && !isNaN(length) && !isNaN(width) && !isNaN(thickness)) {
        const rate = calculateMainRate(length, thickness);
        row[constants.RATE] = calculateMainCost(length, width, thickness, rate, String(row[1]).includes('DT-5') || String(row[1]).includes('DT5')).toFixed(2);
      } else {
        row[constants.RATE] = '';
      }


      const chamfer = parseFloat(row[constants.CORNER_CHAMFER]) || 0;
      const rate = parseFloat(row[constants.RATE]) || 0;
      row[constants.Total] = (chamfer + rate).toFixed(2);


      tableHtml += `<tr>` + visibleColumns.map(i => `<td class="px-3 py-2 text-sm text-gray-700">${row[i] || ''}</td>`).join('') + `</tr>`;
    });

    const totalParts = data.length;
    const totalWeight = calculateTotal(data, constants.WEIGHT);
    const totalCost = calculateTotal(data, constants.RATE);
    const totalChamfer = calculateTotal(data, constants.CORNER_CHAMFER);

    if (id === 'main') {
      const grandTotal = (parseFloat(totalCost) + parseFloat(totalChamfer)).toFixed(2);
      tableHtml += `<tr class='bg-gray-100 font-semibold text-gray-800 text-sm'>` +
        visibleColumns.map(colIndex => {
          if (colIndex === constants.DETAIL_NO) return `<td class='px-3 py-2'>Total</td>`;
          if (colIndex === constants.WEIGHT) return `<td class='px-3 py-2'>${totalWeight}</td>`;
          if (colIndex === constants.RATE) return `<td class='px-3 py-2'>₹${totalCost}</td>`;
          if (colIndex === constants.CORNER_CHAMFER) return `<td class='px-3 py-2'>₹${totalChamfer}</td>`;
          if (colIndex === constants.Total) return `<td class='px-3 py-2'>₹${grandTotal}</td>`;
          return `<td class='px-3 py-2'></td>`; 
        }).join('') + `</tr>`;
    } else {
      tableHtml += `<tr class="bg-gray-50"><td colspan="${visibleColumns.length}" class="px-3 py-2 text-right text-sm font-semibold text-gray-700">Total Parts: ${totalParts} | Total Weight: ${totalWeight}</td></tr>`;
    }

    tableHtml += `</tbody></table>`;
    card.innerHTML = tableHtml;
  }

  function handleDocumentClick(e) {
    if (e.target.classList.contains('tab-button')) {
      switchTab(e);
    } else if (e.target.classList.contains('gear')) {
      const rect = e.target.getBoundingClientRect();
      showColumnPopup(popup, currentTableId = e.target.dataset.gear, rect.left, rect.bottom);
    } else if (!popup.contains(e.target)) {
      popup.classList.add('hidden');
    }
  }

  function switchTab(e) {
    document.querySelectorAll('.tab-button').forEach(b => b.classList.remove('border-blue-600', 'text-blue-600'));
    e.target.classList.add('border-blue-600', 'text-blue-600');
    document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
    document.getElementById(e.target.dataset.target).classList.remove('hidden');
  }

  function exportToPDF() {
    const tab = document.querySelector('.tab-content:not(.hidden)');
    const clone = tab.cloneNode(true);
    const wrapper = document.createElement('div');
    wrapper.style.padding = '20px';
    wrapper.innerHTML = `<h1 style='text-align:center; font-size:20px; font-weight:bold;'>Uday Precision Solutions</h1><div style='text-align:center;'>${document.getElementById('woNumber').innerText + " Blockup Cost Sheet"}</div>`;
    wrapper.appendChild(clone);

    const timestamp = new Date().toISOString().slice(0, 10); // e.g., 2025-05-09
    const woNumber = document.getElementById('woNumber').innerText;
    const fileName = `${woNumber} Blockup Cost Sheet -${timestamp}.pdf`;

    html2pdf().set({ margin: 10, filename: fileName, image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 1.5 }, jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' } }).from(wrapper).save();
  }

  const getBaseHeaders = () => ["Detail No", "ITEM CODE", "Description", "L/DIA", "W", "T", "Qty", "MATERIAL", "Weight", "Rate"];
  const getMainHeaders = () => [...getBaseHeaders(), "Block-up Rate", "Corner Chamfer", "Total"];
  const getRoundHeaders = () => [...getBaseHeaders(), "CG Rate"];
});
