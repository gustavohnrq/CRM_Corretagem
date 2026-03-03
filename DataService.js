var DataService = (function () {

  function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

  function sheet_(name) {
    const sh = ss_().getSheetByName(name);
    if (!sh) throw new Error("Aba não encontrada: " + name);
    return sh;
  }

  function headers_(sh) {
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) throw new Error("Aba sem colunas: " + sh.getName());
    return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
  }

  function headerMap_(headers) {
    const map = {};
    headers.forEach((h, i) => { if (h) map[h] = i + 1; });
    return map;
  }

  function getMeta(sheetName) {
    const sh = sheet_(sheetName);
    const headers = headers_(sh);
    return { sheetName, headers, headerMap: headerMap_(headers), lastRow: sh.getLastRow() };
  }

  function findRowById_(sh, idCol, idVal) {
    if (idVal === undefined || idVal === null || String(idVal).trim() === "") return -1;

    const headers = headers_(sh);
    const map = headerMap_(headers);
    const col = map[idCol];
    if (!col) throw new Error(`Coluna de ID "${idCol}" não existe na aba "${sh.getName()}"`);

    const lr = sh.getLastRow();
    if (lr < 2) return -1;

    const values = sh.getRange(2, col, lr - 1, 1).getValues().flat();
    const target = String(idVal).trim();

    for (let i = 0; i < values.length; i++) {
      if (String(values[i]).trim() === target) return i + 2;
    }
    return -1;
  }

  function rowToObj_(headers, rowValues) {
    const obj = {};
    for (let i = 0; i < headers.length; i++) {
      const k = headers[i];
      if (k) obj[k] = rowValues[i];
    }
    return obj;
  }

  function normalizeToRow_(headers, obj) {
    return headers.map(h => obj.hasOwnProperty(h) ? obj[h] : "");
  }

  function listRecords(sheetName, idCol, labelCols) {
    const sh = sheet_(sheetName);
    const headers = headers_(sh);
    const map = headerMap_(headers);

    if (!map[idCol]) throw new Error(`Coluna ID "${idCol}" não existe em "${sheetName}"`);

    const lr = sh.getLastRow();
    if (lr < 2) return [];

    const data = sh.getRange(2, 1, lr - 1, headers.length).getValues();
    const out = [];

    const idIndex = map[idCol] - 1;

    data.forEach((row) => {
      const id = row[idIndex];
      if (id === "" || id === null) return;

      const labels = (labelCols || []).map(c => {
        const idx = map[c] ? (map[c] - 1) : -1;
        return idx >= 0 ? String(row[idx] ?? "").trim() : "";
      }).filter(Boolean);

      out.push({
        id: String(id).trim(),
        label: labels.join(" • "),
        raw: rowToObj_(headers, row)
      });
    });

    return out;
  }

  function getById(sheetName, idCol, idVal) {
    const sh = sheet_(sheetName);
    const headers = headers_(sh);
    const row = findRowById_(sh, idCol, idVal);
    if (row === -1) return null;

    const values = sh.getRange(row, 1, 1, headers.length).getValues()[0];
    return rowToObj_(headers, values);
  }

  function upsertById(sheetName, idCol, obj) {
    const sh = sheet_(sheetName);
    const headers = headers_(sh);
    const map = headerMap_(headers);

    if (!map[idCol]) throw new Error(`Coluna ID "${idCol}" não existe em "${sheetName}"`);

    const idVal = obj[idCol];
    let row = findRowById_(sh, idCol, idVal);

    if (row === -1 && (idVal === "" || idVal === null || idVal === undefined)) {
      const nextId = getNextNumericId(sheetName, idCol);
      obj[idCol] = nextId;
      row = -1;
    }

    const rowValues = normalizeToRow_(headers, obj);

    if (row === -1) {
      sh.appendRow(rowValues);
      return { ok: true, action: "insert", id: String(obj[idCol]) };
    } else {
      sh.getRange(row, 1, 1, headers.length).setValues([rowValues]);
      return { ok: true, action: "update", id: String(obj[idCol]) };
    }
  }

  function deleteById(sheetName, idCol, idVal) {
    const sh = sheet_(sheetName);
    const row = findRowById_(sh, idCol, idVal);
    if (row === -1) return { ok: false, message: "ID não encontrado" };
    sh.deleteRow(row);
    return { ok: true };
  }

  function getNextNumericId(sheetName, idCol) {
    const sh = sheet_(sheetName);
    const headers = headers_(sh);
    const map = headerMap_(headers);
    const col = map[idCol];
    if (!col) throw new Error(`Coluna ID "${idCol}" não existe em "${sheetName}"`);

    const lr = sh.getLastRow();
    if (lr < 2) return 1;

    const values = sh.getRange(2, col, lr - 1, 1).getValues().flat();
    let max = 0;

    values.forEach(v => {
      const n = Number(String(v).trim());
      if (!isNaN(n) && isFinite(n)) max = Math.max(max, n);
    });

    return max + 1;
  }

  function listClientesForSelect() {
    const records = listRecords("Base_Clientes", "ID", ["Nome Completo", "Telefone"]);
    return records.map(r => ({ id: r.id, label: r.label }));
  }

  function listImoveisForSelect() {
    const records = listRecords("Estoque_Imoveis", "Código", ["Quadra", "Tipo", "Preço"]);
    return records.map(r => ({ id: r.id, label: r.label }));
  }

  return {
    getMeta,
    listRecords,
    getById,
    upsertById,
    deleteById,
    getNextNumericId,
    listClientesForSelect,
    listImoveisForSelect
  };

})();