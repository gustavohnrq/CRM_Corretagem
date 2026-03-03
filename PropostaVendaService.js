/**
 * PropostaVendaService - Fato_Proposta + Fato_Venda
 */


function PV_normKey_(s) {
  if (typeof _normHeader_ === "function") return _normHeader_(s);
  return String(s || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function PV_getSheetObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(h => String(h || "").trim());
  const values = sh.getRange(2, 1, lr - 1, lc).getDisplayValues();

  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

function PV_pickByCandidates_(obj, candidates) {
  if (!obj) return "";
  const keyMap = {};
  Object.keys(obj).forEach(k => { keyMap[PV_normKey_(k)] = k; });

  for (let i = 0; i < candidates.length; i++) {
    const want = PV_normKey_(candidates[i]);
    const found = keyMap[want];
    if (found !== undefined) return String(obj[found] || "").trim();
  }
  return "";
}

function PV_listSheetByIdRobust_(sheetName, idCandidates) {
  return PV_getSheetObjects_(sheetName)
    .map(raw => {
      const id = PV_pickByCandidates_(raw, idCandidates);
      return { id, raw };
    })
    .filter(r => String(r.id || "").trim());
}

function ensurePropostaVendaSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders_(ss, "Fato_Proposta", [
    "Data",
    "Id_Proposta",
    "Valor da Proposta",
    "Modalidade de Pagamento",
    "status",
    "Id_Visita"
  ]);

  ensureSheetWithHeaders_(ss, "Fato_Venda", [
    "Id_Proposta",
    "Data",
    "Id_Venda",
    "Valor da Venda",
    "Forma de Pagamento",
    "Comissão",
    "Data de Recebimento Comissão"
  ]);
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  const lc = sh.getLastColumn();
  const current = lc > 0 ? sh.getRange(1, 1, 1, lc).getDisplayValues()[0].map(h => String(h || "").trim()) : [];

  const equal = current.length === headers.length && headers.every((h, i) => h === current[i]);
  if (equal) return;

  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
}

function PV_listVisitasForSelect_v1() {
  ensureSchema_();

  // Mesmo mecanismo do módulo de PDF:
  // - base em Fato_Visitas
  // - clientes derivados de Fato_Avaliacao + Base_Clientes
  // - ordenação por Id_Visita desc
  if (typeof PDF_listVisitasForSelect_v1 === "function") {
    try {
      return PDF_listVisitasForSelect_v1();
    } catch (e) {
      // fallback abaixo
    }
  }

  // fallback defensivo (não deve ser o caminho principal)
  try {
    return DataService
      .listRecords("Fato_Visitas", "Id_Visita", ["Data_Visita", "Id_Imovel"])
      .map(r => ({ id: r.id, label: r.label || (`Visita ${r.id}`) }));
  } catch (e) {
    return [];
  }
}
function PV_getVisitaContextById_v1(idVisita) {
  ensureSchema_();
  return PDF_getVisitaPayload_v1(idVisita);
}

function PV_tryGetVisitaContext_(idVisita) {
  const id = String(idVisita || "").trim();
  if (!id) return null;
  try {
    return PV_getVisitaContextById_v1(id);
  } catch (e) {
    return null;
  }
}

function PV_listPropostasDetailed_v1() {
  ensureSchema_();

  const propostas = PV_listSheetByIdRobust_("Fato_Proposta", ["Id_Proposta", "ID_ProPOSTA", "id_proposta"]);

  return propostas.map(p => {
    const idVisitaRaw = PV_pickByCandidates_(p.raw, ["Id_Visita", "id_visita"]);
    const valor = PV_pickByCandidates_(p.raw, ["Valor da Proposta", "valor_proposta"]);
    const status = PV_pickByCandidates_(p.raw, ["status", "Status"]);
    const modalidade = PV_pickByCandidates_(p.raw, ["Modalidade de Pagamento", "modalidade_pagamento"]);

    const vis = PV_tryGetVisitaContext_(idVisitaRaw);
    const clientes = vis && vis.clientes_nomes ? vis.clientes_nomes.join(", ") : "-";
    const endereco = vis && vis.imovel ? (vis.imovel["Endereço"] || vis.imovel["Endereco"] || vis.imovel["Quadra/Endereço"] || "-") : "-";

    return {
      id: p.id,
      label: `Proposta ${p.id} • ${valor || "-"} • ${status || "-"}`,
      raw: {
        ...p.raw,
        "Id_Proposta": p.id,
        "Id_Visita": idVisitaRaw,
        "Valor da Proposta": valor,
        "status": status,
        "Modalidade de Pagamento": modalidade
      },
      enrich: {
        clientes,
        endereco,
        dataVisita: vis ? vis.data_visita : "",
        idImovel: vis ? vis.id_imovel : "",
        idAgendamento: vis ? vis.id_agendamento : "",
        notaMediaVisita: vis ? vis.nota_media : ""
      }
    };
  });
}

function PV_getPropostaById_v1(idProposta) {
  ensureSchema_();
  const target = String(idProposta || "").trim();
  if (!target) return null;

  const rows = PV_listSheetByIdRobust_("Fato_Proposta", ["Id_Proposta", "ID_ProPOSTA", "id_proposta"]);
  const hit = rows.find(r => String(r.id || "").trim() === target);
  if (!hit) return null;

  const idVisitaRaw = PV_pickByCandidates_(hit.raw, ["Id_Visita", "id_visita"]);
  const valor = PV_pickByCandidates_(hit.raw, ["Valor da Proposta", "valor_proposta"]);
  const status = PV_pickByCandidates_(hit.raw, ["status", "Status"]);
  const modalidade = PV_pickByCandidates_(hit.raw, ["Modalidade de Pagamento", "modalidade_pagamento"]);
  const data = PV_pickByCandidates_(hit.raw, ["Data", "data"]);

  return {
    ...hit.raw,
    "Data": data,
    "Id_Proposta": target,
    "Valor da Proposta": valor,
    "Modalidade de Pagamento": modalidade,
    "status": status,
    "Id_Visita": idVisitaRaw
  };
}

function PV_getNextIdProposta_v1() {
  ensureSchema_();
  return DataService.getNextNumericId("Fato_Proposta", "Id_Proposta");
}

function PV_upsertProposta_v1(obj) {
  ensureSchema_();
  if (!String(obj["Id_Visita"] || "").trim()) throw new Error("Id_Visita é obrigatório.");
  const res = DataService.upsertById("Fato_Proposta", "Id_Proposta", obj);
  try {
    if (typeof FU_syncFollowUpForRecord_ === "function") {
      FU_syncFollowUpForRecord_("Fato_Proposta", "Id_Proposta", (res && res.id) || obj["Id_Proposta"], obj);
    }
  } catch (e) {}
  return res;
}

function PV_deletePropostaById_v1(idProposta) {
  ensureSchema_();
  return DataService.deleteById("Fato_Proposta", "Id_Proposta", idProposta);
}

function PV_listPropostasForSelect_v1() {
  ensureSchema_();
  const propostas = PV_listPropostasDetailed_v1();
  return propostas.map(p => ({
    id: p.id,
    label: `Proposta ${p.id} • Visita ${p.raw["Id_Visita"] || "-"} • ${p.raw["Valor da Proposta"] || "-"} • ${p.raw["status"] || "-"}`
  }));
}

function PV_getPropostaContextById_v1(idProposta) {
  ensureSchema_();
  const proposta = PV_getPropostaById_v1(idProposta);
  if (!proposta) return null;
  const visita = PV_tryGetVisitaContext_(proposta["Id_Visita"]);
  return { proposta, visita };
}

function PV_listVendasDetailed_v1() {
  ensureSchema_();

  const vendas = PV_listSheetByIdRobust_("Fato_Venda", ["Id_Venda", "id_venda"]);

  return vendas.map(v => {
    const idProposta = PV_pickByCandidates_(v.raw, ["Id_Proposta", "id_proposta"]);
    const valorVenda = PV_pickByCandidates_(v.raw, ["Valor da Venda", "valor_venda"]);
    const forma = PV_pickByCandidates_(v.raw, ["Forma de Pagamento", "forma_pagamento"]);
    const comissao = PV_pickByCandidates_(v.raw, ["Comissão", "Comissao", "comissao"]);
    const dtRec = PV_pickByCandidates_(v.raw, ["Data de Recebimento Comissão", "Data de Recebimento Comissao", "data_recebimento_comissao"]);

    let ctx = null;
    try {
      ctx = idProposta ? PV_getPropostaContextById_v1(idProposta) : null;
    } catch (e) {
      ctx = null;
    }
    const proposta = ctx ? ctx.proposta : null;
    const visita = ctx ? ctx.visita : null;
    const clientes = visita && visita.clientes_nomes ? visita.clientes_nomes.join(", ") : "-";
    const endereco = visita && visita.imovel ? (visita.imovel["Endereço"] || visita.imovel["Endereco"] || visita.imovel["Quadra/Endereço"] || "-") : "-";

    return {
      id: v.id,
      label: `Venda ${v.id} • Proposta ${idProposta || "-"} • ${valorVenda || "-"}`,
      raw: {
        ...v.raw,
        "Id_Venda": v.id,
        "Id_Proposta": idProposta,
        "Valor da Venda": valorVenda,
        "Forma de Pagamento": forma,
        "Comissão": comissao,
        "Data de Recebimento Comissão": dtRec
      },
      enrich: {
        propostaValor: proposta ? (proposta["Valor da Proposta"] || "-") : "-",
        propostaStatus: proposta ? (proposta["status"] || "-") : "-",
        visitaId: proposta ? (proposta["Id_Visita"] || "-") : "-",
        propostaModalidade: proposta ? (proposta["Modalidade de Pagamento"] || "-") : "-",
        visitaData: visita ? (visita.data_visita || "-") : "-",
        idImovel: visita ? (visita.id_imovel || "-") : "-",
        clientes,
        endereco,
        comissao: comissao || "-",
        dataRecebimentoComissao: dtRec || "-"
      }
    };
  });
}

function PV_getVendaById_v1(idVenda) {
  ensureSchema_();
  return DataService.getById("Fato_Venda", "Id_Venda", idVenda);
}

function PV_getNextIdVenda_v1() {
  ensureSchema_();
  return DataService.getNextNumericId("Fato_Venda", "Id_Venda");
}

function PV_upsertVenda_v1(obj) {
  ensureSchema_();
  if (!String(obj["Id_Proposta"] || "").trim()) throw new Error("Id_Proposta é obrigatório.");
  const res = DataService.upsertById("Fato_Venda", "Id_Venda", obj);
  try {
    if (typeof FU_syncFollowUpForRecord_ === "function") {
      FU_syncFollowUpForRecord_("Fato_Venda", "Id_Venda", (res && res.id) || obj["Id_Venda"], obj);
    }
  } catch (e) {}
  return res;
}

function PV_deleteVendaById_v1(idVenda) {
  ensureSchema_();
  return DataService.deleteById("Fato_Venda", "Id_Venda", idVenda);
}
