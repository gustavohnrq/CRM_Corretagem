/**
 * PropostaVendaService - Fato_Proposta + Fato_Venda
 */

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
  const rows = DataService.listRecords("Fato_Visitas", "Id_Visita", ["Data_Visita", "Id_Imovel", "Id_Agendamento"]);
  return rows.map(r => ({ id: r.id, label: `Visita ${r.id} • Data ${r.raw["Data_Visita"] || "-"} • Imóvel ${r.raw["Id_Imovel"] || "-"}` }));
}

function PV_getVisitaContextById_v1(idVisita) {
  ensureSchema_();
  return PDF_getVisitaPayload_v1(idVisita);
}

function PV_listPropostasDetailed_v1() {
  ensureSchema_();
  const propostas = DataService.listRecords("Fato_Proposta", "Id_Proposta", ["Data", "Valor da Proposta", "status", "Id_Visita"]);
  return propostas.map(p => {
    const vis = p.raw["Id_Visita"] ? PV_getVisitaContextById_v1(p.raw["Id_Visita"]) : null;
    const clientes = vis && vis.clientes_nomes ? vis.clientes_nomes.join(", ") : "-";
    const endereco = vis && vis.imovel ? (vis.imovel["Endereço"] || vis.imovel["Endereco"] || vis.imovel["Quadra/Endereço"] || "-") : "-";
    return {
      id: p.id,
      label: `Proposta ${p.id} • ${p.raw["Valor da Proposta"] || "-"} • ${p.raw["status"] || "-"}`,
      raw: p.raw,
      enrich: {
        clientes,
        endereco,
        dataVisita: vis ? vis.data_visita : "",
        idImovel: vis ? vis.id_imovel : ""
      }
    };
  });
}

function PV_getPropostaById_v1(idProposta) {
  ensureSchema_();
  return DataService.getById("Fato_Proposta", "Id_Proposta", idProposta);
}

function PV_getNextIdProposta_v1() {
  ensureSchema_();
  return DataService.getNextNumericId("Fato_Proposta", "Id_Proposta");
}

function PV_upsertProposta_v1(obj) {
  ensureSchema_();
  if (!String(obj["Id_Visita"] || "").trim()) throw new Error("Id_Visita é obrigatório.");
  return DataService.upsertById("Fato_Proposta", "Id_Proposta", obj);
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
  const visita = proposta["Id_Visita"] ? PV_getVisitaContextById_v1(proposta["Id_Visita"]) : null;
  return { proposta, visita };
}

function PV_listVendasDetailed_v1() {
  ensureSchema_();
  const vendas = DataService.listRecords("Fato_Venda", "Id_Venda", ["Id_Proposta", "Data", "Valor da Venda", "Forma de Pagamento"]);
  return vendas.map(v => {
    const ctx = v.raw["Id_Proposta"] ? PV_getPropostaContextById_v1(v.raw["Id_Proposta"]) : null;
    const proposta = ctx ? ctx.proposta : null;
    const visita = ctx ? ctx.visita : null;
    const clientes = visita && visita.clientes_nomes ? visita.clientes_nomes.join(", ") : "-";
    const endereco = visita && visita.imovel ? (visita.imovel["Endereço"] || visita.imovel["Endereco"] || visita.imovel["Quadra/Endereço"] || "-") : "-";

    return {
      id: v.id,
      label: `Venda ${v.id} • Proposta ${v.raw["Id_Proposta"] || "-"} • ${v.raw["Valor da Venda"] || "-"}`,
      raw: v.raw,
      enrich: {
        propostaValor: proposta ? (proposta["Valor da Proposta"] || "-") : "-",
        propostaStatus: proposta ? (proposta["status"] || "-") : "-",
        visitaId: proposta ? (proposta["Id_Visita"] || "-") : "-",
        clientes,
        endereco
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
  return DataService.upsertById("Fato_Venda", "Id_Venda", obj);
}

function PV_deleteVendaById_v1(idVenda) {
  ensureSchema_();
  return DataService.deleteById("Fato_Venda", "Id_Venda", idVenda);
}
