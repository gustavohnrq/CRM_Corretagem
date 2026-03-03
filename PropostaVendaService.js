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
  const propostas = DataService.listRecords("Fato_Proposta", "Id_Proposta", ["Data", "Valor da Proposta", "status", "Id_Visita"]);
  return propostas.map(p => {
    const vis = PV_tryGetVisitaContext_(p.raw["Id_Visita"]);
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
        idImovel: vis ? vis.id_imovel : "",
        idAgendamento: vis ? vis.id_agendamento : "",
        notaMediaVisita: vis ? vis.nota_media : ""
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
  const visita = PV_tryGetVisitaContext_(proposta["Id_Visita"]);
  return { proposta, visita };
}

function PV_listVendasDetailed_v1() {
  ensureSchema_();
  const vendas = DataService.listRecords("Fato_Venda", "Id_Venda", ["Id_Proposta", "Data", "Valor da Venda", "Forma de Pagamento", "Comissão", "Data de Recebimento Comissão"]);
  return vendas.map(v => {
    let ctx = null;
    try {
      ctx = v.raw["Id_Proposta"] ? PV_getPropostaContextById_v1(v.raw["Id_Proposta"]) : null;
    } catch (e) {
      ctx = null;
    }
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
        propostaModalidade: proposta ? (proposta["Modalidade de Pagamento"] || "-") : "-",
        visitaData: visita ? (visita.data_visita || "-") : "-",
        idImovel: visita ? (visita.id_imovel || "-") : "-",
        clientes,
        endereco,
        comissao: v.raw["Comissão"] || "-",
        dataRecebimentoComissao: v.raw["Data de Recebimento Comissão"] || "-"
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
