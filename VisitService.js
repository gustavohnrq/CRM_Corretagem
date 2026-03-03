/**
 * VisitService - Registro de Visita (Fato_Visitas) + Avaliações (Fato_Avaliacao)
 */

function listAgendaForSelect() {
  ensureSchema_();

  const agenda = DataService.listRecords("Agenda_Visitas", "id", ["Data", "Hora", "Id_Cliente", "Imóvel (Código)"]);
  const clientes = DataService.listRecords("Base_Clientes", "ID", ["Nome Completo"]);
  const mapNome = {};
  clientes.forEach(c => mapNome[String(c.id).trim()] = (c.raw["Nome Completo"] || c.label || c.id));

  return agenda.map(a => {
    const raw = a.raw || {};
    const idCli = String(raw["Id_Cliente"] || "").trim();
    const nome = mapNome[idCli] || ("Cliente " + idCli);
    const label = `${raw["Data"] || ""} ${raw["Hora"] || ""} • ${nome} • ${raw["Imóvel (Código)"] || ""} • (id ${a.id})`;
    return { id: a.id, label };
  });
}

/** ✅ NOVO: retorna um registro da Agenda_Visitas por id (preencher Data_Visita e Id_Imovel automaticamente) */
function getAgendaById(idAgenda) {
  ensureSchema_();
  return DataService.getById("Agenda_Visitas", "id", idAgenda);
}

/** Cria registro em Agenda_Visitas e devolve id */
function createAgendaVisit(obj) {
  ensureSchema_();
  return DataService.upsertById("Agenda_Visitas", "id", obj);
}

/** Salva/edita Fato_Visitas (ID = Id_Visita) */
function saveFatoVisita(obj) {
  const idv = String(obj["Id_Visita"] || "").trim();
  if (!idv) throw new Error("Id_Visita é obrigatório em Fato_Visitas.");
  return DataService.upsertById("Fato_Visitas", "Id_Visita", obj);
}

function getFatoVisitaByIdVisita(idVisita) {
  return DataService.getById("Fato_Visitas", "Id_Visita", idVisita);
}

function listFatoVisitas() {
  return DataService.listRecords("Fato_Visitas", "Id_Visita", ["Data_Visita", "Id_Imovel", "Proposta"]);
}

/**
 * UPSERT avaliação por (Id_Visita + Id_Cliente)
 * - se já existir avaliação desse cliente nessa visita, atualiza
 * - senão, cria nova (id_Avaliacao automático)
 */
function upsertAvaliacaoByVisitaCliente(obj) {
  const idVisita = String(obj["Id_Visita"] || "").trim();
  const idCliente = String(obj["Id_Cliente"] || "").trim();
  if (!idVisita) throw new Error("Id_Visita é obrigatório na avaliação.");
  if (!idCliente) throw new Error("Id_Cliente é obrigatório na avaliação.");

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fato_Avaliacao");
  if (!sh) throw new Error('Aba "Fato_Avaliacao" não existe.');

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });

  const lr = sh.getLastRow();
  if (lr >= 2) {
    const colVis = map["Id_Visita"];
    const colCli = map["Id_Cliente"];
    const colIdA = map["id_Avaliacao"];
    if (!colVis || !colCli || !colIdA) throw new Error("Colunas obrigatórias não encontradas em Fato_Avaliacao.");

    const data = sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (String(r[colVis - 1]).trim() === idVisita && String(r[colCli - 1]).trim() === idCliente) {
        obj["id_Avaliacao"] = r[colIdA - 1];
        return DataService.upsertById("Fato_Avaliacao", "id_Avaliacao", obj);
      }
    }
  }

  obj["id_Avaliacao"] = obj["id_Avaliacao"] || "";
  return DataService.upsertById("Fato_Avaliacao", "id_Avaliacao", obj);
}

function listAvaliacoesByVisita(idVisita) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fato_Avaliacao");
  if (!sh) return [];

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return [];

  const headers = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h || "").trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });

  const idxVis = map["Id_Visita"];
  if (idxVis === undefined) return [];

  const data = sh.getRange(2, 1, lr - 1, lc).getValues();
  return data
    .filter(r => String(r[idxVis]).trim() === String(idVisita).trim())
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { if (h) obj[h] = r[i]; });
      return obj;
    });
}

function deleteAvaliacaoById(idAvaliacao) {
  return DataService.deleteById("Fato_Avaliacao", "id_Avaliacao", idAvaliacao);
}