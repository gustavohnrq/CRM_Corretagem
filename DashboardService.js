/**
 * DashboardService - Painel do Menu + Recalculo de métricas
 * Atualiza:
 * - Controle_Semanal (últimas 4 semanas)
 * - Funil_Mensal (mês atual)
 * Lê:
 * - Follow_Up (vencidos / hoje / próximos 7 dias)
 *
 * Fontes principais:
 * - Leads_Compradores: usa "Data Entrada" e "Status"
 * - Agenda_Visitas: usa "Data"
 * - Follow_Up: usa "Próxima Data de Contato" (ou "Próxima Data de Contato" = col exata)
 */

function getDashboardData(){
  // garante schema da agenda
  ensureSchema_();

  // Se quiser, recalcular automaticamente sempre que abrir o Menu:
  // (deixe ligado se sua base não for gigante)
  rebuildControleSemanal();
  rebuildFunilMensal();

  const weekly = readControleSemanalLast4_();
  const monthly = readFunilMensalCurrent_();
  const follow = readFollowUpBuckets_();

  const weeklyTotals = weekly.reduce((acc,r)=>({
    ligacoes: acc.ligacoes + (r.ligacoes||0),
    conversas: acc.conversas + (r.conversas||0),
    visitas: acc.visitas + (r.visitas||0),
    propostas: acc.propostas + (r.propostas||0),
    vendas: acc.vendas + (r.vendas||0),
  }), {ligacoes:0, conversas:0, visitas:0, propostas:0, vendas:0});

  return {
    monthKey: monthly.monthKey,
    weeklyRows: weekly.map(r=>({
      semana: r.semana,
      ligacoes: r.ligacoes,
      conversas: r.conversas,
      visitas: r.visitas,
      propostas: r.propostas,
      vendas: r.vendas
    })),
    weeklyTotals,
    monthly,
    follow
  };
}

/** Recalcula Controle_Semanal com base em Leads_Compradores + Agenda_Visitas */
function rebuildControleSemanal(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOut = ss.getSheetByName("Controle_Semanal");
  if (!shOut) throw new Error('Aba "Controle_Semanal" não existe.');

  const today = new Date();
  const weeks = lastNWeeks_(today, 4); // [{label, start, end}]

  // Métricas derivadas (heurística consistente):
  // - Ligações: Leads_Compradores com Data Entrada na semana
  // - Conversas: Leads_Compradores com Status != "Novo" e Data Entrada na semana
  // - Visitas: Agenda_Visitas com Data na semana
  // - Propostas: Leads_Compradores com Status == "Proposta" na semana (Data Entrada)
  // - Vendas: Leads_Compradores com Status == "Fechado" na semana (Data Entrada)
  const leads = readSheetObjects_("Leads_Compradores");
  const agenda = readSheetObjects_("Agenda_Visitas");

  const rows = weeks.map(w=>{
    const leadWeek = leads.filter(x => inRange_(parseDateAny_(x["Data Entrada"]), w.start, w.end));
    const ligacoes = leadWeek.length;

    const conversas = leadWeek.filter(x => norm_(x["Status"]) && norm_(x["Status"]) !== "novo").length;

    const propostas = leadWeek.filter(x => norm_(x["Status"]) === "proposta").length;
    const vendas = leadWeek.filter(x => norm_(x["Status"]) === "fechado").length;

    const visitas = agenda.filter(x => inRange_(parseDateAny_(x["Data"]), w.start, w.end)).length;

    return {
      semana: w.label,
      ligacoes, conversas, visitas, propostas, vendas
    };
  });

  // Grava no Controle_Semanal (mantém cabeçalho)
  // Formato esperado (do seu modelo): Semana | Ligações | Conversas | Visitas | Propostas | Vendas
  shOut.getRange(2,1, Math.max(shOut.getLastRow()-1,1), shOut.getLastColumn()).clearContent();

  const values = rows.map(r=>[r.semana, r.ligacoes, r.conversas, r.visitas, r.propostas, r.vendas]);
  if (values.length){
    shOut.getRange(2,1, values.length, 6).setValues(values);
  }

  // TOTAL na última linha (opcional)
  const totalRow = 2 + values.length;
  shOut.getRange(totalRow,1).setValue("TOTAL");
  shOut.getRange(totalRow,1).setFontWeight("bold");
  for (let c=2;c<=6;c++){
    const colLetter = String.fromCharCode(64+c);
    shOut.getRange(totalRow,c).setFormula(`=SUM(${colLetter}2:${colLetter}${totalRow-1})`);
  }
}

/** Recalcula Funil_Mensal (mês atual) com base em Leads_Compradores + Agenda_Visitas */
function rebuildFunilMensal(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shOut = ss.getSheetByName("Funil_Mensal");
  if (!shOut) throw new Error('Aba "Funil_Mensal" não existe.');

  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const monthEnd = new Date(now.getFullYear(), now.getMonth()+1, 1); // exclusivo
  const monthKey = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");

  const leads = readSheetObjects_("Leads_Compradores");
  const agenda = readSheetObjects_("Agenda_Visitas");

  const leadMonth = leads.filter(x => inRange_(parseDateAny_(x["Data Entrada"]), monthStart, monthEnd));
  const leadsCount = leadMonth.length;
  const contatos = leadMonth.filter(x => norm_(x["Status"]) && norm_(x["Status"]) !== "novo").length;
  const propostas = leadMonth.filter(x => norm_(x["Status"]) === "proposta").length;
  const vendas = leadMonth.filter(x => norm_(x["Status"]) === "fechado").length;

  const visitas = agenda.filter(x => inRange_(parseDateAny_(x["Data"]), monthStart, monthEnd)).length;

  // Saída minimalista (padrão flexível)
  // Colunas recomendadas: Mês | Leads | Contatos | Visitas | Propostas | Vendas | Taxa Conversão (%)
  // Vamos preencher a linha do mês atual e calcular taxa Vendas/Leads.
  const headers = shOut.getRange(1,1,1,shOut.getLastColumn()).getValues()[0].map(h=>String(h||"").trim());
  const col = (name)=> headers.indexOf(name)+1;

  // Se a estrutura não existir, criamos o cabeçalho padrão
  const needsHeader =
    headers.length < 6 || headers[0] === "" || headers[0].toLowerCase() === "mês" ? false : false;

  // Grava na primeira linha disponível do mês atual (atualiza se já existir)
  const lr = shOut.getLastRow();
  let targetRow = -1;
  if (lr >= 2){
    const monthVals = shOut.getRange(2,1,lr-1,1).getValues().flat().map(x=>String(x||"").trim());
    const idx = monthVals.findIndex(x=>x === monthKey);
    if (idx >= 0) targetRow = idx + 2;
  }
  if (targetRow === -1) targetRow = lr + 1;

  // Se o cabeçalho estiver vazio/errado, padroniza:
  if (shOut.getLastColumn() < 7 || headers[0] === ""){
    shOut.clear();
    shOut.getRange(1,1,1,7).setValues([["Mês","Leads","Contatos","Visitas","Propostas","Vendas","Taxa Conversão (%)"]]);
    shOut.getRange(1,1,1,7).setFontWeight("bold");
    targetRow = 2;
  }

  const taxa = (leadsCount > 0) ? (vendas / leadsCount) : 0;

  shOut.getRange(targetRow,1,1,7).setValues([[monthKey, leadsCount, contatos, visitas, propostas, vendas, taxa]]);
  shOut.getRange(targetRow,7).setNumberFormat("0.0%");
}

/* ===========================
   Leitura para o Menu
=========================== */

function readControleSemanalLast4_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Controle_Semanal");
  if (!sh) return [];
  const lr = sh.getLastRow();
  if (lr < 2) return [];

  // Pega até 4 linhas antes do TOTAL
  const data = sh.getRange(2,1, lr-1, Math.min(sh.getLastColumn(),6)).getValues();
  const cleaned = data.filter(r => String(r[0]||"").trim() && String(r[0]).toUpperCase() !== "TOTAL");

  const last4 = cleaned.slice(-4);
  return last4.map(r=>({
    semana: r[0],
    ligacoes: Number(r[1]||0),
    conversas: Number(r[2]||0),
    visitas: Number(r[3]||0),
    propostas: Number(r[4]||0),
    vendas: Number(r[5]||0),
  }));
}

function readFunilMensalCurrent_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Funil_Mensal");
  if (!sh) return { monthKey:"", funil:{}, rates:{} };

  const now = new Date();
  const monthKey = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");

  const lr = sh.getLastRow();
  if (lr < 2) return { monthKey, funil:{leads:0,contatos:0,visitas:0,propostas:0,vendas:0}, rates:{} };

  const data = sh.getRange(2,1, lr-1, Math.min(sh.getLastColumn(), 7)).getValues();
  const row = data.map(r=>({m:String(r[0]||"").trim(), r})).find(x=>x.m === monthKey);

  const leads = row ? Number(row.r[1]||0) : 0;
  const contatos = row ? Number(row.r[2]||0) : 0;
  const visitas = row ? Number(row.r[3]||0) : 0;
  const propostas = row ? Number(row.r[4]||0) : 0;
  const vendas = row ? Number(row.r[5]||0) : 0;

  const rates = {
    lead_para_contato: (leads>0)? (contatos/leads) : null,
    contato_para_visita: (contatos>0)? (visitas/contatos) : null,
    visita_para_proposta: (visitas>0)? (propostas/visitas) : null,
    proposta_para_venda: (propostas>0)? (vendas/propostas) : null
  };

  return { monthKey, funil:{leads, contatos, visitas, propostas, vendas}, rates };
}

function readFollowUpBuckets_(){
  const rows = readSheetObjects_("Follow_Up"); // se o nome da aba for diferente, ajuste aqui
  const tz = Session.getScriptTimeZone();
  const today = startOfDay_(new Date());

  const plus7 = new Date(today.getTime() + 7*24*60*60*1000);

  // tenta encontrar a coluna de próxima data
  const keyNext = findKey_(rows, ["Próxima Data de Contato","Próxima Data","Próximo Follow-up","Próximo Contato","Próxima Data de Contato"]);

  const mapped = rows.map(x=>{
    const dt = parseDateAny_(x[keyNext]);
    return {
      nome: x["Nome"] || x["Nome Proprietário"] || x["Cliente"] || "",
      tipo: x["Tipo (Comprador/Vendedor)"] || x["Tipo"] || "",
      telefone: x["Telefone"] || "",
      proximo: x[keyNext] || "",
      dt
    };
  }).filter(x=>x.dt); // só com data válida

  const overdue = mapped.filter(x => x.dt < today);
  const todayItems = mapped.filter(x => sameDay_(x.dt, today));
  const week = mapped.filter(x => x.dt > today && x.dt < plus7);

  // ordena
  overdue.sort((a,b)=>a.dt-b.dt);
  todayItems.sort((a,b)=>a.dt-b.dt);
  week.sort((a,b)=>a.dt-b.dt);

  return { overdue, today: todayItems, week };
}

/* ===========================
   Helpers
=========================== */

function readSheetObjects_(sheetName){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1,1,1,lc).getValues()[0].map(h=>String(h||"").trim());
  const data = sh.getRange(2,1,lr-1,lc).getValues();

  return data.map(row=>{
    const obj = {};
    for (let i=0;i<headers.length;i++){
      if (headers[i]) obj[headers[i]] = row[i];
    }
    return obj;
  });
}

function parseDateAny_(v){
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return startOfDay_(v);

  const s = String(v).trim();
  if (!s) return null;

  // yyyy-mm-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(s)){
    const [y,m,d] = s.slice(0,10).split("-").map(Number);
    const dt = new Date(y, m-1, d);
    return isNaN(dt.getTime()) ? null : startOfDay_(dt);
  }

  // dd/mm/yyyy ou dd/mm
  if (/^\d{1,2}\/\d{1,2}(\/\d{2,4})?$/.test(s)){
    const parts = s.split("/");
    const d = Number(parts[0]);
    const m = Number(parts[1]);
    let y = parts[2] ? Number(parts[2]) : (new Date()).getFullYear();
    if (y < 100) y += 2000;
    const dt = new Date(y, m-1, d);
    return isNaN(dt.getTime()) ? null : startOfDay_(dt);
  }

  // fallback: Date.parse
  const t = Date.parse(s);
  if (!isNaN(t)) return startOfDay_(new Date(t));

  return null;
}

function inRange_(dt, start, endExclusive){
  if (!dt) return false;
  return dt >= start && dt < endExclusive;
}

function startOfDay_(dt){
  const x = new Date(dt);
  x.setHours(0,0,0,0);
  return x;
}

function sameDay_(a, b){
  return a && b && a.getTime() === b.getTime();
}

function norm_(s){
  return String(s||"").trim().toLowerCase();
}

function lastNWeeks_(today, n){
  // semana = segunda a domingo (mais útil para seu ciclo seg-sáb)
  const end = startOfDay_(today);
  const day = end.getDay(); // 0 domingo..6 sábado
  const diffToMonday = (day === 0) ? 6 : (day - 1);
  const monday = new Date(end.getTime() - diffToMonday*24*60*60*1000);

  const weeks = [];
  for (let i=n-1;i>=0;i--){
    const start = new Date(monday.getTime() - i*7*24*60*60*1000);
    const endEx = new Date(start.getTime() + 7*24*60*60*1000);
    const label = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-'W'ww");
    weeks.push({ label, start, end: endEx });
  }
  return weeks;
}

function findKey_(rows, candidates){
  if (!rows || rows.length === 0) return candidates[0];
  const keys = Object.keys(rows[0] || {});
  for (const c of candidates){
    const hit = keys.find(k => k.toLowerCase() === c.toLowerCase());
    if (hit) return hit;
  }
  // fallback: primeira candidata
  return candidates[0];
}