/**
 * DashboardService v2 - sem dependência de abas calculadas (Controle_Semanal/Funil_Mensal)
 */

function getDashboardData(filters){
  ensureSchema_();

  const f = filters || {};
  const tz = Session.getScriptTimeZone();
  const now = new Date();

  const weekStart = f.weekStart ? dsParseDateAny_(f.weekStart) : startOfWeek_(now);
  const weekEnd = f.weekEnd ? addDays_(dsParseDateAny_(f.weekEnd), 1) : addDays_(weekStart, 7);

  const funilStart = f.funilStart ? dsParseDateAny_(f.funilStart) : new Date(now.getFullYear(), now.getMonth(), 1);
  const funilEndInclusive = f.funilEnd ? dsParseDateAny_(f.funilEnd) : now;
  const funilEnd = addDays_(funilEndInclusive, 1);
  const monthKey = `${fmtDate_(funilStart)} a ${fmtDate_(funilEndInclusive)}`;

  const leadsCompradores = readSheetObjects_("Leads_Compradores");
  const leadsVendedores = readSheetObjects_("Leads_Vendedores");
  const visitas = readSheetObjects_("Fato_Visitas");
  const propostas = readSheetObjects_("Fato_Proposta");
  const vendas = readSheetObjects_("Fato_Venda");
  const captacoes = readSheetObjects_("Fato_Captacao");

  const weekly = calcWeeklyControl_(weekStart, weekEnd, {
    leadsCompradores, leadsVendedores, visitas, propostas, captacoes
  });

  const monthly = calcMonthlyFunnel_(funilStart, funilEnd, {
    leadsCompradores, leadsVendedores, visitas, propostas, vendas, captacoes
  });

  const follow = readFollowUpBucketsByBoards_();

  return {
    period: {
      weekStart: fmtDate_(weekStart),
      weekEnd: fmtDate_(addDays_(weekEnd, -1)),
      monthKey,
      funilStart: fmtDate_(funilStart),
      funilEnd: fmtDate_(funilEndInclusive)
    },
    weekly,
    monthly,
    kpis: calcKpiCharts_(funilStart, funilEnd, { leadsCompradores, visitas, propostas, vendas, captacoes }),
    follow
  };
}

function rebuildControleSemanal(){
  return { ok:true, mode:"on_the_fly" };
}

function rebuildFunilMensal(){
  return { ok:true, mode:"on_the_fly" };
}

function calcWeeklyControl_(start, endEx, data){
  const metas = {
    ligacoesVendaMin:80, ligacoesVendaMax:120,
    visitasVendaMin:3, visitasVendaMax:5,
    propostaVenda:1,
    ligacoesCaptacaoMin:30, ligacoesCaptacaoMax:50,
    visitasCaptacaoMin:3, visitasCaptacaoMax:5,
    captacao:1
  };

  const ligVenda = data.leadsCompradores.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data Entrada","DataEntrada"])), start, endEx)).length;
  const ligCap = data.leadsVendedores.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data Entrada","DataEntrada"])), start, endEx)).length;

  const visitasVenda = data.visitas.filter(r=>{
    const d = dsParseDateAny_(pick_(r,["Data_Visita","Data Visita","Data"]));
    const tipo = dsNorm_(pick_(r,["Tipo_Visita","Tipo Visita"]));
    return dsInRange_(d,start,endEx) && (tipo === "venda" || tipo === "");
  }).length;

  const visitasCap = data.visitas.filter(r=>{
    const d = dsParseDateAny_(pick_(r,["Data_Visita","Data Visita","Data"]));
    const tipo = dsNorm_(pick_(r,["Tipo_Visita","Tipo Visita"]));
    return dsInRange_(d,start,endEx) && tipo === "captacao";
  }).length;

  const propostasVenda = data.propostas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), start, endEx)).length;

  const captacoesQtd = data.captacoes.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["DataCadastro","Data Cadastro"])), start, endEx)).length;

  return {
    metas,
    rows: [
      metricRow_("Frente 1 - Ligações Venda", ligVenda, metas.ligacoesVendaMin, metas.ligacoesVendaMax),
      metricRow_("Frente 1 - Visitas Venda", visitasVenda, metas.visitasVendaMin, metas.visitasVendaMax),
      metricRow_("Frente 1 - Proposta Venda", propostasVenda, metas.propostaVenda, metas.propostaVenda),
      metricRow_("Frente 2 - Ligações Captação", ligCap, metas.ligacoesCaptacaoMin, metas.ligacoesCaptacaoMax),
      metricRow_("Frente 2 - Visitas Captação", visitasCap, metas.visitasCaptacaoMin, metas.visitasCaptacaoMax),
      metricRow_("Frente 2 - Captação", captacoesQtd, metas.captacao, metas.captacao)
    ]
  };
}

function calcMonthlyFunnel_(start, endEx, data){
  const leads = data.leadsCompradores.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data Entrada"])), start, endEx)).length;
  const visitas = data.visitas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data_Visita","Data"])), start, endEx)).length;
  const propostas = data.propostas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), start, endEx)).length;
  const vendas = data.vendas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), start, endEx)).length;
  const captacoes = data.captacoes.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["DataCadastro"])), start, endEx)).length;

  const rates = {
    leads_para_visitas: leads>0 ? visitas/leads : null,
    visitas_para_propostas: visitas>0 ? propostas/visitas : null,
    propostas_para_vendas: propostas>0 ? vendas/propostas : null,
    lead_vendedor_para_captacao: data.leadsVendedores.length>0 ? captacoes/data.leadsVendedores.length : null
  };

  return { funil:{leads,visitas,propostas,vendas,captacoes}, rates };
}


function calcKpiCharts_(start, endEx, data){
  const buckets = buildWeeklyBuckets_(start, endEx);

  const leads = data.leadsCompradores.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data Entrada"])), start, endEx)).length;
  const visitas = data.visitas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data_Visita","Data"])), start, endEx)).length;
  const propostas = data.propostas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), start, endEx)).length;
  const vendasRows = data.vendas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), start, endEx));
  const vendas = vendasRows.length;
  const captacoes = data.captacoes.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["DataCadastro"])), start, endEx)).length;

  const periodDays = Math.max(1, Math.ceil((endEx.getTime()-start.getTime())/(24*60*60*1000)));
  const weeksInPeriod = Math.max(1, periodDays / 7);

  const seriesLeadsVisitas = buckets.map(b=>{
    const l = data.leadsCompradores.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data Entrada"])), b.start, b.end)).length;
    const v = data.visitas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data_Visita","Data"])), b.start, b.end)).length;
    return l>0 ? v/l : 0;
  });

  const seriesVisitasPropostas = buckets.map(b=>{
    const v = data.visitas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data_Visita","Data"])), b.start, b.end)).length;
    const p = data.propostas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), b.start, b.end)).length;
    return v>0 ? p/v : 0;
  });

  const seriesPropostasVendas = buckets.map(b=>{
    const p = data.propostas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), b.start, b.end)).length;
    const v = data.vendas.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["Data"])), b.start, b.end)).length;
    return p>0 ? v/p : 0;
  });

  const seriesCaptacoesSemana = buckets.map(b=>{
    const c = data.captacoes.filter(r=>dsInRange_(dsParseDateAny_(pick_(r,["DataCadastro"])), b.start, b.end)).length;
    return c;
  });

  const convLeadsVisitas = leads>0 ? visitas/leads : 0;
  const convVisitasPropostas = visitas>0 ? propostas/visitas : 0;
  const convPropostasVendas = propostas>0 ? vendas/propostas : 0;
  const captacoesPorSemana = captacoes / weeksInPeriod;

  return [
    { key:"kpi1", title:"Conversão Leads → Visitas", value: convLeadsVisitas, display: pctValue_(convLeadsVisitas), target: 0.15, targetDisplay:"Meta 15%", series: seriesLeadsVisitas, max: 1 },
    { key:"kpi2", title:"Conversão Visitas → Propostas", value: convVisitasPropostas, display: pctValue_(convVisitasPropostas), target: 0.35, targetDisplay:"Meta 35%", series: seriesVisitasPropostas, max: 1 },
    { key:"kpi3", title:"Conversão Propostas → Vendas", value: convPropostasVendas, display: pctValue_(convPropostasVendas), target: 0.30, targetDisplay:"Meta 30%", series: seriesPropostasVendas, max: 1 },
    { key:"kpi4", title:"Captações por Semana", value: captacoesPorSemana, display: numValue_(captacoesPorSemana), target: 1, targetDisplay:"Meta 1/semana", series: seriesCaptacoesSemana, max: Math.max(1, ...seriesCaptacoesSemana, 1) }
  ];
}

function buildWeeklyBuckets_(start, endEx){
  const out = [];
  let cur = new Date(start);
  while (cur < endEx){
    const next = addDays_(cur, 7);
    out.push({ start:new Date(cur), end: next < endEx ? next : new Date(endEx) });
    cur = next;
    if (out.length > 24) break;
  }
  if (!out.length) out.push({ start:new Date(start), end:new Date(endEx) });
  return out;
}

function dsParseMoney_(v){
  const s = String(v || "").trim();
  if (!s) return 0;
  const n = Number(s.replace(/[R$\s]/g,"").replace(/\./g,"").replace(",","."));
  return isNaN(n) ? 0 : n;
}

function pctValue_(x){
  return `${Math.round((Number(x)||0)*1000)/10}%`;
}

function numValue_(x){
  const n = Number(x||0);
  return (Math.round(n*100)/100).toString().replace('.', ',');
}

function readFollowUpBucketsByBoards_(){
  const boards = {
    leads: readFollowFromSheet_("Leads_Compradores", ["Nome"], ["Telefone"], ["Próximo Follow-up", "Próxima Data de Contato"]),
    captacoes: readFollowFromSheet_("Fato_Captacao", ["Captadores", "Proprietario"], ["Captadores"], ["Próximo Follow-up", "Próxima Data de Contato"]),
    visitas: readFollowFromSheet_("Fato_Visitas", ["Id_Visita"], ["Id_Imovel"], ["Próximo Follow-up", "Próxima Data de Contato"]),
    propostas: readFollowFromSheet_("Fato_Proposta", ["Id_Proposta"], ["Id_Visita"], ["Próximo Follow-up", "Próxima Data de Contato"]),
    vendas: readFollowFromSheet_("Fato_Venda", ["Id_Venda"], ["Id_Proposta"], ["Próximo Follow-up", "Próxima Data de Contato"])
  };
  return boards;
}

function readFollowFromSheet_(sheetName, nameCandidates, phoneCandidates, dateCandidates){
  const rows = readSheetObjects_(sheetName);
  const today = dsStartOfDay_(new Date());
  const plus7 = addDays_(today, 7);

  const all = rows.map(r=>{
    const proximoRaw = pick_(r, dateCandidates);
    const dt = dsParseDateAny_(proximoRaw);
    return {
      nome: pick_(r, nameCandidates),
      telefone: pick_(r, phoneCandidates),
      proximo: proximoRaw,
      dt
    };
  }).filter(x=>x.dt);

  return {
    overdue: all.filter(x=>x.dt < today).sort((a,b)=>a.dt-b.dt),
    today: all.filter(x=>dsSameDay_(x.dt,today)).sort((a,b)=>a.dt-b.dt),
    week: all.filter(x=>x.dt > today && x.dt < plus7).sort((a,b)=>a.dt-b.dt)
  };
}

function metricRow_(label, atual, min, max){
  let status = "red";
  if (atual >= min && atual <= max) status = "green";
  else if (atual > max || (atual >= Math.max(1, Math.floor(min*0.7)))) status = "yellow";
  return { label, atual, min, max, status };
}

function pick_(obj, candidates){
  if (!obj) return "";
  const map = {};
  Object.keys(obj).forEach(k=> map[normHeader_(k)] = k);
  for (const c of (candidates || [])){
    const f = map[normHeader_(c)];
    if (f) return obj[f] || "";
  }
  return "";
}

function normHeader_(s){
  return String(s||"")
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"");
}

function readSheetObjects_(sheetName){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2 || lc < 1) return [];

  const headers = sh.getRange(1,1,1,lc).getDisplayValues()[0].map(h=>String(h||"").trim());
  const data = sh.getRange(2,1,lr-1,lc).getDisplayValues();

  return data.map(row=>{
    const obj = {};
    for (let i=0;i<headers.length;i++) if (headers[i]) obj[headers[i]] = row[i];
    return obj;
  });
}

function dsParseDateAny_(v){
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return dsStartOfDay_(v);
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m){
    const dt = new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
    return isNaN(dt.getTime()) ? null : dsStartOfDay_(dt);
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m){
    const dt = new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
    return isNaN(dt.getTime()) ? null : dsStartOfDay_(dt);
  }
  const t = Date.parse(s);
  return isNaN(t) ? null : dsStartOfDay_(new Date(t));
}

function dsStartOfDay_(dt){ const x = new Date(dt); x.setHours(0,0,0,0); return x; }
function dsSameDay_(a,b){ return a && b && a.getTime() === b.getTime(); }
function dsInRange_(dt,start,endEx){ return !!dt && dt >= start && dt < endEx; }
function addDays_(dt, n){ return new Date(dt.getTime() + n*24*60*60*1000); }
function startOfWeek_(dt){ const d = dsStartOfDay_(dt); const wd=d.getDay(); const diff=(wd===0?6:wd-1); return addDays_(d,-diff); }
function fmtDate_(dt){ return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy"); }
function dsNorm_(s){ return String(s||"").trim().toLowerCase(); }
