function installDailyMetricsTrigger(){
  // remove gatilhos antigos desse handler para evitar duplicar
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction() === "dailyMetricsJob_") ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger("dailyMetricsJob_")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
}

function dailyMetricsJob_(){
  // atualiza ambos
  rebuildControleSemanal();
  rebuildFunilMensal();
}

function installDailyFollowUpTrigger(){
  if (typeof FU_installDailyTrigger_6am !== 'function') {
    throw new Error('FollowUpService não carregado.');
  }
  return FU_installDailyTrigger_6am();
}

function runDailyFollowUpNow(){
  if (typeof FU_dailyFollowUpJob_ !== 'function') {
    throw new Error('FollowUpService não carregado.');
  }
  return FU_dailyFollowUpJob_();
}

function installAllAutomationTriggers(){
  const metrics = installDailyMetricsTrigger();
  const follow = installDailyFollowUpTrigger();
  return { ok: true, metrics, follow };
}
