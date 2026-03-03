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