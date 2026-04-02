"C:\Program Files\Autodesk\Navisworks Manage 2026\roamer.exe" ^
  -NoGui ^
  -log "C:\Users\user\Downloads\testsplit\roamer.log" ^
  -OpenFile "C:\Users\user\Downloads\Snowdon.nwd" ^
  -AddPluginAssembly "C:\Users\user\OneDrive - Autodesk\Documents\ObjectARX 2026\samples\customer\NavisworksBatchBoxPublish\bin\Debug\net48\NavisworksBatchBoxPublish.dll" ^
  -ExecuteAddInPlugin "NavisworksBatchBoxPublish.BatchExport.OAI1" ^
    "outdir=C:\Users\user\Downloads\testsplit" ^
    "prefix=Set_" ^
    "log=C:\Users\user\Downloads\testsplit\NavisworksBatchSearchSetPublish.log" ^
  -Exit
