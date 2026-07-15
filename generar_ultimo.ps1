$here=Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Generar-ReporteATS.ps1')
$bases=@("$env:USERPROFILE\Documents\ATS_EXPORT","$env:USERPROFILE\Documentos\ATS_EXPORT","$env:USERPROFILE\Desktop\ATS_EXPORT","$env:USERPROFILE\OneDrive\Desktop\ATS_EXPORT")
if($args.Count -ge 1){ $bases=@($args[0]) }
$cand=@()
foreach($b in $bases){ if(Test-Path $b){ Get-ChildItem -LiteralPath $b -Directory | ForEach-Object { if((Test-Path (Join-Path $_.FullName 'EVENTS_CURRENT.CSV')) -or (Test-Path (Join-Path $_.FullName 'PROFILE_HISTORY.JSON'))){ $cand+=$_ } } } }
if($cand.Count -eq 0){ Write-Host "No encontre carpetas de exportacion con crudos."; $bases|ForEach-Object{Write-Host "   $_"}; exit 1 }
$ult=($cand|Sort-Object LastWriteTime -Descending)[0]
Write-Host ("Ultima exportacion: "+$ult.FullName)
$r=New-ReporteATS -Crudos $ult.FullName -Plantilla (Join-Path $here 'plantilla_reporte_ats.xlsx')
Write-Host ("[OK] {0}  eventos={1} sucesos={2} score={3} ({4})" -f $r.Salida,$r.Eventos,$r.Sucesos,$r.Score,$r.Banda)
try{ Invoke-Item $r.Salida }catch{}
