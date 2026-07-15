<#
  exportar_ats.ps1  -  Captura EXPORT ALL del modulo ATS por USB-Serial y genera REPORTE_ATS.xlsx.
  PowerShell PURO (Windows nativo). NO instala nada, NO usa Excel, NO usa Python.
  Watchdog de pared: NUNCA se cuelga. Guarda primero los 6 crudos; el xlsx es el ultimo paso.
#>
param([string]$Puerto="", [string]$Destino="$env:USERPROFILE\Documents\ATS_EXPORT")
$ErrorActionPreference='Continue'
$here=Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'Generar-ReporteATS.ps1')

# --- autodeteccion del puerto ---
if(-not $Puerto){
  $ps=[System.IO.Ports.SerialPort]::GetPortNames()
  if($ps.Count -eq 0){ Write-Host "No se detecto el modulo. Conecta el cable USB y reintenta."; Read-Host "Enter para salir"; exit 1 }
  elseif($ps.Count -eq 1){ $Puerto=$ps[0] }
  else { Write-Host ("Puertos: "+($ps -join ", ")); $Puerto=Read-Host "Cual uso (ej: COM3)" }
}
$carpeta=Join-Path $Destino (Get-Date -Format 'yyyy-MM-dd_HHmmss')
New-Item -ItemType Directory -Path $carpeta -Force | Out-Null

$port=New-Object System.IO.Ports.SerialPort $Puerto,115200,([System.IO.Ports.Parity]::None),8,([System.IO.Ports.StopBits]::One)
$port.ReadTimeout=1000; $port.NewLine="`n"; $port.DtrEnable=$false; $port.RtsEnable=$false
try{ $port.Open() }catch{
  Write-Host "ERROR: no se pudo abrir $Puerto. Cierra el Monitor Serial de Arduino y verifica el puerto."; Read-Host "Enter para salir"; exit 1 }

$guardados=@(); $sw=[System.Diagnostics.Stopwatch]::StartNew(); $LIMITE_S=90
try{
  Start-Sleep -Milliseconds 1500; $port.DiscardInBuffer()
  Write-Host "Conectado a $Puerto. Pidiendo EXPORT ALL..."
  $port.WriteLine("STATUS"); Start-Sleep -Milliseconds 300; $port.DiscardInBuffer(); $port.WriteLine("EXPORT ALL")
  $arch=$null; $buf=New-Object System.Collections.Generic.List[string]; $enc=New-Object System.Text.UTF8Encoding($false)
  $inactivo=0; $reintentos=0; $recibio=$false
  while($true){
    if($sw.Elapsed.TotalSeconds -gt $LIMITE_S){ Write-Host "  (watchdog: corto a los $LIMITE_S s para no colgarse)"; break }
    try{
      $linea=$port.ReadLine().TrimEnd("`r"); $inactivo=0; $recibio=$true
      if($linea -match '^----- INICIO (.+) -----$'){ $arch=$Matches[1].Trim(); $buf.Clear(); Write-Host ("Recibiendo "+$arch+"..."); continue }
      if($linea -match '^----- FIN (.+) -----$'){ if($arch){ $r=Join-Path $carpeta $arch; [System.IO.File]::WriteAllText($r,(($buf -join "`n")+"`n"),$enc); $guardados+=$r; Write-Host ("  [OK] "+$arch) }; $arch=$null; continue }
      if($linea -like '===== EXPORT ALL COMPLETO*'){ break }
      if($arch){ [void]$buf.Add($linea) } elseif($linea){ Write-Host ("  modulo> "+$linea) }
    } catch [TimeoutException] {
      $inactivo++
      if(-not $recibio){ if(($inactivo % 3) -eq 0){ if($reintentos -ge 6){ break }; $reintentos++; $port.DiscardInBuffer(); $port.WriteLine("EXPORT ALL") } }
      elseif($inactivo -ge 10){ break }
    }
  }
} finally { try{ $port.Close() }catch{} }

if($guardados.Count -eq 0){
  Write-Host ""; Write-Host "No se recibio respuesta del modulo. Revisa: firmware V4 corriendo, puerto correcto, Monitor Serial CERRADO."
  Remove-Item -LiteralPath $carpeta -Force -Recurse -ErrorAction SilentlyContinue; Read-Host "Enter para salir"; exit 1
}
Write-Host ""; Write-Host ("Crudos guardados ("+$guardados.Count+") en: "+$carpeta)

# --- generar el xlsx (PowerShell puro). Si falla, los crudos quedan igual. ---
try{
  $stamp=Split-Path $carpeta -Leaf
  $rep=New-ReporteATS -Crudos $carpeta -Plantilla (Join-Path $here 'plantilla_reporte_ats.xlsx') -Salida (Join-Path $carpeta ("REPORTE_ATS_"+$stamp+".xlsx"))
  Write-Host ("[OK] REPORTE_ATS.xlsx generado  (eventos="+$rep.Eventos+", sucesos="+$rep.Sucesos+", score="+$rep.Score+")")
} catch {
  Write-Host ("[AVISO] No se pudo generar el xlsx: "+$_.Exception.Message)
  Write-Host "        Los 6 archivos crudos quedaron guardados; puedes correr generar_reporte.bat luego."
}
# --- analisis predictivo (curvas de tendencia) si el analizador esta junto al kit ---
$pred=Join-Path $here 'Analizar-Predictivo.ps1'
if(Test-Path $pred){
  try{ . $pred
    $x=""; if($rep -and $rep.Salida -and (Test-Path $rep.Salida)){ $x=$rep.Salida }
    Invoke-AnalisisPredictivo -Crudos $carpeta -Xlsx $x | Out-Null }
  catch{ Write-Host ("[AVISO] Analisis predictivo no corrio: "+$_.Exception.Message) }
}
if($rep -and $rep.Salida -and (Test-Path $rep.Salida)){ try{ Invoke-Item $rep.Salida }catch{} }
try{ Invoke-Item $carpeta }catch{}
Write-Host ""; Read-Host "Enter para salir"
