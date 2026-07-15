# ATS Exporter — Código Completo y Lógica Funcional

Kit 100% PowerShell que captura los datos del módulo ATS por USB-Serial, guarda 6 archivos crudos y genera un `REPORTE_ATS.xlsx` profesional. **No requiere Python, ni Excel, ni instalar nada** — solo Windows nativo.

---

## Arquitectura y flujo

```
 [ESP32 + Firmware V4]
        │ USB-Serial (115200, DTR=false)
        ▼
 exportar_ats.bat  ──►  exportar_ats.ps1
        │                    │
        │  1. Autodetecta puerto COM
        │  2. Abre serial (DTR=false, RTS=false)
        │  3. Envía STATUS + EXPORT ALL
        │  4. Parsea bloques ----- INICIO/FIN -----
        │  5. Guarda 6 crudos en Documentos\ATS_EXPORT\<fecha>\
        │                    │
        │                    ▼
        │            Generar-ReporteATS.ps1
        │                    │
        │  6. Descomprime plantilla_reporte_ats.xlsx
        │  7. Inyecta datos en el XML interno (sheet1/2/3)
        │  8. Recomprime → REPORTE_ATS.xlsx
        │                    │
        ▼                    ▼
 Carpeta con 6 crudos + REPORTE_ATS.xlsx se abren automáticamente
```

### Los 6 archivos crudos que produce el firmware

| Archivo | Contenido |
|---|---|
| `RENDIMIENTO.TXT` | Resumen legible: score, tendencia, último evento, causa probable |
| `EVENTS_CURRENT.CSV` | Tabla de todos los eventos de transferencia/retransferencia |
| `SUCESOS.CSV` | Log de cambios de condición (apagón, retorno, normalización) |
| `VENTANA.csv` | Muestra de alta resolución del último evento |
| `PROFILE_HISTORY.JSON` | Score actual, banda, tendencia, acumulados |
| `BASELINE.JSON` | Línea base estadística (medianas, IQR, n eventos) |

### El reporte XLSX tiene 3 hojas

- **RESUMEN** — Score, banda, tendencia, confianza, último evento, causa probable, línea base
- **TablaEventos** — Todos los eventos (ID, timestamp, causa, t_transfer, voltajes, corriente, score)
- **TablaSucesos** — Todos los sucesos (timestamp, fuente, cambio, V, Hz, duración, diagnóstico)

---

## 1. `exportar_ats.bat` — Punto de entrada

```batch
@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0exportar_ats.ps1"
```

---

## 2. `exportar_ats.ps1` — Captura serial + genera reporte

```powershell
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
      if($linea -eq '===== EXPORT ALL COMPLETO ====='){ break }
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
  $rep=New-ReporteATS -Crudos $carpeta -Plantilla (Join-Path $here 'plantilla_reporte_ats.xlsx')
  Write-Host ("[OK] REPORTE_ATS.xlsx generado  (eventos="+$rep.Eventos+", sucesos="+$rep.Sucesos+", score="+$rep.Score+")")
  try{ Invoke-Item $rep.Salida }catch{}
} catch {
  Write-Host ("[AVISO] No se pudo generar el xlsx: "+$_.Exception.Message)
  Write-Host "        Los 6 archivos crudos quedaron guardados; puedes correr generar_reporte.bat luego."
}
try{ Invoke-Item $carpeta }catch{}
Write-Host ""; Read-Host "Enter para salir"
```

### Lógica clave de la captura serial

- **DTR=false, RTS=false** — crítico: si DTR está en true, el ESP32 se resetea al abrir el puerto.
- **Protocolo**: envía `STATUS` (para despertar), descarta buffer, luego `EXPORT ALL`.
- **Parseo por delimitadores**: `----- INICIO <archivo> -----` abre un bloque, `----- FIN <archivo> -----` lo cierra y guarda. `===== EXPORT ALL COMPLETO =====` termina.
- **Watchdog de 90s** — nunca se cuelga. Si no recibe nada, reintenta `EXPORT ALL` hasta 6 veces cada 3 timeouts.
- **Resiliencia**: los crudos se guardan PRIMERO; el xlsx es el último paso. Si falla el xlsx, los crudos quedan intactos.

---

## 3. `Generar-ReporteATS.ps1` — Motor del reporte (PowerShell puro)

Este es el corazón del exporter. Descomprime la plantilla `.xlsx` (que es un ZIP de XMLs), inyecta los datos en las celdas y tablas, y recomprime.

```powershell
<#
  Generar-ReporteATS.ps1  -  Genera REPORTE_ATS.xlsx desde los crudos del modulo ATS.
  PowerShell PURO (Windows nativo). NO usa Excel, NO usa Python, NO instala nada.
  Inyecta los datos en el XML de la plantilla .xlsx (descomprime, edita, recomprime).
#>
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

function _XmlEsc([string]$s){ if($null -eq $s){return ""}; $s=$s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'; return $s }
function _IsNum([string]$s){ return ($null -ne $s -and $s -match '^\s*-?\d+(\.\d+)?\s*$') }
function _Cap([string]$s){ if([string]::IsNullOrEmpty($s)){return ""}; return $s.Substring(0,1).ToUpper()+$s.Substring(1) }
function _Inv($n){ if($null -eq $n){return ""}; return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture,"{0}",$n) }

function _SetCell([ref]$xml,[string]$ref,[string]$value,[bool]$isNum){
  $pat='<c r="'+[regex]::Escape($ref)+'"( s="\d+")?[^>]*>.*?</c>'
  if($isNum -and -not (_IsNum $value)){ $isNum=$false }
  $repl={ param($m)
    $st=$m.Groups[1].Value
    if($isNum){ '<c r="'+$ref+'"'+$st+'><v>'+($value.Trim())+'</v></c>' }
    else{ '<c r="'+$ref+'"'+$st+' t="inlineStr"><is><t xml:space="preserve">'+(_XmlEsc $value)+'</t></is></c>' } }
  $xml.Value=[regex]::Replace($xml.Value,$pat,$repl,[System.Text.RegularExpressions.RegexOptions]::Singleline)
}
function _DataCell([string]$col,[int]$row,[string]$value,[bool]$isNum){
  $r=$col+$row
  if([string]::IsNullOrWhiteSpace($value)){ return '<c r="'+$r+'"/>' }
  if($isNum -and (_IsNum $value)){ return '<c r="'+$r+'" t="n"><v>'+($value.Trim())+'</v></c>' }
  return '<c r="'+$r+'" t="inlineStr"><is><t xml:space="preserve">'+(_XmlEsc $value)+'</t></is></c>'
}

function New-ReporteATS{
  param([Parameter(Mandatory=$true)][string]$Crudos,
        [string]$Plantilla=(Join-Path $PSScriptRoot 'plantilla_reporte_ats.xlsx'),
        [string]$Salida=(Join-Path $Crudos 'REPORTE_ATS.xlsx'))
  if(-not(Test-Path $Plantilla)){ throw "No encuentro la plantilla: $Plantilla" }
  function _Read($n){ $p=Join-Path $Crudos $n; if(Test-Path $p){ return (Get-Content -LiteralPath $p -Raw -Encoding UTF8) } return "" }
  function _Json($n){ try{ $t=_Read $n; if($t.Trim()){ return ($t|ConvertFrom-Json) } }catch{}; return $null }
  function _Csv($n){ $p=Join-Path $Crudos $n; if(Test-Path $p){ try{ return @(Import-Csv -LiteralPath $p) }catch{} }; return @() }
  $prof=_Json 'PROFILE_HISTORY.JSON'; $base=_Json 'BASELINE.JSON'
  $ev=_Csv 'EVENTS_CURRENT.CSV'; $su=_Csv 'SUCESOS.CSV'; $rend=_Read 'RENDIMIENTO.TXT'

  $tmp=Join-Path ([System.IO.Path]::GetTempPath()) ("ats_"+[System.Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  try{
    $unz=Join-Path $tmp 'x'
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Plantilla,$unz)
    $enc=New-Object System.Text.UTF8Encoding($false)

    # ---- HOJA 1: RESUMEN ----
    $s1p=Join-Path $unz 'xl/worksheets/sheet1.xml'; $s1=[System.IO.File]::ReadAllText($s1p,$enc)
    _SetCell ([ref]$s1) 'C5' ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) $false
    if($prof){
      _SetCell ([ref]$s1) 'B7' ("$($prof.score_actual)") $true
      _SetCell ([ref]$s1) 'E7' (("$($prof.banda)").ToUpper()) $false
      _SetCell ([ref]$s1) 'C11' (_Cap ("$($prof.tendencia)")) $false
      _SetCell ([ref]$s1) 'C12' ("$($prof.n_eventos_acumulados)") $true
    }
    $nb=0; if($base -and $base.n_eventos_base){ $nb=[int]$base.n_eventos_base }
    $conf= if($nb -ge 10){"Alta (base de $nb eventos)"} elseif($nb -gt 0){"Media (base en formacion, $nb/10)"} else {"Baja (sin linea base)"}
    _SetCell ([ref]$s1) 'C13' $conf $false
    if($ev.Count -gt 0){
      $u=$ev[$ev.Count-1]
      _SetCell ([ref]$s1) 'C16' ("$($u.event_id)") $true
      _SetCell ([ref]$s1) 'C17' ((("$($u.timestamp)") -replace 'T',' ')) $false
      _SetCell ([ref]$s1) 'C18' ("$($u.causa)") $false
      $tt="$($u.t_transfer_ms)"; if(-not $tt.Trim()){ $tt="$($u.t_retransfer_ms)" }
      _SetCell ([ref]$s1) 'C20' $tt $true
      $vr="$($u.Vu_pre)"; if(-not(_IsNum $vr) -or $vr -eq '0'){ $vr="$($u.Vg_start)" }
      _SetCell ([ref]$s1) 'C21' $vr $true
      _SetCell ([ref]$s1) 'C22' ("I="+("$($u.I_load)")+" A, P="+("$($u.P_load)")+" W") $false
    }
    $prob=""
    if($rend -match '(?s)Causa probable:(.*?)(Conclusion:|Nota:|$)'){
      $prob=((($Matches[1]) -split "`n")|ForEach-Object{$_.Trim()}|Where-Object{$_ -and $_ -ne '--'}) -join ' / '
    }
    _SetCell ([ref]$s1) 'C19' $prob $false
    if($base){
      $fc=$base.fecha_construccion
      if($fc -is [datetime]){ $fc=$fc.ToString('yyyy-MM-dd HH:mm:ss') } else { $fc=("$fc") -replace 'T',' ' }
      $lb="t_transf mediana "+(_Inv $base.t_transfer_mediana_ms)+" ms (IQR "+(_Inv $base.t_transfer_iqr_ms)+") | t_retransf mediana "+(_Inv $base.t_retransfer_mediana_ms)+" ms | Vg mediana "+(_Inv $base.Vg_start_mediana_V)+" V | base de $nb eventos ("+$fc+")"
      _SetCell ([ref]$s1) 'C24' $lb $false
    }
    [System.IO.File]::WriteAllText($s1p,$s1,$enc)

    # ---- HOJA 2: TABLA EVENTOS ----
    $s2p=Join-Path $unz 'xl/worksheets/sheet2.xml'; $s2=[System.IO.File]::ReadAllText($s2p,$enc)
    $hdr2= if($s2 -match '(?s)<sheetData>(<row r="1">.*?</row>)'){ $Matches[1] } else {''}
    $rows=New-Object System.Text.StringBuilder; [void]$rows.Append($hdr2); $rn=1
    foreach($e in $ev){ $rn++
      [void]$rows.Append('<row r="'+$rn+'">')
      [void]$rows.Append((_DataCell 'A' $rn ("$($e.event_id)") $true))
      [void]$rows.Append((_DataCell 'B' $rn ((("$($e.timestamp)") -replace 'T',' ')) $false))
      [void]$rows.Append((_DataCell 'C' $rn ("$($e.causa)") $false))
      [void]$rows.Append((_DataCell 'D' $rn ("$($e.t_transfer_ms)") $true))
      [void]$rows.Append((_DataCell 'E' $rn ("$($e.t_retransfer_ms)") $true))
      [void]$rows.Append((_DataCell 'F' $rn ("$($e.Vu_pre)") $true))
      [void]$rows.Append((_DataCell 'G' $rn ("$($e.Vg_start)") $true))
      [void]$rows.Append((_DataCell 'H' $rn ("$($e.I_load)") $true))
      [void]$rows.Append((_DataCell 'I' $rn ("$($e.P_load)") $true))
      [void]$rows.Append((_DataCell 'J' $rn ("$($e.penaliza)") $true))
      [void]$rows.Append((_DataCell 'K' $rn ("$($e.score)") $true))
      [void]$rows.Append('</row>') }
    $nr2=[Math]::Max(2,$rn)
    $s2=[regex]::Replace($s2,'(?s)<sheetData>.*</sheetData>',('<sheetData>'+$rows.ToString()+'</sheetData>'))
    $s2=$s2 -replace '<dimension ref="A1:K2"/>',('<dimension ref="A1:K'+$nr2+'"/>')
    [System.IO.File]::WriteAllText($s2p,$s2,$enc)
    $t1p=Join-Path $unz 'xl/tables/table1.xml'; $t1=[System.IO.File]::ReadAllText($t1p,$enc)
    $t1=$t1 -replace 'ref="A1:K2"',('ref="A1:K'+$nr2+'"'); [System.IO.File]::WriteAllText($t1p,$t1,$enc)

    # ---- HOJA 3: TABLA SUCESOS ----
    $s3p=Join-Path $unz 'xl/worksheets/sheet3.xml'; $s3=[System.IO.File]::ReadAllText($s3p,$enc)
    $hdr3= if($s3 -match '(?s)<sheetData>(<row r="1">.*?</row>)'){ $Matches[1] } else {''}
    $rows3=New-Object System.Text.StringBuilder; [void]$rows3.Append($hdr3); $rn=1
    foreach($x in $su){ $rn++
      $cambio=("$($x.cond_anterior)")+' -> '+("$($x.cond_nueva)")
      [void]$rows3.Append('<row r="'+$rn+'">')
      [void]$rows3.Append((_DataCell 'A' $rn ((("$($x.timestamp)") -replace 'T',' ')) $false))
      [void]$rows3.Append((_DataCell 'B' $rn ("$($x.fuente)") $false))
      [void]$rows3.Append((_DataCell 'C' $rn $cambio $false))
      [void]$rows3.Append((_DataCell 'D' $rn ("$($x.V)") $true))
      [void]$rows3.Append((_DataCell 'E' $rn ("$($x.Hz)") $true))
      [void]$rows3.Append((_DataCell 'F' $rn ("$($x.duracion_anterior_s)") $true))
      [void]$rows3.Append((_DataCell 'G' $rn ("$($x.diagnostico)") $false))
      [void]$rows3.Append('</row>') }
    $nr3=[Math]::Max(2,$rn)
    $s3=[regex]::Replace($s3,'(?s)<sheetData>.*</sheetData>',('<sheetData>'+$rows3.ToString()+'</sheetData>'))
    $s3=$s3 -replace '<dimension ref="A1:G2"/>',('<dimension ref="A1:G'+$nr3+'"/>')
    [System.IO.File]::WriteAllText($s3p,$s3,$enc)
    $t2p=Join-Path $unz 'xl/tables/table2.xml'; $t2=[System.IO.File]::ReadAllText($t2p,$enc)
    $t2=$t2 -replace 'ref="A1:G2"',('ref="A1:G'+$nr3+'"'); [System.IO.File]::WriteAllText($t2p,$t2,$enc)

    # ---- RECOMPRIMIR ----
    $salidaReal=$Salida
    try{ if(Test-Path $Salida){Remove-Item -LiteralPath $Salida -Force}; $fs=[System.IO.File]::Open($Salida,'Create'); $fs.Close(); Remove-Item -LiteralPath $Salida -Force }
    catch{ $salidaReal=[System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($Salida),[System.IO.Path]::GetFileNameWithoutExtension($Salida)+'_'+(Get-Date -Format 'HHmmss')+'.xlsx') }
    $zip=[System.IO.Compression.ZipFile]::Open($salidaReal,'Create')
    try{
      $b=(Resolve-Path $unz).Path.TrimEnd('\','/')
      Get-ChildItem -LiteralPath $unz -Recurse -File | ForEach-Object{
        $rel=$_.FullName.Substring($b.Length+1) -replace '\\','/'
        $en=$zip.CreateEntry($rel,[System.IO.Compression.CompressionLevel]::Optimal)
        $os=$en.Open(); $by=[System.IO.File]::ReadAllBytes($_.FullName); $os.Write($by,0,$by.Length); $os.Close()
      }
    } finally { $zip.Dispose() }
    return [pscustomobject]@{ Salida=$salidaReal; Eventos=$ev.Count; Sucesos=$su.Count; Score=$(if($prof){$prof.score_actual}else{''}); Banda=$(if($prof){$prof.banda}else{''}) }
  } finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

function Split-Captura([string]$Captura,[string]$Destino){
  if(-not(Test-Path $Captura)){ throw "No encuentro la captura: $Captura" }
  New-Item -ItemType Directory -Path $Destino -Force | Out-Null
  $enc=New-Object System.Text.UTF8Encoding($false)
  $arch=$null; $buf=New-Object System.Collections.Generic.List[string]; $n=0
  foreach($l in [System.IO.File]::ReadAllLines($Captura)){
    $t=$l.TrimEnd("`r")
    if($t -match '^----- INICIO (.+) -----$'){ $arch=$Matches[1].Trim(); $buf.Clear(); continue }
    if($t -match '^----- FIN (.+) -----$'){ if($arch){ [System.IO.File]::WriteAllText((Join-Path $Destino $arch),(($buf -join "`n")+"`n"),$enc); $n++ }; $arch=$null; continue }
    if($t -eq '===== EXPORT ALL COMPLETO ====='){ break }
    if($null -ne $arch){ $buf.Add($t) }
  }
  if($arch -and $buf.Count -gt 0){ [System.IO.File]::WriteAllText((Join-Path $Destino $arch),(($buf -join "`n")+"`n"),$enc); $n++ }
  return $n
}
```

### Lógica clave del generador de reportes

- **Plantilla como ZIP**: un `.xlsx` es internamente un ZIP con XMLs. Se descomprime a un directorio temporal, se editan `sheet1.xml`, `sheet2.xml`, `sheet3.xml` y las tablas, y se recomprime.
- **`_SetCell`**: reemplaza una celda existente en el XML por regex. Detecta si el valor es numérico (`<v>`) o texto (`<is><t>` inline string). Preserva el estilo (`s="N"`).
- **`_DataCell`**: genera XML de celda nueva para las filas de datos de las tablas.
- **Escape XML**: `_XmlEsc` maneja `& < > "` para no romper el XML.
- **Tolerancia**: si un archivo crudo no existe o está corrupto, lo salta sin fallar. Si el xlsx de salida está bloqueado, genera uno con sufijo de hora.
- **Split-Captura**: función auxiliar que toma un volcado serial completo (`EXPORT ALL`) y lo separa en los 6 archivos individuales.

---

## 4. Scripts auxiliares

### `generar_ultimo.ps1` — Regenerar reporte de la última exportación

```powershell
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
```

### `generar_reporte.bat`

```batch
@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0generar_ultimo.ps1"
echo.
pause
```

---

## 5. Formato del volcado `EXPORT ALL` (ejemplo real del módulo)

```
  modulo> [STATUS] score=92 banda=Optimo modbus_RED=OK modbus_GEN=OK rtc=OK sd=OK ev=12
----- INICIO RENDIMIENTO.TXT -----
REPORTE ATS - 2026-06-12 20:34:10

Score: 92/100 (Optimo)
Tendencia: deteriorando
Eventos: 12

Ultimo evento (#12, 2026-06-12 20:33:58):
Transferencia lenta (mecanismo)
t: 7.8 ms

Causa probable:
El ATS tardo mas de lo normal en conmutar

Conclusion:
El ATS tardo mas de lo normal en conmutar: revisar mecanismo y resortes.

Nota: porcentajes = grados de confianza estimados por reglas
(ISO 17359). Error: V/I +/-0.5%, f +/-0.1 Hz, t +/-5 ms.
----- FIN RENDIMIENTO.TXT -----
----- INICIO EVENTS_CURRENT.CSV -----
event_id,timestamp,S1_antes,S1_despues,S2_antes,S2_despues,Vu_pre,Fu_pre,Vg_start,Fg_start,I_load,P_load,t_transfer_ms,t_retransfer_ms,causa,penaliza,score
1,2026-06-11T09:15:22,1,0,0,1,127.1,60.00,127.5,60.00,1.05,93,3.5,,Transferencia normal,0,100
2,2026-06-11T11:02:10,1,0,0,1,126.7,60.01,127.2,59.97,1.10,96,3.7,,Transferencia normal,0,100
3,2026-06-11T13:48:55,0,1,1,0,0.0,59.99,127.5,60.00,1.00,89,,3.9,Retransferencia normal,0,100
4,2026-06-11T16:30:04,1,0,0,1,127.1,60.00,127.2,59.97,1.00,90,3.5,,Transferencia normal,0,100
5,2026-06-11T19:05:41,0,1,1,0,0.0,60.01,127.5,60.00,1.08,93,,4.2,Retransferencia normal,0,100
6,2026-06-12T07:10:18,1,0,0,1,127.5,59.99,127.2,59.97,1.10,96,3.3,,Transferencia normal,0,100
7,2026-06-12T08:42:33,0,1,1,0,0.0,60.00,127.5,60.00,1.04,91,,4.05,Retransferencia normal,0,100
8,2026-06-12T10:15:09,1,0,0,1,126.7,60.01,127.2,59.97,1.00,90,3.7,,Transferencia normal,0,100
9,2026-06-12T12:48:50,0,1,1,0,0.0,59.99,127.5,60.00,1.00,89,,3.9,Retransferencia normal,0,100
10,2026-06-12T15:22:14,1,0,0,1,127.1,60.00,127.2,59.97,1.10,96,3.5,,Transferencia normal,0,100
11,2026-06-12T18:01:37,0,1,1,0,0.0,60.01,127.5,60.00,1.08,93,,4.2,Retransferencia normal,0,100
12,2026-06-12T20:33:58,1,0,0,1,127.5,59.99,127.2,59.97,1.00,90,7.8,,Transferencia lenta (mecanismo),8,92
----- FIN EVENTS_CURRENT.CSV -----
----- INICIO SUCESOS.CSV -----
timestamp,fuente,cond_anterior,cond_nueva,V,Hz,duracion_anterior_s,diagnostico
2026-06-11T09:14:02,RED,DESCONOCIDO,OK,127.4,60.02,,Fuente normalizada
2026-06-12T06:55:20,RED,OK,APAGON,0,0,82800,Apagon de red; el ATS debe transferir a generador
2026-06-12T07:10:18,GEN,OK,OK,127.5,59.97,,Generador asume la carga dentro de banda
2026-06-12T20:30:11,RED,APAGON,OK,127.6,60.01,49793,Retorno de red; el ATS retransfiere
----- FIN SUCESOS.CSV -----
----- INICIO VENTANA.csv -----
event_id,evento,t_opto_ms,fase,offset_s,timestamp,V_RED,I_RED,P_RED,Hz_RED,V_GEN,I_GEN,P_GEN,Hz_GEN
12,TRANSFER,7.8,POST,0.0,2026-06-12T20:33:58,0,0,0,0,127.5,1.05,131,59.97
----- FIN VENTANA.csv -----
----- INICIO PROFILE_HISTORY.JSON -----
{
  "score_actual": 92,
  "banda": "Optimo",
  "tendencia": "deteriorando",
  "n_eventos_acumulados": 12,
  "score_min_historico": 92,
  "score_max_historico": 100,
  "ultimo_event_id": 12,
  "ultima_actualizacion": "2026-06-12T20:34:10"
}
----- FIN PROFILE_HISTORY.JSON -----
----- INICIO BASELINE.JSON -----
{
  "n_eventos_base": 10,
  "t_transfer_mediana_ms": 3.5,
  "t_transfer_iqr_ms": 0.6,
  "t_retransfer_mediana_ms": 4.0,
  "t_retransfer_iqr_ms": 0.5,
  "Vg_start_mediana_V": 127.2,
  "fecha_construccion": "2026-06-11T19:06:00"
}
----- FIN BASELINE.JSON -----
===== EXPORT ALL COMPLETO =====
```

---

## 6. Evolución y decisiones de diseño

### Por qué PowerShell puro (no Python ni Excel COM)

1. **Python** fue la primera implementación (`build_reporte.py` con `openpyxl`). Funcionaba, pero requería que el técnico en campo tuviera Python instalado — inaceptable para un entregable profesional.

2. **Excel COM** (`Excel.Application`, `Workbooks.Open`) nunca se usó: se cuelga por Vista Protegida del ZIP, deja instancias zombie, y depende de tener Excel instalado.

3. **PowerShell puro** con `System.IO.Compression` es la solución final: descomprime el xlsx (que es un ZIP), edita los XMLs internos, y recomprime. Cero dependencias — cualquier Windows lo corre.

### El fix de la microSD (causa raíz de "no exportaba")

El firmware V4 originalmente arrancaba la SD a 20 MHz, lo cual fallaba con cables jumper largos. El fix: `montarSD()` prueba velocidades escalonadas (400 kHz → 1 → 4 → 10 MHz) con `SD.end()` + `delay(300)` entre intentos. Además, antes de `spiSD.begin`: `pinMode(PIN_SD_CS,OUTPUT); digitalWrite(PIN_SD_CS,HIGH); delay(100)`.

### Hardening (25-jun-2026)

- `Generar-ReporteATS.ps1` blindado: coerción segura de tipos, lectura tolerante a BOM/JSON corrupto/CSV incompleto, nombre alterno si xlsx bloqueado.
- `exportar_ats.ps1`: watchdog de 90s, `try/finally` que siempre cierra el puerto, guarda crudos PRIMERO.
- 13/13 pruebas adversas PASA (sin profile, 0 eventos, baseline corrupto, caracteres XML especiales).

---

## Inventario de archivos del kit

| Archivo | Función |
|---|---|
| `exportar_ats.bat` | Punto de entrada (doble clic) |
| `exportar_ats.ps1` | Captura serial + genera reporte |
| `Generar-ReporteATS.ps1` | Motor: crudos → REPORTE_ATS.xlsx |
| `plantilla_reporte_ats.xlsx` | Plantilla con formato y estructura |
| `generar_reporte.bat` | Regenerar xlsx de la última exportación |
| `generar_ultimo.ps1` | Busca última carpeta y regenera |
| `Analizar-Predictivo.ps1` | Análisis predictivo del historial de perfiles |
| `analizar_predictivo.bat` | Punto de entrada del análisis predictivo |
| `REPORTE_ATS (ejemplo).xlsx` | Salida de ejemplo ya generada |
| `LEEME.txt` | Instrucciones de uso |
