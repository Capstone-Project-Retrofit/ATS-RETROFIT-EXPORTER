<#
  Analizar-Predictivo.ps1 - Analisis Predictivo de Degradacion del ATS.
  Modulo Retrofit No Invasivo para ATS Legacy - INTEC Capstone Project II

  Corre en la PC (PowerShell puro, sin instalar nada) sobre el EVENTS_CURRENT.CSV
  de una exportacion. NO toca el firmware ni el modulo. Genera:
    - ANALISIS_PREDICTIVO_<carpeta>.xlsx  (titulo, diagnostico por direccion y
      curvas de score y tiempos, desde plantilla_predictivo.xlsx)
    - PREDICTIVO_PC.txt                   (mismo analisis en texto)

  Pipeline por direccion (transferencia RED->GEN y retransferencia GEN->RED):
    1. Estadistica robusta (mediana e IQR exactos) de los tiempos de conmutacion.
    2. Sigma de RUIDO por diferencias sucesivas (mediana|dt|/0.954).
    3. Filtro de Kalman de 2 estados [nivel, pendiente].
    4. CUSUM unilateral superior (k=0.5, h=5; Hawkins & Olwell 1998).
    5. RUL (vida util remanente) por cruce del umbral 1.5 x mediana base,
       con intervalo de confianza 90% (metodo delta). ISO 17359.

  Concepto del nucleo matematico: Edongy Ramirez. Adaptacion a PC: equipo.

  Uso:  doble clic a analizar_predictivo.bat  (usa la exportacion MAS RECIENTE)
        o  Invoke-AnalisisPredictivo -Crudos "C:\...\ATS_EXPORT\<carpeta>"
#>

function _MedianaPS([double[]]$v) {
  if ($v.Count -eq 0) { return [double]::NaN }
  $s = $v | Sort-Object
  $n = $s.Count
  if ($n % 2 -eq 1) { return [double]$s[[int][math]::Floor($n/2)] }
  return ([double]$s[$n/2 - 1] + [double]$s[$n/2]) / 2.0
}
function _IqrPS([double[]]$v) {
  if ($v.Count -lt 4) { return 0.0 }
  $s = @($v | Sort-Object); $n = $s.Count
  function _q($p) {
    $pos = $p * ($n - 1); $i = [int][math]::Floor($pos); $fr = $pos - $i
    $j = [math]::Min($i + 1, $n - 1)
    return [double]$s[$i] + $fr * ([double]$s[$j] - [double]$s[$i])
  }
  return (_q 0.75) - (_q 0.25)
}
function _NumPS([string]$x) {
  $r = 0.0
  if ([double]::TryParse($x, [Globalization.NumberStyles]::Float,
      [Globalization.CultureInfo]::InvariantCulture, [ref]$r)) { return $r }
  return [double]::NaN
}

function _AnalizarSerie([string]$nombre, [double[]]$serie, [datetime[]]$fechas, [System.Text.StringBuilder]$sb) {
  $Z90 = 1.6449; $FACTOR = 1.5; $CK = 0.5; $CH = 5.0
  $QL = 1e-4; $QS = 1e-6; $WARMUP = 5
  $n = $serie.Count
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("== $nombre ($n eventos) ==")
  if ($n -lt 6) {
    [void]$sb.AppendLine("   Datos insuficientes para estimar tendencia (minimo 6 eventos).")
    return @{ Ok=$false; Resumen="Datos insuficientes (menos de 6 eventos con tiempo valido)."
              Base=[double]::NaN; Umbral=[double]::NaN; Nivel=[double]::NaN; Tend=[double]::NaN
              Vida="--" }
  }
  $nBase  = [math]::Max(4, [math]::Min(8, [int][math]::Floor($n/2)))
  $mBase  = _MedianaPS ($serie[0..($nBase-1)])
  $umbral = $FACTOR * $mBase
  $difs = New-Object System.Collections.Generic.List[double]
  for ($i = 1; $i -lt $n; $i++) { $difs.Add([math]::Abs($serie[$i] - $serie[$i-1])) }
  $sigR = (_MedianaPS $difs.ToArray()) / 0.954
  if ($sigR -lt 1e-3) { $sigR = [math]::Max((_IqrPS $serie) * 0.7413, 1e-3) }
  $R = $sigR * $sigR
  $L = $serie[0]; $S = 0.0
  $P00 = $R * 10.0; $P01 = 0.0; $P11 = 1e-2
  $C = 0.0; $alarma = $false; $alarmaEn = 0
  for ($k = 1; $k -lt $n; $k++) {
    $Lp = $L + $S
    $P00p = $P00 + 2*$P01 + $P11 + $QL
    $P01p = $P01 + $P11
    $P11p = $P11 + $QS
    $nu = $serie[$k] - $Lp
    $Sv = $P00p + $R
    $zi = $nu / [math]::Sqrt($Sv)
    if ([math]::Abs($zi) -gt 4.0) { $Sv = $P00p + $R * 25.0; $zi = $nu / [math]::Sqrt($Sv) }
    $K0 = $P00p / $Sv; $K1 = $P01p / $Sv
    $L = $Lp + $K0 * $nu; $S = $S + $K1 * $nu
    $P00 = (1 - $K0) * $P00p; $P01 = (1 - $K0) * $P01p; $P11 = $P11p - $K1 * $P01p
    if ($k -gt $WARMUP) {
      $C = [math]::Max(0.0, $C + $zi - $CK)
      if (($C -gt $CH) -and (-not $alarma)) { $alarma = $true; $alarmaEn = $k + 1 }
    }
  }
  $sdS = [math]::Sqrt([math]::Max($P11, 1e-12))
  [void]$sb.AppendLine(("   Tiempo tipico (mediana base): {0:N2} ms   Umbral de atencion: {1:N2} ms" -f $mBase, $umbral))
  [void]$sb.AppendLine(("   Nivel actual estimado: {0:N2} ms   Tendencia: {1:N4} ms/evento" -f $L, $S))
  if ($alarma) {
    [void]$sb.AppendLine("   [ALERTA] Aceleracion SOSTENIDA del desgaste detectada (CUSUM, evento #$alarmaEn).")
  } else {
    [void]$sb.AppendLine(("   Sin aceleracion sostenida de desgaste (CUSUM {0:N2} de {1:N0})." -f $C, $CH))
  }
  $dtD = [double]::NaN
  for ($i = 1; $i -lt $fechas.Count; $i++) {
    if ($fechas[$i] -le [datetime]::MinValue -or $fechas[$i-1] -le [datetime]::MinValue) { continue }
    $d = ($fechas[$i] - $fechas[$i-1]).TotalDays
    if ($d -gt 1e-6 -and $d -lt 90) {
      if ([double]::IsNaN($dtD)) { $dtD = $d } else { $dtD = $dtD + 0.15 * ($d - $dtD) }
    }
  }
  $vida = "Sin tendencia de desgaste significativa. El ATS esta estable."
  if ($S -gt $Z90 * $sdS) {
    $margen = $umbral - $L
    if ($margen -le 0) {
      [void]$sb.AppendLine("   VIDA UTIL: el tiempo de conmutacion YA alcanzo 1.5x su valor base.")
      [void]$sb.AppendLine("   Recomendacion: inspeccion del mecanismo y resortes del ATS.")
      $vida = "El tiempo de conmutacion YA alcanzo 1.5x su valor base. Inspeccion del mecanismo y resortes recomendada."
    } else {
      $n0 = $margen / $S
      $varN = [math]::Max(0.0, ($P00 + $n0*$n0*$P11 + 2*$n0*$P01) / ($S*$S))
      $sdN = [math]::Sqrt($varN)
      $lo = [math]::Max(0.0, $n0 - $Z90 * $sdN); $hi = $n0 + $Z90 * $sdN
      $lin = ("   VIDA UTIL ESTIMADA: ~{0:N0} transferencias mas (rango 90%: {1:N0} a {2:N0})" -f $n0, $lo, $hi)
      $vida = ("~{0:N0} transferencias mas (rango 90%: {1:N0} a {2:N0})" -f $n0, $lo, $hi)
      if (-not [double]::IsNaN($dtD)) {
        $lin  += (" | ~{0:N0} dias ({1:N0} a {2:N0})" -f ($n0*$dtD), ($lo*$dtD), ($hi*$dtD))
        $vida += (" | ~{0:N0} dias" -f ($n0*$dtD))
      }
      [void]$sb.AppendLine($lin)
      [void]$sb.AppendLine("   (Estimacion estadistica: se refina sola con cada evento nuevo.)")
      $vida += "."
    }
  } else {
    [void]$sb.AppendLine("   VIDA UTIL: sin tendencia de desgaste significativa. El ATS esta estable.")
  }
  $res = "Estable: sin tendencia de desgaste significativa."
  if ($S -gt $Z90 * $sdS) {
    if (($umbral - $L) -le 0) {
      $res = "Tendencia de desgaste: el tiempo YA alcanzo 1.5x su base -> inspeccionar mecanismo y resortes."
    } else {
      $res = "Tendencia de desgaste detectada: vigilar. " + $vida
    }
  }
  if ($alarma) { $res = "ALERTA (evento #" + $alarmaEn + "): deterioro sostenido confirmado. " + $res }
  return @{ Ok=$true; Resumen=$res; Base=$mBase; Umbral=$umbral; Nivel=$L; Tend=$S; Vida=$vida }
}

function _PXml([string]$s){ if($null -eq $s){return ""}; $s=$s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'; return $s }
function _PIsNum([string]$s){ return ($null -ne $s -and $s -match '^\s*-?\d+(\.\d+)?\s*$') }
function _PSetCell([string]$xml,[string]$ref,[string]$value,[bool]$isNum){
  $pat='<c r="'+[regex]::Escape($ref)+'"( s="\d+")?[^>]*>.*?</c>'
  if($isNum -and -not (_PIsNum $value)){ $isNum=$false }
  if($isNum){ $rep='<c r="'+$ref+'"$1><v>'+$value.Trim()+'</v></c>' }
  else      { $rep='<c r="'+$ref+'"$1 t="inlineStr"><is><t xml:space="preserve">'+(_PXml $value)+'</t></is></c>' }
  return [regex]::Replace($xml,$pat,$rep,[System.Text.RegularExpressions.RegexOptions]::Singleline)
}
function _PDataCell([string]$col,[int]$row,[string]$value,[bool]$isNum){
  $r=$col+$row
  if([string]::IsNullOrWhiteSpace($value)){ return '<c r="'+$r+'"/>' }
  if($isNum -and (_PIsNum $value)){ return '<c r="'+$r+'" t="n"><v>'+$value.Trim()+'</v></c>' }
  return '<c r="'+$r+'" t="inlineStr"><is><t xml:space="preserve">'+(_PXml $value)+'</t></is></c>'
}
function _PNumTxt([double]$v,[string]$fmt){
  if([double]::IsNaN($v)){ return "--" }
  return [string]::Format([Globalization.CultureInfo]::InvariantCulture,$fmt,$v)
}

function Add-PredictivoAlReporte([string]$Xlsx,$RT,$RR){
  Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
  $tmp=Join-Path ([System.IO.Path]::GetTempPath()) ("atsp_"+[System.Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tmp -Force | Out-Null
  try{
    $unz=Join-Path $tmp 'x'
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Xlsx,$unz)
    $enc=New-Object System.Text.UTF8Encoding($false)
    $s1p=Join-Path $unz 'xl/worksheets/sheet1.xml'
    if(Test-Path $s1p){
      $s1=[System.IO.File]::ReadAllText($s1p,$enc)
      $txT="$($RT.Resumen)"; if("$($RT.Vida)" -like '~*'){ $txT=$txT+" Vida util: "+$RT.Vida }
      $txR="$($RR.Resumen)"; if("$($RR.Vida)" -like '~*'){ $txR=$txR+" Vida util: "+$RR.Vida }
      $s1=_PSetCell $s1 'C31' $txT $false
      $s1=_PSetCell $s1 'C32' $txR $false
      [System.IO.File]::WriteAllText($s1p,$s1,$enc)
      $out2=$Xlsx+'.tmp'
      if(Test-Path $out2){ Remove-Item -LiteralPath $out2 -Force }
      $zip=[System.IO.Compression.ZipFile]::Open($out2,'Create')
      try{
        $b=(Resolve-Path $unz).Path.TrimEnd('\','/')
        foreach($f in [System.IO.Directory]::GetFiles($unz,'*',[System.IO.SearchOption]::AllDirectories)){
          $rel=$f.Substring($b.Length+1) -replace '\\','/'
          $en2=$zip.CreateEntry($rel,[System.IO.Compression.CompressionLevel]::Optimal)
          $os=$en2.Open(); $by=[System.IO.File]::ReadAllBytes($f); $os.Write($by,0,$by.Length); $os.Close()
        }
      } finally { $zip.Dispose() }
      Move-Item -LiteralPath $out2 -Destination $Xlsx -Force
    }
  } finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

function Invoke-AnalisisPredictivo {
  param([string]$Crudos = "", [string]$Xlsx = "")
  if (-not $Crudos) {
    $raiz = Join-Path $env:USERPROFILE "Documents\ATS_EXPORT"
    if (-not (Test-Path $raiz)) { $raiz = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "ATS_EXPORT" }
    if (Test-Path $raiz) {
      $ult = Get-ChildItem $raiz -Directory | Sort-Object Name -Descending | Select-Object -First 1
      if ($ult) { $Crudos = $ult.FullName }
    }
  }
  if (-not $Crudos -or -not (Test-Path $Crudos)) { throw "No encuentro la carpeta de exportacion. Use -Crudos <carpeta>." }
  $csv = Join-Path $Crudos "EVENTS_CURRENT.CSV"
  if (-not (Test-Path $csv)) { throw "No existe EVENTS_CURRENT.CSV en: $Crudos" }
  $rows = @(Import-Csv -LiteralPath $csv)

  $serT = New-Object System.Collections.Generic.List[double]
  $fT   = New-Object System.Collections.Generic.List[datetime]
  $serR = New-Object System.Collections.Generic.List[double]
  $fR   = New-Object System.Collections.Generic.List[datetime]
  foreach ($r in $rows) {
    $ts = [datetime]::MinValue
    [void][datetime]::TryParse(("$($r.timestamp)" -replace 'T',' '),
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None, [ref]$ts)
    $t  = _NumPS "$($r.t_transfer_ms)"
    $tr = _NumPS "$($r.t_retransfer_ms)"
    if (-not [double]::IsNaN($t))  { $serT.Add($t);  $fT.Add($ts) }
    if (-not [double]::IsNaN($tr)) { $serR.Add($tr); $fR.Add($ts) }
  }

  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("ANALISIS PREDICTIVO DEL ATS - " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
  [void]$sb.AppendLine("Fuente: $csv ($($rows.Count) eventos)")
  [void]$sb.AppendLine("Metodo: Kalman + CUSUM + RUL con IC 90% sobre los tiempos de conmutacion (ISO 17359).")
  $resT = _AnalizarSerie "Transferencia RED->GEN"    $serT.ToArray() $fT.ToArray() $sb
  $resR = _AnalizarSerie "Retransferencia GEN->RED"  $serR.ToArray() $fR.ToArray() $sb
  [void]$sb.AppendLine("")
  [void]$sb.AppendLine("Nota: analisis calculado en la PC sobre la copia exportada; la evidencia")
  [void]$sb.AppendLine("original permanece en la microSD del modulo (offline-first).")

  $txt = $sb.ToString()
  $out = Join-Path $Crudos "PREDICTIVO_PC.txt"
  [System.IO.File]::WriteAllText($out, $txt, (New-Object System.Text.UTF8Encoding($false)))
  Write-Host $txt
  Write-Host "Guardado en: $out" -ForegroundColor Green
  if (-not $Xlsx) {
    $cand = Get-ChildItem -LiteralPath $Crudos -Filter 'REPORTE_ATS*.xlsx' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($cand) { $Xlsx = $cand.FullName }
  }
  if ($Xlsx -and (Test-Path $Xlsx)) {
    try { Add-PredictivoAlReporte $Xlsx $resT $resR; Write-Host "Diagnostico predictivo agregado al reporte." -ForegroundColor Green }
    catch { Write-Host ("[AVISO] No se pudo escribir el diagnostico en el reporte (cierra el Excel si lo tienes abierto): " + $_.Exception.Message) -ForegroundColor Yellow }
  }
  return $out
}

if ($MyInvocation.InvocationName -ne '.') {
  try {
    Invoke-AnalisisPredictivo | Out-Null
    $raiz = Join-Path $env:USERPROFILE "Documents\ATS_EXPORT"
    if (-not (Test-Path $raiz)) { $raiz = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "ATS_EXPORT" }
    $ult = Get-ChildItem $raiz -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    if ($ult) {
      $px = Get-ChildItem -LiteralPath $ult.FullName -Filter 'REPORTE_ATS*.xlsx' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
      if ($px) { try { Invoke-Item $px.FullName } catch {} }
    }
  } catch { Write-Host "ERROR: $_" -ForegroundColor Red }
}
