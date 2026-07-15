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
      $lect="I="+("$($u.I_load)")+" A, P="+("$($u.P_load)")+" W"
      $fpl="$($u.FP_load)".Trim()
      if($fpl){ $lect=$lect+", FP="+$fpl }
      _SetCell ([ref]$s1) 'C22' $lect $false
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
      [void]$rows.Append((_DataCell 'L' $rn ("$($e.S_load_VA)") $true))
      [void]$rows.Append((_DataCell 'M' $rn ("$($e.Q_load_var)") $true))
      [void]$rows.Append((_DataCell 'N' $rn ("$($e.FP_load)") $true))
      [void]$rows.Append((_DataCell 'O' $rn ("$($e.t_sin_suministro_s)") $true))
      [void]$rows.Append('</row>') }
    $nr2=[Math]::Max(2,$rn)
    $s2=[regex]::Replace($s2,'(?s)<sheetData>.*</sheetData>',('<sheetData>'+$rows.ToString()+'</sheetData>'))
    $s2=$s2 -replace '<dimension ref="A1:O2"/>',('<dimension ref="A1:O'+$nr2+'"/>')
    [System.IO.File]::WriteAllText($s2p,$s2,$enc)
    $t1p=Join-Path $unz 'xl/tables/table1.xml'; $t1=[System.IO.File]::ReadAllText($t1p,$enc)
    $t1=$t1 -replace 'ref="A1:O2"',('ref="A1:O'+$nr2+'"'); [System.IO.File]::WriteAllText($t1p,$t1,$enc)
    # GRAFICA: ajustar el rango de las curvas (score y tiempos) al numero real de eventos
    $chDir=Join-Path $unz 'xl/charts'
    if(Test-Path $chDir){
      foreach($cf in [System.IO.Directory]::GetFiles($chDir,'chart*.xml')){
        $cx=[System.IO.File]::ReadAllText($cf,$enc)
        $cx=$cx.Replace('$1001',('$'+$nr2))
        [System.IO.File]::WriteAllText($cf,$cx,$enc)
      }
    }

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
      [void]$rows3.Append((_DataCell 'H' $rn ("$($x.FP)") $true))
      [void]$rows3.Append('</row>') }
    $nr3=[Math]::Max(2,$rn)
    $s3=[regex]::Replace($s3,'(?s)<sheetData>.*</sheetData>',('<sheetData>'+$rows3.ToString()+'</sheetData>'))
    $s3=$s3 -replace '<dimension ref="A1:H2"/>',('<dimension ref="A1:H'+$nr3+'"/>')
    [System.IO.File]::WriteAllText($s3p,$s3,$enc)
    $t2p=Join-Path $unz 'xl/tables/table2.xml'; $t2=[System.IO.File]::ReadAllText($t2p,$enc)
    $t2=$t2 -replace 'ref="A1:H2"',('ref="A1:H'+$nr3+'"'); [System.IO.File]::WriteAllText($t2p,$t2,$enc)

    # ---- RECOMPRIMIR ----
    $salidaReal=$Salida
    try{ if(Test-Path $Salida){Remove-Item -LiteralPath $Salida -Force}; $fs=[System.IO.File]::Open($Salida,'Create'); $fs.Close(); Remove-Item -LiteralPath $Salida -Force }
    catch{ $salidaReal=[System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($Salida),[System.IO.Path]::GetFileNameWithoutExtension($Salida)+'_'+(Get-Date -Format 'HHmmss')+'.xlsx') }
    $zip=[System.IO.Compression.ZipFile]::Open($salidaReal,'Create')
    try{
      $b=(Resolve-Path $unz).Path.TrimEnd('\','/')
      # FIX (30-jun): GetFiles incluye TODOS los archivos -incluido _rels\.rels-, sin el filtro
      # de atributos de Get-ChildItem que botaba el _rels\.rels y hacia que Excel reparara el xlsx.
      foreach($f in [System.IO.Directory]::GetFiles($unz,'*',[System.IO.SearchOption]::AllDirectories)){
        $rel=$f.Substring($b.Length+1) -replace '\\','/'
        $en=$zip.CreateEntry($rel,[System.IO.Compression.CompressionLevel]::Optimal)
        $os=$en.Open(); $by=[System.IO.File]::ReadAllBytes($f); $os.Write($by,0,$by.Length); $os.Close()
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
