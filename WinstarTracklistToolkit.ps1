# Winstar Quang MP3 Toolkit — 2025-11-14
# Fixes: top-right theme (global), dim-dark palette, checkbox @ col#1, play-on-title-only,
# Random(Preview) shuffles only, double-buffered LV, tab icons restored.

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try { $host.Runspace.ApartmentState = 'STA' } catch {}
[System.Windows.Forms.Application]::EnableVisualStyles()

[Console]::InputEncoding  = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---------------- License ----------------
$scriptDir   = Split-Path -Parent $PSCommandPath
$assets      = Join-Path $scriptDir 'assets'
$licenseFile = Join-Path $scriptDir ".license"

function Get-MachineGuid {
  try {
    $k="HKLM:\SOFTWARE\Microsoft\Cryptography"
    $g=(Get-ItemProperty -Path $k -ErrorAction Stop).MachineGuid
    if(-not $g){ $g=(Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Cryptography" -ErrorAction Stop).MachineGuid }
    $g.Trim()
  } catch { $null }
}
$actualGuid = Get-MachineGuid
if (-not $actualGuid) {
  [System.Windows.Forms.MessageBox]::Show("Không thể đọc key cấp quyền.","Error",
    [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null; exit
}
if (Test-Path $licenseFile) {
  $stored = (Get-Content $licenseFile -Raw).Trim() -replace '\s',''
  if ($stored -ne $actualGuid) {
    [System.Windows.Forms.MessageBox]::Show(
      "Key này đã được truy cập trên máy khác. Vui lòng liên hệ Winstar Quang để được kích hoạt.",
      "Không có quyền hợp lệ",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Stop) | Out-Null
    $input = [Microsoft.VisualBasic.Interaction]::InputBox("Hãy nhập key license để truy cập ứng dụng","Liên hệ Winstar Quang để được hỗ trợ","").Trim()
    if ($input -eq $actualGuid) {
      Set-Content -Path $licenseFile -Value $input -Encoding UTF8
      [System.Windows.Forms.MessageBox]::Show("Truy cập thành công!","",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } else {
      [System.Windows.Forms.MessageBox]::Show("Key không chính xác. Ứng dụng sẽ thoát!!!.","",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null; exit
    }
  }
} else {
  $input = [Microsoft.VisualBasic.Interaction]::InputBox("Hãy nhập key license để truy cập ứng dụng","Liên hệ Winstar Quang để được hỗ trợ","").Trim()
  if ($input -eq $actualGuid) {
    Set-Content -Path $licenseFile -Value $input -Encoding UTF8
    [System.Windows.Forms.MessageBox]::Show("Truy cập thành công!","",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  } else {
    [System.Windows.Forms.MessageBox]::Show("Key không chính xác. Ứng dụng sẽ thoát!!!.","",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null; exit
  }
}

# ---------------- FFmpeg locate ----------------
function Find-Exe([string]$exeName){
  $local = Join-Path $scriptDir $exeName
  if (Test-Path $local) { return $local }
  $p=(Get-Command $exeName -ErrorAction SilentlyContinue).Path
  if($p){ return $p }
  $null
}
$ffmpeg  = Find-Exe "ffmpeg.exe"
$ffprobe = Find-Exe "ffprobe.exe"
if (-not $ffmpeg -or -not $ffprobe) {
  [System.Windows.Forms.MessageBox]::Show(
    "ffmpeg/ffprobe không tìm thấy. Hãy đặt cùng thư mục với script hoặc thêm vào PATH.",
    "Thiếu FFmpeg",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null; exit
}

# ---------------- Helpers ----------------
function Get-DurationSec([string]$file){
  $psi=New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName=$ffprobe
  $psi.Arguments="-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 `"$file`""
  $psi.CreateNoWindow=$true; $psi.UseShellExecute=$false; $psi.RedirectStandardOutput=$true
  $p=[System.Diagnostics.Process]::Start($psi)
  $out=$p.StandardOutput.ReadToEnd().Trim(); $p.WaitForExit()
  try { [double]::Parse($out,[System.Globalization.CultureInfo]::InvariantCulture) } catch { 0 }
}
function SecToHMS([double]$sec){ $ts=[TimeSpan]::FromSeconds([math]::Floor([math]::Max($sec,0))); "{0:00}:{1:00}:{2:00}" -f $ts.Hours,$ts.Minutes,$ts.Seconds }
function Parse-Time([string]$t){ $t=$t.Trim().Trim('[',']'); if($t -notmatch '^\d{2}:\d{2}:\d{2}$'){throw "Thời gian không hợp lệ: $t"}; $h,$m,$s=$t.Split(':'); [int]$h*3600+[int]$m*60+[int]$s }
function Fmt-HMS([int]$s){ if($s -lt 0){$s=0}; '{0:00}:{1:00}:{2:00}' -f ([math]::Floor($s/3600)),([math]::Floor(($s%3600)/60)),($s%60) }
function Fmt-SRT([double]$s){ if($s -lt 0){$s=0}; $hh=[math]::Floor($s/3600);$mm=[math]::Floor(($s%3600)/60);$ss=[math]::Floor($s%60);$ms=[math]::Round(($s-[math]::Floor($s))*1000)
  '{0:00}:{1:00}:{2:00},{3:000}' -f $hh,$mm,$ss,$ms }
function Merge-Cuts($cuts){ $cuts=$cuts|Sort-Object Start; $m=New-Object 'System.Collections.Generic.List[object]'; foreach($c in $cuts){ if($m.Count -eq 0 -or $c.Start -gt $m[-1].End){$m.Add([pscustomobject]@{Start=$c.Start;End=$c.End})} elseif($c.End -gt $m[-1].End){$m[-1].End=$c.End} }; $m }
function Total-CutBefore([int]$T,$cuts){ $tot=0; foreach($c in $cuts){ if($T -le $c.Start){break}; $tot += [Math]::Max(0,[Math]::Min($T,$c.End)-$c.Start) }; $tot }
function Read-Tracklist([string]$p){
  $items=New-Object 'System.Collections.Generic.List[object]'
  foreach($line in (Get-Content -LiteralPath $p -Encoding UTF8)){
    if([string]::IsNullOrWhiteSpace($line)){continue}
    $m=[regex]::Match($line,'^\s*\[?(?<h>\d{2}):(?<m>\d{2}):(?<s>\d{2})\]?\s*(?:-\s*)?(?<title>.*)\s*$')
    if($m.Success){ $t=([int]$m.Groups['h'].Value)*3600+([int]$m.Groups['m'].Value)*60+([int]$m.Groups['s'].Value); $items.Add([pscustomobject]@{Time=$t;Title=$m.Groups['title'].Value.Trim()}) }
  } ($items|Sort-Object Time)
}
function Build-SRT($items,[string]$out){
  $TAIL=180; $lines=New-Object System.Collections.Generic.List[string]
  for($i=0;$i -lt $items.Count;$i++){
    $start=[double]$items[$i].Time
    $end= if($i -lt $items.Count-1){ [double]$items[$i+1].Time - 0.001 } else { $start+$TAIL }
    $lines.Add(($i+1)); $lines.Add(("{0} --> {1}" -f (Fmt-SRT $start),(Fmt-SRT $end))); $lines.Add($items[$i].Title); $lines.Add("")
  }
  [IO.File]::WriteAllLines($out,$lines,[Text.UTF8Encoding]::new($false))
}

function New-TrackObject {
  param(
    [string]$Path,
    [string]$Title,
    [double]$DurSec,
    [bool]$Star=$false
  )
  [pscustomobject]@{
    Path   = $Path
    Title  = $Title
    DurSec = [double]$DurSec
    Star   = [bool]$Star
  }
}

function Copy-TrackList {
  param([System.Collections.IEnumerable]$items)
  $copy=[System.Collections.Generic.List[psobject]]::new()
  foreach($t in $items){
    $copy.Add((New-TrackObject -Path $t.Path -Title $t.Title -DurSec $t.DurSec -Star $t.Star))
  }
  return $copy
}

# ---------------- Main Form ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text="Winstar Tracklist Toolkit"
$form.StartPosition="CenterScreen"
$form.Size=[Drawing.Size]::new(1120,700)
$form.MinimumSize=[Drawing.Size]::new(1000,620)
$form.KeyPreview=$true

# Top AppBar (chứa Theme, nằm góc phải trên cùng)
$topBar = New-Object System.Windows.Forms.Panel
$topBar.Dock='Top'; $topBar.Height=44; $topBar.Padding='10,6,10,6'
$form.Controls.Add($topBar)
$topRight = New-Object System.Windows.Forms.FlowLayoutPanel
$topRight.Dock='Right'
$topRight.FlowDirection='LeftToRight'
$topRight.WrapContents=$false
$topRight.AutoSize=$true
$topRight.AutoSizeMode='GrowAndShrink'
$topRight.Margin='0,0,0,0'
$topRight.Padding='0,0,0,0'
$topBar.Controls.Add($topRight)

$lblThemeTop = New-Object System.Windows.Forms.Label
$lblThemeTop.Text="Theme"; $lblThemeTop.AutoSize=$true; $lblThemeTop.Margin='0,6,4,0'
$cmbThemeTop = New-Object System.Windows.Forms.ComboBox
$cmbThemeTop.DropDownStyle='DropDownList'; $cmbThemeTop.Width=140; $cmbThemeTop.Margin='0,2,12,0'
[void]$cmbThemeTop.Items.AddRange(@('Calm Gray','Light','Dark'))
$cmbThemeTop.SelectedIndex=0
$btnResetTop = New-Object System.Windows.Forms.Button
$btnResetTop.Text="Reset"; $btnResetTop.Width=90; $btnResetTop.Margin='0,0,8,0'
$btnCloseTop = New-Object System.Windows.Forms.Button
$btnCloseTop.Text="Close"; $btnCloseTop.Width=90
$topRight.Controls.AddRange(@($lblThemeTop,$cmbThemeTop,$btnResetTop,$btnCloseTop))
$btnCloseTop.Add_Click({ $form.Close() })

# Tabs + Icons
$tabs=New-Object System.Windows.Forms.TabControl
$tabs.Dock='Fill'
$tabs.Multiline=$true
$tabs.SizeMode=[System.Windows.Forms.TabSizeMode]::Fixed
$tabs.ItemSize=[Drawing.Size]::new(180,32)
$tabs.Appearance=[System.Windows.Forms.TabAppearance]::Normal
$tabs.Padding=[Drawing.Point]::new(18,6)
$tabs.DrawMode=[System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$form.Controls.Add($tabs)

$TabImages = New-Object System.Windows.Forms.ImageList
$TabImages.ColorDepth='Depth32Bit'
$TabImages.ImageSize  = [Drawing.Size]::new(24,24)
function Add-TabImage([string]$key,[string]$baseName){
  $ico = Join-Path $assets ("{0}.ico" -f $baseName)
  $png = Join-Path $assets ("{0}.png" -f $baseName)
  if(Test-Path $ico){ $img=(New-Object System.Drawing.Icon($ico,24,24)).ToBitmap(); [void]$TabImages.Images.Add($key,$img); return }
  if(Test-Path $png){ $img=[Drawing.Image]::FromFile($png); [void]$TabImages.Images.Add($key,$img); return }
}
Add-TabImage 'create' 'tab_create'
Add-TabImage 'adjust' 'tab_adjust'
$tabs.ImageList=$TabImages

$script:Colors=@{}

function Set-ThemeControl($ctrl,$panelBg,$fg,$inputBg){
  if ($ctrl -is [System.Windows.Forms.Panel] -or
      $ctrl -is [System.Windows.Forms.TabPage] -or
      $ctrl -is [System.Windows.Forms.TableLayoutPanel] -or
      $ctrl -is [System.Windows.Forms.FlowLayoutPanel]) {
    $ctrl.BackColor=$panelBg; $ctrl.ForeColor=$fg
  } elseif ($ctrl -is [System.Windows.Forms.TabControl]) {
    $ctrl.BackColor=$panelBg; $ctrl.ForeColor=$fg
  } elseif ($ctrl -is [System.Windows.Forms.TextBox] -or $ctrl -is [System.Windows.Forms.ComboBox]) {
    $ctrl.BackColor=$inputBg; $ctrl.ForeColor=$fg
  } elseif ($ctrl -is [System.Windows.Forms.Button]) {
    $ctrl.UseVisualStyleBackColor=$false; $ctrl.BackColor=$panelBg; $ctrl.ForeColor=$fg
  } elseif ($ctrl -is [System.Windows.Forms.ListView]) {
    $ctrl.BackColor=$panelBg; $ctrl.ForeColor=$fg
  } elseif ($ctrl -is [System.Windows.Forms.Label] -or $ctrl -is [System.Windows.Forms.CheckBox] -or $ctrl -is [System.Windows.Forms.RadioButton]) {
    $ctrl.ForeColor=$fg
  }
  foreach($c in $ctrl.Controls){ Set-ThemeControl $c $panelBg $fg $inputBg }
}

function Apply-Theme([string]$mode){
  switch -Regex ($mode) {
    '^(dark|tối)$' {
      $script:Colors=@{
        Panel=[Drawing.Color]::FromArgb(52,55,60)
        Input=[Drawing.Color]::FromArgb(60,63,68)
        Fg   =[Drawing.Color]::FromArgb(235,235,235)
        Grid =[Drawing.Color]::FromArgb(96,96,100)
        Row  =[Drawing.Color]::FromArgb(54,57,62)
        RowAlt=[Drawing.Color]::FromArgb(58,61,66)
        RowSel=[Drawing.Color]::FromArgb(85,130,200)
        RowSelInactive=[Drawing.Color]::FromArgb(76,80,86)
        SelText=[Drawing.Color]::White
        HeaderBg=[Drawing.Color]::FromArgb(60,63,68)
        HeaderFg=[Drawing.Color]::FromArgb(240,240,240)
        TabActiveBg=[Drawing.Color]::FromArgb(70,73,78)
        TabInactiveBg=[Drawing.Color]::FromArgb(58,61,66)
        TabActiveFg=[Drawing.Color]::FromArgb(240,240,240)
        TabInactiveFg=[Drawing.Color]::FromArgb(190,190,190)
        TabBorder=[Drawing.Color]::FromArgb(40,42,46)
      }
      break
    }
    '^(light|sáng)$' {
      $script:Colors=@{
        Panel=[System.Drawing.SystemColors]::Control
        Input=[System.Drawing.SystemColors]::Window
        Fg   =[Drawing.Color]::Black
        Grid =[Drawing.Color]::FromArgb(210,210,210)
        Row  =[Drawing.Color]::White
        RowAlt=[Drawing.Color]::FromArgb(248,248,248)
        RowSel=[Drawing.Color]::FromArgb(204,232,255)
        RowSelInactive=[Drawing.Color]::FromArgb(232,232,232)
        SelText=[Drawing.Color]::Black
        HeaderBg=[Drawing.Color]::FromArgb(242,242,242)
        HeaderFg=[Drawing.Color]::Black
        TabActiveBg=[Drawing.Color]::FromArgb(248,248,248)
        TabInactiveBg=[Drawing.Color]::FromArgb(234,234,234)
        TabActiveFg=[Drawing.Color]::Black
        TabInactiveFg=[Drawing.Color]::FromArgb(90,90,90)
        TabBorder=[Drawing.Color]::FromArgb(200,200,200)
      }
      break
    }
    default {
      $script:Colors=@{
        Panel=[Drawing.Color]::FromArgb(230,232,235)
        Input=[Drawing.Color]::FromArgb(244,245,247)
        Fg   =[Drawing.Color]::FromArgb(36,36,36)
        Grid =[Drawing.Color]::FromArgb(205,207,210)
        Row  =[Drawing.Color]::FromArgb(238,240,242)
        RowAlt=[Drawing.Color]::FromArgb(233,235,238)
        RowSel=[Drawing.Color]::FromArgb(196,214,232)
        RowSelInactive=[Drawing.Color]::FromArgb(214,218,224)
        SelText=[Drawing.Color]::FromArgb(24,24,24)
        HeaderBg=[Drawing.Color]::FromArgb(222,224,228)
        HeaderFg=[Drawing.Color]::FromArgb(40,40,40)
        TabActiveBg=[Drawing.Color]::FromArgb(236,238,241)
        TabInactiveBg=[Drawing.Color]::FromArgb(222,224,228)
        TabActiveFg=[Drawing.Color]::FromArgb(32,32,32)
        TabInactiveFg=[Drawing.Color]::FromArgb(88,88,88)
        TabBorder=[Drawing.Color]::FromArgb(196,198,203)
      }
    }
  }
  $form.BackColor=$script:Colors.Panel; $form.ForeColor=$script:Colors.Fg
  Set-ThemeControl $form $script:Colors.Panel $script:Colors.Fg $script:Colors.Input
  $tabs.Invalidate()
  $form.Refresh()
}

$tabs.Add_DrawItem({
  param($sender,$e)
  $g=$e.Graphics
  $rect=$e.Bounds
  $isSelected=($tabs.SelectedIndex -eq $e.Index)
  $bg=if($isSelected){$script:Colors.TabActiveBg}else{$script:Colors.TabInactiveBg}
  $fg=if($isSelected){$script:Colors.TabActiveFg}else{$script:Colors.TabInactiveFg}
  $brush=[Drawing.SolidBrush]::new($bg)
  try{ $g.FillRectangle($brush,$rect) } finally { $brush.Dispose() }
  if($script:Colors.TabBorder){
    $pen=[Drawing.Pen]::new($script:Colors.TabBorder,1)
    try{
      $g.DrawRectangle($pen,$rect)
    } finally { $pen.Dispose() }
  }
  $image=$null
  $page=$tabs.TabPages[$e.Index]
  if($tabs.ImageList){
    if($page.ImageKey){ $image=$tabs.ImageList.Images[$page.ImageKey] }
    elseif($page.ImageIndex -ge 0){ $image=$tabs.ImageList.Images[$page.ImageIndex] }
  }
  $textRect=[Drawing.Rectangle]::new($rect.X+10,$rect.Y,$rect.Width-20,$rect.Height)
  if($image){
    $imgRect=[Drawing.Rectangle]::new($rect.X+10,$rect.Y+($rect.Height-$image.Height)/2,$image.Width,$image.Height)
    $g.DrawImage($image,$imgRect)
    $textRect.X=$imgRect.Right+6
    $textRect.Width=$rect.Right-$textRect.X-8
  }
  [System.Windows.Forms.TextRenderer]::DrawText($g,$page.Text,$page.Font,$textRect,$fg,
    [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::Left -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis)
})

# =========================================================
# TAB 1
# =========================================================
$tab1=New-Object System.Windows.Forms.TabPage
$tab1.Text="Tạo Tracklist"; $tab1.ImageKey='create'
$tabs.TabPages.Add($tab1)

$tracks = [System.Collections.Generic.List[psobject]]::new()

# ListView: [0]=#, [1]=☑(custom), [2]=Title, [3]=Duration, [4]=Cumulative
$list = New-Object System.Windows.Forms.ListView
$list.View='Details'; $list.FullRowSelect=$true; $list.MultiSelect=$true
$list.GridLines=$true; $list.AllowDrop=$true; $list.HideSelection=$false; $list.Dock='Fill'
# tăng chiều cao dòng ~ 28px + double buffer
$imgH=New-Object System.Windows.Forms.ImageList; $imgH.ImageSize=[Drawing.Size]::new(1,28); $list.SmallImageList=$imgH
$pi=$list.GetType().GetProperty('DoubleBuffered',[System.Reflection.BindingFlags]'NonPublic,Instance'); $pi.SetValue($list,$true,$null)

[void]$list.Columns.Add("#", 42)              # 0
[void]$list.Columns.Add("☑", 36)             # 1 (checkbox)
[void]$list.Columns.Add("File / Title", 520) # 2
[void]$list.Columns.Add("Duration", 120)     # 3
[void]$list.Columns.Add("Cumulative", 120)   # 4
$tab1.Controls.Add($list)

# OwnerDraw (checkbox ở cột 1)
$list.OwnerDraw=$true
$list.Add_DrawColumnHeader({
  $g=$_.Graphics; $r=$_.Bounds
  $brush=[Drawing.SolidBrush]::new($script:Colors.HeaderBg)
  try{ $g.FillRectangle($brush,$r) } finally { $brush.Dispose() }
  [System.Windows.Forms.TextRenderer]::DrawText($g,$_.Header.Text,$_.Font,$r,$script:Colors.HeaderFg,
    [System.Windows.Forms.TextFormatFlags]::Left -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter)
  $pen=[Drawing.Pen]::new($script:Colors.Grid,1)
  try{ $g.DrawLine($pen,$r.Left,$r.Bottom-1,$r.Right,$r.Bottom-1) } finally { $pen.Dispose() }
})
$list.Add_DrawItem({
  $g=$_.Graphics; $r=$_.Bounds
  $row=$_.ItemIndex; $isSel=$_.Item.Selected; $hasFocus=$list.Focused
  $bg = if($isSel){ if($hasFocus){$script:Colors.RowSel}else{$script:Colors.RowSelInactive} } else { if($row%2 -eq 0){$script:Colors.Row}else{$script:Colors.RowAlt} }
  $brush=[Drawing.SolidBrush]::new($bg)
  try{ $g.FillRectangle($brush,$r) } finally { $brush.Dispose() }
})
$list.Add_DrawSubItem({
  $lv=$list; if($lv.Columns.Count -lt 2){ return }
  $g=$_.Graphics; $r=$_.Bounds
  $row=$_.ItemIndex; $isSel=$_.Item.Selected; $hasFocus=$lv.Focused
  $bg = if($isSel){ if($hasFocus){$script:Colors.RowSel}else{$script:Colors.RowSelInactive} } else { if($row%2 -eq 0){$script:Colors.Row}else{$script:Colors.RowAlt} }
  $brush=[Drawing.SolidBrush]::new($bg)
  try{ $g.FillRectangle($brush,$r) } finally { $brush.Dispose() }
  $fg = if($isSel){$script:Colors.SelText}else{$script:Colors.Fg}

  if ($_.ColumnIndex -eq 1) {
    $box=14
    $cx=$r.Left + [Math]::Max(6,[Math]::Floor(($r.Width-$box)/2))
    $cy=$r.Top  + [Math]::Max(1,[Math]::Floor(($r.Height-$box)/2))
    $rect=[Drawing.Rectangle]::new($cx,$cy,$box,$box)
    $state = if ($tracks.Count -gt $row -and $tracks[$row].Star)
      { [System.Windows.Forms.ButtonState]::Checked } else { [System.Windows.Forms.ButtonState]::Normal }
    [System.Windows.Forms.ControlPaint]::DrawCheckBox($g,$rect,$state)
    $pen=[Drawing.Pen]::new($script:Colors.Grid,1)
    try{ $g.DrawLine($pen, $r.Right-1,$r.Top,$r.Right-1,$r.Bottom) } finally { $pen.Dispose() }
  } else {
    [System.Windows.Forms.TextRenderer]::DrawText($g,$_.SubItem.Text,$_.SubItem.Font,$r,$fg,
      [System.Windows.Forms.TextFormatFlags]::Left -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis)
  }
  $pen2=[Drawing.Pen]::new($script:Colors.Grid,1)
  try{ $g.DrawLine($pen2,$r.Left,$r.Bottom-1,$r.Right,$r.Bottom-1) } finally { $pen2.Dispose() }
})

# Playback helper
function Invoke-TrackPlayback([int]$rowIndex){
  if($rowIndex -lt 0 -or $rowIndex -ge $tracks.Count){ return }
  $p=$tracks[$rowIndex].Path
  if(Test-Path $p){ Start-Process -FilePath $p }
}

# Toggle checkbox click (cột 1)
$list.Add_MouseDown({
  $hit=$list.HitTest($_.X,$_.Y)
  if(-not $hit.Item){ return }
  $subIndex = if($hit.SubItem){ $hit.Item.SubItems.IndexOf($hit.SubItem) } else { -1 }
  if($subIndex -eq 1){
    $row=$hit.Item.Index
    if ($row -ge 0 -and $row -lt $tracks.Count) {
      $tracks[$row].Star = -not [bool]$tracks[$row].Star
      $list.Invalidate($hit.SubItem.Bounds)
    }
  }
})

# Double-click: chỉ phát khi click trong cột "File / Title" (col index 2)
$list.Add_MouseDoubleClick({
  $pt=[Drawing.Point]::new($_.X,$_.Y)
  $hit=$list.HitTest($pt)
  if(-not $hit.Item){ return }
  $subIndex = if($hit.SubItem){ $hit.Item.SubItems.IndexOf($hit.SubItem) } else { -1 }
  if($subIndex -eq 2){
    Invoke-TrackPlayback $hit.Item.Index
  }
})

# -------- Bottom panel
$baseBottomH = 240
$bottom = New-Object System.Windows.Forms.Panel
$bottom.Dock='Bottom'; $bottom.Height=$baseBottomH
$bottom.Padding='14,8,14,12'
$tab1.Controls.Add($bottom)
$defaultMinSongs=25
$defaultMaxSongs=30
$defaultOutDir=[Environment]::GetFolderPath('Desktop')
$defaultBaseName='merged'

# Hàng 1: Random + Priority + Preview (1 hàng)
$chkRandom = New-Object System.Windows.Forms.CheckBox; $chkRandom.Text="Tạo list Random"; $chkRandom.AutoSize=$true
$numLists  = New-Object System.Windows.Forms.NumericUpDown; $numLists.Minimum=1; $numLists.Maximum=999; $numLists.Value=1; $numLists.Width=55
$lblList   = New-Object System.Windows.Forms.Label; $lblList.Text="List"; $lblList.AutoSize=$true
$numMin    = New-Object System.Windows.Forms.NumericUpDown; $numMin.Minimum=1; $numMin.Maximum=999; $numMin.Value=$defaultMinSongs; $numMin.Width=55
$lblDen    = New-Object System.Windows.Forms.Label; $lblDen.Text="đến"; $lblDen.AutoSize=$true
$numMax    = New-Object System.Windows.Forms.NumericUpDown; $numMax.Minimum=1; $numMax.Maximum=999; $numMax.Value=$defaultMaxSongs; $numMax.Width=55
$lblSongs  = New-Object System.Windows.Forms.Label; $lblSongs.Text="bài hát"; $lblSongs.AutoSize=$true
$chkPriority = New-Object System.Windows.Forms.CheckBox; $chkPriority.Text="Ưu tiên sắp xếp những bài đánh dấu"; $chkPriority.AutoSize=$true
$btnPreview  = New-Object System.Windows.Forms.Button; $btnPreview.Text="Random (Preview)"; $btnPreview.Width=180

# Hàng 2: Import/Remove/Up/Down/Clear
$btnImport=New-Object System.Windows.Forms.Button; $btnImport.Text="Import"
$btnRemove=New-Object System.Windows.Forms.Button; $btnRemove.Text="Remove"
$btnUp=New-Object System.Windows.Forms.Button; $btnUp.Text="Up"
$btnDown=New-Object System.Windows.Forms.Button; $btnDown.Text="Down"
$btnClear=New-Object System.Windows.Forms.Button; $btnClear.Text="Clear"

# Hàng 3: Output + Browse
$lblOut=New-Object System.Windows.Forms.Label; $lblOut.Text="Output:"
$txtOutDir=New-Object System.Windows.Forms.TextBox; $txtOutDir.Text=$defaultOutDir
$btnBrowse=New-Object System.Windows.Forms.Button; $btnBrowse.Text="Browse"

# Hàng 4: Filename/Total + Format/Generate
$lblBase=New-Object System.Windows.Forms.Label; $lblBase.Text="Filename:"
$txtBase=New-Object System.Windows.Forms.TextBox; $txtBase.Text=$defaultBaseName
$lblTotal=New-Object System.Windows.Forms.Label; $lblTotal.Text='Total: 00:00:00'
$lblFmt=New-Object System.Windows.Forms.Label; $lblFmt.Text='Format:'
$cmbFmt=New-Object System.Windows.Forms.ComboBox; $cmbFmt.DropDownStyle='DropDownList'
[void]$cmbFmt.Items.Add('WAV (48k/16-bit)'); [void]$cmbFmt.Items.Add('MP3 (192 kbps)'); $cmbFmt.SelectedIndex=0
$btnGen=New-Object System.Windows.Forms.Button; $btnGen.Text="Generate"

$bottom.Controls.AddRange(@(
  $chkRandom,$numLists,$lblList,$numMin,$lblDen,$numMax,$lblSongs,$chkPriority,$btnPreview,
  $btnImport,$btnRemove,$btnUp,$btnDown,$btnClear,
  $lblOut,$txtOutDir,$btnBrowse,
  $lblBase,$txtBase,$lblTotal,$lblFmt,$cmbFmt,$btnGen
))

function Reflow-Bottom {
  $w=$bottom.ClientSize.Width; $r=14; $gap=10; $btnW=100
  # hàng 1
  $y=8; $x=10
  $chkRandom.SetBounds($x,$y+6,130,24); $x+=130+$gap
  $numLists.SetBounds($x,$y,55,28);     $x+=55+$gap
  $lblList.SetBounds($x,$y+6,30,24);    $x+=30+$gap
  $numMin.SetBounds($x,$y,55,28);       $x+=55+$gap
  $lblDen.SetBounds($x,$y+6,28,24);     $x+=28+$gap
  $numMax.SetBounds($x,$y,55,28);       $x+=55+$gap
  $lblSongs.SetBounds($x,$y+6,50,24);   $x+=50+($gap*2)
  $chkPriority.SetBounds($x,$y+6,280,24); $x+=280+($gap*2)
  $btnPreview.SetBounds($x,$y-2,180,32)

  # hàng 2
  $y=46
  $btnImport.SetBounds(10,$y,90,32)
  $btnRemove.SetBounds(110,$y,90,32)
  $btnUp.SetBounds(210,$y,70,32)
  $btnDown.SetBounds(290,$y,70,32)
  $btnClear.SetBounds(370,$y,70,32)

  # hàng 3
  $y=88
  $lblOut.SetBounds(10,$y+4,60,24)
  $btnBrowse.SetBounds($w-$r-$btnW,$y-2,$btnW,30)
  $txtOutDir.SetBounds(75,$y,$w-$r-$btnW-75-10,26)

  # hàng 4
  $y=126
  $btnGen.SetBounds($w-$r-$btnW,$y-2,$btnW,34)
  $cmbFmt.SetBounds($btnGen.Left-160-10,$y,160,26)
  $lblFmt.SetBounds($cmbFmt.Left-60,$y+4,55,22)
  $lblBase.SetBounds(10,$y+4,70,22)
  $txtBase.SetBounds(80,$y,220,26)
  $lblTotal.SetBounds(310,$y+4,220,22)
}
$tab1.Add_Resize({ Reflow-Bottom })
Reflow-Bottom

# Status
$lblStatus1=New-Object System.Windows.Forms.Label
$lblStatus1.Text="Ready."; $lblStatus1.AutoEllipsis=$true
$lblStatus1.Dock='Bottom'; $lblStatus1.Padding=New-Object System.Windows.Forms.Padding(10,0,10,6)
$tab1.Controls.Add($lblStatus1)

# Columns + data helpers
function Resize-Columns{
  if($list.Columns.Count -lt 5){return}
  $cNum=42; $cChk=36
  $w=$list.ClientSize.Width
  $remain=[Math]::Max(220,$w-$cNum-$cChk-20)
  $list.Columns[0].Width=$cNum
  $list.Columns[1].Width=$cChk
  $list.Columns[2].Width=[int]($remain*0.60)
  $list.Columns[3].Width=[int]($remain*0.20)
  $list.Columns[4].Width=$remain-$list.Columns[2].Width-$list.Columns[3].Width
}
$list.Add_HandleCreated({ Resize-Columns })
$tab1.Add_Resize({ Resize-Columns })

function Update-TotalLabel{ $sum=0.0; foreach($t in $tracks){$sum+=$t.DurSec}; $lblTotal.Text="Total: " + (SecToHMS $sum) }
function Refresh-ListView{
  $list.BeginUpdate()
  try{
    $list.Items.Clear(); $run=0.0
    for($i=0;$i -lt $tracks.Count;$i++){
      $t=$tracks[$i]; $run+=[double]$t.DurSec
      $it=New-Object System.Windows.Forms.ListViewItem(($i+1).ToString()) # #
      [void]$it.SubItems.Add("")                           # ☑
      [void]$it.SubItems.Add($t.Title)                     # Title
      [void]$it.SubItems.Add((SecToHMS $t.DurSec))         # Duration
      [void]$it.SubItems.Add((SecToHMS $run))              # Cumulative
      [void]$list.Items.Add($it)
    }
  } finally { $list.EndUpdate() }
  Update-TotalLabel
}

function Add-Files([string[]]$paths){
  $okExt=@(".mp3",".wav")
  foreach($p in $paths){
    $ext=[IO.Path]::GetExtension($p).ToLower(); if(-not $okExt.Contains($ext)){continue}
    if($tracks | Where-Object { $_.Path -eq $p }){continue}
    $title=[IO.Path]::GetFileNameWithoutExtension($p); $dur=Get-DurationSec $p
    $tracks.Add((New-TrackObject -Path $p -Title $title -DurSec $dur)) | Out-Null
  }
  Refresh-ListView
}
$script:lastIndex=-1
$list.Add_ItemSelectionChanged({ if ($_.IsSelected) { $script:lastIndex = $_.ItemIndex } })
function Move-Selected([int]$delta){
  $idx = if ($list.SelectedIndices.Count -gt 0){ $list.SelectedIndices[0] } else { $script:lastIndex }
  if($idx -lt 0){return}; $new=$idx+$delta; if($new -lt 0 -or $new -ge $tracks.Count){return}
  $target=$new
  $item=$tracks[$idx]
  if($target -gt $idx){ $target-- }
  $tracks.RemoveAt($idx)
  $tracks.Insert($target,$item)
  Refresh-ListView
  if($target -ge 0 -and $target -lt $list.Items.Count){
    $list.Items[$target].Selected=$true; $list.EnsureVisible($target); $list.Select(); $script:lastIndex=$target
  }
}
function Remove-Selected{
  if($list.SelectedIndices.Count -eq 0){return}
  $idxs=@(); foreach($i in $list.SelectedIndices){ $idxs += [int]$i }
  $idxs=$idxs|Sort-Object -Descending
  foreach($i in $idxs){ if($i -ge 0 -and $i -lt $tracks.Count){ $tracks.RemoveAt($i) } }
  Refresh-ListView; $script:lastIndex=-1
}

# Keyboard + DnD
$list.Add_KeyDown({
  if ($_.Control -and $_.KeyCode -eq 'O')    { $btnImport.PerformClick(); $_.SuppressKeyPress=$true; return }
  if ($_.Control -and $_.KeyCode -eq 'Up')   { Move-Selected -1; $_.SuppressKeyPress=$true; return }
  if ($_.Control -and $_.KeyCode -eq 'Down') { Move-Selected 1;  $_.SuppressKeyPress=$true; return }
  if ($_.KeyCode -eq 'Delete')               { Remove-Selected;  $_.SuppressKeyPress=$true; return }
  if ($_.KeyCode -eq 'Return')               { $btnGen.PerformClick(); $_.SuppressKeyPress=$true; return }
})
$list.InsertionMark.Color=[Drawing.Color]::FromArgb(0,120,215)
$script:dragIndex=-1
$script:dragFormat='LV-TRACKLIST-REORDER'
$list.Add_ItemDrag({
  if($list.SelectedIndices.Count -gt 0){
    $script:dragIndex=$list.SelectedIndices[0]
    $data=New-Object System.Windows.Forms.DataObject
    $data.SetData($script:dragFormat,$script:dragIndex)
    [void]$list.DoDragDrop($data,[System.Windows.Forms.DragDropEffects]::Move)
  }
})
$list.Add_DragEnter({
  if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect=[System.Windows.Forms.DragDropEffects]::Copy }
  elseif ($_.Data.GetDataPresent($script:dragFormat)) { $_.Effect=[System.Windows.Forms.DragDropEffects]::Move }
  else { $_.Effect=[System.Windows.Forms.DragDropEffects]::None }
})
$list.Add_DragOver({
  $pt=$list.PointToClient([Drawing.Point]::new($_.X,$_.Y))
  if(-not $list.ClientRectangle.Contains($pt)){
    $list.InsertionMark.Index=-1
    $_.Effect=[System.Windows.Forms.DragDropEffects]::None
    return
  }
  if ($script:dragIndex -ge 0) {
    $item=$list.GetItemAt($pt.X,$pt.Y)
    if ($item) {
      $list.InsertionMark.Index=$item.Index
      $list.InsertionMark.AppearsAfterItem=($pt.Y -ge ($item.Bounds.Top + $item.Bounds.Height/2))
    } else {
      $list.InsertionMark.Index=$list.Items.Count-1
      $list.InsertionMark.AppearsAfterItem=$true
    }
  }
})
$list.Add_DragLeave({ $list.InsertionMark.Index=-1 })
$list.Add_DragDrop({
  $pt=$list.PointToClient([Drawing.Point]::new($_.X,$_.Y))
  if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
    Add-Files ($_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))
    $script:dragIndex=-1; $list.InsertionMark.Index=-1
    return
  }
  if ($script:dragIndex -lt 0 -or -not $_.Data.GetDataPresent($script:dragFormat)){ $list.InsertionMark.Index=-1; return }
  if(-not $list.ClientRectangle.Contains($pt)){
    $script:dragIndex=-1; $list.InsertionMark.Index=-1; return
  }
  $item=$list.GetItemAt($pt.X,$pt.Y)
  if($item){
    $idx=$item.Index
    if($pt.Y -ge ($item.Bounds.Top + $item.Bounds.Height/2)){ $idx++ }
  } else {
    $idx=$tracks.Count
  }
  if($idx -gt $tracks.Count){ $idx=$tracks.Count }
  $moving=$tracks[$script:dragIndex]
  $tracks.RemoveAt($script:dragIndex)
  if($idx -gt $tracks.Count){ $idx=$tracks.Count }
  $tracks.Insert($idx,$moving)
  Refresh-ListView
  if($idx -ge 0 -and $idx -lt $list.Items.Count){
    $list.Items[$idx].Selected=$true
    $list.EnsureVisible($idx)
    $list.Select()
  }
  $script:dragIndex=-1
  $list.InsertionMark.Index=-1
})

# Buttons
$btnImport.Add_Click({ $ofd=New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter="Audio (MP3, WAV)|*.mp3;*.wav|All files|*.*"; $ofd.Multiselect=$true; if($ofd.ShowDialog() -eq 'OK'){ Add-Files $ofd.FileNames } })
$btnRemove.Add_Click({ Remove-Selected })
$btnUp.Add_Click({ Move-Selected -1 })
$btnDown.Add_Click({ Move-Selected 1 })
$btnClear.Add_Click({ $tracks.Clear(); Refresh-ListView })
$btnBrowse.Add_Click({ $fbd=New-Object System.Windows.Forms.FolderBrowserDialog; if($fbd.ShowDialog() -eq 'OK'){ $txtOutDir.Text=$fbd.SelectedPath } })

function Get-ShuffledTracks([System.Collections.IEnumerable]$items,[bool]$prioritize){
  $rng=[System.Random]::new()
  function Shuffle([psobject[]]$array,[System.Random]$r){
    for($i=$array.Length-1;$i -gt 0;$i--){
      $j=$r.Next($i+1)
      if($j -ne $i){ $tmp=$array[$i]; $array[$i]=$array[$j]; $array[$j]=$tmp }
    }
    return $array
  }
  $source=@($items)
  if($prioritize){
    $stars=Shuffle (@($source | Where-Object { $_.Star })) $rng
    $others=Shuffle (@($source | Where-Object { -not $_.Star })) $rng
    return $stars + $others
  }
  return Shuffle (@($source)) $rng
}

# Random(Preview): chỉ xáo thứ tự danh sách hiện có (không dùng Min/Max)
$btnPreview.Add_Click({
  if ($tracks.Count -le 1) { return }
  $ordered=Get-ShuffledTracks $tracks $chkPriority.Checked
  $tracks.Clear(); foreach($t in $ordered){ $tracks.Add($t) } # giữ nguyên object -> giữ Star
  Refresh-ListView
})

function Merge-Once([System.Collections.IEnumerable]$sel,[string]$base,[bool]$isWav){
  $selList=@($sel)
  if($selList.Count -eq 0){ throw "Không có bài hát để gộp." }
  $outDir=$txtOutDir.Text.Trim()
  $temp=Join-Path ([IO.Path]::GetTempPath()) ("merge_"+[guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Force -Path $temp|Out-Null
  $listFile=Join-Path $temp "list.txt"
  $stamp   =Join-Path $outDir "$base`_timestamps.txt"
  $outA    =Join-Path $outDir ($base+$(if($isWav){".wav"}else{".mp3"}))
  try{
    $concat=@()
    for($i=0;$i -lt $selList.Count;$i++){
      $track=$selList[$i]
      $src=$track.Path
      $seg=Join-Path $temp ("seg_{0:D5}.wav" -f $i)
      $lblStatus1.Text="Normalizing: $([IO.Path]::GetFileName($src))"; $form.Refresh()
      $psi1=New-Object System.Diagnostics.ProcessStartInfo
      $psi1.FileName=$ffmpeg
      $psi1.Arguments="-y -i `"$src`" -vn -ac 2 -ar 48000 -sample_fmt s16 -af aresample=async=1 `"$seg`""
      $psi1.UseShellExecute=$false; $psi1.RedirectStandardError=$true; $psi1.CreateNoWindow=$true
      $p1=[System.Diagnostics.Process]::Start($psi1); $err=$p1.StandardError.ReadToEnd(); $p1.WaitForExit()
      if($p1.ExitCode -ne 0 -or -not (Test-Path $seg)){
        throw "Normalize failed for `"$src`" (exit $($p1.ExitCode)).`n$err"
      }
      $concat+="file '$($seg.Replace("'", "''"))'"
    }
    [IO.File]::WriteAllLines($listFile,$concat,[Text.UTF8Encoding]::new($false))
    $start=0.0; $st=@(); foreach($t in $selList){ $st+=("[{0}] - {1}" -f (SecToHMS $start),$t.Title); $start+=$t.DurSec }
    [IO.File]::WriteAllLines($stamp,$st,[Text.UTF8Encoding]::new($false))
    $lblStatus1.Text="Merging..."; $form.Refresh()
    $psi2=New-Object System.Diagnostics.ProcessStartInfo
    $psi2.FileName=$ffmpeg
    $psi2.Arguments= if($isWav){ "-y -f concat -safe 0 -i `"$listFile`" -c:a pcm_s16le -ar 48000 -ac 2 `"$outA`"" }
                     else       { "-y -f concat -safe 0 -i `"$listFile`" -c:a libmp3lame -ar 48000 -b:a 192k `"$outA`"" }
    $psi2.UseShellExecute=$false; $psi2.RedirectStandardError=$true; $psi2.CreateNoWindow=$true
    $p2=[System.Diagnostics.Process]::Start($psi2); $err2=$p2.StandardError.ReadToEnd(); $p2.WaitForExit()
    if($p2.ExitCode -ne 0 -or -not (Test-Path $outA)){
      throw "ffmpeg concat failed (exit $($p2.ExitCode)).`n$err2"
    }
    return $outA
  } finally { if(Test-Path $temp){ Remove-Item -Recurse -Force $temp } }
}

$btnGen.Add_Click({
  if ($tracks.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No tracks to merge.","",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null; return
  }
  $outDir=$txtOutDir.Text.Trim()
  if (-not (Test-Path $outDir)) {
    [System.Windows.Forms.MessageBox]::Show("Output folder not found.","",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null; return
  }
  $base = if ([string]::IsNullOrWhiteSpace($txtBase.Text)) { "merged" } else { $txtBase.Text }
  $isWav = ($cmbFmt.SelectedIndex -eq 0)

  try{
    if (-not $chkRandom.Checked) {
      $null = Merge-Once -sel:$tracks -base:$base -isWav:$isWav
      [System.Windows.Forms.MessageBox]::Show("Success!`nAudio/Timestamps đã xuất vào:`n$outDir","Completed",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
      return
    }

    $min=[int]$numMin.Value; $max=[int]$numMax.Value
    if ($max -lt $min) { $tmp=$min; $min=$max; $max=$tmp }
    $lists=[int]$numLists.Value
    for ($i=1; $i -le $lists; $i++) {
      $k = $min + (Get-Random -Minimum 0 -Maximum ($max-$min+1))
      $k = [Math]::Min($k, $tracks.Count)
      if($k -le 0){ continue }
      $ordered=Get-ShuffledTracks $tracks $chkPriority.Checked
      $sel=$ordered[0..($k-1)]
      $name = "{0}_{1:D2}" -f $base, $i
      $null = Merge-Once -sel:$sel -base:$name -isWav:$isWav
    }
    [System.Windows.Forms.MessageBox]::Show("Success!`nĐã xuất $lists list vào:`n$outDir","Completed",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
  } catch {
    [System.Windows.Forms.MessageBox]::Show("Lỗi khi gộp: $($_.Exception.Message)","Error",
      [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
  }
})

# =========================================================
# TAB 2: Điều chỉnh Tracklist
# =========================================================
$tab2=New-Object System.Windows.Forms.TabPage
$tab2.Text="Điều chỉnh Tracklist"; $tab2.ImageKey='adjust'
$tabs.TabPages.Add($tab2)

$tlp=New-Object System.Windows.Forms.TableLayoutPanel
$tlp.Dock='Fill'; $tlp.ColumnCount=3; $tlp.RowCount=6; $tlp.Padding='10,10,10,10'
$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,150)))
$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,90)))
$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,58)))
$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,34)))
$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,45)))
$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,34)))
$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,52)))
$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,55)))
$tab2.Controls.Add($tlp)

$txtGuide=New-Object System.Windows.Forms.TextBox
$txtGuide.Multiline=$true; $txtGuide.ReadOnly=$true; $txtGuide.Dock='Fill'
$txtGuide.Text="Nhập Tracklist gốc + các khoảng thời gian bị cắt để tạo Tracklist mới."
$tlp.Controls.Add($txtGuide,0,0); $tlp.SetColumnSpan($txtGuide,3)

$lblTxt2=New-Object System.Windows.Forms.Label; $lblTxt2.Text="Tracklist .txt gốc:"; $lblTxt2.Dock='Fill'
$txtTxt2=New-Object System.Windows.Forms.TextBox; $txtTxt2.Dock='Fill'
$btnTxt2=New-Object System.Windows.Forms.Button; $btnTxt2.Text="Browse"; $btnTxt2.Dock='Fill'
$btnTxt2.Add_Click({ $dlg=New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter="Text (*.txt)|*.txt|All files|*.*"; if($dlg.ShowDialog() -eq 'OK'){ $txtTxt2.Text=$dlg.FileName } })
$tlp.Controls.Add($lblTxt2,0,1); $tlp.Controls.Add($txtTxt2,1,1); $tlp.Controls.Add($btnTxt2,2,1)

$lblCuts=New-Object System.Windows.Forms.Label; $lblCuts.Text="Các đoạn bị cắt`r`n(mỗi dòng):"; $lblCuts.Dock='Fill'
$txtCuts=New-Object System.Windows.Forms.TextBox; $txtCuts.Multiline=$true; $txtCuts.ScrollBars='Vertical'; $txtCuts.Dock='Fill'
$defaultCutsText="00:04:02 - 00:06:07`r`n00:27:32 - 00:30:05"
$txtCuts.Text=$defaultCutsText
$tlp.Controls.Add($lblCuts,0,2); $tlp.Controls.Add($txtCuts,1,2)

$lblOut2=New-Object System.Windows.Forms.Label; $lblOut2.Text="Thư mục xuất:"; $lblOut2.Dock='Fill'
$txtOut2=New-Object System.Windows.Forms.TextBox; $txtOut2.Dock='Fill'
$btnOut2=New-Object System.Windows.Forms.Button; $btnOut2.Text="Browse"; $btnOut2.Dock='Fill'
$btnOut2.Add_Click({ $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog() -eq 'OK'){ $txtOut2.Text=$dlg.SelectedPath } })
$tlp.Controls.Add($lblOut2,0,3); $tlp.Controls.Add($txtOut2,1,3); $tlp.Controls.Add($btnOut2,2,3)

$bar=New-Object System.Windows.Forms.TableLayoutPanel
$bar.Dock='Fill'; $bar.ColumnCount=2; $bar.RowCount=1; $bar.Padding='6,6,6,6'
$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,220)))
$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
$chkSRT=New-Object System.Windows.Forms.CheckBox; $chkSRT.Text="Xuất kèm SRT"; $chkSRT.AutoSize=$true; $chkSRT.Margin='4,10,0,0'
$bar.Controls.Add($chkSRT,0,0)
$btnPanel=New-Object System.Windows.Forms.FlowLayoutPanel; $btnPanel.FlowDirection='RightToLeft'; $btnPanel.WrapContents=$false; $btnPanel.Dock='None'; $btnPanel.Anchor='Top,Right'
$btnRun2=New-Object System.Windows.Forms.Button; $btnRun2.Text="Generate"
$btnClose2=New-Object System.Windows.Forms.Button; $btnClose2.Text="Close"; $btnClose2.Add_Click({ $form.Close() })
$btnPanel.Controls.Add($btnClose2); $btnPanel.Controls.Add($btnRun2)
$bar.Controls.Add($btnPanel,1,0)
$tlp.Controls.Add($bar,0,4); $tlp.SetColumnSpan($bar,3)

$log2=New-Object System.Windows.Forms.TextBox; $log2.Multiline=$true; $log2.ReadOnly=$true; $log2.ScrollBars='Vertical'; $log2.Dock='Fill'
$tlp.Controls.Add($log2,0,5); $tlp.SetColumnSpan($log2,3)

$btnRun2.Add_Click({
  $log2.Clear()
  if([string]::IsNullOrWhiteSpace($txtTxt2.Text)){ $log2.Text="Thiếu tracklist."; return }
  if(-not (Test-Path -LiteralPath $txtTxt2.Text -PathType Leaf)){ $log2.Text="Tracklist không tồn tại."; return }
  if([string]::IsNullOrWhiteSpace($txtOut2.Text)){ $log2.Text="Thiếu thư mục xuất."; return }
  if(-not (Test-Path -LiteralPath $txtOut2.Text -PathType Container)){ $log2.Text="Thư mục xuất không tồn tại."; return }
  try{
    $base=[IO.Path]::GetFileNameWithoutExtension($txtTxt2.Text)
    $outAllTxt=Join-Path $txtOut2.Text ($base+"_adjusted_all.txt")
    $outAllSrt=Join-Path $txtOut2.Text ($base+"_adjusted_all.srt")
    $cutList=New-Object 'System.Collections.Generic.List[object]'
    foreach($line in ($txtCuts.Text -split "`r?`n")){
      $t=$line.Trim(); if(!$t){continue}
      $m=[regex]::Match($t,'^(?<s>\d{2}:\d{2}:\d{2})\s*[-–]\s*(?<e>\d{2}:\d{2}:\d{2})$')
      if($m.Success){ $s=Parse-Time $m.Groups['s'].Value; $e=Parse-Time $m.Groups['e'].Value; if($e -gt $s){ $cutList.Add([pscustomobject]@{Start=$s;End=$e}) } }
    }
    $cutList=Merge-Cuts $cutList
    $items=Read-Tracklist $txtTxt2.Text; if($items.Count -eq 0){ throw "Tracklist không đọc được/sai định dạng." }
    $all=New-Object 'System.Collections.Generic.List[object]'
    for($i=0;$i -lt $items.Count;$i++){
      $T=[int]$items[$i].Time; $newT=$T - (Total-CutBefore $T $cutList)
      $all.Add([pscustomobject]@{Time=[int]$newT;Title=$items[$i].Title})
    }
    $minGap=10; $keep=New-Object 'System.Collections.Generic.List[object]'
    for($i=0;$i -lt $all.Count;$i++){
      $drop=$false; if($i -lt $all.Count-1){ $gap=[int]$all[$i+1].Time - [int]$all[$i].Time; if($gap -lt $minGap){$drop=$true} }
      if(-not $drop){ $keep.Add($all[$i]) }
    }
    $lines=@(); foreach($it in $keep){ $lines+=('[{0}] - {1}' -f (Fmt-HMS $it.Time),$it.Title) }
    [IO.File]::WriteAllLines($outAllTxt,$lines,[Text.UTF8Encoding]::new($false))
    if($chkSRT.Checked){ Build-SRT $keep $outAllSrt }
    $log2.Lines=@("✅ Đã xuất: $outAllTxt")+($(if($chkSRT.Checked){"✅ Đã xuất: $outAllSrt"}))+"Hoàn tất 🎉"
  }catch{ $log2.Text="❌ Lỗi: $($_.Exception.Message)" }
})

function Reset-Toolkit{
  $tracks.Clear()
  if($list.SelectedIndices.Count -gt 0){ $list.SelectedIndices.Clear() }
  Refresh-ListView
  $chkRandom.Checked=$false
  $chkPriority.Checked=$false
  $numLists.Value=1
  $numMin.Value=$defaultMinSongs
  $numMax.Value=$defaultMaxSongs
  $txtOutDir.Text=$defaultOutDir
  $txtBase.Text=$defaultBaseName
  $cmbFmt.SelectedIndex=0
  $lblStatus1.Text="Ready."
  $script:dragIndex=-1
  $script:lastIndex=-1
  $list.InsertionMark.Index=-1
  $txtTxt2.Clear()
  $txtOut2.Clear()
  $txtCuts.Text=$defaultCutsText
  $chkSRT.Checked=$false
  $log2.Clear()
  if($cmbThemeTop.SelectedItem -ne 'Calm Gray'){ $cmbThemeTop.SelectedItem='Calm Gray' } else { Apply-Theme 'Calm Gray' }
  Resize-Columns
  Reflow-Bottom
}
$btnResetTop.Add_Click({ Reset-Toolkit })

$cmbThemeTop.Add_SelectedIndexChanged({ Apply-Theme ($cmbThemeTop.SelectedItem) })

# ---------------- Run ----------------
Apply-Theme 'Calm Gray'
Resize-Columns
Reflow-Bottom
Refresh-ListView
[void]$form.ShowDialog()
