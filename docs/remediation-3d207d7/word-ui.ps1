param([int]$WordProcessId=0,[string]$Action='probe',[string]$Name='',[string]$Target='word',[string]$ExpectedDocument='',[string]$Doc='',[int]$Index=1,[int]$Position=-1,[string]$Value='',[string]$Choice='',[string]$InputFile='',[string]$LatinFont='',[single]$FontSize=0,[string]$Label='ui',[long]$WindowHandle=0,[switch]$Checked,[switch]$KeyClick)
$ErrorActionPreference='Stop'
if(!$WordProcessId){
 $recordedTarget=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw | ConvertFrom-Json
 if(!$recordedTarget.pid -or $recordedTarget.pid -eq 27664){throw 'No owned stage Word process is recorded.'}
 $WordProcessId=[int]$recordedTarget.pid
}
. (Join-Path $PSScriptRoot 'word-connection.ps1')
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$OutputEncoding=[Console]::OutputEncoding
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Windows.Forms,System.Drawing
Add-Type -TypeDefinition @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class AuditUi {
public struct Rect { public int Left,Top,Right,Bottom; }
public struct Point { public int X,Y; }
[DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h,ref Point p);
[DllImport("user32.dll",EntryPoint="SendMessageW")] public static extern IntPtr ListItemRect(IntPtr h,uint m,IntPtr index,ref Rect r);
public delegate bool EnumProc(IntPtr h,IntPtr l);
[DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb,IntPtr l);
[DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent,EnumProc cb,IntPtr l);
[DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
[DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr h);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h,uint flag);
[DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr e);
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h,uint m,UIntPtr w,IntPtr l);
[DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
}
'@
[void][AuditUi]::SetProcessDPIAware()
$out=Join-Path $PSScriptRoot 'evidence';[void][IO.Directory]::CreateDirectory($out)
$script:windows=@()
$script:officeProcessNames=@{}
foreach($process in (Get-Process WINWORD,visualtex -ErrorAction SilentlyContinue)){$script:officeProcessNames[[uint32]$process.Id]=$process.ProcessName}
[void][AuditUi]::EnumWindows({param($h,$l)
if([AuditUi]::IsWindowVisible($h)) { [uint32]$pidValue=0;[void][AuditUi]::GetWindowThreadProcessId($h,[ref]$pidValue);if($script:officeProcessNames.ContainsKey($pidValue)){$s=[Text.StringBuilder]::new(1024);[void][AuditUi]::GetWindowText($h,$s,1024);$script:windows+=@{handle=$h.ToInt64();title=$s.ToString();pid=$pidValue;process=$script:officeProcessNames[$pidValue]}} };return $true
},[IntPtr]::Zero)
function FindEl($root,[string]$n) { return $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty,$n)) }
function ClickEl($el) {
if($null -eq $el){throw "UI element not found: $Name"}
if(-not $KeyClick) { $expand=$null;if($el.Current.ControlType -eq [System.Windows.Automation.ControlType]::MenuItem -and $el.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern,[ref]$expand)) {([System.Windows.Automation.ExpandCollapsePattern]$expand).Expand();return}; $invoke=$null; if($el.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern,[ref]$invoke)) { ([System.Windows.Automation.InvokePattern]$invoke).Invoke(); return }; $select=$null; if($el.Current.ControlType -eq [System.Windows.Automation.ControlType]::TabItem -and $el.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$select)) {([System.Windows.Automation.SelectionItemPattern]$select).Select();return} }
$r=$el.Current.BoundingRectangle
if($r.IsEmpty -or $r.Width -le 0 -or $r.Height -le 0){throw 'UI element has no visible rectangle'}
[void][AuditUi]::SetCursorPos([int]($r.X+$r.Width/2),[int]($r.Y+$r.Height/2));[AuditUi]::mouse_event(2,0,0,0,[UIntPtr]::Zero);Start-Sleep -Milliseconds 60;[AuditUi]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
}
function PasteText([string]$txt) { $old=[System.Windows.Forms.Clipboard]::GetDataObject();try { [System.Windows.Forms.Clipboard]::SetText($txt);[System.Windows.Forms.SendKeys]::SendWait('^v');Start-Sleep -Milliseconds 500 } finally {if($null -ne $old){try{[System.Windows.Forms.Clipboard]::SetDataObject($old,$true)}catch{}}} }
function Tree($r) {
$result=@();$els=$r.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
foreach($e in $els){$c=$e.Current;$v='';$pat=$null;try {if($e.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$pat)){$v=([System.Windows.Automation.ValuePattern]$pat).Current.Value};$pat=$null;if($e.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$pat)){$v=([System.Windows.Automation.TogglePattern]$pat).Current.ToggleState.ToString()}}catch{}
$result+=@{type=$c.ControlType.ProgrammaticName;name=$c.Name;id=$c.AutomationId;enabled=$c.IsEnabled;offscreen=$c.IsOffscreen;value=$v;rect=$c.BoundingRectangle.ToString()}}
return $result
}
if($Action -eq 'windows') { @($script:windows|Where-Object {$_.title})|ConvertTo-Json -Depth 4;exit }
$stageTargetPath=Join-Path $out 'target-word.json'
$stageTarget=if(Test-Path -LiteralPath $stageTargetPath){Get-Content -LiteralPath $stageTargetPath -Raw | ConvertFrom-Json}else{$null}
$expectedStageDocument=if($ExpectedDocument){$ExpectedDocument}elseif($stageTarget -and $stageTarget.pid -eq $WordProcessId){$stageTarget.document}else{''}
if($WindowHandle -eq -1) {
 $matches=@($script:windows|Where-Object {$_.process -eq 'WINWORD' -and $_.title -like '* - Word' -and ($WordProcessId -eq 0 -or $_.pid -eq $WordProcessId) -and (!$expectedStageDocument -or $_.title -eq ($expectedStageDocument+' - Word'))})
 if(!$matches.Count){throw 'The explicitly recorded test document window is unavailable.'}
 $WindowHandle=[long]$matches[0].handle
}
$word=$null
if($WindowHandle -eq 0 -or $Action -in @('position','select-math','select-ole','text','enter','new')) {$word=(Connect-RunningWord $WordProcessId)}
if($Action -eq 'new') {
 if($WordProcessId -eq 27664 -or !$stageTarget -or $stageTarget.pid -ne $WordProcessId){throw 'New stage documents require an explicitly owned Word process.'}
 $d=$word.Documents.Add();$d.Activate();[void][AuditUi]::SetForegroundWindow([IntPtr]$d.ActiveWindow.Hwnd)
 $stageTarget | Add-Member -NotePropertyName document -NotePropertyValue $d.Name -Force
 $stageTarget | ConvertTo-Json | Set-Content -LiteralPath $stageTargetPath -Encoding UTF8
 Write-Output ('NEW_DOCUMENT|'+$d.Name+'|'+$Label);exit
}
if($Action -notin @('windows','probe','native','screenshot')) {
    $guardWord = if($null -ne $word){$word}else{(Connect-RunningWord $WordProcessId)}
    $guardName = $guardWord.ActiveDocument.Name
    if($expectedStageDocument -and $guardName -ne $expectedStageDocument){
        $ownedModal=$false
        if(!$guardName -and $Target -eq 'dialog' -and $WordProcessId -ne 27664){
            $ownerWindow=$script:windows | Where-Object {$_.pid -eq $WordProcessId -and $_.title -eq ($expectedStageDocument+' - Word')} | Select-Object -First 1
            $modalWindow=$script:windows | Where-Object {$_.pid -eq $WordProcessId -and $_.process -eq 'WINWORD' -and $_.title -and $_.title -notlike '* - Word'} | Select-Object -First 1
            $ownedModal=$ownerWindow -and $modalWindow -and ([AuditUi]::GetAncestor([IntPtr]$modalWindow.handle,3).ToInt64() -eq $ownerWindow.handle)
        }
        if(!$ownedModal){throw "Unexpected active document: $guardName; expected $expectedStageDocument. No input was sent."}
    }
    if(($WordProcessId -eq 0 -or $WordProcessId -eq 27664) -and $guardName -match '^(文档|Document)(1|22|26|28|29|33|35|36)(\.docx)?$') {
        throw "Protected original document is active; refusing action $Action on $guardName"
    }
    $windowName = ($script:windows | Where-Object {$_.handle -eq $WindowHandle} | Select-Object -First 1).title
    if(($WordProcessId -eq 0 -or $WordProcessId -eq 27664) -and $windowName -match '^(文档|Document)(1|22|26|28|29|33|35|36)(\.docx)? - Word$') {
        throw "Protected original window selected; refusing action $Action on $windowName"
    }
}
if($Doc -match '^(文档|Document)(1|22|26|28|29|33|35|36)(\.docx)?$' -and ($WordProcessId -eq 0 -or $WordProcessId -eq 27664)){throw 'Protected original document target'}
if($Doc) { $found=$null;for($i=1;$i -le $word.Documents.Count;$i++){$candidate=$word.Documents.Item($i);if($candidate.Name -eq $Doc -or $candidate.Name -eq ($Doc+'.docx')){$found=$candidate;break}};if($null -eq $found){throw "Document not found: $Doc"};$found.Activate() }
$d=$null
if($null -ne $word) {$d=$word.ActiveDocument}
if($null -ne $word -and $Action -in @('position','select-math','select-ole','text','enter')) {[void][AuditUi]::SetForegroundWindow([IntPtr]$word.ActiveWindow.Hwnd);Start-Sleep -Milliseconds 500}
if($Action -eq 'position') {if($Position -lt 0){$Position=$d.Content.End-1};$d.Range($Position,$Position).Select();Write-Output ('POSITION|'+$d.Name+'|'+$Position);exit}
if($Action -eq 'select-math') {$d.OMaths.Item($Index).Range.Select();Write-Output ('SELECT_MATH|'+$d.Name+'|'+$Index);exit}
if($Action -eq 'select-ole') {$d.InlineShapes.Item($Index).Range.Select();Write-Output ('SELECT_OLE|'+$d.Name+'|'+$Index);exit}
if($Action -eq 'text') {if($InputFile){$Value=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $InputFile),[Text.Encoding]::UTF8)};if($Name -eq 'soft-breaks'){$Value=$Value.Replace("`r`n","`n").Replace("`n",[string][char]11)};if($Index -eq 12){$word.Selection.Font.Name=([string][char]0x5B8B+[char]0x4F53);$word.Selection.Font.Size=12};if($FontSize -gt 0){$word.Selection.Font.Size=$FontSize};if($LatinFont){$word.Selection.Font.NameAscii=$LatinFont;$word.Selection.Font.NameOther=$LatinFont};$word.Selection.TypeText($Value);Write-Output ('TYPED_TEXT|'+$d.Name+'|'+$Value.Length);exit}
if($Action -eq 'enter') {$word.Selection.TypeParagraph();Write-Output 'TYPE_PARAGRAPH';exit}
$hwnd=[IntPtr]$WindowHandle
if($WindowHandle -eq 0) {$hwnd=[IntPtr]$word.ActiveWindow.Hwnd}
if($Target -eq 'editor') {$win=$script:windows|Where-Object {$_.process -eq 'visualtex' -and $_.title -like 'Office*'}|Select-Object -First 1;if(!$win){throw 'No Office editor window'};$hwnd=[IntPtr]$win.handle}
elseif($Target -eq 'dialog') {$win=$script:windows|Where-Object {$_.title -and $_.title -notlike '* - Word' -and $_.process -eq 'WINWORD' -and ($WordProcessId -eq 0 -or $_.pid -eq $WordProcessId)}|Select-Object -First 1;if(!$win){throw 'No VisualTeX Word dialog'};$hwnd=[IntPtr]$win.handle}
elseif($Target -eq 'desktop') {$win=$script:windows|Where-Object {$_.process -eq 'visualtex' -and $_.title -notlike 'Office*'}|Select-Object -First 1;if(!$win){throw 'No VisualTeX desktop window'};$hwnd=[IntPtr]$win.handle}
if($Action -in @('native','dlgclick','dlgcheck','dlglist')) { [void][AuditUi]::SetForegroundWindow($hwnd);Start-Sleep -Milliseconds 200; $script:children=@();[void][AuditUi]::EnumChildWindows($hwnd,{param($h,$l) $s=[Text.StringBuilder]::new(8192);$c=[Text.StringBuilder]::new(256);[void][AuditUi]::GetWindowText($h,$s,8192);[void][AuditUi]::GetClassName($h,$c,256);$script:children+=@{handle=$h.ToInt64();name=$s.ToString();class=$c.ToString();id=[AuditUi]::GetDlgCtrlID($h)};return $true},[IntPtr]::Zero);if($Action -eq 'dlglist'){
$list=$script:children|Where-Object {$_.class -like '*LISTBOX*'}|Select-Object -First 1
if(!$list){throw 'Native listbox missing'}
$count=[AuditUi]::SendMessage([IntPtr]$list.handle,0x018B,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32()
if($Index -lt 1 -or $Index -gt $count){throw 'List selection outside available items'}
$itemRect=[AuditUi+Rect]::new()
$rectResult=[AuditUi]::ListItemRect([IntPtr]$list.handle,0x0198,[IntPtr]($Index-1),[ref]$itemRect).ToInt32()
if($rectResult -eq -1 -or $itemRect.Bottom -le $itemRect.Top){throw 'Native list item rectangle unavailable'}
$point=[AuditUi+Point]::new();$point.X=[int](($itemRect.Left+$itemRect.Right)/2);$point.Y=[int](($itemRect.Top+$itemRect.Bottom)/2)
if(![AuditUi]::ClientToScreen([IntPtr]$list.handle,[ref]$point)){throw 'Native list item screen position unavailable'}
[void][AuditUi]::SetCursorPos($point.X,$point.Y);[AuditUi]::mouse_event(2,0,0,0,[UIntPtr]::Zero);Start-Sleep -Milliseconds 60;[AuditUi]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
Start-Sleep -Milliseconds 200
$actual=[AuditUi]::SendMessage([IntPtr]$list.handle,0x0188,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32()
if($actual -ne ($Index-1)){throw "Native listbox selection did not match: actual=$actual expected=$($Index-1)"}
Write-Output ('NATIVE_LIST_SELECTED|'+$Index+'|count='+$count);exit
};if($Action -eq 'native'){[IO.File]::WriteAllText((Join-Path $out ($Label+'-native.json')),($script:children|ConvertTo-Json -Depth 4),[Text.UTF8Encoding]::new($false));$script:children|ConvertTo-Json -Depth 4;exit};$b=$script:children|Where-Object {$_.class -like '*BUTTON*' -and $_.name -like ('*'+$Name+'*')}|Select-Object -First 1;if(!$b){throw "Native button missing: $Name"};if($Action -eq 'dlgcheck') {$state=[AuditUi]::SendMessage([IntPtr]$b.handle,0x00F0,[IntPtr]::Zero,[IntPtr]::Zero).ToInt32();if(($state -eq 1) -eq [bool]$Checked){Write-Output ('NATIVE_CHECK_ALREADY|'+$b.name+'|'+$state);exit}};Write-Output ('UI_NATIVE_BEGIN|'+[DateTimeOffset]::UtcNow.ToString('o')+'|'+$Target+'|'+$b.name);[void][AuditUi]::PostMessage([IntPtr]$b.handle,0x00F5,[UIntPtr]::Zero,[IntPtr]::Zero);Write-Output ('UI_NATIVE_POSTED|'+[DateTimeOffset]::UtcNow.ToString('o')+'|'+$Target+'|'+$b.name);Start-Sleep -Milliseconds 600;Write-Output ('NATIVE_CLICK|'+$b.name);exit }
[void][AuditUi]::SetForegroundWindow($hwnd);Start-Sleep -Milliseconds 250
$r=[System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
if($Action -eq 'activate') {
 $titleCondition=[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::TitleBar)
 $titleBar=$r.FindFirst([System.Windows.Automation.TreeScope]::Descendants,$titleCondition)
 if($null -ne $titleBar){ClickEl $titleBar;Start-Sleep -Milliseconds 300}
 if([AuditUi]::GetForegroundWindow() -ne $hwnd){throw 'Requested Word window did not obtain foreground; Selection was not changed.'}
 Write-Output ('ACTIVATED|'+$WordProcessId+'|'+$hwnd.ToInt64());exit
}
if($Action -eq 'ribbon') {
 $tabs=$r.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::TabItem))
 $matches=@($tabs | Where-Object {$_.Current.Name -eq 'VisualTeX' -and !$_.Current.IsOffscreen -and $_.Current.IsEnabled})
 if($matches.Count -ne 1){throw 'The real visible VisualTeX Ribbon tab is not unique.'}
 $tab=$matches[0];$selectionPattern=$null
 if(!$tab.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern,[ref]$selectionPattern) -or !([System.Windows.Automation.SelectionItemPattern]$selectionPattern).Current.IsSelected){ClickEl $tab}
 Start-Sleep -Milliseconds 400;$r=[System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
}
if($Action -in @('probe','ribbon')) {if($Target -eq 'word') {$ribbonRoot=FindEl $r 'Ribbon';if($null -ne $ribbonRoot){$r=$ribbonRoot}};$tree=Tree $r;[IO.File]::WriteAllText((Join-Path $out ($Label+'-uia.json')),($tree|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false));$tree | Where-Object {$_.name -and $_.type -notin @('ControlType.Pane','ControlType.Group','ControlType.Image')} | ForEach-Object {Write-Output ($_.type+'|'+$_.name+'|'+$_.value+'|enabled='+$_.enabled+'|offscreen='+$_.offscreen)};exit}
if($Action -eq 'click') {$button=FindEl $r $Name;$started=[DateTimeOffset]::UtcNow;Write-Output ('UI_INVOKE_BEGIN|'+$started.ToString('o')+'|'+$Target+'|'+$Name);ClickEl $button;Write-Output ('UI_INVOKE_RETURN|'+[DateTimeOffset]::UtcNow.ToString('o')+'|'+$Target+'|'+$Name);Start-Sleep -Milliseconds 600;Write-Output ('CLICK|'+$Target+'|'+$Name);exit}
if($Action -eq 'combo') {
 $condition=[System.Windows.Automation.AndCondition]::new([System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty,$Name),[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::ComboBox))
 $el=$r.FindFirst([System.Windows.Automation.TreeScope]::Descendants,$condition)
 if(!$el){throw "ComboBox missing: $Name"}
 $valuePattern=$null
 if($Choice -and $el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$valuePattern) -and ([System.Windows.Automation.ValuePattern]$valuePattern).Current.Value -eq $Choice){Write-Output ('COMBO_ALREADY|'+$Name+'|'+$Choice);exit}
 if($Choice){
  ClickEl $el;Start-Sleep -Milliseconds 150
  $itemCondition=[System.Windows.Automation.AndCondition]::new([System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty,$Choice),[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::ListItem))
  $items=@($el.FindAll([System.Windows.Automation.TreeScope]::Descendants,$itemCondition)|Where-Object {$_.Current.IsEnabled -and !$_.Current.IsOffscreen})
  if($items.Count -ne 1){throw "Combo option is not uniquely visible: $Name / $Choice"}
  ClickEl $items[0];Start-Sleep -Milliseconds 400
  $actual=$null
  if(!$el.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern,[ref]$actual) -or ([System.Windows.Automation.ValuePattern]$actual).Current.Value -ne $Choice){throw "Combo did not retain the selected option: $Name / $Choice"}
  Write-Output ('COMBO_SELECTED|'+$Name+'|'+$Choice);exit
 }
 $el.SetFocus()
 for($focusAttempt=0;$focusAttempt -lt 5 -and !$el.Current.HasKeyboardFocus;$focusAttempt++){Start-Sleep -Milliseconds 100}
 if([AuditUi]::GetForegroundWindow() -ne $hwnd -or !$el.Current.HasKeyboardFocus){throw 'Requested combo did not obtain keyboard focus; no keys sent.'}
 ClickEl $el;Start-Sleep -Milliseconds 150
 if([AuditUi]::GetForegroundWindow() -ne $hwnd){throw 'Requested combo window lost foreground; no keys sent.'}
 [System.Windows.Forms.SendKeys]::SendWait($Value);Start-Sleep -Milliseconds 600;Write-Output ('COMBO|'+$Name+'|'+$Value);exit
}
if($Action -eq 'check') {$el=FindEl $r $Name;if(!$el){throw "Checkbox missing: $Name"};$p=$null;if(!$el.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern,[ref]$p)){throw 'No TogglePattern'};$on=([System.Windows.Automation.TogglePattern]$p).Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On;if($on -ne [bool]$Checked){ClickEl $el;Start-Sleep -Milliseconds 400};Write-Output ('CHECK|'+$Name+'|'+[bool]$Checked);exit}
if($Action -eq 'paste') {if($InputFile){$Value=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $InputFile),[Text.Encoding]::UTF8)};if($Name){$el=FindEl $r $Name}else{$els=$r.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty,[System.Windows.Automation.ControlType]::Edit));$el=$els.Item($Index-1)};ClickEl $el;[System.Windows.Forms.SendKeys]::SendWait('^a');PasteText $Value;Write-Output ('PASTE|'+$Target+'|'+$Value.Length);exit}
if($Action -eq 'keys') {if([AuditUi]::GetForegroundWindow() -ne $hwnd){throw 'Requested input window did not obtain foreground; no keys sent.'};[System.Windows.Forms.SendKeys]::SendWait($Value);Start-Sleep -Milliseconds 500;Write-Output ('KEYS|'+$Value);exit}
if($Action -eq 'screenshot') {$bounds=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds;$bmp=[Drawing.Bitmap]::new($bounds.Width,$bounds.Height);$g=[Drawing.Graphics]::FromImage($bmp);$g.CopyFromScreen($bounds.Location,[Drawing.Point]::Empty,$bounds.Size);$bmp.Save((Join-Path $out ($Label+'.png')),[Drawing.Imaging.ImageFormat]::Png);$g.Dispose();$bmp.Dispose();Write-Output ('SCREENSHOT|'+$Label);exit}
throw "Unknown action $Action"
