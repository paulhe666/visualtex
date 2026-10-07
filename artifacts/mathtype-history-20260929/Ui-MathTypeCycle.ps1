param(
 [Parameter(Mandatory=$true)][string]$StatePath,
 [Parameter(Mandatory=$true)][string]$Label,
 [int]$Start=-1,[int]$End=-1,
 [ValidateSet('edit','inline','display','convert','refresh')][string]$Mode='edit',
 [ValidateSet('mathTypeOle','nativeOmml')][string]$ObjectMode='mathTypeOle',
 [string]$Append='+2',[string]$Source='',[switch]$Numbered,
 [int]$WaitSeconds=25,[switch]$OpenOnly,[int]$ExpectedConversions=1,[switch]$SnapshotNativePreflight,
 [switch]$WholeDocument,[switch]$AllowConversionFailure,[switch]$AuthorizedLiveDocument,[switch]$StopAfterConversionClick
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
. (Join-Path $repo 'docs\remediation-3d207d7\word-connection.ps1')
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Windows.Forms
Add-Type @'
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class CycleUi {
 public sealed class Window {public long Handle; public uint Pid; public string Title;}
 public delegate bool Callback(IntPtr h,IntPtr l);
 [DllImport("user32.dll")] public static extern bool EnumWindows(Callback c,IntPtr l);
 [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr d,Callback c,IntPtr l);
 [DllImport("user32.dll",SetLastError=true)] public static extern IntPtr OpenInputDesktop(uint f,bool i,uint a);
 [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr d);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a,uint b,bool attach);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int n);
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] public static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr e);
 [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
 public static Window[] Visible(){var list=new List<Window>();Callback cb=(h,l)=>{var b=new StringBuilder(512);GetWindowText(h,b,512);uint p;GetWindowThreadProcessId(h,out p);if(IsWindowVisible(h))list.Add(new Window{Handle=h.ToInt64(),Pid=p,Title=b.ToString()});return true;};var d=OpenInputDesktop(0,false,0x01C1);try{if(d!=IntPtr.Zero)EnumDesktopWindows(d,cb,IntPtr.Zero);else EnumWindows(cb,IntPtr.Zero);}finally{if(d!=IntPtr.Zero)CloseDesktop(d);}return list.ToArray();}
}
'@
[void][CycleUi]::SetProcessDPIAware()
$state=Get-Content -LiteralPath $StatePath -Raw -Encoding UTF8|ConvertFrom-Json
$process=Get-Process -Id $state.WordPid
$liveAuthorized=$AuthorizedLiveDocument -and [bool]$state.AuthorizedLiveDocument
if((!$state.OwnedTestProcess -and !$liveAuthorized) -or $process.ProcessName -ne 'WINWORD' -or $process.StartTime.ToUniversalTime().ToString('o') -cne $state.WordStarted){throw 'Authorized Word identity mismatch.'}
$out=Join-Path $PSScriptRoot $Label
if(Test-Path -LiteralPath $out){throw 'Evidence directory already exists.'}
[void][IO.Directory]::CreateDirectory($out)
$trace=[Collections.Generic.List[object]]::new()
function Record([string]$stage,$data=$null){$trace.Add([ordered]@{Time=(Get-Date).ToString('o');Stage=$stage;Data=$data});$trace|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $out 'ui-events.json') -Encoding UTF8;Write-Host $stage}
function SnapshotBeforeMutation([string]$stage){
 if(!$SnapshotNativePreflight){return}
 if($doc.Paragraphs.Count -gt 100 -or $word.UndoRecord.IsRecordingCustomRecord){throw 'Read-only stage snapshots require a small fixture outside all Undo records.'}
 $xml=$doc.Content.WordOpenXML
 [IO.File]::WriteAllText((Join-Path $out ($stage+'.xml')),$xml,[Text.UTF8Encoding]::new($false))
 Record ('snapshot-'+$stage) @{Saved=$doc.Saved;Tracking=$doc.TrackRevisions}
}
function NormalizeTestSource([string]$value){
 $v=$value.Replace("`r`n","`n").Trim()
 foreach($shell in @('$$','$')){
  if($v.Length -ge 2*$shell.Length -and $v.StartsWith($shell) -and $v.EndsWith($shell)){
   return $shell+'|'+$v.Substring($shell.Length,$v.Length-2*$shell.Length).Trim()
  }
 }
 return $v
}
function SendPhysicalSource([string]$value){
 [System.Windows.Forms.SendKeys]::SendWait('^a')
 Start-Sleep -Milliseconds 80
 for($i=0;$i -lt $value.Length;$i++){
  $ch=$value[$i]
  if($ch -eq "`r"){
   if($i+1 -lt $value.Length -and $value[$i+1] -eq "`n"){$i++}
   [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
   continue
  }
  if($ch -eq "`n"){[System.Windows.Forms.SendKeys]::SendWait('{ENTER}');continue}
  $key=switch($ch){
   '+' {'{+}'}
   '^' {'{^}'}
   '%' {'{%}'}
   '~' {'{~}'}
   '(' {'{(}'}
   ')' {'{)}'}
   '[' {'{[}'}
   ']' {'{]}'}
   '{' {'{{}'}
   '}' {'{}}'}
   default {[string]$ch}
  }
  [System.Windows.Forms.SendKeys]::SendWait($key)
 }
}
function Release($v){if($null -ne $v -and [Runtime.InteropServices.Marshal]::IsComObject($v)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}}
function Focus([IntPtr]$h,[bool]$maximize=$false){
 [uint32]$pid0=0;$t=[CycleUi]::GetWindowThreadProcessId($h,[ref]$pid0);$f=[CycleUi]::GetWindowThreadProcessId([CycleUi]::GetForegroundWindow(),[ref]$pid0);$c=[CycleUi]::GetCurrentThreadId();$attached=@()
 try{foreach($x in @($t,$f)|Select-Object -Unique){if($x -ne 0 -and $x -ne $c -and [CycleUi]::AttachThreadInput($c,$x,$true)){$attached+=$x}};if($maximize){[void][CycleUi]::ShowWindow($h,3)}elseif([CycleUi]::IsIconic($h)){[void][CycleUi]::ShowWindow($h,9)};[void][CycleUi]::BringWindowToTop($h);[void][CycleUi]::SetForegroundWindow($h)}finally{foreach($x in $attached){[void][CycleUi]::AttachThreadInput($c,$x,$false)}}
 Start-Sleep -Milliseconds 150
 if([CycleUi]::GetForegroundWindow() -ne $h){throw 'Exact owned window did not obtain foreground; no input sent.'}
}
function Root([IntPtr]$h){[Windows.Automation.AutomationElement]::FromHandle($h)}
function Find([IntPtr]$h,[string]$name,[string]$kind=''){
 $matches=@((Root $h).FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$name))|Where-Object {!$_.Current.IsOffscreen -and $_.Current.IsEnabled -and (!$kind -or $_.Current.ControlType.ProgrammaticName -eq $kind)})
 if($matches.Count -ne 1){throw "Expected one visible $kind '$name'; actual=$($matches.Count)."};return $matches[0]
}
function ClickEl([IntPtr]$h,$e){
 Focus $h
 if($e.Current.ControlType -eq [Windows.Automation.ControlType]::Edit){
  try{
   $e.SetFocus();Start-Sleep -Milliseconds 100
   $focused=[Windows.Automation.AutomationElement]::FocusedElement;$walker=[Windows.Automation.TreeWalker]::RawViewWalker
   for($j=0;$j -lt 12 -and $focused;$j++){
    if([Windows.Automation.Automation]::Compare($focused,$e)){return}
    $focused=$walker.GetParent($focused)
   }
  }catch{}
 }
 # UIAutomation invokes the actual visible Ribbon/WebView control; never call
 # add-in/COM product methods as acceptance. Prefer a non-coordinate action for
 # buttons because responsive layout can move the element between UI turns.
 $invoke=$null
 if($e.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$invoke)){
  ([Windows.Automation.InvokePattern]$invoke).Invoke();Start-Sleep -Milliseconds 250;return
 }
 $tp=$null
 if($e.Current.ControlType -eq [Windows.Automation.ControlType]::CheckBox -and $e.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern,[ref]$tp)){
  ([Windows.Automation.TogglePattern]$tp).Toggle();Start-Sleep -Milliseconds 250;return
 }
 $sp=$null
 if($e.Current.ControlType -eq [Windows.Automation.ControlType]::TabItem -and $e.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sp)){
  ([Windows.Automation.SelectionItemPattern]$sp).Select();Start-Sleep -Milliseconds 250;return
 }
 $r=$e.Current.BoundingRectangle
 if($r.IsEmpty -or $r.Width -le 0 -or $r.Height -le 0){throw 'Cannot click an invisible element.'}
 $cx=[int]($r.X+$r.Width/2);$cy=[int]($r.Y+$r.Height/2)
 $hit=[Windows.Automation.AutomationElement]::FromPoint([Windows.Point]::new($cx,$cy));$walker=[Windows.Automation.TreeWalker]::RawViewWalker;$verified=$false
 for($j=0;$j -lt 12 -and $hit;$j++){if([Windows.Automation.Automation]::Compare($hit,$e)){$verified=$true;break};$hit=$walker.GetParent($hit)}
 if(!$verified){throw 'The click point no longer belongs to the inspected UI element; no mouse event sent.'}
 [void][CycleUi]::SetCursorPos($cx,$cy);[CycleUi]::mouse_event(2,0,0,0,[UIntPtr]::Zero);Start-Sleep -Milliseconds 50;[CycleUi]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
 Start-Sleep -Milliseconds 250
}
function Click([IntPtr]$h,[string]$name,[string]$kind='ControlType.Button'){ClickEl $h (Find $h $name $kind);Record 'click' @{Window=$h.ToInt64();Name=$name}}
function Editor(){
 $list=@()
 foreach($win in [CycleUi]::Visible()){
  if(!$win.Title.StartsWith('Office ')){continue};$p=Get-Process -Id $win.Pid
  if($p.ProcessName -ne 'visualtex' -or $p.Path -ine (Join-Path $env:LOCALAPPDATA 'VisualTeX\visualtex.exe')){continue}
  $r=Root ([IntPtr]$win.Handle)
  $docs=$r.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Document))
  foreach($d in $docs){$pat=$null;if($d.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$pat)){$url=([Windows.Automation.ValuePattern]$pat).Current.Value;if($url -match '[?&]sessionId=([^&]+)'){$list+=@{Handle=[IntPtr]$win.Handle;SessionId=$Matches[1];Url=$url}}}}
 }
 if($list.Count -eq 1){return $list[0]};if($list.Count -gt 1){throw 'Multiple active Office sessions; do not choose an arbitrary editor.'};return $null
}
$word=$null;$doc=$null;$range=$null
try{
 Record 'connection-before' @{Pid=$process.Id}
 $word=Connect-RunningWord $process.Id
 Record 'connection-after' @{Pid=$process.Id}
 $doc=$word.ActiveDocument
 if($doc.FullName -ine $state.DocumentPath){throw 'Not the owned test copy.'}
 Record 'document-verified' @{Pid=$process.Id}
 $hwnd=[IntPtr]$doc.ActiveWindow.Hwnd
 Record 'window-resolved' @{Window=$hwnd.ToInt64()}
 Record 'start' @{Pid=$process.Id;Candidate=$state.CandidateSha256;Mode=$Mode;ObjectMode=$ObjectMode;Numbered=[bool]$Numbered;Tracking=$doc.TrackRevisions;Start=$Start;End=$End;AuthorizedLiveDocument=$liveAuthorized}
 SnapshotBeforeMutation 'pre-selection'
 Focus $hwnd $true
 if($Start -ge 0){if($End -lt 0){$End=$Start};$range=$doc.Range($Start,$End);if($Mode -in @('edit','convert')){$required=if($Mode -eq 'convert'){$ExpectedConversions}else{1};if($ObjectMode -eq 'mathTypeOle' -and $range.InlineShapes.Count -ne $required){throw 'Selection does not own the expected OLE count.'};if($ObjectMode -eq 'nativeOmml' -and $range.OMaths.Count -ne 1){throw 'Edit must own exactly one OMath.'}};if($Mode -notin @('edit','convert') -and $Start -ne $End){throw 'Insertion requires an exact caret.'};$range.Select();$doc.ActiveWindow.ScrollIntoView($range,$true)}
 SnapshotBeforeMutation 'post-selection'
 if(Editor){throw 'An Office editor is already active; preserve it.'}
 $tab=Find $hwnd 'VisualTeX' 'ControlType.TabItem';$sp=$null
 if(!$tab.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern,[ref]$sp) -or !([Windows.Automation.SelectionItemPattern]$sp).Current.IsSelected){ClickEl $hwnd $tab}
 if($Mode -eq 'refresh'){
  $log=Join-Path $state.Directory 'word-hook.log';$before=if(Test-Path $log){(Get-Item $log).Length}else{0}
  Record 'refresh-start' @{LogBytesBefore=$before}
  Click $hwnd '更新公式编号'
  $deadline=[DateTime]::UtcNow.AddSeconds($WaitSeconds);$done=$false;$tail=''
  do{
   Start-Sleep -Milliseconds 250;$process.Refresh()
   if(Test-Path $log){$stream=[IO.File]::Open($log,'Open','Read','ReadWrite');try{[void]$stream.Seek($before,'Begin');$reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8);try{$tail=$reader.ReadToEnd()}finally{$reader.Dispose()}}finally{$stream.Dispose()}}
   $done=$tail -match 'stage=transaction-before-end'
  }while(!$done -and !$process.HasExited -and [DateTime]::UtcNow -lt $deadline)
  [IO.File]::WriteAllText((Join-Path $out 'operation-hook.log'),$tail,[Text.UTF8Encoding]::new($false))
  Record 'refresh-observed' @{Completed=$done;Responding=$process.Responding}
  if(!$done -or $tail -notmatch 'stage=transaction-before-end recording=True level=1'){throw 'Refresh did not retain its actual native transaction.'}
  return
 }
 if($Mode -eq 'convert'){
  # Inspect the currently pressed Ribbon scope through the real split menu.
  # Opening the menu is read-only; do not click a scope item or rewrite HKCU.
  $split=Find $hwnd 'MathType → OMML' 'ControlType.SplitButton'
  $more=@($split.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,'更多选项'))|Where-Object {!$_.Current.IsOffscreen -and $_.Current.IsEnabled})
  if($more.Count -ne 1){throw "Expected one MathType → OMML scope menu opener; actual=$($more.Count)."}
  ClickEl $hwnd $more[0]
  $scopeDeadline=[DateTime]::UtcNow.AddSeconds(5);$scopeElement=$null;$scopeWindow=[IntPtr]::Zero
  $scopeName=if($WholeDocument){'默认：转换全文'}else{'默认：转换所选'}
  do{
   foreach($candidate in @([CycleUi]::Visible()|Where-Object {$_.Pid -eq [uint32]$process.Id})){
    try{
     $matches=@((Root ([IntPtr]$candidate.Handle)).FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$scopeName))|Where-Object {!$_.Current.IsOffscreen -and $_.Current.IsEnabled})
     if($matches.Count -eq 1){$scopeElement=$matches[0];$scopeWindow=[IntPtr]$candidate.Handle;break}
    }catch{}
   }
   if(!$scopeElement){Start-Sleep -Milliseconds 100}
  }while(!$scopeElement -and [DateTime]::UtcNow -lt $scopeDeadline)
  if(!$scopeElement){throw "Could not inspect current Ribbon conversion scope '$scopeName'; no conversion click sent."}
  $pressed=$false;$toggle=$null;$selection=$null
  if($scopeElement.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern,[ref]$toggle)){$pressed=([Windows.Automation.TogglePattern]$toggle).Current.ToggleState -eq [Windows.Automation.ToggleState]::On}
  elseif($scopeElement.TryGetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern,[ref]$selection)){$pressed=([Windows.Automation.SelectionItemPattern]$selection).Current.IsSelected}
  Record 'conversion-scope-inspected' @{Name=$scopeName;Pressed=$pressed;Window=$scopeWindow.ToInt64();Type=$scopeElement.Current.ControlType.ProgrammaticName}
  [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
  if(!$pressed){throw "The live Ribbon is not set to '$scopeName'; preserve preference and do not invoke."}
  if($ObjectMode -ne 'mathTypeOle'){throw 'Conversion requires MathType source mode.'}
  if($WholeDocument){
   if($Start -ge 0 -or $End -ge 0){throw 'Whole-document conversion must not prepare a selection.'}
  }elseif($Start -lt 0 -or $Start -ge $End){throw 'Conversion must select one exact MathType source.'}
  $log=Join-Path $state.Directory 'word-hook.log';$before=if(Test-Path $log){(Get-Item $log).Length}else{0}
  Record 'convert-start' @{LogBytesBefore=$before;WholeDocument=[bool]$WholeDocument}
  Click $hwnd 'MathType → OMML'
  if($StopAfterConversionClick){Record 'convert-click-dispatched' @{WholeDocument=[bool]$WholeDocument};return}
  if($WholeDocument){
   $confirmDeadline=[DateTime]::UtcNow.AddSeconds($WaitSeconds);$confirmed=$false
   do{
    Start-Sleep -Milliseconds 250
    foreach($candidate in @([CycleUi]::Visible()|Where-Object {$_.Pid -eq [uint32]$process.Id -and $_.Title -eq 'VisualTeX 公式格式转换'})){
     try{
      $dialogRoot=Root ([IntPtr]$candidate.Handle)
      $buttons=$dialogRoot.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Button))
      $yes=@($buttons|Where-Object {!$_.Current.IsOffscreen -and $_.Current.IsEnabled -and ($_.Current.Name -eq '是(&Y)' -or $_.Current.AutomationId -eq '6')})
      if($yes.Count -eq 1){
       Record 'conversion-confirmation-found' @{Window=$candidate.Handle;Title=$candidate.Title;Button=$yes[0].Current.Name;AutomationId=$yes[0].Current.AutomationId}
       ClickEl ([IntPtr]$candidate.Handle) $yes[0]
       Record 'conversion-confirmation-clicked' @{Window=$candidate.Handle;Button=$yes[0].Current.Name}
       $confirmed=$true
       break
      }
     }catch{Record 'conversion-confirmation-probe-raced' @{Window=$candidate.Handle;Error=$_.Exception.Message}}
    }
   }while(!$confirmed -and !$process.HasExited -and [DateTime]::UtcNow -lt $confirmDeadline)
   if(!$confirmed){throw 'Whole-document confirmation did not appear; preserve the owned Word and evidence.'}
  }
  $deadline=[DateTime]::UtcNow.AddSeconds($WaitSeconds);$done=$false;$tail='';$terminalDialog=$null
  do{
   Start-Sleep -Milliseconds 300;$process.Refresh()
   if(Test-Path $log){$stream=[IO.File]::Open($log,'Open','Read','ReadWrite');try{[void]$stream.Seek($before,'Begin');$reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8);try{$tail=$reader.ReadToEnd()}finally{$reader.Dispose()}}finally{$stream.Dispose()}}
   $done=$tail -match 'format-conversion-(complete|stopped|failed) '
   if(!$done){
    foreach($candidate in @([CycleUi]::Visible()|Where-Object {$_.Pid -eq [uint32]$process.Id -and $_.Title -eq 'VisualTeX 公式格式转换'})){
     try{
      $dialogRoot=Root ([IntPtr]$candidate.Handle)
      $children=$dialogRoot.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)
      $names=@($children|ForEach-Object {$_.Current.Name}|Where-Object {$_})
      if(@($names|Where-Object {$_ -like '*The MathType source changed while its conversion was being prepared.*'}).Count -gt 0){
       $terminalDialog=[ordered]@{Handle=$candidate.Handle;Title=$candidate.Title;Names=$names}
       $terminalDialog|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $out 'source-changed-dialog.json') -Encoding UTF8
       Record 'source-changed-dialog-observed' $terminalDialog
       $done=$true
       break
      }
     }catch{}
    }
   }
  }while(!$done -and !$process.HasExited -and [DateTime]::UtcNow -lt $deadline)
  [IO.File]::WriteAllText((Join-Path $out 'operation-hook.log'),$tail,[Text.UTF8Encoding]::new($false))
  $dialogs=@()
  foreach($candidate in @([CycleUi]::Visible()|Where-Object {$_.Pid -eq [uint32]$process.Id})){
   try{
    $dialogRoot=Root ([IntPtr]$candidate.Handle)
    $children=$dialogRoot.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)
    $named=@($children|ForEach-Object {if($_.Current.Name){[ordered]@{Name=$_.Current.Name;Type=$_.Current.ControlType.ProgrammaticName;AutomationId=$_.Current.AutomationId;Enabled=$_.Current.IsEnabled;Offscreen=$_.Current.IsOffscreen}}})
    if($candidate.Handle -ne $hwnd.ToInt64() -or $candidate.Title -eq 'VisualTeX 公式格式转换'){$dialogs+=,[ordered]@{Handle=$candidate.Handle;Title=$candidate.Title;Children=$named}}
   }catch{$dialogs+=,[ordered]@{Handle=$candidate.Handle;Title=$candidate.Title;ProbeError=$_.Exception.Message}}
  }
  $dialogs|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $out 'terminal-dialogs.json') -Encoding UTF8
  Record 'convert-observed' @{Completed=$done;Responding=$process.Responding;Dialogs=$dialogs.Count;SourceChangedDialog=($null -ne $terminalDialog)}
  if($AllowConversionFailure){if(!$done){throw 'Conversion produced no terminal trace; preserve evidence.'};return}
  if(!$done -or $tail -notmatch ('format-conversion-complete .*converted='+$ExpectedConversions+' failed=0')){throw 'Conversion did not report the exact expected successful target count; preserve evidence.'}
  return
 }
 $button=switch($Mode){'edit'{'编辑所选公式'}'inline'{if($ObjectMode -eq 'nativeOmml'){'OMML 行内公式'}else{'OLE 行内公式'}}'display'{if($ObjectMode -eq 'nativeOmml'){'OMML 行间公式'}else{'OLE 行间公式'}}}
 Click $hwnd $button
 $deadline=[DateTime]::UtcNow.AddSeconds(12);$editor=$null
 do{$editor=Editor;if(!$editor){Start-Sleep -Milliseconds 200}}while(!$editor -and [DateTime]::UtcNow -lt $deadline)
 if(!$editor){throw 'Editor did not open.'};$eh=$editor.Handle
 Record 'editor-open' $editor
 SnapshotBeforeMutation 'post-editor-open'
 $state|Add-Member VisualTeXEditorHandle $eh.ToInt64() -Force
 $ep=([CycleUi]::Visible()|Where-Object Handle -eq $eh.ToInt64())[0].Pid;$edp=Get-Process -Id $ep
 $state|Add-Member VisualTeXEditorPid $ep -Force;$state|Add-Member VisualTeXEditorStarted $edp.StartTime.ToUniversalTime().ToString('o') -Force
 $state|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $StatePath -Encoding UTF8
 if($OpenOnly){return}
 # A reused WebView window exposes its new session URL before the editor model
 # has finished hydrating. Do not write into the old control during navigation.
 Start-Sleep -Milliseconds 1200
 Focus $eh
 if($ObjectMode -eq 'mathTypeOle'){
 $format=Find $eh '公式对象格式' 'ControlType.ComboBox';$vp=$null
 if(!$format.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$vp)){throw 'Format value not readable.'}
 if(([Windows.Automation.ValuePattern]$vp).Current.Value -ne 'MathType OLE'){
  ClickEl $eh $format;Click $eh 'MathType OLE' 'ControlType.ListItem'
 }
 }
 # Numbered display is supported by both native OMML and MathType. Verify the
 # actual checkbox for either mode; do not rely on persisted editor preferences.
 if($Mode -eq 'display'){$box=Find $eh '编号' 'ControlType.CheckBox';$tp=$null;if(!$box.TryGetCurrentPattern([Windows.Automation.TogglePattern]::Pattern,[ref]$tp)){throw 'Numbering not readable.'};if((([Windows.Automation.TogglePattern]$tp).Current.ToggleState -eq [Windows.Automation.ToggleState]::On) -ne [bool]$Numbered){ClickEl $eh $box}}
 $candidates=@()
 for($attempt=0;$attempt -lt 2;$attempt++){
  $edits=(Root $eh).FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Edit))
  $candidates=@(foreach($e in $edits){$vp=$null;$b=$e.Current.BoundingRectangle;if(!$e.Current.IsOffscreen -and $e.Current.IsEnabled -and !$b.IsEmpty -and $b.Width -gt 50 -and $b.Height -gt 40 -and $e.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$vp)){@{Element=$e;Value=([Windows.Automation.ValuePattern]$vp).Current.Value;Pattern=$vp}}})
  if($candidates.Count -eq 1){break};Click $eh 'LaTeX 源码' 'ControlType.TabItem'
 }
 if($candidates.Count -ne 1){throw "Source pane not unique: $($candidates.Count)."}
 $old=[string]$candidates[0].Value
 Record 'source-ready' @{SessionId=$editor.SessionId;Initial=$old}
 if($Source){$new=$Source}
 elseif($Mode -ne 'edit'){$new=if($Mode -eq 'inline'){'$x+1$'}else{'$$x+1$$'}}
 else{if($Append -notmatch '^\+[0-9]+$'){throw 'Only numeric test term append allowed.'};$trim=$old.TrimEnd();$shell=if($trim.EndsWith('$$')){'$$'}elseif($trim.EndsWith('$')){'$'}else{throw 'Unexpected source wrapper.'};$new=$trim.Substring(0,$trim.Length-$shell.Length).TrimEnd()+$Append+"`r`n"+$shell}
 ClickEl $eh $candidates[0].Element
 # Set the inspected actual source edit control through accessibility. This is
 # native UI input, not a service/host callback, and avoids racing the global
 # clipboard with a user or another Office editor. If WebView drops the first
 # accessibility write during hydration, allow one verified retry ONLY while the
 # same owned session still contains the exact untouched pre-write source.
 for($inputAttempt=0;$inputAttempt -lt 3;$inputAttempt++){
  if($inputAttempt -gt 0){
   $liveEditor=Editor
   if(!$liveEditor -or $liveEditor.SessionId -cne $editor.SessionId -or $liveEditor.Handle -ne $eh){throw 'Owned editor session changed before input retry; no further input sent.'}
   Record 'source-input-not-applied-retry' @{SessionId=$editor.SessionId;Original=$actual;Attempt=$inputAttempt+1}
   Start-Sleep -Milliseconds 700
   Focus $eh
   try{Click $eh 'LaTeX 源码' 'ControlType.TabItem'}catch{}
   # WebView can replace the accessibility Edit node while hydrating. Never
   # reuse a stale ValuePattern across retries; reacquire the one live source
   # control and prove it still contains the untouched original before writing.
   $retryEdits=(Root $eh).FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Edit))
   $retryCandidates=@(foreach($e in $retryEdits){$vp=$null;$b=$e.Current.BoundingRectangle;if(!$e.Current.IsOffscreen -and $e.Current.IsEnabled -and !$b.IsEmpty -and $b.Width -gt 50 -and $b.Height -gt 40 -and $e.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$vp)){@{Element=$e;Value=([Windows.Automation.ValuePattern]$vp).Current.Value;Pattern=$vp}}})
   if($retryCandidates.Count -ne 1){throw "Source pane not unique during input retry: $($retryCandidates.Count)."}
   $candidates=$retryCandidates
   $actual=[string]$candidates[0].Pattern.Current.Value
   if((NormalizeTestSource $actual) -cne (NormalizeTestSource $old)){break}
  }
  ClickEl $eh $candidates[0].Element
  ([Windows.Automation.ValuePattern]$candidates[0].Pattern).SetValue($new)
  $valueDeadline=[DateTime]::UtcNow.AddSeconds(3)
  do{Start-Sleep -Milliseconds 150;$actual=[string]$candidates[0].Pattern.Current.Value}while((NormalizeTestSource $actual) -cne (NormalizeTestSource $new) -and [DateTime]::UtcNow -lt $valueDeadline)
  if((NormalizeTestSource $actual) -ceq (NormalizeTestSource $new)){break}
  if((NormalizeTestSource $actual) -cne (NormalizeTestSource $old)){break}
 }
 # The actual editor can normalize only the whitespace immediately inside the
 # selected dollar shell. Preserve/compare all mathematical body text verbatim.
 if((NormalizeTestSource $actual) -cne (NormalizeTestSource $new)){Record 'source-mismatch' @{Expected=$new;Actual=$actual};throw 'UI source does not match intended test edit.'}
 Record 'source-edited' @{Old=$old;New=$actual}
 # ValuePattern.SetValue can update Chromium's accessibility value without
 # dispatching the same DOM input/change path as a user's keystroke. Probe that
 # state with one reversible real key event. If the controlled editor rerenders
 # the OLD value, re-enter the complete intended source with physical keyboard
 # input on this exact owned/focused Edit control. Never treat synthetic visible
 # text as proof that the product model received a user edit.
 Focus $eh
 $candidates[0].Element.SetFocus()
 [System.Windows.Forms.SendKeys]::SendWait("^{END} {BACKSPACE}")
 Start-Sleep -Milliseconds 250
 $actual=[string]$candidates[0].Pattern.Current.Value
 if((NormalizeTestSource $actual) -ceq (NormalizeTestSource $old)){
  Record 'source-synthetic-state-stale' @{Expected=$new;Actual=$actual;SessionId=$editor.SessionId}
  Focus $eh
  $candidates[0].Element.SetFocus()
  SendPhysicalSource $new
  Start-Sleep -Milliseconds 350
  # A controlled WebView can replace the accessibility node while real input is
  # processed. Reacquire the single live source editor before validating it.
  $physicalEdits=(Root $eh).FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Edit))
  $physicalCandidates=@(foreach($e in $physicalEdits){$vp=$null;$b=$e.Current.BoundingRectangle;if(!$e.Current.IsOffscreen -and $e.Current.IsEnabled -and !$b.IsEmpty -and $b.Width -gt 50 -and $b.Height -gt 40 -and $e.TryGetCurrentPattern([Windows.Automation.ValuePattern]::Pattern,[ref]$vp)){@{Element=$e;Value=([Windows.Automation.ValuePattern]$vp).Current.Value;Pattern=$vp}}})
  if($physicalCandidates.Count -ne 1){throw "Source pane not unique after physical input: $($physicalCandidates.Count)."}
  $candidates=$physicalCandidates
  $actual=[string]$candidates[0].Pattern.Current.Value
  Record 'source-physical-reentry' @{SessionId=$editor.SessionId;Value=$actual}
 }else{
  Record 'source-physical-sync' @{SessionId=$editor.SessionId;Value=$actual}
 }
 if((NormalizeTestSource $actual) -cne (NormalizeTestSource $new)){Record 'source-physical-sync-mismatch' @{Expected=$new;Actual=$actual};throw 'Physical source synchronization did not produce the intended test text.'}
 # Give the editor's source -> MathML/export synchronization one UI turn before
 # clicking Apply; real typing naturally has this gap.
 Start-Sleep -Milliseconds 350
 $log=Join-Path $state.Directory 'word-hook.log'
 $before=if(Test-Path $log){(Get-Item $log).Length}else{0}
 SnapshotBeforeMutation 'pre-apply'
 Record 'apply-start' @{SessionId=$editor.SessionId;LogBytesBefore=$before}
 $sw=[Diagnostics.Stopwatch]::StartNew()
 $apply=if($Mode -eq 'edit'){'更新公式'}else{'完成并插入'}
 Click $eh $apply
 $deadline=[DateTime]::UtcNow.AddSeconds($WaitSeconds);$tail='';$done=$false
 do{
  Start-Sleep -Milliseconds 300;$process.Refresh()
  if(Test-Path $log){$stream=[IO.File]::Open($log,'Open','Read','ReadWrite');try{[void]$stream.Seek($before,'Begin');$reader=[IO.StreamReader]::new($stream,[Text.Encoding]::UTF8);try{$tail=$reader.ReadToEnd()}finally{$reader.Dispose()}}finally{$stream.Dispose()}}
  $done=$tail.Contains('ribbon-session-operation-released sessionId='+$editor.SessionId) -or $tail.Contains('ribbon-session-failed')
 }while(!$done -and !$process.HasExited -and [DateTime]::UtcNow -lt $deadline)
 [IO.File]::WriteAllText((Join-Path $out 'operation-hook.log'),$tail,[Text.UTF8Encoding]::new($false))
 Record 'apply-observed' @{ElapsedMs=$sw.ElapsedMilliseconds;Completed=$done;Responding=$process.Responding;Exited=$process.HasExited;SessionId=$editor.SessionId}
 if(!$done -or $tail.Contains('ribbon-session-failed')){throw 'Operation did not complete successfully; preserve process and evidence.'}
 Start-Sleep -Milliseconds 300;$sel=$word.Selection.Range
 try{$local=$sel.Paragraphs[1].Range;$xml=$local.WordOpenXML;[IO.File]::WriteAllText((Join-Path $out 'after-paragraph.xml'),$xml,[Text.UTF8Encoding]::new($false));Record 'after-local' @{SelectionStart=$sel.Start;SelectionEnd=$sel.End;ParagraphStart=$local.Start;ParagraphEnd=$local.End;Text=$local.Text;Alignment=[int]$local.ParagraphFormat.Alignment;LocalShapes=$local.InlineShapes.Count;LocalOMaths=$local.OMaths.Count;TrackRevisions=$doc.TrackRevisions};Release $local}finally{Release $sel}
}finally{Release $range;Release $doc;Release $word}
