param([string]$Label='reference-jump',[int]$Index=1,[int]$ScreenX=-1,[int]$ScreenY=-1)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw | ConvertFrom-Json
if(!$target.pid -or $target.pid -eq 27664){throw 'Only a stage-owned Word process may receive navigation input.'}
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class ReferenceClick {
[StructLayout(LayoutKind.Sequential)] public struct MouseInput { public int x,y; public uint data,flags,time; public UIntPtr extra; }
[StructLayout(LayoutKind.Sequential)] public struct Input { public uint type; public MouseInput mouse; }
[DllImport("user32.dll",SetLastError=true)] public static extern uint SendInput(uint count,Input[] inputs,int size);
public static void DoubleClick() {
 var inputs=new Input[4];
 for(int i=0;i<4;i++) inputs[i].mouse.flags=(i%2==0)?2u:4u;
 if(SendInput(4,inputs,Marshal.SizeOf(typeof(Input)))!=4) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
}
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
[DllImport("user32.dll")] public static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr e);
}
'@
[void][ReferenceClick]::SetProcessDPIAware()
$word=Connect-RunningWord $target.pid
$document=$word.ActiveDocument
if(!$target.document -or $document.Name -ne $target.document){throw 'Active Word document differs from the recorded stage target; no navigation input was sent.'}
$field=$null;$seen=0
foreach($candidate in $document.Fields){
 if([int]$candidate.Type -eq 50){$seen++;if($seen -eq $Index){$field=$candidate;break}}
}
if(!$field){throw 'Requested real GOTOBUTTON reference is missing.'}
$nested=$field.Code.Fields.Item(1)
$match=[regex]::Match($nested.Code.Text,'(?i)^\s*REF\s+([^\s\\]+)')
if(!$match.Success){throw 'Nested REF target unavailable.'}
$bookmarkName=$match.Groups[1].Value
if(!$document.Bookmarks.Exists($bookmarkName)){throw 'Reference target bookmark missing.'}
$navigationMatch=[regex]::Match($field.Code.Text,'(?i)^\s*GOTOBUTTON\s+([^\s\\]+)')
if(!$navigationMatch.Success){throw 'Real GOTOBUTTON navigation target is missing.'}
$navigationName=$navigationMatch.Groups[1].Value
if(!$document.Bookmarks.Exists($navigationName)){throw 'Navigation bookmark missing.'}
$visible=$nested.Result
$window=$document.ActiveWindow
[void][ReferenceClick]::SetForegroundWindow([IntPtr]$window.Hwnd)
$window.ScrollIntoView($visible,$true)
Start-Sleep -Milliseconds 500
[int]$left=0;[int]$top=0;[int]$width=0;[int]$height=0
$window.GetPoint([ref]$left,[ref]$top,[ref]$width,[ref]$height,$visible)
$actualReferenceRectangle=@{left=$left;top=$top;width=$width;height=$height}
$completeReferenceRectangle=$null
if($ScreenX -lt 0 -and $ScreenY -lt 0 -and ($width -le 0 -or $height -le 0)){
 # A nested REF is inside GOTOBUTTON's code. Word returns zero width for that
 # hidden-code range even while its text is visible. The complete real field
 # has the visible rectangle; obtain it from Word rather than guessing pixels.
 $completeReference=$document.Range($field.Code.Start-1,$field.Result.End+1)
 $window.GetPoint([ref]$left,[ref]$top,[ref]$width,[ref]$height,$completeReference)
 $completeReferenceRectangle=@{left=$left;top=$top;width=$width;height=$height}
}
if($ScreenX -ge 0 -and $ScreenY -ge 0){
 if($height -gt 0 -and ($ScreenY -lt $top -or $ScreenY -ge ($top+$height))){throw 'Observed point is outside the current reference text line; no click was sent.'}
 $hit=$window.RangeFromPoint($ScreenX,$ScreenY)
 if($hit.Start -lt ($field.Code.Start-1) -or $hit.End -gt ($field.Result.End+1)){throw 'Observed screen point does not belong to the real reference field.'}
 $left=$ScreenX;$top=$ScreenY;$width=1;$height=1
}
elseif($width -le 0 -or $height -le 0){throw 'Visible reference screen rectangle unavailable.'}
$clickX=$left+[int]($width/2);$clickY=$top+[int]($height/2)
$clickRange=$window.RangeFromPoint($clickX,$clickY)
if($clickRange.Start -lt ($field.Code.Start-1) -or $clickRange.End -gt ($field.Result.End+1)){throw 'Resolved point does not belong to the real reference field; no click was sent.'}
$before=@{start=$word.Selection.Start;end=$word.Selection.End}
[void][ReferenceClick]::SetCursorPos($clickX,$clickY)
[ReferenceClick]::DoubleClick()
Start-Sleep -Milliseconds 650
$destination=$document.Bookmarks.Item($navigationName).Range
$after=@{start=$word.Selection.Start;end=$word.Selection.End;text=$word.Selection.Text}
$passed=$after.start -ge $destination.Start -and $after.end -le $destination.End
$expandedField=$null
if(!$passed){
 foreach($targetField in $document.Fields){
  if([int]$targetField.Type -ne 12){continue}
  $targetCode=$targetField.Code;$targetResult=$targetField.Result
  if($targetResult.Start -le $destination.Start -and $targetResult.End -ge $destination.End -and
     $after.start -eq ($targetCode.Start-1) -and $after.end -eq ($targetResult.End+1) -and
     $after.text -eq $destination.Text){
   $expandedField=@{type=[int]$targetField.Type;code=$targetCode.Text;start=$targetCode.Start-1;end=$targetResult.End+1;resultStart=$targetResult.Start;resultEnd=$targetResult.End}
   $passed=$true;break
  }
 }
}
$macroTarget=$null
if(!$passed -and $after.start -eq $after.end){
 foreach($targetField in $document.Fields){
  if([int]$targetField.Type -ne 51 -or $targetField.Code.Text -notmatch '^\s*MACROBUTTON\s+MTPlaceRef\b'){continue}
  $targetCode=$targetField.Code
  if($after.start -eq ($targetCode.Start-1) -and $targetCode.Start -lt $destination.Start -and $destination.End -le $targetCode.End -and $targetCode.StoryType -eq $destination.StoryType){
   $macroTarget=@{type=51;codeStart=$targetCode.Start;codeEnd=$targetCode.End;code=$targetCode.Text;nativeSelectionType=[int]$word.Selection.Type}
   $passed=$true;break
  }
 }
}
$selectionMatched=$passed
$destinationFrames=$destination.Frames.Count
$passed=$passed -and $destinationFrames -eq 0
$record=@{time=[DateTime]::UtcNow.ToString('o');pid=$target.pid;document=$document.Name;action='real mouse double-click';bookmark=$bookmarkName;navigationBookmark=$navigationName;target=@{start=$destination.Start;end=$destination.End;text=$destination.Text;frames=$destinationFrames};rectangle=@{left=$left;top=$top;width=$width;height=$height};actualReferenceRectangle=$actualReferenceRectangle;before=$before;after=$after;wordExpandedTargetField=$expandedField;wordMacroButtonTarget=$macroTarget;selectionMatched=$selectionMatched;passed=$passed;visualReviewRequired=$true}
$record.completeReferenceRectangle=$completeReferenceRectangle
$record | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$record | ConvertTo-Json -Depth 6
if(!$passed){throw 'Real reference double-click did not select its target bookmark.'}
