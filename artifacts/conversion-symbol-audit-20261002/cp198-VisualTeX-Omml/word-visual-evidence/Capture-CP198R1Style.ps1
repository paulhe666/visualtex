param(
 [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$InputManifestPath,
 [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$OutputDirectory,
 [switch]$ValidateOnly,
 [ValidateSet('Direct','Shell')][string]$LaunchMode='Direct',
 [switch]$ActivateApplication,
 [string]$WordExecutable='C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
)
# This retains the original capture API. It does not change protection settings or bypass AMSI.
# A build-evidence file documents the saved input's origin; this script performs no conversion.
$ErrorActionPreference='Stop'
$root=$null;$source=$null;$formulaSource=$null;$owned=$null;$manifest=$null
$outputCreated=$false;$originalShowAll=$null;$failure=$null;$cleanupErrors=@()
$process=$null;$native=$null;$app=$null;$doc=$null;$plain=$null;$header=$null;$ownedVerified=$false
$summary=[ordered]@{kind='owned-saved-OMML-Word-screenshot';startedUtc=[DateTime]::UtcNow;status='validating-input';captures=@();sourceMetadata=@();noConversionOrRenderInvoked=$true;validateOnly=[bool]$ValidateOnly}
function Save-Summary {
 if($outputCreated){$summary | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath (Join-Path $root 'capture-summary.json') -Encoding UTF8}
}
function Release-Com($o) { if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}} }
function Resolve-InputFile([string]$value,[string]$baseDirectory) {
 if([string]::IsNullOrWhiteSpace($value)){throw 'Manifest input path is empty.'}
 if(-not [IO.Path]::IsPathRooted($value)){$value=Join-Path $baseDirectory $value}
 $path=[IO.Path]::GetFullPath($value)
 if(-not [IO.File]::Exists($path)){throw ('Input file not found: '+$path)}
 return $path
}
function Assert-FileHash([string]$path,[string]$expected) {
 if($expected -notmatch '^[A-Fa-f0-9]{64}$'){throw ('Missing or invalid SHA256 for '+$path)}
 $actual=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
 if($actual -ine $expected){throw ('Input SHA256 mismatch: '+$path)}
 return $actual
}
function Initialize-CaptureHost {
Add-Type -AssemblyName System.Drawing,System.Windows.Forms,System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class OwnedWordCapture {
 public delegate bool EnumerateWindow(IntPtr hwnd, IntPtr data);
 [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left,Top,Right,Bottom; }
 [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X,Y; public POINT(int x,int y){X=x;Y=y;} }
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumerateWindow callback,IntPtr data);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent,EnumerateWindow callback,IntPtr data);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd,StringBuilder name,int length);
 [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint objectId,ref Guid iid,[MarshalAs(UnmanagedType.Interface)]out object result);
 [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd,out RECT rect);
 [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern IntPtr GetThreadDesktop(uint threadId);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd,int n);
 [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
 [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint first,uint second,bool attach);
 [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern IntPtr SetActiveWindow(IntPtr hwnd);
 public static void ActivateOwned(IntPtr hwnd,int wantedPid){
  if(Pid(hwnd)!=(uint)wantedPid)throw new Exception("Cannot activate non-owned window.");
  uint unused;uint current=GetCurrentThreadId();uint target=GetWindowThreadProcessId(hwnd,out unused);uint foreground=GetWindowThreadProcessId(GetForegroundWindow(),out unused);
  bool a=false,b=false;
  try{
   if(current!=foreground)a=AttachThreadInput(current,foreground,true);
   if(current!=target)b=AttachThreadInput(current,target,true);
   ShowWindow(hwnd,3);BringWindowToTop(hwnd);SetForegroundWindow(hwnd);SetActiveWindow(hwnd);SetFocus(hwnd);
  }finally{if(b)AttachThreadInput(current,target,false);if(a)AttachThreadInput(current,foreground,false);}
 }
 [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT point);
 public static uint Pid(IntPtr hwnd){uint p;GetWindowThreadProcessId(hwnd,out p);return p;}
 public static object FindNative(int wantedPid) {
  var handles=new List<IntPtr>();
  EnumWindows((window,unused)=>{uint p;GetWindowThreadProcessId(window,out p);if(p==(uint)wantedPid)
   EnumChildWindows(window,(child,x)=>{var name=new StringBuilder(128);GetClassName(child,name,name.Capacity);if(name.ToString()=="_WwG")handles.Add(child);return true;},IntPtr.Zero);return true;},IntPtr.Zero);
  foreach(var h in handles){var iid=new Guid("00020400-0000-0000-C000-000000000046");object value;if(AccessibleObjectFromWindow(h,0xFFFFFFF0,ref iid,out value)==0 && value!=null)return value;}
  return null;
 }
}
'@
[void][OwnedWordCapture]::SetProcessDPIAware()
}
function Read-Metadata {
 $zip=[IO.Compression.ZipFile]::OpenRead($formulaSource)
 try {
  $reader=New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
  try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $ns=New-Object Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $ns.AddNamespace('v','urn:schemas-microsoft-com:vml')
  $ns.AddNamespace('o','urn:schemas-microsoft-com:office:office')
  $ns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $objects=@($xml.SelectNodes('//w:object[o:OLEObject]',$ns))
  $summary.oleCount=$objects.Count
  for($i=0;$i -lt $objects.Count;$i++){
   $obj=$objects[$i];$shape=$obj.SelectSingleNode('v:shape',$ns);$ole=$obj.SelectSingleNode('o:OLEObject',$ns)
   $para=$obj.SelectSingleNode('ancestor::w:p[1]',$ns)
   $alt=$shape.GetAttribute('alt');$meta=$null
   if($alt.StartsWith('visualtex:v1:deflate:')){
    $encoded=$alt.Substring('visualtex:v1:deflate:'.Length).Replace('-','+').Replace('_','/')
    while($encoded.Length % 4){$encoded+='='}
    $bytes=[Convert]::FromBase64String($encoded)
    $stream=New-Object IO.MemoryStream(,$bytes)
    $deflate=New-Object IO.Compression.DeflateStream($stream,[IO.Compression.CompressionMode]::Decompress)
    $reader=New-Object IO.StreamReader($deflate)
    try{$meta=$reader.ReadToEnd() | ConvertFrom-Json}finally{$reader.Dispose();$deflate.Dispose();$stream.Dispose()}
   }
   $summary.sourceMetadata+=,[ordered]@{sourceIndex=($i+1);progId=$ole.GetAttribute('ProgID');shapeId=$shape.GetAttribute('id');shapeStyle=$shape.GetAttribute('style');paragraphText=($para.SelectNodes('.//w:t',$ns)|ForEach-Object{$_.InnerText}) -join '';paragraphStyle=[string]$para.SelectSingleNode('w:pPr/w:pStyle/@w:val',$ns).Value;previewMetadata=$meta}
  }
  if($objects.Count -ne $summary.expectedFormulaCount -or @($summary.sourceMetadata|Where-Object {$_.progId -ne 'VisualTeX.Formula.1'}).Count -ne 0){throw 'Saved VisualTeX OLE count/type does not match the manifest.'}
 }finally{$zip.Dispose()}
}
function Assert-Owned {
 if($process.HasExited -or $app.Documents.Count -ne 1 -or $doc.FullName -ine $owned -or [OwnedWordCapture]::Pid([IntPtr][int]$native.Hwnd) -ne $process.Id){throw 'Owned document/process proof no longer matches.'}
 $h=[OwnedWordCapture]::GetAncestor([IntPtr][int]$doc.ActiveWindow.Hwnd,2)
 if([OwnedWordCapture]::Pid($h) -ne $process.Id){throw 'Owned root HWND PID mismatch.'}
 return $h
}
function Record-ActivationState([string]$stage,$activationResult=$null) {
 $h=Assert-Owned
 $f=[OwnedWordCapture]::GetAncestor([OwnedWordCapture]::GetForegroundWindow(),2)
 [uint32]$ownedPid=0;[uint32]$foregroundPid=0
 $ownedThread=[OwnedWordCapture]::GetWindowThreadProcessId($h,[ref]$ownedPid)
 $foregroundThread=[OwnedWordCapture]::GetWindowThreadProcessId($f,[ref]$foregroundPid)
 $current=[Diagnostics.Process]::GetCurrentProcess()
 $foregroundSession=$null
 if($foregroundPid -gt 0){try{$foregroundSession=[Diagnostics.Process]::GetProcessById([int]$foregroundPid).SessionId}catch{}}
 $entry=[ordered]@{stage=$stage;utc=[DateTime]::UtcNow;activationResult=$activationResult;ownedHwnd=$h.ToInt64();foregroundHwnd=$f.ToInt64();ownedPid=$ownedPid;foregroundPid=$foregroundPid;ownedSession=$process.SessionId;callerSession=$current.SessionId;foregroundSession=$foregroundSession;ownedThread=$ownedThread;foregroundThread=$foregroundThread;ownedDesktop=[OwnedWordCapture]::GetThreadDesktop($ownedThread).ToInt64();callerDesktop=[OwnedWordCapture]::GetThreadDesktop([OwnedWordCapture]::GetCurrentThreadId()).ToInt64();foregroundDesktop=[OwnedWordCapture]::GetThreadDesktop($foregroundThread).ToInt64();ownedMinimized=[OwnedWordCapture]::IsIconic($h);ownedVisible=[OwnedWordCapture]::IsWindowVisible($h)}
 if(-not $summary.Contains('activationTrace')){$summary.activationTrace=@()}
 $summary.activationTrace+=,$entry
 Save-Summary
}
function Assert-Foreground {
 $h=Assert-Owned
 $f=[OwnedWordCapture]::GetAncestor([OwnedWordCapture]::GetForegroundWindow(),2)
 if($f -ne $h -or [OwnedWordCapture]::Pid($f) -ne $process.Id){throw ('Foreground root mismatch; no screenshot taken. Expected hwnd='+$h+' pid='+$process.Id+'; actual hwnd='+$f+' pid='+[OwnedWordCapture]::Pid($f))}
 return $h
}
function Capture-Rect([string]$name,[int]$x,[int]$y,[int]$width,[int]$height,[int]$index,[string]$scope,$point) {
 $h=Assert-Foreground
 $wr=New-Object OwnedWordCapture+RECT
 if(-not [OwnedWordCapture]::GetWindowRect($h,[ref]$wr)){throw 'Could not inspect owned window rectangle.'}
 $screen=[Windows.Forms.SystemInformation]::VirtualScreen
 $left=[Math]::Max([Math]::Max($x,$wr.Left),$screen.Left)
 $top=[Math]::Max([Math]::Max($y,$wr.Top),$screen.Top)
 $right=[Math]::Min([Math]::Min($x+$width,$wr.Right),$screen.Right)
 $bottom=[Math]::Min([Math]::Min($y+$height,$wr.Bottom),$screen.Bottom)
 $width=$right-$left;$height=$bottom-$top
 if($width -lt 10 -or $height -lt 10){throw 'Capture rect is empty or too small.'}
 foreach($xx in @(($left+1),([int](($left+$right)/2)),($right-2))){
  foreach($yy in @(($top+1),([int](($top+$bottom)/2)),($bottom-2))){
   $at=[OwnedWordCapture]::WindowFromPoint((New-Object OwnedWordCapture+POINT($xx,$yy)))
   if([OwnedWordCapture]::GetAncestor($at,2) -ne $h){throw 'Capture area includes a window outside owned Word; refused.'}
  }
 }
 $bmp=$null;$g=$null
 $path=Join-Path $root ($name+'.png')
 try {
  if(Test-Path -LiteralPath $path){throw 'Screenshot output already exists; refusing overwrite.'}
  $bmp=New-Object Drawing.Bitmap($width,$height)
  $g=[Drawing.Graphics]::FromImage($bmp)
  $g.CopyFromScreen($left,$top,0,0,$bmp.Size)
  [void](Assert-Foreground)
  $bmp.Save($path,[Drawing.Imaging.ImageFormat]::Png)
 }finally{if($null -ne $g){$g.Dispose()};if($null -ne $bmp){$bmp.Dispose()}}
 $record=[ordered]@{name=$name;path=$path;width=$width;height=$height;left=$left;top=$top;scope=$scope;sourceIndex=$index;getPoint=$point;foregroundRootHwnd=$h.ToInt64();foregroundPid=$process.Id;timestampUtc=[DateTime]::UtcNow;sha256=(Get-FileHash $path -Algorithm SHA256).Hash}
 $summary.captures+=,$record;Save-Summary
 Write-Output ('ACTUAL_WORD_IMAGE '+($record|ConvertTo-Json -Compress -Depth 10))
}
function Capture-Formula([int]$index,[string]$label) {
 $math=$null;$range=$null
 try {
  [void](Assert-Owned)
  $math=$doc.OMaths.Item($index);$range=$math.Range
  $doc.ActiveWindow.ScrollIntoView($range,$true)
  Start-Sleep -Milliseconds 450
  [void](Assert-Foreground)
  [int]$l=0;[int]$t=0;[int]$w=0;[int]$ht=0
  $doc.ActiveWindow.GetPoint([ref]$l,[ref]$t,[ref]$w,[ref]$ht,$range)
  $point=[ordered]@{left=$l;top=$t;width=$w;height=$ht;rangeStart=[int]$range.Start;rangeEnd=[int]$range.End;mathType=[int]$math.Type;mathText=[string]$range.Text;paragraphText=[string]$range.Paragraphs.Item(1).Range.Text;hostFontName=[string]$range.Font.Name;hostFontSize=[double]$range.Font.Size;rangeFontPosition=[double]$range.Font.Position}
  $name=$label
  $point.selectionStart=[int]$app.Selection.Start;$point.selectionEnd=[int]$app.Selection.End;$point.selectionMathCount=[int]$app.Selection.OMaths.Count;$point.selectionText=[string]$app.Selection.Text
  Capture-Rect $name ($l-20) ($t-20) ($w+40) ($ht+40) $index 'actual saved target Word native OMML; formula indices refer to the explicit input manifest' $point
 }finally{Release-Com $range;Release-Com $math}
}
try {
 $manifestPath=Resolve-InputFile $InputManifestPath (Get-Location).Path
 $manifestDirectory=Split-Path -Parent $manifestPath
 $manifest=Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
 if($manifest.schemaVersion -ne 1){throw 'Expected capture manifest schemaVersion 1.'}
 if([string]::IsNullOrWhiteSpace([string]$manifest.runLabel) -or [string]::IsNullOrWhiteSpace([string]$manifest.buildLabel)){throw 'Manifest must explicitly name the run and saved-input build.'}
 if([string]$manifest.expectedFormulaCount -notmatch '^[1-9][0-9]*$'){throw 'Expected formula count must be a positive integer.'}
 $summary.expectedFormulaCount=[int]$manifest.expectedFormulaCount
 $indices=@($manifest.formulaIndices)
 if($indices.Count -eq 0){throw 'No formulas selected.'}
 foreach($index in $indices){if([string]$index -notmatch '^[1-9][0-9]*$' -or [int]$index -gt $summary.expectedFormulaCount){throw 'Formula index is outside the declared source range.'}}
 if(@($indices | Select-Object -Unique).Count -ne $indices.Count){throw 'Duplicate formula indices would overwrite evidence.'}
 $root=[IO.Path]::GetFullPath($OutputDirectory)
 if(Test-Path -LiteralPath $root){throw 'Output directory already exists; choose a fresh evidence directory.'}
 $source=Resolve-InputFile $manifest.sourcePath $manifestDirectory
 $formulaSource=Resolve-InputFile $manifest.formulaSourcePath $manifestDirectory
 $buildEvidence=Resolve-InputFile $manifest.buildEvidencePath $manifestDirectory
 $summary.source=$source;$summary.formulaSource=$formulaSource
 $summary.sourceSha256Before=Assert-FileHash $source $manifest.sourceSha256
 $summary.formulaSourceSha256Before=Assert-FileHash $formulaSource $manifest.formulaSourceSha256
 $summary.buildEvidenceSha256Before=Assert-FileHash $buildEvidence $manifest.buildEvidenceSha256
 $summary.buildEvidencePath=$buildEvidence;$summary.buildLabel=$manifest.buildLabel;$summary.runLabel=$manifest.runLabel
 $summary.formulaIndices=$indices;$summary.inputManifestPath=$manifestPath
 $summary.inputManifestSha256=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
 $summary.scriptSha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
 $WordExecutable=Resolve-InputFile $WordExecutable (Get-Location).Path
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 Read-Metadata
 $zip=[IO.Compression.ZipFile]::OpenRead($source)
 try{
  $entry=$zip.GetEntry('word/document.xml')
  if($null -eq $entry){throw 'Target document has no word/document.xml.'}
  $reader=New-Object IO.StreamReader($entry.Open())
  try{[xml]$targetXml=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $targetNs=New-Object Xml.XmlNamespaceManager($targetXml.NameTable)
  $targetNs.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
  $summary.savedOmmlCount=@($targetXml.SelectNodes('//m:oMath',$targetNs)).Count
  if($summary.savedOmmlCount -ne $summary.expectedFormulaCount){throw 'Saved native OMML count does not match the manifest.'}
 }finally{$zip.Dispose()}
 if($ValidateOnly){$summary.status='validated';return}
 $null=New-Item -ItemType Directory -Path $root -ErrorAction Stop
 $outputCreated=$true;$owned=Join-Path $root 'owned-document.docx';$summary.owned=$owned
 Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $root 'input-manifest.json')
 Save-Summary
 Initialize-CaptureHost
 Copy-Item -LiteralPath $source -Destination $owned
 $summary.ownedSha256Before=(Get-FileHash $owned -Algorithm SHA256).Hash
 $protected=@(Get-Process WINWORD -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
 $summary.protectedWordPids=$protected
 $start=New-Object Diagnostics.ProcessStartInfo
 $start.FileName=$WordExecutable
 $start.Arguments='/x "'+$owned+'"';$start.UseShellExecute=($LaunchMode -eq 'Shell');$start.WorkingDirectory=$root
 $summary.launchMode=$LaunchMode;$summary.activateApplication=[bool]$ActivateApplication
 $process=[Diagnostics.Process]::Start($start)
 if($null -eq $process){throw 'Launch returned no owned process handle; no document access attempted.'}
 if($protected -contains $process.Id){throw 'Returned an existing Word PID.'}
 $summary.wordPid=$process.Id;$summary.wordStartedUtc=$process.StartTime.ToUniversalTime();$summary.wordCommand=$start.FileName+' '+$start.Arguments
 $summary.status='waiting-for-owned-native';Save-Summary
 Write-Output ('OWNED_WORD_STARTED pid='+$process.Id)
 $deadline=[DateTime]::UtcNow.AddSeconds(40)
 while([DateTime]::UtcNow -lt $deadline -and -not $process.HasExited){
  try{
   $native=[OwnedWordCapture]::FindNative($process.Id)
   if($null -ne $native){
    $app=$native.Application;$doc=$native.Document
    if([OwnedWordCapture]::Pid([IntPtr][int]$native.Hwnd) -ne $process.Id -or $doc.FullName -ine $owned -or $app.Documents.Count -ne 1){throw 'Native object model ownership mismatch.'}
    $ownedVerified=$true;break
   }
  }catch [Runtime.InteropServices.COMException]{}
  Start-Sleep -Milliseconds 100
 }
 if(-not $ownedVerified){throw 'Owned Word native object model unavailable.'}
 $summary.ownedPathVerified=$doc.FullName;$summary.ownedPidVerified=[OwnedWordCapture]::Pid([IntPtr][int]$native.Hwnd)
 $summary.omathCount=[int]$doc.OMaths.Count
 if($summary.omathCount -ne $summary.expectedFormulaCount){throw 'Opened native OMML count does not match the manifest.'}
 Record-ActivationState 'before-visible-activate'
 $app.DisplayAlerts=0;$app.Visible=$true;$doc.Activate()
 Record-ActivationState 'after-document-activate'
 $h=Assert-Owned
 [void][OwnedWordCapture]::ShowWindow($h,3)
 $doc.ActiveWindow.View.Type=3
 $doc.ActiveWindow.View.Zoom.Percentage=150
 Start-Sleep -Milliseconds 2500
 # Do not send global keys to whichever application currently has focus.
 if($ActivateApplication){
  Record-ActivationState 'before-application-activate'
  $app.Activate()
  Record-ActivationState 'after-application-activate'
 }
 Record-ActivationState 'before-set-foreground'
 $foregroundAccepted=[OwnedWordCapture]::SetForegroundWindow($h)
 Record-ActivationState 'after-set-foreground' $foregroundAccepted
 $doc.Activate()
 Record-ActivationState 'after-second-document-activate'
 Start-Sleep -Milliseconds 150
 Record-ActivationState 'before-first-foreground-check'
 [void](Assert-Foreground)
 $summary.zoomPercentage=150;$summary.status='capturing';Save-Summary
 $originalShowAll=[bool]$doc.ActiveWindow.View.ShowAll
 $summary.showAllBefore=$originalShowAll
 $doc.ActiveWindow.View.ShowAll=$false
 $summary.showAllDuring=[bool]$doc.ActiveWindow.View.ShowAll
 $header=$doc.Paragraphs.Item(1).Range
 if($header.OMaths.Count -ne 0 -or $header.Text.Length -lt 3){throw 'First paragraph is not verified plain prose.'}
 $summary.headerText=[string]$header.Text
 $plain=$doc.Range(($header.Start+2),($header.Start+2))
 $plain.Select()
 if($app.Selection.OMaths.Count -ne 0){throw 'Header selection unexpectedly contains math.'}
 foreach($index in $indices){Capture-Formula ([int]$index) ('F{0:D2}-Word-OMML' -f [int]$index)}
 if($summary.captures.Count -ne $indices.Count){throw 'Capture count does not match the selected formulas.'}
 $summary.status='complete'
}catch{
 $failure=$_
 $summary.status='failed';$summary.error=$_.Exception.ToString();$summary.errorLine=$_.InvocationInfo.PositionMessage;$summary.errorStack=$_.ScriptStackTrace
 Write-Output ('CAPTURE_FAILED '+$_.Exception.Message)
}finally{
 Release-Com $plain;Release-Com $header
 if($ownedVerified -and $null -ne $app){
  try{
   [void](Assert-Owned)
   if($null -ne $originalShowAll){$doc.ActiveWindow.View.ShowAll=$originalShowAll;$summary.showAllRestored=[bool]$doc.ActiveWindow.View.ShowAll}
   $doc.Close(0);$app.Quit(0);$summary.ownedWordClosed=$true
  }catch{$cleanupErrors+=('Owned Word cleanup: '+$_.Exception.Message)}
 }
 Release-Com $doc;Release-Com $app;Release-Com $native
 if($null -ne $process){
  try{[void]$process.WaitForExit(5000);$summary.ownedWordExited=$process.HasExited;if(-not $process.HasExited){$cleanupErrors+='Owned Word remains running; no process was forcibly terminated.'}}catch{$cleanupErrors+=('Process exit check: '+$_.Exception.Message)}
  finally{$process.Dispose()}
 }
 foreach($item in @(@{key='source';path=$source},@{key='formulaSource';path=$formulaSource},@{key='buildEvidence';path=$summary.buildEvidencePath},@{key='owned';path=$owned})){
  $key=$item.key
  if($null -ne $summary[($key+'Sha256Before')]){
   try{
    $summary[($key+'Sha256After')]=(Get-FileHash -LiteralPath $item.path -Algorithm SHA256).Hash
    $summary[($key+'Unchanged')]=($summary[($key+'Sha256Before')] -eq $summary[($key+'Sha256After')])
    if(-not $summary[($key+'Unchanged')]){$cleanupErrors+=($key+' changed during the run.')}
   }catch{$cleanupErrors+=($key+' final hash: '+$_.Exception.Message)}
  }
 }
 $summary.cleanupErrors=$cleanupErrors
 if($cleanupErrors.Count -gt 0){$summary.status='failed';if($null -eq $failure){$summary.error=$cleanupErrors -join '; '}}
 $summary.completedUtc=[DateTime]::UtcNow
 try{Save-Summary}catch{$summary.status='failed';if($null -eq $failure){$failure=$_};$summary.summaryWriteError=$_.Exception.Message}
 [ordered]@{status=$summary.status;validateOnly=[bool]$ValidateOnly;runLabel=$summary.runLabel;buildLabel=$summary.buildLabel;wordPid=$summary.wordPid;ownedWordClosed=$summary.ownedWordClosed;ownedWordExited=$summary.ownedWordExited;sourceUnchanged=$summary.sourceUnchanged;ownedUnchanged=$summary.ownedUnchanged;captureCount=$summary.captures.Count;error=$summary.error;cleanupErrors=$cleanupErrors;summaryPath=$(if($outputCreated){Join-Path $root 'capture-summary.json'}else{$null})}|ConvertTo-Json -Depth 10
 if($summary.status -eq 'failed'){
  if($null -ne $failure){throw $failure}
  throw ($cleanupErrors -join '; ')
 }
}
