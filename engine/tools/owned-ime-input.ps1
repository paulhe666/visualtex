param(
  [Parameter(Mandatory=$true)][int]$OwnedProcessId,
  [Parameter(Mandatory=$true)][string]$ExpectedTitle,
  [ValidateSet('inspect','activate','type','key','layout','capture')][string]$Action='inspect',
  [string]$Text='',
  [ValidateSet('space','enter','escape','backspace','shift','shiftenter')][string]$Key='space',
  [long]$KeyboardLayout=0,
  [string]$OutputPath=''
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
if ($ExpectedTitle -notmatch '^VisualTeX-IME-[a-z0-9-]{8,80}$') { throw 'An exact owned test title is required' }
$process=Get-CimInstance Win32_Process -Filter "ProcessId=$OwnedProcessId"
if (!$process -or $process.Name -notin @('msedge.exe','chrome.exe') -or $process.CommandLine -notmatch 'visualtex-owned-svg-review-') {
  throw 'The target is not the isolated browser process launched by this test'
}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class OwnedImeInput {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr data);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr data);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern IntPtr GetLastActivePopup(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool GetGUIThreadInfo(uint thread,ref GUIINFO info);
  [StructLayout(LayoutKind.Sequential)] public struct GUIINFO {public uint size,flags;public IntPtr active,focus,capture,menuOwner,moveSize,caret;public RECT caretRect;}
  public static GUIINFO Gui(uint thread) {var info=new GUIINFO();info.size=(uint)Marshal.SizeOf(typeof(GUIINFO));if(!GetGUIThreadInfo(thread,ref info))throw new InvalidOperationException("GUI thread context unavailable");return info;}
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hwnd, bool altTab);
  [DllImport("user32.dll")] public static extern IntPtr GetKeyboardLayout(uint thread);
  [DllImport("user32.dll")] public static extern int GetKeyboardLayoutList(int count, [Out] IntPtr[] list);
  [DllImport("user32.dll")] public static extern short GetAsyncKeyState(int key);
  [DllImport("user32.dll",SetLastError=true)] public static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wp, IntPtr lp);
  [DllImport("user32.dll",SetLastError=true)] public static extern uint SendInput(uint count, INPUT[] input, int size);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
  [DllImport("user32.dll",SetLastError=true)] public static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x,int y,int width,int height,uint flags);
  [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW",SetLastError=true)] public static extern IntPtr GetWindowLongPtr(IntPtr hwnd,int index);
  [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT point);
  [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT point);
  [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);
  [DllImport("user32.dll")] public static extern int GetSystemMetrics(int index);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr value);
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X,Y; }
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left,Top,Right,Bottom; }
  [StructLayout(LayoutKind.Sequential)] public struct KEYBOARD {public ushort vk,scan; public uint flags,time; public UIntPtr extra;}
  [StructLayout(LayoutKind.Sequential)] public struct MOUSE {public int x,y;public uint data,flags,time;public UIntPtr extra;}
  [StructLayout(LayoutKind.Explicit)] public struct UNION {[FieldOffset(0)]public KEYBOARD keyboard;[FieldOffset(0)]public MOUSE mouse;}
  [StructLayout(LayoutKind.Sequential)] public struct INPUT {public uint type;public UNION data;}
  public static string Title(IntPtr hwnd) {var b=new StringBuilder(1024);GetWindowText(hwnd,b,b.Capacity);return b.ToString();}
  public static IntPtr[] Windows(uint pid,string title) {
    var found=new List<IntPtr>();EnumWindows((h,d)=>{uint p;GetWindowThreadProcessId(h,out p);
      if(p==pid&&IsWindowVisible(h)&&Title(h).StartsWith(title,StringComparison.Ordinal))found.Add(h);return true;},IntPtr.Zero);return found.ToArray();
  }
  public static void ClickOwnedTitlebar(IntPtr target) {
    foreach(int mod in new[]{0x10,0x11,0x12,0x5B,0x5C})if((GetAsyncKeyState(mod)&0x8000)!=0)
      throw new InvalidOperationException("A user modifier is held; no activation click sent");
    // Temporarily raise ONLY the UUID/PID-validated owned test window. Restore
    // its original topmost state even if exposure verification or input fails.
    bool wasTopmost=(GetWindowLongPtr(target,-20).ToInt64()&8)!=0;
    try {
    if(!SetWindowPos(target,new IntPtr(-1),0,0,0,0,0x0013))
      throw new InvalidOperationException("Could not expose the owned test window: "+Marshal.GetLastWin32Error());
    System.Threading.Thread.Sleep(200);
    RECT rect;if(!GetWindowRect(target,out rect))throw new InvalidOperationException("Owned window bounds unavailable");
    POINT p=new POINT{X=rect.Left+(rect.Right-rect.Left)/2,Y=rect.Top+16};
    var covering=GetAncestor(WindowFromPoint(p),2);
    if(covering!=target)throw new InvalidOperationException("Owned titlebar is not exposed; no activation click sent. point="+p.X+","+p.Y+" bounds="+rect.Left+","+rect.Top+","+rect.Right+","+rect.Bottom+" covering="+covering+" title="+Title(covering));
    int left=GetSystemMetrics(76),top=GetSystemMetrics(77),width=GetSystemMetrics(78),height=GetSystemMetrics(79);
    if(width<=1||height<=1)throw new InvalidOperationException("Desktop coordinates unavailable");
    var move=new INPUT[1];
    move[0].data.mouse.x=(int)Math.Round((p.X-left)*65535.0/(width-1));
    move[0].data.mouse.y=(int)Math.Round((p.Y-top)*65535.0/(height-1));
    move[0].data.mouse.flags=0xC001;
    if(SendInput(1,move,Marshal.SizeOf(typeof(INPUT)))!=1)throw new InvalidOperationException("Owned activation mouse move was not accepted");
    System.Threading.Thread.Sleep(120);
    POINT actual;
    if(!GetCursorPos(out actual))throw new InvalidOperationException("Cannot verify actual mouse location; no click sent");
    var actualRoot=GetAncestor(WindowFromPoint(actual),2);
    if(actualRoot!=target || Math.Abs(actual.X-p.X)>3 || Math.Abs(actual.Y-p.Y)>3)
      throw new InvalidOperationException("Mouse coordinates or target differ after motion; no click sent. requested="+p.X+","+p.Y+" actual="+actual.X+","+actual.Y+" desktop="+left+","+top+","+width+","+height+" actualRoot="+actualRoot+" title="+Title(actualRoot));
    var clicks=new INPUT[2];clicks[0].data.mouse.flags=2;clicks[1].data.mouse.flags=4;
    if(SendInput(2,clicks,Marshal.SizeOf(typeof(INPUT)))!=2)throw new InvalidOperationException("Owned activation mouse click was not accepted");
    System.Threading.Thread.Sleep(150);
    } finally {
      if(!wasTopmost)SetWindowPos(target,new IntPtr(-2),0,0,0,0,0x0013);
    }
  }
  public static void Press(IntPtr target,ushort key) {
    if(GetForegroundWindow()!=target)throw new InvalidOperationException("Owned test window lost foreground; no key sent");
    foreach(int mod in new[]{0x10,0x11,0x12,0x5B,0x5C})if((GetAsyncKeyState(mod)&0x8000)!=0)
      throw new InvalidOperationException("A user modifier is held; no key sent");
    var input=new INPUT[2];input[0].type=1;input[0].data.keyboard.vk=key;
    input[1].type=1;input[1].data.keyboard.vk=key;input[1].data.keyboard.flags=2;
    if(SendInput(2,input,Marshal.SizeOf(typeof(INPUT)))!=2)throw new InvalidOperationException("SendInput did not accept both test key events");
  }
  public static void PressShiftEnter(IntPtr target) {
    if(GetForegroundWindow()!=target)throw new InvalidOperationException("Owned test window lost foreground; no chord sent");
    foreach(int mod in new[]{0x10,0x11,0x12,0x5B,0x5C})if((GetAsyncKeyState(mod)&0x8000)!=0)
      throw new InvalidOperationException("A user modifier is held; no chord sent");
    var input=new INPUT[4];
    input[0].type=1;input[0].data.keyboard.vk=0x10;
    input[1].type=1;input[1].data.keyboard.vk=0x0D;
    input[2].type=1;input[2].data.keyboard.vk=0x0D;input[2].data.keyboard.flags=2;
    input[3].type=1;input[3].data.keyboard.vk=0x10;input[3].data.keyboard.flags=2;
    if(SendInput(4,input,Marshal.SizeOf(typeof(INPUT)))!=4)throw new InvalidOperationException("SendInput did not accept the owned Shift+Enter chord");
  }
}
'@
# Only this test process's current thread changes coordinate interpretation.
# No user's system DPI or application settings are changed.
[void][OwnedImeInput]::SetThreadDpiAwarenessContext([IntPtr](-4))
$windows=@([OwnedImeInput]::Windows([uint32]$OwnedProcessId,$ExpectedTitle))
if($windows.Count -ne 1){throw "Expected exactly one owned visible test window, found $($windows.Count)"}
$window=$windows[0]
[uint32]$windowPid=0
$thread=[OwnedImeInput]::GetWindowThreadProcessId($window,[ref]$windowPid)
$beforeLayout=[OwnedImeInput]::GetKeyboardLayout($thread).ToInt64()
$beforeForeground=[OwnedImeInput]::GetForegroundWindow().ToInt64()
$popupInfo=$null
$popup=[OwnedImeInput]::GetLastActivePopup($window)
if($Action -eq 'inspect' -and $popup -ne $window -and [OwnedImeInput]::IsWindowVisible($popup)) {
  [uint32]$popupPid=0
  [void][OwnedImeInput]::GetWindowThreadProcessId($popup,[ref]$popupPid)
  $popupInfo=[ordered]@{hwnd=$popup.ToInt64();processId=$popupPid;title=[OwnedImeInput]::Title($popup);elements=@()}
  if($popupPid -eq $OwnedProcessId) {
    Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
    $rootElement=[System.Windows.Automation.AutomationElement]::FromHandle($popup)
    $elements=$rootElement.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
    $popupInfo.elements=@(for($i=0;$i -lt [Math]::Min(100,$elements.Count);$i++) {
      $item=$elements.Item($i).Current
      if($item.Name -or $item.ControlType.ProgrammaticName -eq 'ControlType.Button') {
        [ordered]@{name=$item.Name;type=$item.ControlType.ProgrammaticName;automationId=$item.AutomationId;enabled=$item.IsEnabled}
      }
    })
  }
}
if($Action -eq 'activate') {
  $activationNotes=New-Object 'System.Collections.Generic.List[string]'
  $activationNotes.Add('initialForeground='+[OwnedImeInput]::GetForegroundWindow().ToInt64()+' title='+[OwnedImeInput]::Title([OwnedImeInput]::GetForegroundWindow()))
  $activated=[OwnedImeInput]::SetForegroundWindow($window)
  $activationNotes.Add('SetForegroundWindow='+$activated)
  Start-Sleep -Milliseconds 200
  if([OwnedImeInput]::GetForegroundWindow() -ne $window){
    # Standard accessibility focus on this exact owned window. Do not attach
    # input threads, send keys to the user's window, or change focus policies.
    Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
    try {
      $element=[System.Windows.Automation.AutomationElement]::FromHandle($window)
      $element.SetFocus()
    } catch { $activationNotes.Add('UIA='+$_.Exception.Message) }
    Start-Sleep -Milliseconds 200
  }
  if([OwnedImeInput]::GetForegroundWindow() -ne $window){
    try {[OwnedImeInput]::ClickOwnedTitlebar($window)} catch { $activationNotes.Add('ownedTitlebar='+$_.Exception.Message) }
    Start-Sleep -Milliseconds 250
  }
  if([OwnedImeInput]::GetForegroundWindow() -ne $window){
    try {[OwnedImeInput]::SwitchToThisWindow($window,$true);$activationNotes.Add('SwitchToThisWindow=called')} catch { $activationNotes.Add('SwitchToThisWindow='+$_.Exception.Message) }
    Start-Sleep -Milliseconds 350
  }
  $deadline=[DateTime]::UtcNow.AddSeconds(20)
  while([OwnedImeInput]::GetForegroundWindow() -ne $window -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
  if([OwnedImeInput]::GetForegroundWindow() -ne $window){throw ('Owned test window still needs manual foreground focus; no keys were sent. '+($activationNotes -join ' | '))}
}
if($Action -eq 'layout') {
  $count=[OwnedImeInput]::GetKeyboardLayoutList(0,$null)
  $layouts=New-Object IntPtr[] $count
  [void][OwnedImeInput]::GetKeyboardLayoutList($count,$layouts)
  if(@($layouts|Where-Object{$_.ToInt64() -eq $KeyboardLayout}).Count -ne 1){throw 'Only an already-loaded layout may be selected'}
  if(![OwnedImeInput]::PostMessage($window,0x50,[IntPtr]1,[IntPtr]$KeyboardLayout)){throw 'Layout request failed'}
  Start-Sleep -Milliseconds 250
}
if($Action -eq 'type') {
  if($Text -notmatch '^[a-z]{1,80}$'){throw 'This test input only accepts short ASCII pinyin, never injected Unicode text'}
  foreach($char in $Text.ToUpperInvariant().ToCharArray()) {
    [OwnedImeInput]::Press($window,[uint16][char]$char)
    Start-Sleep -Milliseconds 100
  }
}
if($Action -eq 'key') {
  if($Key -eq 'shiftenter') {[OwnedImeInput]::PressShiftEnter($window)}
  else {
    $keys=@{space=32;enter=13;escape=27;backspace=8;shift=16}
    [OwnedImeInput]::Press($window,[uint16]$keys[$Key])
  }
}
if($Action -eq 'capture') {
  if([OwnedImeInput]::GetForegroundWindow() -ne $window){throw 'Capture requires the owned foreground window'}
  $expectedRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../target/owned-ime-review'))
  $destination=[IO.Path]::GetFullPath($OutputPath)
  if(!$destination.StartsWith($expectedRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)-or [IO.Path]::GetExtension($destination) -ne '.png'){throw 'Capture output must be an owned test PNG'}
  Add-Type -AssemblyName System.Drawing
  $rect=New-Object OwnedImeInput+RECT
  if(![OwnedImeInput]::GetWindowRect($window,[ref]$rect)){throw 'Window bounds unavailable'}
  $width=$rect.Right-$rect.Left;$height=$rect.Bottom-$rect.Top
  if($width -le 0 -or $height -le 0 -or $width -gt 5000 -or $height -gt 5000){throw 'Unexpected test window bounds'}
  $image=New-Object Drawing.Bitmap($width,$height)
  $g=[Drawing.Graphics]::FromImage($image)
  try{$g.CopyFromScreen($rect.Left,$rect.Top,0,0,$image.Size);$image.Save($destination,[Drawing.Imaging.ImageFormat]::Png)}
  finally{$g.Dispose();$image.Dispose()}
}
$count=[OwnedImeInput]::GetKeyboardLayoutList(0,$null)
$layouts=New-Object IntPtr[] $count
[void][OwnedImeInput]::GetKeyboardLayoutList($count,$layouts)
[ordered]@{action=$Action;ownedProcessId=$OwnedProcessId;hwnd=$window.ToInt64();title=[OwnedImeInput]::Title($window);
  thread=$thread;beforeLayout=$beforeLayout;afterLayout=[OwnedImeInput]::GetKeyboardLayout($thread).ToInt64();
  enabled=[OwnedImeInput]::IsWindowEnabled($window);style=[OwnedImeInput]::GetWindowLongPtr($window,-16).ToInt64().ToString('X');extendedStyle=[OwnedImeInput]::GetWindowLongPtr($window,-20).ToInt64().ToString('X');
  activePopup=[OwnedImeInput]::GetLastActivePopup($window).ToInt64();popup=$popupInfo;ownedGui=[OwnedImeInput]::Gui($thread);foregroundGui=[OwnedImeInput]::Gui(0);
  beforeForeground=$beforeForeground;foreground=[OwnedImeInput]::GetForegroundWindow().ToInt64();
  loadedLayouts=@($layouts|ForEach-Object{[ordered]@{handle=$_.ToInt64();hex=$_.ToInt64().ToString('X');languageId=$_.ToInt64()-band 65535}});
  unicodeInjection=$false;clipboardUsed=$false} | ConvertTo-Json -Depth 6
