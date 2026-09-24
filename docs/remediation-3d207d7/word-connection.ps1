# Read-only COM connection to an already running Word process. No service/add-in
# instance is created here. A second visible Word process has its own native OM.
function Connect-RunningWord([int]$WordProcessId=0) {
 if($WordProcessId -eq 0){return [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')}
 if(!('WordNativeConnection' -as [type])){ Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class WordNativeConnection {
 public delegate bool EnumProc(IntPtr window, IntPtr data);
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc proc, IntPtr data);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr data);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr window, StringBuilder text, int count);
 [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object result);
}
'@
 }
 # A completed Ribbon operation can briefly rebuild Word's _WwG surface.
 # Retry only this read-only connection; never repeat the user operation.
 $connectionDeadline=[DateTime]::UtcNow.AddSeconds(5)
 do {
 $script:nativeWordWindows=@()
 [void][WordNativeConnection]::EnumWindows({param($window,$data)
   [uint32]$ownerProcess=0;[void][WordNativeConnection]::GetWindowThreadProcessId($window,[ref]$ownerProcess)
   if($ownerProcess -eq $WordProcessId){$script:nativeWordWindows+=$window};return $true
 },[IntPtr]::Zero)
 $script:connectedWord=$null
 foreach($parent in $script:nativeWordWindows){
  [void][WordNativeConnection]::EnumChildWindows($parent,{param($window,$data)
   $class=[Text.StringBuilder]::new(64);[void][WordNativeConnection]::GetClassName($window,$class,64)
   if($class.ToString() -eq '_WwG'){
    $iid=[Guid]'00020400-0000-0000-C000-000000000046';$native=$null
    if([WordNativeConnection]::AccessibleObjectFromWindow($window,4294967280,[ref]$iid,[ref]$native) -eq 0){
      $script:connectedWord=$native.Application
      [void][Runtime.InteropServices.Marshal]::ReleaseComObject($native)
      return $false
    }
   };return $true
  },[IntPtr]::Zero)
  if($null -ne $script:connectedWord){return $script:connectedWord}
 }
 if([DateTime]::UtcNow -lt $connectionDeadline){Start-Sleep -Milliseconds 250}
 } while([DateTime]::UtcNow -lt $connectionDeadline)
 throw "No native Word document window in process $WordProcessId"
}
