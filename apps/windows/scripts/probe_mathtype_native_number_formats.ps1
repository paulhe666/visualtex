param(
    [ValidateSet("RibbonTree", "DialogTree", "ReferenceTree", "ReferenceComplete", "ReferenceFormats", "Full")]
    [string]$Mode = "RibbonTree",
    [string]$ArtifactRoot = "src-windows/artifacts/mathtype-native-number-format-probe",
    [int]$TimeoutSeconds = 30
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$artifactPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ArtifactRoot))
[IO.Directory]::CreateDirectory($artifactPath) | Out-Null

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Accessibility
$nativeProbeSource = @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Accessibility;

public static class VisualTeXMathTypeNativeProbe {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint msg, UIntPtr wParam, IntPtr lParam);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd,
        uint objectId,
        ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out object accessible);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleChildren(
        [MarshalAs(UnmanagedType.Interface)] object container,
        int childStart,
        int childCount,
        [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] children,
        out int obtained);

    [DllImport("oleacc.dll", CharSet=CharSet.Unicode)]
    private static extern uint GetRoleText(uint role, StringBuilder roleText, uint roleTextMax);

    private const uint ObjIdClient = 0xFFFFFFFC;
    private static readonly Guid IidAccessible = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");

    public static string[] DumpAccessibleTree(IntPtr hwnd) {
        IAccessible root = GetAccessibleRoot(hwnd);
        var lines = new List<string>();
        if (root == null) {
            lines.Add("MSAA_ERROR|ROOT_NOT_IACCESSIBLE");
            return lines.ToArray();
        }

        int count = 0;
        DumpAccessibleNode(root, 0, 0, lines, ref count, 0);
        return lines.ToArray();
    }

    public static int GetAccessibleStateByName(IntPtr hwnd, string exactName, int occurrence) {
        AccessibleTarget target = FindAccessibleByName(hwnd, exactName, occurrence);
        if (target == null) throw new InvalidOperationException(
            "Accessible control not found: '" + exactName + "' occurrence=" + occurrence + ".");
        return ToInt(SafeObject(delegate { return target.Accessible.get_accState(target.ChildId); }));
    }

    public static string GetAccessibleValueByName(IntPtr hwnd, string exactName, int occurrence) {
        AccessibleTarget target = FindAccessibleByName(hwnd, exactName, occurrence);
        if (target == null) throw new InvalidOperationException(
            "Accessible control not found: '" + exactName + "' occurrence=" + occurrence + ".");
        return SafeString(delegate { return target.Accessible.get_accValue(target.ChildId); });
    }

    public static string InvokeAccessibleByName(IntPtr hwnd, string exactName, int occurrence) {
        AccessibleTarget target = FindAccessibleByName(hwnd, exactName, occurrence);
        if (target == null) throw new InvalidOperationException(
            "Accessible control not found: '" + exactName + "' occurrence=" + occurrence + ".");
        string action = SafeString(delegate { return target.Accessible.get_accDefaultAction(target.ChildId); });
        target.Accessible.accDoDefaultAction(target.ChildId);
        return "name=" + exactName
            + "|occurrence=" + occurrence
            + "|role=" + ToInt(SafeObject(delegate { return target.Accessible.get_accRole(target.ChildId); }))
            + "|action=" + Escape(action);
    }

    public static string SetAccessibleValueByRoleAndCurrentValue(
        IntPtr hwnd,
        int role,
        string currentValueA,
        string currentValueB,
        int occurrence,
        string newValue) {
        IAccessible root = GetAccessibleRoot(hwnd);
        if (root == null) throw new InvalidOperationException("The dialog client is not IAccessible.");
        int seen = 0;
        int visited = 0;
        AccessibleTarget target = FindAccessibleByRoleAndCurrentValue(
            root,
            0,
            role,
            currentValueA ?? string.Empty,
            currentValueB ?? string.Empty,
            occurrence,
            ref seen,
            ref visited,
            0);
        if (target == null) throw new InvalidOperationException(
            "Accessible value control not found: role=" + role
            + ", values='" + currentValueA + "'/'" + currentValueB
            + "', occurrence=" + occurrence + ".");
        string valueBefore = SafeString(delegate { return target.Accessible.get_accValue(target.ChildId); });
        target.Accessible.set_accValue(target.ChildId, newValue ?? string.Empty);
        string valueAfter = SafeString(delegate { return target.Accessible.get_accValue(target.ChildId); });
        if (!string.Equals(valueAfter, newValue ?? string.Empty, StringComparison.Ordinal))
            throw new InvalidOperationException(
                "Accessible control rejected value '" + newValue + "'; actual='" + valueAfter + "'.");
        return "role=" + role
            + "|before=" + Escape(valueBefore)
            + "|after=" + Escape(valueAfter);
    }

    private sealed class AccessibleTarget {
        public IAccessible Accessible;
        public object ChildId;
    }

    private static IAccessible GetAccessibleRoot(IntPtr hwnd) {
        object raw;
        Guid iid = IidAccessible;
        int hr = AccessibleObjectFromWindow(hwnd, ObjIdClient, ref iid, out raw);
        if (hr < 0 || raw == null) return null;
        return raw as IAccessible;
    }

    private static AccessibleTarget FindAccessibleByName(IntPtr hwnd, string exactName, int occurrence) {
        IAccessible root = GetAccessibleRoot(hwnd);
        if (root == null) return null;
        int seen = 0;
        int visited = 0;
        return FindAccessibleByName(root, 0, exactName ?? string.Empty, occurrence, ref seen, ref visited, 0);
    }

    private static AccessibleTarget FindAccessibleByRoleAndCurrentValue(
        IAccessible accessible,
        object childId,
        int exactRole,
        string currentValueA,
        string currentValueB,
        int occurrence,
        ref int seen,
        ref int visited,
        int depth) {
        if (accessible == null || depth > 18 || visited++ >= 2000) return null;
        int role = ToInt(SafeObject(delegate { return accessible.get_accRole(childId); }));
        string value = SafeString(delegate { return accessible.get_accValue(childId); });
        if (role == exactRole
            && (string.Equals(value, currentValueA, StringComparison.Ordinal)
                || string.Equals(value, currentValueB, StringComparison.Ordinal))) {
            if (seen == occurrence)
                return new AccessibleTarget { Accessible = accessible, ChildId = childId };
            seen++;
        }

        if (!(childId is int) || (int)childId != 0) {
            IAccessible direct = null;
            try { direct = accessible.get_accChild(childId) as IAccessible; } catch { }
            return direct == null
                ? null
                : FindAccessibleByRoleAndCurrentValue(
                    direct,
                    0,
                    exactRole,
                    currentValueA,
                    currentValueB,
                    occurrence,
                    ref seen,
                    ref visited,
                    depth + 1);
        }

        int childCount = 0;
        try { childCount = accessible.accChildCount; } catch { }
        if (childCount <= 0) return null;
        object[] children = new object[childCount];
        int obtained = 0;
        int childHr = AccessibleChildren(accessible, 0, childCount, children, out obtained);
        if (childHr < 0) return null;
        for (int index = 0; index < obtained; index++) {
            object child = children[index];
            AccessibleTarget match;
            IAccessible nested = child as IAccessible;
            if (nested != null) {
                match = FindAccessibleByRoleAndCurrentValue(
                    nested,
                    0,
                    exactRole,
                    currentValueA,
                    currentValueB,
                    occurrence,
                    ref seen,
                    ref visited,
                    depth + 1);
            } else if (child != null) {
                match = FindAccessibleByRoleAndCurrentValue(
                    accessible,
                    child,
                    exactRole,
                    currentValueA,
                    currentValueB,
                    occurrence,
                    ref seen,
                    ref visited,
                    depth + 1);
            } else {
                match = null;
            }
            if (match != null) return match;
        }
        return null;
    }

    private static AccessibleTarget FindAccessibleByName(
        IAccessible accessible,
        object childId,
        string exactName,
        int occurrence,
        ref int seen,
        ref int visited,
        int depth) {
        if (accessible == null || depth > 18 || visited++ >= 2000) return null;
        string name = SafeString(delegate { return accessible.get_accName(childId); });
        if (string.Equals(name, exactName, StringComparison.Ordinal)) {
            if (seen == occurrence) {
                return new AccessibleTarget { Accessible = accessible, ChildId = childId };
            }
            seen++;
        }

        if (!(childId is int) || (int)childId != 0) {
            IAccessible direct = null;
            try { direct = accessible.get_accChild(childId) as IAccessible; } catch { }
            return direct == null
                ? null
                : FindAccessibleByName(direct, 0, exactName, occurrence, ref seen, ref visited, depth + 1);
        }

        int childCount = 0;
        try { childCount = accessible.accChildCount; } catch { }
        if (childCount <= 0) return null;
        object[] children = new object[childCount];
        int obtained = 0;
        int childHr = AccessibleChildren(accessible, 0, childCount, children, out obtained);
        if (childHr < 0) return null;

        for (int index = 0; index < obtained; index++) {
            object child = children[index];
            AccessibleTarget match;
            IAccessible nested = child as IAccessible;
            if (nested != null) {
                match = FindAccessibleByName(nested, 0, exactName, occurrence, ref seen, ref visited, depth + 1);
            } else if (child != null) {
                match = FindAccessibleByName(accessible, child, exactName, occurrence, ref seen, ref visited, depth + 1);
            } else {
                match = null;
            }
            if (match != null) return match;
        }
        return null;
    }

    private static void DumpAccessibleNode(
        IAccessible accessible,
        object childId,
        int depth,
        List<string> lines,
        ref int count,
        int siblingIndex) {
        if (accessible == null || depth > 18 || count >= 2000) return;
        count++;

        string name = SafeString(delegate { return accessible.get_accName(childId); });
        string value = SafeString(delegate { return accessible.get_accValue(childId); });
        string description = SafeString(delegate { return accessible.get_accDescription(childId); });
        string shortcut = SafeString(delegate { return accessible.get_accKeyboardShortcut(childId); });
        string defaultAction = SafeString(delegate { return accessible.get_accDefaultAction(childId); });
        object roleValue = SafeObject(delegate { return accessible.get_accRole(childId); });
        object stateValue = SafeObject(delegate { return accessible.get_accState(childId); });
        int left = 0, top = 0, width = 0, height = 0;
        try { accessible.accLocation(out left, out top, out width, out height, childId); } catch { }

        int roleNumber = ToInt(roleValue);
        int stateNumber = ToInt(stateValue);
        lines.Add(
            new string(' ', depth * 2)
            + "ACC|depth=" + depth
            + "|index=" + siblingIndex
            + "|child=" + Escape(childId)
            + "|role=" + roleNumber + ":" + Escape(RoleText(roleNumber))
            + "|state=0x" + stateNumber.ToString("X")
            + "|name=" + Escape(name)
            + "|value=" + Escape(value)
            + "|description=" + Escape(description)
            + "|shortcut=" + Escape(shortcut)
            + "|action=" + Escape(defaultAction)
            + "|rect=" + left + "," + top + "," + width + "," + height);

        IAccessible childAccessible = null;
        if (!(childId is int) || (int)childId != 0) {
            try { childAccessible = accessible.get_accChild(childId) as IAccessible; } catch { }
            if (childAccessible != null) {
                DumpAccessibleNode(childAccessible, 0, depth + 1, lines, ref count, 0);
            }
            return;
        }

        int childCount = 0;
        try { childCount = accessible.accChildCount; } catch { }
        if (childCount <= 0) return;
        object[] children = new object[childCount];
        int obtained = 0;
        int childHr = AccessibleChildren(accessible, 0, childCount, children, out obtained);
        if (childHr < 0) {
            lines.Add(new string(' ', (depth + 1) * 2) + "MSAA_CHILDREN_ERROR|HRESULT=0x" + childHr.ToString("X8"));
            return;
        }

        for (int index = 0; index < obtained; index++) {
            object child = children[index];
            var nested = child as IAccessible;
            if (nested != null) {
                DumpAccessibleNode(nested, 0, depth + 1, lines, ref count, index);
            } else if (child != null) {
                DumpAccessibleNode(accessible, child, depth + 1, lines, ref count, index);
            }
        }
    }

    private delegate string StringGetter();
    private delegate object ObjectGetter();

    private static string SafeString(StringGetter getter) {
        try { return getter() ?? string.Empty; } catch { return string.Empty; }
    }

    private static object SafeObject(ObjectGetter getter) {
        try { return getter(); } catch { return null; }
    }

    private static int ToInt(object value) {
        if (value == null) return 0;
        try { return Convert.ToInt32(value); } catch { return 0; }
    }

    private static string RoleText(int role) {
        if (role <= 0) return string.Empty;
        var builder = new StringBuilder(128);
        uint copied = GetRoleText((uint)role, builder, (uint)builder.Capacity);
        return copied > 0 ? builder.ToString() : string.Empty;
    }

    private static string Escape(object value) {
        if (value == null) return string.Empty;
        return value.ToString()
            .Replace("\\", "\\\\")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("|", "\\|");
    }
}
"@
Add-Type -TypeDefinition $nativeProbeSource -ReferencedAssemblies ([Accessibility.IAccessible].Assembly.Location)

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Get-TopLevelWindows {
    $result = New-Object Collections.Generic.List[object]
    [VisualTeXMathTypeNativeProbe]::EnumWindows({
        param([IntPtr]$hwnd, [IntPtr]$unused)
        if (-not [VisualTeXMathTypeNativeProbe]::IsWindowVisible($hwnd)) { return $true }
        [uint32]$pidValue = 0
        [void][VisualTeXMathTypeNativeProbe]::GetWindowThreadProcessId($hwnd, [ref]$pidValue)
        if ($pidValue -eq 0) { return $true }
        try { $process = Get-Process -Id $pidValue -ErrorAction Stop }
        catch { return $true }
        if ($process.ProcessName -notin @("WINWORD", "MathType", "MathTypeLib")) { return $true }
        $titleLength = [VisualTeXMathTypeNativeProbe]::GetWindowTextLength($hwnd)
        $titleBuilder = New-Object Text.StringBuilder ([Math]::Max(2, $titleLength + 1))
        [void][VisualTeXMathTypeNativeProbe]::GetWindowText($hwnd, $titleBuilder, $titleBuilder.Capacity)
        $classBuilder = New-Object Text.StringBuilder 256
        [void][VisualTeXMathTypeNativeProbe]::GetClassName($hwnd, $classBuilder, $classBuilder.Capacity)
        $result.Add([pscustomobject]@{
            Handle = [Int64]$hwnd
            ProcessId = [int]$pidValue
            ProcessName = $process.ProcessName
            Title = $titleBuilder.ToString()
            ClassName = $classBuilder.ToString()
        })
        return $true
    }, [IntPtr]::Zero) | Out-Null
    return $result.ToArray()
}

function Write-WindowInventory([string]$fileName) {
    $path = Join-Path $artifactPath $fileName
    Get-TopLevelWindows |
        Sort-Object ProcessName, ProcessId, Handle |
        ConvertTo-Json -Depth 4 |
        Set-Content -LiteralPath $path -Encoding UTF8
    Write-Output "WINDOW_INVENTORY=$path"
}

function Get-NativeWindowDescription([IntPtr]$hwnd) {
    $titleLength = [VisualTeXMathTypeNativeProbe]::GetWindowTextLength($hwnd)
    $titleBuilder = New-Object Text.StringBuilder ([Math]::Max(2, $titleLength + 1))
    [void][VisualTeXMathTypeNativeProbe]::GetWindowText($hwnd, $titleBuilder, $titleBuilder.Capacity)
    $classBuilder = New-Object Text.StringBuilder 256
    [void][VisualTeXMathTypeNativeProbe]::GetClassName($hwnd, $classBuilder, $classBuilder.Capacity)
    $rect = New-Object VisualTeXMathTypeNativeProbe+RECT
    [void][VisualTeXMathTypeNativeProbe]::GetWindowRect($hwnd, [ref]$rect)
    return [pscustomobject]@{
        Handle = [Int64]$hwnd
        Parent = [Int64][VisualTeXMathTypeNativeProbe]::GetParent($hwnd)
        ControlId = [VisualTeXMathTypeNativeProbe]::GetDlgCtrlID($hwnd)
        ClassName = $classBuilder.ToString()
        Title = $titleBuilder.ToString()
        Visible = [VisualTeXMathTypeNativeProbe]::IsWindowVisible($hwnd)
        Enabled = [VisualTeXMathTypeNativeProbe]::IsWindowEnabled($hwnd)
        Rectangle = "$($rect.Left),$($rect.Top),$($rect.Right),$($rect.Bottom)"
    }
}

function Dump-NativeChildWindowTree([IntPtr]$rootHwnd, [string]$path) {
    $nodes = New-Object Collections.Generic.List[object]
    $nodes.Add((Get-NativeWindowDescription $rootHwnd))
    [VisualTeXMathTypeNativeProbe]::EnumChildWindows($rootHwnd, {
        param([IntPtr]$childHwnd, [IntPtr]$unused)
        $nodes.Add((Get-NativeWindowDescription $childHwnd))
        return $true
    }, [IntPtr]::Zero) | Out-Null
    $nodes | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
    Write-Output "NATIVE_CHILD_TREE=$path"
}

function Dump-AccessibleTree([IntPtr]$rootHwnd, [string]$path) {
    $lines = [VisualTeXMathTypeNativeProbe]::DumpAccessibleTree($rootHwnd)
    [IO.File]::WriteAllLines($path, $lines, [Text.UTF8Encoding]::new($false))
    Write-Output "MSAA_TREE=$path"
}

function Write-Utf8Text([string]$path, [AllowNull()][string]$text) {
    [IO.File]::WriteAllText(
        $path,
        $(if ($null -eq $text) { "" } else { $text }),
        [Text.UTF8Encoding]::new($false))
}

function ConvertTo-WordDebugText([AllowNull()][string]$text) {
    if ($null -eq $text) { return "" }
    $builder = New-Object Text.StringBuilder
    foreach ($character in $text.ToCharArray()) {
        $codePoint = [int][char]$character
        $token = switch ($codePoint) {
            1 { "<OLE>" }
            7 { "<CELL>" }
            9 { "<TAB>" }
            10 { "<LF>" }
            11 { "<LINE_BREAK>" }
            12 { "<PAGE_BREAK>" }
            13 { "<PARA>" }
            19 { "<FIELD_BEGIN>" }
            20 { "<FIELD_SEPARATOR>" }
            21 { "<FIELD_END>" }
            default { $null }
        }
        if ($null -ne $token) { [void]$builder.Append($token) }
        elseif ([char]::IsControl($character)) {
            [void]$builder.Append(("<U+{0:X4}>" -f $codePoint))
        } else {
            [void]$builder.Append($character)
        }
    }
    return $builder.ToString()
}

function Get-WordCharacterSnapshot([object]$document, [int]$position) {
    $content = $null
    $range = $null
    try {
        $content = $document.Content
        if ($position -lt $content.Start -or $position -ge $content.End) {
            return [pscustomobject]@{ Position = $position; CodePoint = $null; Text = "<OUT_OF_RANGE>" }
        }
        $range = $document.Range($position, $position + 1)
        $text = [string]$range.Text
        $codePoint = if ($text.Length -gt 0) { [int][char]$text[0] } else { $null }
        return [pscustomobject]@{
            Position = $position
            CodePoint = $codePoint
            Text = ConvertTo-WordDebugText $text
        }
    } finally {
        Release-ComObject $range
        Release-ComObject $content
    }
}

function Normalize-WordOpenXml([AllowNull()][string]$xml) {
    if ([string]::IsNullOrWhiteSpace($xml)) { return "" }
    $normalized = [regex]::Replace(
        $xml,
        '<pkg:binaryData>.*?</pkg:binaryData>',
        '<pkg:binaryData>[BINARY OMITTED]</pkg:binaryData>',
        [Text.RegularExpressions.RegexOptions]::Singleline)
    $normalized = [regex]::Replace(
        $normalized,
        '\s+(?:w:rsid[A-Za-z0-9]*|w14:paraId|w14:textId|w15:paraId)="[^"]*"',
        '')
    return $normalized
}

function Get-ParagraphTabStopsSnapshot([object]$paragraphFormat) {
    $tabStops = $null
    $tabStop = $null
    $result = New-Object Collections.Generic.List[object]
    try {
        $tabStops = $paragraphFormat.TabStops
        for ($index = 1; $index -le $tabStops.Count; $index++) {
            Release-ComObject $tabStop
            $tabStop = $tabStops.Item($index)
            $result.Add([pscustomobject][ordered]@{
                Index = $index
                Position = [double]$tabStop.Position
                Alignment = [int]$tabStop.Alignment
                Leader = [int]$tabStop.Leader
            })
        }
        return $result.ToArray()
    } finally {
        Release-ComObject $tabStop
        Release-ComObject $tabStops
    }
}

function Get-RangeStyleName([object]$range) {
    $styleValue = $null
    try {
        $styleValue = $range.Style
        if ($styleValue -is [string]) { return [string]$styleValue }
        try { return [string]$styleValue.NameLocal } catch { return [string]$styleValue }
    } catch { return "<unreadable>" }
    finally { Release-ComObject $styleValue }
}

function Write-RangeWordOpenXml([object]$range, [string]$path) {
    try {
        Write-Utf8Text $path ([string]$range.WordOpenXML)
        return $true
    } catch {
        Write-Utf8Text ($path + ".error.txt") $_.Exception.ToString()
        return $false
    }
}

function Capture-WordStructureSnapshot(
    [object]$document,
    [string]$stageName,
    [string]$documentPath
) {
    $stagePath = Join-Path $artifactPath $stageName
    [IO.Directory]::CreateDirectory($stagePath) | Out-Null
    [IO.Directory]::CreateDirectory((Join-Path $stagePath "field-ranges")) | Out-Null
    [IO.Directory]::CreateDirectory((Join-Path $stagePath "paragraph-ranges")) | Out-Null
    [IO.Directory]::CreateDirectory((Join-Path $stagePath "bookmark-ranges")) | Out-Null

    if ([string]::IsNullOrWhiteSpace([string]$document.Path)) {
        $document.SaveAs2($documentPath, 12)
    } else {
        $document.Save()
    }
    $stageDocumentPath = Join-Path $stagePath ($stageName + ".docx")
    [IO.File]::Copy([string]$document.FullName, $stageDocumentPath, $true)

    $content = $null
    $rawFlatOpc = ""
    try {
        $content = $document.Content
        $rawFlatOpc = [string]$content.WordOpenXML
        Write-Utf8Text (Join-Path $stagePath "content-flat-opc.raw.xml") $rawFlatOpc
        $normalizedFlatOpc = Normalize-WordOpenXml $rawFlatOpc
        Write-Utf8Text (Join-Path $stagePath "content-flat-opc.normalized.xml") $normalizedFlatOpc

        try {
            $packageDocument = New-Object Xml.XmlDocument
            $packageDocument.PreserveWhitespace = $true
            $packageDocument.LoadXml($normalizedFlatOpc)
            $namespaceManager = New-Object Xml.XmlNamespaceManager($packageDocument.NameTable)
            $namespaceManager.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/xmlPackage")
            $documentXmlNode = $packageDocument.SelectSingleNode(
                "/pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/*",
                $namespaceManager)
            if ($null -ne $documentXmlNode) {
                Write-Utf8Text (Join-Path $stagePath "word-document.normalized.xml") $documentXmlNode.OuterXml
            }
        } catch {
            Write-Utf8Text (Join-Path $stagePath "word-document-extraction.error.txt") $_.Exception.ToString()
        }

        Write-Utf8Text (Join-Path $stagePath "document-text.txt") (ConvertTo-WordDebugText ([string]$content.Text))
    } finally { Release-ComObject $content }

    $fieldEntries = New-Object Collections.Generic.List[object]
    $fields = $null
    $field = $null
    $code = $null
    $result = $null
    $codeFields = $null
    $fullSpan = $null
    $fieldParagraphs = $null
    $fieldParagraph = $null
    $fieldParagraphRange = $null
    try {
        $fields = $document.Fields
        for ($index = 1; $index -le $fields.Count; $index++) {
            Release-ComObject $fieldParagraphRange; $fieldParagraphRange = $null
            Release-ComObject $fieldParagraph; $fieldParagraph = $null
            Release-ComObject $fieldParagraphs; $fieldParagraphs = $null
            Release-ComObject $fullSpan; $fullSpan = $null
            Release-ComObject $codeFields; $codeFields = $null
            Release-ComObject $result; $result = $null
            Release-ComObject $code; $code = $null
            Release-ComObject $field; $field = $fields.Item($index)

            $code = $field.Code
            $result = $field.Result
            $codeFields = $code.Fields
            $fullStart = [Math]::Max(0, [int]$code.Start - 1)
            $fullEnd = [Math]::Min([int]$document.Content.End, [int]$result.End + 1)
            $fullSpan = $document.Range($fullStart, $fullEnd)
            $fieldXmlName = "field-{0:D3}-{1}-{2}.xml" -f $index, $fullStart, $fullEnd
            [void](Write-RangeWordOpenXml $fullSpan (Join-Path (Join-Path $stagePath "field-ranges") $fieldXmlName))

            $paragraphStart = $null
            $paragraphEnd = $null
            $paragraphStyle = ""
            try {
                $fieldParagraphs = $fullSpan.Paragraphs
                if ($fieldParagraphs.Count -gt 0) {
                    $fieldParagraph = $fieldParagraphs.Item(1)
                    $fieldParagraphRange = $fieldParagraph.Range
                    $paragraphStart = [int]$fieldParagraphRange.Start
                    $paragraphEnd = [int]$fieldParagraphRange.End
                    $paragraphStyle = Get-RangeStyleName $fieldParagraphRange
                }
            } catch { }

            $fieldEntries.Add([pscustomobject][ordered]@{
                Index = $index
                Type = [int]$field.Type
                Locked = [bool]$field.Locked
                ShowCodes = [bool]$field.ShowCodes
                Begin = Get-WordCharacterSnapshot $document ([int]$code.Start - 1)
                CodeStart = [int]$code.Start
                CodeEnd = [int]$code.End
                CodeText = ConvertTo-WordDebugText ([string]$code.Text)
                Separator = Get-WordCharacterSnapshot $document ([int]$code.End)
                ResultStart = [int]$result.Start
                ResultEnd = [int]$result.End
                ResultText = ConvertTo-WordDebugText ([string]$result.Text)
                End = Get-WordCharacterSnapshot $document ([int]$result.End)
                FullStart = $fullStart
                FullEnd = $fullEnd
                FullText = ConvertTo-WordDebugText ([string]$fullSpan.Text)
                NestedCodeFieldCount = [int]$codeFields.Count
                ParentFieldIndex = $null
                ParagraphStart = $paragraphStart
                ParagraphEnd = $paragraphEnd
                ParagraphStyle = $paragraphStyle
                WordOpenXmlFile = "field-ranges/$fieldXmlName"
            })
        }
    } finally {
        Release-ComObject $fieldParagraphRange
        Release-ComObject $fieldParagraph
        Release-ComObject $fieldParagraphs
        Release-ComObject $fullSpan
        Release-ComObject $codeFields
        Release-ComObject $result
        Release-ComObject $code
        Release-ComObject $field
        Release-ComObject $fields
    }

    foreach ($child in $fieldEntries) {
        $bestParent = $null
        $bestSpan = [int]::MaxValue
        foreach ($candidate in $fieldEntries) {
            if ($candidate.Index -eq $child.Index) { continue }
            if ($candidate.FullStart -gt $child.FullStart -or $candidate.FullEnd -lt $child.FullEnd) { continue }
            $candidateSpan = [int]$candidate.FullEnd - [int]$candidate.FullStart
            $childSpan = [int]$child.FullEnd - [int]$child.FullStart
            if ($candidateSpan -le $childSpan -or $candidateSpan -ge $bestSpan) { continue }
            $bestParent = $candidate.Index
            $bestSpan = $candidateSpan
        }
        $child.ParentFieldIndex = $bestParent
    }
    Write-Utf8Text (Join-Path $stagePath "fields.json") ($fieldEntries | ConvertTo-Json -Depth 8)

    $paragraphEntries = New-Object Collections.Generic.List[object]
    $paragraphs = $null
    $paragraph = $null
    $paragraphRange = $null
    $paragraphFormat = $null
    try {
        $paragraphs = $document.Paragraphs
        for ($index = 1; $index -le $paragraphs.Count; $index++) {
            Release-ComObject $paragraphFormat; $paragraphFormat = $null
            Release-ComObject $paragraphRange; $paragraphRange = $null
            Release-ComObject $paragraph; $paragraph = $paragraphs.Item($index)
            $paragraphRange = $paragraph.Range
            $paragraphFormat = $paragraph.Format
            $paragraphXmlName = "paragraph-{0:D3}-{1}-{2}.xml" -f $index, $paragraphRange.Start, $paragraphRange.End
            [void](Write-RangeWordOpenXml $paragraphRange (Join-Path (Join-Path $stagePath "paragraph-ranges") $paragraphXmlName))
            $paragraphEntries.Add([pscustomobject][ordered]@{
                Index = $index
                Start = [int]$paragraphRange.Start
                End = [int]$paragraphRange.End
                Text = ConvertTo-WordDebugText ([string]$paragraphRange.Text)
                Style = Get-RangeStyleName $paragraphRange
                Alignment = [int]$paragraphFormat.Alignment
                LeftIndent = [double]$paragraphFormat.LeftIndent
                RightIndent = [double]$paragraphFormat.RightIndent
                FirstLineIndent = [double]$paragraphFormat.FirstLineIndent
                SpaceBefore = [double]$paragraphFormat.SpaceBefore
                SpaceAfter = [double]$paragraphFormat.SpaceAfter
                LineSpacing = [double]$paragraphFormat.LineSpacing
                LineSpacingRule = [int]$paragraphFormat.LineSpacingRule
                KeepTogether = [int]$paragraphFormat.KeepTogether
                KeepWithNext = [int]$paragraphFormat.KeepWithNext
                PageBreakBefore = [int]$paragraphFormat.PageBreakBefore
                OutlineLevel = [int]$paragraphFormat.OutlineLevel
                TabStops = @(Get-ParagraphTabStopsSnapshot $paragraphFormat)
                FieldCount = [int]$paragraphRange.Fields.Count
                InlineShapeCount = [int]$paragraphRange.InlineShapes.Count
                WordOpenXmlFile = "paragraph-ranges/$paragraphXmlName"
            })
        }
    } finally {
        Release-ComObject $paragraphFormat
        Release-ComObject $paragraphRange
        Release-ComObject $paragraph
        Release-ComObject $paragraphs
    }
    Write-Utf8Text (Join-Path $stagePath "paragraphs.json") ($paragraphEntries | ConvertTo-Json -Depth 8)

    $shapeEntries = New-Object Collections.Generic.List[object]
    $inlineShapes = $null
    $shape = $null
    $shapeRange = $null
    $oleFormat = $null
    try {
        $inlineShapes = $document.InlineShapes
        for ($index = 1; $index -le $inlineShapes.Count; $index++) {
            Release-ComObject $oleFormat; $oleFormat = $null
            Release-ComObject $shapeRange; $shapeRange = $null
            Release-ComObject $shape; $shape = $inlineShapes.Item($index)
            $shapeRange = $shape.Range
            $progId = ""
            try {
                $oleFormat = $shape.OLEFormat
                $progId = [string]$oleFormat.ProgID
            } catch { }
            $shapeEntries.Add([pscustomobject][ordered]@{
                Index = $index
                Type = [int]$shape.Type
                Start = [int]$shapeRange.Start
                End = [int]$shapeRange.End
                RangeText = ConvertTo-WordDebugText ([string]$shapeRange.Text)
                ProgId = $progId
                Width = [double]$shape.Width
                Height = [double]$shape.Height
                ScaleWidth = $(try { [double]$shape.ScaleWidth } catch { $null })
                ScaleHeight = $(try { [double]$shape.ScaleHeight } catch { $null })
                AlternativeText = $(try { [string]$shape.AlternativeText } catch { "" })
                Title = $(try { [string]$shape.Title } catch { "" })
            })
        }
    } finally {
        Release-ComObject $oleFormat
        Release-ComObject $shapeRange
        Release-ComObject $shape
        Release-ComObject $inlineShapes
    }
    Write-Utf8Text (Join-Path $stagePath "inline-shapes.json") ($shapeEntries | ConvertTo-Json -Depth 6)

    $bookmarkEntries = New-Object Collections.Generic.List[object]
    $bookmarks = $null
    $bookmark = $null
    $bookmarkRange = $null
    $previousShowHidden = $false
    try {
        $bookmarks = $document.Bookmarks
        try {
            $previousShowHidden = [bool]$bookmarks.ShowHidden
            $bookmarks.ShowHidden = $true
        } catch { }
        for ($index = 1; $index -le $bookmarks.Count; $index++) {
            Release-ComObject $bookmarkRange; $bookmarkRange = $null
            Release-ComObject $bookmark; $bookmark = $bookmarks.Item($index)
            $bookmarkRange = $bookmark.Range
            $bookmarkXmlName = "bookmark-{0:D3}-{1}.xml" -f $index, ([regex]::Replace([string]$bookmark.Name, '[^A-Za-z0-9_.-]', '_'))
            [void](Write-RangeWordOpenXml $bookmarkRange (Join-Path (Join-Path $stagePath "bookmark-ranges") $bookmarkXmlName))
            $bookmarkEntries.Add([pscustomobject][ordered]@{
                Index = $index
                Name = [string]$bookmark.Name
                Start = [int]$bookmarkRange.Start
                End = [int]$bookmarkRange.End
                Text = ConvertTo-WordDebugText ([string]$bookmarkRange.Text)
                Empty = [bool]$bookmark.Empty
                Column = [bool]$bookmark.Column
                WordOpenXmlFile = "bookmark-ranges/$bookmarkXmlName"
            })
        }
    } finally {
        if ($null -ne $bookmarks) { try { $bookmarks.ShowHidden = $previousShowHidden } catch { } }
        Release-ComObject $bookmarkRange
        Release-ComObject $bookmark
        Release-ComObject $bookmarks
    }
    Write-Utf8Text (Join-Path $stagePath "bookmarks.json") ($bookmarkEntries | ConvertTo-Json -Depth 6)

    $customPropertyEntries = New-Object Collections.Generic.List[object]
    $customProperties = $null
    $customProperty = $null
    try {
        $customProperties = $document.CustomDocumentProperties
        for ($index = 1; $index -le $customProperties.Count; $index++) {
            Release-ComObject $customProperty; $customProperty = $customProperties.Item($index)
            $customPropertyEntries.Add([pscustomobject][ordered]@{
                Index = $index
                Name = $(try { [string]$customProperty.Name } catch { "" })
                Type = $(try { [int]$customProperty.Type } catch { $null })
                Value = $(try { [string]$customProperty.Value } catch { "" })
                LinkToContent = $(try { [bool]$customProperty.LinkToContent } catch { $null })
                LinkSource = $(try { [string]$customProperty.LinkSource } catch { "" })
            })
        }
    } catch {
        $customPropertyEntries.Add([pscustomobject]@{ Error = $_.Exception.Message })
    } finally {
        Release-ComObject $customProperty
        Release-ComObject $customProperties
    }
    Write-Utf8Text (Join-Path $stagePath "custom-document-properties.json") ($customPropertyEntries | ConvertTo-Json -Depth 6)

    $variableEntries = New-Object Collections.Generic.List[object]
    $variables = $null
    $variable = $null
    try {
        $variables = $document.Variables
        for ($index = 1; $index -le $variables.Count; $index++) {
            Release-ComObject $variable; $variable = $variables.Item($index)
            $variableEntries.Add([pscustomobject][ordered]@{
                Index = $index
                Name = [string]$variable.Name
                Value = [string]$variable.Value
            })
        }
    } finally {
        Release-ComObject $variable
        Release-ComObject $variables
    }
    Write-Utf8Text (Join-Path $stagePath "document-variables.json") ($variableEntries | ConvertTo-Json -Depth 4)

    $styleEntries = New-Object Collections.Generic.List[object]
    $styles = $null
    $style = $null
    $styleParagraphFormat = $null
    $styleFont = $null
    try {
        $styles = $document.Styles
        for ($index = 1; $index -le $styles.Count; $index++) {
            Release-ComObject $styleFont; $styleFont = $null
            Release-ComObject $styleParagraphFormat; $styleParagraphFormat = $null
            Release-ComObject $style; $style = $styles.Item($index)
            $name = $(try { [string]$style.NameLocal } catch { "" })
            if (-not $name.StartsWith("MT", [StringComparison]::OrdinalIgnoreCase)) { continue }
            try {
                $styleParagraphFormat = $style.ParagraphFormat
                $styleFont = $style.Font
                $styleEntries.Add([pscustomobject][ordered]@{
                    Index = $index
                    Name = $name
                    Type = [int]$style.Type
                    Alignment = [int]$styleParagraphFormat.Alignment
                    LeftIndent = [double]$styleParagraphFormat.LeftIndent
                    RightIndent = [double]$styleParagraphFormat.RightIndent
                    FirstLineIndent = [double]$styleParagraphFormat.FirstLineIndent
                    SpaceBefore = [double]$styleParagraphFormat.SpaceBefore
                    SpaceAfter = [double]$styleParagraphFormat.SpaceAfter
                    LineSpacingRule = [int]$styleParagraphFormat.LineSpacingRule
                    TabStops = @(Get-ParagraphTabStopsSnapshot $styleParagraphFormat)
                    FontName = $(try { [string]$styleFont.Name } catch { "" })
                    FontSize = $(try { [double]$styleFont.Size } catch { $null })
                    FontHidden = $(try { [int]$styleFont.Hidden } catch { $null })
                    FontColor = $(try { [int]$styleFont.Color } catch { $null })
                })
            } catch {
                $styleEntries.Add([pscustomobject][ordered]@{
                    Index = $index
                    Name = $name
                    Error = $_.Exception.Message
                })
            }
        }
    } finally {
        Release-ComObject $styleFont
        Release-ComObject $styleParagraphFormat
        Release-ComObject $style
        Release-ComObject $styles
    }
    Write-Utf8Text (Join-Path $stagePath "mathtype-styles.json") ($styleEntries | ConvertTo-Json -Depth 8)

    $pageSetup = $null
    try {
        $pageSetup = $document.PageSetup
        $summary = [pscustomobject][ordered]@{
            Stage = $stageName
            FullName = [string]$document.FullName
            SavedCopy = $stageDocumentPath
            DocumentStart = [int]$document.Content.Start
            DocumentEnd = [int]$document.Content.End
            ParagraphCount = [int]$document.Paragraphs.Count
            FieldCount = [int]$document.Fields.Count
            InlineShapeCount = [int]$document.InlineShapes.Count
            BookmarkCount = [int]$document.Bookmarks.Count
            PageWidth = [double]$pageSetup.PageWidth
            LeftMargin = [double]$pageSetup.LeftMargin
            RightMargin = [double]$pageSetup.RightMargin
            UsableWidth = [double]($pageSetup.PageWidth - $pageSetup.LeftMargin - $pageSetup.RightMargin)
            CapturedAt = [DateTimeOffset]::Now.ToString("O")
        }
        Write-Utf8Text (Join-Path $stagePath "summary.json") ($summary | ConvertTo-Json -Depth 5)
    } finally { Release-ComObject $pageSetup }

    Write-Output ("WORD_STRUCTURE_SNAPSHOT|stage="+$stageName+"|path="+$stagePath+"|fields="+$fieldEntries.Count+"|paragraphs="+$paragraphEntries.Count+"|shapes="+$shapeEntries.Count+"|bookmarks="+$bookmarkEntries.Count)
}

function Find-ElementExact(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$name,
    [System.Windows.Automation.ControlType]$controlType
) {
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $name)
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        $controlType)
    return $root.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.AndCondition($nameCondition, $typeCondition)))
}

function Wait-ElementExact(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$name,
    [System.Windows.Automation.ControlType]$controlType,
    [int]$seconds = 15
) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        $element = Find-ElementExact $root $name $controlType
        if ($null -ne $element) { return $element }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 120
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out locating UI element '$name' ($($controlType.ProgrammaticName))."
}

function Invoke-Element(
    [System.Windows.Automation.AutomationElement]$element,
    [string]$description
) {
    if ($null -eq $element) { throw "UI element not found: $description" }
    $patternObject = $null
    if (-not $element.TryGetCurrentPattern(
        [System.Windows.Automation.InvokePattern]::Pattern,
        [ref]$patternObject)) {
        throw "UI element has no InvokePattern: $description"
    }
    ([System.Windows.Automation.InvokePattern]$patternObject).Invoke()
}

function Wait-NewTopLevelWindow(
    [Int64[]]$existingHandles,
    [string[]]$processNames,
    [int]$seconds
) {
    $existing = @{}
    foreach ($handle in $existingHandles) { $existing[$handle] = $true }
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        $candidate = Get-TopLevelWindows |
            Where-Object {
                $processNames -contains $_.ProcessName -and
                -not $existing.ContainsKey([Int64]$_.Handle)
            } |
            Sort-Object @{ Expression = { [string]::IsNullOrWhiteSpace($_.Title) } }, ProcessName, ProcessId |
            Select-Object -First 1
        if ($null -ne $candidate) { return $candidate }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 120
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out waiting for a new top-level window owned by: $($processNames -join ', ')."
}

function Wait-DocumentInlineShapeCount([object]$document, [int]$expected, [int]$seconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        try {
            if ($document.InlineShapes.Count -ge $expected) { return }
        } catch { }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out waiting for Word to contain $expected inline MathType object(s)."
}

function Invoke-NativeMathTypeNumberedEquation(
    [object]$document,
    [IntPtr]$wordHwnd,
    [ValidateSet("left", "right")][string]$side,
    [string]$latex,
    [int]$expectedShapeCount
) {
    $wordRoot = Select-RibbonTab $wordHwnd "MathType"
    $buttonName = if ($side -eq "left") { "左编号   " } else { "右编号" }
    $button = Wait-ElementExact $wordRoot $buttonName ([System.Windows.Automation.ControlType]::Button) $TimeoutSeconds
    $beforeEditorHandles = @(
        Get-TopLevelWindows |
            ForEach-Object { [Int64]$_.Handle })

    Invoke-Element $button "MathType native $side-numbered display equation"
    Start-Sleep -Milliseconds 900
    Write-WindowInventory ("windows-after-native-" + $side + "-click.json")
    try {
        $editor = Wait-NewTopLevelWindow $beforeEditorHandles @("MathType", "MathTypeLib", "WINWORD") $TimeoutSeconds
    } catch {
        $desktop = [System.Windows.Automation.AutomationElement]::RootElement
        Dump-AutomationTree $desktop (Join-Path $artifactPath ("desktop-after-native-" + $side + "-timeout-uia-tree.txt"))
        throw
    }
    Write-Output ("MATHTYPE_EDITOR|side=$side|pid="+$editor.ProcessId+"|hwnd="+$editor.Handle+"|title="+$editor.Title+"|class="+$editor.ClassName)
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$editor.Handle)
    Dump-AutomationTree $editorRoot (Join-Path $artifactPath ("native-" + $side + "-first-window-uia-tree.txt"))
    Dump-NativeChildWindowTree ([IntPtr]$editor.Handle) (Join-Path $artifactPath ("native-" + $side + "-first-window-native-tree.json"))
    Dump-AccessibleTree ([IntPtr]$editor.Handle) (Join-Path $artifactPath ("native-" + $side + "-first-window-msaa-tree.txt"))
    [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow([IntPtr]$editor.Handle)
    Start-Sleep -Milliseconds 500
    if ($editor.ClassName -eq "ThunderDFrame" -and $editor.Title -eq "插入公式编号") {
        $beforeContinueHandles = @(
            Get-TopLevelWindows |
                ForEach-Object { [Int64]$_.Handle })
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Start-Sleep -Milliseconds 900
        Write-WindowInventory ("windows-after-native-" + $side + "-number-dialog-enter.json")
        try {
            $editor = Wait-NewTopLevelWindow $beforeContinueHandles @("MathType", "MathTypeLib", "WINWORD") $TimeoutSeconds
        } catch {
            $desktop = [System.Windows.Automation.AutomationElement]::RootElement
            Dump-AutomationTree $desktop (Join-Path $artifactPath ("desktop-after-native-" + $side + "-number-dialog-enter-timeout-uia-tree.txt"))
            throw
        }
        Write-Output ("MATHTYPE_EDITOR_AFTER_NUMBER_DIALOG|side=$side|pid="+$editor.ProcessId+"|hwnd="+$editor.Handle+"|title="+$editor.Title+"|class="+$editor.ClassName)
        $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$editor.Handle)
        Dump-AutomationTree $editorRoot (Join-Path $artifactPath ("native-" + $side + "-second-window-uia-tree.txt"))
        [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow([IntPtr]$editor.Handle)
        Start-Sleep -Milliseconds 500
    } elseif ($editor.ClassName -eq "ThunderDFrame") {
        [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
        throw "MathType native $side-numbered insertion opened an unexpected VBA dialog '$($editor.Title)'."
    }

    [System.Windows.Forms.SendKeys]::SendWait("^a")
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.SendKeys]::SendWait("{DELETE}")
    Start-Sleep -Milliseconds 180
    [System.Windows.Forms.Clipboard]::SetText($latex, [System.Windows.Forms.TextDataFormat]::UnicodeText)
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    Start-Sleep -Milliseconds 900
    [System.Windows.Forms.SendKeys]::SendWait("%{F4}")

    Wait-DocumentInlineShapeCount $document $expectedShapeCount $TimeoutSeconds
    Start-Sleep -Milliseconds 700
    Write-Output ("NATIVE_MATHTYPE_INSERT|side=$side|inlineShapes="+$document.InlineShapes.Count+"|fields="+$document.Fields.Count)
}

function Get-FirstNativeMathTypePlaceRefField([object]$document) {
    $fields = $null
    $field = $null
    $code = $null
    try {
        $fields = $document.Fields
        for ($index = 1; $index -le $fields.Count; $index++) {
            Release-ComObject $code; $code = $null
            Release-ComObject $field; $field = $fields.Item($index)
            $code = $field.Code
            if (([string]$code.Text) -match '^\s*MACROBUTTON\s+MTPlaceRef\b') {
                $result = $field
                $field = $null
                return $result
            }
        }
        return $null
    } finally {
        Release-ComObject $code
        Release-ComObject $field
        Release-ComObject $fields
    }
}

function Wait-NativeMathTypeReferenceCompleted([object]$document, [int]$seconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        $bookmarks = $null
        try {
            $bookmarks = $document.Bookmarks
            $pending = $bookmarks.Exists("MTReference")
            $hasTarget = $false
            for ($index = 1; $index -le $bookmarks.Count; $index++) {
                $bookmark = $null
                try {
                    $bookmark = $bookmarks.Item($index)
                    if (([string]$bookmark.Name) -like "ZEqnNum*") {
                        $hasTarget = $true
                        break
                    }
                } finally { Release-ComObject $bookmark }
            }
            if (-not $pending -and $hasTarget) { return }
        } finally { Release-ComObject $bookmarks }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "MathType native reference did not replace MTReference with a ZEqnNum bookmark."
}

function Complete-NativeMathTypeReference(
    [object]$document,
    [IntPtr]$wordHwnd,
    [string]$documentPath
) {
    $targetField = $null
    $targetCode = $null
    $targetResult = $null
    $selection = $null
    try {
        $targetField = Get-FirstNativeMathTypePlaceRefField $document
        if ($null -eq $targetField) {
            throw "No native MTPlaceRef field is available for the reference target."
        }
        $targetCode = $targetField.Code
        $targetResult = $targetField.Result
        Write-Output ("NATIVE_REFERENCE_TARGET|codeStart="+$targetCode.Start+"|codeEnd="+$targetCode.End+"|resultStart="+$targetResult.Start+"|resultEnd="+$targetResult.End+"|code="+(ConvertTo-WordDebugText ([string]$targetCode.Text)))

        [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow($wordHwnd)
        $targetField.Select()
        $selection = $document.Application.Selection
        Write-Output ("NATIVE_REFERENCE_TARGET_SELECTION|start="+$selection.Start+"|end="+$selection.End+"|text="+(ConvertTo-WordDebugText ([string]$selection.Text)))
        Start-Sleep -Milliseconds 300

        try {
            $targetField.DoClick()
            Write-Output "NATIVE_REFERENCE_TARGET_ACTION=Field.DoClick"
        } catch {
            Write-Output ("NATIVE_REFERENCE_DOCLICK_WARNING="+$_.Exception.Message)
            $document.Application.Run("MTPlaceRef")
            Write-Output "NATIVE_REFERENCE_TARGET_ACTION=Application.Run(MTPlaceRef)"
        }
        Wait-NativeMathTypeReferenceCompleted $document $TimeoutSeconds
        Start-Sleep -Milliseconds 700
        Capture-WordStructureSnapshot $document "reference-completed" $documentPath
        Write-Output ("NATIVE_REFERENCE_COMPLETED|fields="+$document.Fields.Count+"|bookmarks="+$document.Bookmarks.Count)
    } finally {
        Release-ComObject $selection
        Release-ComObject $targetResult
        Release-ComObject $targetCode
        Release-ComObject $targetField
    }
}

function Probe-NativeMathTypeReferenceEntry(
    [object]$document,
    [IntPtr]$wordHwnd,
    [string]$documentPath,
    [bool]$completeReference = $false
) {
    $selection = $null
    try {
        $selection = $document.Application.Selection
        $selection.SetRange($document.Content.End - 1, $document.Content.End - 1)
        $selection.TypeText("Native reference: ")
    } finally { Release-ComObject $selection }

    $wordRoot = Select-RibbonTab $wordHwnd "MathType"
    $button = Wait-ElementExact $wordRoot "插入引用" ([System.Windows.Automation.ControlType]::Button) $TimeoutSeconds
    $beforeHandles = @(Get-TopLevelWindows | ForEach-Object { [Int64]$_.Handle })
    Invoke-Element $button "MathType native Insert Reference"
    Start-Sleep -Milliseconds 1200
    Write-WindowInventory "windows-after-native-insert-reference-click.json"

    $newWindows = @(
        Get-TopLevelWindows |
            Where-Object {
                $beforeHandles -notcontains [Int64]$_.Handle -and
                -not [string]::IsNullOrWhiteSpace($_.Title) -and
                $_.ClassName -ne "MSO_BORDEREFFECT_WINDOW_CLASS"
            })
    if ($newWindows.Count -eq 0) {
        $wordRootAfter = [System.Windows.Automation.AutomationElement]::FromHandle($wordHwnd)
        Dump-AutomationTree $wordRootAfter (Join-Path $artifactPath "word-after-native-insert-reference-no-dialog-uia.txt")
        Capture-WordStructureSnapshot $document "reference-entry-no-dialog" $documentPath
        Write-Output "NATIVE_REFERENCE_ENTRY=NO_NEW_TOP_LEVEL_WINDOW"
        if ($completeReference) {
            Complete-NativeMathTypeReference $document $wordHwnd $documentPath
        }
        return
    }

    foreach ($window in $newWindows) {
        $windowHwnd = [IntPtr]$window.Handle
        Write-Output ("NATIVE_REFERENCE_WINDOW|pid="+$window.ProcessId+"|hwnd="+$window.Handle+"|title="+$window.Title+"|class="+$window.ClassName)
        try {
            $root = [System.Windows.Automation.AutomationElement]::FromHandle($windowHwnd)
            Dump-AutomationTree $root (Join-Path $artifactPath ("native-reference-window-" + $window.Handle + "-uia.txt"))
            Dump-NativeChildWindowTree $windowHwnd (Join-Path $artifactPath ("native-reference-window-" + $window.Handle + "-native.json"))
            Dump-AccessibleTree $windowHwnd (Join-Path $artifactPath ("native-reference-window-" + $window.Handle + "-msaa.txt"))
        } catch { }
    }
    Capture-WordStructureSnapshot $document "reference-entry-dialog-open" $documentPath

    foreach ($window in $newWindows) {
        try {
            [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow([IntPtr]$window.Handle)
            [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
            Start-Sleep -Milliseconds 250
        } catch { }
    }
    Write-Output "NATIVE_REFERENCE_ENTRY=DIALOG_TREE_COMPLETE"
}

function Open-NativeMathTypeNumberFormatDialog([IntPtr]$wordHwnd) {
    $wordRoot = Select-RibbonTab $wordHwnd "MathType"
    $splitButton = Wait-ElementExact $wordRoot "插入编号" ([System.Windows.Automation.ControlType]::SplitButton) $TimeoutSeconds
    $expandObject = $null
    if (-not $splitButton.TryGetCurrentPattern(
        [System.Windows.Automation.ExpandCollapsePattern]::Pattern,
        [ref]$expandObject)) {
        throw "MathType Insert Number split button has no ExpandCollapsePattern."
    }
    ([System.Windows.Automation.ExpandCollapsePattern]$expandObject).Expand()
    Start-Sleep -Milliseconds 500

    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $formatItem = $null
    foreach ($controlType in @(
        [System.Windows.Automation.ControlType]::MenuItem,
        [System.Windows.Automation.ControlType]::Button
    )) {
        $formatItem = Find-ElementExact $desktop "格式化..." $controlType
        if ($null -ne $formatItem) { break }
    }
    if ($null -eq $formatItem) {
        Dump-AutomationTree $desktop (Join-Path $artifactPath "desktop-after-insert-number-expand-uia-tree.txt")
        throw "MathType native '格式化...' menu item was not found after expanding Insert Number."
    }

    $beforeHandles = @(Get-TopLevelWindows | ForEach-Object { [Int64]$_.Handle })
    [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow($wordHwnd)
    $formatItem.SetFocus()
    Start-Sleep -Milliseconds 150
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    $dialog = Wait-NewTopLevelWindow $beforeHandles @("WINWORD", "MathType") $TimeoutSeconds
    [Console]::WriteLine("MATHTYPE_NUMBER_FORMAT_DIALOG|pid="+$dialog.ProcessId+"|hwnd="+$dialog.Handle+"|title="+$dialog.Title+"|class="+$dialog.ClassName)
    return $dialog
}

function Test-AccessibleChecked([IntPtr]$dialogHwnd, [string]$name) {
    $state = [VisualTeXMathTypeNativeProbe]::GetAccessibleStateByName($dialogHwnd, $name, 0)
    return (($state -band 0x10) -ne 0)
}

function Set-AccessibleChecked(
    [IntPtr]$dialogHwnd,
    [string]$name,
    [bool]$checked
) {
    $before = Test-AccessibleChecked $dialogHwnd $name
    if ($before -ne $checked) {
        $action = [VisualTeXMathTypeNativeProbe]::InvokeAccessibleByName($dialogHwnd, $name, 0)
        Write-Output ("MSAA_ACTION|"+$action+"|targetChecked="+$checked)
        Start-Sleep -Milliseconds 250
    }
    $after = Test-AccessibleChecked $dialogHwnd $name
    Write-Output ("MSAA_STATE|name="+$name+"|before="+$before+"|after="+$after+"|target="+$checked)
    if ($after -ne $checked) {
        throw "Accessible control '$name' did not reach checked=$checked."
    }
}

function Wait-TopLevelWindowClosed([Int64]$handle, [int]$seconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        $exists = Get-TopLevelWindows | Where-Object { [Int64]$_.Handle -eq $handle } | Select-Object -First 1
        if ($null -eq $exists) { return }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 120
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out waiting for top-level window $handle to close."
}

function Get-ComIdentityPointer([object]$value) {
    if ($null -eq $value -or -not [Runtime.InteropServices.Marshal]::IsComObject($value)) { return 0L }
    $pointer = [IntPtr]::Zero
    try {
        $pointer = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($value)
        return [Int64]$pointer
    } finally {
        if ($pointer -ne [IntPtr]::Zero) {
            [void][Runtime.InteropServices.Marshal]::Release($pointer)
        }
    }
}

function Capture-DocumentFieldIdentitySet([object]$document) {
    $entries = New-Object Collections.Generic.List[object]
    $fields = $null
    $field = $null
    $code = $null
    $result = $null
    $paragraphs = $null
    $paragraph = $null
    $paragraphRange = $null
    try {
        $fields = $document.Fields
        for ($index = 1; $index -le $fields.Count; $index++) {
            Release-ComObject $paragraphRange; $paragraphRange = $null
            Release-ComObject $paragraph; $paragraph = $null
            Release-ComObject $paragraphs; $paragraphs = $null
            Release-ComObject $result; $result = $null
            Release-ComObject $code; $code = $null
            $field = $fields.Item($index)
            $code = $field.Code
            $result = $field.Result
            $paragraphStart = $null
            try {
                $paragraphs = $field.Result.Paragraphs
                if ($paragraphs.Count -gt 0) {
                    $paragraph = $paragraphs.Item(1)
                    $paragraphRange = $paragraph.Range
                    $paragraphStart = [int]$paragraphRange.Start
                }
            } catch { }
            $entries.Add([pscustomobject][ordered]@{
                Index = $index
                Pointer = Get-ComIdentityPointer $field
                Type = [int]$field.Type
                CodeStart = [int]$code.Start
                CodeEnd = [int]$code.End
                ResultStart = [int]$result.Start
                ResultEnd = [int]$result.End
                ParagraphStart = $paragraphStart
                CodeText = ConvertTo-WordDebugText ([string]$code.Text)
                ResultText = ConvertTo-WordDebugText ([string]$result.Text)
                FieldObject = $field
            })
            # Ownership of this COM proxy transfers to the returned identity set.
            $field = $null
        }
        return ,$entries
    } finally {
        Release-ComObject $paragraphRange
        Release-ComObject $paragraph
        Release-ComObject $paragraphs
        Release-ComObject $result
        Release-ComObject $code
        Release-ComObject $field
        Release-ComObject $fields
    }
}

function Release-DocumentFieldIdentitySet([object]$entries) {
    if ($null -eq $entries) { return }
    foreach ($entry in $entries) {
        try { Release-ComObject $entry.FieldObject } catch { }
        try { $entry.FieldObject = $null } catch { }
    }
}

function Write-FieldIdentityTransition(
    [object]$beforeEntries,
    [object]$document,
    [string]$path
) {
    $afterEntries = $null
    try {
        $afterEntries = Capture-DocumentFieldIdentitySet $document
        $beforeReport = New-Object Collections.Generic.List[object]
        foreach ($before in $beforeEntries) {
            $heldAccessible = $false
            $heldCode = ""
            $heldResult = ""
            $heldCodeStart = $null
            $heldResultStart = $null
            $heldError = ""
            $heldCodeRange = $null
            $heldResultRange = $null
            try {
                $heldCodeRange = $before.FieldObject.Code
                $heldResultRange = $before.FieldObject.Result
                $heldCode = ConvertTo-WordDebugText ([string]$heldCodeRange.Text)
                $heldResult = ConvertTo-WordDebugText ([string]$heldResultRange.Text)
                $heldCodeStart = [int]$heldCodeRange.Start
                $heldResultStart = [int]$heldResultRange.Start
                $heldAccessible = $true
            } catch {
                $heldError = $_.Exception.Message
            } finally {
                Release-ComObject $heldResultRange
                Release-ComObject $heldCodeRange
            }
            $sameIdentity = @(
                $afterEntries |
                    Where-Object { [Int64]$_.Pointer -eq [Int64]$before.Pointer })
            $beforeReport.Add([pscustomobject][ordered]@{
                BeforeIndex = $before.Index
                Pointer = $before.Pointer
                BeforeType = $before.Type
                BeforeParagraphStart = $before.ParagraphStart
                BeforeCodeStart = $before.CodeStart
                BeforeCodeText = $before.CodeText
                BeforeResultText = $before.ResultText
                SameIdentityCountAfter = $sameIdentity.Count
                SameIdentityAfterIndexes = @($sameIdentity | ForEach-Object Index)
                SameIdentityAfterCodes = @($sameIdentity | ForEach-Object CodeText)
                HeldReferenceAccessibleAfter = $heldAccessible
                HeldReferenceCodeStartAfter = $heldCodeStart
                HeldReferenceResultStartAfter = $heldResultStart
                HeldReferenceCodeAfter = $heldCode
                HeldReferenceResultAfter = $heldResult
                HeldReferenceError = $heldError
            })
        }

        $afterReport = New-Object Collections.Generic.List[object]
        foreach ($after in $afterEntries) {
            $existedBefore = @(
                $beforeEntries |
                    Where-Object { [Int64]$_.Pointer -eq [Int64]$after.Pointer })
            $afterReport.Add([pscustomobject][ordered]@{
                AfterIndex = $after.Index
                Pointer = $after.Pointer
                Type = $after.Type
                ParagraphStart = $after.ParagraphStart
                CodeStart = $after.CodeStart
                CodeText = $after.CodeText
                ResultText = $after.ResultText
                SameIdentityCountBefore = $existedBefore.Count
                SameIdentityBeforeIndexes = @($existedBefore | ForEach-Object Index)
            })
        }
        $report = [pscustomobject][ordered]@{
            BeforeCount = $beforeEntries.Count
            AfterCount = $afterEntries.Count
            BeforeFields = $beforeReport.ToArray()
            AfterFields = $afterReport.ToArray()
        }
        Write-Utf8Text $path ($report | ConvertTo-Json -Depth 10)
        Write-Output ("FIELD_IDENTITY_TRANSITION="+$path)
    } finally {
        Release-DocumentFieldIdentitySet $afterEntries
        Release-DocumentFieldIdentitySet $beforeEntries
    }
}

function Set-MathTypeNumberFormatSeparator(
    [IntPtr]$dialogHwnd,
    [ValidateSet(".", "-")][string]$separator
) {
    try {
        $result = [VisualTeXMathTypeNativeProbe]::SetAccessibleValueByRoleAndCurrentValue(
            $dialogHwnd,
            42,
            ".",
            "-",
            0,
            $separator)
    } catch {
        # MathType clears this edit box after a one-component format such as
        # continuous numbering. In that state the separator is the only empty
        # editable-text node in the simple-format group.
        $result = [VisualTeXMathTypeNativeProbe]::SetAccessibleValueByRoleAndCurrentValue(
            $dialogHwnd,
            42,
            "",
            "",
            0,
            $separator)
    }
    Write-Output "MATHTYPE_FORMAT_VALUE|separator=$separator|$result"
    Start-Sleep -Milliseconds 250
}

function Apply-NativeMathTypeNumberFormat(
    [object]$document,
    [IntPtr]$wordHwnd,
    [string]$stageName,
    [bool]$chapter,
    [bool]$section,
    [string]$documentPath,
    [ValidateSet(".", "-")][string]$separator = "."
) {
    $fieldIdentitiesBefore = Capture-DocumentFieldIdentitySet $document
    $beforeHandles = @(Get-TopLevelWindows | ForEach-Object { [Int64]$_.Handle })
    $dialog = Open-NativeMathTypeNumberFormatDialog $wordHwnd
    $dialogHwnd = [IntPtr]$dialog.Handle
    try {
        Dump-AccessibleTree $dialogHwnd (Join-Path $artifactPath ($stageName + "-dialog-before-msaa.txt"))
        Set-AccessibleChecked $dialogHwnd "简单格式" $true
        # Keep MathType's native separator enabled throughout. Turning it off
        # clears the adjacent text box; re-checking it does not restore '.'.
        # With only one active number component the separator is naturally unused.
        Set-AccessibleChecked $dialogHwnd "分隔符:" $true
        Set-AccessibleChecked $dialogHwnd "章编号:" $chapter
        Set-AccessibleChecked $dialogHwnd "节编号:" $section
        Set-AccessibleChecked $dialogHwnd "公式编号:" $true
        Set-AccessibleChecked $dialogHwnd "附件:" $true
        Set-MathTypeNumberFormatSeparator $dialogHwnd $separator
        Set-AccessibleChecked $dialogHwnd "新的公式编号" $true
        Set-AccessibleChecked $dialogHwnd "整篇文档" $true
        Set-AccessibleChecked $dialogHwnd "自动更新公式编号" $true
        Set-AccessibleChecked $dialogHwnd "用作新文档的默认格式" $false
        Dump-AccessibleTree $dialogHwnd (Join-Path $artifactPath ($stageName + "-dialog-configured-msaa.txt"))

        $okAction = [VisualTeXMathTypeNativeProbe]::InvokeAccessibleByName($dialogHwnd, "确定", 0)
        Write-Output ("MSAA_ACTION|"+$okAction+"|stage="+$stageName)
        Wait-TopLevelWindowClosed ([Int64]$dialog.Handle) $TimeoutSeconds
    } finally { }

    Start-Sleep -Milliseconds 1200
    $newWindows = @(
        Get-TopLevelWindows |
            Where-Object {
                $beforeHandles -notcontains [Int64]$_.Handle -and
                -not [string]::IsNullOrWhiteSpace($_.Title) -and
                $_.ClassName -ne "MSO_BORDEREFFECT_WINDOW_CLASS"
            })
    if ($newWindows.Count -gt 0) {
        Write-WindowInventory ($stageName + "-unexpected-windows.json")
        foreach ($newWindow in $newWindows) {
            $newHwnd = [IntPtr]$newWindow.Handle
            try {
                $newRoot = [System.Windows.Automation.AutomationElement]::FromHandle($newHwnd)
                Dump-AutomationTree $newRoot (Join-Path $artifactPath ($stageName + "-unexpected-" + $newWindow.Handle + "-uia.txt"))
                Dump-AccessibleTree $newHwnd (Join-Path $artifactPath ($stageName + "-unexpected-" + $newWindow.Handle + "-msaa.txt"))
            } catch { }
        }
        throw "MathType format stage '$stageName' opened an unexpected follow-up window; evidence was captured."
    }

    $stagePath = Join-Path $artifactPath $stageName
    [IO.Directory]::CreateDirectory($stagePath) | Out-Null
    Write-FieldIdentityTransition $fieldIdentitiesBefore $document (Join-Path $stagePath "field-com-identity-transition.json")
    $fieldIdentitiesBefore = $null
    Capture-WordStructureSnapshot $document $stageName $documentPath
    Write-Output ("NATIVE_FORMAT_APPLIED|stage="+$stageName+"|chapter="+$chapter+"|section="+$section+"|separator="+$separator+"|fields="+$document.Fields.Count+"|paragraphs="+$document.Paragraphs.Count)
}

function Select-RibbonTab(
    [IntPtr]$wordHwnd,
    [string]$tabName
) {
    [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow($wordHwnd)
    Start-Sleep -Milliseconds 500
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($wordHwnd)
    $tab = Wait-ElementExact $root $tabName ([System.Windows.Automation.ControlType]::TabItem) $TimeoutSeconds
    $selectionObject = $null
    if (-not $tab.TryGetCurrentPattern(
        [System.Windows.Automation.SelectionItemPattern]::Pattern,
        [ref]$selectionObject)) {
        throw "Ribbon tab '$tabName' has no SelectionItemPattern."
    }
    ([System.Windows.Automation.SelectionItemPattern]$selectionObject).Select()
    Start-Sleep -Milliseconds 700
    return [System.Windows.Automation.AutomationElement]::FromHandle($wordHwnd)
}

function Get-SupportedPatternNames([System.Windows.Automation.AutomationElement]$element) {
    $names = New-Object Collections.Generic.List[string]
    foreach ($pattern in $element.GetSupportedPatterns()) {
        $names.Add($pattern.ProgrammaticName)
    }
    return ($names -join ",")
}

function Dump-AutomationTree(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$path
) {
    $lines = New-Object Collections.Generic.List[string]
    $rootRect = $root.Current.BoundingRectangle
    $lines.Add(
        "ROOT|$($root.Current.ControlType.ProgrammaticName)|NAME=$($root.Current.Name)|ID=$($root.Current.AutomationId)|CLASS=$($root.Current.ClassName)|RECT=$rootRect|PATTERNS=$(Get-SupportedPatternNames $root)")
    $elements = $root.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($element in $elements) {
        $name = $element.Current.Name
        $automationId = $element.Current.AutomationId
        $className = $element.Current.ClassName
        if ([string]::IsNullOrEmpty($name) -and
            [string]::IsNullOrEmpty($automationId) -and
            [string]::IsNullOrEmpty($className)) {
            continue
        }
        $rect = $element.Current.BoundingRectangle
        $patterns = Get-SupportedPatternNames $element
        $lines.Add(
            "EL|$($element.Current.ControlType.ProgrammaticName)|NAME=$name|ID=$automationId|CLASS=$className|ENABLED=$($element.Current.IsEnabled)|OFFSCREEN=$($element.Current.IsOffscreen)|RECT=$rect|PATTERNS=$patterns")
    }
    [IO.File]::WriteAllLines($path, $lines, [Text.UTF8Encoding]::new($false))
    Write-Output "UIA_TREE=$path"
}

$word = $null
$document = $null
$comAddIns = $null
$visualTeXAddIn = $null
$ownedWordPids = @()
$wordHwnd = [IntPtr]::Zero
$documentPath = Join-Path $artifactPath "MathType-Native-Number-Format-Probe.docx"
$beforeWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$beforeMathTypePids = @(Get-Process MathType -ErrorAction SilentlyContinue | ForEach-Object Id)
$clipboardBackup = $null
try { $clipboardBackup = [System.Windows.Forms.Clipboard]::GetDataObject() } catch { }

try {
    Write-Output ("PREEXISTING_WORD_PIDS=" + ($beforeWordPids -join ","))
    Write-Output ("PREEXISTING_MATHTYPE_PIDS=" + ($beforeMathTypePids -join ","))
    Write-WindowInventory "windows-before.json"

    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $document.Activate()
    Start-Sleep -Milliseconds 700

    $ownedWordPids = @(
        Get-Process WINWORD -ErrorAction SilentlyContinue |
            Where-Object { $beforeWordPids -notcontains $_.Id } |
            ForEach-Object Id)
    if ($ownedWordPids.Count -ne 1) {
        throw "The probe requires exactly one new isolated Word process; new PIDs=$($ownedWordPids -join ',')."
    }
    $wordHwnd = [IntPtr]$word.ActiveWindow.Hwnd
    Write-Output "OWNED_WORD_PID=$($ownedWordPids[0])"
    Write-Output "OWNED_WORD_HWND=$([Int64]$wordHwnd)"

    try {
        $comAddIns = $word.COMAddIns
        $visualTeXAddIn = $comAddIns.Item("VisualTeX.WordVsto")
        # Do not toggle Connect here. Office can block while unloading a VSTO
        # add-in from a freshly created automation instance. This probe invokes
        # only MathType's own Ribbon commands and never calls VisualTeX callbacks.
        Write-Output "VISUALTEX_ADDIN_CONNECTED_IN_PROBE=$($visualTeXAddIn.Connect)"
    } catch {
        Write-Output "VISUALTEX_ADDIN_STATE_WARNING=$($_.Exception.Message)"
    }

    Write-Output "TEST_DOCUMENT=UNSAVED_ISOLATED_DOCUMENT"
    $wordRoot = Select-RibbonTab $wordHwnd "MathType"
    Dump-AutomationTree $wordRoot (Join-Path $artifactPath "mathtype-ribbon-uia-tree.txt")
    Write-WindowInventory "windows-after-ribbon.json"

    if ($Mode -eq "RibbonTree") {
        Write-Output "MATHTYPE_NATIVE_NUMBER_PROBE=RIBBON_TREE_COMPLETE"
        return
    }

    $document.Range($document.Content.Start, $document.Content.Start).Select()
    Invoke-NativeMathTypeNumberedEquation $document $wordHwnd "right" "x+1" 1
    $selection = $word.Selection
    try {
        $selection.SetRange($document.Content.End - 1, $document.Content.End - 1)
        $selection.TypeParagraph()
    } finally { Release-ComObject $selection }
    Invoke-NativeMathTypeNumberedEquation $document $wordHwnd "left" "y+2" 2
    Write-WindowInventory "windows-after-native-inserts.json"
    Capture-WordStructureSnapshot $document "00-native-baseline" $documentPath

    if ($Mode -eq "ReferenceTree" -or $Mode -eq "ReferenceComplete" -or $Mode -eq "ReferenceFormats") {
        Probe-NativeMathTypeReferenceEntry $document $wordHwnd $documentPath ($Mode -ne "ReferenceTree")
        if ($Mode -eq "ReferenceFormats") {
            Apply-NativeMathTypeNumberFormat $document $wordHwnd "01-native-chapter-section-with-reference" $true $true $documentPath "."
            Apply-NativeMathTypeNumberFormat $document $wordHwnd "02-native-chapter-with-reference" $true $false $documentPath "."
            Apply-NativeMathTypeNumberFormat $document $wordHwnd "03-native-continuous-with-reference" $false $false $documentPath "."
            Apply-NativeMathTypeNumberFormat $document $wordHwnd "04-native-section-with-reference" $false $true $documentPath "."
            Apply-NativeMathTypeNumberFormat $document $wordHwnd "05-native-chapter-dash-with-reference" $true $false $documentPath "-"
            Apply-NativeMathTypeNumberFormat $document $wordHwnd "06-native-chapter-section-dash-with-reference" $true $true $documentPath "-"
        }
        Write-Output ("MATHTYPE_NATIVE_NUMBER_PROBE=" + $Mode.ToUpperInvariant() + "_COMPLETE")
        return
    }

    $dialog = Open-NativeMathTypeNumberFormatDialog $wordHwnd
    try {
        $dialogRoot = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$dialog.Handle)
        Dump-AutomationTree $dialogRoot (Join-Path $artifactPath "mathtype-number-format-dialog-uia-tree.txt")
        Dump-NativeChildWindowTree ([IntPtr]$dialog.Handle) (Join-Path $artifactPath "mathtype-number-format-dialog-native-tree.json")
        Dump-AccessibleTree ([IntPtr]$dialog.Handle) (Join-Path $artifactPath "mathtype-number-format-dialog-msaa-tree.txt")
        Write-WindowInventory "windows-with-number-format-dialog.json"
        [void][VisualTeXMathTypeNativeProbe]::SetForegroundWindow([IntPtr]$dialog.Handle)
        [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
        Start-Sleep -Milliseconds 500
    } finally { }

    if ($Mode -eq "DialogTree") {
        Write-Output "MATHTYPE_NATIVE_NUMBER_PROBE=DIALOG_TREE_COMPLETE"
        return
    }

    Apply-NativeMathTypeNumberFormat $document $wordHwnd "01-native-continuous" $false $false $documentPath
    Apply-NativeMathTypeNumberFormat $document $wordHwnd "02-native-chapter" $true $false $documentPath
    Apply-NativeMathTypeNumberFormat $document $wordHwnd "03-native-section" $false $true $documentPath
    Apply-NativeMathTypeNumberFormat $document $wordHwnd "04-native-chapter-section" $true $true $documentPath
    Write-Output "MATHTYPE_NATIVE_NUMBER_PROBE=FULL_COMPLETE"
}
finally {
    if ($null -ne $clipboardBackup) {
        try { [System.Windows.Forms.Clipboard]::SetDataObject($clipboardBackup, $true) } catch { }
    }
    Release-ComObject $visualTeXAddIn
    Release-ComObject $comAddIns
    if ($null -ne $document) {
        try { $document.Close(0) } catch { }
    }
    Release-ComObject $document
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
    }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 700

    foreach ($ownedPid in $ownedWordPids) {
        $process = Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($null -eq $process) { continue }
        try {
            $cim = Get-CimInstance Win32_Process -Filter "ProcessId=$ownedPid" -ErrorAction SilentlyContinue
            if ($null -ne $cim -and $cim.CommandLine -like "*/Automation -Embedding*") {
                Stop-Process -Id $ownedPid -Force
                Write-Output "FORCED_CLEANUP_OWNED_WORD_PID=$ownedPid"
            } else {
                Write-Output "OWNED_WORD_CLEANUP_SKIPPED_UNEXPECTED_COMMANDLINE=$ownedPid"
            }
        } catch {
            Write-Output "OWNED_WORD_CLEANUP_WARNING=$ownedPid|$($_.Exception.Message)"
        }
    }
}
