-- Source form for the compiled AppleScriptTask file installed as
-- ~/Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXPowerPoint.scpt

use scripting additions
use framework "Foundation"

property runtimeSuffix : "Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime"
property maximumRelativePathLength : 1024
property expectedHost : "powerpoint"
property cachedVisualTeXExecutable : ""

on OpenVisualTeXSession(sessionId)
    try
        set safeSessionId to my validateSessionId(sessionId as text)
        set visualTeXURL to "visualtex://office/open?session=" & safeSessionId
        my launchVisualTeXURL(visualTeXURL)
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end OpenVisualTeXSession

on WriteAndOpenVisualTeXSession(argumentText)
    set startedAt to my monotonicSeconds()
    try
        set {requestedHost, sessionId, encodedData} to my splitTriple(argumentText as text)
        set safeHost to my validateHostName(requestedHost)
        set safeSessionId to my validateSessionId(sessionId)
        set targetPath to my absoluteRuntimePath("OfficeSessions/" & safeSessionId & "/request.json")
        set validatedAt to my monotonicSeconds()
        my writeEncodedFileAtomically(targetPath, encodedData)
        set writtenAt to my monotonicSeconds()
        set visualTeXURL to "visualtex://office/open?session=" & safeSessionId
        my launchVisualTeXURL(visualTeXURL)
        set launchedAt to my monotonicSeconds()
        return "ok|host=" & safeHost & ";validationMs=" & my elapsedMilliseconds(startedAt, validatedAt) & ";writeMs=" & my elapsedMilliseconds(validatedAt, writtenAt) & ";launchMs=" & my elapsedMilliseconds(writtenAt, launchedAt) & ";totalMs=" & my elapsedMilliseconds(startedAt, launchedAt)
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end WriteAndOpenVisualTeXSession

on PrewarmVisualTeXApplication(hostName)
    set startedAt to my monotonicSeconds()
    try
        set safeHost to my validateHostName(hostName as text)
        -- Resolve a fully launched resident while Office itself is starting.
        -- This keeps every later formula click on the in-process AppKit fast path.
        set cachedVisualTeXExecutable to my runningVisualTeXExecutable()
        set finishedAt to my monotonicSeconds()
        return "ok|host=" & safeHost & ";prewarmMs=" & my elapsedMilliseconds(startedAt, finishedAt)
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end PrewarmVisualTeXApplication

on OpenVisualTeXApplication(ignoredValue)
    try
        do shell script "/usr/bin/open -b " & quoted form of "com.visualtex.studio"
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end OpenVisualTeXApplication

on EnsureVisualTeXDirectory(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        my ensureDirectory(targetPath)
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end EnsureVisualTeXDirectory

on WriteVisualTeXFile(argumentText)
    try
        set {relativePath, encodedData} to my splitPair(argumentText as text)
        set targetPath to my absoluteRuntimePath(relativePath)
        my writeEncodedFileAtomically(targetPath, encodedData)
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end WriteVisualTeXFile

on AppendVisualTeXFile(argumentText)
    try
        set {relativePath, encodedData} to my splitPair(argumentText as text)
        set targetPath to my absoluteRuntimePath(relativePath)
        set parentPath to do shell script "/usr/bin/dirname " & quoted form of targetPath
        my ensureDirectory(parentPath)
        set normalizedData to my normalizeBase64Url(encodedData)
        do shell script "umask 077; /usr/bin/printf %s " & quoted form of normalizedData & " | /usr/bin/base64 -D >> " & quoted form of targetPath & " && /bin/chmod 600 " & quoted form of targetPath
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end AppendVisualTeXFile

on ReadVisualTeXFile(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        do shell script "/bin/test -f " & quoted form of targetPath
        set encodedData to do shell script "/usr/bin/base64 < " & quoted form of targetPath & " | /usr/bin/tr -d '\r\n'"
        set encodedData to my replaceText(encodedData, "+", "-")
        set encodedData to my replaceText(encodedData, "/", "_")
        repeat while encodedData ends with "="
            if (count characters of encodedData) is 1 then
                set encodedData to ""
            else
                set encodedData to text 1 thru -2 of encodedData
            end if
        end repeat
        return "ok|" & encodedData
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end ReadVisualTeXFile

on VisualTeXFileExists(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        try
            do shell script "/bin/test -f " & quoted form of targetPath
            return "ok|1"
        on error
            return "ok|0"
        end try
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end VisualTeXFileExists

on DeleteVisualTeXFile(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        do shell script "/bin/rm -f " & quoted form of targetPath
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end DeleteVisualTeXFile

on absoluteRuntimePath(relativePath)
    set safeRelativePath to my validateRelativePath(relativePath)
    set rootPath to my ensureRuntimeRoot()
    return rootPath & "/" & safeRelativePath
end absoluteRuntimePath

on ensureRuntimeRoot()
    set homePath to POSIX path of (path to home folder)
    set rootPath to homePath & runtimeSuffix
    my ensureDirectory(rootPath)
    return rootPath
end ensureRuntimeRoot

on ensureDirectory(targetPath)
    do shell script "/bin/mkdir -p " & quoted form of targetPath & " && /bin/chmod 700 " & quoted form of targetPath
end ensureDirectory

on launchVisualTeXURL(visualTeXURL)
    set safeURL to visualTeXURL as text
    if safeURL does not start with "visualtex://office/open?session=" then error "VisualTeX launch URL is invalid" number 7127
    set executablePath to my runningVisualTeXExecutable()
    set diagnosticPath to my absoluteRuntimePath("Tests/powerpoint-launch-diagnostic.txt")
    set diagnosticText to "executable=" & executablePath & linefeed
    try
        do shell script "/bin/test -x " & quoted form of executablePath
        set diagnosticText to diagnosticText & "executableTest=ok" & linefeed
    on error testMessage number testNumber
        set diagnosticText to diagnosticText & "executableTest=fail:" & testNumber & ":" & testMessage & linefeed
        my writeDiagnosticText(diagnosticPath, diagnosticText)
        error "VisualTeX executable is not accessible from PowerPoint AppleScriptTask: " & executablePath number 7128
    end try
    try
        set processIds to do shell script "/usr/bin/pgrep -x " & quoted form of "visualtex"
        set diagnosticText to diagnosticText & "pgrep=" & processIds & linefeed
    on error processMessage number processNumber
        set diagnosticText to diagnosticText & "pgrep=fail:" & processNumber & ":" & processMessage & linefeed
    end try
    try
        -- PowerPoint's sandbox cannot reliably join Tauri's single-instance IPC
        -- when it launches a second VisualTeX executable directly. Deliver the
        -- validated URL through LaunchServices to the already-running production
        -- bundle instead. The heartbeat check above still guarantees that the
        -- resident process is the expected VisualTeX executable before dispatch.
        do shell script "/usr/bin/open -b " & quoted form of "com.visualtex.studio" & space & quoted form of safeURL
        set diagnosticText to diagnosticText & "launchServices=ok" & linefeed
        my writeDiagnosticText(diagnosticPath, diagnosticText)
    on error launchMessage number launchNumber
        set diagnosticText to diagnosticText & "launchServices=fail:" & launchNumber & ":" & launchMessage & linefeed
        my writeDiagnosticText(diagnosticPath, diagnosticText)
        error "VisualTeX resident URL forwarding failed: " & launchMessage & " [executable=" & executablePath & "]" number 7128
    end try
end launchVisualTeXURL

on writeDiagnosticText(targetPath, diagnosticText)
    set diagnosticData to current application's NSString's stringWithString:(diagnosticText as text)
    set writeSucceeded to diagnosticData's writeToFile:targetPath atomically:true encoding:(current application's NSUTF8StringEncoding) |error|:(missing value)
    if not writeSucceeded then return false
    try
        do shell script "/bin/chmod 600 " & quoted form of targetPath
    end try
    return true
end writeDiagnosticText

on runningVisualTeXExecutable()
    set executableSuffix to "/VisualTeX.app/Contents/MacOS/visualtex"

    -- Bind PowerPoint to the exact resident that wrote its FastOpen heartbeat.
    -- This avoids forwarding to another installed/development VisualTeX process
    -- when several app bundles are running at the same time.
    set runningExecutable to my readyResidentVisualTeXExecutable(executableSuffix)
    if runningExecutable is not "" then
        set cachedVisualTeXExecutable to runningExecutable
        return runningExecutable
    end if

    if cachedVisualTeXExecutable is not "" then
        if my isRunningVisualTeXExecutable(cachedVisualTeXExecutable, executableSuffix) then return cachedVisualTeXExecutable
        set cachedVisualTeXExecutable to ""
    end if

    -- Compatibility fallback for residents from before the PID/path heartbeat.
    set runningExecutable to my firstRunningVisualTeXExecutable(executableSuffix)
    if runningExecutable is not "" then
        set cachedVisualTeXExecutable to runningExecutable
        return runningExecutable
    end if

    do shell script "/usr/bin/open -gj -b " & quoted form of "com.visualtex.studio" & " --args --office-background"
    repeat with attemptIndex from 1 to 80
        delay 0.05
        set runningExecutable to my readyResidentVisualTeXExecutable(executableSuffix)
        if runningExecutable is "" then set runningExecutable to my firstRunningVisualTeXExecutable(executableSuffix)
        if runningExecutable is not "" then
            delay 0.5
            set verifiedExecutable to my readyResidentVisualTeXExecutable(executableSuffix)
            if verifiedExecutable is "" then set verifiedExecutable to my firstRunningVisualTeXExecutable(executableSuffix)
            if verifiedExecutable is not "" then
                set cachedVisualTeXExecutable to verifiedExecutable
                return verifiedExecutable
            end if
        end if
    end repeat
    error "The prewarmed VisualTeX executable is not running" number 7128
end runningVisualTeXExecutable

on readyResidentVisualTeXExecutable(executableSuffix)
    set markerPath to my fastOpenReadyMarkerPath()
    try
        set markerText to do shell script "/bin/cat " & quoted form of markerPath
    on error
        return ""
    end try
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to linefeed
    set markerLines to text items of markerText
    set AppleScript's text item delimiters to previousDelimiters
    if (count of markerLines) < 4 then return ""
    if item 1 of markerLines is not "visualtex-fast-open-ready-v2" then return ""
    set processId to item 3 of markerLines as text
    set candidatePath to item 4 of markerLines as text
    if not my isDecimalProcessId(processId) then return ""
    if candidatePath does not end with executableSuffix then return ""
    try
        do shell script "/bin/test -x " & quoted form of candidatePath
        set actualPath to do shell script "/bin/ps -p " & quoted form of processId & " -o comm="
        if actualPath is candidatePath then return candidatePath
    end try
    return ""
end readyResidentVisualTeXExecutable

on fastOpenReadyMarkerPath()
    set homePath to POSIX path of (path to home folder)
    return homePath & "Library/Containers/com.microsoft.Powerpoint/Data/Library/Application Support/VisualTeX/FastOpen/powerpoint/resident-ready"
end fastOpenReadyMarkerPath

on isRunningVisualTeXExecutable(candidatePath, executableSuffix)
    if candidatePath is "" or candidatePath does not end with executableSuffix then return false
    set processIds to ""
    try
        do shell script "/bin/test -x " & quoted form of candidatePath
        set processIds to do shell script "/usr/bin/pgrep -x " & quoted form of "visualtex"
    on error
        return false
    end try
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to linefeed
    set processIdItems to text items of processIds
    set AppleScript's text item delimiters to previousDelimiters
    repeat with processIdItem in processIdItems
        set processId to processIdItem as text
        if my isDecimalProcessId(processId) then
            try
                set actualPath to do shell script "/bin/ps -p " & quoted form of processId & " -o comm="
                if actualPath is candidatePath then return true
            end try
        end if
    end repeat
    return false
end isRunningVisualTeXExecutable

on firstRunningVisualTeXExecutable(executableSuffix)
    set processIds to ""
    try
        set processIds to do shell script "/usr/bin/pgrep -x " & quoted form of "visualtex"
    end try
    if processIds is "" then return ""
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to linefeed
    set processIdItems to text items of processIds
    set AppleScript's text item delimiters to previousDelimiters
    repeat with processIdItem in processIdItems
        set processId to processIdItem as text
        if my isDecimalProcessId(processId) then
            try
                set candidatePath to do shell script "/bin/ps -p " & quoted form of processId & " -o comm="
                if candidatePath ends with executableSuffix then
                    do shell script "/bin/test -x " & quoted form of candidatePath
                    return candidatePath
                end if
            end try
        end if
    end repeat
    return ""
end firstRunningVisualTeXExecutable

on isDecimalProcessId(candidate)
    set candidate to candidate as text
    if candidate is "" then return false
    repeat with currentCharacter in characters of candidate
        if "0123456789" does not contain (currentCharacter as text) then return false
    end repeat
    return true
end isDecimalProcessId

on writeEncodedFileAtomically(targetPath, encodedData)
    try
        set parentPath to ((current application's NSString's stringWithString:targetPath)'s stringByDeletingLastPathComponent()) as text
        my ensureDirectory(parentPath)
        set normalizedData to my normalizeBase64Url(encodedData)
        set decodedData to current application's NSData's alloc()'s initWithBase64EncodedString:normalizedData options:0
        if decodedData is missing value then error "VisualTeX file bridge Base64URL payload is invalid" number 7125
        set writeSucceeded to (decodedData's writeToFile:targetPath atomically:true) as boolean
        if not writeSucceeded then error "VisualTeX could not write the local Session request" number 7129
        do shell script "/bin/chmod 600 " & quoted form of targetPath
    on error errorMessage number errorNumber
        error errorMessage number errorNumber
    end try
end writeEncodedFileAtomically

on validateRelativePath(candidate)
    set candidate to candidate as text
    if candidate is "" then error "VisualTeX runtime path is empty" number 7120
    if (count characters of candidate) > maximumRelativePathLength then error "VisualTeX runtime path is too long" number 7121
    if candidate starts with "/" or candidate ends with "/" or candidate is "." or candidate contains ".." or candidate contains "//" then error "VisualTeX runtime path is unsafe" number 7122
    set allowedCharacters to "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-/"
    repeat with currentCharacter in characters of candidate
        if allowedCharacters does not contain (currentCharacter as text) then error "VisualTeX runtime path contains an unsupported character" number 7123
    end repeat
    return candidate
end validateRelativePath

on splitPair(value)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to "|"
    set fields to text items of value
    set AppleScript's text item delimiters to previousDelimiters
    if (count fields) is not 2 then error "VisualTeX file bridge payload is invalid" number 7124
    return {item 1 of fields, item 2 of fields}
end splitPair

on splitTriple(value)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to "|"
    set fields to text items of value
    set AppleScript's text item delimiters to previousDelimiters
    if (count fields) is not 3 then error "VisualTeX write-and-launch payload is invalid" number 7126
    return {item 1 of fields, item 2 of fields, item 3 of fields}
end splitTriple

on normalizeBase64Url(encodedData)
    set normalizedData to my replaceText(encodedData as text, "-", "+")
    set normalizedData to my replaceText(normalizedData, "_", "/")
    set remainderValue to (count characters of normalizedData) mod 4
    if remainderValue is 1 then error "VisualTeX file bridge Base64URL payload is invalid" number 7125
    if remainderValue is 2 then set normalizedData to normalizedData & "=="
    if remainderValue is 3 then set normalizedData to normalizedData & "="
    return normalizedData
end normalizeBase64Url

on replaceText(sourceText, searchText, replacementText)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to searchText
    set sourceItems to text items of sourceText
    set AppleScript's text item delimiters to replacementText
    set resultText to sourceItems as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end replaceText

on validateSessionId(candidate)
    if (count characters of candidate) is not 36 then error "Invalid VisualTeX Session id" number 7101
    if character 9 of candidate is not "-" or character 14 of candidate is not "-" or character 19 of candidate is not "-" or character 24 of candidate is not "-" then error "Invalid VisualTeX Session id" number 7102
    if character 15 of candidate is not "4" then error "Invalid VisualTeX Session version" number 7103
    if "89ab" does not contain character 20 of candidate then error "Invalid VisualTeX Session variant" number 7104

    set allowedHex to "0123456789abcdef"
    repeat with characterIndex from 1 to 36
        set currentCharacter to character characterIndex of candidate
        if characterIndex is 9 or characterIndex is 14 or characterIndex is 19 or characterIndex is 24 then
            if currentCharacter is not "-" then error "Invalid VisualTeX Session id" number 7105
        else if allowedHex does not contain currentCharacter then
            error "Invalid VisualTeX Session id" number 7106
        end if
    end repeat
    return candidate
end validateSessionId

on validateHostName(candidate)
    set candidate to candidate as text
    if candidate is not expectedHost then error "VisualTeX Office host does not match its Application Script" number 7107
    return candidate
end validateHostName

on monotonicSeconds()
    return (current application's NSProcessInfo's processInfo()'s systemUptime()) as real
end monotonicSeconds

on elapsedMilliseconds(startedAt, finishedAt)
    return (round ((finishedAt - startedAt) * 1000)) as integer
end elapsedMilliseconds

on errorResponse(errorNumber, errorMessage)
    return "error|" & (errorNumber as text) & "|" & my safeError(errorMessage)
end errorResponse

on safeError(value)
    set cleanValue to value as text
    set AppleScript's text item delimiters to {return, linefeed, "|"}
    set cleanItems to text items of cleanValue
    set AppleScript's text item delimiters to " "
    set cleanValue to cleanItems as text
    set AppleScript's text item delimiters to ""
    if (count characters of cleanValue) > 240 then set cleanValue to text 1 thru 240 of cleanValue
    return cleanValue
end safeError
