-- Source form for the compiled AppleScriptTask file installed as
-- ~/Library/Application Scripts/com.microsoft.Word/VisualTeXWord.scpt

use scripting additions
use framework "Foundation"

property runtimeSuffix : "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime"
property maximumRelativePathLength : 1024
property expectedHost : "word"
property numberingPreferenceFileName : "VisualTeXNumberingPreference.txt"
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

on WriteFormulaRestoreAndOpenVisualTeXSession(argumentText)
    set startedAt to my monotonicSeconds()
    try
        set {requestedHost, sessionId, encodedRequest, encodedSource} to my splitQuadruple(argumentText as text)
        set safeHost to my validateHostName(requestedHost)
        set safeSessionId to my validateSessionId(sessionId)
        set sessionDirectory to "OfficeSessions/" & safeSessionId
        set requestPath to my absoluteRuntimePath(sessionDirectory & "/request.json")
        set sourcePath to my absoluteRuntimePath(sessionDirectory & "/formula-restore-source.txt")
        set validatedAt to my monotonicSeconds()
        my writeEncodedFileAtomically(requestPath, encodedRequest)
        my writeEncodedFileAtomically(sourcePath, encodedSource)
        set writtenAt to my monotonicSeconds()
        set visualTeXURL to "visualtex://office/open?session=" & safeSessionId
        my launchVisualTeXURL(visualTeXURL)
        set launchedAt to my monotonicSeconds()
        return "ok|host=" & safeHost & ";validationMs=" & my elapsedMilliseconds(startedAt, validatedAt) & ";writeMs=" & my elapsedMilliseconds(validatedAt, writtenAt) & ";launchMs=" & my elapsedMilliseconds(writtenAt, launchedAt) & ";totalMs=" & my elapsedMilliseconds(startedAt, launchedAt)
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end WriteFormulaRestoreAndOpenVisualTeXSession

on ConvertOmmlBatch(argumentText)
    set startedAt to my monotonicSeconds()
    try
        set rawArgument to argumentText as text
        if rawArgument contains "|" then
            set {sessionId, encodedSource} to my splitPair(rawArgument)
        else
            set sessionId to rawArgument
            set encodedSource to ""
        end if
        set safeSessionId to my validateSessionId(sessionId)
        set sessionDirectory to "OfficeSessions/" & safeSessionId
        set inputPath to my absoluteRuntimePath(sessionDirectory & "/formula-restore-source.txt")
        set outputPath to my absoluteRuntimePath(sessionDirectory & "/omml-latex-result.txt")
        if encodedSource is not "" then
            my writeEncodedFileAtomically(inputPath, encodedSource)
        end if
        set writtenAt to my monotonicSeconds()
        set executablePath to "/Applications/VisualTeX.app/Contents/MacOS/visualtex"
        set fileManager to current application's NSFileManager's defaultManager()
        if not ((fileManager's isExecutableFileAtPath:executablePath) as boolean) then
            set executablePath to my runningVisualTeXExecutable()
        end if
        set errorPipe to current application's NSPipe's pipe()
        set convertTask to current application's NSTask's alloc()'s init()
        convertTask's setLaunchPath:executablePath
        convertTask's setArguments:{"--office-omml-to-latex-batch", inputPath, outputPath}
        convertTask's setStandardOutput:(current application's NSFileHandle's fileHandleWithNullDevice())
        convertTask's setStandardError:errorPipe
        convertTask's |launch|()
        convertTask's waitUntilExit()
        set finishedAt to my monotonicSeconds()
        if (convertTask's terminationStatus() as integer) is not 0 then
            set errorData to errorPipe's fileHandleForReading()'s readDataToEndOfFile()
            set errorText to current application's NSString's alloc()'s initWithData:errorData encoding:(current application's NSUTF8StringEncoding)
            if errorText is missing value then set errorText to "VisualTeX OMML conversion failed"
            error (errorText as text) number 7130
        end if
        set resultData to current application's NSData's dataWithContentsOfFile:outputPath
        if resultData is missing value or (resultData's |length|() as integer) is 0 then error "VisualTeX OMML conversion returned no result" number 7131
        set encodedResult to (resultData's base64EncodedStringWithOptions:0) as text
        set encodedResult to my replaceText(encodedResult, "+", "-")
        set encodedResult to my replaceText(encodedResult, "/", "_")
        repeat while encodedResult ends with "="
            if (count characters of encodedResult) is 1 then
                set encodedResult to ""
            else
                set encodedResult to text 1 thru -2 of encodedResult
            end if
        end repeat
        return "ok|writeMs=" & my elapsedMilliseconds(startedAt, writtenAt) & ";convertMs=" & my elapsedMilliseconds(writtenAt, finishedAt) & ";totalMs=" & my elapsedMilliseconds(startedAt, finishedAt) & "|" & encodedResult
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end ConvertOmmlBatch

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

on WriteVisualTeXNumberingPreference(encodedData)
    try
        set targetPath to my numberingPreferencePath()
        my writeEncodedFileAtomically(targetPath, encodedData as text)
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end WriteVisualTeXNumberingPreference

on ReadVisualTeXNumberingPreference(ignoredValue)
    try
        set targetPath to my numberingPreferencePath()
        set fileManager to current application's NSFileManager's defaultManager()
        if not ((fileManager's fileExistsAtPath:targetPath) as boolean) then return "ok|"
        set preferenceData to current application's NSData's dataWithContentsOfFile:targetPath
        if preferenceData is missing value then error "VisualTeX could not read the numbering preference" number 7132
        if (preferenceData's |length|() as integer) > 256 then error "VisualTeX numbering preference is too large" number 7133
        set encodedPreference to (preferenceData's base64EncodedStringWithOptions:0) as text
        set encodedPreference to my replaceText(encodedPreference, "+", "-")
        set encodedPreference to my replaceText(encodedPreference, "/", "_")
        repeat while encodedPreference ends with "="
            if (count characters of encodedPreference) is 1 then
                set encodedPreference to ""
            else
                set encodedPreference to text 1 thru -2 of encodedPreference
            end if
        end repeat
        return "ok|" & encodedPreference
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end ReadVisualTeXNumberingPreference

on ReadVisualTeXImageInkCenter(formulaId)
    try
        set formulaId to formulaId as text
        if (count characters of formulaId) is not 36 then error "VisualTeX formula id is invalid" number 7134
        set fileManager to current application's NSFileManager's defaultManager()
        set executablePath to "/Applications/VisualTeX.app/Contents/MacOS/visualtex"
        if not ((fileManager's isExecutableFileAtPath:executablePath) as boolean) then
            set executablePath to my runningVisualTeXExecutable()
        end if
        set ratioText to do shell script quoted form of executablePath & " --office-image-ink-center " & quoted form of formulaId
        return "ok|" & ratioText
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end ReadVisualTeXImageInkCenter

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

on numberingPreferencePath()
    set homePath to POSIX path of (path to home folder)
    return homePath & "Library/Application Scripts/com.microsoft.Word/" & numberingPreferenceFileName
end numberingPreferencePath

on ensureDirectory(targetPath)
    do shell script "/bin/mkdir -p " & quoted form of targetPath & " && /bin/chmod 700 " & quoted form of targetPath
end ensureDirectory

on launchVisualTeXURL(visualTeXURL)
    set safeURL to visualTeXURL as text
    if safeURL does not start with "visualtex://office/open?session=" then error "VisualTeX launch URL is invalid" number 7127
    set executablePath to my runningVisualTeXExecutable()
    -- Keep the detached forwarding helper from the stable cold-launch path so
    -- Office never falls back to another foreground application while the
    -- resident editor window is being raised.
    do shell script "/usr/bin/nohup " & quoted form of executablePath & space & quoted form of safeURL & " >/dev/null 2>&1 &"
end launchVisualTeXURL

on runningVisualTeXExecutable()
    -- Prewarm resolves and caches the exact resident executable before the user
    -- opens a formula. Do not re-run pgrep/ps/test on every hot editor launch.
    if cachedVisualTeXExecutable is not "" then return cachedVisualTeXExecutable

    set executableSuffix to "/VisualTeX.app/Contents/MacOS/visualtex"
    set runningExecutable to my firstRunningVisualTeXExecutable(executableSuffix)
    if runningExecutable is not "" then
        set cachedVisualTeXExecutable to runningExecutable
        return runningExecutable
    end if

    -- Cold launch retains the validated b201fde behavior: start VisualTeX in
    -- background-only mode and wait until the real process is observable before
    -- forwarding the first Session URL.
    do shell script "/usr/bin/open -gj -b " & quoted form of "com.visualtex.studio" & " --args --office-background"
    repeat with attemptIndex from 1 to 80
        delay 0.05
        set runningExecutable to my firstRunningVisualTeXExecutable(executableSuffix)
        if runningExecutable is not "" then
            delay 0.5
            set runningExecutable to my firstRunningVisualTeXExecutable(executableSuffix)
            if runningExecutable is not "" then
                set cachedVisualTeXExecutable to runningExecutable
                return runningExecutable
            end if
        end if
    end repeat
    error "The prewarmed VisualTeX executable is not running" number 7128
end runningVisualTeXExecutable

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

on splitQuadruple(value)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to "|"
    set fields to text items of value
    set AppleScript's text item delimiters to previousDelimiters
    if (count fields) is not 4 then error "VisualTeX formula-restore write-and-launch payload is invalid" number 7127
    return {item 1 of fields, item 2 of fields, item 3 of fields, item 4 of fields}
end splitQuadruple

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
