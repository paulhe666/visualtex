-- Source form for the compiled AppleScriptTask file installed as
-- ~/Library/Application Scripts/com.microsoft.Word/VisualTeXWord.scpt

use scripting additions

property runtimeSuffix : "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime"
property maximumRelativePathLength : 1024

on OpenVisualTeXSession(sessionId)
    try
        set safeSessionId to my validateSessionId(sessionId as text)
        set visualTeXURL to "visualtex://office/open?session=" & safeSessionId
        do shell script "/usr/bin/open " & quoted form of visualTeXURL
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end OpenVisualTeXSession

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
    set temporaryPath to ""
    try
        set {relativePath, encodedData} to my splitPair(argumentText as text)
        set targetPath to my absoluteRuntimePath(relativePath)
        set parentPath to do shell script "/usr/bin/dirname " & quoted form of targetPath
        my ensureDirectory(parentPath)
        set normalizedData to my normalizeBase64Url(encodedData)
        set temporaryPath to do shell script "/usr/bin/mktemp " & quoted form of (targetPath & ".tmp.XXXXXX")
        do shell script "/usr/bin/printf %s " & quoted form of normalizedData & " | /usr/bin/base64 -D > " & quoted form of temporaryPath
        do shell script "/bin/chmod 600 " & quoted form of temporaryPath & " && /bin/mv -f " & quoted form of temporaryPath & space & quoted form of targetPath
        set temporaryPath to ""
        return "ok|1"
    on error errorMessage number errorNumber
        if temporaryPath is not "" then
            try
                do shell script "/bin/rm -f " & quoted form of temporaryPath
            end try
        end if
        return my errorResponse(errorNumber, errorMessage)
    end try
end WriteVisualTeXFile

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

on validateRelativePath(candidate)
    set candidate to candidate as text
    if candidate is "" then error "VisualTeX runtime path is empty" number 7120
    if (count characters of candidate) > maximumRelativePathLength then error "VisualTeX runtime path is too long" number 7121
    if candidate starts with "/" or candidate contains ".." or candidate contains "//" then error "VisualTeX runtime path is unsafe" number 7122
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
