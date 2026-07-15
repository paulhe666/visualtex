-- Source form for the compiled AppleScriptTask file installed as
-- ~/Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXPowerPoint.scpt

on OpenVisualTeXSession(sessionId)
    try
        set safeSessionId to my validateSessionId(sessionId as text)
        set visualTeXURL to "visualtex://office/open?session=" & safeSessionId
        do shell script "/usr/bin/open " & quoted form of visualTeXURL
        return "ok|1"
    on error errorMessage number errorNumber
        return "error|" & (errorNumber as text) & "|" & my safeError(errorMessage)
    end try
end OpenVisualTeXSession

on OpenVisualTeXApplication(ignoredValue)
    try
        do shell script "/usr/bin/open -b " & quoted form of "com.visualtex.studio"
        return "ok|1"
    on error errorMessage number errorNumber
        return "error|" & (errorNumber as text) & "|" & my safeError(errorMessage)
    end try
end OpenVisualTeXApplication

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
