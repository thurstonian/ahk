#Requires AutoHotkey v2.0

; Functions that have to do with email.
;; TODO
;; Get master list of categories from helpdesk inbox somehow

; Fetches current selected email within Outlook
GetCurrentEmail() {
    Return ComObjActive("Outlook.Application").ActiveExplorer.Selection.Item(1)
}

; Fetches the Full Name of the sender of the COM MailItem Object, or the currently selected email
GetSender(email) {
    Return email.SenderName
}

; Wrapper for GetSender to get Sender from Current Email
GetCurrentSender() {
    Return GetSender(GetCurrentEmail())
}

; Converts a Full Name into a Standard Name by removing the Middle Initial
; If none passed in, defaults to sender of Current Email
GetStandardName(name := "") {
    If (name = "") {
        name := GetCurrentSender()
    }
    Return RegExReplace(name, "(\w+)(?:, )(\w+)", "$2 $1")
}

; Extracts the first name from a Standard Name
; If none, defaults to sender of Current email
GetFirstName(name := "") {
    If (name = "") {
        name := GetStandardName()
    }
    Return RegExReplace(name, "\s.*")
}

; Extracts First and Last initial from a passed in name
GetInitials(name) {
    Return RegExReplace(GetStandardName(name), "(?!\b\w).")
}

; Trims out the contents of an email after the name provided
; If no sender provided, passes empty string internally
; If no email provided, uses current selected email in outlook
GetEmailBody(email := "", name := "") {
    whitespace := " `t`n`r"
    regexstr := "s)(?=From:|" . GetFirstName(name) . "|" . GetStandardName(name) . "|"

    If (email = "") {
        email := GetCurrentEmail().Body
    }

    If (name = "") {
        regexstr := regexstr . GetCurrentSender() . ").*"
    } Else {
        regexstr := regexstr . name . ").*"
    }
    cleanstr := RegExReplace(email, regexstr) ; Initial clearance of anything past first detection of full name for basic sig removal
    cleanstr := RegExReplace(cleanstr, "\h+", " ") ; Cleaning of extraneus horizontal whitespace
    cleanstr := RegExReplace(cleanstr, "\v+", "`n") ; Cleaning of extraneous vertical whitespace
    cleanstr := Trim(cleanstr, whitespace) ; Removes trailing/preceding tabs, spaces, and returns.
    cleanstr := RegExReplace(cleanstr, "\s{2,}", "`n`n") ; Handling bulk mixed whitespace.
    cleanstr := RegExReplace(cleanstr, "([!+#^{}])", "`{${1}`}") ; Wrap AHK symbols for safe pasting
    Return cleanstr
}

; Sets the Category of the email to mark that it was handled by me, and marks as read
SetAsHandled() {
    email := GetCurrentEmail()
    If (email.Categories = "") {
        email.Categories := olCat
    } Else If (!InStr(email.Categories, olCat)) {
        email.Categories := email.Categories . ", " . olCat
    }
    email.UnRead := False
    email.Save
    Return
}

; Returns a random boilerplate email greeting
GenerateGreeting() {
    greetings := ["", "Hi", "Hey", "Good $time"]
    hour := FormatTime(, "H")
    If (hour >= 7) and (hour < 12) {
        timestr := "morning"
    } Else If (hour >= 12) and (hour < 18) {
        timestr := "afternoon"
    } Else {
        timestr := "day" ; In case hour is outside of normal operating hours
    }
    idx := Random(1, greetings.Length)
    greet := greetings[idx]
    If (greet != "") {
        greet := greet . " "
    }
    Return StrReplace(greet, "$time", timestr)
}