#Requires AutoHotkey v2.0
SendMode "Input"
SetWorkingDir A_ScriptDir
CoordMode "Mouse", "Window"
AHKPath := "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
SetTitleMatchMode "RegEx"

; ahk-scripts by Simon Thurston
; 0.0.2 (January 12, 2024)
; Originally Made with love in New Jersey
; Now developed with angst and malice in Pennsylvania

; Loads common settings from the ini file
olCat := IniRead("settings.ini", "scriptconf", "olCat")
sig := StrReplace(IniRead("settings.ini", "scriptconf", "sig"), "``n", "`n")
podUser := IniRead("secrets.ini", "secrets". "podUser")
podPwd := IniRead("secrets.ini", "secrets", "podPwd")

; Grabs the users full name
fullName := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo", "UserName")

; Function Definitions
#Include EmailFuncs.ahk

; General Hotkeys

; Reloads Script
^!+r:: Reload

; Debug Key: Rebind for testing
NumpadDot:: Click

:*:/shrug::¯\_(ツ)_/¯

; VSCodium Hotkeys
#HotIf (WinActive("ahk_exe VSCodium.exe"))
{
	; Generates a Table of Contents using Headers of a Markdown File
	;; Should work fine for general applications but is untested outside of current usecase context
	^!m:: {
		A_Clipboard := ""
		Send("^a^x")
		ClipWait
		instr := RegExReplace(A_Clipboard, "m)^(?!#+.*).*\R")
		instr := RegExReplace(instr, "^.*[^\s]")
		outstr := ""
		Loop Parse instr, "`n"
		{
			RegExReplace(A_LoopField, "#", "#", &rCount)
			rCount -= 1
			tspace := ""
			Loop (%rCount%)
				tspace := tspace . A_Tab
			str := %A_LoopField%
			; Replace headings with bullets
			str := %RegExReplace(str, "#+", tspace . "*")%
			; Place headings into Markdown Link Format
			str := %RegExReplace(str, "(?<=\*\s)(.*)", "[$1](#$1)")%
			; Trim Newlines
			str := %RegExReplace(str, "m)\R")%
			; Condense filenames and hotkeys to links
			str := %RegExReplace(str, "[\.+](?!\S+\])")%
			; Replace spaces with plusses
			str := %RegExReplace(str, "#(\w+)\s(\w+)", "#$1+$2")%
			outstr := outstr . str . "`n"
		}
		out := ""
		RegExMatch(A_Clipboard, "m)^# (?!Table of Contents)(.*\R)*.*", &out)
		A_Clipboard := "# Table of Contents`r`n" . outstr . out
		Send("^v")
		Return
	}
}

; Outlook Hotkeys
#HotIf (WinActive("ahk_exe OUTLOOK.EXE"))
{
	^r:: {
		Send("^r")
		WinWaitActive("RE: ")
		Send("!m{Enter}{Up 3}{Del}")
	}

	; Forwards the email to ServiceNow
	^!+f:: {
		SetAsHandled()
		mailitem := GetCurrentEmail()
		fwdItem := mailItem.Forward
		fwdItem.Recipients.add("cozen@service-now.com")
		fwdItem.Recipients.ResolveAll
		fwdItem.Display
		WinWaitActive("FW:")
		Send("!m{Enter}^{Enter}")
		If (WinWaitActive("Spelling: ", , 3)) {
			Send("{Esc}")
			WinWaitActive("Microsoft Outlook", , , " - Outlook")
			Send("{Enter}")
		}
		Return
	}

	^!+s:: {
		SetAsHandled()
		mailitem := GetCurrentEmail()
		replyItem := mailItem.Reply
		replyItem.Display
		WinWaitActive("RE:")
		Send("!m{Enter}{Up 2}")
		SendText(GetFirstName())
		Send("-{Enter 2}This is spam and can be deleted.")
	}
}

; Slack
;; Pretty much just to revert to the old layout, honestly.
#HotIf (WinActive("ahk_exe Slack.exe"))
{
	^!+s:: {
		A_Clipboard := "localStorage.setItem(`"localConfig_v2`", localStorage.getItem(`"localConfig_v2`").replace(/\`"is_unified_user_client_enabled\`":true/g, `"\`"is_unified_user_client_enabled\`":false`"))"
		Send("^!i")
		WinWaitActive("Developer Tools")
		Sleep(500)
		Send("^v")
		Sleep(500)
		Send("{Enter}")
		Sleep(500)
		Send("!{F4}^r")
	}
}

; Google Chrome
;; Unfortunately, what my company uses. Thanks, javascript.
#HotIf (WinActive("ahk_exe chrome.exe"))
{
	#HotIf (WinActive("Incident"))
	{
		^+s:: SendText(sig)
	}

	#HotIf (WinActive("Proofpoint"))
	{
		^+l:: {
			Send("{Tab}+{Tab}")
			SendText(podUser)
			Send("{Tab}")
			SendText(podPwd)
		}
	}
}