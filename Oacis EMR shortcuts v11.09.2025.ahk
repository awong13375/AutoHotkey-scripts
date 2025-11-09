#Requires AutoHotkey v2.0
;----------------------------------------------------------------------------------------

; Message box with tutorial
text := "Script is designed to improve your workflow efficiency and reduce the number of clicks on OACIS.`n`n"
    . "REQUIREMENTS: Please have Inteleviewer opened to the specific patient and study, and are logged into OACIS.`n`n"
    . "IMPORTANT: Do not touch keyboard or mouse while shortcut is running.`n`n"
    . "SHORTCUTS:`n"
    . "Ctrl+Shift+O = Retrieves patient file and opens labs by default (can be changed, see below)`n"
    . "Ctrl+Shift+I = To show these instructions again at any time`n`n"
    . "Additional shortcuts require patient file to already be retrieved on Oacis (using Ctrl+Shift+O above)`n"
    . "Ctrl+Shift+D = Opens Documents viewer`n"
    . "Ctrl+Shift+P = Opens Pathology`n"
    . "Ctrl+Shift+L = Opens Labs`n"
    . "Ctrl+Shift+S = Opens OPERA surgical procedures`n`n`n"
    . "Additional script settings:`n"
    . "Ctrl+Alt+Shift+O = To change default window that opens on OACIS using the Ctrl+Shift+O shortcut`n"
    . "Ctrl+Alt+Shift+H = To set script for remote OACIS access from home via Citrix versus onsite access (default onsite access)`n`n`n"
    . "Handsfree dictation:`n"
    . "Backwards apostrophe (left of '1' on keyboard) = toggle dictation on/off on Powerscribe`n`n`n"
    . "To extend OACIS timeout from logging out every 5 minutes:`n"
    . "Send email to sabina.chowdury@muhc.mcgill.ca and tell her to extend OACIS timeout to 10 hours. You will need to provide her the Term. ID (MUHC-AA###AAA## or PC name if using home computer), found on bottom right of OACIS.`n`n`n"
    . "Script creator: Alexander Wong.`n`n"
    . "Version Nov 9, 2025."

MsgBox text, "INSTRUCTIONS"

; Manual input username and password for OACIS
;A := InputBox("Please enter your OACIS username:","Username").value
;B := InputBox("Please enter your OACIS password","Password", "password").value

; Keep OACIS from idle log off
;timestamp1 := A_TickCount
;while WinActive("ahk_exe java.exe"){
;    timestamp1 := A_TickCount
;}
;while !WinActive("ahk_exe java.exe"){
;    timeinactive := A_Tickcount - timestamp1
;    if (timeinactive > 180000 ){
;        ControlSend("{Esc}", WinTitle := "ahk_exe java.exe")
;    }
;}

; To show instructions again
^+i::{
    MsgBox text, "INSTRUCTIONS"
}

; Re-enter OACIS username/password
;^+q::{
;    global A
;    global B
;
;    MsgBox "Stored username is: " . A . "`n`nStored password is: " . B , "Verify and re-enter username and password"
;
;    A := InputBox("Please enter your OACIS username:", "Username").value
;    B := InputBox("Please enter your OACIS password:", "Password", "password").value
;}

; To set default OACIS window to open using Ctrl+Shift+O
A:= 1
^!+o::{
    global A
    A:= InputBox("Please enter number corresponding to default OACIS window to open using Ctrl+Shift+O`n"
    . "1 = Labs, 2 = Documents viewer, 3 = Pathology, 4 = Surgical history"
    ).value
}

H:= 1
^!+h::{
    global H
    H:= InputBox("Please enter number corresponding to remote OACIS access from home (via Citrix) vs onsite (natively installed version)`n"
    . "1 = Onsite, 2 = Home remote access via Citrix"
    ).value
}

;EMR patient launcher
^+o::{

MyGui := Gui(, "Script in progress")
MyGui.Opt("+AlwaysOnTop +Disabled -SysMenu +Owner")  ; +Owner avoids a taskbar button.
MyGui.Add("Text",, "Please do not touch keyboard or mouse while script is active.")
MyGui.Show("NoActivate")  ; NoActivate avoids deactivating the currently active window.


;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MyGui.Destroy()
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if H==1{

    if not WinExist("ahk_exe java.exe"){
        MyGui.Destroy()
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

} else if H==2{

    if not WinExist("ahk_exe wfica32.exe"){
        MyGui.Destroy()
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

} else {

    if not WinExist("ahk_exe java.exe"){
        MyGui.Destroy()
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

}

;Copy MRN from PACS
;Ensure activation of search tool window
HWNDs := WinGetList("ahk_exe InteleViewer.exe")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "Search Tool"){
        WinActivate(id)
        Sleep 500
        WinActivate(id)
    }
} 
Sleep 100

A_Clipboard := ""
A_Clipboard := ""
MouseClick "left", 50, 105, 2
Sleep 100
Send "^c"
Sleep 10
Send "^c"
Sleep 10
Send "^c"
if( !Clipwait(1,1) ){
    MyGui.Destroy()
    MsgBox "Please select active patient in InteleViewer Search Tool and try again.", "Script error"
    Return
}

mrn := A_Clipboard

if H==1 {

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 400
            WinActivate (id)
        }
    }

} else if H==2 {

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe wfica32.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 400
            WinActivate (id)
        }
    }

} else {

;Ensure activation of vOACIS
HWNDs := WinGetList("ahk_exe java.exe")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "OACIS"){
        WinActivate(id)
        Sleep 400
        WinActivate (id)
    }
}

}

Sleep 500
Send "{Esc}"
Sleep 50
Send "{Esc}"
Sleep 50
Send "!d"
Sleep 500
Send "{Tab}"
Sleep 50
Send "{Tab}"
Sleep 50
Send "{Enter}"

;Open patient lookup
Sleep 750
Send "{F10}"
Sleep 50
Send "{Right}"
Sleep 50
Send "!c"
Sleep 500

;for LAC MRNs
if (InStr(mrn, "L"))!= 0 {
    pos := InStr(mrn, "L")
    NewStr := SubStr(mrn, 1, -1)
    Send NewStr
    Sleep 100
    Send "{Tab}"
    Sleep 50  
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Space}"
}


;for MCH MRNs
if (InStr(mrn, "MCH"))!= 0 {
    pos := InStr(mrn, "MCH")
    NewStr := SubStr(mrn, 1, -3)
    Send NewStr
    Sleep 100
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Space}"
}


;for MGH MRNs
if (InStr(mrn, "MGH"))!= 0 {
    pos := InStr(mrn, "MGH")
    NewStr := SubStr(mrn, 1, -3)
    Send NewStr
    Sleep 100
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Space}"
}


;for RVH MRNs
if (InStr(mrn, "RVH"))!= 0 {
    pos := InStr(mrn, "RVH")
    NewStr := SubStr(mrn, 1, -3)
    Send NewStr
    Sleep 100
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Space}"
}


;for MNH MRNs (all under RVH database)
if (InStr(mrn, "V"))!= 0 {
    pos := InStr(mrn, "V")
    NewStr := SubStr(mrn, 1, -1)
    Send NewStr
    Sleep 100
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
    Sleep 50
    Send "{Tab}"
    Sleep 50
    Send "{Space}"
}

Sleep 200
Send "!d"
Sleep 100
Send "!o"
Sleep 750
Send "+{Tab}"
Sleep 50
Send "+{Tab}"
Sleep 50

; for default OACIS window to open
; 1 = Labs
; 2 = Documents viewer
; 3 = Pathology
; 4 = Surgical history
if A == 1 {
    Send "!r"
    Sleep 50
    Send "!l"
} else if A == 2 {
    Send "!r"
    Sleep 50
    Send "!d"
} else if A == 3 {
    Send "!r"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Down}"
    Sleep 50
    Send "{Enter}"
} else if A == 4 {
    Send "{F10}"
    Sleep 50
    Send "{Right}"
    Sleep 50
    Send "{Right}"
    Sleep 50
    Send "{Right}"
    Sleep 50
    Send "{Up}"
    Sleep 50
    Send "{Enter}"
} else {
    Send "!r"
    Sleep 50
    Send "!l"
}

A_Clipboard := ""
A_Clipboard := ""

MyGui.Destroy()
Return

}

;----------------------------------------------------------------------------------------
;Open Documents Viewer
^+d::{

;Check if logged in and patient selected
;if (PixelGetColor(42, 12) != 000000){
;    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
;    Return
;}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if H==1 {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

} else if H==2 {

    if not WinExist("ahk_exe wfica32.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe wfica32.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

} else {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

}

Sleep 100

Send "!r"
Sleep 50
Send "!d"

Return
}


;----------------------------------------------------------------------------------------
;Open Labs
^+l::{

;Check if logged in and patient selected
;if (PixelGetColor(42, 12) != 000000){
;    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
;    Return
;}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if H==1 {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }


} else if H==2 {

    if not WinExist("ahk_exe wfica32.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe wfica32.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

} else {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

}

Sleep 100
Send "!r"
Sleep 50
Send "!l"
Return
}

;----------------------------------------------------------------------------------------
;Open Pathology
^+p::{

;Check if logged in and patient selected
;if (PixelGetColor(42, 12) != 000000){
;    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
;    Return
;}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if H==1 {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }


} else if H==2 {

    if not WinExist("ahk_exe wfica32.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe wfica32.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

} else {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

}

Sleep 100
Send "!r"
Sleep 50
Send "{Down}"
Sleep 50
Send "{Down}"
Sleep 50
Send "{Down}"
Sleep 50
Send "{Down}"
Sleep 50
Send "{Down}"
Sleep 50
Send "{Enter}"

Return
}

;----------------------------------------------------------------------------------------
;Open Surgical history
^+s::{

;Check if logged in and patient selected
;if (PixelGetColor(42, 12) != 000000){
;    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
;    Return
;}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if H==1 {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }


} else if H==2 {

    if not WinExist("ahk_exe wfica32.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe wfica32.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

} else {

    if not WinExist("ahk_exe java.exe"){
        MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
        Return
    }

    ;Ensure activation of vOACIS
    HWNDs := WinGetList("ahk_exe java.exe")
    For id in HWNDs{
        title := WinGetTitle(id)
        if InStr(title, "OACIS"){
            WinActivate(id)
            Sleep 500
            WinActivate (id)
        }
    }

}

Sleep 100
Send "{F10}"
Sleep 50
Send "{Right}"
Sleep 50
Send "{Right}"
Sleep 50
Send "{Right}"
Sleep 50
Send "{Up}"
Sleep 50
Send "{Enter}"
Return
}

;----------------------------------------------------------------------------------------
;Powerscribe dictation shortcuts

`::{

if not WinExist("ahk_exe Nuance.PowerScribe360.exe"){
    MsgBox "Powerscribe is not running, please launch Powerscribe first and try again.", "Script error"
    Return
}

WinExist("A")

WinActivate ("ahk_exe Nuance.PowerScribe360.exe")
Sleep 10
Send "{F4}"
Sleep 10

WinActivate

}
