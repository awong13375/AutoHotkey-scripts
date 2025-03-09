#Requires AutoHotkey v2.0
;----------------------------------------------------------------------------------------

; Auto shutdown script after 18 hours to avoid other people using login
SetTimer AutoShutdown, 64800000
AutoShutdown()
{
    MsgBox "Oacis EMR script has automatically terminated after running for an extended period of time. Please relaunch the script and re-enter your username and password.", "Oacis EMR script Auto-shutdown."
    ExitApp
}

; Manual input username and password for Oacis
A := InputBox("Please enter your OACIS username:","Username").value
B := InputBox("Please enter your OACIS password","Password", "password").value

;EMR patient launcher
^+o::{

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if not WinExist("ahk_exe wfica32.EXE"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

;Copy MRN from PACS

;;Ensure activation of search tool window
HWNDs := WinGetList("ahk_exe InteleViewer.exe")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "Search Tool"){
        WinActivate(id)
        Sleep 500
        WinActivate (id)
    }
} 
Sleep 300
MouseClick "left", 50, 105, 2
Sleep 300
Send "^c"
ClipWait
mrn := A_Clipboard

;Ensure activation of vOACIS
HWNDs := WinGetList("ahk_exe wfica32.EXE")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "OACIS"){
        WinActivate(id)
        Sleep 400
        WinActivate (id)
    }
}
Sleep 100
Send "{Esc}"

;Check if logged in, if not then log in
if (PixelGetColor(100, 40) != 000000){
    Sleep 200
    Send "!u"
    Sleep 200
    Send "^a"
    Send "{Backspace}"
    Send "{Ctrl Up}"
    Sleep 200
    ;Send "{text}r30470"
    SendText A
    Sleep 250
    Send "{Tab}"
    Sleep 200
    ;Send "{text}Spongeb0b????"
    SendText B
    Send "{Tab}"
    Sleep 250
    Send "!l"
    Sleep 1250
}

;Open single patient lookup
MouseClick "left", 300, 70
Sleep 750

;Search specific database based on hospital MRN

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


Sleep 50
Send "!s"
Sleep 100
Send "!o"
Sleep 300
MouseClick "left", 130, 140
Sleep 550
MouseClick "left", 130, 140
Sleep 100
MouseClick "left", 450, 70
Return
}

;----------------------------------------------------------------------------------------
;Open Documents Viewer
^+d::{

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if not WinExist("ahk_exe wfica32.EXE"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

;Ensure activation of vOACIS
HWNDs := WinGetList("ahk_exe wfica32.EXE")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "OACIS"){
        WinActivate(id)
        Sleep 500
        WinActivate (id)
    }
}
Sleep 300
MouseClick "left", 450, 70
Sleep 100
MouseClick "left", 450, 70
Return
}


;----------------------------------------------------------------------------------------
;Open Labs
^+l::{

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if not WinExist("ahk_exe wfica32.EXE"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

;Ensure activation of vOACIS
HWNDs := WinGetList("ahk_exe wfica32.EXE")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "OACIS"){
        WinActivate(id)
        Sleep 500
        WinActivate (id)
    }
}
Sleep 300
Send "!r"
Sleep 100
Send "!l"
Return
}

;----------------------------------------------------------------------------------------
;Open Pathology
^+p::{

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if not WinExist("ahk_exe wfica32.EXE"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

;Ensure activation of vOACIS
HWNDs := WinGetList("ahk_exe wfica32.EXE")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "OACIS"){
        WinActivate(id)
        Sleep 500
        WinActivate (id)
    }
}
Sleep 300
Send "!r"
Sleep 200
MouseClick "left", 340, 170
Sleep 200
Return
}

;----------------------------------------------------------------------------------------
;Open Surgical history
^+s::{

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

if not WinExist("ahk_exe wfica32.EXE"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

;Ensure activation of vOACIS
HWNDs := WinGetList("ahk_exe wfica32.EXE")
For id in HWNDs{
    title := WinGetTitle(id)
    if InStr(title, "OACIS"){
        WinActivate(id)
        Sleep 500
        WinActivate (id)
    }
}
Sleep 300
Send "!c"
Sleep 100
Send "!o"
Return
}

