#Requires AutoHotkey v2.0
;----------------------------------------------------------------------------------------
A := InputBox("Please enter your OACIS username:","Username").value
B := InputBox("Please enter your OACIS password","Password", "password").value

;EMR patient launcher
^!+o::{

;Copy MRN
Send "^c"
ClipWait
mrn := A_Clipboard

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
Send "{Esc}"
Sleep 50
Send "{Esc}"

;Check if logged in, if not then log in
if (PixelGetColor(100, 40) != 000000){
    MouseClick "left", 710, 290
    Sleep 100
    SendText A
    Sleep 100
    Send "{Tab}"
    Sleep 100
    SendText B
    Sleep 100
    Send "!l"
    Sleep 1750
}

;Open single patient lookup
MouseClick "left", 300, 70
Sleep 500

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
Sleep 500
MouseClick "left", 130, 140, 2
Return
}

