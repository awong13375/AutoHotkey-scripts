#Requires AutoHotkey v2.0
;----------------------------------------------------------------------------------------

; Auto shutdown script after 18 hours to avoid other people using login
SetTimer AutoShutdown, 64800000
AutoShutdown()
{
    MsgBox "Oacis EMR script has automatically terminated after running for an extended period of time. Please relaunch the script and re-enter your username and password.", "Oacis EMR script Auto-shutdown."
    ExitApp
}

; Message box with tutorial
text := "Script is designed to improve your workflow efficiency and reduce the number of clicks on OACIS.`n`n"
    . "REQUIREMENTS: Please have Inteleviewer opened to the specific patient and study, and OACIS opened.`n`n"
    . "IMPORTANT: Do not touch keyboard or mouse while shortcut is running.`n`n"
    . "SHORTCUTS:`n"
    . "Ctrl+Shift+O = Login to Oacis (if not already logged in), retrieves patient file and opens labs`n"
    . "Ctrl+Shift+I = To show these instructions again at any time`n"
    . "Ctrl+Shift+Q = To re-enter OACIS username/password if typo made on first try`n`n"
    . "The below shortcuts require patient to already be retrieved on Oacis (using Ctrl+Shift+O above)`n"
    . "Ctrl+Shift+D = Opens Documents viewer`n"
    . "Ctrl+Shift+P = Opens Pathology`n"
    . "Ctrl+Shift+L = Opens Labs`n"
    . "Ctrl+Shift+S = Opens OPERA surgical procedures`n`n`n"
    . "For those with handsfree dictation setups:`n"
    . "Backwards apostrophe (left of '1' on keyboard) = toggle dictation on/off on Powerscribe`n`n`n"
    . "Script creator: Alexander Wong.`n`n"
    . "Version: 1.0, released Apr 5, 2025."

MsgBox text, "INSTRUCTIONS"

; Manual input username and password for OACIS
A := InputBox("Please enter your OACIS username:","Username").value
B := InputBox("Please enter your OACIS password","Password", "password").value

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
^+q::{
    global A
    global B

    MsgBox "Stored username is: " . A . "`n`nStored password is: " . B , "Verify and re-enter username and password"

    A := InputBox("Please enter your OACIS username:", "Username").value
    B := InputBox("Please enter your OACIS password:", "Password", "password").value
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

if not WinExist("ahk_exe java.exe"){
    MyGui.Destroy()
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
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
Sleep 100
Send "{Esc}"

;Define Pastetext function to paste user/pass
Pastetext(text){
    ClipSaved := ClipboardAll()
    A_Clipboard := text
    Send "^v"
    Sleep 100
    A_Clipboard := ""
    Sleep 100
    A_Clipboard := ClipSaved
}

;Check if logged in, if not then log in

Sleep 150
Send "!p"
Sleep 150
Pastetext(B)
Sleep 250
Send "+{Tab}"
Sleep 100
Pastetext(A)
Sleep 250
Send "!l"
Sleep 2000 

while (PixelGetColor(42, 12) != 000000){
    Send "{Esc}"
    Send "{Esc}"

    if (PixelGetColor(42, 12) == 000000){
        Break
    } 

    Sleep 150
    Send "!p"
    Sleep 150 
    Send "^a"
    Sleep 100
    Send "{Backspace}"
    Sleep 150
    Pastetext(B)
    Sleep 250
    Send "+{Tab}"
    Sleep 150
    Send "^a"
    Sleep 100
    Send "{Backspace}"
    Sleep 150
    Pastetext(A)
    Sleep 250
    Send "!l"
    Sleep 1500    

    if (A_Index > 5){
        MyGui.Destroy()
        MsgBox "Please restart script and re-enter correct username and password.", "Incorrect Username/Password"
        global A
        global B
    
        MsgBox "Stored username is: " . A . "`n`nStored password is: " . B , "Verify and re-enter username and password"
    
        A := InputBox("Please enter your OACIS username:", "Username").value
        B := InputBox("Please enter your OACIS password:", "Password", "password").value
        
        Return
    }
}

;Open single patient lookup
MouseClick "left", 291, 39
Sleep 750

; Make sure MRN search is selected
;CoordMode "Pixel", "Client"
;if (PixelGetColor(21, 54) != 333333){
;        Send "+{Tab}"
;	Sleep 150
;	while (PixelGetColor(21, 54) != 333333){
;		Send "{Left}"
;		Sleep 250
;	}
;
;	Send "{Tab}"
;} 

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
MouseClick "left", 33, 117
Sleep 700
MouseClick "left", 33, 117
Sleep 50
Send "!r"
Sleep 50
Send "!l"

MyGui.Destroy()
Return

}

;----------------------------------------------------------------------------------------
;Open Documents Viewer
^+d::{

;Check if logged in and patient selected
if (PixelGetColor(42, 12) != 000000){
    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
    Return
}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

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
Sleep 100
MouseClick "left", 440, 42
Sleep 50
MouseClick "left", 440, 42
Return
}


;----------------------------------------------------------------------------------------
;Open Labs
^+l::{

;Check if logged in and patient selected
if (PixelGetColor(42, 12) != 000000){
    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
    Return
}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

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

WinExist("A")
WinActivate ("ahk_exe InteleViewer.exe")
Sleep 50
WinActivate

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
if (PixelGetColor(42, 12) != 000000){
    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
    Return
}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

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

WinExist("A")
WinActivate ("ahk_exe InteleViewer.exe")
Sleep 50
WinActivate

Sleep 100
Send "!r"
Sleep 50
MouseClick "left", 320, 140
Return
}

;----------------------------------------------------------------------------------------
;Open Surgical history
^+s::{

;Check if logged in and patient selected
if (PixelGetColor(42, 12) != 000000){
    MsgBox "Please run 'Ctrl+Shift+O' shortcut first (to select patient on Oacis) before attempting other shortcuts.", "Script error"   
    Return
}

;Check if oacis and Inteleviewer launched, if not send error message
if not WinExist("ahk_exe InteleViewer.exe"){
    MsgBox "Either Inteleviewer and/or Oacis is not running. Please launch both programs before proceeding.", "Script error"
    Return
}

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

WinExist("A")
WinActivate ("ahk_exe InteleViewer.exe")
Sleep 50
WinActivate

Sleep 100
Send "!c"
Sleep 50
Send "!o"
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
