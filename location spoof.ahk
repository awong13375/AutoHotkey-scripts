#Requires AutoHotkey v2.0

SetTimer CheckTime, 3300000

CheckTime(){
if (A_Hour==20){
;Teleport MGH
WinActivate("ahk_class VMPlayerFrame")
Sleep 1000
Send "{Esc}"
Sleep 1000
MouseClick "left", 838, 768
Sleep 1000
Send "{Esc}"
Sleep 1000
Send "{Esc}"
Sleep 2000
MouseClick "left", 838, 768
Sleep 2000
MouseClick "left", 220, 290
Sleep 2000
Send "{LWin}o"
Sleep 2000
;change here for location
MouseClick "left", 375, 240
;
Sleep 2000
Send "{Enter}"
Sleep 2000
MouseClick "left", 505, 505
}

if (A_Hour==08){
;Teleport home
WinActivate("ahk_class VMPlayerFrame")
Sleep 1000
Send "{Esc}"
Sleep 1000
MouseClick "left", 838, 768
Sleep 1000
Send "{Esc}"
Sleep 1000
Send "{Esc}"
Sleep 2000
MouseClick "left", 838, 768
Sleep 2000
MouseClick "left", 220, 290
Sleep 2000
Send "{LWin}o"
Sleep 2000
;change here for location
MouseClick "left", 375, 200
;
Sleep 2000
Send "{Enter}"
Sleep 2000
MouseClick "left", 505, 505
}

;test
if (A_Hour==14){
    ;Teleport JGH
    WinActivate("ahk_class VMPlayerFrame")
    Sleep 250
    Send "{Esc}"
    Sleep 250
    MouseClick "left", 838, 768
    Sleep 250
    Send "{Esc}"
    Sleep 250
    MouseClick "left", 838, 768
    Sleep 250
    MouseClick "left", 220, 290
    Sleep 250
    Send "{LWin}o"
    Sleep 250
    ;change here for location
    MouseClick "left", 375, 220
    ;
    Sleep 250
    Send "{Enter}"
    Sleep 250
    MouseClick "left", 505, 505
}
} 
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;Teleport home
WinActivate("ahk_class VMPlayerFrame")
Sleep 1000
Send "{Esc}"
Sleep 1000
MouseClick "left", 838, 768
Sleep 1000
Send "{Esc}"
Sleep 1000
Send "{Esc}"
Sleep 2000
MouseClick "left", 838, 768
Sleep 2000
MouseClick "left", 220, 290
Sleep 2000
Send "{LWin}o"
Sleep 2000
;change here for location
MouseClick "left", 375 200
;
Sleep 2000
Send "{Enter}"
Sleep 2000
MouseClick "left", 505 505


;Teleport MGH
WinActivate("ahk_class VMPlayerFrame")
Sleep 1000
Send "{Esc}"
Sleep 1000
MouseClick "left", 838, 768
Sleep 1000
Send "{Esc}"
Sleep 1000
Send "{Esc}"
Sleep 2000
MouseClick "left", 838, 768
Sleep 2000
MouseClick "left", 220, 290
Sleep 2000
Send "{LWin}o"
Sleep 2000
;change here for location
MouseClick "left", 375, 240
;
Sleep 2000
Send "{Enter}"
Sleep 2000
MouseClick "left", 505 505



;Teleport JGH
WinActivate("ahk_class VMPlayerFrame")
Sleep 1000
Send "{Esc}"
Sleep 1000
MouseClick "left", 838, 768
Sleep 1000
Send "{Esc}"
Sleep 1000
Send "{Esc}"
Sleep 2000
MouseClick "left", 838, 768
Sleep 2000
MouseClick "left", 220, 290
Sleep 2000
Send "{LWin}o"
Sleep 2000
;change here for location
MouseClick "left", 375, 220
;
Sleep 2000
Send "{Enter}"
Sleep 2000
MouseClick "left", 505 505



