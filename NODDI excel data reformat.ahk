#Requires AutoHotkey v2.0

^!+v::{
    n := 19
; MNI = 8
; HO_cort = 47
; HO_subcort = 20
; JHU_label = 49
; JHU_tract = 19
    k := 25
    Loop n {
        Send "^{Down}"
        Sleep k
        Send "^+{Right}"
        Sleep k
        Send "^x"
        Sleep k
        Send "^{Up}"
        Sleep k
        Send "^{Right}"
        Sleep k
        Send "{Right}"
        Sleep k
        Send "^v"
        Sleep k
        Send "^{Left}"
        Sleep k
    }
}