;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; freeplane ;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#IfWinActive, ahk_class SunAwtFrame
    RButton::
        Keywait, RButton, U T0.1
        if (ErrorLevel=1){
            ;; hold
            MouseClick LEFT , , , , , D,
            Keywait, RButton, U
            MouseClick LEFT , , , , , U,
        }else{
            ;; click
            MouseClick RIGHT , , , , , D,
            MouseClick RIGHT , , , , , U,
        }
        return
    wheelup::
        send, ^{WheelUp}
        return
    wheeldown::
        send, ^{WheelDown}
        return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; draw io ;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#IfWinActive, ahk_exe draw.io.exe
wheelup::
    send, ^{WheelUp}
    return
wheeldown::
    send, ^{WheelDown}
    return

#IfWinActive