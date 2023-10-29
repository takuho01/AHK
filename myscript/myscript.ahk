;; 2022/05/21 Taku Honda

;--- config ---

; 変数定義は一番最初で実行しないと、謎のエラーが起きる
global hoge := 1

;;; general variable ;;;
global win01 := 1839474
global win02 := 10881984
global win03 := 56365772
global win04 := 1836078
global win05 := 124327408
global win06 := 724632
global win07 := 0
global win08 := 0
global winsel := 1
global array := [10881984,56365772,1836078,124327408,724632]
global win_index := 1
global array_big := [1839474,1248898,19468538]
global win_index_big := 1
global switch_big := 0


;;; free plane ;;;
global fp_mode := 0

;;; IME change function ;;;
IME_GET(WinTitle="A")  {
    ControlGet,hwnd,HWND,,,%WinTitle%
    if    (WinActive(WinTitle))    {
        ptrSize := !A_PtrSize ? 4 : A_PtrSize
        VarSetCapacity(stGTI, cbSize:=4+4+(PtrSize*6)+16, 0)
        NumPut(cbSize, stGTI,  0, "UInt")   ;    DWORD   cbSize;
        hwnd := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
                 ? NumGet(stGTI,8+PtrSize,"UInt") : hwnd
    }

    return DllCall("SendMessage"
          , UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hwnd)
          , UInt, 0x0283  ;Message : WM_IME_CONTROL
          ,  Int, 0x0005  ;wParam  : IMC_GETOPENSTATUS
          ,  Int, 0)      ;lParam  : 0
}

IME_SET(SetSts, WinTitle="A")    {
    ControlGet,hwnd,HWND,,,%WinTitle%
    if    (WinActive(WinTitle))    {
        ptrSize := !A_PtrSize ? 4 : A_PtrSize
        VarSetCapacity(stGTI, cbSize:=4+4+(PtrSize*6)+16, 0)
        NumPut(cbSize, stGTI,  0, "UInt")   ;    DWORD   cbSize;
        hwnd := DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
                 ? NumGet(stGTI,8+PtrSize,"UInt") : hwnd
    }

    return DllCall("SendMessage"
          , UInt, DllCall("imm32\ImmGetDefaultIMEWnd", Uint,hwnd)
          , UInt, 0x0283  ;Message : WM_IME_CONTROL
          ,  Int, 0x006   ;wParam  : IMC_SETOPENSTATUS
          ,  Int, SetSts) ;lParam  : 0 or 1
}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; basic key setting ;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; cursor move ;;;
    ^Up::Send    {PgUp}
    ^Down::Send  {PgDn}
    ^Left::Send  {Home}
    ^Right::Send {End}
    ^+Up::Send    +{PgUp}
    ^+Down::Send  +{PgDn}
    ^+Left::Send  +{Home}
    ^+Right::Send +{End}

;---2 times push key sumple---
; F13::
;     Keywait, F13, U
;     Keywait, F13, D T0.2
;     if (ErrorLevel=1){
;         MsgBox, a  
;     }else{
;         MsgBox, b
;     }
;     return

;---num pad---
    ; Numpad0::
    ; Numpad1::
    ; Numpad2::
    ; Numpad3::
    ; Numpad4::
    ; Numpad5::
    ; Numpad6::
    ; Numpad7::
    ; Numpad8::
    ; Numpad9::

;---other---
    ; sc079 : henkan
    ; sc07B : muhenkan


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; window layer ;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

winswitch(){
    global array

    if (switch_big == 1){
        switch_big := 0
    }else{
        win_index := win_index + 1
        if (win_index > array.MaxIndex()){
            win_index := 1
        }
    }
    win_id_tmp := array[win_index]
    WinActivate, ahk_id %win_id_tmp%
    switch_big := 0
}
winswitch_rev(){
    global array

    if (switch_big == 1){
        switch_big := 0
    }else{
        win_index := win_index - 1
        if (win_index < 1){
            win_index := array.MaxIndex()
        }
    }
    win_id_tmp := array[win_index]
    WinActivate, ahk_id %win_id_tmp%
    switch_big := 0
}
winswitch_big(){
    global array_big
    global switch_big
    if (switch_big == 0){
        win_id_tmp := array_big[win_index_big]
        WinActivate, ahk_id %win_id_tmp%
        switch_big := 1
    }else{
        if (win_index_big == array_big.MaxIndex()){
            win_index_big := 1
        }else{
            win_index_big := win_index_big + 1
        }
        win_id_tmp := array_big[win_index_big]
        WinActivate, ahk_id %win_id_tmp%
    }
}
swithc_big_base(){
    if (switch_big == 0){
        win_id_tmp := array_big[win_index_big]
        WinActivate, ahk_id %win_id_tmp%
        switch_big := 1
    }else if (switch_big == 1){
        win_id_tmp := array[win_index]
        WinActivate, ahk_id %win_id_tmp%
        switch_big := 0
    }
}
sc079 & 2::winswitch()
sc079 & 1::winswitch_rev()
sc079 & tab::swithc_big_base()
sc079 & q::winswitch_big()

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Onenote ;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe ONENOTE.EXE
        ^0::
            send, ^+n
            return
        ^1::
            send, ^!1
            send, ^8
            return
        ^2::
            send, ^!2
            send, ^9
            return
        ^3::
            send, ^!3
            return
        ^q::
            send, !w
            Sleep, 100
            send, t
            Return
        ^+Enter::
            ; set_insert_mode()
            send, !j
            Sleep, 100
            send, l
            Sleep, 100
            send, v
            ; set_normal_mode()
            Return
        ^BackSpace::
            send, !j
            Sleep, 100
            send, l
            Sleep, 100
            send, w
            Return
        Numpad0:: 
            send, {Alt}
            Sleep, 100
            send, w
            Sleep, 100
            send, z
            Sleep, 100
            send, w
            Sleep, 100
            send, o
            Return
    #IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Excel ;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe EXCEL.exe
        ^Enter::
            send {F2}
            return
        +Enter::
            send !{Enter}
            return
        sc07B & Enter::
            send {Down}
            Sleep, 100
            send +{Space}
            Sleep, 100
            send ^+;
            Sleep, 100
            send {Up}{Down}
            return
        ^BackSpace::
            send +{Space}
            send ^-
            send {Up}{Down}
            return
        ^Up::Send    ^{Up}
        ^Down::Send  ^{Down}
        ^Left::Send  ^{Left}
        ^Right::Send ^{Right}
        ^+Up::Send    ^+{Up}
        ^+Down::Send  ^+{Down}
        ^+Left::Send  ^+{Left}
        ^+Right::Send ^+{Right}
        +WheelUp:: send +^{WheelUp}
        +WheelDown:: send +^{WheelDown}
        ^Tab:: send !6
        ^+Tab:: send !7
        
    #IfWinActive
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Power Point ;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe POWERPNT.EXE
        ^Enter::
            send !2
            send !{Up}
            send {down}
            return
        +^Enter::
            send !1
            send !{Up}
            send {down}
            return
        ^BackSpace::
            send !5
            send {Left}{Right}
            return
        Tab:: send !8
        +Tab:: send !7
        ^Up::Send    ^{Up}
        ^Down::Send  ^{Down}
        ^Left::Send  ^{Left}
        ^Right::Send ^{Right}
        ^+Up::Send    ^+{Up}
        ^+Down::Send  ^+{Down}
        ^+Left::Send  ^+{Left}
        ^+Right::Send ^+{Right}
    #IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; edge ;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    ; #IfWinActive,xlsx ahk_exe msedge.exe
    #IfWinActive, xls ahk_exe msedge.exe
                ^Enter:: send, {F2}
                sc07B & Enter::
                    send {Down}
                    Sleep, 100
                    send +{Space}
                    Sleep, 100
                    send ^+;
                    Sleep, 300
                    send {Up}{Down}
                    return
                ^BackSpace::
                    send +{Space}
                    Sleep, 100
                    send ^-
                    Sleep, 300
                    send {Up}{Down}
                    return
                +Enter::
                    send !{Enter}
                    return
                ^+Enter::
                    send +{Enter}
                    return
    #IfWinActive


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; freeplane ;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
#IfWinActive, ahk_class SunAwtFrame
    ^Enter:: 
        send, {F2}
        fp_mode := 1
        return
    ^+c::send, ^c 
    ^+v::send, ^v
    i::
        if (fp_mode==0){
            send, {F2}
            fp_mode := 1
        }else if (fp_mode==1){
            send, i
        }else if (fp_mode==2){
            send, i
        }
        return 
    h::
        if (fp_mode==0){
            send, {left}
        }else if (fp_mode==1){
            send, h 
        }else if (fp_mode==2){
            send, h 
        }
        return 
    j::
        if (fp_mode==0){
            send, {down}
        }else if (fp_mode==1){
            send, j 
        }else if (fp_mode==2){
            send, j 
        }
        return 
    k::
        if (fp_mode==0){
            send, {up}
        }else if (fp_mode==1){
            send, k 
        }else if (fp_mode==2){
            send, k 
        }
        return 
    l::
        if (fp_mode==0){
            send, {right}
        }else if (fp_mode==1){
            send, l 
        }else if (fp_mode==2){
            send, l 
        }
        return 
    x::
        if (fp_mode==0){
            send, ^x
        }else if (fp_mode==1){
            send, x 
        }else if (fp_mode==2){
            send, x 
        }
        return 
    y::
        if (fp_mode==0){
            send, ^c 
        }else if (fp_mode==1){
            send, y 
        }else if (fp_mode==2){
            send, y 
        }
        return 
    p::
        if (fp_mode==0){
            send, ^v 
        }else if (fp_mode==1){
            send, p 
        }else if (fp_mode==2){
            send, p 
        }
        return 
    Enter::
        send, {Enter}
        fp_mode := 1
        return
    ^c::
        if (fp_mode==0){
            send, {Esc}
            fp_mode := 0
        }else if (fp_mode==1){
            send, {Esc}
            fp_mode := 0
        }else if (fp_mode==2){
            send, {Enter}
            ; send, {Esc}
            fp_mode := 0
        }
        return 
    u::
        if (fp_mode==0){
            send, ^z
        }else if (fp_mode==1){
            send, u 
        }else if (fp_mode==2){
            send, u
        }
        return 
    ^r::
        if (fp_mode==0){
            send, ^y
        }else if (fp_mode==1){
            send, ^r
        }else if (fp_mode==2){
            send, ^r 
        }
        return 
    Tab::
        send, {Tab}
        fp_mode := 1 
        return
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; logseq ;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive, ahk_exe Logseq.exe
    ^c::
        send {esc}
        return
    #IfWinActive

