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
global array1 := [ ]
global array2 := [ ]
global win_index1 := 1
global win_index2 := 1
global array_big := [ ]
global win_index_big := 1
global switch_big := 0
global toggle_array1 := 0
global open_window_id := 0
global arr1_posX := 30
global arr1_posY := 30
global arr1_width := 1500
global arr1_height := 1000

browser_id := 2165268
terminal_id := 56365772
mindmap_id := 1248898
onenote_id := 124327408
logseq_id := 1836078
excel_id := 724632
; ppt_id := 
; array.Push(browser_id)
; array.Push(terminal_id)
; array.Push(mindmap_id)
; array.Push(onenote_id)
; array.Push(logseq_id)
; array.Push(excel_id)

vscode_id := 1839474
; hoge_id := 10203040
array_big.Push(vscode_id)


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

; reset_id(){
;     global array
;     global array_big
;     array.Clear()
;     array_big.Clear()
    
;     browser_id := 11078592
;     terminal_id := 56365772
;     mindmap_id := 1248898
;     onenote_id := 124327408
;     excel_id := 724632
;     ; ppt_id := 
;     array.Push(browser_id)
;     array.Push(terminal_id)
;     array.Push(mindmap_id)
;     array.Push(onenote_id)
;     array.Push(excel_id)

;     vscode_id := 1839474
;     ; hoge_id := 10203040
;     array_big.Push(vscode_id)
; }

show(){
    WinGet, active_id, ID, A ; 'A' はアクティブなウィンドウを意味する
    MsgBox, %active_id%
}

edit_array1(){
    global array1
    MouseGetPos, , , active_id
    push_flag := 1
    for index, value in array1{
        val := array1[index]
        if (active_id == val){
            array1.RemoveAt(index)
            MsgBox, 4, , removed, 0.5
            push_flag := 0
        }
    }
    if (push_flag == 1){
        array1.Push(active_id)
        MsgBox, 4, , inserted, 0.5
    }
}
edit_array2(){
    global array2
    MouseGetPos, , , active_id
    push_flag := 1
    for index, value in array2{
        val := array2[index]
        if (active_id == val){
            array2.RemoveAt(index)
            MsgBox, 4, , removed, 0.5
            push_flag := 0
        }
    }
    if (push_flag == 1){
        array2.Push(active_id)
        MsgBox, 4, , inserted, 0.5
    }
}

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
winswitch_fixedpos1(){
    global array1
    win_index1 := win_index1 + 1
    if (win_index1 > array1.MaxIndex()){
        win_index1 := 1
    }
    WinGetPos, posX, posY, width, height, A
    win_id_tmp := array1[win_index1]
    if (win_index1 == 1){
        WinMove, ahk_id %win_id_tmp%,, posX-200*(array1.MaxIndex()-1), posY, width, height
    }else{
        WinMove, ahk_id %win_id_tmp%,, posX+200, posY, width, height
    }
    WinActivate, ahk_id %win_id_tmp%
}
winswitch_fixedpos1_rev(){
    global array1
    win_index1 := win_index1 - 1
    if (win_index1 == 0){
        win_index1 := array1.MaxIndex()
    }
    WinGetPos, posX, posY, width, height, A
    win_id_tmp := array1[win_index1]
    if (win_index1 == 1){
        WinMove, ahk_id %win_id_tmp%,, posX-100*(array1.MaxIndex()-1), posY, width, height
    }else{
        WinMove, ahk_id %win_id_tmp%,, posX+100, posY, width, height
    }
    WinActivate, ahk_id %win_id_tmp%
}
winswitch_fixedpos2(){
    global array2
    win_index2 := win_index2 + 1
    if (win_index2 > array2.MaxIndex()){
        win_index2 := 1
    }
    WinGetPos, posX, posY, width, height, A
    win_id_tmp := array2[win_index2]
    if (win_index2 == 1){
        WinMove, ahk_id %win_id_tmp%,, posX-30*(array2.MaxIndex()-1), posY, width, height
    }else{
        WinMove, ahk_id %win_id_tmp%,, posX+30, posY, width, height
    }
    WinActivate, ahk_id %win_id_tmp%
}
winswitch_fixedpos2_rev(){
    global array2
    win_index2 := win_index2 - 1
    if (win_index2 == 0 ){
        win_index2 := array2.MaxIndex()
    }
    WinGetPos, posX, posY, width, height, A
    win_id_tmp := array2[win_index2]
    if (win_index2 == 1){
        WinMove, ahk_id %win_id_tmp%,, posX-30*(array2.MaxIndex()-1), posY, width, height
    }else{
        WinMove, ahk_id %win_id_tmp%,, posX+30, posY, width, height
    }
    WinActivate, ahk_id %win_id_tmp%
}
check_array(){
    WinGet, active_id, ID, A ; 'A' はアクティブなウィンドウを意味する
    for index, value in array1{
        val := array1[index]
        if (active_id == val){
            winswitch_fixedpos1()
            break
        }
    }
    for index, value in array2{
        val := array2[index]
        if (active_id == val){
            winswitch_fixedpos2()
            break
        }
    }
}
check_array_rev(){
    WinGet, active_id, ID, A ; 'A' はアクティブなウィンドウを意味する
    for index, value in array1{
        val := array1[index]
        if (active_id == val){
            winswitch_fixedpos1_rev()
            break
        }
    }
    for index, value in array2{
        val := array2[index]
        if (active_id == val){
            winswitch_fixedpos2_rev()
            break
        }
    }
}

listup1(){
    WinGetPos, posX, posY, width, height, ahk_id %open_window_id%
    if (width>500 && height > 500){
        arr1_posX := posX
        arr1_posY := posY
        arr1_width := width
        arr1_height := height
    }
    for index, value in array1{
        win_id_tmp := array1[index]
        WinMove, ahk_id %win_id_tmp%,, 0+(index-1)*300, 0, 500, 500
        WinActivate, ahk_id %win_id_tmp%
    }
    toggle_array1 := 1
}

listup(){
    MouseGetPos, , , mouse_id
    for index, value in array1{
        val := array1[index]
        if (mouse_id == val){
            WinMove, ahk_id %mouse_id%,, arr1_posX, arr1_posY, arr1_width, arr1_height
            WinActivate, ahk_id %mouse_id%
            open_window_id := mouse_id
            toggle_array1 := 0
            break
        }
    }
    if (toggle_array1 == 1){
        listup1()
    }
}

switch_array1(){
    if (toggle_array1 == 0){
        listup1()
    }else{
        listup()
    }
}


sc07B & o::winswitch_big()
sc07B & k::winswitch()
sc07B & j::winswitch_rev()

sc07B & 1::edit_array1()
sc07B & 2::edit_array2()
sc07B & LButton::switch_array1()
sc07B & RButton::edit_array1()
sc07B & tab::switch_array1()
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Onenote ;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe ONENOTE.EXE
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
        ^Up::Send    {PgUp}
        ^Down::Send  {PgDn}
        ^Left::Send  {Home}
        ^Right::Send {End}
        ^+Left::Send  +{Home}
        ^+Right::Send +{End}
        +WheelUp:: send +^{WheelUp}
        +WheelDown:: send +^{WheelDown}
        ^Tab:: send !6
        ^+Tab:: send !7
        
    #IfWinActive
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Power Point ;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe POWERPNT.EXE
        ; ^Entter::
        ;     send !2
        ;     send !{Up}
        ;     send {down}
        ;     return
        ; +^Enter::
        ;     send !1
        ;     send !{Up}
        ;     send {down}
        ;     return
        ; ^BackSpace::
        ;     send !5
        ;     send {Left}{Righ}
        ;     return
        Tab:: send !2
        +Tab:: send !1
        ^Up::Send    {PgUp}
        ^Down::Send  {PgDn}
        ^Left::Send  {Home}
        ^Right::Send {End}
        ^+Left::Send  +{Home}
        ^+Right::Send +{End}
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
    ^Enter:: send, {F2}
#IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; logseq ;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive, ahk_exe Logseq.exe
    ^c::
        send {esc}
        return
    #IfWinActive

