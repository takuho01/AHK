;; 2022/05/21 Taku Honda

;--- config ---
;; for vscode wheel script
#MaxHotkeysPerInterval, 200


;---spjlekcial move---
; 変数定義は一番最初で実行しないと、謎のエラーが起きる
global winpane_on := 1

;;; general variable ;;;
global memo_tgl := 0
global memo_tgl2 := 0

;;; winpane ;;;
global wp_init_flag := 0
global Xrate
global Yrate
global Wrate
global Hrate
global Xmou_rate
global Ymou_rate
global moni_sel
global middle_rate
global xedge := 0.01
global yedge := 0.03
global side_open_xrate := 0.01
global side_open_yrate := 0.01
global side_open_hrate := 0.95
global center_open_yrate := 0.01
global center_open_hrate := 0.95
global midcenter_open_xrate := 0.05
global midcenter_open_yrate := 0.03
global midcenter_open_wrate := 0.9
global midcenter_open_hrate := 0.9
global rclick_resize1_w := 0.15
global rclick_resize1_h := 0.25
global rclick_resize2_w := 0.4
global rclick_resize2_h := 0.5
global rclick_status := 0

global m1_moni_left
global m1_moni_top
global m1_moni_width
global m1_moni_height
global m1_middle_rate := 0.5
global m1_c_buf_x := 0.4
global m1_c_buf_y := 0.4
global m1_c_buf_w := 0.3
global m1_c_buf_h := 0.4
global m1_mc_buf_x := 0.4
global m1_mc_buf_y := 0.3
global m1_mc_buf_w := 0.3
global m1_mc_buf_h := 0.4
global m1_l_buf_x := 0.2
global m1_l_buf_y := 0.4
global m1_l_buf_w := 0.3
global m1_l_buf_h := 0.4
global m1_r_buf_x := 0.6
global m1_r_buf_y := 0.4
global m1_r_buf_w := 0.3
global m1_r_buf_h := 0.4

global m2_moni_left
global m2_moni_top
global m2_moni_width
global m2_moni_height
global m2_middle_rate := 0.5
global m2_c_buf_x := 0.4
global m2_c_buf_y := 0.4
global m2_c_buf_w := 0.3
global m2_c_buf_h := 0.4
global m2_mc_buf_x := 0.4
global m2_mc_buf_y := 0.3
global m2_mc_buf_w := 0.3
global m2_mc_buf_h := 0.4
global m2_l_buf_x := 0.2
global m2_l_buf_y := 0.4
global m2_l_buf_w := 0.3
global m2_l_buf_h := 0.4
global m2_r_buf_x := 0.6
global m2_r_buf_y := 0.4
global m2_r_buf_w := 0.3
global m2_r_buf_h := 0.4

global m3_moni_left
global m3_moni_top
global m3_moni_width
global m3_moni_height
global m3_middle_rate := 0.5
global m3_c_buf_x := 0.4
global m3_c_buf_y := 0.4
global m3_c_buf_w := 0.3
global m3_c_buf_h := 0.4
global m3_mc_buf_x := 0.4
global m3_mc_buf_y := 0.3
global m3_mc_buf_w := 0.3
global m3_mc_buf_h := 0.4
global m3_l_buf_x := 0.2
global m3_l_buf_y := 0.4
global m3_l_buf_w := 0.3
global m3_l_buf_h := 0.4
global m3_r_buf_x := 0.6
global m3_r_buf_y := 0.4
global m3_r_buf_w := 0.3
global m3_r_buf_h := 0.4

;;; keep memo ;;;
global km_mode := 0
global km_pid := 28176
global km_x := 600
global km_y := 300
global km_w := 1200
global km_h := 800
global km2_mode := 0

;;; miro ;;;
global rclick_flag := 0

;;; vscode ;;;

;;; excel vim ;;;
global ev_mode := 0
global normal_mode := 0
global insert_mode := 1
global visual_mode := 2
global visual_line_mode := 3
global space_mode := 4
global visual_line_cp := 0
global space_cnt := 0
global Enter_cnt := 0

;;; onenote ;;;

;;; power point ;;;
global scroll_val := 0

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
;;; reset key ;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ; F13 & q::send {Blind}^q
    ; F13 & w::send {Blind}^w
    ; F13 & e::send {Blind}^e
    ; F13 & r::send {Blind}^r
    ; F13 & t::send {Blind}^t
    ; F13 & y::send {Blind}^y
    ; F13 & u::send {Blind}^u
    ; F13 & i::send {Blind}^i
    ; F13 & o::send {Blind}^o
    ; F13 & p::send {Blind}^p
    ; F13 & a::send {Blind}^a
    ; F13 & s::send {Blind}^s
    ; F13 & d::send {Blind}^d
    ; F13 & f::send {Blind}^f
    ; F13 & g::send {Blind}^g
    ; F13 & h::send {Blind}^h
    ; F13 & j::send {Blind}^j
    ; F13 & k::send {Blind}^k
    ; F13 & l::send {Blind}^l
    ; F13 & z::send {Blind}^z
    ; F13 & x::send {Blind}^x
    ; F13 & c::send {Blind}^c
    ; F13 & v::send {Blind}^v
    ; F13 & b::send {Blind}^b
    ; F13 & n::send {Blind}^n
    ; F13 & m::send {Blind}^m
    ; F13 & 1::send {Blind}^1
    ; F13 & 2::send {Blind}^2
    ; F13 & 3::send {Blind}^3
    ; F13 & 4::send {Blind}^4
    ; F13 & 5::send {Blind}^5
    ; F13 & 6::send {Blind}^6
    ; F13 & 7::send {Blind}^7
    ; F13 & 8::send {Blind}^8
    ; F13 & 9::send {Blind}^9
    ; F13 & 0::send {Blind}^0
    ; F13 & -::send {Blind}^-
    ; F13 & /::send {Blind}^/
    ; F13 & `;::send {Blind}^`;
    ; F13 & Tab::send {Blind}^{Tab}
    ; F13 & [::send {Blind}^[
    ; F13 & ]::send {Blind}^]
    ; F13 & Space::send {Blind}^{Space}
    ; F13 & Enter::send {Blind}^{Enter}


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; basic key setting ;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    ;---2 times F13---
    ; F13::
    ;     Keywait, F13, U
    ;     Keywait, F13, D T0.2
    ;     if (ErrorLevel=1){
    ;         MsgBox, a  
    ;     }else{
    ;         MsgBox, b
    ;     }
    ;     return

    ;---cursol move---
    ; F13 & K::send {Blind}{Up}
    ; F13 & J::send {Blind}{Down}
    ; F13 & H::send {Blind}{Left}
    ; F13 & L::send {Blind}{Right}
    ; F13 & I::send {Blind}{Home}
    ; F13 & O::send {Blind}{End}
    ; F13 & K::send {Blind}{Down}
    ; F13 & J::send {Blind}{Left}
    ; F13 & L::send {Blind}{Right}
    ; F13 & I::send {Blind}{Up}

    ^Up::Send    {PgUp}
    ^Down::Send  {PgDn}
    ^Left::Send  {Home}
    ^Right::Send {End}
    ^+Up::Send    +{PgUp}
    ^+Down::Send  +{PgDn}
    ^+Left::Send  +{Home}
    ^+Right::Send +{End}
    

;---num pad---
    Numpad0::send ^#t
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
    sc07B & Tab::send !^{Tab}
    
    ; sc079 : henkan
    ; sc07B : muhenkan


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; winpane ;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
        ; ^MButton::
        ;     CoordMode, Mouse, Screen ;; mouse absolute pos setting
        ;     MouseGetPos, Xmou, Ymou, winid
        ;     WinActivate, ahk_id %winid%
        ;     WinGetClass, win_class, A
        ;     get_moni2()
        ;     rate_setting()
        ;     if (moni_sel==1){
        ;         m1_middle_rate := Xmou_rate
        ;     }else if (moni_sel==2){
        ;         m2_middle_rate := Xmou_rate
        ;     }else if (moni_sel==3){
        ;         m3_middle_rate := Xmou_rate
        ;     }
        ;     return

        ; ^RButton::
        ;     CoordMode, Mouse, Screen ;; mouse absolute pos setting
        ;     MouseGetPos, Xmou, Ymou, winid
        ;     WinActivate, ahk_id %winid%
        ;     WinGetClass, win_class, A
        ;     if (win_class=="WorkerW" || win_class=="Shell_TrayWnd" || win_class=="Shell_SecondaryTrayWnd"){
        ;     }else {
        ;         Keywait, RButton, U T0.05
        ;         if (ErrorLevel=1){
        ;             ;; hold
        ;             get_moni2()
        ;             rate_setting()
        ;             resizewin2(moni_sel, Xmou_rate-0.1, Ymou_rate-0.1, 0.15, 0.25)
        ;         }else{
        ;             ;; click
        ;             if (rclick_status==1){
        ;                 rclick_status := 0
        ;             }else {
        ;                 rclick_status := rclick_status + 1
        ;             }
        ;             get_moni2()
        ;             rate_setting()
        ;             if (rclick_status==0){
        ;                 resizewin2(moni_sel, Xmou_rate-0.01, Ymou_rate-0.01, rclick_resize1_w, rclick_resize1_h)
        ;             }else if (rclick_status==1){
        ;                 resizewin2(moni_sel, Xmou_rate-0.01, Ymou_rate-0.01, rclick_resize2_w, rclick_resize2_h)
        ;             }
        ;         }
        ;     }
        ;     return

        ; ^LButton::
        ;     LButton_count := 0
        ;     Xmou_hold := 0
        ;     Ymou_hold := 0
        ;     CoordMode, Mouse, Screen ;; mouse absolute pos setting
        ;     MouseGetPos, Xmou, Ymou, winid
        ;     WinActivate, ahk_id %winid%
        ;     WinGetClass, win_class, A
        ;     WinGetPos,X,Y,W,H,A
        ;     if (win_class=="WorkerW" || win_class=="Shell_TrayWnd" || win_class=="Shell_SecondaryTrayWnd"){
        ;             keywait, LButton, D T0.1
        ;             if (ErrorLevel==1){
        ;                 ;; single click
        ;             }else{
        ;                 ;; double click
        ;                 get_moni2()
        ;                 rate_setting()
        ;                 if (moni_sel==1){
        ;                     m1_middle_rate := Xmou_rate
        ;                 }else if (moni_sel==2){
        ;                     m2_middle_rate := Xmou_rate
        ;                 }else if (moni_sel==3){
        ;                     m3_middle_rate := Xmou_rate
        ;                 }
        ;             }
        ;     }else {
        ;         Xmou_pre := Xmou
        ;         Ymou_pre := Ymou
        ;         xt := X+0*(W/3)
        ;         x1 := X+1*(W/3)
        ;         x2 := X+2*(W/3)
        ;         xb := X+3*(W/3)
        ;         yt := Y+0*(H/3)
        ;         y1 := Y+1*(H/3)
        ;         y2 := Y+2*(H/3)        
        ;         yb := Y+3*(H/3)
        ;         rsz_lt := 0
        ;         rsz_lm := 1
        ;         rsz_lb := 2
        ;         rsz_mt := 3
        ;         rsz_mm := 4
        ;         rsz_mb := 5
        ;         rsz_rt := 6
        ;         rsz_rm := 7
        ;         rsz_rb := 8
                
        ;         if ((Xmou>xt && Xmou<x1)&&(Ymou>yt && Ymou<y1 )){
        ;             resize_type := rsz_lt
        ;             X_fix := X+W
        ;             Y_fix := Y+H
        ;             Xmou_edge := X - Xmou
        ;             Ymou_edge := Y - Ymou
        ;         }else if ((Xmou>xt && Xmou<x1)&&(Ymou>y1 && Ymou<y2 )){
        ;             resize_type := rsz_lm
        ;             X_fix := X+W
        ;             ; Y_fix := Y+H
        ;             Xmou_edge := X - Xmou
        ;             ; Ymou_edge := Y - Ymou
        ;         }else if ((Xmou>xt && Xmou<x1)&&(Ymou>y2 && Ymou<yb )){
        ;             resize_type := rsz_lb
        ;             X_fix := X+W
        ;             Y_fix := Y
        ;             Xmou_edge := X - Xmou
        ;             Ymou_edge := Y+H - Ymou
        ;         }else if ((Xmou>x1 && Xmou<x2)&&(Ymou>yt && Ymou<y1 )){
        ;             resize_type := rsz_mt
        ;             ; X_fix := X+W
        ;             Y_fix := Y+H
        ;             ; Xmou_edge := X - Xmou
        ;             Ymou_edge := Y - Ymou
        ;         }else if ((Xmou>x1 && Xmou<x2)&&(Ymou>y1 && Ymou<y2 )){
        ;             resize_type := rsz_mm
        ;             Xmou_edge := X - Xmou
        ;             Ymou_edge := Y - Ymou
        ;         }else if ((Xmou>x1 && Xmou<x2)&&(Ymou>y2 && Ymou<yb )){
        ;             resize_type := rsz_mb
        ;             ; X_fix := X
        ;             Y_fix := Y
        ;             ; Xmou_edge := X+W-Xmou
        ;             Ymou_edge := Y+H-Ymou
        ;         }else if ((Xmou>x2 && Xmou<xb)&&(Ymou>yt && Ymou<y1 )){
        ;             resize_type := rsz_rt
        ;             X_fix := X
        ;             Y_fix := Y+H            
        ;             Xmou_edge := X+W - Xmou
        ;             Ymou_edge := Y - Ymou            
        ;         }else if ((Xmou>x2 && Xmou<xb)&&(Ymou>y1 && Ymou<y2 )){
        ;             resize_type := rsz_rm
        ;             ; X_fix := X
        ;             ; Y_fix := Y+H            
        ;             Xmou_edge := X+W - Xmou
        ;             ; Ymou_edge := Y - Ymou            
        ;         }else if ((Xmou>x2 && Xmou<xb)&&(Ymou>y2 && Ymou<yb )){
        ;             resize_type := rsz_rb
        ;             X_fix := X
        ;             Y_fix := Y
        ;             Xmou_edge := X+W-Xmou
        ;             Ymou_edge := Y+H-Ymou
        ;         }
                
        ;         Keywait, LButton, U T0.1
        ;         if (ErrorLevel=1){
        ;             CoordMode, Mouse, Screen ;; mouse absolute pos setting
        ;             MouseGetPos, Xmou, Ymou, winid
        ;             if (abs(Xmou_pre-Xmou)<1 || abs(Ymou_pre-Ymou)<1){
        ;                 get_moni2()
        ;                 rate_setting()
        ;                 resize_long_click()
        ;             }else{
        ;                 Loop{
        ;                     ; Sleep, 1
        ;                     LButton_count := LButton_count + 1
        ;                     CoordMode, Mouse, Screen ;; mouse absolute pos setting
        ;                     MouseGetPos, Xmou, Ymou, winid
        ;                     Xmou_diff := Xmou - Xmou_pre
        ;                     Ymou_diff := Ymou - Ymou_pre
        ;                     Xmou_hold := Xmou_hold + Xmou_diff
        ;                     Ymou_hold := Ymou_hold + Ymou_diff
        ;                     Xmou_pre := Xmou
        ;                     Ymou_pre := Ymou
        ;                     WinGetPos,X,Y,W,H,A
        ;                     if (resize_type==rsz_lt){
        ;                         resizewin(Xmou+Xmou_edge, Ymou+Ymou_edge, X_fix-(Xmou+Xmou_edge), Y_fix-(Ymou+Ymou_edge))
        ;                     }else if (resize_type==rsz_lm){
        ;                         resizewin(Xmou+Xmou_edge, Y, X_fix-(Xmou+Xmou_edge), H)
        ;                     }else if (resize_type==rsz_lb){
        ;                         resizewin(Xmou+Xmou_edge, Y_fix, X_fix-(Xmou+Xmou_edge), Ymou+Ymou_edge-Y_fix)    
        ;                     }else if (resize_type==rsz_rt){
        ;                         resizewin(X_fix, Ymou+Ymou_edge, Xmou+Xmou_edge-X_fix, Y_fix-(Ymou+Ymou_edge))
        ;                     }else if (resize_type==rsz_rm){
        ;                         resizewin(X, Y, Xmou+Xmou_edge-X, H)
        ;                     }else if (resize_type==rsz_rb){
        ;                         resizewin(X, Y, Xmou-X+Xmou_edge, Ymou-Y+Ymou_edge)
        ;                     }else if (resize_type==rsz_mt){
        ;                         resizewin(X, Ymou+Ymou_edge, W, Y_fix-(Ymou+Ymou_edge))
        ;                     }else if (resize_type==rsz_mb){
        ;                         resizewin(X, Y, W, Ymou-Y+Ymou_edge)
        ;                     }else if (resize_type==rsz_mm){
        ;                         resizewin(Xmou + Xmou_edge, Ymou + Ymou_edge, W, H)
        ;                     }
        ;                     GetKeyState, state, LButton, P
        ;                     If state = U			; The key has been released, so break out of the loop.
        ;                         Break
        ;                 }
        ;             }         
        ;         }else{
        ;             keywait, LButton, D T0.15
        ;             if (ErrorLevel==1){
        ;                 ;; single click
        ;                 get_moni2()
        ;                 rate_setting()
        ;                 resize_short_click()
        ;             }else{
        ;                 ;; double click
        ;                 get_moni2()
        ;                 rate_setting()
        ;                 resize_double_click()
        ;             }
        ;         }
        ;     }
        ;     return

        ; ;; alternative ctrl click
        F13 & LButton::
            send {CtrlDown}
            MouseClick LEFT , , , , , D,
            Keywait LButton, 
            MouseClick LEFT , , , , , U,
            send {CtrlUp}
            return

        ; ;; resize func
        ; resizewin(Xpos, Ypos, Width, Height){
        ;     WinMove, A, , Xpos, Ypos, Width, Height
        ; }

        ; resizewin2(moni, Xrate, Yrate, Wrate, Hrate){
        ;     if (moni==1){
        ;         rszX := m1_moni_left + m1_moni_width * Xrate
        ;         rszY := m1_moni_top + m1_moni_height * Yrate
        ;         rszW := m1_moni_width * Wrate
        ;         rszH := m1_moni_height * Hrate
        ;     }else if (moni==2){
        ;         rszX := m2_moni_left + m2_moni_width * Xrate
        ;         rszY := m2_moni_top + m2_moni_height * Yrate
        ;         rszW := m2_moni_width * Wrate
        ;         rszH := m2_moni_height * Hrate
        ;     }else if(moni==3){
        ;         rszX := m3_moni_left + m3_moni_width * Xrate
        ;         rszY := m3_moni_top + m3_moni_height * Yrate
        ;         rszW := m3_moni_width * Wrate
        ;         rszH := m3_moni_height * Hrate
        ;     }
        ;     WinMove, A, , rszX, rszY, rszW, rszH
        ; }

        ; rate_setting(){
        ;     CoordMode, Mouse, Screen ;; mouse absolute pos setting
        ;     MouseGetPos, Xmou, Ymou, winid
        ;     WinActivate, ahk_id %winid%
        ;     WinGetPos,X,Y,W,H,A
            
        ;     if (Xmou>m1_moni_left && Xmou<m1_moni_left+m1_moni_width){
        ;         moni_sel := 1
        ;     }else if (Xmou>m2_moni_left && Xmou<m2_moni_left+m2_moni_width){
        ;         moni_sel := 2
        ;     }else if (Xmou>m3_moni_left && Xmou<m3_moni_left+m3_moni_width){
        ;         moni_sel := 3 
        ;     }

        ;     if (moni_sel==1){
        ;         Xrate := (X - m1_moni_left) / m1_moni_width
        ;         Yrate := (Y - m1_moni_top) / m1_moni_height
        ;         Wrate := W / m1_moni_width
        ;         Hrate := H / m1_moni_height
        ;         Xmou_rate := (Xmou - m1_moni_left)/m1_moni_width
        ;         Ymou_rate := (Ymou - m1_moni_top)/m1_moni_height
        ;         middle_rate := m1_middle_rate
        ;     }else if (moni_sel==2){
        ;         Xrate := (X - m2_moni_left) / m2_moni_width
        ;         Yrate := (Y - m2_moni_top) / m2_moni_height
        ;         Wrate := W / m2_moni_width
        ;         Hrate := H / m2_moni_height
        ;         Xmou_rate := (Xmou - m2_moni_left)/m2_moni_width
        ;         Ymou_rate := (Ymou - m2_moni_top)/m2_moni_height
        ;         middle_rate := m2_middle_rate
        ;     }else if(moni_sel==3){
        ;         Xrate := (X - m3_moni_left) / m3_moni_width
        ;         Yrate := (Y - m3_moni_top) / m3_moni_height
        ;         Wrate := W / m3_moni_width
        ;         Hrate := H / m3_moni_height
        ;         Xmou_rate := (Xmou - m3_moni_left)/m3_moni_width
        ;         Ymou_rate := (Ymou - m3_moni_top)/m3_moni_height
        ;         middle_rate := m3_middle_rate
        ;     }

        ; }

        ; resize_short_click(){
        ;     if (Xrate==midcenter_open_xrate && Wrate==midcenter_open_wrate){
        ;             ;; close middle center
        ;             resize_midcenter_buf(moni_sel)
        ;     }else if (Wrate>=0.9){
        ;         ;; close center 
        ;         resize_center_buf(moni_sel)
        ;     }else if (Hrate>0.8){
        ;         ;; close side
        ;         if (Xmou_rate<middle_rate){
        ;             ;; close left
        ;             ; reset_middle(moni_sel, 0)
        ;             resize_left_buf(moni_sel)
        ;         }else {
        ;             ;; close right
        ;             ; reset_middle(moni_sel, 1)
        ;             resize_right_buf(moni_sel)
        ;         }   
        ;     }else {
        ;         ;; open side
        ;         if (Xmou_rate<middle_rate){
        ;             ;; open left side
        ;             resizewin2(moni_sel, midcenter_open_xrate, midcenter_open_yrate, midcenter_open_wrate, midcenter_open_hrate)
        ;             get_midcenter_buf(moni_sel, Xrate, Yrate, Wrate, Hrate)
        ;         }else{
        ;             ;; opne right side
        ;             resizewin2(moni_sel, midcenter_open_xrate, midcenter_open_yrate, midcenter_open_wrate, midcenter_open_hrate)
        ;             get_midcenter_buf(moni_sel, Xrate, Yrate, Wrate, Hrate)
        ;         }
        ;     }
        ; }

        ; resize_long_click(){
        ;     if (Xrate==midcenter_open_xrate && Wrate==midcenter_open_wrate){
        ;             ;; close middle center
        ;             resize_midcenter_buf(moni_sel)
        ;     }else if (Wrate>=0.9){
        ;         ;; close center 
        ;         resize_center_buf(moni_sel)
        ;     }else if (Hrate>0.8){
        ;         ;; close side
        ;         if (Xmou_rate<middle_rate){
        ;             ;; close left
        ;             ; reset_middle(moni_sel, 0)
        ;             resize_left_buf(moni_sel)
        ;         }else {
        ;             ;; close right
        ;             ; reset_middle(moni_sel, 1)
        ;             resize_right_buf(moni_sel)
        ;         }   
        ;     }else {
        ;         ;; open side
        ;         if (Xmou_rate<middle_rate){
        ;             ;; open left side
        ;             resizewin2(moni_sel, side_open_xrate, side_open_yrate, middle_rate-side_open_xrate, side_open_hrate)
        ;             get_left_buf(moni_sel, Xrate, Yrate, Wrate, Hrate)
        ;         }else{
        ;             ;; opne right side
        ;             resizewin2(moni_sel, middle_rate, side_open_yrate, 1-middle_rate-side_open_xrate, side_open_hrate)
        ;             get_right_buf(moni_sel, Xrate, Yrate, Wrate, Hrate)
        ;         }
        ;     }
        ; }

        ; resize_double_click(){
        ;     get_center_buf(moni_sel, Xrate, Yrate, Wrate, Hrate)
        ;     resizewin2(moni_sel, xedge, center_open_yrate, 1-xedge*2, center_open_hrate)
        ; }

        ; reset_middle(moni, side){
        ;     if (moni==1){
        ;         if (side==0){
        ;             m1_middle_rate := Wrate
        ;         }else {
        ;             m1_middle_rate := Xrate
        ;         }
        ;     }else if (moni==2){
        ;         if (side==0){
        ;             m2_middle_rate := Wrate
        ;         }else {
        ;             m2_middle_rate := Xrate
        ;         }
        ;     }else if (moni==3)
        ;         if (side==0){
        ;             m3_middle_rate := Wrate
        ;         }else {
        ;             m3_middle_rate := Xrate
        ;         }
        ; }

        ; resize_center_buf(moni){
        ;     if (moni==1) {
        ;         resizewin2(moni, m1_c_buf_x, m1_c_buf_y, m1_c_buf_w, m1_c_buf_h)
        ;     }else if (moni==2){
        ;         resizewin2(moni, m2_c_buf_x, m2_c_buf_y, m2_c_buf_w, m2_c_buf_h)
        ;     }else if (moni==3){
        ;         resizewin2(moni, m3_c_buf_x, m3_c_buf_y, m3_c_buf_w, m3_c_buf_h)
        ;     }
        ; }

        ; resize_midcenter_buf(moni){
        ;     if (moni==1) {
        ;         resizewin2(moni, m1_mc_buf_x-0.05, m1_mc_buf_y-0.05, m1_mc_buf_w+0.1, m1_mc_buf_h+0.1)
        ;         resizewin2(moni, m1_mc_buf_x, m1_mc_buf_y, m1_mc_buf_w, m1_mc_buf_h)
        ;     }else if (moni==2){
        ;         resizewin2(moni, m2_mc_buf_x-0.05, m2_mc_buf_y-0.05, m2_mc_buf_w+0.1, m2_mc_buf_h+0.1)
        ;         resizewin2(moni, m2_mc_buf_x, m2_mc_buf_y, m2_mc_buf_w, m2_mc_buf_h)
        ;     }else if (moni==3){
        ;         resizewin2(moni, m3_mc_buf_x-0.05, m3_mc_buf_y-0.05, m3_mc_buf_w+0.1, m3_mc_buf_h+0.1)
        ;         resizewin2(moni, m3_mc_buf_x, m3_mc_buf_y, m3_mc_buf_w, m3_mc_buf_h)
        ;     }
        ; }


        ; resize_left_buf(moni){
        ;     if (moni==1) {
        ;         resizewin2(moni, m1_l_buf_x-0.05, m1_l_buf_y-0.05, m1_l_buf_w+0.1, m1_l_buf_h+0.1)
        ;         resizewin2(moni, m1_l_buf_x, m1_l_buf_y, m1_l_buf_w, m1_l_buf_h)
        ;     }else if (moni==2){
        ;         resizewin2(moni, m2_l_buf_x-0.05, m2_l_buf_y-0.05, m2_l_buf_w+0.1, m2_l_buf_h+0.1)
        ;         resizewin2(moni, m2_l_buf_x, m2_l_buf_y, m2_l_buf_w, m2_l_buf_h)
        ;     }else if (moni==3){
        ;         resizewin2(moni, m3_l_buf_x-0.05, m3_l_buf_y-0.05, m3_l_buf_w+0.1, m3_l_buf_h+0.1)
        ;         resizewin2(moni, m3_l_buf_x, m3_l_buf_y, m3_l_buf_w, m3_l_buf_h)
        ;     }
        ; }

        ; resize_right_buf(moni){
        ;     if (moni==1) {
        ;         resizewin2(moni, m1_r_buf_x-0.05, m1_r_buf_y-0.05, m1_r_buf_w+0.1, m1_r_buf_h+0.1)
        ;         resizewin2(moni, m1_r_buf_x, m1_r_buf_y, m1_r_buf_w, m1_r_buf_h)
        ;     }else if (moni==2){
        ;         resizewin2(moni, m2_r_buf_x-0.05, m2_r_buf_y-0.05, m2_r_buf_w+0.1, m2_r_buf_h+0.1)
        ;         resizewin2(moni, m2_r_buf_x, m2_r_buf_y, m2_r_buf_w, m2_r_buf_h)
        ;     }else if (moni==3){
        ;         resizewin2(moni, m3_r_buf_x-0.05, m3_r_buf_y-0.05, m3_r_buf_w+0.1, m3_r_buf_h+0.1)
        ;         resizewin2(moni, m3_r_buf_x, m3_r_buf_y, m3_r_buf_w, m3_r_buf_h)
        ;     }
        ; }

        ; get_center_buf(moni, xbuf, ybuf, wbuf, hbuf) {
        ;     if (moni==1) {
        ;         m1_c_buf_x := xbuf
        ;         m1_c_buf_y := ybuf
        ;         m1_c_buf_w := wbuf
        ;         m1_c_buf_h := hbuf
        ;     }else if (moni==2){
        ;         m2_c_buf_x := xbuf
        ;         m2_c_buf_y := ybuf
        ;         m2_c_buf_w := wbuf
        ;         m2_c_buf_h := hbuf
        ;     }else if (moni==3){
        ;         m3_c_buf_x := xbuf
        ;         m3_c_buf_y := ybuf
        ;         m3_c_buf_w := wbuf
        ;         m3_c_buf_h := hbuf
        ;     }
        ; }

        ; get_midcenter_buf(moni, xbuf, ybuf, wbuf, hbuf) {
        ;     if (moni==1) {
        ;         m1_mc_buf_x := xbuf
        ;         m1_mc_buf_y := ybuf
        ;         m1_mc_buf_w := wbuf
        ;         m1_mc_buf_h := hbuf
        ;     }else if (moni==2){
        ;         m2_mc_buf_x := xbuf
        ;         m2_mc_buf_y := ybuf
        ;         m2_mc_buf_w := wbuf
        ;         m2_mc_buf_h := hbuf
        ;     }else if (moni==3){
        ;         m3_mc_buf_x := xbuf
        ;         m3_mc_buf_y := ybuf
        ;         m3_mc_buf_w := wbuf
        ;         m3_mc_buf_h := hbuf
        ;     }
        ; }

        ; get_left_buf(moni, xbuf, ybuf, wbuf, hbuf) {
        ;     if (moni==1) {
        ;         m1_l_buf_x := xbuf
        ;         m1_l_buf_y := ybuf
        ;         m1_l_buf_w := wbuf
        ;         m1_l_buf_h := hbuf
        ;     }else if (moni==2){
        ;         m2_l_buf_x := xbuf
        ;         m2_l_buf_y := ybuf
        ;         m2_l_buf_w := wbuf
        ;         m2_l_buf_h := hbuf
        ;     }else if (moni==3){
        ;         m3_l_buf_x := xbuf
        ;         m3_l_buf_y := ybuf
        ;         m3_l_buf_w := wbuf
        ;         m3_l_buf_h := hbuf
        ;     }
        ; }

        ; get_right_buf(moni, xbuf, ybuf, wbuf, hbuf) {
        ;     if (moni==1) {
        ;         m1_r_buf_x := xbuf
        ;         m1_r_buf_y := ybuf
        ;         m1_r_buf_w := wbuf
        ;         m1_r_buf_h := hbuf
        ;     }else if (moni==2){
        ;         m2_r_buf_x := xbuf
        ;         m2_r_buf_y := ybuf
        ;         m2_r_buf_w := wbuf
        ;         m2_r_buf_h := hbuf
        ;     }else if (moni==3){
        ;         m3_r_buf_x := xbuf
        ;         m3_r_buf_y := ybuf
        ;         m3_r_buf_w := wbuf
        ;         m3_r_buf_h := hbuf
        ;     }
        ; }

        ; get_moni(){
        ;     if (wp_init_flag==0){
        ;         SysGet, moni, Monitor, 1
        ;         if (moniRight>10){
        ;             m1_moni_width := moniRight
        ;             m1_moni_height := moniBottom
        ;         }
        ;         if (moniLeft<-10){
        ;             m2_moni_width := moniLeft
        ;             m2_moni_height := moniBottom
        ;         }
        ;         SysGet, moni, Monitor, 2
        ;         if (moniRight>10){
        ;             m1_moni_width := moniRight
        ;             m1_moni_height := moniBottom
        ;         }
        ;         if (moniLeft<-10){
        ;             m2_moni_width := -1*moniLeft
        ;             m2_moni_height := moniBottom
        ;         }      
        ;         wp_init_flag := 1
        ;     }        
        ; }

        ; get_moni2(){
        ;     if (wp_init_flag==0){
        ;         SysGet, moni, Monitor, 1
        ;         m1_moni_width := moniRight - moniLeft
        ;         m1_moni_height := moniBottom - moniTop
        ;         m1_moni_left := moniLeft
        ;         m1_moni_top := moniTop
        ;         SysGet, moni, Monitor, 2
        ;         m2_moni_width := moniRight - moniLeft
        ;         m2_moni_height := moniBottom - moniTop
        ;         m2_moni_left := moniLeft
        ;         m2_moni_top := moniTop
        ;         SysGet, moni, Monitor, 3
        ;         m3_moni_width := moniRight - moniLeft
        ;         m3_moni_height := moniBottom - moniTop
        ;         m3_moni_left := moniLeft
        ;         m3_moni_top := moniTop
        ;         wp_init_flag := 1
        ;     } 
        ; }

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; keep memo ;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

    ; RAlt::
    ;     if (memo_tgl==0){
    ;         WinActivate, MEMO01
    ;         memo_tgl := 1
    ;     }else {
    ;         WinMinimize, MEMO01
    ;         memo_tgl := 0
    ;     }
    ;     return
    ; RShift::
    ;     if (memo_tgl2==0){
    ;         WinActivate, ppt_memo2
    ;         memo_tgl2 := 1
    ;     }else {
    ;         WinMinimize, ppt_memo2
    ;         memo_tgl2 := 0
    ;     }
    ;     return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Miro ;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe Miro.exe
        ;; Double Click as create new TEXT
        ; LButton::
        ;     If (A_PriorHotKey == A_ThisHotKey and A_TimeSincePriorHotkey < 300){
        ;         send t
        ;         MouseClick LEFT , , , , , D,
        ;         MouseClick LEFT , , , , , U,
        ;     }else{
        ;         MouseClick LEFT , , , , , D,
        ;         Keywait LButton, 
        ;         MouseClick LEFT , , , , , U,
        ;     }
        ;     Return
        LButton::
            keywait, LButton, U T0.1
            if (ErrorLevel==1){
                ;; single click hold
                MouseClick LEFT , , , , , D,
                Keywait LButton, 
                MouseClick LEFT , , , , , U,
            }else{
                keywait, LButton, D T0.1
                if (ErrorLevel==1){
                    ;; single click
                    MouseClick LEFT , , , , , D,
                    MouseClick LEFT , , , , , U,
                }else{
                    keywait, LButton, U
                    keywait, LButton, D T0.1
                    if (ErrorLevel==1){
                        ;; double click
                        click 2
                    }else {
                        ;; triple click
                        send t
                        MouseClick LEFT , , , , , D,
                        MouseClick LEFT , , , , , U,                        
                    }
                }
            }
            return
    #IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; VScode ;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe Code.exe
    #IfWinActive

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; Onenote ;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    #IfWinActive,ahk_exe ONENOTE.EXE
        ;; Double Click as create new TEXT
        ; LButton::
        ;     If (A_PriorHotKey == A_ThisHotKey and A_TimeSincePriorHotkey < 300){
        ;         send {CtrlDown}
        ;         MouseClick LEFT , , , , , D,
        ;         MouseClick LEFT , , , , , U,
        ;         send {CtrlUp}            
        ;     }else{
        ;         MouseClick LEFT , , , , , D,
        ;         Keywait LButton, 
        ;         MouseClick LEFT , , , , , U,
        ;     }
        ;     Return
        ^1::
            send, ^!1
            return
        ^2::
            send, ^!2
            return
        ^3::
            send, ^!3
            return
        Numpad1::
            send, ^m
            Sleep, 500
            send, ^n
            Sleep, 500
            send, {Alt}wt
            Sleep, 500
            send, ^+.
            send, ^+.
            send, ^+.
            send, ^+.
            send, ^b
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
        Numpad2::
            send, ^c
            Sleep, 500
            send, #e
            Sleep, 500
            send, !d
            Sleep, 500
            send, ^v
            Sleep, 500
            send, {Enter}
            return
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
            send +{Space}
            send ^+;
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
        sc07B & right:: send {Tab}{End}
        sc07B & left:: send +{Tab}{End}
        sc07B & down:: send {Tab}+{Tab}{End}{Down}{Tab}+{Tab}{End}
        sc07B & up:: send {Tab}+{Tab}{Home}{Up}{Tab}+{Tab}{End}
        ; ~Lshift & WheelUp::ComObjActive("PowerPoint.Application").ActiveWindow.SmallScroll(0,0,0,1)
        ; ~Lshift & WheelDown::ComObjActive("PowerPoint.Application").ActiveWindow.SmallScroll(0,0,1,0)
        ; ^WheelUp::
        ;     if (scroll_val>-10 && scroll_val<15) {
        ;         if (scroll_val == 14){
        ;         }else{
        ;             scroll_val := scroll_val + 1
        ;         }
        ;     }
        ;     if (scroll_val==0){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 100 
        ;     }else if (scroll_val==1){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 110 
        ;     }else if (scroll_val==2){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 120 
        ;     }else if (scroll_val==3){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 130 
        ;     }else if (scroll_val==4){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 140 
        ;     }else if (scroll_val==5){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 150 
        ;     }else if (scroll_val==6){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 160 
        ;     }else if (scroll_val==7){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 170 
        ;     }else if (scroll_val==8){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 180 
        ;     }else if (scroll_val==9){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 190 
        ;     }else if (scroll_val==10){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 200 
        ;     }else if (scroll_val==11){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 210 
        ;     }else if (scroll_val==12){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 220 
        ;     }else if (scroll_val==13){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 230 
        ;     }else if (scroll_val==14){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 240 
        ;     }else if (scroll_val==15){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 250 
        ;     }else if (scroll_val==-1){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 90 
        ;     }else if (scroll_val==-2){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 80 
        ;     }else if (scroll_val==-3){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 70 
        ;     }else if (scroll_val==-4){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 60 
        ;     }else if (scroll_val==-5){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 50
        ;     }else if (scroll_val==-6){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 40 
        ;     }else if (scroll_val==-7){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 30 
        ;     }else if (scroll_val==-8){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 20 
        ;     }else if (scroll_val==-9){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 10 
        ;     }
        ;     return
        ; ^WheelDown::
        ;     if (scroll_val>-10 && scroll_val<15) {
        ;         if (scroll_val == -9){
        ;         }else{
        ;             scroll_val := scroll_val - 1
        ;         }
        ;     }
        ;     if (scroll_val==0){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 100 
        ;     }else if (scroll_val==1){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 110 
        ;     }else if (scroll_val==2){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 120 
        ;     }else if (scroll_val==3){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 130 
        ;     }else if (scroll_val==4){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 140 
        ;     }else if (scroll_val==5){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 150 
        ;     }else if (scroll_val==6){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 160 
        ;     }else if (scroll_val==7){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 170 
        ;     }else if (scroll_val==8){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 180 
        ;     }else if (scroll_val==9){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 190 
        ;     }else if (scroll_val==10){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 200 
        ;     }else if (scroll_val==11){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 210 
        ;     }else if (scroll_val==12){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 220 
        ;     }else if (scroll_val==13){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 230 
        ;     }else if (scroll_val==14){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 240 
        ;     }else if (scroll_val==15){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 250 
        ;     }else if (scroll_val==-1){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 90 
        ;     }else if (scroll_val==-2){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 80 
        ;     }else if (scroll_val==-3){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 70 
        ;     }else if (scroll_val==-4){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 60 
        ;     }else if (scroll_val==-5){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 50
        ;     }else if (scroll_val==-6){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 40 
        ;     }else if (scroll_val==-7){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 30 
        ;     }else if (scroll_val==-8){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 20 
        ;     }else if (scroll_val==-9){
        ;         ComObjActive("PowerPoint.Application").ActiveWindow.View.Zoom := 10 
        ;     }
        ;     return
    #IfWinActive


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

