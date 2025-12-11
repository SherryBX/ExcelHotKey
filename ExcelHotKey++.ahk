; 仅在 Excel 中生效
#IfWinActive ahk_exe EXCEL.EXE
^]::Send ^{PgDn}  ; Ctrl+] 映射为 Ctrl+PageDown
^[::Send ^{PgUp}  ; Ctrl+[ 映射为 Ctrl+PageUp

; 选中当前行
^+h::  ; Ctrl+Shift+H
    CoordMode, Caret, Screen
    CaretY := A_CaretY
    
    if (CaretY > 0)
    {
        clickY := CaretY + 18
        Click 20, %clickY%
    }
    else
    {
        MouseGetPos, , ypos
        clickY := ypos + 18
        Click 20, %clickY%
    }
return

; 合并单元格并居中
!c::  ; Alt+C
    Send !h  ; 打开开始菜单
    Sleep 50
    Send m    ; 合并菜单
    Sleep 50
    Send c    ; 合并并居中
return

#IfWinActive