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
        clickY := CaretY + 20
        Click 20, %clickY%
    }
    else
    {
        MouseGetPos, , ypos
        clickY := ypos + 15
        Click 20, %clickY%
    }
return

#IfWinActive