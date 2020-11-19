; This script was created using Pulover's Macro Creator
; www.macrocreator.com

#NoEnv
SetWorkingDir %A_ScriptDir%
CoordMode, Mouse, Window
SendMode Input
#SingleInstance Force
SetTitleMatchMode 2
#WinActivateForce
SetControlDelay 1
SetWinDelay 0
SetKeyDelay -1
SetMouseDelay -1
SetBatchLines -1


WIPINVOICES:
SetNumLockState, On
SetScrollLockState, Off
MsgBox, 4, , 
(LTrim
Welcome to the Sands invoicing macro. This program has been written to quickly and accurately populate a Sands invoice in Adagio.

Before proceeding please ensure that Adagio is opened with an invoice created and that Google chrome is opened and logged into the work order system. 

For best results have as few programs open as possible and do not touch the keyboard or mouse until prompted. 

Would you like to include the excel project line in the ship to?
)
BlockInput, ON
IfMsgBox, Yes
{
    AddressFormat := 1
}
Else
{
    AddressFormat := 0
}
Sleep, 600
Loop
{
    WinActivate, ahk_exe EXCEL.EXE
    Sleep, 333
    IfWinActive, ahk_exe EXCEL.EXE
    {
        Break
    }
}
Send, {Enter}
Sleep, 400
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, C8
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 30
Send, {c}
Send, {LControl Up}
Sleep, 30
Description := Clipboard
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Alt Down}
Sleep, 50
Send, {h}
Sleep, 60
Send, {v}
Sleep, 50
Send, {Alt Up}
Sleep, 50
Send, {LControl Down}
Sleep, 25
Send, {v}
Send, {LControl Up}
Sleep, 25
Sleep, 100
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, L6
Sleep, 150
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
SalesConv(Clipboard)
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab 2}
Sleep, 100
Send, %Reference%
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, E18
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, L2
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
WOValue := Clipboard
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab}
Sleep, 50
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Send, {Tab 2}
Sleep, 100
Send, %Header%
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, A18
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab 6}
Sleep, 50
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
BlockInput, Off
MsgBox, 0, , Please review the tax information and return the cursor to the "ship via" input box after making any necessary changes.
BlockInput, ON
Sleep, 600
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, L2
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab 6}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Send, {LShift Down}
Sleep, 50
Send, {Tab 5}
Sleep, 50
Send, {LShift Up}
Sleep, 30
IfWinActive, Warning
{
    Sleep, 100
    Send, {Escape}
}
Send, {LAlt Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {b}
Sleep, 50
Send, {LAlt Up}
Sleep, 60
BlockInput, Off
MsgBox, 0, , Please review the bill to information and email recipients before hitting OK
BlockInput, ON
Sleep, 600
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab}
Sleep, 50
Send, {LAlt Down}
Sleep, 100
Send, {h}
Sleep, 50
Send, {v}
Sleep, 50
Send, {s}
Sleep, 50
Send, {LAlt Up}
Sleep, 100
If AddressFormat = 1
{
    Loop
    {
        WinActivate, Excel
        Sleep, 333
        IfWinActive, Excel
        {
            Break
        }
    }
    Send, {Enter}
    Sleep, 50
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, C8
    Sleep, 100
    Send, {Enter}
    Sleep, 100
    Send, {LControl Down}
    Sleep, 50
    Send, {c}
    Sleep, 50
    Send, {LControl Up}
    Sleep, 50
    Loop
    {
        WinActivate, Invoices
        Sleep, 333
        IfWinActive, Invoices
        {
            Break
        }
    }
    Send, {Tab}
    Sleep, 100
    Send, {LControl Down}
    Sleep, 50
    Send, {v}
    Sleep, 50
    Send, {LControl Up}
    Sleep, 50
}
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Send, {Enter}
Sleep, 50
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I12
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I13
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
If AddressFormat = 1
{
    Send, {Tab 2}
    Sleep, 50
}
Else
{
    Send, {Tab}
    Sleep, 50
}
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I14
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
If AddressFormat = 0
{
    Send, {Tab 2}
    Sleep, 50
}
Else
{
    Send, {Tab}
    Sleep, 50
}
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I15
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Loop
{
    WinActivate, Excel
    Sleep, 333
    IfWinActive, Excel
    {
        Break
    }
}
Loop
{
    Send, {F5}
    IfWinExist, Go To
    {
        Break
    }
}
SendRaw, I16
Sleep, 100
Send, {Enter}
Sleep, 100
Send, {LControl Down}
Sleep, 30
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 30
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {Tab}
Sleep, 100
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 30
Send, {Tab}
Sleep, 100
Send, {Delete}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Delete}
Sleep, 100
Send, {Tab}
Sleep, 100
Send, {Delete}
Sleep, 100
BlockInput, OFF
MsgBox, 0, , Please review the ship to information before hitting OK
BlockInput, ON
Sleep, 600
Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Invoices?woNo=%WOValue%
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {LAlt Down}
Sleep, 50
Send, {h}
Sleep, 50
Send, {v}
Sleep, 50
Send, {t}
Sleep, 50
Send, {LAlt Up}
Sleep, 60
Send, {Tab}
Sleep, 50
Send, {LControl Down}
Sleep, 30
Send, {c}
Send, {LControl Up}
Sleep, 30
InvoiceNum := Clipboard
Clipboard := ""
Loop
{
    WinActivate, Chrome
    Sleep, 333
    IfWinExist, Chrome
    {
        Break
    }
}
Loop
{
    Send, {LControl Down}
    Sleep, 30
    Send, {a}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 180
    IfInString, Clipboard, Invoice #
    {
        Break
    }
}
Send, {LControl Down}
Sleep, 50
Send, {LShift Down}
Sleep, 50
Send, {Home}
Sleep, 50
Send, {LShift Up}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Send, {Tab 18}
Sleep, 100
Send, {Enter}
Sleep, 1000
Send, {Tab 3}
Sleep, 50
Clipboard := InvoiceNum
Send, {LControl Down}
Sleep, 50
Send, {v}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Send, {Tab 2}
Sleep, 100
Send, {Enter}
Sleep, 1000
Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Edit?woNo=%WOValue%
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {LAlt Down}
Sleep, 50
Send, {i}
Send, {LAlt Up}
Sleep, 50
Send, {n}
SendRaw, TOT--
Sleep, 150
Send, {Enter}
Sleep, 500
Send, {e}
Sleep, 50
Send, {Tab}
Sleep, 50
Clipboard := ""
Loop
{
    WinActivate, Chrome
    Sleep, 333
    IfWinExist, Chrome
    {
        Break
    }
}
Loop
{
    Send, {LControl Down}
    Sleep, 30
    Send, {a}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 180
    IfInString, Clipboard, Comments
    {
        Break
    }
}
Send, {LControl Down}
Sleep, 50
Send, {LShift Down}
Sleep, 50
Send, {Home}
Sleep, 50
Send, {LShift Up}
Sleep, 50
Send, {LControl Up}
Sleep, 50
Send, {Tab 31}
Sleep, 50
Send, {LControl Down}
Sleep, 30
Send, {a}
Sleep, 50
Send, {c}
Sleep, 50
Send, {LControl Up}
Sleep, 400
Loop
{
    WinActivate, Chrome
    Sleep, 333
    IfWinActive, Chrome
    {
        Break
    }
}
Loop
{
    WinActivate, Invoices
    Sleep, 333
    IfWinActive, Invoices
    {
        Break
    }
}
Send, {LControl Down}{v}
Send, {LControl Up}
Sleep, 100
Send, {Tab}
Sleep, 50
StringLen, Len1, Clipboard
StringGetPos, Pos1, Clipboard, $, R1
StringGetPos, Pos2, Clipboard, HST, R1
Len1 -= Pos2
StringTrimRight, Clipboard, Clipboard, %Len1%
StringTrimLeft, Clipboard, Clipboard, %Pos1%
Send, {LControl Down}{v}
Send, {LControl Up}
Sleep, 100
Send, {Tab}
Sleep, 30
BlockInput, OFF
MsgBox, 0, , 
(LTrim
Task complete. Please finish filling out the invoicing lines. 

Remember to create a holdback invoice if required. 
)
Return

SalesConv(SalesMan)
{
    global
    If Clipboard contains BREEN
    {
        Account := 3104
        Reference := "BREEN"
        Header := "BREEN"
        Initials := "DA"
    }
    If Clipboard contains PETER
    {
        Account := 3101
        Header := "BREEN"
        Reference := "SANDS"
        Initials := "PE"
    }
    If Clipboard contains WEATHERALL
    {
        Account := 3107
        Header := "WEATHE"
        Reference := "WEATHERALL"
        Initials := "WE"
    }
    If Clipboard contains JEFFEREY
    {
        Account := 3108
        Header := "JEFFER"
        Reference := "JEFFEREY"
        Initials := "DJ"
    }
    If Clipboard contains DUNK
    {
        Account := 3110
        Header := "DUNK"
        Reference := "DUNK"
        Initials := "AD"
    }
    If Clipboard contains RUFFIEUX
    {
        Account := 3106
        Header := "RUFFIE"
        Reference := "RUFFIEUX"
        Initials := "ER"
    }
    If Clipboard contains CAHILL
    {
        Header := "CAHILL"
        Reference := "CAHILL"
        Account := 3102
        Initials := "KC"
    }
    If Clipboard contains BRYAN
    {
        Reference := "SANDS B"
        Header := "BRYAN"
        Account := 3114
        Initials := "BS"
    }
    If Clipboard contains THOMPSON
    {
        Reference := "THOMPSON"
        Header := "THOMPS"
        Account := 3121
        Initials := "MT"
    }
    If Clipboard contains WADDEN
    {
        Account := 3119
        Header := "WADDEN"
        Reference := "WADDEN"
        Initials := "DW"
    }
    If Clipboard contains PARIC
    {
        Account := 3122
        Header := "PARIC"
        Reference := "PARIC"
        Initials := "DP"
    }
    If Clipboard contains COWARD
    {
        Account := 3113
        Header := "COWARD"
        Reference := "COWARD"
        Initials := "WC"
    }
    If Clipboard contains Barclay
    {
        Account := 3115
        Header := "BARCLAY"
        Reference := "BARCLAY"
        Initials := "PB"
    }
    If Clipboard contains NELSON
    {
        Reference := "BREEN/NELSON"
        Header := "NELSON"
        Account := 3103
        Initials := "JN"
    }
    If Clipboard contains Owen
    {
        Reference := "OWEN"
        Header := "RJOWEN"
        Account := 3111
        Initials := "RJ"
    }
}


F8::ExitApp
