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


OpeningJobs:
SendMode, Input
SetNumLockState, On
SetScrollLockState, Off
MsgBox, 0, , 
(LTrim
This macro will enter all jobs from the open work order report. Once initiated please do not use the mouse or keyboard until the "task complete" message appears. 

Please Ensure: 
1. Job Costs (with "estimates" and "jobs" tabs opened) is open.
2. Google Chrome is open & logged into the work order system.
3. The open work order report excel sheet is opened.

Hit OK to continue!
)
IfWinNotExist, JobCost
{
    MsgBox, 0, Error, Adagio JobCost does not seem to be opened. Program will now terminate.
    ExitApp
}
IfWinNotExist, OpenWO
{
    MsgBox, 0, , Open work order report does not seem to be open in excel. Program will now terminate. 
    ExitApp
}
Loop
{
    WinActivate, OpenWO
    Sleep, 333
    IfWinActive, OpenWO
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
SendRaw, A1
Sleep, 150
Send, {Enter}
Sleep, 100
Send, {LShift Down}
Sleep, 50
Send, {Down 4}
Sleep, 30
Send, {LShift Up}
Sleep, 50
Send, {LControl Down}
Loop
{
    Send, {NumpadSub}
    IfWinExist, Delete
    {
        Break
    }
}
Send, {LControl Up}
Send, {r}
Sleep, 50
Send, {Enter}
Sleep, 200
Loop
{
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, A1
    Sleep, 100
    Send, {Enter}
    Sleep, 150
    While Clipboard!=""
    {
        Clipboard := ""
    }
    Send, {LControl Down}
    Sleep, 30
    Send, {c}{LControl Up}
    ClipWait
    WOValue := Clipboard
    StringTrimRight, WOValue, WOValue, 2
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, D1
    Sleep, 100
    Send, {Enter}
    Sleep, 150
    While Clipboard!=""
    {
        Clipboard := ""
    }
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 30
    ClipWait
    Description := Clipboard
    StringTrimRight, Description, Description, 2
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, E1
    Sleep, 150
    Send, {Enter}
    Sleep, 100
    While Clipboard!=""
    {
        Clipboard := ""
    }
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 80
    ClipWait
    SalesConv()
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, F1
    Sleep, 150
    Send, {Enter}
    Sleep, 100
    While Clipboard!=""
    {
        Clipboard := ""
    }
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 80
    ClipWait
    ContractValue := Clipboard
    StringTrimRight, ContractValue, ContractValue, 2
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    SendRaw, G1
    Sleep, 150
    Send, {Enter}
    Sleep, 100
    While Clipboard!=""
    {
        Clipboard := ""
    }
    Send, {LControl Down}
    Sleep, 30
    Send, {c}
    Send, {LControl Up}
    Sleep, 80
    ClipWait
    JobCostValue := Clipboard
    StringTrimRight, JobCostValue, JobCostValue, 2
    Loop
    {
        WinActivate, JobCost
        Sleep, 333
        IfWinActive, ahk_exe JobCost.exe
        {
            Break
        }
    }
    Loop
    {
        ControlClick, TSButton12, ahk_exe JobCost.exe,, Left, 1,  x0 y0 NA
        Sleep, 10
        Send, {Alt Down}
        Sleep, 30
        Send, {n}
        Send, {Alt Up}
        Sleep, 30
        IfWinActive, New Job
        {
            Break
        }
    }
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinExist, New Job
        {
            Break
        }
    }
    SendRaw, %WOValue%
    Sleep, 200
    Send, {Tab}
    Sleep, 100
    SendRaw, %Description%
    Sleep, 400
    Send, {Tab 2}
    Sleep, 150
    Send, %Initials%
    Sleep, 100
    BlockInput, ON
    Run, "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" http://win2012:8088/WorkOrder/Customer?woNo=%WOValue%
    While Clipboard!=""
    {
        Clipboard := ""
    }
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
        IfInString, Clipboard, Cust #
        {
            Break
        }
    }
    BlockInput, OFF
    IfInString, Clipboard, ]
    {
        StringGetPos, CustPos1, Clipboard, [, L1, 0
        CustPos1 += 1
        StringTrimLeft, String1, Clipboard, %CustPos1%
        StringGetPos, CustPos, String1, ], R1, 0
        StringLeft, Clipboard, String1, %CustPos%
    }
    Else
    {
        StringGetPos, CustPos, Clipboard, Cust #:, L1, 0
        CustPos += 9
        StringTrimLeft, String1, Clipboard, %CustPos%
        StringGetPos, CustPos, String1, Name:, R1, 0
        CustPos -= 2
        StringLeft, Clipboard, String1, %CustPos%
    }
    Send, {LControl Down}
    Sleep, 30
    Send, {w}
    Sleep, 20
    Send, {LControl Up}
    Sleep, 180
    Loop
    {
        WinActivate, New Job
        Sleep, 333
        IfWinExist, New Job
        {
            Break
        }
    }
    Send, {Tab}
    Sleep, 100
    Send, {LControl Down}
    Sleep, 30
    Send, {v}
    Send, {LControl Up}
    Sleep, 80
    Send, {Tab}
    Sleep, 1000
    IfWinExist, Customer Alert
    {
        WinActivate, Customer Alert
        Sleep, 333
        WinWaitActive, Customer Alert
        Sleep, 333
        Send, {Enter}
    }
    Send, {Tab 10}
    Sleep, 150
    Sleep, 100
    SendRaw, %ContractValue%
    Sleep, 200
    Send, {LAlt Down}
    Sleep, 30
    Send, {o}
    Sleep, 30
    Send, {LAlt Up}
    Sleep, 30
    WinWaitNotActive, New Job
    Sleep, 333
    IfWinExist, Error
    {
        Send, {Enter}
        Sleep, 100
        Send, {Escape}
        Sleep, 500
        Send, {n}
        Sleep, 50
    }
    Else
    {
        Sleep, 500
        Loop
        {
            WinActivate, JobCost
            Sleep, 333
            IfWinActive, JobCost
            {
                Break
            }
        }
        Loop
        {
            Send, {LControl Down}
            Sleep, 30
            Send, {Tab}
            Sleep, 30
            Send, {LControl Up}
            Sleep, 30
            Send, {Alt Down}
            Sleep, 50
            Send, {n Down}
            Sleep, 50
            Send, {Alt Up}
            Sleep, 50
            IfWinActive, New Job Estimate
            {
                Break, 1
            }
            Send, {Escape}
            Sleep, 50
        }
        SendRaw, %WOValue%
        Sleep, 200
        Send, {Tab}
        Sleep, 100
        Send, {1}
        Sleep, 100
        Send, {Tab}
        Sleep, 100
        Send, {1}
        Sleep, 50
        Send, {Tab 5}
        Sleep, 100
        SendRaw, %JobCostValue%
        Sleep, 100
        Send, {Enter}
        Sleep, 100
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
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    Sleep, 300
    SendRaw, A1
    Send, {Enter}
    Sleep, 300
    Send, {LControl Down}{NumpadSub}{LControl Up}
    Sleep, 50
    Send, {r}{Enter}
    Sleep, 150
    Loop
    {
        Send, {F5}
        IfWinExist, Go To
        {
            Break
        }
    }
    Sleep, 100
    SendRaw, A1
    Send, {Enter}
    Sleep, 150
    Send, {LControl Down}{c}{LControl Up}
    Sleep, 150
    IfNotInString, Clipboard, 1
    {
        IfNotInString, Clipboard, 9
        {
            IfNotInString, Clipboard, 8
            {
                Break
            }
        }
    }
}
MsgBox, 0, , Task Complete!
Return

SalesConv()
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
