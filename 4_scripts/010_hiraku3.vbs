'Option Explicit

'Const vbNormalNoFocus = 4    '通常のウィンドウで、最前面ではない

Dim objWShell

arg1 = "FAX_Order_Summary_1.csv"

' あんたバカねぇ！Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

objWshShell.Run "excel.exe " & arg1

WScript.Sleep 1000

arg2 = "FAX_Order_Summary_11.xlsx"

' あんたバカねぇ！Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

objWshShell.Run "excel.exe " & arg2

WScript.Sleep 1000

arg3 = "FAX_Order_Summary_2.csv"

' あんたバカねぇ！Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

objWshShell.Run "excel.exe " & arg3

WScript.Sleep 1000

arg4 = "FAX_Order_Summary_22.xlsx"

' あんたバカねぇ！Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

objWshShell.Run "excel.exe " & arg4

WScript.Sleep 1000

arg5 = "df_todays_orders_display.csv"

' あんたバカねぇ！Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

objWshShell.Run "excel.exe " & arg5