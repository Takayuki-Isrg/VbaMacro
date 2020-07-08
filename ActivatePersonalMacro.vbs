'変数宣言
Dim objApp, objExcel, wshShell

	'個人用マクロブックを開く準備
	Set objApp = CreateObject("Excel.Application.16")
	Set wshShell = WScript.CreateObject("WScript.Shell")

	'個人用マクロブックを開く
	Set objExcel = GetObject("C:\PERSONAL2.XLSB")
	
	'MicroPrintToPDF1を実行する
	objExcel.Application.Run "PERSONAL2.XLSB!MicroPrintToPDF1.MicroPrintToPDF1"
	
'終了処理
Set wshShell = Nothing
Set objExcel = Nothing
Set objApp   = Nothing
WScript.Quit