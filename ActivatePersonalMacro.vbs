'�ϐ��錾
Dim objApp, objExcel, wshShell

	'�l�p�}�N���u�b�N���J������
	Set objApp = CreateObject("Excel.Application.16")
	Set wshShell = WScript.CreateObject("WScript.Shell")

	'�l�p�}�N���u�b�N���J��
	Set objExcel = GetObject("C:\PERSONAL.XLSB")
	
	'MicroPrintToPDF1�����s����
	objExcel.Application.Run "PERSONAL.XLSB!MicroPrintToPDF1.MicroPrintToPDF1"
	
'�I������
Set wshShell = Nothing
Set objExcel = Nothing
Set objApp   = Nothing
WScript.Quit