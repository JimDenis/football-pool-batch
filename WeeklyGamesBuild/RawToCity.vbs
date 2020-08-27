Dim fs, tsi, tso
Dim OutLine, OutlineCnt, OutCount
Dim InLine, InLineCnt, InWord, InCount

Const ForReading         = 1
Const ForWriting         = 2
Const ForAppending       = 8
Const TristateUseDefault = -2
Const TristateTrue       = -1
Const TristateFalse      = 0

Dim vbQuestion: vbQuestion=32
Dim vbYesNo: vbYesNo=4 
Dim vbYes: vbYes=6
Dim vbNo: vbNo=7 

InCount  = 0
OutCount = 0
InLineCnt = 0
OutlineCnt = 0 

 Set fs = CreateObject("Scripting.FileSystemObject")

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesBuild\WeekPS4Raw", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesBuild\WeekPS4CityNames", ForWriting, True)

Do Until tsi.AtEndOfStream

	InLine = tsi.ReadLine
	InLineCnt = InLineCnt + 1
	InWord = Split(InLine,vbTab) 
	InLineCnt = 0

	for each x in InWord
		InLineCnt = InLineCnt + 1
		OutlineCnt = 0

		if InLineCnt = 1 Then 
			Outline = " "
			OutLine = x
			OutlineCnt = InStr(Outline," at ")
		'	OutLine = OutLine & " " & OutlineCnt
		End If	

		if  OutlineCnt > 0 Then 
			OutLine = Replace(OutLine," at ","?")
			OutLine = Replace(OutLine," ","") 
			OutLine = Replace(OutLine,"?"," ")
			tso.writeLine OutLine
			OutCount = OutCount + 1
		End If	
	next

loop

WScript.Echo "Input count is " & InCount
WScript.Echo "Output count is " & OutCount 