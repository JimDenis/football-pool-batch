Dim tsi, InCount
Dim fs, tso, OutLine, OutLine2, OutCount
Dim GameNum

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
GameNum = 0
InAwayTeam = ""
InHomeTeam = ""
OutCount = 0
OutLine  = "AWAY                 "
OutLine2 = "HOME                 "
OutLine3 = "PLAYER                                                                              POINTS"


 Set fs = CreateObject("Scripting.FileSystemObject")

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\code\homework\football_1\src\data\Week6.js", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\Football_Pool\WeeklyGamesResults\Week6Schedule", ForWriting, True)

OutLine = "                    Welcome to week 6 " 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

InAwayTeam = tsi.ReadLine


Do Until tsi.AtEndOfStream

	InAwayTeam = tsi.ReadLine

	SplitCount = 0
	a=Split(InAwayTeam,":")
	for each x in a
		SplitCount = SplitCount + 1 
		'OutLine = x & " " & SplitCount
		if SplitCount = 3 Then 
			HoldHome = x
		End If	
		if SplitCount = 4 Then 
			HoldAway = x
		End If	
	next

	SplitCount = 0
	b=Split(HoldHome,"""")
	for each x in b
		SplitCount = SplitCount + 1 
		if SplitCount = 2 Then
			HoldHome = x 
		End If	
	next

	SplitCount = 0
	c=Split(HoldAway,"""")
	for each x in c
		SplitCount = SplitCount + 1 
		if SplitCount = 2 Then
			HoldAway = x 
		End If	
	next


	InCount = InCount + 1
	InCount = InCount + 1
	GameNum = GameNum + 1

	'WScript.Echo Len(InHomeTeam)
	Disp = Len(HoldAway)
	Diff = 12 - Disp
	Filler = String(Diff," ")

	If GameNum < 10 Then
		OutLine = "      Game  " & GameNum & " "  
	Else
		OutLine = "      Game " & GameNum & " " 
	End If	

	If disp > 0 Then

		OutLine = OutLine & "____ " & HoldAway & Filler & " @ " & HoldHome & " ____"  
		tso.writeLine OutLine

		OutLine = " "
		tso.writeLine OutLine
	
		OutCount = OutCount + 1

	End If

	HoldHome = ""
	HoldAway = "" 
	
Loop

OutLine = "" 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

OutLine = "Your Team Name ______________________________________"  
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

OutLine = "Your Tie Breaking Score _____________________________" 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

WScript.Echo "Teams Input is " & InCount
WScript.Echo "Games Output is " & OutCount 