Dim tsi, InCount
Dim fs, tso, OutLine, OutLine2, OutCount
Dim GameNum

Dim ArrayCounter
Dim TeamNameList
	TeamNameList = Array("Cowboys", "Bears", "Panthers", "Falcons", "Ravens", "Bills", "Bengals", "Browns", "Redskins", "Packers", "Broncos", _
						 "Texans", "Lions", "Vikings", "49ers", "Saints", "Dolphins", "Jets", "Colts", "Buccaneers", "Chargers", "Jaguars", _
						 "Steelers", "Cardinals", "Chiefs", "Patriots", "Titans", "Raiders", "Seahawks", "Rams", "Giants", "Eagles")


Dim TeamNameShort  
	TeamNameShort = Array("DAL", "CHI", "CAR", "ATL", "BAL", "BUF", "CIN", "CLE", "WAS", "GB ", "DEN", _
						  "HOU", "DET", "MIN", "SF ", "NO ", "MIA", "NYJ", "IND", "TB ", "LAC", "JAC", _
						  "PIT", "ARI", "KC ", "NE ", "TEN", "LV ", "SEA", "LAR", "NYG", "PHI")

Dim HomeArray(16)	
Dim AwayArray(16)					  



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

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\code\homework\football_1\src\data\Week7.js", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\Football_Pool\WeeklyGamesResults\Week7Picks", ForWriting, True)

InAwayTeam = tsi.ReadLine
InAwayTeam = tsi.ReadLine

Do Until tsi.AtEndOfStream

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
			InHomeTeam = x 
		End If	
	next

	SplitCount = 0
	c=Split(HoldAway,"""")
	for each x in c
		SplitCount = SplitCount + 1 
		if SplitCount = 2 Then
			InAwayTeam = x 
		End If	
	next

	InCount = InCount + 2
	GameNum = GameNum + 1

	ArrayCounter = 0
	Do While ArrayCounter < 32
		' WScript.Echo "X is " & X
		' WScript.Echo "Array is " & 	TeamListBad(ArrayCounter)	
		 If InAwayTeam =  TeamNameList(ArrayCounter) Then
			AwayTeamOut = TeamNameShort(ArrayCounter) 
		 	ArrayCounter = 33
		Else
			ArrayCounter = ArrayCounter + 1
			If ArrayCounter = 32 Then
				WScript.Echo "Unmatched Team " & InAwayTeam	
			End If
		End If  
	Loop   

	
	OutLine = OutLine _ 
	 		  & AwayTeamOut & " " 

		ArrayCounter = 0
		Do While ArrayCounter < 32
		' WScript.Echo "X is " & X
		' WScript.Echo "Array is " & 	TeamListBad(ArrayCounter)	
		 If InHomeTeam =  TeamNameList(ArrayCounter) Then
			HomeTeamOut = TeamNameShort(ArrayCounter) 
		 	ArrayCounter = 33
		Else
			ArrayCounter = ArrayCounter + 1
			If ArrayCounter = 32 Then
				WScript.Echo "Unmatched Team " & InHomeTeam	
			End If
		End If  
	Loop   
	 

	OutLine2 = OutLine2 _ 
	 		  & HomeTeamOut & " " 

	
	OutCount = OutCount + 1

	InAwayTeam = tsi.ReadLine
	LenOfInput = Len(InAwayTeam)

	'if LenOfInput < 10 Then
	'	InAwayTeam = tsi.ReadLine
	'End If	

	
Loop

tso.writeLine OutLine
tso.writeLine OutLine2
tso.writeLine(String(95,"="))
tso.writeLine OutLine3
tso.writeLine(String(95,"="))

OutLine  = "" 

WScript.Echo "Teams Input is " & InCount
WScript.Echo "Games Output is " & OutCount 