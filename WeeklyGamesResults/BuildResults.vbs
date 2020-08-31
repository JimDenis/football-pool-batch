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

Dim KeyArray(17)



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

InEmail = ""  
InCount  = 0

 Set fs = CreateObject("Scripting.FileSystemObject")

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesResults\Week1Picks", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesResults\Week1Results", ForWriting, True)

InPicks = tsi.ReadLine

InName = Mid(InPicks,1,3)

If InName = "Key" Then
	InCount = InCount + 1
Else
	WScript.Echo "Error key = " + InName	
End If

KeyAtEnd = "N"
PosCount = 22
KeyArrCnt = 0
Do Until KeyAtEnd = "Y"
	InTeam = Mid(InPicks,PosCount,3)  

	if IsNumeric(InTeam) Then
		WScript.Echo "Is Numeric"
		HoldScore = InTeam
		KeyAtEnd = "Y"
	Else
		KeyArray(KeyArrCnt) = InTeam
		KeyArrCnt = KeyArrCnt + 1
		PosCount = PosCount + 4
	End If	

Loop

WScript.Echo HoldScore 

InPicks = tsi.ReadLine
tso.writeLine InPicks

InPicks = tsi.ReadLine
tso.writeLine InPicks

InPicks = tsi.ReadLine
tso.writeLine InPicks

InPicks = tsi.ReadLine
tso.writeLine InPicks

InPicks = tsi.ReadLine
tso.writeLine InPicks

InPicks = tsi.ReadLine
tso.writeLine InPicks

HoldKeyArrCnt = KeyArrCnt
WScript.Echo "Keys Loaded = " & HoldKeyArrCnt
KeyAtEnd = "N"
KeyArrCnt = 0
Do Until KeyAtEnd = "Y"
	
	If KeyArrCnt >= HoldKeyArrCnt Then
		KeyAtEnd = "Y"	
	Else
		OutLine = OutLine + KeyArray(KeyArrCnt) + " "
'		tso.writeLine OutLine
		KeyArrCnt = KeyArrCnt + 1
	End If	

Loop	

Do Until tsi.AtEndOfStream

	InPicks = tsi.ReadLine
	tso.writeLine InPicks

		PicksAtEnd = "N"
		PosCount = 22
		KeyArrCnt = 0
		CntWins = 0
		'OutLine = "12345678901234567890 "
		OutLine = "Results              "
		
		Do Until KeyArrCnt = HoldKeyArrCnt

			InPick = Mid(InPicks,PosCount,3) 
			If InPick = KeyArray(KeyArrCnt) Then
				OutLine = OutLine & "WIN "
				CntWins = CntWins + 1
			Else
				OutLine = OutLine & "LSE "
			End If

			KeyArrCnt = KeyArrCnt + 1 
			PosCount = PosCount + 4
		Loop	

		'PosCount = PosCount + 4
		InPick = Mid(InPicks,PosCount,3) 
		WScript.Echo "Hold " & HoldScore
		WScript.Echo "Input " & InPick

		Diff = HoldScore - InPick

		OutLine = OutLine & CntWins & " " & Diff 
		tso.writeLine OutLine  

Loop