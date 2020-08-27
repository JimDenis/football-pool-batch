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

InEmail = ""  
InCount  = 0

 Set fs = CreateObject("Scripting.FileSystemObject")

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesResults\EmailPicks", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesResults\Week1Results", ForWriting, True)

Do Until tsi.AtEndOfStream

	InEmail = tsi.ReadLine
	InCount = InCount + 1

	WhereIsAt = InStr(InEmail," is ")

	If WhereIsAt > 0 Then
		WhereIsAt = WhereIsAt + 4
		TeamIn = mid(InEmail,WhereIsAt,22)
	End If	

	If WhereIsAt = 14 Then
		Disp = Len(TeamIn)
		WScript.Echo Disp
		Diff = 22 - Disp
		Filler = String(Diff," ")
		OutLine = OutLine + TeamIn + Filler
		WhereIsAt = 10
	End If

	If WhereIsAt = 27 Then
		Disp = Len(TeamIn)
		WScript.Echo Disp
		HoldPoints = TeamIn 
		WhereIsAt = 10
	End If

	If WhereIsAt > 10 Then
				
		ArrayCounter = 0
		Do While ArrayCounter < 32

		 	If TeamIn = TeamNameList(ArrayCounter) Then
			 	TeamShort = TeamNameShort(ArrayCounter) 
				OutLine = OutLine + TeamShort + " "
				ArrayCounter = 33
			Else
				ArrayCounter = ArrayCounter + 1
			End If

		Loop   

	End If 

	If WhereIsAt = 0 Then
		OutLine = OutLine + " " + HoldPoints
		tso.writeLine OutLine
		OutLine = ""
	End If

Loop

WScript.Echo "Input  count is " & InCount
WScript.Echo "Output count is " & OutCount 