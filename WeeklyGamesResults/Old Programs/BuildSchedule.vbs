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

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\Football_Pool\WeeklyGamesResults\Data\Week6TeamNames", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\Football_Pool\WeeklyGamesResults\Week6Schedule", ForWriting, True)

OutLine = "                    Welcome to week 6 " 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine

OutLine = "" 
tso.writeLine OutLine


Do Until tsi.AtEndOfStream

	InAwayTeam = tsi.ReadLine
	InCount = InCount + 1
	InHomeTeam = tsi.ReadLine
	InCount = InCount + 1
	GameNum = GameNum + 1

	'WScript.Echo Len(InHomeTeam)
	Disp = Len(InAwayTeam)
	Diff = 12 - Disp
	Filler = String(Diff," ")

	If GameNum < 10 Then
		OutLine = "      Game  " & GameNum & " "  
	Else
		OutLine = "      Game " & GameNum & " " 
	End If	

	OutLine = OutLine & "____ " & InAwayTeam & Filler & " @ " & InHomeTeam & " ____"  
	tso.writeLine OutLine

	OutLine = " "
	tso.writeLine OutLine
	
	OutCount = OutCount + 1
	
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