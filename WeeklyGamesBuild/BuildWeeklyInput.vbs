Dim tsi, InChar, InCount
Dim fs, tso, OutLine, OutCount
Dim GameNum

Dim ArrayCounter
Dim TeamListGood
	TeamListGood = Array("Cowboys", "Bears", "Panthers", "Falcons", "Ravens", "Bills", "Bengals", "Browns", "Redskins", "Packers", "Broncos", _
						 "Texans", "Lions", "Vikings", "49ers", "Saints", "Dolphins", "Jets", "Colts", "Buccaneers", "Chargers", "Jaguars", _
						 "Steelers", "Cardinals", "Chiefs", "Patriots", "Titans", "Raiders", "Seahawks", "Rams", "Giants", "Eagles")


Dim TeamListBad  
	TeamListBad = Array("Dallas", "Chicago", "Carolina", "Atlanta", "Baltimore", "Buffalo", "Cincinnati", "Cleveland", "Washington", "GreenBay", "Denver", _
						"Houston", "Detroit", "Minnesota", "SanFrancisco", "NewOrleans", "Miami", "NYJets", "Indianapolis", "Tampa", "LAChargers", "Jacksonville",_
						"Pittsburgh", "Arizona", "KansasCity", "NewEngland", "Tennessee", "LasVegas", "Seattle", "LARams", "NYGiants", "Philadelphia")

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
ArrayCounter = 0
GameNum = 0
OutCount = 0
Output = ""
InAwayTeam = ""
InHomeTeam = ""


 Set fs = CreateObject("Scripting.FileSystemObject")

 Set tsi = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesBuild\WeekPS2TeamNames", ForReading)
 Set tso = fs.OpenTextFile("C:\Users\jimde\Desktop\hold_folder_react_app\WeeklyGamesBuild\WeekPS2Input", ForWriting, True)

Do Until tsi.AtEndOfStream

	InAwayTeam = tsi.ReadLine
	InHomeTeam = tsi.ReadLine
	GameNum = GameNum + 1

	OutLine = "{" _
			  & " " & "num:" & " " &  """" & GameNum & """," _ 
			  & " " & "home:" & " " & """" & InHomeTeam & """," _
			  & " " & "away:" & " " & """" & InAwayTeam & """," _
			  & " " & "picked:" & " " & """" & "none" & ""","  _
			  & " " & "}" & ","


	tso.writeLine OutLine
	OutCount = OutCount + 1
	OutLine = ""

	
	
Loop

WScript.Echo "Output count is " & OutCount 