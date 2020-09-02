# The Football Pools batch jobs

Created two groups of batch jobs (VBScript). One group (WeeklyGamesBuild) takes the input for a website & creates a weekly football schedule for the year. This schedule will feed the on-line screen (web based) Football Pool. The second set of jobs (WeeklyGamesResults) takes in the picks from The Football Pool and outputs all the playes picks for that week. This second group of job will also, after all the games complete, will count up the playes win & determine a winner.

### Instructions:

-   To create a weekly schedule (WeeklyGamesBuild)
-   Cut schedule from https://fbschedules.com/nfl-schedule/ & place into a text file
-   Enter correct week into job RawToCity.vbs & then run it
-   Enter correct week into job CityToTeam2Lines.vbs & then run it
-   Enter correct week into job BuildWeeklyInput.vbs & then run it
-   Move output from above Week#Input to football_1/src/data/Week#.js
-   Added "export let week# = [" to the top of football_1/src/data/Week#.js
-   Added "[;" to the bottom of football_1/src/data/Week#.js
-
-   To create a weekly picks & winners (WeeklyGamesResults)
-   Cut returned e-mails & place into a text file, Can be one to many in the file
-   Enter correct week into job BuildPicksHeader.vbs & then run it
-   Enter correct input file into job BuildPicksDetail.vbs & then run it
-   Add NFL winner into the above & run it. Then move that line to the top
-   Enter correct week into job BuildPicksResults.vbs & then run it

### Prerequisites

None

### Installing

Installed using GitHub with following commands:

-   git add -A
-   git commit -m"comment goes here"
-   git push

## Running the tests

No automated testing

## Built With

-   VBScript

## Authors

-   **Jim Denis** - _Initial work_ - [JimDenis](https://github.com/JimDenis)

This app can is not on-line
