# VBA-Challenge 2
 Module 2 challenge
I worked on this challenge with Jaxon Keller

Sources
	Creditcard checker
	Student census
	CheckerBoard 
We did look up how to use a "lastrow = ws.Cells(Rows.Count, 1).End(x1Up).Row" on Google

Linking all worksheets was found in Slack notes during class 
	Nick Sneed
	  9:19 PM
	' Steps:
	' ----------------------------------------------------------------------------
	' Part I:
	' 1. Extract the number before the phrase "_census_data" to figure out the year.
	' 2. Add the year to the first column of each spreadsheet.
	' 3. Split the "Place" column into "County" and "State".
	' 4. Convert the household and per capita income columns to currency values for all cells.
	Sub Census_pt1()
	    ' --------------------------------------------
	    ' LOOP THROUGH ALL SHEETS
     		 For Each ws. In Worksheets...
	
