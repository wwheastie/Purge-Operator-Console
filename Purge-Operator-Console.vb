sub zmain()
	'Verify user would like to run
	messageBox1 = msgbox("Are you sure you would like to run this macro?", 3, "Purge NetStat")
	
	'Run macro if user selects "Yes"
	if messageBox1 = "Yes" then
	
		'Declare variables
		dim rowCount, columnCount, getHowMany, stringText, deleteString
		
		'Set rowCount, columnCount, and getHowMany
		rowCount = 3
		columnCount = 2
		getHowMany = 25
		
		'Send cursor home and pause screen
		sendhostkeys("<Home>PAUSE<Enter>")
		
		'Row loop
		do until rowCount = 21
			'Column loop
			do until columnCount > 54
			stringText = GetString(rowCount, columnCount, getHowMany)
			stringText = Replace(stringText, "?", " ")
			'Check to see if PR or IN is OFF, and remove anything that isn't
			if mid(stringText, 1, 2) = "PR" OR mid(stringText, 1, 2) = "IN" then
				if mid(stringText, 10, 3) <> "OFF" AND mid(stringText, 10, 4) <> "STOP" AND mid(stringText, 1, 7) <> "       " then
					deleteString = "DELETE '" + mid(stringText, 1, 8) + mid(stringText, 10, 5) + "',PURGE<ENTER>"
					sendhostkeys(deleteString)
					WaitForHostUpdate(10)
				end if
			'Remove anything else that isn't PR or IN and OFF or DOWN (for visa stations)
			elseif mid(stringText, 10, 3) <> "OFF" AND mid(stringText, 10, 4) <> "STOP" AND mid(stringText, 10, 4) <> "DOWN" AND mid(stringText, 1, 7) <> "       " then
					deleteString = "DELETE '" + mid(stringText, 1, 8) + mid(stringText, 10, 5) + "',PURGE<ENTER>"
					sendhostkeys(deleteString)
					WaitForHostUpdate(10)
			end if
			'Check rowCount, columnCount, and stringText conditions
			if rowCount = 20 and columnCount = 54 and stringText <> "                         " then
				sendhostkeys("<HOME><ENTER>")
				WaitForHostUpdate(10)
				rowCount = 3
				columnCount = 2
			else	
			'Add to columnCount
			columnCount = columnCount + 26
			end if
			loop
		'Reset columnCount and add 1 to rowCount
		columnCount	= 2
		rowCount = rowCount + 1
		loop	

		'Send cursor home and change pause to 15
		sendhostkeys("<Home>Pause 15<Enter>")
		sendhostkeys("<Enter>")
		
	end if	
end sub