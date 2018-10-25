Dim strInputPath, strOutputPath, strStatus
Dim objFSO, objTextIn, objTextOut1,objTextOut2,colMatches,tmpStr
Dim wshShell,oResult
Const ForReading = 1


'Varnow = now
'vardate = Day(varnow) & "-" & Month(varnow) & "-" & Year(varnow) & ".log"


dToday = Date()
sToday = Right("0" & Day(dToday), 2) & MonthName(Month(dToday), True) & Year(dToday)

'LogPath = ("C:\MY Logs\" & VarDate)

strInputPath = "C:\Users\kollam\Desktop\IMPFiles\LoadComputersList.txt" 'Change the path
strOutputPath1 = "C:\Users\kollam\Desktop\IMPFiles\TRaceRouteLog_ " & sToday & ".txt" 'Change the location but not the file name
'strOutputPath2 = "C:\Users\kollam\Desktop\IMPFiles\LoadComputersStatus2.txt" 'Change the location but not the file name

Set wshShell = WScript.CreateObject ("WSCript.shell")

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objTextIn = objFSO.OpenTextFile( strInputPath, ForReading)

Set objTextOut1 = objFSO.CreateTextFile( strOutputPath1 )

'Set objTextOut3 = objFSO.CreateTextFile( strOutputPath2 )

'objTextOut3.WriteLine("computerName,IPAddress,status")

Do Until objTextIn.AtEndOfStream = True

	strLine = objTextIn.ReadLine

	If Instr(strLine,"||") Then
	arrFields = Split(strLine, "||")

     	strIpAddress = arrFields(0)
     	strComputer =arrFields(1)
		
	Set oResult = wshShell.Exec("tracert -d " & strIpAddress) 
	
	Do Until oResult.StdOut.AtEndOfStream = True


        objTextOut1.WriteLine oResult.StdOut.Readline 

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objTextOut2 = objFSO.OpenTextFile(strOutputPath1, ForReading)


	'Do While objTextOut2.AtEndOfStream = False

	'ln1 = objTextOut2.Readline

	'str1 = Mid(ln1,9,1)

    	'If str1 = "Request TimeOut" Then

        	'objTextOut3.Write strIpAddress & "||"	
		'objTextOut3.Write strComputer & "||" 
        	'objTextOut3.Write "Offline" & ""
    	'Else
		'objTextOut3.Write strIpAddress & "||"	
		'objTextOut3.Write strComputer & "||" 
        	'objTextOut3.Write "Online" & ""
        'Exit Do
    'End If


'Loop


'Wscript.Quit
Loop
	
End If
Loop

Set objOutlook = CreateObject("Outlook.Application")
'Set myItem = objOutlook.CreateItem(olNoteItem)

Set objMail = objOutlook.CreateItem(0)
objMail.Display   'To display message
objMail.To = "madhavilatha.kolla@cognizant.timeinc.com" 'Change the email addddress
objMail.cc = "madhavilatha.kolla@cognizant.timeinc.com;madhavilatha.kolla@cognizant.timeinc.com" ' Chnage the email address
objMail.Subject = "Trace Status of the LoadComputers <Auto Generated Email>"
objMail.Body = "Hi All, Please find the tracelog of all the loadcomputers below -    " & objTextOut2.ReadAll &   "Note: If you see Request Timeout from any of the above computers trace log, pelase condiser them as Down status /Unable to connect"
objMail.Attachments.Add(strOutputPath1) 
objMail.Attachments.Add(strInputPath)
objMail.Send
objOutlook.Quit
Set objMail = Nothing
Set objOutlook = Nothing





