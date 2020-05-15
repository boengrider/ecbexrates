''============================================================
'' Program:   01_Exchange rate processing
'' Desc:      Automation of exchange rate import to SAP
'' Called by: End user OR task scheduler
'' Call:      wscript 01_ExRate.vbs
'' Arguments: none
'' Comments:  Downloads XML from ECB web
''	      https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml
''	      Parses the file 	
''            and outputs formatted file that can be imported		
'' 	      by SAP
'' Changes---------------------------------------------------
'' Date			Programmer   	Change		Contact
'' 2020-04-21	Tomas Ac     	Written		tomas.ac@volvo.com;tomasac.22@gmail.com
'' 
'' 
''								
'' 
''								
''


Option Explicit
'On Error Resume Next

' ------------------------------------------------------------------------
' ----------------- H o u s e k e e p i n g ------------------------------
' ------------------------------------------------------------------------
Const Xml = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml" ' Todays rate
Const Xml90 = "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-hist-90d.xml?a3173423c4ae84dd89e4c898d0313231" '90 Days
Const dirPath = "ExRate"
Const ns = "xmlns:gesmes='http://www.gesmes.org/xml/2002-08-01' xmlns='http://www.ecb.int/vocabulary/2002-08-01/eurofxref'" ' XML namespace
Dim oFSO,oWSH,oXML,oHTTP,oFD,oTCD,oXML90,oLogfile,oRemoteLog,oNET,loops
Dim n_ChildNodes,n_ChildNode,a_Attributes,a_Attribute,strFname,strDate,strDrive,strCubeDate,strToday,strUser,strComputer,logremote
Set oNET = CreateObject("Wscript.Network")
Set oTCD = CreateObject("Scripting.Dictionary") ' Associative array holding ECB holiday calendar
Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set oWSH = CreateObject("WScript.Shell")
Set oXML = CreateObject("MSXML2.DOMDocument")   ' Current ECB rate, valid for next working day
Set oXML90 = CreateObject("MSXML2.DOMDocument") ' Historical ECB rates, up to 90 days
Set oHTTP = CreateObject("MSXML2.XMLHTTP")
strDrive = oWSH.ExpandEnvironmentStrings("%SYSTEMDRIVE%")
loops = 0
Call MakeOutputDir(strDrive & "\ExRate") ' Make output directory
Set oLogfile = oFSO.OpenTextFile(strDrive & "\ExRate\" & "log.txt",8,True)
strToday = Date()
strDate = FindDate(strToday) ' Call this ONCE and work with variable
strComputer = oNET.ComputerName
strUser = oNET.UserName


' ECB published TCD (Target closing day) calendar
' On these days there will be no new exchange rate 
' published at 16:00 that is typically valid
' for next working day

'2020
oTCD.Add "01012020","New Year's Day"
oTCD.Add "10042020","Good Friday"
oTCD.Add "13042020","Easter Monday"
oTCD.Add "01052020","Labour Day"
oTCD.Add "09052020","Anniversary of Robert Schuman's Declaration"
oTCD.Add "21052020","Ascension Day"
oTCD.Add "01062020","Whit Monday"
oTCD.Add "11062020","Corpus Cristy"
oTCD.Add "03102020","Day of German Unity"
oTCD.Add "01112020","All Saints' Day"
oTCD.Add "24122020","Christmas Eve"
oTCD.Add "25122020","Christmas Day"
oTCD.Add "26122020","Christmas Holiday"
oTCD.Add "31122020","New Year's Eve"
'2021
oTCD.Add "01012021","New Year's Day"
oTCD.Add "02042021","Good Friday"
oTCD.Add "05042021","Easter Monday"
oTCD.Add "01052021","Labour Day"
oTCD.Add "09052021","Anniversary of Robert Schuman's Declaration"
oTCD.Add "13052021","Ascension Day"
oTCD.Add "24052021","Whit Monday"
oTCD.Add "03062021","Corpus Cristy"
oTCD.Add "03102021","Day of German Unity"
oTCD.Add "01112021","All Saints' Day"
oTCD.Add "24122021","Christmas Eve"
oTCD.Add "25122021","Christmas Day"
oTCD.Add "26122021","Christmas Holiday"
oTCD.Add "31122021","New Year's Eve"
'2022
oTCD.Add "01012022","New Year's Day"
oTCD.Add "15042022","Good Friday"
oTCD.Add "18042022","Easter Monday"
oTCD.Add "01052022","Labour Day"
oTCD.Add "09052022","Anniversary of Robert Schuman's Declaration"
oTCD.Add "26052022","Ascension Day"
oTCD.Add "06062022","Whit Monday"
oTCD.Add "16062022","Corpus Cristy"
oTCD.Add "03102022","Day of German Unity"
oTCD.Add "01112022","All Saints' Day"
oTCD.Add "24122022","Christmas Eve"
oTCD.Add "25122022","Christmas Day"
oTCD.Add "26122022","Christmas Holiday"
oTCD.Add "31122022","New Year's Eve"
' ------------------------------------------------------------------------
' ----------------- H o u s e k e e p i n g   e n d ----------------------
' ------------------------------------------------------------------------



































' =============================================================
' =================== M A I N =================================
' =============================================================


' Check if it is TCD and if it is the correct time. If not,  nothing
' and instruct the user to come later and repeat the action
' Change Hour and Minute as required. 
' ECB publishes on non TCD weekdays at 16:00 CET / 4:00 PM CET
' I run this script at 17:00 CET/ 5:00 PM CET on weekdays

If Hour(Time()) < 17 And Minute(Time()) < 59 Then 
		LogEvent oLogfile,"Time less than 17:00","WARNING"
		LogEvent oLogfile,"Script exited with code: 3","INFORMATION"
		WScript.Quit(3) ' Bad time
End If 


' If the below code executes it means it is not a TCD and it is the correct time i.e after 17:00


' Download the latest XML
oHTTP.open "GET", Xml,False
oHTTP.send

If oHTTP.status <> 200 Then
	LogEvent oLogfile,"Error downloading XML file. HTTP error: " & oHTTP.status,"ERROR"
	LogEvent oLogfile,"Script exited with code: " & oHTTP.status,"INFORMATION"
	WScript.Quit(oHTTP.status)
Else
	oXML.load(oHTTP.responseXML)   ' Load XML object
	oXML.setProperty "SelectionNamespaces", ns ' Set proper namespace
End If 

' Download historical XML (90 days)
oHTTP.open "GET", Xml90,False
oHTTP.send

If oHTTP.status <> 200 Then
	LogEvent oLogfile,"Error downloading XML file. HTTP error: " & oHTTP.status,"ERROR"
	LogEvent oLogfile,"Script exited with code: " & oHTTP.status,"INFORMATION"
	WScript.Quit(oHTTP.status)
Else
	oXML90.load(oHTTP.responseXML) ' Load XML90 object
	oXML90.setProperty "SelectionNamespaces", ns ' Set proper namespace
End If


' If the below code executes it means both files were downloaded successfully.
' We have both files downloaded




' Sunday -> 1
' Monday -> 2
' Tuesday -> 3
' ...
' Saturday -> 7
' I don't check anything regarding system settings
' First day of week is sunday (1) by default I guess



Select Case Weekday(strToday)


		' On Monday we're looking for the new exchange rate published at 16:00 on Monday. This rate will be used on Tuesday i.e Next working day
		Case 2 ' Monday
		LogEvent oLogfile, "Determining the weekday: Monday","INFORMATION"
		
		If loops < 1 Then ' We will be using todays XML
			If Not ValidateXMLDate(oXML) Then  ' Date() and Cube time are not the same. Critical  Error. Quit
				LogEvent oLogfile, "Invalid XML date. Expected:" & strDate & " Found:" & strCubeDate,"ERROR"
				WScript.Quit(4) ' Bad Cube date
			End If
			LogEvent oLogfile, "Today ( " & strToday & " ) is not a TCD","INFORMATION"
			LogEvent oLogfile, "Valid XML date " & strCubeDate,"INFORMATION"
			Set oFD = oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			'ParseXML
			ParseXML strDate,oFD,oXML
		Else ' Use XML90
			Set oFD =  oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			ParseXML90 strDate,oFD,oXML90
			LogEvent oLogfile, "Today ( " & strToday & " ) is a TCD","INFORMATION"
			LogEvent oLogfile, "Parsing XML90","INFORMATION"
		End If 
		
		
		
		
		
	' On Tuesday we're looking for the new exchange rate published at 16:00 on Tuesday. This rate will be used on Wednesday i.e Next working day
	Case 3 ' Tuesday
		LogEvent oLogfile, "Determining the weekday: Tuesday","INFORMATION"
		

		If loops < 1 Then ' We will be using todays XML
			If Not ValidateXMLDate(oXML) Then  ' Date() and Cube time are not the same. Critical  Error. Quit
				LogEvent oLogfile, "Invalid XML date. Expected:" & strDate & " Found:" & strCubeDate,"ERROR"
				LogEvent oLogfile,"Script exited with code: 4","INFORMATION"
				WScript.Quit(4) ' Bad Cube date
			End If
			LogEvent oLogfile, "Today ( " & strToday & " ) is not a TCD","INFORMATION"
			LogEvent oLogfile, "Valid XML date " & strCubeDate,"INFORMATION"
			Set oFD = oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			'ParseXML
			ParseXML strDate,oFD,oXML
		Else ' Use XML90
			Set oFD =  oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			ParseXML90 strDate,oFD,oXML90
			LogEvent oLogfile, "Today ( " & strToday & " ) is a TCD","INFORMATION"
			LogEvent oLogfile, "Parsing XML90","INFORMATION"
		End If 
		
		
		
		
		
		
	' On Wednesday we're looking for the new exchange rate published at 16:00 on Wednesday. This rate will be used on Thursday i.e Next working day	
	Case 4 ' Wednesday
		LogEvent oLogfile, "Determining the weekday: Wednesday","INFORMATION"
		

		If loops < 1 Then ' We will be using todays XML
			If Not ValidateXMLDate(oXML) Then  ' Date() and Cube time are not the same. Critical  Error. Quit
				LogEvent oLogfile, "Invalid XML date. Expected:" & strDate & " Found:" & strCubeDate,"ERROR"
				LogEvent oLogfile,"Script exited with code: 4","INFORMATION"
				WScript.Quit(4) ' Bad Cube date
			End If
			LogEvent oLogfile, "Today ( " & strToday & " ) is not a TCD","INFORMATION"
			LogEvent oLogfile, "Valid XML date " & strCubeDate,"INFORMATION"
			Set oFD = oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			'ParseXML
			ParseXML strDate,oFD,oXML
		Else ' Use XML90
			Set oFD =  oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			ParseXML90 strDate,oFD,oXML90
			LogEvent oLogfile, "Today ( " & strToday & " ) is a TCD","INFORMATION"
			LogEvent oLogfile, "Parsing XML90","INFORMATION"
		End If 
		
		
		
		
		
	' On Thursday we're looking for the new exchange rate published at 16:00 on Thursday. This rate will be used on Friday i.e Next working day
	' AND Saturday and Sunday
	Case 5 ' Thursday. Filename +1 +2 + 3
		LogEvent oLogfile, "Determining the weekday: Thursday","INFORMATION"
		
		
		If loops < 1 Then ' We will be using todays XML
			If Not ValidateXMLDate(oXML) Then  ' Date() and Cube time are not the same. Critical  Error. Quit
				LogEvent oLogfile, "Invalid XML date. Expected:" & strDate & " Found:" & strCubeDate,"ERROR"
				LogEvent oLogfile,"Script exited with code: 4","INFORMATION"
				WScript.Quit(4) ' Bad Cube date
			End If
			LogEvent oLogfile, "Today ( " & strToday & " ) is not a TCD","INFORMATION"
			LogEvent oLogfile, "Valid XML date " & strCubeDate,"INFORMATION"
			Set oFD = oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for Friday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt","INFORMATION"
			'ParseXML
			ParseXML strDate,oFD,oXML
			
			oFSO.CopyFile strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", strDrive & "\ExRate\" & FormatDate(strToday + 2,2) & ".txt" ' Copy Friday file for Saturday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 2,2) & ".txt","INFORMATION"
			oFSO.CopyFile strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt" ' Copy Friday file for Sunday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt","INFORMATION"
		Else ' Use XML90
			Set oFD =  oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", True) ' Make file for Friday
			ParseXML90 strDate,oFD,oXML90
			LogEvent oLogfile, "Today ( " & strToday & " ) is a TCD","INFORMATION"
			LogEvent oLogfile, "Parsing XML90","INFORMATION"
			
			oFSO.CopyFile strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", strDrive & "\ExRate\" & FormatDate(strToday + 2,2) & ".txt" ' Copy Friday file for Saturday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 2,2) & ".txt","INFORMATION"
			oFSO.CopyFile strDrive & "\ExRate\" & FormatDate(strToday + 1,2) & ".txt", strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt" ' Copy Friday file for Sunday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt","INFORMATION"
		End If 
		
		
		
		
		
	' On Friday we're looking for the new exchange rate published at 16:00 on Friday. This rate will be used on Monday i.e Next working day
	' Monday is the first working day following Friday
	Case 6 ' Friday
	LogEvent oLogfile, "Determining the weekday: Friday","INFORMATION"
	

		If loops < 1 Then ' We will be using todays XML
			If Not ValidateXMLDate(oXML) Then  ' Date() and Cube time are not the same. Critical  Error. Quit
				LogEvent oLogfile, "Invalid XML date. Expected:" & strDate & " Found:" & strCubeDate,"ERROR"
				LogEvent oLogfile,"Script exited with code: 4","INFORMATION"
				WScript.Quit(4) ' Bad Cube date
			End If
			LogEvent oLogfile, "Today ( " & strToday & " ) is not a TCD","INFORMATION"
			LogEvent oLogfile, "Valid XML date " & strCubeDate,"INFORMATION"
			Set oFD = oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt") ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt","INFORMATION"
			'ParseXML
			ParseXML strDate,oFD,oXML
		Else ' Use XML90
			
			Set oFD =  oFSO.CreateTextFile(strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt") ' Make file for tuesday
			LogEvent oLogfile, "Created a file " & strDrive & "\ExRate\" & FormatDate(strToday + 3,2) & ".txt","INFORMATION"
			ParseXML90 strDate,oFD,oXML90
			LogEvent oLogfile, "Today ( " & strToday & " ) is a TCD","INFORMATION"
			LogEvent oLogfile, "Parsing XML90","INFORMATION"
		End If
		
End Select 





' =============================================================
' =================== M A I N   E N D  ========================
' =============================================================


























































' =============================================================
' ======= F U N C T I O N S  &  S U B R O U T I N E S =========
' =============================================================


' ================== GetRate() ===============================
' Function returns string containing the current exchange rate
' for chosen currency. The string is formatted to comply with SAP 
' requirements
Function GetRate(rate,boolAdjust)
	Dim temp
	temp = Replace(rate,".",",") ' Replace delimiter character
	temp = CDbl(temp) ' Convert do double
	temp = 1 / temp   ' Do this so that we have enough digits
	If boolAdjust Then
		temp = temp * 100 ' Adjust for CZK and HUF 
		GetRate = Round(temp,5) & vbTab & "100" & vbTab & "1"  
		Exit Function
	End If
	GetRate = Round(temp,5) & vbTab & "1" & vbTab & "1"
End Function


' =============== ParseXML(targetDate,outputFile,oXML ============================
' Parameters: targetDate -> Date to search in XML
' Parameters: f -> FilesystemObject i.e open file descriptor
' Parameters: oXML -> Loaded XML object. From HTTP request
Function ParseXML(targetDate,f,oXML)
	Set n_ChildNodes = oXML.getElementsByTagName("Cube")
	
	For Each n_ChildNode In n_ChildNodes 
		If n_ChildNode.attributes.length > 1 Then
		
			
			Select Case n_ChildNode.attributes.getNamedItem("currency").text
					
						Case "NOK"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
							
						Case "USD"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
						
						Case "PLN"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
						
						Case "DKK"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
							
						Case "GBP"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
							
						Case "SEK"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
							
						Case "CHF"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,False) & vbCrLf
							
						Case "HUF"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,True) & vbCrLf
							
						Case "CZK"
							f.Write n_ChildNode.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(n_ChildNode.attributes.getNamedItem("rate").text,True) & vbCrLf
							
				End Select 
		End If
	Next
						
End Function


' =============== ParseXML90(targetDate,outputFile) ==============================
' Parameters: targetDate -> Date to search in XML
' Parameters: f -> FilesystemObject i.e open file descriptor
' Parameters: oXML -> Loaded XML object. From HTTP request
'
' Function parses XML with 90 days historical exchange 
' rates
' targetDate in format YYYYMMDD
' Subroutine loops through the XML until if finds cube="time" 
' with value of "targetDate"
' Then it retreives exchange rates as usual
Function ParseXML90(targetDate,f,oXML)
	Dim childNodes,child
	Set n_ChildNodes = oXML.getElementsByTagName("Cube")	
	For Each n_ChildNode In n_ChildNodes ' Loop through child nodes
		If n_ChildNode.attributes.length = 1 Then ' First child node having only one attribute tells us the date
			If n_ChildNode.attributes.getNamedItem("time").text = targetDate Then ' Let's be sure it is the "time" attribute
				Set childNodes = n_ChildNode.childNodes ' Collection of child nodes cca 32 children
				'32 children nodes are processed here
				For Each child In childNodes
					Select Case child.attributes.getNamedItem("currency").text
					
						Case "NOK"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "USD"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "PLN"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "DKK"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "GBP"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "SEK"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "CHF"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,False)
							
						Case "HUF"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,True)
							
						Case "CZK"
							f.WriteLine child.attributes.getNamedItem("currency").text & vbTab & "EUR" & vbTab & GetRate(child.attributes.getNamedItem("rate").text,True)
							
					End Select 
				Next
				Exit For 
			ElseIf Replace(n_ChildNode.attributes.getNamedItem("time").text,"-","") < Replace(targetDate,"-","") Then ' Exit if we encounter an element with date lower than the target date. Dont loop through all 2000+ entries
				ParseXML90 = 1 ' Error. No such date found
				Exit Function
			End If 
		End If 
	Next
End Function

' ============== MakeOutputDir() =============================
' Function creates output directory. Function returns 0 if
' dir is successfully created or if it already exists
Function MakeOutputDir(strFullDirPath)
	Dim oFSO,oWSH
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oWSH = CreateObject("WScript.Shell")
	If oFSO.FolderExists(strFullDirPath) Then
		MakeOutputDir = 0
		Exit Function
	End If 
	oFSO.CreateFolder strFullDirPath
	MakeOutputDir = 0 ' Return 0. Successfully created a new output directory
End Function


' ============= FindDate() ====================================
' Function searches for targetDate which is later 
' passed as an argument to ParseXML90 function
Function FindDate(d)
	Dim key
	loops = 0
	Do While Weekday(d) = 7 Or Weekday(d) = 1 Or IsTCD(FormatDate(d,1)) = True
		d = d - 1 ' Previous days
		loops = loops + 1
	Loop
	
	FindDate = Year(d) & "-" & right("00" & Month(d),2) & "-" & right("00" & Day(d),2) ' Returns the last valid date in the format YYYY-MM-DD
		
End Function


' ========== FormatDate() ====================================
' Parameters: D -> Date
' Parameters: f -> format type
' format type 1 = DDMMYYYY
' format type 2 = YYYYMMDD
' formate type 3 = YYYY-MM-DD
Function FormatDate(D,f)
	Select Case f
	
		Case 1
			FormatDate = Right("00" & Day(D),2) & Right("00" & Month(D),2) & Year(D)
			Exit Function
		Case 2
			FormatDate = Year(D) & Right("00" & Month(D),2) & Right("00" & Day(D),2)
			Exit Function
		Case 3
			FormatDate = Year(D) & "-" & Right("00" & Month(D),2) & "-" & Right("00" & Day(D),2)
			Exit Function
			
	End Select 
End Function

Function IsTCD(D)
	Dim key
	For Each key In oTCD.Keys
		If key = D Then
			IsTCD = True
			Exit Function
		End If
	Next
		
		IsTCD = False
End Function


Function ValidateXMLDate(oXML)
	Dim n_ChildNode,n_ChildNodes,attrs,attr
	Set n_ChildNodes = oXML.getElementsByTagName("Cube")
	For Each n_ChildNode In n_ChildNodes
		Set attrs = n_ChildNode.attributes
		For Each attr In attrs
			If attr.baseName = "time" Then
				strCubeDate = attr.text
				If attr.text = FormatDate(Date(),3) Then
					ValidateXMLDate = True	' Todays date matches Cube date
					Exit Function
				Else 
					ValidateXMLDate = False ' Todays date doesn't match Cube date
					Exit Function
				End If 
			End if
		Next
	Next
End Function

Function WarnUser(message)
	MsgBox "Nespravny cas",16,message
End Function 


Sub LogEvent(file,message,severity)
	file.writeline Date() & vbTab & Time() & vbTab & strUser & vbTab & strComputer & vbTab & WScript.ScriptName & vbTab & message & vbTab & severity
End Sub 







