Option Explicit
Const ns = "xmlns:gesmes='http://www.gesmes.org/xml/2002-08-01' xmlns='http://www.ecb.int/vocabulary/2002-08-01/eurofxref'" ' XML namespace
Dim oFSO,oWSH,oXML,oHTTP,oFD
Dim n_ChildNodes,n_ChildNode,a_Attributes,a_Attribute,strDate
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWSH = CreateObject("WScript.Shell")
Set oXML = CreateObject("MSXML2.DOMDocument")
Set oHTTP = CreateObject("MSXML2.XMLHTTP")
oHTTP.open "GET", "https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml",False
oHTTP.send ' Send HTTP request



If oHTTP.status = 200 Then
	oXML.load(oHTTP.responseXML)
	oXML.setProperty "SelectionNamespaces", ns
	Set n_ChildNodes = oXML.getElementsByTagName("Cube")
	
	For Each n_ChildNode In n_ChildNodes ' Loop through child nodes
		If n_ChildNode.attributes.length = 1 Then ' First child node having only one attribute tells us the date
			If n_ChildNode.attributes.item(0).baseName = "time" Then ' Let's be sure it is the "time" attribute
				strDate = Replace(n_ChildNode.attributes.getNamedItem("time").text,"-",".") ' Output file name i.e 2020.04.09.txt
				Set oFD = oFSO.CreateTextFile(oWSH.SpecialFolders("Desktop") & "\" & strDate & ".txt",True,False) ' Output file
			End If 
		End If
		If n_ChildNode.attributes.length > 1 Then ' Nodes with 2 attributes(currency,rate) hold the currency rates
			Set a_Attributes = n_ChildNode.attributes
			For Each a_Attribute In a_Attributes
				Select Case a_Attribute.text
				    ' Copy and paste the Case for each currency you want and change the Case expression
					Case "NOK"
						oFD.WriteLine a_Attribute.text & vbTab & Round(n_ChildNode.attributes.getNamedItem("rate").text,5) &_
						vbTab & "1"
					Case "USD"
						oFD.WriteLine a_Attribute.text & vbTab & Round(n_ChildNode.attributes.getNamedItem("rate").text,5) &_
						vbTab & "1"
					Case "PLN"
						oFD.WriteLine a_Attribute.text & vbTab & Round(n_ChildNode.attributes.getNamedItem("rate").text,5) &_
						vbTab & "1"
					Case "HUF"
						oFD.WriteLine a_Attribute.text & vbTab & Round((n_ChildNode.attributes.getNamedItem("rate").text / 100),5) &_
						vbTab & "100"
					Case "CZK"
						oFD.WriteLine a_Attribute.text & vbTab & Round((n_ChildNode.attributes.getNamedItem("rate").text / 100),5) &_
						vbTab & "100"
					Case "DKK"
						oFD.WriteLine a_Attribute.text & vbTab & Round(n_ChildNode.attributes.getNamedItem("rate").text,5) &_
						vbTab & "1"
						
				End Select
			Next
		End If
	Next
	WScript.Quit(0) ' SUCCESS
	
Else 
	WScript.Quit(1) ' ERROR obtaining xml file
End If




