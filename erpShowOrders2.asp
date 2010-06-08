<% @ Language="VBScript" %>
<html>

<head>
    <title>CeTIM ERP Simulation</title>
</head>
 
<body>
<!-- #include file="erpClasses.asp" -->
<!-- #include file="erpProcedures.asp" -->
<!-- form -->

<!--
<form method="POST" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">
<h3>Bestellungen von Kaufteilen:</h3>
<div id="kteilePatternDiv">
-->
<%

'get usergroup
call getUsergroup(usergroup)

'get game period
call getPeriod(period)

'initialise DB
call initDB(Connection)

'initialise KTeile and ETeile
call initETeile(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190) 'last 5 entries intermediate products of e14,e15,e16,e18,e19
call initKTeile(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)


response.write "Die folgenden Bestellungen und Arbeitsaufträge wurden von Gruppe " &usergroup& " für die Spielperiode " &period& " erfasst: <br><br>"

'*********************************************************
'delete selected entries
'*********************************************************
deleteID = Request.Form("deleteID")
deleteTable = Request.Form("deleteTable")
response.write "deleteID: " & deleteID & "<br>"
response.write "deleteTable: " & deleteTable & "<br>"

'response.write "vartype(id): " & vartype(id) & "<br>"

if deleteID <> vbEmpty then
	response.write "in...<br>"
	'response.write "id: " & id & "<br>"
	'response.write "table: " & table & "<br>"
	sqlDelete = "DELETE FROM " &deleteTable& " WHERE id="&deleteID&""
	Connection.execute(sqlDelete)
end if



'**********************************************************
'write Purchase Pieces (KTeile)
'**********************************************************
response.write "<h3>Kaufteilbestellungen</h3>"

sql = "SELECT * FROM Kaufteilbestellungen WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND geliefert = 0 ORDER BY id"
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

If Recordset.EOF Then
	Response.Write("Es sind noch keine Kaufteilbestellungen von Gruppe " &usergroup& " für Spielperiode " &period& "vorhanden.<br>")
Else	
	'form
	
	
	'table header
	response.write "<table>"
	response.write "<tr>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp K-TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Anzahl &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "</tr>"
	
	Do While NOT Recordset.Eof   
		kteilnr = Recordset("KTeilNr")
		kteilAmount = Recordset("Anzahl")
		deleteID = Recordset("id")
		deleteTable = "Kaufteilbestellungen"
		'response.write "id: " & id & "<br>"
		call setKTeilObject(kteilnr,kteilObject)
		
		response.write "<tr>"
		response.write "<td align='center'>"&kteilnr&"</td>"
		response.write "<td align='center'>"&kteilObject.name&"</td>"
		response.write "<td align='center'>"&kteilAmount&"</td>"
		response.write "<form method=""POST"" action='erpShowOrders2.asp' >"
		response.write "<input type=""hidden"" name=""deleteID"" value="&deleteID&">"
		response.write "<input type=""hidden"" name=""deleteTable"" value="&deleteTable&">"
		response.write "<td align='center'><input type=""submit"" name=""delete"" value='löschen'></td>"
		response.write "</form>"	
		response.write "</tr>"
		
		Recordset.MoveNext
	Loop
	
	response.write "</table>"

end if
response.write "<br><br>"

Recordset.close
set Recordset = Nothing




'********************************************
'write production pieces (ETeile)
'********************************************
response.write "<h3>Produktionsaufträge für Eigenfertigunsteile</h3>"

sql = "SELECT * FROM Produktionsauftraege WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND abgeschlossen = 0 "
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

If Recordset.EOF Then
	Response.Write("Es sind noch keine Produktionsaufträge von Gruppe " &usergroup& " für Spielperiode " &period& " vorhanden.<br>")
Else	
	'header
	response.write "<table>"
	response.write "<tr>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp AuftragsNr &nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp E-TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Losgröße &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp Tag &nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "</tr>"
	
	Do While NOT Recordset.Eof   
		eteilnr = CInt(Recordset("ETeilnr"))
		if eteilnr = 140 OR eteilnr = 150 OR eteilnr = 160 OR eteilnr = 180 OR eteilnr = 190 then
			'continue (with next iteration)
			Recordset.MoveNext
			set eteilnr = Nothing
		else
			deleteID = Recordset("id")
			deleteTable = "Produktionsauftraege"
			Set pOrder = New ProdOrder
			with pOrder
				.id = CInt(Recordset("id"))
				.prodOrderNo = CInt(Recordset("Auftragsnr"))
				.eteilnr = CInt(Recordset("ETeilnr"))
				.day  = CInt(Recordset("Tag"))
				.batchsizeRequired = CInt(Recordset("Losgroesse"))
				.finished = CInt(Recordset("abgeschlossen")) 'Anzahl der bereits gefertigten Teile vom Vortag, falls Los nicht komplett abgearbeitet werden konnte
				'.batchsize
			end with
				
			call setETeilObject(pOrder.eteilnr,eteilObject)
			
			response.write "<tr>"
			response.write "<td align='center'>"&pOrder.prodOrderNo&"</td>"
			response.write "<td align='center'>"&pOrder.eteilnr&"</td>"
			response.write "<td align='center'>"&eteilObject.name&"</td>"
			response.write "<td align='center'>"&pOrder.batchsizeRequired&"</td>"
			response.write "<td align='center'>"&pOrder.day&"</td>"
			response.write "<form method=""POST"" action='erpShowOrders2.asp' >"
			response.write "<input type=""hidden"" name=""deleteID"" value="&deleteID&">"
			response.write "<input type=""hidden"" name=""deleteTable"" value="&deleteTable&">"
			response.write "<td align='center'><input type=""submit"" name=""delete"" value='löschen'></td>"
			response.write "</form>"
			response.write "</tr>"
			
			Recordset.MoveNext
		end if
	Loop
	
	response.write "</table>"
end if


Recordset.close
set Recordset = Nothing

'close DB connection
Connection.close
set Connection = Nothing
%>

<!--
<input type="submit" value=" Absenden ">
<input type="reset" value=" Abbrechen">
</form>
-->

</body>
</html>