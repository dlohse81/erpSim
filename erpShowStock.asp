<% @Language="VBScript" %>
<html>

<head>
    <title>CeTIM ERP Simulation</title>

	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
</head>
 
<body>

<!-- form -->

<!--
<form method="POST" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">
<h3>Bestellungen von Kaufteilen:</h3>
<div id="kteilePatternDiv">
-->
<%

'get usergroup
set groupObj = New ERPGroup
call getUsergroup(usergroup, groupObj)

'get game period
call getPeriod(period)

'initialise DB
call initDB(Connection)

'initialise KTeile and ETeile
call initETeile(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190) 'last 5 entries intermediate products of e14,e15,e16,e18,e19
call initKTeile(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)


response.write "Folgender Lagerbestand liegt für Gruppe " &groupObj.name& " in der Spielperiode " &period& " vor: <br><br>"


'**********************************************************
'Purchase Pieces (KTeile)
'**********************************************************
response.write "<h3>Lagerbestand an Kaufteilen</h3>"

sql = "SELECT * FROM Lager WHERE usergroup = "&usergroup&" AND Teilnr > 19 AND Teilnr < 45 " 'AND Periode = "&period&" "
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

If Recordset.EOF Then
	Response.Write("Es ist noch kein Lagerbestand für " &usergroup& " in Spielperiode " &period& "vorhanden. Bitte Lagerbestand anlegen.<br>")
Else	
	'header
	response.write "<table>"
	response.write "<tr>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Anzahl &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "</tr>"
	
	Do While NOT Recordset.Eof   
		kteilnr = Recordset("Teilnr")
		kteilAmount = Recordset("Lagerbestand")
		'call setKTeilObject(kteilnr, kteilObject)
		call setItemObject(kteilnr, kteilObject)
		
		
		response.write "<tr>"
		response.write "<td align='center'>"&kteilObject.nr&"</td>"
		response.write "<td align='center'>"&kteilObject.name&"</td>"
		response.write "<td align='center'>"&kteilAmount&"</td>"
		response.write "</tr>"
		
		Recordset.MoveNext
	Loop
	
	response.write "</table>"
end if
response.write "<br><br>"

Recordset.close
set Recordset = Nothing




'********************************************
'production pieces (ETeile)
'********************************************
response.write "<h3>Lagerbestand für Eigenfertigunsteile</h3>"

sql = "SELECT * FROM Lager WHERE usergroup = "&usergroup&" AND Teilnr < 20 " 'AND Periode = "&period&" "
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

If Recordset.EOF Then
	Response.Write("Esn ist noch kein Lagerbestand für Gruppe " &usergroup& " für Spielperiode " &period& " angelegt. Bitte Lagerbestand anlegen.<br>")
Else	
	'header
	response.write "<table>"
	response.write "<tr>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp E-TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
	response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Bestand &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
	response.write "</tr>"
	
	Do While NOT Recordset.Eof   
		eteilnr = CInt(Recordset("Teilnr"))
		eteilHolding = Recordset("Lagerbestand")
		'call setETeilObject(eteilnr, eteilObject)
		call setItemObject(eteilnr, eteilObject)
		
		response.write "<tr>"
		response.write "<td align='center'>"&eteilObject.nr&"</td>"
		response.write "<td align='center'>"&eteilObject.name&"</td>"
		response.write "<td align='center'>"&eteilHolding&"</td>"
		response.write "</tr>"
		
		Recordset.MoveNext
	
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