<!--***********************
erpInitStockFormular.asp
*************************-->
<html>
<head>
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->
    
	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
</head>
<body>
<h3>Lagerbestand für Gruppen initialisieren</h3>
<div><p>Für welche Gruppe soll der aktuelle Lagerbestand neu initialisiert werden?<br>Achtung, der momentane Lagerbestand wird ersetzt durch den Anfangslagerbestand.</p></div>
<form method="POST" onsubmit="return checkInput();" action="erpInitStock.asp">
	<select name="usergroup">
<%
	'initialise DB
	dim Connection
	call initDB(Connection)

	sql = "SELECT * FROM Gruppen"
	Set Recordset=Server.CreateObject("ADODB.Recordset")	
	Recordset.Open sql, Connection
	
	If Recordset.EOF Then
		Response.Write "ERROR. Keine Eintrag in der Tabelle 'Gruppen' vorhanden.<br>"
	Else
		Do While NOT Recordset.Eof
			response.write "<option value='"&Recordset(0)&"'>"&Recordset(1)&"/"&Recordset(2)&"</option>"
			
			Recordset.MoveNext()
		Loop
	end if
	'close Recordset
	Recordset.close
	set Recordset = Nothing
%>
	
	</select>
	
	<input type="submit" value="Absenden">
	<input type="reset" value = "Abbrechen">
</form>

</body>
</html>