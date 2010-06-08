<!--***********************
erpInitStock.asp 
*************************-->
<% @Language="VBScript" %>
<html>
<head>
	<title>Initialise Stock</title>

	<!-- #include file="erpProcedures.asp" -->
	<!-- #include file="erpClasses.asp" -->
	
	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
</head>
<body>


<%
'get usergroup (id) to initialise from formular
usergroup = Request.Form("usergroup")
'response.write "usergroup to init: " & usergroup & "<br>"
'set groupObj = New ERPGroup
'call getUsergroup(usergroup, groupObj)
%>


<%
'get game period
'set period = 1 as the period-field in the table "Lager" is currently not used 
period = 0
'call getPeriod(period)

'initialise DB
call initDB(Connection)


dim lagerbestandArray
lagerbestandArray=Array(0,0,19,38,19,19,19,38,38,19,19,130,65,390,520,130,390,260,260,171,38,19,171,76,437,323,19,38,228,1040,1040,2210,260,260,0,260,520,2990,190,470,195,105,520,65,0,0,0,0,0) 'last 5 entries represent the intermediate products no. 140, 150, 160, 180, 190

'look up usergroup in table "Lager"
sql = "SELECT * FROM Lager WHERE usergroup = "&usergroup&" " 'AND Periode = "&period&""
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

'if no usegroup entry is found in table "Lager" then create completely new entry for usergroup
if Recordset.EOF then
	for i=1 to 44
		id = id + 1
		'Response.write "i: " & i & "<br>"
		j=i-1
		sqlInsert = "INSERT INTO Lager (usergroup, Periode, Teilnr, Lagerbestand) VALUES ("&usergroup&", "&period&", "&i&", "&lagerbestandArray(j)&")"
		
		'executing the SQL statement
		Connection.execute(sqlInsert)
	next
	i=140
	sqlInsert = "INSERT INTO Lager (usergroup, Periode, Teilnr, Lagerbestand) VALUES ("&usergroup&", "&period&", "&i&", "&lagerbestandArray(j)&")"
	Connection.execute(sqlInsert)
	
	i=150
	sqlInsert = "INSERT INTO Lager (usergroup, Periode, Teilnr, Lagerbestand) VALUES ("&usergroup&", "&period&", "&i&", "&lagerbestandArray(j)&")"
	Connection.execute(sqlInsert)
	
	i=160
	sqlInsert = "INSERT INTO Lager (usergroup, Periode, Teilnr, Lagerbestand) VALUES ("&usergroup&", "&period&", "&i&", "&lagerbestandArray(j)&")"
	Connection.execute(sqlInsert)
	
	i=180
	sqlInsert = "INSERT INTO Lager (usergroup, Periode, Teilnr, Lagerbestand) VALUES ("&usergroup&", "&period&", "&i&", "&lagerbestandArray(j)&")"
	Connection.execute(sqlInsert)
	
	i=190
	sqlInsert = "INSERT INTO Lager (usergroup, Periode, Teilnr, Lagerbestand) VALUES ("&usergroup&", "&period&", "&i&", "&lagerbestandArray(j)&")"
	Connection.execute(sqlInsert)
	
'if usergroup entry is found in table "Lager" then reset (update) the stock to its initial values	
else
	for i=1 to 44+5 'because of intermediate products no. 140, 150, 160, 180, 190
		id = id + 1
		'Response.write "i: " & i & "<br>"
		j=i-1
		sqlUpdate = "UPDATE Lager SET Lagerbestand = "&lagerbestandArray(j)&" WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Teilnr = "&i&""
		
		'executing the SQL statement
		Connection.execute(sqlUpdate)
	next
	
	lagerbestandArray(j) = 0
	i = 140
	sqlUpdate = "UPDATE Lager SET Lagerbestand = "&lagerbestandArray(j)&" WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Teilnr = "&i&""
	Connection.execute(sqlUpdate)
	
	i = 150
	sqlUpdate = "UPDATE Lager SET Lagerbestand = "&lagerbestandArray(j)&" WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Teilnr = "&i&""
	Connection.execute(sqlUpdate)
	
	i = 160
	sqlUpdate = "UPDATE Lager SET Lagerbestand = "&lagerbestandArray(j)&" WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Teilnr = "&i&""
	Connection.execute(sqlUpdate)
	
	i = 180
	sqlUpdate = "UPDATE Lager SET Lagerbestand = "&lagerbestandArray(j)&" WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Teilnr = "&i&""
	Connection.execute(sqlUpdate)
	
	i = 190
	sqlUpdate = "UPDATE Lager SET Lagerbestand = "&lagerbestandArray(j)&" WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Teilnr = "&i&""
	Connection.execute(sqlUpdate)
	
	
end if


'user feedback
sql = "SELECT * FROM Gruppen WHERE id="&usergroup&" "
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

response.write "Lagerbestand für Gruppe " &Recordset(1)& "/" &Recordset(2)& " erfolgreich initialisiert.<br>"
Set stockObj = New Stock
stockObj.getHolding()
stockObj.writeHolding()

Recordset.close
Connection.close



%>

</body>
</html>