<html>
<head>
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->
    
	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
</head>
<body>
<%
'get usergroup
set groupObj = New ERPGroup
call getUsergroup(usergroup, groupObj)

'initialise DB
dim Connection
call initDB(Connection)
%>
<h3>Simulationsergebnisse einsehen</h3>
<div><p>Vorhandene Simulationsergebnisse für Gruppe <%=groupObj.name%>:</p></div>
<!--<form method="POST" onsubmit="return checkInput();" action="erpShowResults.asp">-->
<form method="POST" action="erpShowResults.asp">
	<select name="resultId">
		<%
		sql = "SELECT * FROM Ergebnisse WHERE usergroup="&usergroup&" "
		Set Recordset=Server.CreateObject("ADODB.Recordset")	
		Recordset.Open sql, Connection
		
		If Recordset.EOF Then
			Response.Write "ERROR. Noch keine Simulationsergebnisse für Gruppe "&groupObj.name&" vorhanden.<br>"
		Else
			Do While NOT Recordset.Eof
				response.write "<option value='"&Recordset(0)&"'>ID: '"&Recordset(0)&"', Periode: '"&Recordset(2)&"', Zeitstempel: '"&Recordset(3)&"'</option>"
				
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

<%
' sql = "SELECT * FROM Ergebnisse WHERE usergroup="&usergroup&" "
' Set Recordset=Server.CreateObject("ADODB.Recordset")	
' Recordset.Open sql, Connection

' If Recordset.EOF Then
	' Response.Write "ERROR. Noch keine Simulationsergebnisse für Gruppe "&groupObj.name&" vorhanden.<br>"
' Else
	' Do While NOT Recordset.Eof
		' response.write "id: "&Recordset(0)&", Periode: "&Recordset(2)&", Zeitstempel: "&Recordset(3)&" "
		
		' Recordset.MoveNext()
	' Loop
' end if

'close Recordset
' Recordset.close
' set Recordset = Nothing
%>

<%
Connection.close
set Connection=Nothing
%>
</body>
</html>