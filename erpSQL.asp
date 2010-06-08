<html>
<head>
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
</head>

<body>
<%
'**************************************
'initialisations
'**************************************
'get usergroup
set groupObj = New ERPGroup
call getUsergroup(usergroup, groupObj)

'get game period
call getPeriod(period)

'initialise DB
call initDB(Connection)

'check for sql-query from recursive call
sql = Request.Form("erpSQL")
if sql <> vbEmpty then
	Connection.execute(sql)
end if
%>
<h3>ERP-Sim SQL-Query Prompt zum schnellen SQL-Zugriff auf erpSim.mdb</h3>
<form method="POST" action="erpSQL.asp">
	<textarea name="erpSQL" cols="80" rows="10"></textarea>
	<br>
	<input type="submit" value="Absenden">
	<input type="reset" value = "Abbrechen">
</form>

</body>
</html>