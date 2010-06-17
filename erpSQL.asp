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
sqlType = Request.Form("query")
if sql <> vbEmpty then
	if sqlType = "query" then
		Set rs = Server.CreateObject("ADODB.Recordset")
	
		rs.Open sql, Connection
		
		if rs.EOF Then
			Response.Write "Kein Abfrageergebnis!"
		else
			for each i in rs.fields
  				response.write(i.name)
				response.write(" = ")
				response.write(i.value)
				response.write "<br>"
			next
		end if

		rs.close
		set rs=Nothing
	else
		Connection.execute(sql)
	end if
end if

Connection.close
set connection=nothing

%>
<h3>ERP-Sim SQL-Query Prompt zum schnellen SQL-Zugriff auf erpSim.mdb</h3>
<form method="POST" action="erpSQL.asp">
	<textarea name="erpSQL" cols="80" rows="10"></textarea>
	<br>
	<input id="query" name="query" value="query" type="checkbox"/>
	<br><br>
	<input type="submit" value="Absenden">
	<input type="reset" value = "Abbrechen">
</form>

</body>
</html>