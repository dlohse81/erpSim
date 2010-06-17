<!--**************************
erpGroupsAssignMember.asp 
****************************-->
<!--#include file="../lib.inc"-->
<!--#include file="../defaults.inc"-->

<!-- #include file="erpClasses.asp" -->
<!-- #include file="erpProcedures.asp" -->
<%

' connect to DB (VE-Forum, not erp-mdb)
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open DB

' read formular data
userAlias = request.form("userAlias") 
groupID = request.form("groupID") 'erp-groupId in erp.mdb "Gruppen"
response.write "userAlias: " & userAlias & "<br>"
response.write "groupID: " & groupID & "<br>"

'*****************************
'get userID from "User" table
'*****************************
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT IDuser FROM [User] WHERE Alias LIKE '"&userAlias&"'"
'start query
rs.Open sql, Conn
'parse results
If rs.EOF Then
	Response.Write "ERROR in erpGroupsAssignMember.asp. User not present in table Members.<br>"
Else
	'******************************************************************************************************
	' Assign user to selected group by writing the groupID into his "Custom1" field in the "Members" table
	'update entry in "Members" table
	'******************************************************************************************************
	sqlUpdate = "UPDATE Members SET Custom1 = '"&groupID&"' WHERE IDuser LIKE '"&rs(0)&"' "
	'sqlUpdate = "UPDATE Members SET Custom1 = '"&groupID&"' WHERE IDuser LIKE '"&rs(0)&"' AND IDProject = 520 "
	Conn.Execute(sqlUpdate)

	response.write "rs(0)==IDuser: " & rs(0) & "<br>"
	
	'positive feedback
	response.write "Der Nutzer wurde erfolgreich der gewählten Gruppe zugeordnet.<br><br>"
	
	'try to move rs-pointer forward (actually that shouldn't work, there should be only 1 result)
	rs.MoveNext()
	Do While NOT rs.Eof
		response.write "ERROR in erpGroupsAssignMember.asp. sql-query returned more than one result for a single userId in the VEForum table 'Members'.<br>"
		'fetch next entry
		rs.MoveNext()
	Loop
end if


'clean memory, destroy objects
rs.close
set rs = Nothing


'Response.Clear
'Response.Redirect("http://ve-forum.org/apps/pubs.asp?Q=3&T=Simulationsprogramm")
'Response.End
'response.write "Der Nutzer wurde erfolgreich der gewählten Gruppe zugeordnet.<br><br>"
'Response.Write("<script>location.href = 'http://ve-forum.org/apps/pubs.asp?Q=3&T=Simulationsprogramm';</script>")
'Response.Write("<script>window.open('http://ve-forum.org/apps/pubs.asp?Q=3&T=Simulationsprogramm');</script>") 
'response.write("<script>document.write 'TEST TEST TEST'</script>")


'Link back to overview
'response.write "<a href='http://ve-forum.org/apps/pubs.asp?Q=3&T=Simulationsprogramm'>zurück zur Übersicht</a>"


%>