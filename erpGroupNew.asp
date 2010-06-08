<!--#include file="../lib.inc"-->
<!--#include file="../defaults.inc"-->

<!-- #include file="erpClasses.asp" -->
<!-- #include file="erpProcedures.asp" -->
<%
'initialise DB
dim Connection
call initDB(Connection)

' read formular data
course = request.form("course")
groupString = request.form("group")
group = CLng(groupString)
'response.write "course: " & course & " group: " & group' & " id: " & id

'The table "Gruppen" contains only the information which groups exist. 
'The users are assigned to a certain group by setting the "Custom Values" 
'of their entry in the "Members" table

' Set rs = Conn.Execute("SELECT Max(id) FROM erp_groups")
' id = rs(0) + 1
'rs.close


sqlInsert = "INSERT INTO Gruppen (Kurs, Gruppe, Periode) VALUES ('"&course&"', '" &group&"', '0')"
Connection.execute(sqlInsert)
'Conn.execute("INSERT INTO erp_work_plan (matNo_fk, amount, group_fk, period) VALUES ('5', '10', '1', '1')")
'Conn.execute("INSERT INTO erp_groups (course, [group], period) VALUES ('" & course &"', '" & group &"', '1')")

response.write "Die Gruppe wurde erfolgreich erstellt<br>"

' DB disconnect
Connection.close
%>