<!--***********************
erpWorkPlanInsert.asp
*************************-->
<!--#include file="../lib.inc"-->
<!--#include file="../defaults.inc"-->

<%

'connect to DB
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open DB


for each i in request.form
	'read form data
	response.write Request.Form(i) & "<br>"
	inputDataArray = split(Request.Form(i), ", ")
	
	'get material
	mat = inputDataArray(0)
	
	'get amount
	if inputDataArray(1) <> "" then
		amount = CLng(trim(inputDataArray(1)))
	else
		amount = 0
	end if
	
	'get manufacturingNo
	manufacturingNo = inputDataArray(2)
	
	'get info if entry shall be deleted
	if UBound(inputDataArray) = 3 then
		if inputDataArray(3) = "delete" then
			delete = "yes"
		else 
			response.write "invalid checkbox value for deletion."
		end if
	else
		delete = "no"	
	end if
	
	
	' response.write "priority: " & priority & "<br>"
	' response.write "mat: " & mat & "<br>"
	' response.write "amount: " & amount & "<br><br>"
	'response.write "delete: " & delete & "<br>"
	
	if amount <> 0 then
		'read foreign key from material-table
		'set rs = Conn.execute("SELECT matNo FROM erp_material_data WHERE description LIKE '" & mat & "'")
		
		'Get current user
		user = Session("user")
		userID = Session("teamid")
		
		'Get user's group
		set rs = Conn.Execute("SELECT Custom1 FROM Members WHERE ID LIKE '" & userID & "'")
		groupFk = rs(0)
		rs.close
		'rs = Nothing
		
		'Get user's current period
		set rs2 = Conn.Execute("SELECT period FROM erp_groups WHERE id LIKE '"& groupFk &"'")
		period = rs2(0)
		rs2.close
		
		
		'get current manufacturing orders for user's group
		'set rs = Conn.execute("SELECT * FROM erp_work_plan WHERE group_fk LIKE '"& groupFK &"' ")
		set rs = Conn.execute("SELECT * FROM erp_work_plan WHERE manufacturingNo LIKE '"& manufacturingNo &"' ")
		
		'if there is an entry, it has to be updated, otherwise a new entry has to be added
		if not rs.eof then
			'************
			'update entry
			'************					
			'get necessary data from tables
			set rs2 = Conn.execute("SELECT matNo FROM erp_material_data WHERE description LIKE '"& mat &"'")
			matNoFK = rs2(0)
			rs2.close
			
			sqlQuery = "UPDATE erp_work_plan " &_
						 "SET matNo_fk = '"& matNoFk &"', amount = '"& amount &"', period = '"& period &"' " &_
						 "WHERE manufacturingNo LIKE '"& manufacturingNo &"'"
			Conn.Execute(sqlQuery)
		else
			'**************
			'add new entry
			'**************
			'get necessary data from tables
			set rs2 = Conn.execute("SELECT matNo FROM erp_material_data WHERE description LIKE '"& mat &"'")
			matNoFK = rs2(0)
			rs2.close
			
			sqlQuery = "INSERT INTO erp_work_plan (matNo_fk, amount, group_fk, period) " & _ 
						"VALUES ('" & matNoFk & "', '" & amount & "', '" & groupFk & "', '" & period & "')"
			Conn.Execute(sqlQuery)
		end if
	end if
	
	'****************************
	'delete recordset
	'****************************
	if delete = "yes" then
		Conn.execute("DELETE FROM erp_work_plan WHERE manufacturingNo LIKE '" & manufacturingNo & "'")
		response.write "Die ausgewählten Datensätze wurden erfolgreich gelöscht!"
	end if
next



'positive feedback
response.write "Der Arbeitsaufträge wurden erfolgreich hinzugefügt.<br><br>"

'Link back to overview
response.write "<a href='http://ve-forum.org/*520-P2947/e'>zurück zur Arbeitsplanerstellung</a>"

'DB disconnect
Conn.close
%>