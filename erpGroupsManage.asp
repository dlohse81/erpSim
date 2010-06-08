<!--***********************
erpGroupsManage.asp
*************************-->
<html>
<head>
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
</head>
<body>

<br>
<!--
The "erp_groups" table contains only the information which groups exist. 
The users are assigned to a certain group by setting the "Custom Values" of their entry in the "Members" table

*********************************************
Formular to assign a user to a certain group
*********************************************
-->
<h3>VE-Forum-Nutzer einer Gruppe zuordnen</h3><br><br>
<form action='erpGroupsAssignMember.asp' method='Post'>
	<table>
		<tr>
			<td>Nutzerkennung: </td>
			<td><input type='text' name='userAlias' id='userAlias'></td>
		</tr>

		<!--create drop down menu to choose group-->
		<tr>
			<td>Gruppe zuweisen: </td>

			<td>
			<select name='groupID'> <!-- erpGroupId in erp.mdb "Gruppen"-->
				<%
				'fill out drop-down menu	
				'initialise DB
				dim Connection
				call initDB(Connection)
				
				'query db
				sql = "SELECT * FROM Gruppen"
				Set Recordset=Server.CreateObject("ADODB.Recordset")	
				Recordset.Open sql, Connection
				
				'parse results
				If Recordset.EOF Then
					Response.Write "ERROR in erpGroupsManage. Keine Eintrag in der Tabelle 'Gruppen' vorhanden.<br>"
				Else
					Do While NOT Recordset.Eof
						response.write "<option value='"&Recordset(0)&"'>"&Recordset(1)&"/"&Recordset(2)&"</option>"
						
						Recordset.MoveNext()
					Loop
				end if
				
				'close objects
				Recordset.close
				set Recordset = Nothing
								
				Connection.close
				set Connection = Nothing
				%>
			</select>
			</td>
		</tr>
	</table>

	<!-- submit button -->
	<input style='position: relative; left: 60%;' type='submit' value='submit'>
</form>

<br><br>



<!--
**********************************
 create new courses and groups
**********************************
create formular to add new groups
-->

<h3>Neue Gruppe erstellen</h3><br><br>
<form action='erpGroupNew.asp' method='Post'>
	<table>
		<tr>
			<td>Kurs (z.B. 'LRT09'): </td>
			<td><input size='5' type='text' name='course' id='course'></td>
		</tr>

		<tr>
			<td>Gruppe (z.B. '2'):</td>
			<td><input size='2' type='text' name='group' id='group'></td>
		</tr>
	</table>

	<!--submit button-->
	<input style='position: relative; left: 60%;' type='submit' value='submit'>
</form>


</body>
<html>