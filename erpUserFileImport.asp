<!--***********************
erpUserFileImport.asp
*************************-->
<!--#include file="../lib.inc"-->
<!--#include file="../defaults.inc"-->

<!-- #include file="erpClasses.asp" -->
<!-- #include file="erpProcedures.asp" -->
<%
' connect to DB
'initialise DB (own erp.mdb)
dim Connection
call initDB(Connection)

'connect to ve-forum db
Set ConnVE = Server.CreateObject("ADODB.Connection")
ConnVE.Open DB

'Variablen & Konstanten erstellen
Dim fso, fsoFile, Pfad
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
'fpath = "http://www.technology-management.de/apps/erpTest.asp"
fpath = "D:\Inetpub\VEForum\Apps\erpUsers.txt"

'Objekte erstellen
set fso = Server.CreateObject("Scripting.FileSystemObject")

If fso.FileExists(fpath) Then
  set fsoFile = fso.OpenTextFile(fpath, ForReading , true, TristateFalse)
Else
  Response.Write("File 'erpUsers.txt' cannot be found!")
End if

'Text aus Datei lesen und ausgeben
course = ""
group = ""
userAlias = ""
Do while not fsoFile.AtEndOfStream
	'response.write "Starting to loop through lines in file.<br>"
	line = fsoFile.ReadLine
	'response.write "line: " & line & "<br>"
	if instr(line, "<course>") > 0 AND instr(line, "</course>") > 7 Then
		course = replace(line, "<course>","")
		course = replace(course, "</course>","")
		'response.write "len(course): " &len(course)& "; course: " & course & "<br>"
	elseif instr(line, "<group>") > 0 AND instr(line, "</group>") > 7 Then 'second "in string" must be > 8 cause closing tag must follow opening tag + at least 1 char content
		group = replace(line, "<group>", "")
		group = replace(group, "</group>", "")
		'response.write "len(group): " & len(group) & "; group: " & group & "<br>"
		'get groupID from table "Gruppen"
		'*****************************
		if group<>"" AND course<>"" then
			'response.write "course: " & course & "; group: " & group & "<br>" 
			sql = "SELECT * FROM Gruppen WHERE Kurs LIKE '"&course&"' AND Gruppe LIKE '"&group&"'"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.open sql, Connection
			'parse results
			If rs.EOF Then
				Response.Write "ERROR in erpUserFileImport.asp. Group " & course & "/" & group & " not present in table 'Gruppen'.<br>"
			Else
				groupID = rs(0)
			end if
			rs.close
			set rs = Nothing
		end if
	elseif instr(line, "<user>") > 0 AND instr(line, "</user>") > 6 Then 'second "in string" must be > 7 cause closing tag must follow opening tag + at least 1 char content	
		userAlias = replace(line, "<user>", "")
		userAlias = replace(userAlias, "</user>", "")
		'response.write "len(userAlias): " & len(userAlias) & "; userAlias: " & userAlias & "<br>"
		if groupID<>vbEmpty then
			'get userID from "User" table in ve-forum db
			'**************************************
			Set rs = Server.CreateObject("ADODB.Recordset")
			sql = "SELECT IDuser FROM [User] WHERE Alias LIKE '"&userAlias&"'"
			'start query
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open sql, ConnVE
			'parse results
			If rs.EOF Then
				Response.Write "ERROR in erpUserFileImport.asp. User not present in ve-forum's db; table Members.<br>"
			Else
				'******************************************************************************************************
				' Assign user to selected group by writing the groupID into his "Custom1" field in the "Members" table
				'update entry in "Members" table
				'******************************************************************************************************
				sqlUpdate = "UPDATE Members SET Custom1 = '"&groupID&"' WHERE IDuser LIKE '"&rs(0)&"' "
				ConnVE.Execute(sqlUpdate)
				'response.write "rs(0)==IDuser: " & rs(0) & "<br>"
				
				'positive feedback
				response.write "Der Nutzer '"&userAlias&"' wurde erfolgreich der Gruppe '"&course&"/"&group&"' zugeordnet.<br><br>"
				
				'try to move rs-pointer forward (actually that shouldn't work, there should be only 1 result)
				rs.MoveNext()
				Do While NOT rs.Eof
					response.write "ERROR in erpUserFileImport.asp. sql-query returned more than one result for a single userId in the VEForum table 'Members'.<br>"
					'fetch next entry
					rs.MoveNext()
				Loop
			end if
			'clean memory, destroy objects
			rs.close
			set rs = Nothing
		else
			response.write "ERROR in erpUserFileImport.asp. groupID=vbEmpty."
		end if
	elseif line="" then
		'empty line, do nothing
	else
		response.write("Invalid line in erpUsers.txt")
  end if
Loop

'TextDatei schliessen und Objekte terminieren
fsoFile.close
set fsoFile = nothing
set fso = nothing
Connection.close
set Connection=Nothing
ConnVE.close
set ConnVE=Nothing
%>