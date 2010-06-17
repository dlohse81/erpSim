<!--***************************
erpSetPeriodFormular.asp
****************************-->
<html>
<head>
<title>Spielperiode festlegen</title>
</head>

<body>
<h1>Spielperiode festlegen</h1>
<% @ Language="VBScript" %>
<% 
'********************************
'read current period from period.txt
'*******************************
Dim fso, fsoFile, path
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
path = "D:\Inetpub\VEForum\Apps\erpPeriod.txt"

'create filesystem object
set fso = Server.CreateObject("Scripting.FileSystemObject")

If fso.FileExists(path) Then
  set fsoFile = fso.OpenTextFile(path, ForReading , false, TristateFalse)
Else
  Response.Write("ERROR in erpSetPeriodFormular.asp. file 'erpPeriod.txt' not found in 'D:\Inetpub\VEForum\Apps\erpPeriod.txt' <br>")
End if

'read period from period.txt
'Do while not fsoFile.AtEndOfStream
  period = fsoFile.ReadLine
'Loop

'close file, kill objects
fsoFile.close
set fsoFile = nothing
set fso = nothing
%>
Derzeit ist Spielperiode <%response.write period%> aktiv.

<p>Bitte wählen Sie die Spielperiode aus, die <i>global</i> für alle Gruppen aktiviert werden soll.</p>

<form method="POST" action="erpSetPeriod.asp">
  <p>
		Spielperiode: 
    <select name="period" size="1">
      <%
			for i=1 to 15 
				response.write "<option>" &i& "</option>"
			next
			%>
    </select>
		&nbsp; 
		<input type="submit" value="Absenden">
    <input type="reset" value="Abbrechen">
  </p>
</form>

</body>
</html>
