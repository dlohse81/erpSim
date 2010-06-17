<!--**************************
erpSetPeriod.asp 
****************************-->
<% @ Language="VBScript" %>
<%
' read formular data
period = request.form("period") 
response.write "period: " & period & "<br>"

'Variablen & Konstanten erstellen
Dim fso, fsoFile, path
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
path = "D:\Inetpub\VEForum\Apps\erpPeriod.txt"

'Objekte erstellen
set fso = Server.CreateObject("Scripting.FileSystemObject")

'open file 'period.txt' and write period into it
If fso.FileExists(path) Then
  set fsoFile = fso.OpenTextFile(path, ForWriting , false, TristateFalse)
	fsoFile.WriteLine period
	response.write "Periode " &period& " wurde erfolgreich aktiviert.<br>"
Else
  Response.Write("ERROR in erpSetPeriod.asp. file 'erpPeriod.txt' not found in 'D:\Inetpub\VEForum\Apps\erpPeriod.txt' <br>")
End if



'TextDatei schliessen und Objekte terminieren
fsoFile.close
set fsoFile = nothing
set fso = nothing
%>