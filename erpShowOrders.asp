<!--***********************
erpShowResults.asp
*************************-->
<% @Language="VBScript" %>
<html>

<head>
    <title>CeTIM ERP Simulation</title>

	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->


	<%
	'get usergroup
	set groupObj = New ERPGroup
	call getUsergroup(usergroup, groupObj)

	'get game period
	call getPeriod(period)

	'initialise DB
	call initDB(Connection)

	'get formular data
	resultId = Request.Form("resultId")
	'response.write "resultId: " & resultId & "<br>"

	sql = "SELECT * FROM Ergebnisse WHERE id="&resultId&" "
	Set Recordset=Server.CreateObject("ADODB.Recordset")	
	Recordset.Open sql, Connection

	'***********************************
	'iterate through purchase orders
	'***********************************
	If Recordset.EOF Then
		Response.Write("Error in erpShowResults.asp. Das ausgewählte Ergebnis konnte in der Db nicht gefunden werden.<br>")
	Else
		htmlCodeUnicode = Recordset("Ergebnis")
		'response.write "htmlCodeUnicode: " & htmlCodeUnicode & "<br>"
		if htmlCodeUnicode <>"" then
			response.write "htmlCodeUnicode received!<br>"
		end if
		'Do While NOT Recordset.Eof   
	end if
		
	'clear memory, destroy objects
	Recordset.close
	set Recordset = Nothing
	Connection.close
	set Connection = Nothing


	%>
	<!-- JavaScript functions -->
	<script type="text/javascript">	
		function showPage() {
			//alert("IN");
			//alert("<%=htmlCodeUnicode%>");
			var htmlCodeUnicode = "<%=htmlCodeUnicode%>";
			var htmlCode = decodeURIComponent(htmlCodeUnicode);
			document.write(htmlCode);
		}
		/*
		function encodeToHex(str){
			var r="";
			var e=str.length;
			var c=0;
			var h;
			while(c<e){
				h=str.charCodeAt(c++).toString(16);
				while(h.length<3) h="0"+h;
				r+=h;
			}
			return r;
		}

		function decodeFromHex(str){
			var r="";
			var e=str.length;
			var s;
			while(e>0){
				s=e-3;
				r=String.fromCharCode("0x"+str.substring(s,e))+r;
				e=s;
			}
			return r;
		}
		*/
	</script>
</head>

 
<body onload='showPage()' style='height:100%'>
</body>
</html>