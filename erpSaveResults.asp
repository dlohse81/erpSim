<!--***********************
erpSaveResults.asp
*************************-->
<% @Language="VBScript" %>
<html id="html">
<head>
	<title>CeTIM ERP Simulation</title>
	
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->
		
	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->

	
	<!-- continue with <head> -->	
	<%
	'get usergroup
	dim usergroup
	set groupObj = New ERPGroup
	call getUsergroup(usergroup, groupObj)
	'response.write "usergroup: " & usergroup & "<br>"

	'get game period
	dim period
	call getPeriod(period)

	'initialise DB
	dim Connection
	call initDB(Connection)

	'get formular data
	htmlCodeUnicode = Request.Form("htmlCodeUnicode")
	' htmlCode = Request.Form("htmlCode")
	' response.write "succesfully loaded<br>"
	' response.write "htmlCodeUnicode: " & htmlCodeUnicode
	' response.write "<br><br>"
	' response.write "htmlCode: " & htmlCode
	
	'get date-time-group
	dtg = CDbl(Now())

	'check if their is already an entry for the usergroup in the current period, otherwise write entry
	sql = "SELECT * FROM Ergebnisse WHERE usergroup="&usergroup&" AND Periode="&period&" "
	set rs = Server.CreateObject("ADODB.Recordset")	
	rs.Open sql, Connection
	'If rs.EOF Then
		sqlInsert = "INSERT INTO Ergebnisse (usergroup, Periode, Zeitstempel, Lagerbestand, Eingabe, Ergebnis) "_
				& "VALUES ("&usergroup&", "&period&", "&dtg&", 'text', 'text', '"&htmlCodeUnicode&"')"
		Connection.execute(sqlInsert)
	'else
		'response.write "no insert"
	'end if

	'destroy object
	rs.close
	set rs=Nothing

	'response.write	 "<body onload='showPage("&htmlCodeUnicode&")' style='height:100%'>"
	%>

	<!-- JavaScript functions -->
	<script type="text/javascript">	
		function showPage() {
			//alert("IN");
			//alert("<%=htmlCodeUnicode%>");
			
			var htmlCodeUnicode = "<%=htmlCodeUnicode%>";
			var htmlCode
			htmlCode = decodeURIComponent(htmlCodeUnicode);
			document.write(htmlCode);
			/*
			for (var i = 0; i <= htmlCode.length-1; i++) {
				
				htmlCode = htmlCode + htmlCode.fromCharCode(i);
			}
			*/
			/*
			alert(htmlCodeUnicode);
			
			//alert(htmlCodeUnicode);
			//var htmlCode = decodeFromHex(htmlCodeUnicode);
			*/
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

<%
'end of block
%>

</body>
</html>