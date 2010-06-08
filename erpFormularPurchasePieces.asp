<!--**************************
erpFormularPurchasePieces.asp
****************************-->
<% @Language="VBScript" %>
<html>
<head>
	<title>ERP Simulation: Nutzereingaben</title>
	
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpTest2.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
	
	<!-- JavaScript functions -->
	<script type="text/javascript">
	
	function addKTeilRow() {
		//get input row (TR = Table Row = <tr>) and clone it
		var TR_K, TR_K_Clone;
		TR_K = document.getElementById("TR_K");
		TR_K_Clone = TR_K.cloneNode(true);
		
		//set inputs of clone back to zero
		var ausgabe = ""
		for (var i = 0; i <= TR_K_Clone.childNodes.length-1; i++) {
			ausgabe = ausgabe + "name: " + TR_K_Clone.childNodes[i].nodeName + "; value: " + TR_K_Clone.childNodes[i].nodeValue + "\n";
			// && TR_K_Clone.childNodes[i].nodeName == "td"
			
			if (i>0 && TR_K_Clone.childNodes[i].nodeName == "TD") {
				TR_K_Clone.childNodes[i].firstChild.value = 0;
				//alert("in");
			}			
		}	
		//alert(ausgabe);
		//alert("type: " + TR_K_Clone.childNodes[1].nodeType + "; name: " + TR_K_Clone.childNodes[1].nodeName + "; value: " + TR_K_Clone.childNodes[1].nodeValue);
		
		
		//append cloned row to document tree
		table = document.getElementById("tableK").tBodies[0];
		table.appendChild(TR_K_Clone);
	}
	
	/*
	function addETeilRow(tableId) {
		//find corresponding table
		table = document.getElementById(tableId).tBodies[0];
		//copy the last row
		var TRclone = table.rows[table.rows.length-1].cloneNode(true)
		//increment production number
		//TRclone.cells[0].childNodes[0].nodeValue++ 
		TRclone.cells[0].childNodes[0].value = 0
		TRclone.cells[0].childNodes[0].nodeValue = "noch keine"
		//set amount back to zero
		TRclone.cells[2].childNodes[0].value = 0 
		table.appendChild(TRclone)
	}
	*/
	function checkInput() {
		
		//return false;
		
	}
	
	function arrangeDOM() {
		
		var divContainer = document.getElementById("divContainer");
		//var divETeile = document.getElementById("divETeile");
		var divKTeile = document.getElementById("divKTeile");
		//var divDelE = document.getElementById("divDelE");
		var divDelK = document.getElementById("divDelK");
		var submitForm = document.getElementById("submitForm");
		
		//move delete-buttons to table "Kaufteilbestellungen"
		var rows = divDelK.firstChild.firstChild.childNodes;
		var tableK = document.getElementById("tableK").firstChild;
		for (var i=1; i<rows.length; i++) { //start with i=1 as rows[0] == header-row
			tableK.childNodes[i].appendChild(rows[i].firstChild);
		}
		
		
		
		//***********************************************************
		
		//move delete-buttons to table "Produktionsauftraege"		
		/*
		for (var day=1; day<=5; day++) {
			var tableId="tableE" + day
			//alert("tableID:"+tableId)
			var tableDel = divDelE.firstChild.firstChild
			var delButtonRows = tableDel.childNodes;
			var tableE = document.getElementById(tableId).firstChild; //tableBody
			
			var rows = tableE.childNodes;
			//alert(tableId + "; " + rows.length)
			
			//append delete-buttons from table "tableDel" to table "tableE"
			for (var i=1; i<rows.length-1; i++) { //start with i=1 as rows[0] == "header-row"
				delButtonRows[i].firstChild.vAlign = "bottom"
				//alert(rows[i].firstChild.vAlign)
				tableE.childNodes[i].appendChild(delButtonRows[i].firstChild);
			}
			
			//remove delete-buttons from table "tableDel"
			for (var i=1; i<rows.length-1; i++) {
				tableDel.removeChild(tableDel.childNodes[1])
			}
			
			
		}
		*/
	}
	
	</script>
</head>

<!-- END OF HEAD -->

<body onload="arrangeDOM()" style="height:100%">


<!-- ----------------------------------- DELETE SECTION ---------------------- -->
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

'initialise workstations, KTeile and ETeile
call initWS(ws1,ws2,ws3,ws4,ws5,ws6,ws7)
call initETeile(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190) 'last 5 entries intermediate products of e14,e15,e16,e18,e19
call initKTeile(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)


'assign objects to arrays
kteileArray = Array(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)
eteileArray = Array(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190)
wsArray = Array(ws1,ws2,ws3,ws4,ws5,ws6,ws7)

'dynamic array for order id's				
dim idArrayK()
dim idArrayE()

'*********************************************************
'delete selected entries 
'after recursive call of script
'*********************************************************
deleteID = Request.Form("deleteID")
deleteTable = Request.Form("deleteTable")
'response.write "DeleteID: " & DeleteID & "<br>"
'response.write "deleteTable: " & deleteTable & "<br>"

'response.write "vartype(id): " & vartype(id) & "<br>"
if deleteID <> vbEmpty then
	if deleteTable = "Produktionsauftraege" then
		deleteIDArray = split(deleteID, ",")
		deleteID = ""
		if deleteIDArray(1) = "void" then
			deleteID = deleteIDArray(0)
		else
			deleteID = deleteIDArray(0)
			deleteID2 = deleteIDArray(1)
			sqlDelete = "DELETE FROM " &deleteTable& " WHERE id="&deleteID2&""
			'response.write "deleteID2: " & deleteID2 & "<br>"
			Connection.execute(sqlDelete)
		end if
	end if

	'response.write "deleteID: " & deleteID & "<br>"
	'response.write "table: " & deleteTable & "<br>"
	sqlDelete = "DELETE FROM " &deleteTable& " WHERE id="&deleteID&""
	Connection.execute(sqlDelete)
end if

%>

<!------------------------------------- END OF DELETE SECTION ---------------------- -->


<div id="divContainer" style="position: absolute; border: 0px solid blue">

<!-- formular -->
<form method="POST" onsubmit="return checkInput();" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">
<input type="hidden" name="formularType" value="purchase">
<input type="hidden" name="d" value="void">
<!--<form method="POST" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">-->


<div id ="divKTeile" style="position: relative; border: 0px solid red; width: 90%; float: none; overflow: visible;">

<%
response.write "<h1 align='center'>Eingabeformular für Gruppe: " & groupObj.name & ", Spielperiode: " & period & "</h1>"
%>
<br><br>

<%
'*****************************************************
' Formular "Kaufteile"
'*****************************************************
response.write "<h3>Bestellungen von Kaufteilen für die <br>kommende Woche/Spielperiode Nr." &period& "</h3>"
sql = "SELECT * FROM Kaufteilbestellungen WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND geliefert = 0 ORDER BY id"
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection

positionFlag = "first"

If Recordset.EOF Then
	Response.Write("Es sind noch keine Kaufteilbestellungen von Gruppe " &usergroup& " für Spielperiode " &period& " vorhanden.<br>")
	positionFlag = "firstLast"
	Set purchOrder = New PurchaseOrder
	with purchOrder
		.id = "void"
		.purchaseOrderNo = 0
		.kteilnr = 0
		.amount = 0
		.delivered = 0 
	end with
	call writeInputRowK(purchOrder, positionFlag)
Else
	count = 1
	Do While NOT Recordset.Eof   
		Set purchOrder = New PurchaseOrder
		with purchOrder
			.id = CInt(Recordset("id"))
			.purchaseOrderNo = 0
			.kteilnr = CInt(Recordset("KTeilNr"))
			.amount = CInt(Recordset("Anzahl"))
			.delivered = CInt(Recordset("geliefert")) 
		end with
		
		' kteilnr = Recordset("KTeilNr")
		' kteilAmount = Recordset("Anzahl")
		' id = Recordset("id")
		Redim Preserve idArrayK(count)
		idArrayK(count-1) = purchOrder.id
		count = count + 1
		'response.write "id: " & id & "<br>"
		'call setKTeilObject(kteilnr,kteilObject)

		'******************************************
		'write input row containing current order
		'******************************************
		call writeInputRowK(purchOrder, positionFlag)
		'response.write "<br><br>"
		
		'next entry in recordset
		Recordset.MoveNext
		
		'if last entry is reached, write additional row of input fields
		if Recordset.EOF then		
			if positionFlag = "" then
				positionFlag = "last"
			elseif positionFlag = "first" then
				'just one entry
				positionFlag = "firstLast"
			end if
			
			'***************************************
			'write empty input row for next order
			'***************************************
			with purchOrder
				.id = "void"
				.purchaseOrderNo = 0
				.kteilnr = ""
				.amount = 0
				.delivered = 0 
			end with
			call writeInputRowK(purchOrder, positionFlag)		
		end if
	Loop
End If
Recordset.Close
set Recordset = Nothing

%>
</div> <!-- K-Teile -->

<!-- close formular, add send-button -->
<div id="submitForm" style="position: relative; left: 40%; margin-bottom:0px; border: 0px solid red; float: none;">
	<input type="submit" value="Absenden">
	<input type="reset" value="Abbrechen">
</div>
</form>


<%
'***********************************
'K-Teile Delete-Buttons
'***********************************
response.write "<div id='divDelK' style=""position: relative; top: 0%; left: 0%; border: 0px solid red; width: 10%; float: left; "">"
response.write "<table>"
response.write "<tr><th></th></tr>"

deleteTable = "Kaufteilbestellungen"
for each deleteID in idArrayK
	'response.write "deleteID: " & deleteID & "<br>"
	if deleteID <> vbEmpty then
		response.write "<tr>"
		response.write "<td>"
		response.write "<form method=""POST"" action='erpFormularPurchasePieces.asp' style='margin-bottom: 0;'>"
		response.write "<input type=""hidden"" name=""deleteID"" value="&deleteID&">"
		response.write "<input type=""hidden"" name=""deleteTable"" value="&deleteTable&">"
		'response.write "<td align='center'><input type=""submit"" name=""delete"" value='löschen'></td>"
		response.write "<div style='position:relative; top:0px;'><input type=""submit"" name=""delete"" value='löschen'></div>"
		response.write "</form>"
		response.write "</td>"
		response.write "</tr>"
	end if
next

response.write "</table>"
response.write "</div>"



Connection.close
set Connection = Nothing
%>

</div> <!-- container -->
</body>
</html>