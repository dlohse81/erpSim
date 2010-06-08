<% @ Language="VBScript" %>
<html>
<head>
	<title>ERP Simulation: Nutzereingaben</title>
	<!-- JavaScript functions -->
	<script type="text/javascript">
	
	function addKTeilRow() {
		table = document.getElementById("tableK").tBodies[0];
		//copy the last row
		var TRclone = table.rows[table.rows.length-1].cloneNode(true)
		//increment production number
		//TRclone.cells[0].childNodes[0].nodeValue++ 
		//set amount back to zero
		TRclone.cells[1].childNodes[0].value = 0 
		table.appendChild(TRclone)
	}
	
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
	
	function checkInput() {
		
		//return false;
		
	}
	
	function arrangeDOM() {
		var divContainer = document.getElementById("divContainer");
		var divETeile = document.getElementById("divETeile");
		var divKTeile = document.getElementById("divKTeile");
		var divDelE = document.getElementById("divDelE");
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
	}
	
	</script>
</head>

<!-- END OF HEAD -->

<body onload="arrangeDOM()" style="height:100%">
<!-- #include file="erpClasses.asp" -->
<!-- #include file="erpProcedures.asp" -->



<!-- ----------------------------------- DELETE SECTION ---------------------- -->
<%
'**************************************
'initialisations
'**************************************
'get usergroup
call getUsergroup(usergroup)

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

<!-- ----------------------------------- END OF DELETE SECTION ---------------------- -->


<div id="divContainer" style="position: absolute; border: 0px solid blue">

<!-- form -->
<form method="POST" onsubmit="return checkInput();" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">
<!--<form method="POST" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">-->


<div id ="divKTeile" style="position: relative; border: 0px solid red; width: 90%; float: none; overflow: visible;">

<%
response.write "<h1 align='center'>Eingabeformular für Gruppe: " & usergroup & " Spielperiode: " & period & "</h1>"
%>
<br><br>

<%
'*****************************************************
' Formular "Kaufteile"
'*****************************************************
response.write "<h3>Bestellungen von Kaufteilen:</h3>"
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


<div id="divETeile" style="position: relative; border: 0px solid green; width: 100%; float: none; overflow: visible;">




<%
'*****************************************************
' Formular "days"
'*****************************************************
count = 1
'iterate through weekdays (Monday - Friday)
for d=1 to 5
	response.write "<h2>Tag " & d & "</h2>"
	response.write "<h3>Produktionsaufträge für Eigenfertigungsteile:</h3>"
	
	'**************
	'E-Teile
	'**************
	sql = "SELECT * FROM Produktionsauftraege WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Tag="&d&" AND abgeschlossen = 0 ORDER BY Auftragsnr, id"
	Set Recordset=Server.CreateObject("ADODB.Recordset")	
	Recordset.Open sql, Connection
	positionFlag = "first"

	If Recordset.EOF Then
		Response.Write("Es sind noch keine Produktionsaufträge von Gruppe " &usergroup& " für Tag "&d&" in Spielperiode " &period& " vorhanden.<br>")
		positionFlag = "firstLast"
		Set pOrder = New ProdOrder
		with pOrder
			.id = "void"
			.prodOrderNo = "void"
			.eteilnr = 0
			.day  = d
			.batchsizeRequired = 0
			.finished = 0'Anzahl der bereits gefertigten Teile vom Vortag, falls Los nicht komplett abgearbeitet werden konnte
			'.batchsize
		end with
		call writeInputRowE(pOrder, positionFlag)
	Else
		Do While NOT Recordset.Eof  		
			eteilnr = CInt(Recordset("ETeilnr"))
			if eteilnr = 140 OR eteilnr = 150 OR eteilnr = 160 OR eteilnr = 180 OR eteilnr = 190 then
				'continue (with next iteration)
				'response.write "intermediate product<br>"
				Recordset.MoveNext
				
				set eteilnr = Nothing
			else
				Set pOrder = New ProdOrder
				with pOrder
					.id = CInt(Recordset("id"))
					.prodOrderNo = CInt(Recordset("Auftragsnr"))
					.eteilnr = CInt(Recordset("ETeilnr"))
					.day  = CInt(Recordset("Tag"))
					.batchsizeRequired = CInt(Recordset("Losgroesse"))
					.finished = CInt(Recordset("abgeschlossen")) 'Anzahl der bereits gefertigten Teile vom Vortag, falls Los nicht komplett abgearbeitet werden konnte
					'.batchsize
				end with
				
				Redim Preserve idArrayE(count)
				idArrayE(count-1) = pOrder.id & "," & pOrder.eteilnr & "," & pOrder.prodOrderNo
				count = count + 1
				
				
				'get the object that represents the eteil
				'call setETeilObject(pOrder.eteilnr,eteilObject)
				
				'write next input row
				call writeInputRowE(pOrder, positionFlag)
				
				'next entry in recordset
				Recordset.MoveNext
			end if				
				
		
			'if last entry is reached, write additional row of input fields
			'response.write "last?<br>"
			if Recordset.EOF then
				pOrder.id = "void"
				pOrder.batchsizeRequired = 0
				pOrder.eteilnr = ""
				pOrder.day = d
				pOrder.prodOrderNo = "void"
				'pOrder.prodOrderNo = pOrder.prodOrderNo + 1
				
				if positionFlag = "" then
					positionFlag = "last"
				elseif positionFlag = "first" then
					positionFlag = "firstLast"
				end if
				call writeInputRowE(pOrder, positionFlag)
			end if
		Loop	
	end if
		
	Recordset.Close
	set Recordset = Nothing
	
	'************************
	'workingTime
	'************************	
	Response.Write "<h3>Arbeitszeiten</h3>"
	Response.Write "<table name='workingTimes'cellpadding=10>"
	Response.Write "<tr>"
	Response.Write "<th>Arbeitsplatz</th><th>Arbeitszeit</th>"
	Response.Write "</tr>"
	
	'iterate through workstations
	for i=1 to 7 
	    'get workingtime on workstation "i"
	    'call setWSObject(i, wsObject)
		set wsObject = wsArray(i-1)
	    wsObject.loadWorkingTime(d)
	 
	    Response.Write "<tr>"
	    Response.Write "<td align='center'>"& i & "</td>"
	    Response.Write "<td align='center'>"
	    Response.Write "<select name=""selectWorkingTime"">"
	   
	    'iterate through workingtime-hours
        for j=0 to 24
            if j=wsObject.workingTime then 
                response.Write "<option selected>"&j&" h </option>"
            else
                response.Write "<option>"&j&" h</option>"
            end if
        next      
        Response.Write "</select>"
        Response.Write "</td>"
        Response.Write "</tr>"
    next
    
    Response.Write "</table>"
next
%>	


</div> <!-- E-Teile -->
<br>
<div id="submitForm" style="position: relative; left: 75%; margin-bottom:0px; border: 0px solid red; float: none;">
	<input type="submit" value="Absenden">
	<input type="reset" value="Abbrechen">
</div>
</form>




<%
'***********************************
'K-Teile Delete-Button
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
		response.write "<form method=""POST"" action='erpUserInput.asp' style='margin-bottom: 0;'>"
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


'***********************************
'E-Teile Delete-Button
'***********************************
response.write "<div id='divDelE' style=""position: relative; top: 0px; left: 0%; border: 0px solid red; width: 10%; float: left;"">"
response.write "<table>"
response.write "<tr><th></th></tr>"

deleteTable = "Produktionsauftraege"
for each i in idArrayE
		'response.write "i in idArrayE: " & i & "<br>"
		if i <> vbEmpty then
			iArray = split(i, ",")
			deleteID1 = iArray(0)
			eteilnr = iArray(1)
			prodOrderNo = iArray(2)
			' response.write "deleteID1: " & deleteID1 & "<br>"
			' response.write "eteilnr: " & eteilnr & "<br>"
			' response.write "prodOrderNo: " & prodOrderNo & "<br><br>"

			if eteilnr = 14 OR eteilnr = 15 OR eteilnr = 16 OR eteilnr = 18 OR eteilnr = 19 then
				select case eteilnr
					case 14
						eteilnr = 140
					case 15
						eteilnr = 150
					case 16
						eteilnr = 160
					case 18
						eteilnr = 180
					case 19
						eteilnr = 190				
				end select 
				sql = "SELECT * FROM Produktionsauftraege WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Auftragsnr = "&prodOrderNo&" " _
					& " AND ETeilnr = "&eteilnr&" "
				Set Recordset=Server.CreateObject("ADODB.Recordset")	
				Recordset.Open sql, Connection
				
				if Recordset.EOF then
					response.write "ERROR! Cannot delete corresponding intermediate product. SQL found no entry in DB.<br>"
					deleteID2 = "void"
				else
					deleteID2 = Recordset("id")
					' response.write "proOrderNo: " & prodOrderNo & "<br>"
					' response.write "eteilnr: " & eteilnr & "<br>"
					' response.write "deleteID2: " & deleteID2 & "<br><br>"
				end if
				
				Recordset.Close
				set Recordset = Nothing
			else
				deleteID2 = "void"
			end if	
				
			' response.write "deleteID2: " & deleteID2 & "<br>"	
			' response.write "deleteID: " & deleteID & "<br><br>"		
			deleteID = deleteID1 & "," & deleteID2
			
			response.write "<tr>"
			response.write "<td align='center'>"
			response.write "<form method=""POST"" action='erpUserInput.asp' style='margin-bottom: 0;'>"
			response.write "<input type=""hidden"" name=""deleteID"" value="&deleteID&">"
			response.write "<input type=""hidden"" name=""deleteTable"" value="&deleteTable&">"
			response.write "<input type=""submit"" name=""delete"" value='löschen'>"
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