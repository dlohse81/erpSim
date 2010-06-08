<% @ Language="VBScript" %>
<html>
<head>
	<title>ERP Simulation: Nutzereingaben</title>
	<!-- JavaScript functions -->
	<script type="text/javascript">
	
	function addKTeil() {
		table = document.getElementById("tableK").tBodies[0];
		//copy the last row
		var TRclone = table.rows[table.rows.length-1].cloneNode(true)
		//increment production number
		//TRclone.cells[0].childNodes[0].nodeValue++ 
		//set amount back to zero
		TRclone.cells[1].childNodes[0].value = 0 
		table.appendChild(TRclone)
	}
	function addETeil() {
		table = document.getElementById("tableE").tBodies[0];
		//copy the last row
		var TRclone = table.rows[table.rows.length-1].cloneNode(true)
		//increment production number
		TRclone.cells[0].childNodes[0].nodeValue++ 
		//set amount back to zero
		TRclone.cells[3].childNodes[0].value = 0 
		table.appendChild(TRclone)
	}
	function arrangeDOM() {
		var divContainer = document.getElementById("divContainer");
		var divETeile = document.getElementById("divETeile");
		var divKTeile = document.getElementById("divKTeile");
		var divDelE = document.getElementById("divDelE");
		var divDelK = document.getElementById("divDelK");
		var submitForm = document.getElementById("submitForm");
		
		//add delete-buttons to table "Kaufteilbestellungen"
		var rows = divDelK.firstChild.firstChild.childNodes;
		var tableK = document.getElementById("tableK").firstChild;
		for (var i=1; i<rows.length; i++) { //start with i=1 as rows[0] == header-row
			tableK.childNodes[i].appendChild(rows[i].firstChild);
		}
		
		//add delete-buttons to table "Produktionsauftraege"		
		var rows = divDelE.firstChild.firstChild.childNodes;
		var tableE = document.getElementById("tableE").firstChild;
		for (var i=1; i<rows.length; i++) { //start with i=1 as rows[0] == header-row
			//alert(rows[i].firstChild.chOff)
			rows[i].firstChild.vAlign = "bottom"
			//alert(rows[i].firstChild.vAlign)
			tableE.childNodes[i].appendChild(rows[i].firstChild);
			//tableE.replaceChild(rows[i].firstChild, tableE.childNodes[i]);
		}
		
		/*
		var childNodes = divContainer.childNodes
		//alert(childNodes.length)
		for (i=childNodes.length-1; i>=0; i--) {
			//alert(i)
			//alert(childNodes[i])
			//alert(childNodes[i].nodeType)
			
			chNode = childNodes[i]
			//alert(chNode.nodeType)
			
			if (chNode.nodeType == 1) {
				//alert("")
				divContainer.removeChild(chNode)
			}
		}
		
		
		divContainer.appendChild(divKTeile);
		divContainer.appendChild(divDelK);
		divContainer.appendChild(divETeile);
		//divContainer.appendChild(divDelE);
		divContainer.appendChild(submitForm);
		*/
	}
	
	</script>
</head>

<!-- END OF HEAD -->

<body onload="arrangeDOM()" style="heigth:100%">
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

'initialise KTeile and ETeile
call initETeile(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190) 'last 5 entries intermediate products of e14,e15,e16,e18,e19
call initKTeile(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)

'dynamic array for order id's				
dim idArrayK()
dim idArrayE()

'*********************************************************
'delete selected entries
'*********************************************************
deleteID = Request.Form("deleteID")
deleteTable = Request.Form("deleteTable")
response.write "DeleteID: " & DeleteID & "<br>"
response.write "deleteTable: " & deleteTable & "<br>"

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
			'Connection.execute(sqlDelete)
		end if
	end if

	'response.write "deleteID: " & deleteID & "<br>"
	'response.write "table: " & deleteTable & "<br>"
	sqlDelete = "DELETE FROM " &deleteTable& " WHERE id="&deleteID&""
	'Connection.execute(sqlDelete)
end if

%>

<!-- ----------------------------------- END OF DELETE SECTION ---------------------- -->


<div id="divContainer" style="position: absolute; border: 1px solid blue">

<!-- form -->
<form method="POST" action="erpUserInputWrite2DB.asp" name="usrInput" id="usrInputForm">


<div id ="divKTeile" style="position: relative; border: 1px solid red; width: 90%; float: none; overflow: visible;">
<h3>Bestellungen von Kaufteilen:</h3>
<%
'*****************************************************
' Kaufteile
'*****************************************************
sql = "SELECT * FROM Kaufteilbestellungen WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND geliefert = 0 ORDER BY id"
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection
positionFlag = "first"

If Recordset.EOF Then
	Response.Write("Es sind noch keine Kaufteilbestellungen von Gruppe " &usergroup& " für Spielperiode " &period& "vorhanden.<br>")
	positionFlag = "firstLast"
	call writeInputRowK("", "0", positionFlag)
Else
	count = 1
	Do While NOT Recordset.Eof   
		kteilnr = Recordset("KTeilNr")
		kteilAmount = Recordset("Anzahl")
		id = Recordset("id")
		Redim Preserve idArrayK(count)
		idArrayK(count-1) = id
		count = count + 1
		'response.write "id: " & id & "<br>"
		'call setKTeilObject(kteilnr,kteilObject)

		'write next input row
		call writeInputRowK(kteilnr, kteilAmount, positionFlag)
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
	
			call writeInputRowK("", "0", positionFlag)		
		end if
	Loop
End If
Recordset.Close
set Recordset = Nothing
%>
</div> <!-- K-Teile -->


<div id="divETeile" style="position: relative; border: 1px solid green; width: 100%; float: none; overflow: visible;">
<h3>Produktionsaufträge für Eigenfertigungsteile:</h3>

<%
'*****************************************************
' Eigenfertigungsteile
'*****************************************************
sql = "SELECT * FROM Produktionsauftraege WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND abgeschlossen = 0 ORDER BY Auftragsnr"
Set Recordset=Server.CreateObject("ADODB.Recordset")	
Recordset.Open sql, Connection
positionFlag = "first"

If Recordset.EOF Then
	Response.Write("Es sind noch keine Produktionsaufträge von Gruppe " &usergroup& " für Spielperiode " &period& "vorhanden.<br>")
	positionFlag = "firstLast"
	call writeInputRowE(pOrder, positionFlag)
Else
	count = 1
	Do While NOT Recordset.Eof  		
		eteilnr = CInt(Recordset("ETeilnr"))
		if eteilnr = 140 OR eteilnr = 150 OR eteilnr = 160 OR eteilnr = 180 OR eteilnr = 190 then
			'continue (with next iteration)
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
			
			if eteilnr = 140 OR eteilnr = 150 OR eteilnr = 160 OR eteilnr = 180 OR eteilnr = 190 then
				'continue
			else
				Redim Preserve idArrayE(count)
				idArrayE(count-1) = pOrder.id & "," & pOrder.eteilnr & "," & pOrder.prodOrderNo
				count = count + 1
			end if 	
			
			
			'get the object that represents the eteil
			'call setETeilObject(pOrder.eteilnr,eteilObject)
			
			'write next input row
			call writeInputRowE(pOrder, positionFlag)
			
			'next entry in recordset
			Recordset.MoveNext
						
			'if last entry is reached, write additional row of input fields
			if Recordset.EOF then
				pOrder.batchsizeRequired = 0
				pOrder.eteilnr = ""
				pOrder.day = ""
				pOrder.prodOrderNo = pOrder.prodOrderNo + 1
				
				if positionFlag = "" then
					positionFlag = "last"
				elseif positionFlag = "first" then
					positionFlag = "firstLast"
				end if
			
				call writeInputRowE(pOrder, positionFlag)
			end if
		end if
	Loop	
end if
	
Recordset.Close
set Recordset = Nothing

%>	
</div> <!-- E-Teile -->
<br>
<div id="submitForm" style="position: relative; left: 75%; margin-bottom:0px; border: 1px solid red; float: none;">
	<input type="submit" value="Absenden">
	<input type="reset" value="Abbrechen">
</div>
</form>




<%
'***********************************
'K-Teile Delete-Button
'***********************************
response.write "<div id='divDelK' style=""position: relative; top: 0%; left: 0%; border: 1px solid red; width: 10%; float: left; "">"
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
response.write "<div id='divDelE' style=""position: relative; top: 0px; left: 0%; border: 1px solid red; width: 10%; float: left;"">"
response.write "<table>"
response.write "<tr><th></th></tr>"

deleteTable = "Produktionsauftraege"
for each i in idArrayE
		'response.write "i: " & vartype(i) & "<br>"
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
				
			'response.write "deleteID2: " & deleteID2 & "<br>"	
			'response.write "deleteID: " & deleteID & "<br><br>"		
			deleteID = deleteID1 & "," & deleteID2
			
			response.write "<tr>"
			response.write "<td align='center'>"
			response.write "<form method=""POST"" action='erpUserInput.asp' >"
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