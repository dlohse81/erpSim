<!--***********************
erpUserInputWrite2DB.asp
*************************-->
<% 
@Language="VBScript" 
%>
<html>
<head>
	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->

	<!-- JavaScript -->
	<script type="text/javascript">
		function jsRedirect(formularType, d) {
			//alert("formType: " + formularType + "; d: " + d);
			
			//check formular type of calling script and redirect/reload to it
			if (formularType == "production") {
				//alert("in: production");
				location.href="erpFormularProduction.asp?d=" + d;
			}
			
			if (formularType == "purchase") {
				//alert("in: purchase");
				location.href = "erpFormularPurchasePieces.asp";
			}
		}
	</script>
</head>

<!--
<body onload="jsRedirect()" style="height:100%">
-->
<%
'get usergroup
set groupObj = New ERPGroup
call getUsergroup(usergroup, groupObj)

'get game period
call getPeriod(period)

'initialise DB
call initDB(Connection)

'******************************************
'get user input data from formular
'and write orders to DB
'******************************************
formularType = Request.Form("formularType")
'response.write "formularType: " & formularType & "<br>"
d = Request.Form("d")
'response.write "d: " & d & "<br>"

response.write "Bitte warten Sie während Aufträge und Bestellungen werden in die Datenbank eingetragen werden.<br><br>"
response.write "Die aktualisierte Eingabemaske sollte sich in wenigen Sekunden öffnen, andernfalls schließen Sie das Popup-Fenster.<br>"

'response.write "<body onload='jsRedirect()' style='height:100%'>"
'response.write "<body onload=""jsRedirect()"" style=""height:100%"">"
'response.write "<body onload=""jsRedirect(formularType)"" style=""height:100%"">"
response.write "<body onload=""jsRedirect('"&formularType&"', '"&d&"')"" style=""height:100%"">"


'***************
'purchase pieces
'***************
if formularType = "purchase" then
	Set purchOrders = New PurchaseOrder
	with purchOrders
		.id = split(Request.Form("purchOrderId"), ",")
		.purchaseOrderNo = 0
		.kteilnr = split(Request.Form("kteile"), ",")
		.amount = split(Request.Form("amountKTeile"), ",")
		.delivered = "void"
	end with

	'**********************************
	'write kteile-orders to db
	'**********************************	
	' response.write "purchOrders.id: " & purchOrders.id(0) & "<br>"
	' response.write "UBound(purchOrders.id): " & UBound(purchOrders.id) & "<br>"
	' response.write "UBound(purchOrders.kteilnr): " & UBound(purchOrders.kteilnr) & "<br>"
	if UBound(purchOrders.kteilnr) <> -1 then 'if there is no purchase order at all and none is made
		for i=0 to UBound(purchOrders.id)
			if trim(purchOrders.kteilnr(i))<>"" then 'last row contains no entries
				'response.write "in<br>"
				if trim(purchOrders.id(i))="void" then 'new entry (not yet in DB)
					sqlInsert = "INSERT INTO Kaufteilbestellungen (usergroup, Periode, KTeilNr, Anzahl, geliefert) " _
						& "VALUES ("&usergroup&", "&period&", "&purchOrders.kteilnr(i)&", "&purchOrders.amount(i)&", 0)"
					'response.write "sqlInsert: " & sqlInsert & "<br>"
					Connection.execute(sqlInsert)
				else 'old entry, update DB
					sqlUpdate = "UPDATE Kaufteilbestellungen SET KTeilNr="&purchOrders.kteilnr(i)&", Anzahl="&purchOrders.amount(i)&" WHERE id="&purchOrders.id(i)&" "
					'response.write "sqlUpdate: " & sqlUpdate & "<br>"
					Connection.execute(sqlUpdate)
				end if
			end if
		next
	end if

	
'******************
'production pieces
'******************	
elseif formularType = "production" then		
	Set pOrders = New ProdOrder
	with pOrders
		.id = split(Request.Form("pOrderId"), ",") 'id of purchase order in DB
		.prodOrderNo = split(Request.Form("prodOrderNo"), ",") 'production order number
		.eteilnr = split(Request.Form("eteile"), ",")
		.day  = split(Request.Form("tage"), ",")
		.batchsizeRequired = split(Request.Form("batchsizeRequired"), ",")
		.finished = "void" 'Anzahl der bereits gefertigten Teile vom Vortag, falls Los nicht komplett abgearbeitet werden konnte
		'.batchsize
	end with


	'response.write "wsworkingtime(i): " & vartype(pOrders)<>vbEmpty & "<br>"

	wsWorkingTime = replace(Request.Form("selectWorkingTime"), " h, ", ",")
	wsWorkingTime = replace(wsWorkingTime, " h", "")
	wsWorkingTime = split(wsWorkingTime, ",")


	'response.write "purchOrders.id: " & purchOrders.id(1) & "; .prodOrderNo: " & purchOrders.kteilnr(1) & "<br>"



	'****************************************
	'write workingtime of workstations to db
	'****************************************
	'counter=1 '(workstation) counter
	'd=1 'day
	for i=0 to 6 '7 workstations 
		'response.write i & ": wsworkingtime(i): " & wsworkingtime(i) & "<br>"
		sqlUpdate = "UPDATE Arbeitszeiten SET Arbeitszeit="&wsWorkingTime(i)&" " _
			& "WHERE usergroup="&usergroup&" AND Periode="&period&" AND Tag="&d&" AND ArbeitsplatzNr="&i+1&" "
		Connection.Execute(sqlUpdate)
		'response.Write "sqlUpdate: " & sqlUpdate & "<br>"
	next

	' for i=0 to 34 '7 workstations * 5 days = 35
		' response.write i & ": wsworkingtime(i): " & wsworkingtime(i) & "<br>"
		' sqlUpdate = "UPDATE Arbeitszeiten SET Arbeitszeit="&wsWorkingTime(i)&" " _
			' & "WHERE usergroup="&usergroup&" AND Periode="&period&" AND Tag="&d&" AND ArbeitsplatzNr="&counter&" "
		' Connection.Execute(sqlUpdate)
		' response.Write "sqlUpdate: " & sqlUpdate & "<br>"
		
		' if counter=7 then
			' d=d+1
			' counter=0
		' end if
		
		' counter=counter+1
	' next


	'response.Write "<br><br>"


	'**************************************
	'write eteile-orders to db
	'**************************************
	'get maximum production order number and increase it
	'----------------------------------------------------
	sql = "SELECT usergroup, MAX(Auftragsnr) AS Auftragsnr FROM Produktionsauftraege GROUP BY usergroup"
	Set Recordset=Server.CreateObject("ADODB.Recordset")	
	Recordset.Open sql, Connection
	groupFound = 0

	If Recordset.EOF Then
			'currently, there are no production orders at all
			'response.write "in<br>"
			prodOrderNo = 1
	Else
		Do While NOT Recordset.Eof   
			ugroup = Recordset("usergroup")
			prodOrderNo = Recordset("Auftragsnr")
			'response.write "prodOrderNo: " & prodOrderNo & "<br>"
			
			if ugroup = usergroup then
				'increase production order number
				prodOrderNo = prodOrderNo + 1 
				groupFound = 1
				exit do
			end if
			
			'next entry in recordset
			Recordset.MoveNext
		Loop
		
		if groupFound = 0 then
			'currently, there is no production order from the usergroup
			prodOrderNo = 1
		end if
	end if
	Recordset.close
	Set Recordset = Nothing

	'response.write "<br>"
	'response.write "prodOrderNo: " & prodOrderNo & "<br>"



	'iterate through production orders
	'-------------------------------------
	'response.write "UBound(pOrders.id): " & UBound(pOrders.id) & "<br>"

	if UBound(pOrders.eteilnr) <> -1 then 'if there is no production order at all and none is made
		for i=0 to UBound(pOrders.id)
			'response.write "entry<br>"
			if trim(pOrders.eteilnr(i))<>"" then 'last row contains no entries
				select case pOrders.eteilnr(i)
					case 14
						e = 140
					case 15
						e = 150
					case 16
						e = 160
					case 18
						e = 180
					case 19
						e = 190
					case else 
						e=0
				end select
				
				
				if trim(pOrders.id(i))<>"void" then   
					'***********************
					'old entry, update DB	
					'***********************
					sql = "SELECT * FROM Produktionsauftraege WHERE id="&pOrders.id(i)&" "
					Set Recordset2=Server.CreateObject("ADODB.Recordset")	
					Recordset2.Open sql, Connection
					eteilnrOld = CInt(Recordset2("ETeilnr"))
					'prodOrderNoOld = CInt(Recordset2("Auftragsnr")
					Recordset2.close
					set Recordset2 = Nothing
					
					'if new order differs from old one (otherwise do nothing as nothing has changed)
					if eteilnrOld <> pOrders.eteilnr(i) then
						'if there is a manufacturing chain for the desired eteil, write additional production order for the intermediate product to db
						'(but with the same production-order-number as the intermediate piece belongs directly to its original production order 
						'and must not exist separately!)

						'check if there is already an old intermediate-piece order
						select case eteilnrOld
							case 14
								e_i = 140
							case 15
								e_i = 150
							case 16
								e_i = 160
							case 18
								e_i = 180
							case 19
								e_i = 190
							case else
								e_i = 0
						end select
						'response.write "e_i: " & e_i & "<br>"
						
						if e<>0 then
							'updated order is a production chain
							'====================================					
							if e_i<>0 then
								'old order was production chain, too
								'------------------------------------
								'update intermediate-piece order 
								sqlUpdate = "UPDATE Produktionsauftraege SET Tag="&pOrders.day(i)&", ETeilnr="&e&", Losgroesse="&pOrders.batchsizeRequired(i)&" " _
									& "WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Auftragsnr = "&pOrders.prodOrderNo(i)&" AND ETeilnr = "&e_i&" "
								'response.write "sqlUpdate: " & sqlUpdate & "<br>"
								Connection.execute(sqlUpdate)
								
							else
								'old order was no production chain (updated order is)
								'-----------------------------------------------------
								'write new intermediate-piece order
								sqlInsert = "INSERT INTO Produktionsauftraege (Auftragsnr, usergroup, Periode, Tag, ETeilNr, Losgroesse, abgeschlossen) " _
									& "VALUES ("&pOrders.prodOrderNo(i)&", "&usergroup&", "&period&", "&pOrders.day(i)&", "&e&", "&pOrders.batchsizeRequired(i)&", 0)"
								'response.write "sqlInsert: " & sqlInsert & "<br>"
								Connection.execute(sqlInsert)
							end if
						else
							'updated order is no production chain
							'=================================
							if eteilnrOld=14 OR eteilnrOld=15 OR eteilnrOld=16 OR eteilnrOld=18 OR eteilnrOld=19 then
								'old order was a production chain
								'---------------------------------
								'delete the intermediate-piece order
								sqlDelete = "DELETE FROM Produktionsauftraege WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Auftragsnr = "&pOrders.prodOrderNo(i)&" " _
									& " AND ETeilnr = "&e_i&" "
								'response.write "sqlDelete: " & sqlDelete & "<br>"
								Connection.execute(sqlDelete)
							end if
						end if
						
						'update order 
						sqlUpdate = "UPDATE Produktionsauftraege SET Tag="&pOrders.day(i)&", ETeilnr="&pOrders.eteilnr(i)&", Losgroesse="&pOrders.batchsizeRequired(i)&" " _
							& "WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Auftragsnr = "&pOrders.prodOrderNo(i)&" AND ETeilnr = "&eteilnrOld&" "
						'response.write "sqlUpdate: " & sqlUpdate & "<br>"
						Connection.execute(sqlUpdate)
					end if 'new order differs from old one
				
				else 
					'***************************
					'new entry (not yet in DB)
					'***************************
					'response.write "pOrders.id(i) == ""void"" <br>"
					if e<>0 then
						'new order is a production chain
						'=================================	
						'write new intermediate-piece order
						sqlInsert = "INSERT INTO Produktionsauftraege (Auftragsnr, usergroup, Periode, Tag, ETeilNr, Losgroesse, abgeschlossen) " _
							& "VALUES ("&prodOrderNo&", "&usergroup&", "&period&", "&pOrders.day(i)&", "&e&", "&pOrders.batchsizeRequired(i)&", 0)"
						'response.write "sqlInsert: " & sqlInsert & "<br>"
						Connection.execute(sqlInsert)
						'prodOrderNo = prodOrderNo + 1
					end if	

					' "normal" production orders
					sqlInsert = "INSERT INTO Produktionsauftraege (Auftragsnr, usergroup, Periode, Tag, ETeilNr, Losgroesse, abgeschlossen) " _
						& "VALUES ("&prodOrderNo&", "&usergroup&", "&period&", "&pOrders.day(i)&", "&pOrders.eteilnr(i)&", "&pOrders.batchsizeRequired(i)&", 0)"
					Connection.execute(sqlInsert)
					'response.write "sqlInsert: " & sqlInsert & "<br>"
					
					'increase prodOrderNo
					prodOrderNo = prodOrderNo + 1
				
				end if 'old (update) or new entry (add only)
			end if
		next
	end if
else
	'response.write "ERROR. Variable 'formularType' contains neither the value 'production' nor 'purchase'<br>"
end if


'close DB connection
Connection.close
set Connection = Nothing


'if formularType = "purchase" then
	'response.write "redirect-case 'purchase'<br>"
	'response.write "<body onload='jsRedirect('purchase')'>"
	'Response.Redirect("erpFormularPurchasePieces.asp")
	'Response.End
'elseif formularType = "production" then
	'response.write "redirect-case 'production'<br>"
	'response.write "<body onload='jsRedirect('production')'>"
	'Response.Redirect("erpFormularProduction.asp")
	'Response.End
'else 
	'Response.write "ERROR. Variable 'formularType' contains neither the value 'production' nor 'purchase'<br>"
'end if
%>

</body>
</html>