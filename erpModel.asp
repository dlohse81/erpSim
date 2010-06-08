<% @Language="VBScript" %>
<html id="html">
<head>
	<title>CeTIM ERP Simulation</title>

	<!-- #include file="erpClasses.asp" -->
	<!-- #include file="erpProcedures.asp" -->

	<!--#include file="../lib.inc"-->
	<!--#include file="../defaults.inc"-->
	
	<!-- JavaScript functions -->
	<script type="text/javascript">			
		function savePage() {
			var htmlCode = document.getElementById('container').innerHTML;
			htmlCode = "<html><head></head><body>" + htmlCode + "</body></html>";
			//alert(htmlCode);
			
			var htmlCodeUnicode;
			htmlCodeUnicode = encodeURIComponent(htmlCode);
			//alert(htmlCodeUnicode);
			
			//clear screen, otherwise form.submit() won't work for a strange reason
			document.write("Speichere Resultate in Datenbank zum späteren Abruf. Dies kann einen Moment in Anspruch nehmen. Bitte warten.");
			
			//create form with hidden input field sending page-HTML as HEX to savePage.asp
			document.write("<form method='POST' action='erpSaveResults.asp' name='formularSavePage'><input type='hidden' name='htmlCodeUnicode' value='"+htmlCodeUnicode+"'></form>");
			//document.write("<form method='POST' action='erpSaveResults.asp' name='formularSavePage'><input type='hidden' name='htmlCodeUnicode' value='"+htmlCodeUnicode+"'><input type='hidden' name='htmlCode' value='"+htmlCode+"'></form>");
			
			//submit form, i.e. call script erpSavePage.asp
			document.forms.formularSavePage.submit();
			
			
			/*
			var htmlCodeHex = encodeToHex(htmlCode);
			
			var htmlCodeUnicode;
			htmlCodeUnicode = htmlCode.charCodeAt(0);;
			for (var i = 1; i <= htmlCode.length-1; i++) {
				htmlCodeUnicode = htmlCodeUnicode + "%" + htmlCode.charCodeAt(i);
			}
			alert(htmlCodeUnicode);
			*/
			
			/*
			//document.write("htmlCodeHex: " + htmlCodeHex+"<br>"); 
			//alert("htmlCodeHex.length: " +htmlCodeHex.length);
			
			//htmlCode = decodeFromHex(htmlCodeHEx);
			//document.write("htmlCode.length: " +htmlCode.length+"<br>");
			//document.write("htmlCode: " + htmlCode+"<br>");
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
<body onload="savePage()" style="height:100%">



<div id="container">

<%

'dynamic array for stock holding				
dim holding()

'dynamic array for workload of a working station
'dim workload()

'get usergroup
dim usergroup
set groupObj = New ERPGroup
call getUsergroup(usergroup, groupObj)
'response.write "usergroup: " & usergroup & "<br>"


'get game period
dim period
call getPeriod(period)

'initialise first day of the week
dim dayOld
dayOld = 0

'initialise DB
dim Connection
call initDB(Connection)

'create objects: workstations, ETeile, KTeile, stock
call initWS(ws1,ws2,ws3,ws4,ws5,ws6,ws7)
call initETeile(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190) 'last 5 entries intermediate products of e14,e15,e16,e18,e19
call initKTeile(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)

'assign objects to arrays
kteileArray = Array(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)
eteileArray = Array(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190)
wsArray = Array(ws1,ws2,ws3,ws4,ws5,ws6,ws7)

'set stockObject = New Stock
set purchOrder = New PurchaseOrder

'dim Recordset, sql, Connection

' set out = new Output
' with out
	' .inputs = ""
	' .results = ""
	' .stock = ""	
	' .dtg = "21:45:00"
' end with
'out.echo("<html><body>")

Set stockObj = New Stock
stockObj.init()

'backup current holding for later statistics
stockObj.backupHolding()

'**************************************
'iterate through days
'**************************************
for d=1 to 5		
	call writeHeader(usergroup, period, d)
	positionFlag = "first"
	'***********************
	'get purchase orders
	'***********************

	sql = "SELECT * FROM Kaufteilbestellungen WHERE usergroup="&usergroup&" AND Periode="&period&" AND geliefert=0 ORDER BY id"
	Set Recordset5=Server.CreateObject("ADODB.Recordset")	
	Recordset5.Open sql, Connection
	
	'***********************************
	'iterate through purchase orders
	'***********************************
	If Recordset5.EOF Then
		Response.Write("Keine Lagerzugänge aus Bestellungen.<br>")
	Else	
		Do While NOT Recordset5.Eof   
			
			kteilnr = Recordset5("KTeilNr")
			
			'get the object that represents the KTeil given via "kteilnr"
			dim kteilObject
			'call setKTeilObject(kteilnr,kteilObject)
			call setItemObject(kteilnr, kteilObject)
			'set kteilObject = kteileArray(kteilnr-20)
			'response.write kteilObject.name & "<br>"			
			
			with purchOrder
				.id = Recordset5("id")
				.purchaseOrderNo = 0
				.kteilnr = Recordset5("KTeilNr")
				.amount = Recordset5("Anzahl")
				.delivered = Recordset5("geliefert")
			end with
			
			deliveryDay = kteilObject.deliveryTime + 1 'because the delivery can be found on stock not until the next day after the given delivery time
			if deliveryDay = d then
				'calculate and save new holding of kteil
				'response.write "kteilObject.holding: " & kteilObject.holding & "<br>"
				'response.write "setting kteilAmount " & kteilObject.holding + purchOrder.amount & "<br>"
				kteilObject.setHolding(kteilObject.holding + purchOrder.amount)
				
				'deliver K-Teile
				purchOrder.deliver()
				
				'next recordset
				Recordset5.MoveNext
				
				if Recordset5.EOF AND positionFlag = "" then
					positionFlag = "last"
				elseif Recordset5.EOF AND positionFlag = "first" then
					'only one purchase order
					positionFlag="firstLast"
				end if
				
				call purchOrder.writePurchases(kteilObject, positionFlag)
				'call writePurchases(positionFlag)
			else 
				Recordset5.MoveNext
			end if		
		Loop
		if positionFlag="first" then
			response.write "Keine Lagerzugänge aus Bestellungen.<br>"
		end if
		response.write "</table>"
	end if
	
	Recordset5.close
	set Recordset5 = Nothing
	response.write("<br><br><br>")
	
	
	
	
	
	'*************************************************************************************
	
		
	'***********************
	'get production orders
	'***********************
	response.write "<h4>Aus Produktion:</h4>"
	sql = "SELECT * FROM Produktionsauftraege WHERE usergroup="&usergroup&" AND Periode="&period&" AND Tag<=" &d& " AND abgeschlossen < Losgroesse ORDER BY abgeschlossen DESC, Tag, Auftragsnr, ETeilnr DESC"
	'call readDB(sql, Recordset, Connection)
	Set Recordset=Server.CreateObject("ADODB.Recordset")	
	Recordset.Open sql, Connection
	
	positionFlag = "first"
	
	'load workstations for statistics
	for each workstation in wsArray
		workstation.loadWorkingTime(d)
		'workstation.reset()
	next
	
	'***********************************
	'iterate through production orders
	'***********************************
	If Recordset.EOF Then
		Response.Write("Keine Produktionsaufträge für Tag " & d & " vorhanden.<br><br>")
	Else
		Do While NOT Recordset.Eof   	
			'set production-order object for current order
			'----------------------------------------------
			Set pOrder = New ProdOrder
			with pOrder
				.id = CInt(Recordset("id"))
				.prodOrderNo = CInt(Recordset("Auftragsnr"))
				.eteilnr = CInt(Recordset("ETeilnr"))
				.day  = CInt(Recordset("Tag"))
				.batchsizeRequired = CInt(Recordset("Losgroesse"))
				.finished = CInt(Recordset("abgeschlossen")) 
				.abort = 0 
				.setBatchsize()
			end with
			'response.write "pOrder.eteilnr: " & pOrder.eteilnr & "<br>"
			'response.write "pOrder.batchsize: " & pOrder.batchsize & "<br>"
			
						
			'*******************************************
			'check if required workstation is available
			'*******************************************
			'get the object that represents the ETeil given in "pOrder.eteilnr"
			'--------------------------------------------------------------------
			dim eteilObject
			'call setETeilObject(pOrder.eteilnr, eteilObject)
			call setItemObject(pOrder.eteilnr, eteilObject)
			'eteilObject.getHolding() 'holding is loaded automatically when calling setItemObject()
			'response.write eteilObject.name & "<br>"
			
			dim wsObject
			'call setWSObject(eteilObject.ws, wsObject)
			set wsObject = wsArray(eteilObject.ws-1)
			wsObject.loadWorkingTime(d)
			'response.write "wsObject.nr: " & wsObject.nr & "; wsObject.t: " & wsObject.t & "; wsObject.workinTime: " & wsObject.workingTime & "<br>"
			
			'check if eteil-production is a manufacturing chain, if yes, check if intermediate pieces have been manufactured before 
			'(otherwise set pOrder.abort=1)
			pOrder.checkManufacturingChain()
			
			'*******************************************
			'check if there is sufficient working time 
			'left at the required workstation for the 
			'whole batch
			'*******************************************
			'call checkWSavailability(wsObject,pOrder)
			if pOrder.abort = 0 then
				wsObject.checkAvailability(pOrder) 'updates pOrder.batchsize to the batchsize that can actually be produced in the remaining time
			end if
			
			'if remaining time is sufficient
			if pOrder.abort = 0 then
				'**************************************************
				'check if there are sufficient input items on stock
				'**************************************************
				pOrder.checkHolding() 
				
				'set size of dynamic array "stock-holding"
				'ReDim holding(UBound(eteilObject.inputItemsAmount))
			end if	
			
		
			'*****************************
			'start the production
			'*****************************
			if pOrder.abort = 0 then	
				'******************************************
				'DATA UPDATES
				'******************************************
				'WORKSTATIONS
				'-------------
				'update working time of the used workstations
				'workload stored as a string: "wsNr,starttime,stoptime,duration,prodOrderNo,batchsize,eteilnr; ..."
				'Response.write "pOrder.prodOrderNo: " & pOrder.prodOrderNo & "pOrder.batchsize: " & pOrder.batchsize & "pOrder.eteilnr: " & pOrder.eteilnr & "<br>"
				wsObject.setWorkload(pOrder)
				'response.write "wsObject.workload: " & wsObject.workload & "<br>"
				'response.write "wsObject.workload: " & wsObject.workload & "<br>"
												
				
				'production order
				'------------------
				pOrder.manufacture()
				eteilObject.manufacture(pOrder.batchsize)
				'response.write "fertigung läuft...<br>"
				'response.write "pOrder.batchsizeReq: " & pOrder.batchsizeRequired & "; pOrder.batchsize: " & pOrder.batchsize & "<br>"
				
				
				
				
				
				'move finished orders into backup table
				'---------------------------------------
				'if production order is completely finished, erase it from order list
				if pOrder.batchsizeRequired = pOrder.batchsize then
					'copy finished production task from table "Produktionauftraege" into "ProduktionsauftraegeAlt"
					sqlInsert = "INSERT INTO ProduktionsauftraegeAlt SELECT Auftragsnr, usergroup, Periode, Tag, ETeilnr, Losgroesse, abgeschlossen FROM Produktionsauftraege WHERE id = "&pOrder.id&""
					'Connection.execute(sqlInsert)
					'delete finished production task from table "Produktionauftraege"
					sqlDelete = "DELETE FROM Produktionsauftraege WHERE id="&pOrder.id&""
					'Connection.execute(sqlDelete)
				elseif pOrder.batchsize < pOrder.batchsizeRequired then
					' response.write "Auftrag Nr. konnte nicht vollständig abgeschlossen werden. Es konnten von ETeil " &pOrder.eteilnr& " nur " &batchsize& " statt " &batchsizeRequired& _
					' " Stücke gefertigt werden. Der Auftrag wird daher morgen fortgesetzt.<br>"
				end if
			end if
		
			'***************************
			'footer and next order
			'***************************
			'fetch next production order
			Recordset.MoveNext 
			
			'set position flag for creation of output table
			if Recordset.EOF AND positionFlag = "" then
				positionFlag = "last"
			elseif Recordset.EOF AND positionFlag = "first" then
				'only one purchase order
				positionFlag="firstLast"
			end if
			
			
			call pOrder.writeProdOrder(eteilObject, positionFlag)
			'call writeProduction(pOrder, eteilObject, positionFlag)
			
			
		Loop 'fetch next production order
	End If
	
	
	
	'write workstations' workload for current day
	for each workstation in wsArray 
		'response.write "workstation.workingTime: " & workstation.workingTime & "<br>"
		workstation.writeSummary()
	next
	for each workstation in wsArray 
		workstation.writeWorkload()
	next
	
	'save weekly (total) state and reset workstations
	for each workstation in wsArray
		workstation.setTotals()
		workstation.reset()
	next
	'call initWS(ws1,ws2,ws3,ws4,ws5,ws6,ws7)
	'wsArray = Array(ws1,ws2,ws3,ws4,ws5,ws6,ws7)
next 'switch to the next day



'set day=="void" 
d = "void"

'header "Wochenzusammenfassung"
call writeHeader(usergroup, period, d)

for each i in wsArray
	i.writeTotalSummary()
next

'unfinished productions
Set pOrder = New ProdOrder
pOrder.writeUnfinishedOrdersWeek()

'"Lagerbestand"/stock holding
stockObj.getHolding()
stockObj.writeHolding()

'"Ergebnisrechnung"/profit-and-loss-statement
Set finance = New Finances
finance.writeProfitAndLossStatement(stockObj)

'close DB
Recordset.close
set Recordset = Nothing

'close DB connection
Connection.close
set Connection=Nothing

'out.echo("</body><html>")

'response.Write "out.results: " & out.results & "<br>"
%>
</div> <!-- container -->
</body>
</html>