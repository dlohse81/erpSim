<!--***********************
erpClasses.asp
*************************-->
<!--************************
	ASP classes
**************************-->
<%
class ETeil
	'properties
	public name 'ETeil name/task
	public nr 'nr des ETeils
	public stockCoeff '
	public ws 'required workstations in production process
	public cost	'(Material-)kosten, die bei der Produktion verursacht werden
	public usp 'unit sales price = Verkaufspreis
	public inputItems 'required items to produce the E-Teil
	public inputItemsAmount
	public holding
	
	'methods
	public function setHolding(holding)
		'saves the given amount/holding to DB and the object. 
		'ATTENTION: method replaces the current value by the given one. 
		'If an amount should be added to the previous value, add it to "holding" before you call the method (current holding via ETeil.getHolding)

		if holding<0 then
			response.write "ERROR in ETeil.setHolding(holding). Bestand < 0 wird gesetzt auf 0.<br>"
			holding = 0
		end if
	
		me.holding = holding
		
		'update database
		' if me.nr=38 then
			' response.write "updating eteilnr 38<br>"
		' end if
		
		'response.write "Saving item "new holding "&holding&" to stock.<br>"
		sqlUpdate = "UPDATE Lager SET Lagerbestand="&holding&" WHERE usergroup="&usergroup&" AND Teilnr="&me.nr&" " 'AND Periode="&period&" "
		Connection.execute(sqlUpdate)
	end function
	
	public function getHolding()
		'query
		sql = "SELECT * FROM Lager WHERE usergroup="&usergroup&" AND Teilnr="&me.nr&" " 'AND Periode="&period&" "		
		Set RecordsetE=Server.CreateObject("ADODB.Recordset")	
		RecordsetE.Open sql, Connection
		If RecordsetE.EOF Then
			Response.Write("ERROR! Teilnr " &me.nr& " ist nicht im Lager vorhanden.<br>")
		Else
			'store
			me.holding = RecordsetE("Lagerbestand")
		end if
		RecordsetE.close
		set RecordsetE=Nothing
		
		'return holding
		getHolding = me.holding
	end function
	
	public function manufacture(batchsize)
		'response.write "Manufacturing E-Teil "&me.nr&", batchsize "&batchsize&"<br>"
		'store new holding
		me.setHolding(me.holding + batchsize)
		'set counter for later statistics in stockObj
		call stockObj.setInput(me, batchsize)
		
		'reduce holding of input items
		'response.write "manufacturing eteil.nr: " & me.nr & "<br>"
		for i=0 to UBound(me.inputItems)
			'response.write "i: " & i & "; me.inputItems(i): " & me.inputItems(i) & "<br>"
			
			'get the itemObject
			call setItemObject(me.inputItems(i),itemObject)
			'response.write "reducing item.nr: " & item.nr & " to amount " & item.holding - me.inputItemsAmount(i)*batchsize & "<br>"
			
			'store new holding of item
			'response.write "setting holding of item " &item.nr& "; holding: " &item.holding - me.inputItemsAmount(i)*batchsize&"<br>"
			itemObject.setHolding(itemObject.holding - me.inputItemsAmount(i)*batchsize)
			
			'set counter for later statistics in stockObj
			call stockObj.setDispatch(itemObject, batchsize)
		next
	end function
end class

'****************************************************************************************************************

class KTeil
	'properties
	public name 'KTeil name/task
	public nr 'nr des KTeils
	public stockCoeff '
	public deliveryTime
	public cost 'purchase price of the KTeil
	public holding
	
	
	'methods
	public function getHolding()
		'query
		sql = "SELECT * FROM Lager WHERE usergroup="&usergroup&" AND Teilnr="&me.nr&" "	' AND Periode="&period&" "
		Set RecordsetK=Server.CreateObject("ADODB.Recordset")	
		RecordsetK.Open sql, Connection
		If RecordsetK.EOF Then
			Response.Write("ERROR! Teilnr " &me.nr& " ist nicht im Lager vorhanden.<br>")
		Else
			'store
			me.holding = RecordsetK("Lagerbestand")
		end if
		RecordsetK.close
		set RecordsetK=Nothing
		
		'return holding
		getHolding = me.holding
	end function
	
	public function setHolding(holding)
		'saves the given amount/holding to DB and the object. 
		'ATTENTION: method replaces the current value by the given one. 
		'If an amount should be added to the previous value, add it to "holding" before you call the method (current holding via KTeil.getHolding)
		if holding<0 then
			response.write "ERROR in KTeil.setHolding(holding). Bestand < 0 wird gesetzt auf 0.<br>"
			holding = 0
		end if
		
		'setter
		me.holding = holding
		
		'update database
		'response.write "updating kteil " & me.nr & "<br>"
		sqlUpdate = "UPDATE Lager SET Lagerbestand="&holding&" WHERE usergroup="&usergroup&" AND Teilnr="&me.nr&"" 'AND Periode="&period&" "
		Connection.execute(sqlUpdate)
		'sqlUpdate = "UPDATE Kaufteilbestellungen SET geliefert = "&kteilAmount&" WHERE id="&kteilOrderID&""
	end function
end class
	
'*********************************************************************************************************	
	
class ERPGroup
	'properties
	public id	'the id of the group in the erp.mdb table Gruppen
	public course 'the course, e.g. LRT04
	public group	'the group, e.g. 1
	public name		'the name of the group, combining "course/group" as LRT04/1
end class	
	
'*********************************************************************************************************	
	
class ProdOrder
	'properties
	public prodOrderNo 'production order number
	public eteilnr 'number of the eteil that shall be manfactured
	public day 'the day the prodOrder gets active and the eteil shall be manufactured
	public batchsizeRequired 'the batchsize that was ordered
	public batchsize 'the batchsize that was actually manufactured (might be smaller because of insufficient working time, material,...)
	public id 'production order id in database
	public finished 'amount of already finished pieces; this field is != 0 when the job was began on the previous day but was not finished
	public abort 'flag that indicates whether an order is to abort, because some of the preconditions (sufficient time, input pieces, ...) were not fulfilled
	
	'methods
	public function setBatchsize()
		if me.finished > 0 then
			me.batchsize = me.batchsizeRequired - me.finished
		else
			me.batchsize = me.batchsizeRequired
		end if
	end function
	
	public function checkHolding()
		'call setETeilObject(me.eteilnr, eteilObject)
		call setItemObject(me.eteilnr, eteilObject)
		eteilObject.getHolding()
		'iterate through required input items
		for i=0 to UBound(eteilObject.inputItems)
			item = CInt(eteilObject.inputItems(i))
			itemAmount = CInt(eteilObject.inputItemsAmount(i))
			'response.write "item: " & item & "; "
			'response.write "itemAmount: " & itemAmount & "<br>"

			'response.write "eteilObject.inputItems(i): " & eteilObject.inputItems(i) & "<br>"
			if eteilObject.inputItems(i) < 20 OR eteilObject.inputItems(i) > 100 then 
				'call setETeilObject(eteilObject.inputItems(i), inputItemObject)
				call setItemObject(eteilObject.inputItems(i), inputItemObject)
			elseif eteilObject.inputItems(i) > 19 then
				'call setKTeilObject(eteilObject.inputItems(i), inputItemObject)
				'set inputItemObject = kteileArray(eteilObject.inputItems(i)-20)
				call setItemObject(eteilObject.inputItems(i), inputItemObject)
			end if
			'set inputItemObject = new ETeil
			'inputItemObject.holding = 100
			'inputItemObject.getHolding()
			
			if inputItemObject.holding = 0 then
				'response.write "Es sind keine Stücke von Teilnr. " & item & " mehr im Lager vorhanden, E-Teilnr. " & me.eteilnr	&_
				response.write "Es sind keine Stücke von Teilnr. " & inputItemObject.nr & " mehr im Lager vorhanden, Auftragsnr. " & me.prodOrderNo	&_
				" kann daher heute nicht mehr gefertigt werden.<br><br>"
				'response.write "fall holding(i)=0 eingetreten.<br>"
				'preconditions(1)=0
				me.abort = 1
			else
				'reduce batchsize to find a batchsize that can be produced with the items on stock
				for j = me.batchsize to 1 step -1
					if inputItemObject.holding >= eteilObject.inputItemsAmount(i) * j then
						if j = me.batchsize then			
							'Response.write "Es sind ausreichend Teile vorhanden. Es werden " & itemAmount & "*" & batchsize & "="_
							'& itemAmount*batchsize &" Teile benötigt, " &holding(i) & " Stücke sind im Lager vorhanden.<br>"
							exit for 'only inner for-loop
						else
							me.batchsize = j
							response.write "<u>Auftrag " & me.prodOrderNo & ":</u><br>"
							response.write "Es können nur " &me.batchsize& " statt " &me.batchsizeRequired& " Stück von Teil " _
								&me.eteilnr& " gefertigt werden, da Teil " &item& " nur noch " &inputItemObject.holding&" mal statt " _
								&j*itemAmount& " mal auf Lager ist.<br>"
							exit for 'only inner for-loop
						end if
					else
						if j=1 then
							' response.write "Es sind nur noch " & holding(i) & " Teile von Teilnr. " & item & _
							' " im Lager vorhanden, diese genügen nicht mehr, um ein oder mehrere Stücke von E-Teilnr." & eteilnr & _
							' " zu fertigen.<br>"
							'response.write "fall me.batchsize=0 eingetreten.<br>"
							me.abort = 1
						end if
					end if
				next
			end if
		next
	end function
	
	public function checkManufacturingChain()	
		'if eteil is manufactured in production chain(e14,e15,e16,e18,e19) then check if intermediate products (e140,...,e190) 
		'have already been finished (otherwise break)
		select case me.eteilnr
			case 14
				set intermediateETeilObject = e140
				set intermediateWSObject = ws1
				mchain = 1
			case 15
				set intermediateETeilObject = e150
				set intermediateWSObject = ws2
				mchain = 1
			case 16			
				set intermediateETeilObject = e160
				set intermediateWSObject = ws1
				mchain = 1
			case 18
				set intermediateETeilObject = e180
				set intermediateWSObject = ws1
				mchain = 1
			case 19
				set intermediateETeilObject = e190
				set intermediateWSObject = ws2
				mchain = 1
			case else
				mchain = 0
		end select 
		'response.write "m.eteilnr: " & me.eteilnr & "; mchain: " & mchain & "<br>"
		
		
		if mchain = 1 then	
			'zugehöriger Produktionsauftrag des Zwischenprodukts muss abgeschlossen vorliegen
			sql = "SELECT * FROM Produktionsauftraege WHERE usergroup="&usergroup&" AND Tag="&me.day&" AND Periode="&period&" " _
				& "AND Auftragsnr="&me.prodOrderNo&" AND abgeschlossen=Produktionsauftraege.Losgroesse ORDER BY id" 
			Set Recordset7=Server.CreateObject("ADODB.Recordset")	
			Recordset7.Open sql, Connection
			
			If Recordset7.EOF Then
				response.write "Auftrag " & me.prodOrderNo & " kann an AP "&wsObject.nr&" nicht abgearbeitet werden, " _
					& "da die Arbeiten am vorangegangen Arbeitsplatz der Produktionskette noch nicht abgeschlossen wurden.<br><br>"
				'preconditions(0) = 0
				me.abort = 1
			Else
				Do While NOT Recordset7.Eof   
					'get batchsize of intermediate product
					intermediateBatchsize = Recordset7("abgeschlossen")
					
					'response.write "Manufacturing chain for production order no. " & me.prodOrderNo & " found.<br>"
					' response.write "wsObject.nr: " & wsObject.nr & "<br>"
					' response.write "wsObject.t: " & wsObject.t & "<br>"
					' response.write "me.batchsize: " & me.batchsize & "<br>"
					' response.write "(intermediateWSObject.prepTime + intermediateWSObject.prodTime) * me.batchsize: " & (intermediateWSObject.prepTime + intermediateWSObject.prodTime) & "<br>"
					
					
					'wait at workstation until intermediate pieces are finished
					'wenn e140,...,e190 heute schon zuvor gefertigt (siehe db produktionsaufträge) und abgeschlossen UND wenn 
					'aktuelle eteilobject.ws.t < produktionszeit für e140,...e190, dann ws.t=produktionszeit (warten auf Zwischenprodukte)
					if wsObject.t < intermediateWSObject.prepTime + (intermediateWSObject.prodTime * intermediateBatchsize) then
						t = intermediateWSObject.prepTime + (intermediateWSObject.prodTime * intermediateBatchsize)
						if t > wsObject.workingTime*60 then
							t = wsObject.workingTime*60
						end if
						wsObject.setT(t)
						'workload stored as a string: "wsNr,starttime,stoptime,duration,prodOrderNo,batchsize,eteilnr; ..."
						wsObject.waitingFlag = 1
						'wsObject.workload = wsObject.workload & wsObject.nr & ",waiting..., , , , , ;" 
						'response.write "waiting...<br>"
					end if
					Recordset7.MoveNext
				Loop
			End If
			Recordset7.close
			Set Recordset7 = Nothing
		end if
	end function
	
	
	public function manufacture()
		'save current production state (amount of finished items)
		me.finished = me.finished + me.batchsize
		'response.write "me.nr: " & me.prodOrderNo & "; me.finished: " & me.finished & "; me.batchsize: " & me.batchsize & "<br>"
		if me.finished > me.batchsizeRequired then
			Response.write "ERROR in 'prodOrder.manufacture'! Aktuelle Losgröße > angeforderte Losgröße<br>"
		end if
		sqlUpdate = "UPDATE Produktionsauftraege SET abgeschlossen="&me.finished&" WHERE id="&me.id&""
		Connection.execute(sqlUpdate)
	end function
	
	public function writeProdOrder(eteilObject, positionFlag)
		'response.write "positionFlag: " & positionFlag & "<br>"
		if positionFlag = "first" OR positionFlag = "firstLast" then
			'response.write "Aus Produktion:<br><br>"
			response.write "<table>"
			response.write "<tr>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp AuftragsNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Zugang &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp noch zu fertigen &nbsp&nbsp</th>"
			response.write "</tr>"
			
			if positionFlag = "first" then
				positionFlag = ""
			end if
		end if
		
		response.write "<tr>"
		'auftragsnr
		response.write "<td align='center'>" & me.prodOrderNo & "</td>"
		'TeilNr
		response.write "<td align='center'>" & eteilObject.nr & "</td>"
		'Bezeichnung
		response.write "<td align='center'>" & eteilObject.name & "</td>"
		' batchsize
		response.write "<td align='center'>" & me.finished & "</td>"
		' noch zu fertigen
		response.write "<td align='center'>" & me.batchsizeRequired - me.finished & "</td>"	
		response.write "</tr>"
		
		if positionFlag = "last" OR positionFlag = "firstLast" then
			response.write "</table>"
			response.write "<br><br>"
		end if
	end function
	
	public function writeUnfinishedOrdersWeek()
		'get orders that could not be finished in this week
		sql = "SELECT * FROM Produktionsauftraege WHERE usergroup="&usergroup&" AND Periode="&period&" AND abgeschlossen < Losgroesse ORDER BY abgeschlossen DESC, Auftragsnr, ETeilnr DESC" 
		Set Recordset8=Server.CreateObject("ADODB.Recordset")	
		Recordset8.Open sql, Connection
		If Recordset8.EOF Then
			Response.Write "Alle Produktionsaufträge abgearbeitet!<br> Keine Fehlteile!<br>"
		Else
			Do While NOT Recordset8.Eof  
				with me
					.id = CInt(Recordset8("id"))
					.prodOrderNo = CInt(Recordset8("Auftragsnr"))
					.eteilnr = CInt(Recordset8("ETeilnr"))
					.day = CInt(Recordset8("Tag"))
					.batchsizeRequired = CInt(Recordset8("Losgroesse"))
					.finished = CInt(Recordset8("abgeschlossen")) 
					.abort = 0 
					.setBatchsize()
				end with
				
				response.write "Folgende Produktionsaufträge sind noch nicht abgeschlossen und müssen in der folgenden Periode erneut beauftragt werden:<br><br>"
				response.write "Auftragsnr.: " & me.prodOrderNo & "<br>"
				response.write "E-Teilnr.: " & me.eteilnr & "<br>"
				response.write "Losgröße: " & me.batchsizeRequired & "<br>"
				response.write "abgeschlossen: " & me.finished & "<br>"
				response.write "noch offen: " & me.batchsizeRequired - me.finished & "<br><br>"
				
				'fetch next entry
				Recordset8.MoveNext
			Loop		
		end if	
		Recordset8.close
		set Recordset8 = Nothing
	end function
	
	
end class

'*************************************************************************************************

class PurchaseOrder
	'properties
	public id 'purchase order id in database
	public purchaseOrderNo 'purchase order number
	public kteilnr 'number of the kteil that shall be manfactured
	public amount 'amount of the kteil that is ordered
	public delivered 'amount of already delivered pieces; this field is != 0 when not all pieces where previously delivered (feature not yet implemented, therefore always 0
	
	'methods
	public function deliver()
		me.delivered = me.amount
	
		sqlUpdate = "UPDATE Kaufteilbestellungen SET geliefert = "&me.amount&" WHERE id="&me.id&""  
		Connection.execute(sqlUpdate)
		
		me.amount = 0
	end function
	
	public function writePurchases(kteilObject, positionFlag)
		if positionFlag = "first" OR positionFlag = "firstLast" then
			response.write "<table>"
			response.write "<tr>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Zugang &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "</tr>"
			
			positionFlag = ""
		end if
		
		response.write "<tr>"
		response.write "<td align='center'>" & kteilObject.nr & "</td>"
		response.write "<td align='center'>" & kteilObject.name & "</td>"
		response.write "<td align='center'>" & me.delivered & "</td>"
		response.write "</tr>"
		
		if positionFlag = "last" OR positionFlag = "firstLast" then
			response.write "</table>"
			response.write "<br><br>"
		end if
	end function
end class
	
'**********************************************************************************************
	
class WS
	'************
	'properties
	'**************
	'object constants (initiated in sub initWS(ws1,ws2,ws3,ws4,ws5,ws6,ws7))
	public name 'workstation's name/task
	public nr
	public prepTime 'preparation time
	public prodTime 'production time for one piece
	public wage 'wage/h
	
	'object variables (daily reset)
	public t 'current operating time this day [min]
	public workload 'stores an 2 dim array; workload[day(1-5)]["starttime;stoptime;prodOrderNo;batchsize;eteilnr"]
	public workingTime 'max. working time this day [h]
	public prepAmount 'amount of workstation-preparations
	public waitingFlag 'flag signalising whether the workstation was waiting (for other items being finished first) before starting its production
	
	'object variables (weekly reset)
	public workingTimeTotal 'total max. working time / week [h]
	public tTotal			'total operating time [min]
	public prepAmountTotal	'total amount of workstation-preparations
	
	'***************
	'methods
	'**************
	public function reset()
		me.t = 0
		me.workload = ""
		me.workingTime = 0
		me.prepAmount = 0
		me.waitingFlag = 0
	end function
	
	public function setT(t)
		if t > me.workingTime*60 then
			response.write "ERROR in WS.setT()! me.t > me.workingTime, setze me.t = me.workingTime<br>"
			me.t = me.workingTime
		else 
			me.t = t
		end if
	end function
	
	public function setTotals()
		'response.write "me.workingTime: " & me.workingTime & "<br>"
		me.tTotal = me.tTotal + me.t '[min]
		me.workingTimeTotal = me.workingTimeTotal + me.workingTime '[h]
		me.prepAmountTotal = me.prepAmountTotal + me.prepAmount
	end function
		
	public function updateWStime()
		t = me.t + me.preptime + (pOrder.batchsize * me.prodTime)
		me.setT(t)
		me.prepAmount = me.prepAmount + 1
	end function
	
	public function loadWorkingTime(day)
		'response.write "loading workingtime for WS " & me.nr & "<br>"
		sql = "SELECT * FROM Arbeitszeiten WHERE usergroup = "&usergroup&" AND Periode = "&period&" AND Tag = "&day&" AND ArbeitsplatzNr = "&me.nr&" ORDER BY id"
		Set Recordset2=Server.CreateObject("ADODB.Recordset")	
		Recordset2.Open sql, Connection
		
		if Recordset2.EOF then
			'response.write "Keine Arbeitszeit für Arbeitsplatz " &me.nr& " angegeben, es werden 8h gesetzt.<br>"
			me.workingTime = 8
			sqlInsert = "INSERT INTO Arbeitszeiten (usergroup, Periode, Tag, ArbeitsplatzNr, Arbeitszeit) " _
				& "VALUES ("&usergroup&", "&period&", "&d&", "&me.nr&", 8)"
			Connection.execute(sqlInsert)
		else
			me.workingTime = Recordset2("Arbeitszeit")
		end if
		
		Recordset2.close
		set Recordset2=Nothing
	end function
	
	public function checkAvailability(pOrder)
		'check if there is sufficient working time left for the desired batchsize, otherwise reduce the batchsize
		batchsizeOld = pOrder.batchsize
		'response.write "ws.nr: " & me.nr & "; me.t: " & me.t & "; me.workingTime: " & me.workingTime & "<br>"
		'response.write "pOrder.prodOrderNo: " & porder.prodOrderNo & "; pOrder.eteilnr: " & pOrder.eteilnr & "; pOrder.batchsize: " & pOrder.batchsize & "; pOrder.finished: " & pOrder.finished & "<br>"
		for j=pOrder.batchsize to 1 step -1
			'response.write "j: " & j & "<br>"
			if (me.t + me.preptime + (j * me.prodTime)) <= (me.workingTime * 60) then
				'set new batchsize
				pOrder.batchsize = j
				
				if pOrder.batchsize = batchsizeOld then
					'response.write "fall 1"
					' response.write "Es können alle gefordeten Teile im Los (" & batchsizeRequired & " Stück) in der verbleibenden Arbeitszeit an AP " &me.nr& " von "_
					' & (8*60 - me.t) & " Minuten produziert werden.<br>"
					exit for
				else
					'response.write "fall 2"
					 response.write "<u>Auftrag " &pOrder.prodOrderNo& ":</u><br> Teilnr. " &pOrder.eteilnr& ", gewünschte Stückzahl: " _
						& batchsizeOld & ". ACHTUNG: Verbleibende Arbeitszeit an AP " &me.nr& " von " _
						& (me.workingTime*60 - me.t) & " Minuten genügt nur für " & pOrder.batchsize & " Teile. Der restliche Auftrag wird am nächsten Arbeitstag fortgesetzt.<br><br>"
					exit for
				end if
				
			else
				'response.write "fall 3"
				if j=1 then
					pOrder.abort = 1
					response.write "<u>Auftrag " &pOrder.prodOrderNo& ":</u> <br>Teilnr. " &pOrder.eteilnr& ", gewünschte Stückzahl: " _
						& pOrder.batchsizeRequired & ". ACHTUNG: Verbleibende Arbeitszeit an AP " & me.nr & " von " & (me.workingTime*60 - me.t) _
						& " Minuten genügt nicht, um ein Teil dieses Loses zu fertigen. " & "Der Auftrag wird am nächsten Arbeitstag aufgenommen.<br><br>"
				elseif j<1 then 
					response.write "ERROR. Unzulässiger Variablenwert für j in WS.checkAvailability().<br>"
				end if
			end if
		next
		' if batchsizeOld < batchsize then
			' Response.write "Der erste (vorherige) Arbeitsplatz in der Arbeitsplatzkette beschränkt die Losgrösse weiter.<br>"
			' batchsize = batchsizeOld
		' end if
	end function
	
	
	public function setWorkload(pOrder)
		'workload stored as a string: "wsNr,starttime,stoptime,duration,prodOrderNo,batchsize,eteilnr; ..."
		workloadString = me.workload & me.nr & "," & me.t & "," 
		starttime = me.t
		me.updateWStime()
		workloadString = workloadString & me.t & "," & me.t-starttime & "," & pOrder.prodOrderNo & "," & pOrder.batchsize & "," & pOrder.eteilnr & ";"
		me.workload = workloadString
	end function
	
	public function writeWorkload()
		if me.nr = 1 then
			'write table header
			response.write "<h4>Maschinenbelegung</h4>"
			response.write "<table>"
			response.write "<tr>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp ArbPlatz &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Beginn &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Ende &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Dauer &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp AuftragsNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp Anzahl gefertigte Teile &nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp E-TeilNr &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "</tr>"	
		end if
		
		if me.waitingFlag = 1 then
			response.write "<tr><td align='center'>" & me.nr & "</td><td align='center'>warte...</td></tr>"
		end if
		
		'split workload string "wsNr,starttime,stoptime,duration,prodOrderNo,batchsize,eteilnr;..."
		wl = split(me.workload,";")
		for each i in wl
			wlFields = split(i,",")
			count = 0
			for each k in wlFields
				'[workstation number, starttime, stoptime, duration, batchsize, eteilnr]
				if count = 0 then
					'workstation number
					response.write "<tr>"
					response.write "<td align='center'>" & k & "</td>"
					count = count + 1
				elseif count = 6 then
					'eteilnr
					response.write "<td align='center'>" & k & "</td>"
					response.write  "</tr>"
					count = 0
				else 
					response.write "<td align='center'>" & k & "</td>"
					count = count + 1
				end if
				
			next
			separator = 1
		next
		
		if separator = 1 then
			response.write "<tr>"
			response.write "<td align='center' colspan=7>------------------------------------------------------------------------------------------------------------------------------------------------------------------------</td>"
			response.write "<tr>"
			separator=0
		else
			response.write "<tr>"
			response.write "<td align='center'>" & me.nr & "</td>"
			response.write "<td align='center'>-</td>"
			response.write "</tr>"
			response.write "<tr>"
			response.write "<td align='center' colspan=7>------------------------------------------------------------------------------------------------------------------------------------------------------------------------</td>"
			response.write "<tr>"
		end if
		
		if me.nr = 7 then	
			response.write "</table>"
			response.write "<br><br>"
		end if
	end function
	
	
	public function writeSummary()
		'response.write "positionFlag: " & positionFlag & "<br>"
		if me.nr = 1 then
			'response.write "Aus Produktion:<br><br>"
			response.write "<h3>ARBEITSPLÄTZE</h3>"
			response.write "<h4>Zusammenfassung</h4>"
			response.write "<table>"
			response.write "<tr>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp ArbPlatz &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Leermin &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Auslastung &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Rüstvorgänge &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th> &nbsp&nbsp Rüstminuten &nbsp&nbsp</th>"
			response.write "</tr>"
		end if
		
		response.write "<tr>"
		'ArbPlatz
		response.write "<td align='center'>" & me.nr & "</td>"
		'Leermin
		response.write "<td align='center'>" & me.workingTime*60 - me.t & "</td>"
		'Auslastung[%] = (Fertigungszeit + Rüstzeit) / Gesamtarbeitszeit * 100
		if me.workingTime <> 0 then 
			efficiency = 100 * (me.t + me.prepAmount*me.prepTime)/(me.workingTime*60)
		else 
			efficiency = 0 
		end if
		'FormatNumber(Zahl, Nachkommastellen, fuehrendeNull,KlammernFürNegativeZahlen, Zahlengruppen) ...
		'mit -1 == "Parameter gesetzt", 0=="Parameter nicht setzen", -2 == "Parameter wie in Ländereinstellungen des Computers"
		response.write "<td align='center'>" & FormatNumber(efficiency,2,-1,0,-1) & " %</td>"
		' Rüstvorgänge
		response.write "<td align='center'>" & me.prepAmount & "</td>"
		' Rüstminuten
		response.write "<td align='center'>" & me.prepAmount*me.prepTime & "</td>"	
		response.write "</tr>"
		
		if me.nr = 7 then
			response.write "</table>"
			response.write "<br><br>"
		end if
	end function
	
	
	public function writeTotalSummary()
		if me.nr = 1 then
			response.write "<h3>PRODUKTIONSSTAND</h3>"
			'response.write "<h4>Zusammenfassung</h4>"
			response.write "<table>"
			response.write "<tr>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp ArbPlatz &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp VorgabeMin &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp LeerMin &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp RüstMin &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp FertMin &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Produktivität &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "<th>&nbsp&nbsp&nbsp&nbsp Auslastung &nbsp&nbsp&nbsp&nbsp</th>"
			response.write "</tr>"
		end if
		
		response.write "<tr>"
		'ArbPlatz
		response.write "<td align='center'>" & me.nr & "</td>"
		'VorgabeMin
		response.write "<td align='center'>" & me.workingTimeTotal*60 & "</td>"
		'LeerMin
		response.write "<td align='center'>" & me.workingTimeTotal*60 - me.tTotal & "</td>"
		' Rüstminuten
		response.write "<td align='center'>" & me.prepAmountTotal*me.prepTime & "</td>"	
		'Fertigungsminuten
		response.write "<td align='center'>" & me.tTotal & "</td>"	
		'Produktivität[%] = Fertigungszeit/Gesamtarbeitszeit * 100
		if me.workingTimeTotal <> 0 then 
			productivity = 100*(me.tTotal)/(me.workingTimeTotal*60)
		else 
			productivity = 0 
		end if
		response.write "<td align='center'>" & FormatNumber(productivity,2,-1,0,-1) & " %</td>"
		'Auslastung[%] = (Fertigungszeit + Rüstzeit) / Gesamtarbeitszeit * 100
		if me.workingTimeTotal <> 0 then 
			efficiency = 100*(me.tTotal + me.prepAmountTotal * me.prepTime)/(me.workingTimeTotal*60)
		else 
			efficiency = 0 
		end if
		'FormatNumber(Zahl, Nachkommastellen, fuehrendeNull,KlammernFürNegativeZahlen, Zahlengruppen) ...
		'mit -1 == "Parameter gesetzt", 0=="Parameter nicht setzen", -2 == "Parameter wie in Ländereinstellungen des Computers"
		response.write "<td align='center'>" & FormatNumber(efficiency,2,-1,0,-1) & " %</td>"
		
		response.write "</tr>"
		
		if me.nr = 7 then
			for each i in wsArray
				VorgabeMinGes = VorgabeMin + i.workingTimeTotal*60
				LeerMinGes = LeerMinGes + i.tTotal
				RuestMinGes = RuestMinGes + i.prepAmountTotal*i.prepTime
				FertMinGes = FertMinGes + i.tTotal
				if i.workingTimeTotal <> 0 then
					produktivitaet = 100*(i.tTotal)/(i.workingTimeTotal*60)
				else
					produktivitaet = 0
				end if
				ProduktivitaetGes = ProduktivitaetGes + produktivitaet
				
				if i.workingTimeTotal <> 0 then
					auslastung = 100*(i.tTotal + i.prepAmountTotal * i.prepTime)/(i.workingTimeTotal*60)
				else
					auslastung = 0
				end if
				AuslastungGes = AuslastungGes + auslastung
			next
			response.write "<tr><td align='center' colspan=7>===========================================================================================</td></tr>"
			response.write "<tr>"
			'ArbPlatz
			response.write "<td align='center'>Gesamt: </td>"
			'VorgabeMin
			response.write "<td align='center'>" & VorgabeMinGes & "</td>"
			'LeerMin
			response.write "<td align='center'>" & LeerMinGes & "</td>"
			' Rüstminuten
			response.write "<td align='center'>" & RuestMinGes & "</td>"	
			'Fertigungsminuten
			response.write "<td align='center'>" & FertMinGes & "</td>"	
			'Produktivität[%] = Fertigungszeit/Gesamtarbeitszeit * 100
			response.write "<td align='center'>" & FormatNumber(ProduktivitaetGes/7,2,-1,0,-1) & " %</td>"	
			'Auslastung[%] = (Fertigungszeit + Rüstzeit) / Gesamtarbeitszeit * 100
			response.write "<td align='center'>" & FormatNumber(AuslastungGes/7,2,-1,0,-1) & " %</td>"	
			response.write "</tr>"							
			response.write "<tr><td align='center' colspan=7>--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</td></tr>"
			response.write "</table>"
			response.write "<br><br>"
		end if
	end function 
	
end class

'*******************************************************************************************************

class Output
	public inputs
	public results
	public stock
	public dtg 'day-time-group
	
	'methods
	public function getInputs()
		'sql = "SELECT * FROM 
	end function
		 
	public function echo(str)
		me.results = me.results & str
	end function
	
	public function save()
		sqlInsert = "INSERT INTO Ergebnisse (usergroup, Periode, Zeitstempel, Lagerbestand, Eingabe, Ergebnis) " _
			& "VALUES ("&usergroup&", "&period&", "&me.dtg&", "&me.stock&", "&me.inputs&", "&me.results&")" 
		Connection.execute(sqlInsert)
	end function
	
end class


'********************************************************************

class Stock
	public holdingETeil
	
	'Arrays (array range [0-48])
	public holdingOld(49)  'holds the initial stock holding for later statistics
	public holding(49) 'represents the the current stock state, the current DB amount of each item is saved in the array, item.nr-1 == index
	public input(49) 'saves the amount of incoming items (used for statistics only, not used as state)
	public dispatch(49) 'saves the amount of outgoing items (used for statistics only, not used as state)
	public undone(49)
	public storageArea(49) 
	public bilValue(49) 
	public stockValue(49)
	
	'remember: regular item numbers range from 1 to 44 [i.e. array[0-43]), the remaining 5 slots in the array are used for 
	'the intermediate items e140, e150, e160, e180, e190
	
	
	'methods
	public function init()
		for i=0 to 48
			with me
				.input(i) = 0
				.dispatch(i) = 0
				.undone(i) = ""
			end with
		next
	end function
	
	function setDispatch(itemObject, amount)
		'counting the dispatch amount of items in the attribute item.dispatch ([array], itemnr-1 == index)
		'amount contains the used respectively shipped (for item No.1 "Lopez" and 2 "Komb. Arbeiten 1") 
		'amount of itemObject 
		if itemObject.nr=140 OR  itemObject.nr=150 OR itemObject.nr=160 OR itemObject.nr=180 OR itemObject.nr=190 then
			select case itemObject.nr
				case 140
					index = 44
				case 150
					index = 45
				case 160
					index = 46
				case 180		
					index = 47
				case 190
					index = 48
			end select
			me.dispatch(index) = me.dispatch(index) + amount
		else
			'set amount of incoming item
			me.dispatch(itemObject.nr-1) = me.dispatch(itemObject.nr-1) + amount
		end if
		'me.dispatch(itemObject.nr-1) = me.dispatch(itemObject.nr-1) + amount
	end function
	
	function setInput(itemObject, amount)
		'counting the incoming amount of items in the attribute item.input ([array], itemnr-1 == index)
		if itemObject.nr=140 OR  itemObject.nr=150 OR itemObject.nr=160 OR itemObject.nr=180 OR itemObject.nr=190 then
			select case itemObject.nr
				case 140
					index = 44
				case 150
					index = 45
				case 160
					index = 46
				case 180		
					index = 47
				case 190
					index = 48
			end select
			me.input(index) = me.input(index) + amount
		else
			'set amount of incoming item
			me.input(itemObject.nr-1) = me.input(itemObject.nr-1) + amount
		end if
	end function
	
	' public function updateHolding(item)
		'**********************************
		'save current stock holding to DB
		'**********************************
		'********
		'iterate through inputItemsAmount and substract the pieces from stock
		'for i=0 to UBound(eteilObject.inputItems)
			'item = eteilObject.inputItems(i)
			'itemAmount = eteilObject.inputItemsAmount(i)
			'bestand = me.holding(item.nr-1)-(pOrder.batchsize * itemAmount)
		'*********
		
		' if item.algebraicSign = "+" then
			' bestand = me.holding(item.nr-1)+(item.amount)
			' me.holding(item.nr-1) = bestand
		' elseif item.algebraicSign = "-" then
			' bestand = me.holding(item.nr-1)-(item.amount)
			' if bestand < 0 then
				' bestand = 0
				' response.write "ERROR in Stock.updateHolding(item)! Bestand < 0, dies sollte nicht sein! Bestand wird auf 0 gesetzt.<br>"
			' end if
			' me.holding(item.nr-1) = bestand
		' end if
		
		
		'response.write "item: " &item& "; itemAmount: " &itemAmount& "<br>"
		'response.write "holding(i): " &holding(i)& "; pOrder.batchsize: " &pOrder.batchsize& "; itemAmount: " &itemAmount& "; bestand: " &bestand& "<br>"
		'subtract itemAmount of input items for E-Teil from stock

		'sqlUpdate = "UPDATE Lager SET Lagerbestand="&bestand&" WHERE usergroup="&usergroup&" AND Periode="&period&" AND Teilnr="&item.nr&""
		'Connection.execute(sqlUpdate)
		'next
	' end function
	
	public function getHolding()
		'**********************************************************************************
		'gets the current stock holding from the DB. Used for statistical output only, 
		'the states used for calculations are handled by the objects of ETeil and KTeil
		'**********************************************************************************
		sql = "SELECT * FROM Lager WHERE usergroup="&usergroup&" "'AND Periode="&period&" "	
		Set Recordset4=Server.CreateObject("ADODB.Recordset")	
		Recordset4.Open sql, Connection
		If Recordset4.EOF Then
			Response.Write "Kein Lager für Nutzergruppe " &usergroup& " eingerichtet.<br>"
		Else
			count=0
			Do While NOT Recordset4.Eof
				if count>48 then
					response.write "ERROR. Resetting the counter in Stock.getHolding().<br>"
					count=0
				end if
				me.holding(count) = Recordset4("Lagerbestand")
				count=count+1
				
				'teilnr = Recordset4("Teilnr")
				'response.write "count: " & count & "; teilnr: " & teilnr & "; me.holdingOld(count): " & me.holdingOld(count) & "<br>"
				
				'fetch next entry
				Recordset4.MoveNext()
			Loop
		end if
		Recordset4.close
		set Recordset4 = Nothing
	end function
	
	public function backupHolding()
		'backups the current (initial) holding in stock.holdingOld array for later statistics
		sql = "SELECT * FROM Lager WHERE usergroup="&usergroup&" "'AND Periode="&period&" "	
		Set Recordset4=Server.CreateObject("ADODB.Recordset")	
		Recordset4.Open sql, Connection
		If Recordset4.EOF Then
			Response.Write "Kein Lager für Nutzergruppe " &usergroup& " eingerichtet.<br>"
		Else
			count=0
			Do While NOT Recordset4.Eof
				if count>48 then
					response.write "ERROR. Resetting the counter in Stock.backupHolding().<br>"
					count=0
				end if
				me.holdingOld(count) = Recordset4("Lagerbestand")
				count=count+1
				
				'teilnr = Recordset4("Teilnr")
				'response.write "count: " & count & "; teilnr: " & teilnr & "; me.holdingOld(count): " & me.holdingOld(count) & "<br>"
				
				'fetch next entry
				Recordset4.MoveNext()
			Loop
		end if
		Recordset4.close
		set Recordset4 = Nothing
	end function
	
	public function writeHolding()
		response.write "<h3>LAGERBESTAND</h3>"
		'response.write "<h4>Zusammenfassung</h4>"
		response.write "<table>"
		response.write "<tr>"
		response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Nr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp AltBest &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th>&nbsp&nbsp&nbsp&nbsp Abgang &nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Zugang &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Offen &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Bestand &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Lagerfläche &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"	
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Bil.Wert &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"	
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Lagerwert &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"	
		response.write "</tr>"
	
			
		for i = 0 to 43
			response.write "<tr>"
			'Nr
			response.write "<td align='center'>" & i+1 & "</td>"
			'AltBest
			response.write "<td align='center'>" & me.holdingOld(i) & "</td>"
			'Abgang
			response.write "<td align='center'>" & me.dispatch(i) & "</td>"
			'Zugang
			response.write "<td align='center'>" & me.input(i) & "</td>"
			'Offen
			response.write "<td align='center'>" & me.undone(i) & "</td>"	
			'Bestand
			response.write "<td align='center'>" & me.holding(i) & "</td>"	
			'Lagerfläche
			response.write "<td align='center'>" & me.storageArea(i) & "</td>"	
			'Bil.Wert
			response.write "<td align='center'>" & me.bilValue(i) & "</td>"	
			'Lagerwert
			response.write "<td align='center'>" & me.stockValue(i) & "</td>"	
			response.write "</tr>"
		next		
	
		response.write "</table>"
		response.write "<br><br>"

	end function
end class
	
class Finances
	'methods
	public function writeProfitAndLossStatement(stockObj)
		response.write "<h3>ERGEBNISRECHNUNG</h3>"
		'response.write "<h4>Zusammenfassung</h4>"
		response.write "<table>"
		response.write "<tr>"
		response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp Vertriebsteil &nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Wunsch &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th>&nbsp&nbsp&nbsp&nbsp verkauft &nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Preis &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Umsatz &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
		response.write "</tr>"
	
		'iterating through eteil 1 and 2 (Lopez & Arbeiten 1)	
		for i = 1 to 2
			response.write "<tr>"
			'Vertriebsteil
			call setItemObject(i, itemObject)
			response.write "<td align='center'>" & itemObject.name & "</td>"
			'Wunsch
			response.write "<td align='center'>" & "Wunsch" & "</td>"
			'verkauft
			response.write "<td align='center'>" & stockObj.input(i-1) & "</td>"
			'Preis
			response.write "<td align='center'>" & itemObject.usp & "</td>"
			'Umsatz
			response.write "<td align='center'>" & itemObject.usp*stockObj.input(i-1) & "</td>"	

			response.write "</tr>"
		next	
	
	end function

end class	
	
	
	
'************************************
'			"constructors"/
'		object initialisations
'************************************

'***************************
'create ETeile
'***************************
sub initETeile(e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e140,e150,e160,e180,e190)
	Set e1 = New ETeil
	with e1
		.name = "Lopez Arbeitsplatz"
		.nr = 1
		.stockCoeff = 0
		.ws = 5
		.cost = 2178.94
		.usp = 2600
		.inputItems = Array(3,4,5,25,27,28,29,39,40)
		.inputItemsAmount = Array(1,2,1,6,1,2,12,3,4)
	end with
	
	Set e2 = New ETeil
	with e2
		.name = "Kombination Arbeiten 1"
		.nr = 2
		.stockCoeff = 0
		.ws = 5
		.cost = 2620.85
		.usp = 2500
		.inputItems = Array(12,13,36,37,38,39,40)
		.inputItemsAmount = Array(2,1,4,8,16,2,6)
	end with
	
	Set e3 = New ETeil
	with e3
		.name = "Allroundplatte 90°"
		.nr = 3
		.stockCoeff = 0
		.ws = 4
		.cost = 425.87
		.usp = 0
		.inputItems = Array(6,7,26)
		.inputItemsAmount = Array(1,1,4)
	end with
	
	Set e4 = New ETeil
	with e4
		.name = "Besprechungsplatte 120°"
		.nr = 4
		.stockCoeff = 0
		.ws = 4
		.cost = 351.29
		.usp = 0
		.inputItems = Array(8,9,26)
		.inputItemsAmount = Array(1,1,4)
	end with
	
	Set e5 = New ETeil
	with e5
		.name = "Ergänzungsplatte"
		.nr = 5
		.stockCoeff = 0
		.ws = 4
		.cost = 356
		.usp = 0
		.inputItems = Array(10,11,26)
		.inputItemsAmount = Array(1,1,5)
	end with
	
	Set e6 = New ETeil
	with e6
		.name = "A-Platte"
		.nr = 6
		.stockCoeff = 0
		.ws = 1
		.cost = 232.19
		.usp = 0
		.inputItems = Array(20)
		.inputItemsAmount = Array(3)
	end with
	
	Set e7 = New ETeil
	with e7
		.name = "A-Rahmen"
		.nr = 7
		.stockCoeff = 0
		.ws = 3
		.cost = 101.22
		.usp = 0
		.inputItems = Array(21,22,23,25)
		.inputItemsAmount = Array(1,1,2,4)
	end with
	
	Set e8 = New ETeil
	with e8
		.name = "B-Platte"
		.nr = 8
		.stockCoeff = 0
		.ws = 1
		.cost = 162.1
		.usp = 0
		.inputItems = Array(20)
		.inputItemsAmount = Array(2)
	end with
	
	Set e9 = New ETeil
	with e9
		.name = "B-Rahmen"
		.nr = 9
		.stockCoeff = 0
		.ws = 3
		.cost = 96.72
		.usp = 0
		.inputItems = Array(24,23,25)
		.inputItemsAmount = Array(2,2,4)
	end with
	
	Set e10 = New ETeil
	with e10
		.name = "E-Platte"
		.nr = 10
		.stockCoeff = 0
		.ws = 1
		.cost = 162.1
		.usp = 0
		.inputItems = Array(20)
		.inputItemsAmount = Array(2)
	end with


	Set e11 = New ETeil
	with e11
		.name = "E-Rahmen"
		.nr = 11
		.stockCoeff = 0
		.ws = 3
		.cost = 100.98
		.usp = 0
		.inputItems = Array(21,23,25)
		.inputItemsAmount = Array(1,3,5)
	end with
	
	Set e12 = New ETeil
	with e12
		.name = "Container (klein)"
		.nr = 12
		.stockCoeff = 0
		.ws = 4
		.cost = 672.87
		.usp = 0
		.inputItems = Array(14,15,16,17,33,34,38,41,42,43)
		.inputItemsAmount = Array(3,4,1,2,1,2,12,1,1,2)
	end with
	
	Set e13 = New ETeil
	with e13
		.name = "Container (gross)"
		.nr = 13
		.stockCoeff = 0
		.ws = 4
		.cost = 1051.19
		.usp = 0
		.inputItems = Array(18,19,17,33,38,41,42,43,44)
		.inputItemsAmount = Array(4,4,2,2,6,1,1,4,1)
	end with
	
	Set e14 = New ETeil
	with e14
		.name = "Seitenteil 100"
		.nr = 14
		.stockCoeff = 0
		.ws = 6
		.cost = 78.63
		.usp = 0
		.inputItems = Array(140,31)
		.inputItemsAmount = Array(1,1)
	end with
	
	Set e140 = New ETeil
	with e140
		.name = "Seitenteil 100 Zwischenprodukt"
		.nr = 140
		.stockCoeff = 0
		.ws = 1
		.cost = 0
		.usp = 0
		.inputItems = Array(30)
		.inputItemsAmount = Array(1)
	end with
	
	Set e15 = New ETeil
	with e15
		.name = "Rundstollen 100"
		.nr = 15
		.stockCoeff = 0
		.ws = 7
		.cost = 36.12
		.usp = 0
		.inputItems = Array(150,32)
		.inputItemsAmount = Array(1,2)
	end with
	
	Set e150 = New ETeil
	with e150
		.name = "Rundstollen 100 Zwischenprodukt"
		.nr = 150
		.stockCoeff = 0
		.ws = 2
		.cost = 0
		.usp = 0
		.inputItems = Array(32)
		.inputItemsAmount = Array(2)
	end with
	
	Set e16 = New ETeil
	with e16
		.name = "Tür"
		.nr = 16
		.stockCoeff = 0
		.ws = 6
		.cost = 78.63
		.usp = 0
		.inputItems = Array(160,31)
		.inputItemsAmount = Array(1,1)
	end with	
	
	Set e160 = New ETeil
	with e160
		.name = "Tür Zwischenprodukt"
		.nr = 160
		.stockCoeff = 0
		.ws = 1
		.cost = 0
		.usp = 0
		.inputItems = Array(30)
		.inputItemsAmount = Array(1)
	end with
	
	Set e17 = New ETeil
	with e17
		.name = "Rundstollen 40"
		.nr = 17
		.stockCoeff = 0
		.ws = 2
		.cost = 19.77
		.usp = 0
		.inputItems = Array(32)
		.inputItemsAmount = Array(1)
	end with
	
	Set e18 = New ETeil
	with e18
		.name = "Seitenteil 210"
		.nr = 18
		.stockCoeff = 0
		.ws = 6
		.cost = 124.72
		.usp = 0
		.inputItems = Array(180,31)
		.inputItemsAmount = Array(1,2)
	end with	
	
	'Zwischenprodukt von AP1
	Set e180 = New ETeil
	with e180
		.name = "Seitenteil 210 Zwischenprodukt"
		.nr = 180
		.stockCoeff = 0
		.ws = 1
		.cost = 0
		.inputItems = Array(30)
		.inputItemsAmount = Array(2)
	end with
	
	Set e19 = New ETeil
	with e19
		.name = "Rundstollen 210"
		.nr = 19
		.stockCoeff = 0
		.ws = 7
		.cost = 47.24
		.usp = 0
		.inputItems = Array(190)
		.inputItemsAmount = Array(1)
	end with
	
	Set e190 = New ETeil
	with e190
		.name = "Rundstollen 210 Zwischenprodukt"
		.nr = 190
		.stockCoeff = 0
		.ws = 2
		.cost = 0
		.usp = 0
		.inputItems = Array(32)
		.inputItemsAmount = Array(3)
	end with

end sub


'******************************************
' KTeile
'******************************************
sub initKTeile(k20,k21,k22,k23,k24,k25,k26,k27,k28,k29,k30,k31,k32,k33,k34,k35,k36,k37,k38,k39,k40,k41,k42,k43,k44)
	Set k20 = New KTeil
	with k20
		.name = "Platte Lopez"
		.nr = 20
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 64.9
	end with
	
	Set k21 = New KTeil
	with k21
		.name = "Rahmen Innenbogen"
		.nr = 21
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 10.86
	end with
	
	Set k22 = New KTeil
	with k22
		.name = "Rahmen Aussenbogen"
		.nr = 22
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 15.86
	end with
	
	Set k23 = New KTeil
	with k23
		.name = "Rahmen 80 gerade"
		.nr = 23
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 8.72
	end with
	
	Set k24 = New KTeil
	with k24
		.name = "Rahmen 120 gerade"
		.nr = 24
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 11.82
	end with
	
	Set k25 = New KTeil
	with k25
		.name = "Tischverbinder"
		.nr = 25
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 6.92
	end with
	
	Set k26 = New KTeil
	with k26
		.name = "Verbindunszapfen mit Federring"
		.nr = 26
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 0.42
	end with
	
	Set k27 = New KTeil
	with k27
		.name = "Sichtblende 90° mit Beschlag"
		.nr = 27
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 32.1
	end with
	
	Set k28 = New KTeil
	with k28
		.name = "Sichtblende 120° mit Beschlag"
		.nr = 28
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 34.28
	end with
	
	Set k29 = New KTeil
	with k29
		.name = "Tischbein"
		.nr = 29
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 48.56
	end with
	
	Set k30 = New KTeil
	with k30
		.name = "Platte Arbeiten 1"
		.nr = 30
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 40.23
	end with
	
	Set k31 = New KTeil
	with k31
		.name = "Lack Alu"
		.nr = 31
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 2.45
	end with
	
	Set k32 = New KTeil
	with k32
		.name = "Rundstollen"
		.nr = 32
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 10.3
	end with
	
	Set k33 = New KTeil
	with k33
		.name = "Heisskleber"
		.nr = 33
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 1.8
	end with
	
	Set k34 = New KTeil
	with k34
		.name = "Türbeschläge"
		.nr = 34
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 3.18
	end with
	
	Set k35 = New KTeil
	with k35
		.name = "Möbelschrauben M4x15"
		.nr = 35
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 0.07
	end with
	
	Set k36 = New KTeil
	with k36
		.name = "Glasboden"
		.nr = 36
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 25.73
	end with
	
	Set k37 = New KTeil
	with k37
		.name = "Traverse"
		.nr = 37
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 6.78
	end with
	
	Set k38 = New KTeil
	with k38
		.name = "Schrauben"
		.nr = 38
		.stockCoeff = 0
		.deliveryTime = 2 
		.cost = 0.07
	end with
	
	Set k39 = New KTeil
	with k39
		.name = "Wellpappe"
		.nr = 39
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 2.2
	end with
	
	Set k40 = New KTeil
	with k40
		.name = "Versandkarton"
		.nr = 40
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 3.5
	end with
	
	Set k41 = New KTeil
	with k41
		.name = "Containersockel"
		.nr = 41
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 43.8
	end with
	
	Set k42 = New KTeil
	with k42
		.name = "Containeroberboden"
		.nr = 42
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 43.8
	end with
	
	Set k43 = New KTeil
	with k43
		.name = "Containereinlegeboden"
		.nr = 43
		.stockCoeff = 0
		.deliveryTime = 3 
		.cost = 19.34
	end with
	
	Set k44 = New KTeil
	with k44
		.name = "Rolladentür"
		.nr = 44
		.stockCoeff = 0
		.deliveryTime = 3
		.cost = 46.87
	end with	

end sub



'***************************
'create workstations
'***************************
sub initWS(ws1,ws2,ws3,ws4,ws5,ws6,ws7)
	'initialise workstations
	Set ws1 = New WS
	with ws1
		.name = "Sägen und Fräsen 1 (Platte)"
		.nr = 1
		.t = 0
		.prepTime = 3
		.prodTime = 4
		.wage = 173.9
		.workload = ""
		.prepAmount = 0
	end with
	
	Set ws2 = New WS
	with ws2
		.name = "Sägen und Fräsen 2 (Rundstollen)"
		.nr = 2
		.t = 0
		.prepTime = 2
		.prodTime = 3
		.wage = 96.12
		.workload = ""
		.prepAmount = 0
	end with
	
	Set ws3 = New WS
	with ws3
		.name = "Kommissionieren und Verpacken"
		.nr = 3
		.t = 0
		.prepTime = 4
		.prodTime = 8
		.wage = 109.39
		.workload = ""
		.prepAmount = 0
	end with
	
	Set ws4 = New WS
	with ws4
		.name = "Montage"
		.nr = 4
		.t = 0
		.prepTime = 10
		.prodTime = 25
		.wage = 143.9
		.workload = ""
		.prepAmount = 0
	end with
		
	Set ws5 = New WS
	with ws5
		.name = "Verpacken und Versenden"
		.nr = 5
		.t = 0
		.prepTime = 5
		.prodTime = 20
		.wage = 56.76
		.workload = ""
		.prepAmount = 0
	end with
		
	Set ws6 = New WS
	with ws6
		.name = "Lackieren und Trocknen"
		.nr = 6
		.t = 0
		.prepTime = 1
		.prodTime = 6
		.wage = 84.29
		.workload = ""
		.prepAmount = 0
	end with
		
	Set ws7 = New WS
	with ws7
		.name = "Bohren"
		.nr = 7
		.t = 0
		.prepTime = 2
		.prodTime = 2
		.wage = 72.53
		.workload = ""
		.prepAmount = 0
	end with
end sub
%>



			