<!--***********************
erpProcedures.asp
*************************-->
<%
sub setItemObject(itemNr, itemObject)
    'response.write "eteilnr: " & eteilnr & "<br>"
    select case itemNr
        case 1
            Set itemObject = e1
        case 2
            Set itemObject = e2
        case 3
            Set itemObject = e3
        case 4
            Set itemObject = e4
        case 5
            Set itemObject = e5
        case 6
            Set itemObject = e6
        case 7
            Set itemObject = e7
        case 8
            Set itemObject = e8
        case 9
            Set itemObject = e9
        case 10
            Set itemObject = e10
        case 11
            Set itemObject = e11
        case 12
            Set itemObject = e12
        case 13
            Set itemObject = e13
        case 14
            Set itemObject = e14
        case 15
            Set itemObject = e15
        case 16
            Set itemObject = e16
        case 17
            Set itemObject = e17
        case 18
            Set itemObject = e18
        case 19
            Set itemObject = e19
        case 140
            Set itemObject = e140
        case 150
            Set itemObject = e150
        case 160
            Set itemObject = e160
        case 180
            Set itemObject = e180
        case 190
            Set itemObject = e190
        case 20
            Set itemObject = k20
        case 21
            Set itemObject = k21
        case 22
            Set itemObject = k22
        case 23
            Set itemObject = k23
        case 24
            Set itemObject = k24
        case 25
            Set itemObject = k25
        case 26
            Set itemObject = k26
        case 27
            Set itemObject = k27
        case 28
            Set itemObject = k28
        case 29
            Set itemObject = k29
        case 30
            Set itemObject = k30
        case 31
            Set itemObject = k31
        case 32
            Set itemObject = k32
        case 33
            Set itemObject = k33
        case 34
            Set itemObject = k34
        case 35
            Set itemObject = k35
        case 36
            Set itemObject = k36
        case 37
            Set itemObject = k37
        case 38
            Set itemObject = k38
        case 39
            Set itemObject = k39
        case 40
            Set itemObject = k40
        case 41
            Set itemObject = k41
        case 42
            Set itemObject = k42
        case 43
            Set itemObject = k43
        case 44
            Set itemObject = k44
            
    end select
    
    'response.write "vartype(eteilObject): " & vartype(eteilObject) & "<br>"
    
    
    'load current holding of item from db
    if vartype(itemObject)<>vbEmpty then
        itemObject.getHolding()    
    end if
end sub

' sub setKTeilObject(kteilnr, kteilObject)
    'response.write "kteilnr: " & kteilnr & "<br>"
    ' select case kteilnr
        ' case 20
            ' Set kteilObject = k20
        ' case 21
            ' Set kteilObject = k21
        ' case 22
            ' Set kteilObject = k22
        ' case 23
            ' Set kteilObject = k23
        ' case 24
            ' Set kteilObject = k24
        ' case 25
            ' Set kteilObject = k25
        ' case 26
            ' Set kteilObject = k26
        ' case 27
            ' Set kteilObject = k27
        ' case 28
            ' Set kteilObject = k28
        ' case 29
            ' Set kteilObject = k29
        ' case 30
            ' Set kteilObject = k30
        ' case 31
            ' Set kteilObject = k31
        ' case 32
            ' Set kteilObject = k32
        ' case 33
            ' Set kteilObject = k33
        ' case 34
            ' Set kteilObject = k34
        ' case 35
            ' Set kteilObject = k35
        ' case 36
            ' Set kteilObject = k36
        ' case 37
            ' Set kteilObject = k37
        ' case 38
            ' Set kteilObject = k38
        ' case 39
            ' Set kteilObject = k39
        ' case 40
            ' Set kteilObject = k40
        ' case 41
            ' Set kteilObject = k41
        ' case 42
            ' Set kteilObject = k42
        ' case 43
            ' Set kteilObject = k43
        ' case 44
            ' Set kteilObject = k44
    ' end select
    
    'response.write "vartype(kteilObject): " & vartype(kteilObject) & "<br>"
    ' if vartype(kteilObject)<>vbEmpty then
        ' kteilObject.getHolding()    
    ' end if
' end sub

'*********************************************************************

sub setETeilObject(eteilnr, eteilObject)
    'response.write "eteilnr: " & eteilnr & "<br>"
    select case eteilnr
        case 1
            Set eteilObject = e1
        case 2
            Set eteilObject = e2
        case 3
            Set eteilObject = e3
        case 4
            Set eteilObject = e4
        case 5
            Set eteilObject = e5
        case 6
            Set eteilObject = e6
        case 7
            Set eteilObject = e7
        case 8
            Set eteilObject = e8
        case 9
            Set eteilObject = e9
        case 10
            Set eteilObject = e10
        case 11
            Set eteilObject = e11
        case 12
            Set eteilObject = e12
        case 13
            Set eteilObject = e13
        case 14
            Set eteilObject = e14
        case 15
            Set eteilObject = e15
        case 16
            Set eteilObject = e16
        case 17
            Set eteilObject = e17
        case 18
            Set eteilObject = e18
        case 19
            Set eteilObject = e19
        case 140
            Set eteilObject = e140
        case 150
            Set eteilObject = e150
        case 160
            Set eteilObject = e160
        case 180
            Set eteilObject = e180
        case 190
            Set eteilObject = e190
    end select
    
    'response.write "vartype(eteilObject): " & vartype(eteilObject) & "<br>"
    if vartype(eteilObject)<>vbEmpty then
        eteilObject.getHolding()    
    end if
end sub

'**********************************************************************

' sub setWSObject(wsNo, wsObject)
    ' select case wsNo
        ' case 1
            ' set wsObject = ws1
        ' case 2
            ' set wsObject = ws2
        ' case 3
            ' set wsObject = ws3
        ' case 4
            ' set wsObject = ws4
        ' case 5
            ' set wsObject = ws5
        ' case 6
            ' set wsObject = ws6
        ' case 7
            ' set wsObject = ws7
    ' end select
' end sub

'**********************************************************************
' sub checkWSavailability(ws,pOrder)
    ' batchsizeOld = pOrder.batchsize
    ' for j=pOrder.batchsizeRequired to 1 step -1
        ' if (ws.t + ws.preptime + (j * ws.prodTime)) <= (8 * 60) then
            ' pOrder.batchsize = j
            ' if pOrder.batchsize = pOrder.batchsizeRequired then
                'response.write "fall 1"
                ' preconditions(0)=1
                ' response.write "Es können alle gefordeten Teile im Los (" & batchsizeRequired & " Stück) in der verbleibenden Arbeitszeit an AP " &ws.nr& " von "_
                ' & (8*60 - ws.t) & " Minuten produziert werden.<br>"
                ' exit for
            ' else
                'response.write "fall 2"
                'preconditions(0)=1
                 ' response.write "Es sollen " & batchsizeRequired & " Teile im aktuellen Los produziert werden. Die verbleibende Arbeitszeit an AP " &ws.nr& " von "_
                     ' & (8*60 - ws.t) & " Minuten genügt für " & batchsize & " Teile.<br>"
                ' exit for
            ' end if
        ' else
            'response.write "fall 3"
            'preconditions(0)=0
            ' response.write "Die verbleibende Arbeitszeit an AP" & ws.nr & " von " & (8*60 - ws.t) & " genügt nicht, um ein Teil dieses Loses zu fertigen."_
                ' & "Der Auftrag wird am nächsten Arbeitstag fortgesetzt."
        ' end if
    ' next
    ' if batchsizeOld < batchsize then
        ' Response.write "Der erste (vorherige) Arbeitsplatz in der Arbeitsplatzkette beschränkt die Losgrösse weiter.<br>"
        ' batchsize = batchsizeOld
    ' end if
' end sub

'************************************************************
sub getUsergroup(usergroup, groupObj)
    '********************************************************************************************
    'getUsergroup() uses the userId (from VEForum-table "User", available as a session variable)
    'to determine the group he is assigned to and returns the id of the corresponding erp group
    '********************************************************************************************
    userId = Session("UserID") 'userId in VE-Forum
    response.write "userId: " & userId & "<br>"

   'connect to VEForum-Db
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open DB
    'dim ConnectionVEForum
    'call initDbVEForum(ConnectionVEForum)
    
    'get userId from session variable
'    userId = Session("UserID") 'userId in VE-Forum
   ' response.write "userId: " & userId & "<br>"
    if userId = "" then
        response.write "Es konnte keine Nutzerkennung aus den Session-Variablen ermittelt werden.<br>"
    '    & "Dieses Problem tritt bei der Nutzung des VE-Forums mit <b>Mozilla Firefox</b> auf "_
     '   & "und existiert unabhängig von der ERP-Simulation.<br>"_
      '  & "Bitte benutzen Sie den <b>Microsoft Internet-Explorer</b> als Browser.<br>"
    else
    'search assigned erp group to the userId in VEForum-table "Members" sql-query    
    'sql = "SELECT * FROM Members WHERE IDuser = "&userId&" "
    sql = "SELECT * FROM Members WHERE IDuser = "&userId&" AND IDProject = 520 "
    Set Recordset11 = Server.CreateObject("ADODB.Recordset")
    
    'start query
    Recordset11.Open sql, Conn
    'Recordset11.Open sql, ConnectionVEForum
    
    'parse results
    If Recordset11.EOF Then
        Response.Write "ERROR in erpProcedure.asp 'getUsergroup()'. User not present in table Members.<br>"
    Else
        'get erp-usergroup
        usergroup = Recordset11("Custom1")
        'response.write "usergroup: " & usergroup & "<br>"
        if usergroup=vbempty OR usergroup="" then
            response.write "Sie sind noch keiner Spielgruppe zugeordnet. Bitte melden Sie sich beim Spielleiter.<br>"
            'response.end
        end if
        'try to move rs-pointer forward (actually that shouldn't work, there should be only 1 result)
				'Note: Apparently this can happen. User 9722 is there currently twice (as he is working for two different projects???) -> now checking for IDProject=520 in SQL-query. NEED TO VERIFY THIS!!!
        Recordset11.MoveNext()
        Do While NOT Recordset11.Eof
            'response.write "ERROR in getUsergroup(). sql-query returned more than one result for for a single userId in the VEForum table 'Members'.<br>"
            'fetch next entry
            Recordset11.MoveNext()
        Loop
    end if
    
    'clean memory, destroy objects
    Recordset11.close
    set Recordset11 = Nothing
    Conn.close
    set Conn = Nothing
    ' ConnectionVEForum.close
    ' set ConnectionVEForum = Nothing
    
    '************************
    'create groupObj
    '************************
    'get group-values from erpSim.mdb table "Gruppen"
    sql = "SELECT * FROM Gruppen WHERE id="&usergroup&" "
    'initialise DB
    call initDB(Connection)
    Set rs=Server.CreateObject("ADODB.Recordset")    
    rs.Open sql, Connection
    
    '***********************************
    'iterate through purchase orders
    '***********************************
    If rs.EOF Then
        Response.Write("ERROR in erpProcedures 'getUsergroup'. Group "&usergroup&"not found in erpSim.mdb table 'Gruppen'.<br>")
    else    
        'assign values to groupObj
        groupObj.id = usergroup ' ==rs(0)
        groupObj.course = rs(1)
        groupObj.group = rs(2)
        groupObj.name = rs(1) & "/" & rs(2)
    end if    
    
    'clear memory, destroy objects
    rs.close
    set rs=Nothing
    Connection.close
    set Connection = Nothing
    
    'for use on localhost only
    'usergroup=1
    end if
end sub


'***********************************************************

sub getPeriod(period)
    period=1
end sub

'***********************************************************

sub closeDB(Recordset, Connection)
    Recordset.Close
    Set Recordset=Nothing    
    Connection.close
    Set Connection = Nothing
end sub

'*****************************************************************************************

'************************************************
' output of purchases and production orders
'************************************************
sub writeHeader(usergroup, period, d)
    if d="void" then
        response.write "***************************************************************************************************<br>"
        response.write "GRUPPE: " &groupObj.name& " &nbsp &nbsp &nbsp &nbsp Periode: " &period& " &nbsp &nbsp Wochenzusammenfassung <br>"
        response.write "***************************************************************************************************"
    else
        response.write "***************************************************************************************************<br>"
        response.write "GRUPPE: " &groupObj.name& " &nbsp &nbsp &nbsp &nbsp Periode: " &period& " &nbsp &nbsp &nbsp &nbsp Tag: " &d& " <br>"
        response.write "***************************************************************************************************"
        response.write "<h3>LAGERZUGANG</h3>"
        response.write "<h4>Aus Bestellungen:</h4>"
    end if
end sub

'************************************************

' sub writePurchases(positionFlag)    

    ' if positionFlag = "first" OR positionFlag = "firstLast" then
        ' response.write "<table>"
        ' response.write "<tr>"
        ' response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Zugang &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "</tr>"
        
        ' positionFlag = ""
    ' end if
    
    ' response.write "<tr>"
    ' response.write "<td align='center'>" & kteilObject.nr & "</td>"
    ' response.write "<td align='center'>" & kteilObject.name & "</td>"
    ' response.write "<td align='center'>" & purchOrder.delivered & "</td>"
    ' response.write "</tr>"
    
    ' if positionFlag = "last" OR positionFlag = "firstLast" then
        ' response.write "</table>"
        ' response.write "<br><br>"
    ' end if
' end sub

'**************************************************

' sub writeProduction(pOrder, eteilObject, positionFlag)
    'response.write "positionFlag: " & positionFlag & "<br>"
    ' if positionFlag = "first" OR positionFlag = "firstLast" then
        'response.write "Aus Produktion:<br><br>"
        ' response.write "<table>"
        ' response.write "<tr>"
        ' response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp AuftragsNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp TeilNr &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "<th>&nbsp&nbsp&nbsp&nbsp Bezeichnung &nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Zugang &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        ' response.write "<th> &nbsp&nbsp noch zu fertigen &nbsp&nbsp</th>"
        ' response.write "</tr>"
        
        ' if positionFlag = "first" then
            ' positionFlag = ""
        ' end if
    ' end if
    
    ' response.write "<tr>"
    'auftragsnr
    ' response.write "<td align='center'>" & pOrder.prodOrderNo & "</td>"
    'TeilNr
    ' response.write "<td align='center'>" & eteilObject.nr & "</td>"
    'Bezeichnung
    ' response.write "<td align='center'>" & eteilObject.name & "</td>"
    ' batchsize
    ' response.write "<td align='center'>" & pOrder.finished & "</td>"
    ' noch zu fertigen
    ' response.write "<td align='center'>" & pOrder.batchsizeRequired - pOrder.finished & "</td>"    
    ' response.write "</tr>"
    
    ' if positionFlag = "last" OR positionFlag = "firstLast" then
        ' response.write "</table>"
        ' response.write "<br><br>"
    ' end if
' end sub


'**********************************************************************************************************

'*********************************
'OUTPUT OF WORKSTATION WORKLOAD
'*********************************
sub writeWL(ws)
    if ws.nr = 1 then
        'output table header
        response.write "<h3>MASCHINENBELEGUNG</h3>"
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
    
    'split workload string "wsNr,starttime,stoptime,duration,prodOrderNo,batchsize,eteilnr;..."
    wl = split(ws.workload,";")
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
        response.write "<td align='center'>" & ws.Nr & "</td>"
        response.write "<td align='center'>-</td>"
        response.write "</tr>"
        response.write "<tr>"
        response.write "<td align='center' colspan=7>------------------------------------------------------------------------------------------------------------------------------------------------------------------------</td>"
        response.write "<tr>"
    end if
    
    if ws.nr = 7 then
        response.write "</table>"
        response.write "<br><br>"
    end if
end sub

'**********************************************************************

'********************************
'initialise database connection
'********************************
sub initDB(Connection)
    'create an ADO connection object
    Set Connection=Server.CreateObject("ADODB.Connection")
    'declare the variable that will hold the connection string
    Dim ConnectionString
    'define connection string, specify database driver and location of the database
    ConnectionString="PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source= d:\Inetpub\VEForum\mdb\erpSim_4.mdb"
    'ConnectionString="PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    '"Data Source= c:\inetpub\wwwroot\db\erpSim.mdb"
    
    'open the connection to the database
    Connection.Open ConnectionString
end sub

'***********************************************************************

' sub initDbVEForum(ConnectionVEForum)
    'create an ADO connection object
    ' Set ConnectionVEForum = Server.CreateObject("ADODB.Connection")
    'declare the variable that will hold the connection string
    ' Dim ConnectionStringVEForum
    'define connection string, specify database driver and location of the database
    ' ConnectionStringVEForum = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
    ' "Data Source= d:\Inetpub\VEForum\mdb\VEForum.mdb"
    'open the connection to the database
    ' ConnectionVEForum.Open ConnectionString
' end sub


'*************************************
'write input rows for purchase pieces
'*************************************
sub writeInputRowK(purchOrder, positionFlag)
    '*************************
    'table header
    '*************************
    if positionFlag = "first" OR positionFlag = "firstLast" then
        'response.write "<table cellpadding='10' id='tableK' border='1'>"
        response.write "<table id='tableK' border='0' cellpadding='10'>"
        response.write "<tr>"
        response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Kaufteil &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp Menge &nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        response.write "</tr>"
        
        if positionFlag = "first" then
            positionFlag = ""
        end if
    end if
    
    '***************************
    'write row of input fields
    '***************************
    'give <tr> an id for the last entry
    if positionFlag = "last" OR positionFlag = "firstLast" then
        response.write "<tr id='TR_K'>"
    else
        response.write "<tr>"
    end if
    'response.write "<input type=""hidden"" name=""kOrderId"" value="&kOrder.id&">"
    
    'drop-down list
    '---------------        
    'response.write "Kaufteil: <select name=""kteile"">"    
    response.write "<td align='center'><select name=""kteile"">"    
    response.write "<option></option>"
    
    for i=20 to 44
        'call setKTeilObject(i,kteilObject)
        set kteilObject = kteileArray(i-20)
        if i=purchOrder.kteilnr then
            'selected
            response.write "<option selected value='"&i&"'>" &i& ": " &kteilObject.name& "</option>"
        else
            'not selected, but in list
            response.write "<option value='"&i&"'>" &i& ": " &kteilObject.name& "</option>"
        end if
    next        
    response.write "</select>"
    response.write "</td>"

    'text-fields
    '------------
    'response.write " Menge: <input name=""amountKTeile"" type=""text"" size=""4"" maxlength=""4"" value="&kteilAmount&" onchange=""addKTeilRow()"">"
    'give <input> an id for the last entry
    'if positionFlag = "last" OR positionFlag = "firstLast" then
    response.write " <td align='center'><input name=""amountKTeile"" id=""amountKTeile"" type=""text"" size=""4"" maxlength=""4"" value="&purchOrder.amount&" onchange=""addKTeilRow()""></td>"
    response.write "<input type=""hidden"" name=""purchOrderId"" value="&purchOrder.id&">"

    
    'place holder for delete-button (which is inserted by JavaScript because it has to be defined later as forms cannot be nested)
    'response.write "<td align='center'></td>"
    
    response.write "</tr>"
    
    
    '****************
    'table footer
    '****************
    if positionFlag = "last" OR positionFlag = "firstLast" then
        response.write "</table>"
        response.write "<br><br>"
    end if
end sub

'**************************************************

'***********************************************
'write input rows for production order (ETeile)
'***********************************************
sub writeInputRowE(pOrder, positionFlag)
    tableId = "tableE" '& pOrder.day
    '*************************
    'table header
    '*************************
    if positionFlag = "first" OR positionFlag = "firstLast" then
        'response.write "creating table, id: " & tableId & "<br>"
        response.write "<table cellpadding='10' id="&tableId&">"
        response.write "<tr>"
        response.write "<th>AuftragsNr</th>"
        response.write "<th>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp Eigenfertigungsteil &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        'response.write "<th>&nbsp&nbsp&nbsp&nbsp Produktionstag &nbsp&nbsp&nbsp&nbsp</th>"
        response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp Losgröße &nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        response.write "<th> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>"
        response.write "</tr>"
        if positionFlag = "first" then
            positionFlag = ""
        end if
    end if
    
    '***************************
    'write row of input fields
    '***************************
    if positionFlag = "last" OR positionFlag = "firstLast" then
        response.write "<tr id='TR_E'>"
    else
        response.write "<tr>"
    end if
    
    'production-order number
    '------------------------
    if pOrder.prodOrderNo = "void" then
        response.write "<td align='center'>noch keine</td>"
    else
        response.write "<td align='center'>"&pOrder.prodOrderNo&"</td>"
    end if
    response.write "<input type=""hidden"" name=""prodOrderNo"" value="&pOrder.prodOrderNo&">"
    response.write "<input type=""hidden"" name=""pOrderId"" value="&pOrder.id&">"
    response.write "<input type=""hidden"" name=""tage"" value="&pOrder.day&">"
    
    
    'drop-down lists
    '---------------        
    'list 1 (eteil)
    response.write "<td align='center'><select name=""eteile"" size=""1"">"
    response.write "<option></option>"
    for i=1 to 19
        'call setETeilObject(i,eteilObject)
        call setItemObject(i,eteilObject)
        if i = pOrder.eteilnr then
            'selected
            response.write "<option selected value='"&i&"'>" &i& ": " &eteilObject.name& "</option>"
        else
            'not selected, but in list
            response.write "<option value='"&i&"'>" &i& ": " &eteilObject.name& "</option>"
        end if
    next
    response.write "</select>"
    response.write "</td>"
    
    'list 2 (day)
    'response.write "<td align='center'><select name=""tage"" size=""1"">"
    '    response.write "<option></option>"
    '    for i=1 to 5
    '        if i = pOrder.day then
                'selected
    '            response.write "<option selected>" &i& "</option>"
    '        else
                'not selected, but in list
    '            response.write "<option>" &i& "</option>"
    '        end if
    '    next
    'response.write "</select>"
    'response.write "</td>"
    
    
    'text-fields
    '------------
    'response.write " Menge: <input name=""amountETeile"" type=""text"" size=""4"" maxlength=""4"" value="&pOrder.batchsizeRequired&" onchange=""addETeilRow(document.usrInput.kteile)"">"
    response.write "<td align='center'><input name=""batchsizeRequired"" type=""text"" size=""4"" maxlength=""4"" value="&pOrder.batchsizeRequired&" onchange=""addETeilRow('"&tableId&"')""></td>"
    
    response.write "</tr>"
    
    
    '****************
    'table footer
    '****************
    if positionFlag = "last" OR positionFlag = "firstLast" then
        response.write "</table>"
        response.write "<br><br>"
    end if
    
end sub


'****************************************************
' Simple functions to convert the first 256 characters
' of the Windows character set from and to UTF-8.

' Written by Hans Kalle for Fisz
' http://www.fisz.nl

'IsValidUTF8
'  Tells if the string is valid UTF-8 encoded
'Returns:
'  true (valid UTF-8)
'  false (invalid UTF-8 or not UTF-8 encoded string)
function IsValidUTF8(s)
  dim i
  dim c
  dim n

  IsValidUTF8 = false
  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      select case n
      case 1
        exit function
      case 2
        if (c and &HE0) <> &HC0 then
          exit function
        end if
      case 3
        if (c and &HF0) <> &HE0 then
          exit function
        end if
      case 4
        if (c and &HF8) <> &HF0 then
          exit function
        end if
      case else
        exit function
      end select
      i = i + n
    else
      i = i + 1
    end if
  loop
  IsValidUTF8 = true
end function

'DecodeUTF8
'  Decodes a UTF-8 string to the Windows character set
'  Non-convertable characters are replace by an upside
'  down question mark.
'Returns:
'  A Windows string
function DecodeUTF8(s)
  dim i
  dim c
  dim n

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c and &H80 then
      n = 1
      do while i + n < len(s)
        if (asc(mid(s,i+n,1)) and &HC0) <> &H80 then
          exit do
        end if
        n = n + 1
      loop
      if n = 2 and ((c and &HE0) = &HC0) then
        c = asc(mid(s,i+1,1)) + &H40 * (c and &H01)
      else
        c = 191
      end if
      s = left(s,i-1) + chr(c) + mid(s,i+n)
    end if
    i = i + 1
  loop
  DecodeUTF8 = s
end function

'EncodeUTF8
'  Encodes a Windows string in UTF-8
'Returns:
'  A UTF-8 encoded string
function EncodeUTF8(s)
  dim i
  dim c

  i = 1
  do while i <= len(s)
    c = asc(mid(s,i,1))
    if c >= &H80 then
      s = left(s,i-1) + chr(&HC2 + ((c and &H40) / &H40)) + chr(c and &HBF) + mid(s,i+1)
      i = i + 1
    end if
    i = i + 1
  loop
  EncodeUTF8 = s
end function


%>


