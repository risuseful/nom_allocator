Option Explicit

Private Sub CommandButton1_Click()

    '//--------------------------------------------------------------
    '-----------------copying well priority status---------------
    
        'Author: Akmal Aulia
        'Last update: 30/3/2016
        
        'declare variables
        Dim sqlStr2 As String
        Dim i As Integer
        Dim wnam(130) As String
        
        'declare connection variables
        Dim oConn2 As ADODB.Connection
        Dim rs2 As ADODB.Recordset
        Dim sConnString2 As String
        
         'declare worksheet variable
        Dim wdat As Worksheet, wc As Worksheet
        
        'set worksheets
        Set wdat = Sheets("WMI Input")
        Set wc = Sheets("CO2 Real Time")
        
        'create connection string
        sConnString2 = "Provider=******;data source=******;Initial Catalog=******;User Id= ******;Password= ****** "
        
        'create the connection and recordset objects
        Set oConn2 = New ADODB.Connection
        Set rs2 = New ADODB.Recordset
        
        'Open connection and execute
        oConn2.Open sConnString2
        
        'unlock sheets
        wc.Unprotect "###"
        wdat.Unprotect "###"
        
        'record well input data
        For i = 1 To 128
        
            'store well names
            wnam(i) = wdat.Cells(1 + i, 2)
            
            'record MSFR
            'sqlStr = "SELECT MSFR FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
            'Set rs = oConn.Execute(sqlStr)
            'wdat.Cells(12 + i, 3).CopyFromRecordset rs
            
            'record CO2
            'sqlStr = "SELECT CO2 FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
            'Set rs = oConn.Execute(sqlStr)
            'wdat.Cells(12 + i, 4).CopyFromRecordset rs
            
            'record WELL_PRIORITY
            sqlStr2 = "SELECT WELL_PRIORITY FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
            Set rs2 = oConn2.Execute(sqlStr2)
            wdat.Cells(1 + i, 3).CopyFromRecordset rs2
            
            'record MSFR
            sqlStr2 = "SELECT MSFR FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
            Set rs2 = oConn2.Execute(sqlStr2)
            wc.Cells(21 + i, 6).CopyFromRecordset rs2
            
            'record CO2
            sqlStr2 = "SELECT CO2 FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
            Set rs2 = oConn2.Execute(sqlStr2)
            wc.Cells(21 + i, 3).CopyFromRecordset rs2
            wc.Cells(21 + i, 3) = wc.Cells(21 + i, 3) / 100
            
            'record GHV (or CV, i.e. calorific value)
            sqlStr2 = "SELECT GHV FROM IPDSStaging.dbo.vIPDS_Latest_MSFR_CO2 WHERE Well='" & wnam(i) & "';"
            Set rs2 = oConn2.Execute(sqlStr2)
            wc.Cells(21 + i, 5).CopyFromRecordset rs2
    
        Next i
        
        're-lock sheet "CO2 Real Time"
        wc.Protect Password:="*****", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True
        
        wdat.Protect Password:="*****", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True
        
        'Clean up
        If CBool(oConn2.State And adStateOpen) Then oConn2.Close
        Set oConn2 = Nothing
        Set rs2 = Nothing
    
    '--------------------copying well priority status---------------
    '--------------------------------------------------------------//
    
    'Akmal Aulia, 13/4/2016
    'This Sub executes the *Snapshot* button in the "Shapshot" sheet, followed by
    'the *Get Snapshot* button in the "ProdAlloc" sheet.
    
    Worksheets("CO2 Real Time").Protect "###", UserInterfaceOnly:=True
    Worksheets("Snapshot").Protect "###", UserInterfaceOnly:=True
    
    Application.ScreenUpdating = False 'don't show other sheets when activated
    
    '--------------------------------
    'execute *Snapshot* button
    '--------------------------------
    
    Worksheets("Snapshot").Activate
    'author: Akmal Aulia
    'update: March 23, 2016
    
    '///----------------DECLARE VARIABLES-------------------
    
    'integer
    Dim j As Integer, nrow1 As Integer, nrow2 As Integer
    Dim nrow3 As Integer, offSet1 As Integer, offSet2 As Integer, offSet3 As Integer
    
    
    '----------------DECLARE VARIABLES-------------------///
    
    '///----------------INITIALIZE VARIABLES-------------------
    nrow1 = ####
    nrow2 = ####
    nrow3 = ####
    offSet1 = ####
    offSet2 = ####
    offSet3 = ####
    
    '----------------INITIALIZE VARIABLES-------------------///
    
    'clear "Errors" sheet
    Worksheets("Errors").Cells.Clear
    
    'declare error counter
    Dim cnt As Integer
    cnt = 1
    
    'copy CO2 and rates
    For i = 1 To nrow1
    
        'copy CO2
        If IsEmpty(Worksheets("CO2 Real Time").Cells(i + offSet1, 3).Value) = False Then
            Worksheets("Snapshot").Cells(i + 21, 3) = Worksheets("CO2 Real Time").Cells(i + offSet1, 3)
        Else
            Worksheets("Errors").Cells(cnt, 1) = Now
            Worksheets("Errors").Cells(cnt, 2) = "Could not retrieve CO2 value from well " & Worksheets("CO2 Real Time").Cells(i + offSet1, 2) & "."
            cnt = cnt + 1
        End If
        
        'copy rate
        If IsEmpty(Worksheets("CO2 Real Time").Cells(i + offSet1, 4).Value) = False Then
            Worksheets("Snapshot").Cells(i + 21, 4) = Worksheets("CO2 Real Time").Cells(i + offSet1, 4)
        Else
            Worksheets("Errors").Cells(cnt, 1) = Now
            Worksheets("Errors").Cells(cnt, 2) = "Could not retrieve gross rate value from well " & Worksheets("CO2 Real Time").Cells(i + offSet1, 2) & "."
            cnt = cnt + 1
        End If
        
        
        'copy MSFR
        If IsEmpty(Worksheets("CO2 Real Time").Cells(i + offSet1, 6).Value) = False Then
            Worksheets("Snapshot").Cells(i + 21, 5) = Worksheets("CO2 Real Time").Cells(i + offSet1, 6)
        Else
            Worksheets("Errors").Cells(cnt, 1) = Now
            Worksheets("Errors").Cells(cnt, 2) = "Could not retrieve MSFR value from well " & Worksheets("CO2 Real Time").Cells(i + offSet1, 2) & "."
            cnt = cnt + 1
        End If
    
        
        
        'copy CV (calorific value)
        If IsEmpty(Worksheets("CO2 Real Time").Cells(i + offSet1, 5).Value) = False Then
            Worksheets("Snapshot").Cells(i + 21, 24) = Worksheets("CO2 Real Time").Cells(i + offSet1, 5)
        Else
            Worksheets("Errors").Cells(cnt, 1) = Now
            Worksheets("Errors").Cells(cnt, 2) = "Could not retrieve GHV value from well " & Worksheets("CO2 Real Time").Cells(i + offSet1, 2) & "."
            cnt = cnt + 1
        End If
        
        'copy Well Priority Index
        If IsEmpty(Worksheets("WMI Input").Cells(i + 1, 4).Value) = False Then
            Worksheets("Snapshot").Cells(i + 21, 27) = Worksheets("WMI Input").Cells(i + 1, 4)
        Else
            Worksheets("Errors").Cells(cnt, 1) = Now
            Worksheets("Errors").Cells(cnt, 2) = "Could not retrieve Well Priority Index value from well " & Worksheets("CO2 Real Time").Cells(i + offSet1, 2) & "."
            cnt = cnt + 1
        End If
        
        
        
    Next i
    
    'copy platform rates (non-unit)
    For i = 1 To nrow2
        Worksheets("Snapshot").Cells(i + 23, 9) = Worksheets("CO2 Real Time").Cells(i + offSet2, 11)
    Next i
    
    'copy platform rates (unit)
    For i = 1 To nrow3
        Worksheets("Snapshot").Cells(i + 37, 9) = Worksheets("CO2 Real Time").Cells(i + offSet3, 11)
    Next i
    
    Worksheets("Snapshot").Cells(18, 12) = Date & ", " & Time
    
    ' --------the following is deactivated (Oct. 10th, 2017)--------------
    '--------------------------------
    'execute *Get Snapshot* button
    '--------------------------------
    
    'Worksheets("ProdAlloc").Activate
    ''declare integers
    'Dim offSet11 As Integer, offSet12 As Integer
    'Dim offSet21 As Integer, offSet22 As Integer
    
    ''declare worksheet
    'Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    
    ''initialize variables
    'offSet11 = 23
    'offSet12 = 37
    'offSet21 = 18
    'offSet22 = 30
    'nrow1 = ### 'number of non-unit WHPs
    'nrow2 = ### 'number of unit WHPs
    'Set ws1 = Sheets("Snapshot")
    'Set ws2 = Worksheets("ProdAlloc")
    'ws2.Unprotect "###"
    
    
    ''read in data (non-unit)
    'For i = 1 To nrow1
    
        ''min rate
        'ws2.Cells(i + offSet21, 4) = ws1.Cells(i + offSet11, 12)
     
        ''current rate
        'ws2.Cells(i + offSet21, 5) = ws1.Cells(i + offSet11, 10)
        
        ''max rate
        'ws2.Cells(i + offSet21, 6) = ws1.Cells(i + offSet11, 11)
        
        ''gross CO2
        'ws2.Cells(i + offSet21, 7) = ws1.Cells(i + offSet11, 9)
    'Next i
    
    ''read in data (unit)
    'For i = 1 To nrow2
    
        ''min rate
        'ws2.Cells(i + offSet22, 4) = ws1.Cells(i + offSet12, 12)
        
        ''current rate
        'ws2.Cells(i + offSet22, 5) = ws1.Cells(i + offSet12, 10)
        
        ''max rate
        'ws2.Cells(i + offSet22, 6) = ws1.Cells(i + offSet12, 11)
        
        ''gross CO2
        'ws2.Cells(i + offSet22, 7) = ws1.Cells(i + offSet12, 9)
    'Next i
    
    ''reprotect sheet "ProdAlloc"
    'ws2.Protect Password:="*****", _
    'DrawingObjects:=True, _
    'Contents:=True, _
    'Scenarios:=True
    
    '//------checking for possibility of a wells being idle; temperature < threshold (40 C) -----
    'note: rate reading might gave the wrong reading
    Dim wIdle As String
    
    wIdle = ""
    For i = 1 To nrow1
    
        'search for possibly idle wells
        If wc.Cells(i + 21, 4) > 0.0000001 And wc.Cells(i + 21, 7) < wc.Cells(9, 7) Then
            wIdle = wIdle & wc.Cells(i + 21, 2) & ", "
        End If
    
    Next i
    

    
    If wIdle <> "" Then
    
        'replace last character with a dot
        wIdle = Left(wIdle, Len(wIdle) - 2)
        wIdle = wIdle & "."
    
        MsgBox "Warning! These wells might be idle: " & wIdle
    End If
    
    Worksheets("EASY").Activate
    
    If cnt = 1 Then
        MsgBox "Snapshot is successful."
    Else
        MsgBox "Snapshot had failed, Jim.  Check the log at the *Errors* sheet. The system found " & (cnt - 1) & " errors."
    End If
    
    
    
    

    
    
    '-------------------------------------------------------------------------//



End Sub

Private Sub CommandButton2_Click()

    'record current time
    Dim tInit As Date, tElaps As Date
    tInit = Now
    
    'variables
    Dim i As Integer
    
    'unprotect sheet "ProdAlloc" to enable Solver
    Worksheets("ProdAlloc").Activate
    ActiveSheet.Unprotect "###"
    
    
    ' //---------------------OPTIMIZATION STARTS HERE---------------------------
    
    'find solution gross rates
    
        'initialize optimized gross vectors (see column "S")
        'For i = 34 To 161
            'Worksheets("ProdAlloc").Range("I" & i) = Worksheets("ProdAlloc").Range("G" & i)
        'Next i
    
        'remove previous added constraints
        SolverReset
        
        'configure optimizer
        'SolverOptions Precision:=0.001, MaxTime:=300, Convergence:=0.0001, MutationRate:=0.1, PopulationSize:=150, RandomSeed:=0, MaxTimeNoImp:=50
        SolverOptions MaxTime:=300, MaxTimeNoImp:=80
        
        'set objective function and control parameters (i.e. every platform's gross rate)
        'note: MaxMinVal=2 means minimize
        SolverOK SetCell:="$H$7", MaxMinVal:=2, ByChange:="$I$34:$I$161", Engine:=3
            
        '*****************
        'add constraints
        '*****************
        'SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$G$12"), relation:=2, formulaText:=Worksheets("ProdAlloc").Range("$E$12") '### Unit constraint (sales)
        'SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$G$13"), relation:=2, formulaText:=Worksheets("ProdAlloc").Range("$E$13") '### Unit constraint (sales)
        'SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$G$14"), relation:=2, formulaText:=Worksheets("ProdAlloc").Range("$E$14") '### South Unit constraint (sales)
        'SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$G$15"), relation:=2, formulaText:=Worksheets("ProdAlloc").Range("$E$15") '###-Unit constraint (sales)
        
        SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$I$34:$I$161"), relation:=1, formulaText:=Worksheets("ProdAlloc").Range("$H$34:$H$161") 'control parameters max constraint
        SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$I$34:$I$161"), relation:=3, formulaText:=Worksheets("ProdAlloc").Range("$G$34:$G$161") 'control parameters min constraint
        
        SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$I$5"), relation:=3, formulaText:=Worksheets("ProdAlloc").Range("$I$4") 'co2 min constraint
        'SolverAdd cellRef:=Worksheets("ProdAlloc").Range("$S$14"), relation:=3, formulaText:=Worksheets("ProdAlloc").Range("$T$14")
        
        'run optimizer
        SolverSolve UserFinish:=True
    
    ' ---------------------OPTIMIZATION ENDS HERE---------------------------//
    
    
    
    
    '//--------------revised gross rates and copy to sheet "EASY"---------------------
    
    'unprotect sheet "EASY"
    Worksheets("EASY").Unprotect "###"
    
    Dim gs As Double, ms As Double, gBLS As Double, gSY As Double, mBLS As Double, mSY As Double
    Dim addBumi As Double, addSY As Double, addBLS As Double, delSY As Double, delBLS As Double
    
    'revised variables
    Dim rBLS As Double, rBMA As Double, rBMB As Double, rSYA As Double, rSYB As Double
    Dim offBumiU As Double, offSuriyaU As Double, offBlsU As Double
    
    'specify addition
    add##### = ###
    add### = ###
    add### = ###
    
    '-- ### Unit (add ### mmscfd to mimic actual shrinkage factor)
    ms = Worksheets("EASY").Cells(25, 6) 'msfr
    gs = Worksheets("ProdAlloc").Cells(12, 10) 'computed gross (i.e. numerical solution)
    
    If (ms - gs) > add### Then
        Worksheets("EASY").Cells(25, 5) = gs + add###
    Else
        Worksheets("EASY").Cells(25, 5) = ms
    End If
    
    '-- ### Unit (add ### mmscfd to mimic actual shrinkage factor)
    gs = Worksheets("ProdAlloc").Cells(13, 10)
    ms = Worksheets("EASY").Cells(26, 6)
    If (ms - gs) > addSY Then
        Worksheets("EASY").Cells(26, 5) = gs + addSY
    Else
        Worksheets("EASY").Cells(26, 5) = ms
        delSY = gs + addSY - ms 'the remaining from addSY
    End If
    
    '-- ### ### Unit (add ### mmscfd to mimic actual shrinkage factor)
    gs = Worksheets("ProdAlloc").Cells(14, 10)
    ms = Worksheets("EASY").Cells(27, 6)
    If (ms - gs) > addBLS Then
        Worksheets("EASY").Cells(27, 5) = gs + addBLS
    Else
        Worksheets("EASY").Cells(27, 5) = ms
        delBLS = gs + addBLS - ms 'the remaining from add###
    End If
    
    '//-----------the following lines were deactivated on 13/2/2018---------
    '### and ### ### reconciliation
    'g### = Worksheets("EASY").Cells(26, 5)
    'm### = Worksheets("EASY").Cells(26, 6)
    'g### = Worksheets("EASY").Cells(27, 5)
    'm### = Worksheets("EASY").Cells(27, 6)
    

    '------------------------------------------------------------------//
    
    
    
    
    'computed revised WHP gross rates
    offBumiU = Worksheets("EASY").Cells(25, 5) - Worksheets("ProdAlloc").Cells(12, 10)
    offSuriyaU = Worksheets("EASY").Cells(26, 5) - Worksheets("ProdAlloc").Cells(13, 10)
    offBlsU = Worksheets("EASY").Cells(27, 5) - Worksheets("ProdAlloc").Cells(14, 10)
    
    rBLS = Worksheets("ProdAlloc").Cells(24, 38) + offBlsU
    rBMA = Worksheets("ProdAlloc").Cells(25, 38) + (Worksheets("ProdAlloc").Cells(25, 38)) / (Worksheets("ProdAlloc").Cells(25, 38) + Worksheets("ProdAlloc").Cells(26, 38)) * offBumiU
    rBMB = Worksheets("ProdAlloc").Cells(26, 38) + (Worksheets("ProdAlloc").Cells(26, 38)) / (Worksheets("ProdAlloc").Cells(25, 38) + Worksheets("ProdAlloc").Cells(26, 38)) * offBumiU
    rSYA = Worksheets("ProdAlloc").Cells(27, 38) + (Worksheets("ProdAlloc").Cells(27, 38)) / (Worksheets("ProdAlloc").Cells(27, 38) + Worksheets("ProdAlloc").Cells(28, 38)) * offSuriyaU
    rSYB = Worksheets("ProdAlloc").Cells(28, 38) + (Worksheets("ProdAlloc").Cells(28, 38)) / (Worksheets("ProdAlloc").Cells(27, 38) + Worksheets("ProdAlloc").Cells(28, 38)) * offSuriyaU
    
    Worksheets("EASY").Cells(36, 14) = rBLS
    Worksheets("EASY").Cells(37, 14) = rBMA
    Worksheets("EASY").Cells(38, 14) = rBMB
    Worksheets("EASY").Cells(39, 14) = rSYA
    Worksheets("EASY").Cells(40, 14) = rSYB
    
    
    
    're-protect sheet "EASY"
    Worksheets("EASY").Protect "###"
    
    '--------------revised gross rates and copy to sheet "EASY"---------------------//
    
    
    
    
    'reprotect sheet "ProdAlloc"
    ActiveSheet.Protect Password:="******", DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Application.ScreenUpdating = True 'don't show other sheets when activated
    Worksheets("EASY").Activate
    
    '//--------------------push allocated values to server------------------------
    
    If Worksheets("Snapshot").Range("T15") = "y" Then
    
        'declare variables
        Dim dateStamp As Date, plat As String, sqlStr As String
        Dim sales As Double, gross As Double
        
        'declare connection variables
        Dim oConn As ADODB.Connection
        Dim rs As ADODB.Recordset
        Dim sConnString As String
        
        'create connection string
        sConnString = "Provider=******;data source=******;Initial Catalog=******;User Id= ******;Password= ****** "
        
        'create the connection and recordset objects
        Set oConn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        'Open connection and execute
        oConn.Open sConnString
        
        'assign values for current time
        dateStamp = Now()
        
        'push data to server
        For i = 1 To 17
        
            If i <> 11 Then '11th row is for SYA_0, which does not exist.
                'assign values
                plat = Cells(23 + i, 17)
                sales = Cells(23 + i, 13)
                gross = Cells(23 + i, 14)
            
                'execute SQL's insert statement
                sqlStr = "INSERT INTO AllocNormGas(Date_Stamp,Platform,NormGrossGas,NormSalesGas)  VALUES ('" & Format(Now(), "yyyy-mmm-dd hh:mm") & "','" & plat & "'," & gross & "," & sales & ");"
                Set rs = oConn.Execute(sqlStr)
            End If
                
        Next i
        
        'Clean up
        If CBool(oConn.State And adStateOpen) Then oConn.Close
        Set oConn = Nothing
        Set rs = Nothing
    
    End If
    
    '--------------------push allocated values to server------------------------//
    
    
    
    'report last allocation time
    ActiveSheet.Unprotect "###"
    Worksheets("EASY").Cells(16, 8) = Date & ", " & Time
    ActiveSheet.Protect "###"
    
    'record end time and display
    tElaps = tInit - Now
    MsgBox "Allocation completion time: " & DatePart("n", tElaps) & " minutes and " & DatePart("s", tElaps) & " seconds."


End Sub
