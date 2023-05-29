
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''NEW SOLUTION'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub btnPrintInvoice_Click()
                                    
'-----------------------------------—-------------------------------------------
    Dim sDate As String
    Dim Answer_1 As String
    Dim Sec_ As String
    Sec_ = Sheet20.Range("K1")
    If (Range("R14").Value > 8) Or (Range("O16").Value = "") Then
        MsgBox ("Cela ukhethe usuku azoDilivelwa ngalo")
        GoTo Quit
    Else
        Answer_1 = MsgBox("Did you select a delivery day!", vbYesNo)
        If Not IsEmpty(invoice1.Range("L6").Value) And (Answer_1 = vbYes) Then
            sDate = Format(Now(), "ddmmyy") & " - " & Format(Now(), "hhmmss")
            invoice1.Range("A1:M42").ExportAsFixedFormat xlTypePDF, Filename:="C:\Users\Mean_Machine\Downloads\Documents\" & invoice1.Range("D7").Value & "  " & invoice1.Range("D8").Value & " " & invoice1.Range("L6").Value & " " & (sDate), openafterpublish:=False
            Application.Dialogs(xlDialogPrint).Show
        Else
            GoTo Quit
        End If
    End If
'-----------------------------------—-------------------------------------------
    If invoice1.Range("C25") = Empty Then
        Dim Item, invoiceID As Range
        Dim NewCol, CurrRow, iRemainder, i, j, rows As Integer
        Dim dSum, LSum, iAmount, iRem, pAmount, Sum, iBal, sSum, bSum As Double
        Dim sTarget, sInvoice, sName, sSurname, sLocation, sItem, sStatus, sContact, sID As String
        sTarget = invoice1.Range("L6")
        sName = invoice1.Range("D7").Value
        sSurname = invoice1.Range("D8").Value
        sLocation = invoice1.Range("D9").Value
        sContact = "'" & invoice1.Range("D10").Value
        iRemainder = "=[Quantity]-[Delivered]"
        sStatus = "=If([Remainder] + [Balance] = 0, " & Chr(34) & "Closed" & Chr(34) & "," & Chr(34) & "Active" & Chr(34) & " )"
        LSum = invoice1.Range("Z1").Value
        sSum = invoice1.Range("U1").Value
        bSum = invoice1.Range("Y1").Value
        rows = 3
        If Left(invoice1.Cells(14, 3).Value, 2) <> "Pr" Then
        Application.ScreenUpdating = False
        For i = 14 To 26
        For Each Item In invoice1.Cells(i, rows)
            dSum = 0
''''''''''''''''''''''''''''''''''''''''GO TO Blocks ACOUNTS''''''''''''''''''''''''''''''''C:\Users\Mawox Business Hub\Desktop\Saved Receipts\''''''''''''''''
                If (Left(Item, 6) = "Blocks") And invoice1.Range("AD1") = "Y" Then
                    Sheet4.Activate
                    dSum = dSum + invoice1.Range("Y1").Value
                    If Left(invoice1.Cells(i + 1, 3).Value, 1) = "B" And (bSum >= invoice1.Cells(i, 11).Value) Then
                        bSum = bSum - invoice1.Cells(i, 11).Value
                        dSum = invoice1.Cells(i, 11).Value
                    Else 'For Balance
                        dSum = bSum
                        bSum = 0
                    End If
                    GoTo Update
                End If
    ''''''''''''''''''''''''''''''''''''''''''GO TO Partition ACOUNTS''''''''''''''''''''''''''''''''''''''''''''''''
                If (Item = "Partition") And invoice1.Range("AD1") = "Y" Then
                    Sheet9.Activate
                    dSum = dSum + invoice1.Range("X1").Value
                    GoTo Update
                End If
    ''''''''''''''''''''''''''''''''''''''''GO TO Cement ACOUNTS''''''''''''''''''''''''''''''''''''''''''''''''
                If Item = "Cement" And invoice1.Range("AD1") = "Y" Then
                    Sheet22.Activate
                    dSum = dSum + invoice1.Range("W1").Value
                    GoTo Update
                End If
    ''''''''''''''''''''''''''''''''''''''''''GO TO Lintels ACOUNTS''''''''''''''''''''''''''''''''''''''''''''''''
                If Left(Item, 2) = "Le" And invoice1.Range("AD1") = "Y" Then
                    Sheet5.Activate
                    If (Left(invoice1.Cells(i + 1, 3).Value, 1) = "L") And (LSum >= invoice1.Cells(i, 11).Value) Then
                        LSum = LSum - invoice1.Cells(i, 11).Value
                        dSum = invoice1.Cells(i, 11).Value
                    Else 'For Balance
                        dSum = LSum
                        LSum = 0
                    End If
                    GoTo Update
                End If

    ''''''''''''''''''''''''''''''''''''''''''''GO TO Sand ACOUNTS''''''''''''''''''''''''''''''''''''''''''''''''
                If Left(Item, 1) = "S" And invoice1.Range("AD1") = "Y" Then
                    Sheet6.Activate
                    dSum = dSum + invoice1.Range("U1").Value
                    If Left(invoice1.Cells(i + 1, 3).Value, 1) = "S" And (sSum >= invoice1.Cells(i, 11).Value) Then
                        sSum = sSum - invoice1.Cells(i, 11).Value
                        dSum = invoice1.Cells(i, 11).Value
                    Else 'For Balance
                        dSum = sSum
                        sSum = 0
                    End If
                    GoTo Update
                End If
                
    '''''''''''''''''''''''''''''''''''''''''''GO TO DELIVERY ACOUNTS''''''''''''''''''''''''''''''''''''''''''''''''
                If invoice1.Range("C26") <> Empty And invoice1.Range("AD1") = "Y" Then
                    i = 26
                    Sheet11.Activate
                    dSum = dSum + invoice1.Range("V1").Value
                    GoTo Update
                End If
    '''''''''''''''''''''''''''''''''''''''''''''ACCOUNTS UPDATE CODE''''''''''''''''''''''''''''''''''''''''''''''''
Update:
                With ActiveSheet
                    CurrRow = 3
                    NewCol = 0
                    .Unprotect (Sec_)
                    .Range("A3").Select
                    ActiveCell.EntireRow.Insert shift:=xlDown
                    .Cells(CurrRow, NewCol + 1).Value = sTarget                               'Invoice ID
                    .Cells(CurrRow, NewCol + 2).Value = sName                                 'Name
                    .Cells(CurrRow, NewCol + 3).Value = sSurname                              'Surname
                    .Cells(CurrRow, NewCol + 4).Value = invoice1.Cells(i, 3).Value            'Product Type
                    .Cells(CurrRow, NewCol + 5).Value = invoice1.Cells(i, 9).Value            'No. of Items
                    .Cells(CurrRow, NewCol + 7).Value = "=[Quantity]-[Delivered]"             'Remainder
                    .Cells(CurrRow, NewCol + 9).Value = invoice1.Range("O16")                 'Promised Delivery Date
                    .Cells(CurrRow, NewCol + 10).Value = sContact                             'Cell No
                    .Cells(CurrRow, NewCol + 11).Value = sLocation                            'Customers' Location
                    .Cells(CurrRow, NewCol + 12).Value = sStatus                              'Status
                    .Cells(CurrRow, NewCol + 13).Value = dSum                                 'Amount Paid
                    .Cells(CurrRow, NewCol + 14).Value = invoice1.Cells(i, 11).Value          'Actual Amount
                    .Cells(CurrRow, NewCol + 15).Value = Date                                 'Current Date
                    .Cells(CurrRow, NewCol + 16).Value = invoice1.Range("H27").Value
                    .Cells(CurrRow, NewCol + 8).Value = (.Range("N3").Value - .Range("M3").Value)                                   'Balance
                    .Protect (Sec_)
                End With
'''''''''''''''''''''''''''''''''''''''''''''''''''''CLOSING THE CODE'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Next Item
        Next i
        ''''''''''''''''''''''''''''''''''''''''''GO TO PREVIOUS BALANCE''''''''''''''''''''''''''''''''''''''''''''''''
        Else
            j = 0
            Do While j < 7
                invoice1.Unprotect (Sec_)
                If invoice1.Range("Y1") > 0 Then ''''''''''''''''''''''''''Blocks
                    Sheet4.Activate
                    With ActiveSheet
                        .Unprotect (Sec_)
                         i = 3
                        iRem = 1
                        Do While i >= 3 And iRem > 0
                            With ActiveSheet
                                sInvoice = .Cells(i, 1).Value
                                If sInvoice = sTarget Then
                                    pAmount = invoice1.Range("Y1").Value
                                    iRem = pAmount - .Cells(i, 8).Value
                                    If iRem >= 0 Then
                                        iBal = 0
                                        iAmount = pAmount - iRem
                                    Else
                                        iBal = iRem * (-1)
                                        iAmount = pAmount
                                    End If
                                    invoice1.Range("Y1").Value = iRem
                                    .Cells(i, 8).Value = iBal
                                    .Cells(i, 13).Value = iAmount
                                    .Cells(i, 15).Value = Date
                                    .Cells(i, 16).Value = invoice1.Range("H27").Value
                                    If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
                                        i = 0
                                    End If
                                End If
                                i = i + 1
                            End With
                        Loop
                        invoice1.Range("Y1").Value = 0
                        .Protect (Sec_)
                        GoTo UpdateBal
                    End With
                End If
                If invoice1.Range("X1") > 0 Then  '''''''''''''''''''''''''Partition
                    Sheet9.Activate
                    With ActiveSheet
                        .Unprotect (Sec_)
                        iAmount = invoice1.Range("X1").Value
                        iBal = iAmount - Application.WorksheetFunction.VLookup(sTarget, Range("Table824"), 8, 0)
                        iBal = iBal
                        invoice1.Range("X1").Value = 0
                        .Protect (Sec_)
                        GoTo UpdateBal
                    End With
                End If
                If invoice1.Range("W1") > 0 Then '''''''''''''''''''''''''''Cement
                    Sheet22.Activate
                    With ActiveSheet
                        .Unprotect (Sec_)
                        iAmount = invoice1.Range("W1").Value
                        iBal = iAmount - Application.WorksheetFunction.VLookup(sTarget, Range("Table65"), 8, 0)
                        invoice1.Range("W1").Value = 0
                        .Protect (Sec_)
                        GoTo UpdateBal
                    End With
                End If
                If invoice1.Range("Z1") > 0 Then ''''''''''''''''''''''''''Lentils & Poles
                    Sheet5.Activate
                    Sheet5.Unprotect (Sec_)
                    i = 3
                    iRem = 1
                    Do While i >= 3 And iRem > 0
                        With ActiveSheet
                            sInvoice = .Cells(i, 1).Value
                            If (sInvoice = sTarget) And (.Cells(i, 8).Value > 0) Then
                                pAmount = invoice1.Range("Z1").Value
                                iRem = pAmount - .Cells(i, 8).Value
                                If iRem >= 0 Then
                                    iBal = 0
                                    iAmount = pAmount - iRem
                                Else
                                    iBal = iRem * (-1)
                                    iAmount = pAmount
                                End If
                                invoice1.Range("Z1").Value = iRem
                                .Cells(i, 8).Value = iBal
                                .Cells(i, 13).Value = iAmount
                                .Cells(i, 15).Value = Date
                                .Cells(i, 16).Value = invoice1.Range("H27").Value
                                'Loop Escape
                                If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
                                    i = 0
                                End If
                            End If
                            i = i + 1
                        End With
                    Loop
                    Sheet5.Protect Sec_
                    invoice1.Range("Z1").Value = 0
                    dSum = "Done"
                End If
                If invoice1.Range("U1") > 0 Then  '''''''''''''''''''''''''''Sand
                    Sheet6.Activate
                    Sheet6.Unprotect (Sec_)
                    i = 3
                    iRem = 1
                    Do While i >= 3 And iRem > 0
                        With ActiveSheet
                            sInvoice = .Cells(i, 1).Value
                            If sInvoice = sTarget Then
                                pAmount = invoice1.Range("U1").Value
                                iRem = pAmount - .Cells(i, 8).Value
                                If iRem >= 0 Then
                                    iBal = 0
                                    iAmount = pAmount - iRem
                                Else
                                    iBal = iRem * (-1)
                                    iAmount = pAmount
                                End If
                                invoice1.Range("U1").Value = iRem
                                .Cells(i, 8).Value = iBal
                                .Cells(i, 13).Value = iAmount
                                .Cells(i, 15).Value = Date
                                .Cells(i, 16).Value = invoice1.Range("H27").Value
                                If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
                                    i = 0
                                End If
                            End If
                            i = i + 1
                        End With
                    Loop
                    invoice1.Range("U1").Value = 0
                    Sheet6.Protect Sec_
                    dSum = "Done"
                End If
                If invoice1.Range("V1") > 0 Then ''''''''''''''''''''''''''Delivery
                    Sheet11.Activate
                    With ActiveSheet
                        .Unprotect (Sec_)
                        iAmount = invoice1.Range("V1").Value
                        iBal = iAmount - Application.WorksheetFunction.VLookup(sTarget, Range("Table712"), 8, 0)
                        invoice1.Range("V1").Value = 0
                        .Protect (Sec_)
                        GoTo UpdateBal
                    End With
                End If
                invoice1.Protect (Sec_)
            
            If dSum = "Done" Then
                GoTo Quit
            End If
UpdateBal:
                i = 3
                Do While i >= 3
                    sInvoice = ActiveSheet.Cells(i, 1).Value
                    If sInvoice = sTarget Then
                        With ActiveSheet
                            .Unprotect Sec_
                            .Cells(i, 8).Value = iBal * (-1)
                            .Cells(i, 13).Value = iAmount
                            .Cells(i, 15).Value = Date
                            .Cells(i, 16).Value = invoice1.Range("H27").Value
                            .Cells(i, 17).Value = ""
                            .Protect Sec_
                            If .Cells(i, 1).Value <> .Cells(i + 1, 1).Value Then
                                i = 0
                            End If
                        End With
                        
                    End If
                    i = i + 1
                Loop
                j = j + 1
            Loop
                If iRem < 0 Then
                    GoTo Quit
                End If
                
End If
End If
Quit:
    invoice1.Activate
    ActiveWorkbook.Save
End Sub
Sub ShowServices_Click()
    frmServices.Show
End Sub


