# Next-Day-Securities-Rec
Reconcile next day USD securities vs system and clypso file
Sub USDRecMovement()
Call USDNextDayMovement
'open Rec go to USD Recsheet
Dim wb, wbSecRec As Workbook
Dim wsUSD As Worksheet
Dim iLastrow, iLastrow1, iLastrowUSD As Long
Dim strFind As String
Set wb = ThisWorkbook

On Error GoTo ErrHandler:
Set wbSecRec = Workbooks.Open(Inbound.Range("AW18"))
Set wsUSD = wbSecRec.Sheets("USD Report")

'identify if any data is created
If wsUSD.Range("B2") = "" Then
wsUSD.Range("B2").Value = "MOVEMENT"
wsUSD.Range("C2").Value = "ACTION"
wsUSD.Range("D2").Value = "ADAPTIV CODE"
wsUSD.Range("E2").Value = "NOTIONAL"
wsUSD.Range("F2").Value = "CUSIP"
End If
'formatsheet and trasfer data
iLastrowUSD = wsUSD.Cells(3000, 4).End(xlUp).Row
'Transfer BMO plege
If UsdND.Range("A3") <> "" Then
    iLastrow = wsUSD.Cells(3000, 2).End(xlUp).Row
    
    Range(UsdND.Cells(3, 1), UsdND.Cells(3000, 2).End(xlUp)).Copy
    wsUSD.Range("D" & iLastrow + 1).PasteSpecial xlValues
    
    Range(UsdND.Cells(3, 4), UsdND.Cells(3000, 4).End(xlUp)).Copy
    wsUSD.Range("F" & iLastrow + 1).PasteSpecial xlValues
    
    iLastrow1 = wsUSD.Cells(3000, 4).End(xlUp).Row
    wsUSD.Range(Cells(iLastrow + 1, 2), Cells(iLastrow1, 2)) = "DELIVER"
    wsUSD.Range(Cells(iLastrow + 1, 3), Cells(iLastrow1, 3)) = "DELIVER"
End If

'Transfer Counterparty return
If UsdND.Range("F3") <> "" Then
    iLastrow = wsUSD.Cells(3000, 2).End(xlUp).Row
    
    Range(UsdND.Cells(3, 6), UsdND.Cells(3000, 7).End(xlUp)).Copy
    wsUSD.Range("D" & iLastrow + 1).PasteSpecial xlValues
    
    Range(UsdND.Cells(3, 9), UsdND.Cells(3000, 9).End(xlUp)).Copy
    wsUSD.Range("F" & iLastrow + 1).PasteSpecial xlValues
    
    iLastrow1 = wsUSD.Cells(3000, 4).End(xlUp).Row
    wsUSD.Range(Cells(iLastrow + 1, 2), Cells(iLastrow1, 2)) = "RECEIVE"
    wsUSD.Range(Cells(iLastrow + 1, 3), Cells(iLastrow1, 3)) = "RETURN"
End If


'Transfer New counterparty pledges
If UsdND.Range("M3") <> "" Then
    iLastrow = wsUSD.Cells(3000, 2).End(xlUp).Row
    
    Range(UsdND.Cells(3, 13), UsdND.Cells(3000, 14).End(xlUp)).Copy
    wsUSD.Range("D" & iLastrow + 1).PasteSpecial xlValues
    
    Range(UsdND.Cells(3, 16), UsdND.Cells(3000, 16).End(xlUp)).Copy
    wsUSD.Range("F" & iLastrow + 1).PasteSpecial xlValues
    
    iLastrow1 = wsUSD.Cells(3000, 4).End(xlUp).Row
    wsUSD.Range(Cells(iLastrow + 1, 2), Cells(iLastrow1, 2)).Value = "RECEIVE"
    wsUSD.Range(Cells(iLastrow + 1, 3), Cells(iLastrow1, 3)).Value = "DELIVER"
End If


'Transfer BMO return
If UsdND.Range("R3") <> "" Then
    iLastrow = wsUSD.Cells(3000, 2).End(xlUp).Row
    
    Range(UsdND.Cells(3, 18), UsdND.Cells(3000, 19).End(xlUp)).Copy
    wsUSD.Range("D" & iLastrow + 1).PasteSpecial xlValues
    
    Range(UsdND.Cells(3, 21), UsdND.Cells(3000, 21).End(xlUp)).Copy
    wsUSD.Range("F" & iLastrow + 1).PasteSpecial xlValues
    
    iLastrow1 = wsUSD.Cells(3000, 4).End(xlUp).Row
    wsUSD.Range(Cells(iLastrow + 1, 2), Cells(iLastrow1, 2)) = "DELIVER"
    wsUSD.Range(Cells(iLastrow + 1, 3), Cells(iLastrow1, 3)) = "RETURN"
End If

'check for LGY

Dim Rng As Range
wsUSD.Activate

iLastrow = wsUSD.Cells(300, 4).End(xlUp).Row
For i = iLastrowUSD + 1 To iLastrow
    If wsUSD.Cells(i, 2) = "RECEIVE" Then
        strFind = wsUSD.Cells(i, 2).Offset(0, 2)
        count = Application.WorksheetFunction.CountIf(Inbound.Columns(3), strFind)
    iRow = Inbound.Columns(3).Find(what:=strFind).Row
    If wsUSD.Cells(i, 5) = Inbound.Cells(iRow, 7) And wsUSD.Cells(i, 6) = Inbound.Cells(iRow, 6) Then
 
        strLGY = Inbound.Columns(3).Find(what:=strFind).Offset(0, 11)
        If strLGY = "LGY" Then
         wsUSD.Cells(i, 4) = strFind + "LGY"
        End If
    Else
        ThisWorkbook.Activate
        Sheets(1).Activate
        Cells(iRow, 3).Activate
        For j = 2 To count
        iRow = Inbound.Columns(3).FindNext(After:=ActiveCell).Row
            If wsUSD.Cells(i, 5) = Inbound.Cells(iRow, 7) And wsUSD.Cells(i, 6) = Inbound.Cells(iRow, 6) Then
            strLGY = Inbound.Cells(iRow, 3).Offset(0, 11)
             j = count
            End If
        Next
         
        If strLGY = "LGY" Then
        wsUSD.Cells(i, 4) = strFind + "LGY"
        End If
    End If
    End If

    If wsUSD.Cells(i, 2) = "DELIVER" Then
        strFind = wsUSD.Cells(i, 2).Offset(0, 2)
        count = Application.WorksheetFunction.CountIf(Outbound.Columns(3), strFind)
        iRow = Outbound.Columns(3).Find(what:=strFind).Row
    If wsUSD.Cells(i, 5) = Outbound.Cells(iRow, 7) And wsUSD.Cells(i, 6) = Outbound.Cells(iRow, 6) Then
 
        strLGY = Outbound.Columns(3).Find(what:=strFind).Offset(0, 11)
        If strLGY = "LGY" Then
         wsUSD.Cells(i, 4) = strFind + "LGY"
        End If
    Else
        ThisWorkbook.Activate
        Sheets(2).Activate
        Cells(iRow, 3).Activate
        
        For j = 2 To count
            iRow = Outbound.Columns(3).FindNext(After:=ActiveCell).Row
            If wsUSD.Cells(i, 5) = Outbound.Cells(iRow, 7) And wsUSD.Cells(i, 6) = Outbound.Cells(iRow, 6) Then
                strLGY = Outbound.Cells(iRow, 3).Offset(0, 11)
                j = count
            End If
        Next
        
        If strLGY = "LGY" Then
        wsUSD.Cells(i, 4) = strFind + "LGY"
        End If
    End If
    End If
Next


Exit Sub
ErrHandler:
If Err.Description Like "*file name*" Then
MsgBox ("Please check if file path is set properly and try again")
Call DisplayResetFoldersForm
Else
MsgBox Err.Description
End If
End Sub
