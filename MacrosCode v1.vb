Sub FormatPayrollData()

    Columns("C:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Call RemoveUnwantedRows
    
    Call SplitFullNames
    
    Columns("H:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Call ComputePaymentData
    
    Call ObtainRoutingAndAccountNumbers
    
    Columns("J:J").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "ExternalID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "FirstName"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "LastName"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "RoutingNumber"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "AccountNumber"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "MonthlySubtotal"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "TotalTax"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "NetPayment"
    Range("I1").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    
    Dim FileName
    FileName = "Payroll_Company_" & Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2) & ".csv"
    
    ActiveWorkbook.SaveAs FileName:= _
        "C:\temp\Excel examples\" & FileName, FileFormat:=xlCSVUTF8, _
        CreateBackup:=False

End Sub



Private Sub ObtainRoutingAndAccountNumbers()
    
    Workbooks.Open ("C:\temp\Excel examples\Bank Accounts.xlsx")
    Range("A1").Select
    Selection.End(xlDown).Select
    Dim BALastRow
    BALastRow = ActiveCell.Row
    
    Workbooks(1).Sheets("Sheet1").Activate
    Range("A1").Select
    Selection.End(xlDown).Select
    
    Dim EmpLastRow
    EmpLastRow = ActiveCell.Row
    Range("A1").Select
    
    For i = 2 To EmpLastRow
        Dim EEID
        EEID = Range("A" & i).Value
        
        For j = 2 To BALastRow
            If EEID = Workbooks(2).Sheets("Sheet1").Range("A" & j).Value Then
                Range("D" & i).Value = Workbooks(2).Sheets("Sheet1").Range("B" & j).Value
                Range("E" & i).Value = Workbooks(2).Sheets("Sheet1").Range("C" & j).Value
                Exit For
            End If
        Next j
    Next i
    

    Workbooks(2).Close
    
End Sub



Private Sub ComputePaymentData()
    Range("A1").Select
    Selection.End(xlDown).Select
    
    Dim LastRow
    LastRow = ActiveCell.Row
    
    For i = 2 To LastRow
        YearlySubtotal = Range("F" & i).Value + Range("F" & i).Value * Range("J" & i).Value
        MonthlySubtotal = YearlySubtotal / 12
        TaxBracket = GetTaxBracket(YearlySubtotal, Year(Date))
        TaxOnMonthly = MonthlySubtotal * TaxBracket
        NetPayment = MonthlySubtotal - TaxOnMonthly
        
        Range("G" & i).Value = MonthlySubtotal
        Range("H" & i).Value = TaxOnMonthly
        Range("I" & i).Value = NetPayment
    Next i
    
End Sub



Private Sub SplitFullNames()
    Range("A1").Select
    Selection.End(xlDown).Select
    Dim LastRow
    LastRow = ActiveCell.Row
    
    Range("A2").Select

    For i = 2 To LastRow
        Range("C" & i).Value = GetLastName(Range("B" & i).Value)
        Range("B" & i).Value = GetFirstName(Range("B" & i).Value)
        
    Next i

End Sub



Private Sub RemoveUnwantedRows()
    Range("A1").Select
    Selection.End(xlDown).Select
    
    Do Until ActiveCell.Row = 1
        
        If Range("H" & ActiveCell.Row).Value <> "United States" Or Range("I" & ActiveCell.Row).Value <> "" Then
            Rows(ActiveCell.Row & ":" & ActiveCell.Row).Select
            Selection.Delete Shift:=xlUp
        End If
        
        ActiveCell.Offset(-1, 0).Select
    Loop

End Sub



Function GetFirstName(FullName)
    GetFirstName = Left(FullName, InStr(FullName, " ") - 1)
End Function



Function GetLastName(FullName)
    GetLastName = StrReverse(Left(StrReverse(FullName), InStr(StrReverse(FullName), " ") - 1))
End Function



Function GetTaxBracket(AnnualIncome, Yr)

    Dim TaxRate
    
    If Yr >= 2022 Then
        If AnnualIncome < 0 Then
            TaxRate = 0
        ElseIf AnnualIncome < 11001 Then
            TaxRate = 0.1
        ElseIf AnnualIncome < 44726 Then
            TaxRate = 0.12
        ElseIf AnnualIncome < 95376 Then
            TaxRate = 0.22
        ElseIf AnnualIncome < 182101 Then
            TaxRate = 0.24
        ElseIf AnnualIncome < 231251 Then
            TaxRate = 0.32
        ElseIf AnnualIncome < 578126 Then
            TaxRate = 0.35
        Else
            TaxRate = 0.37
        End If
    ElseIf Yr < 2022 And Yr > 2015 Then
        If AnnualIncome < 0 Then
            TaxRate = 0
        ElseIf AnnualIncome < 11001 Then
            TaxRate = 0.1
        ElseIf AnnualIncome < 44726 Then
            TaxRate = 0.12
        ElseIf AnnualIncome < 95376 Then
            TaxRate = 0.22
        ElseIf AnnualIncome < 182101 Then
            TaxRate = 0.24
        ElseIf AnnualIncome < 231251 Then
            TaxRate = 0.32
        ElseIf AnnualIncome < 578126 Then
            TaxRate = 0.35
        Else
            TaxRate = 0.37
        End If
    Else
        If AnnualIncome < 0 Then
            TaxRate = 0
        ElseIf AnnualIncome < 11001 Then
            TaxRate = 0.1
        ElseIf AnnualIncome < 44726 Then
            TaxRate = 0.12
        ElseIf AnnualIncome < 95376 Then
            TaxRate = 0.22
        ElseIf AnnualIncome < 182101 Then
            TaxRate = 0.24
        ElseIf AnnualIncome < 231251 Then
            TaxRate = 0.32
        ElseIf AnnualIncome < 578126 Then
            TaxRate = 0.35
        Else
            TaxRate = 0.37
        End If
    End If
    
    GetTaxBracket = TaxRate

End Function






