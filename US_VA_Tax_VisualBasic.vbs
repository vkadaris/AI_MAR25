Sub CalculateTaxes()
    Dim ws As Worksheet
    Dim sheetName As String
    sheetName = "Tax Calculation"
    
    ' Check if sheet exists, create if not
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    End If
    
    ' Clear previous contents
    ws.Cells.Clear
    
    ' Headers
    ws.Cells(1, 1).Value = "Income"
    ws.Cells(1, 2).Value = "401K"
    ws.Cells(1, 3).Value = "TaxOn"
    ws.Cells(1, 4).Value = "US_Tax"
    ws.Cells(1, 5).Value = "VA_Tax"
    ws.Cells(1, 6).Value = "Total_Tax"
    ws.Cells(1, 7).Value = "Eff Rate"
    ws.Cells(1, 8).Value = "Marg Rate"
    
    ' Constants
    Const StartIncome As Double = 50000
    Const Increment As Double = 10000
    Const Max401K As Double = 23000 ' 2024 limit for married couple
    Const FedStandardDeduction As Double = 29200 ' 2024 standard deduction for married filing jointly
    Const VAStandardDeduction As Double = 16900 ' Virginia standard deduction for married filing jointly
    
    ' Variables
    Dim i As Integer, row As Integer
    Dim income As Double, taxableIncome As Double
    Dim federalTax As Double, vaTax As Double, totalTax As Double
    Dim effectiveTaxRate As Double, marginalTaxRate As Double
    
    row = 2
    
    ' Loop through income levels
    For i = 0 To 10
        income = StartIncome + (i * Increment)
        Dim contribution401K As Double
        contribution401K = WorksheetFunction.Min(income * 0.15, Max401K)
        taxableIncome = WorksheetFunction.Max(income - contribution401K - FedStandardDeduction, 0)
        
        federalTax = CalculateFederalTax(taxableIncome)
        vaTax = CalculateVATax(WorksheetFunction.Max(income - contribution401K - VAStandardDeduction, 0))
        totalTax = federalTax + vaTax
        
        effectiveTaxRate = Round((totalTax / income) * 100, 1)
        marginalTaxRate = CalculateFederalMarginalRate(taxableIncome)
        
        ' Output values
        ws.Cells(row, 1).Value = income
        ws.Cells(row, 2).Value = contribution401K
        ws.Cells(row, 3).Value = taxableIncome
        ws.Cells(row, 4).Value = federalTax
        ws.Cells(row, 5).Value = vaTax
        ws.Cells(row, 6).Value = totalTax
        ws.Cells(row, 7).Value = effectiveTaxRate & "%"
        ws.Cells(row, 8).Value = marginalTaxRate & "%"
        
        row = row + 1
    Next i
    
    ' Formatting
    ws.Columns("A:F").NumberFormat = "$#,##0"
    ws.Columns("G:H").NumberFormat = "0.0%"
    ws.Columns.AutoFit
End Sub

Function CalculateFederalTax(income As Double) As Double
    Dim tax As Double
    Dim brackets As Variant, rates As Variant
    
    brackets = Array(0, 23200, 70550, 167100, 262100, 418850, 628300)
    rates = Array(0.1, 0.12, 0.22, 0.24, 0.32, 0.35, 0.37)
    
    tax = ApplyTaxBrackets(income, brackets, rates)
    CalculateFederalTax = tax
End Function

Function CalculateVATax(income As Double) As Double
    Dim tax As Double
    Dim brackets As Variant, rates As Variant
    
    brackets = Array(0, 3000, 5000, 17000)
    rates = Array(0.02, 0.03, 0.05, 0.0575)
    
    tax = ApplyTaxBrackets(income, brackets, rates)
    CalculateVATax = tax
End Function

Function ApplyTaxBrackets(income As Double, brackets As Variant, rates As Variant) As Double
    Dim tax As Double, i As Integer
    
    For i = UBound(brackets) To 1 Step -1
        If income > brackets(i) Then
            tax = tax + (income - brackets(i)) * rates(i)
            income = brackets(i)
        End If
    Next i
    
    ApplyTaxBrackets = Round(tax, 0)
End Function

Function CalculateFederalMarginalRate(income As Double) As Double
    Dim brackets As Variant, rates As Variant
    
    brackets = Array(0, 23200, 70550, 167100, 262100, 418850, 628300)
    rates = Array(0.1, 0.12, 0.22, 0.24, 0.32, 0.35, 0.37)
    
    Dim i As Integer
    For i = UBound(brackets) To 1 Step -1
        If income > brackets(i) Then
            CalculateFederalMarginalRate = rates(i) * 100
            Exit Function
        End If
    Next i
    
    CalculateFederalMarginalRate = 10
End Function

