Attribute VB_Name = "Module1"
Sub LossData()
    
    'this is to filter data for further data quality review
    
    '1. define variables and sheets
    Dim Company As String
    Dim Source As Worksheet
    Dim Target As Worksheet
    Dim LastRow As Long
    Dim filtered_data As Range
    
    'data in Source will be filtered and pasted into Target
    Set Source = ThisWorkbook.Sheets("ppauto_pos98-07")
    Set Target = ThisWorkbook.Sheets("Data")
    
    ThisWorkbook.Sheets("Summary").Activate
    
    '2. find filter value and number of last row in Source
    Company = Range("CompanyName").Value
    LastRow = Source.Cells(Source.Rows.Count, 1).End(xlUp).Row
    
    '3. filter Source and select visible range
    Source.Range("A1:M" & LastRow).AutoFilter field:=2, Criteria1:=Company
    Set filtered_data = Source.Range("A2:M" & LastRow).SpecialCells(xlCellTypeVisible)
    
    '4. clear contents in Target before pasting filtered data
    Target.Cells.ClearContents
    filtered_data.Copy Destination:=Target.Range("A1")
    
    '5. close auto filter mode
    Source.AutoFilterMode = False
    
End Sub

Sub Loss_Triangle()

    'this is to create loss triangle for selected company
    
    '1. define variables
    Dim Development_Lag As Variant
    Dim Accident_Year As Variant
    Dim LastRow As Long
    Dim AY_ind As Integer
    Dim DevLag_ind As Integer
    Dim i As Long, row_id As Integer, col_id As Integer
    Dim triangle() As Variant
    Dim method_id As Integer
    
    Dim data As Worksheet
    Set data = ThisWorkbook.Sheets("Data")
    
    '2. activate sheet Summary and find number of rows, Development_Lag and Accident_Year
    ThisWorkbook.Sheets("Summary").Activate
        
    LastRow = data.Cells(data.Rows.Count, 5).End(xlUp).Row
    Development_Lag = Application.WorksheetFunction.Unique(data.Range("E1:E" & LastRow))
    Accident_Year = Application.WorksheetFunction.Unique(data.Range("C1:C" & LastRow))
    method_id = Range("Method_ID").Value
        
    '3. create column header for Development_Lag and row header for Accident_year
    Range("Accident_Year").Resize(10, 1).ClearContents
    Range("Development_Lag").Resize(1, 10).ClearContents
    
    Range("Accident_Year").Resize(UBound(Accident_Year, 1), 1).Value = Accident_Year
    Range("Development_Lag").Resize(1, UBound(Development_Lag, 1)).Value = Application.WorksheetFunction.Transpose(Development_Lag)
    
    '4. define a matrix for storing loss values, row labelled as Accident_Year, and column labelled as Development_Lag
    ReDim triangle(1 To UBound(Accident_Year, 1), 1 To UBound(Development_Lag, 1))
    For i = 1 To LastRow
        If data.Cells(i, 4).Value < 2008 Then
            row_id = Application.WorksheetFunction.Match(data.Cells(i, 3), Accident_Year, 0)
            col_id = Application.WorksheetFunction.Match(data.Cells(i, 5), Development_Lag, 0)
            
            triangle(row_id, col_id) = data.Cells(i, method_id).Value
        End If
    Next i
    
    '5. clear contents in triangle and paste matrix of triangle into defined cells
    Range("Triangle").Resize(10, 10).ClearContents
    Range("Triangle").Resize(UBound(Accident_Year, 1), UBound(Development_Lag, 1)).Value = triangle
    
End Sub

Sub LDF()

    Dim Age2AgeLDF() As Variant
    Dim Age2UltLDF() As Variant
    Dim rowNum As Integer
    Dim colNum As Integer
    Dim triangle() As Variant
    Dim Total As Double
    Dim Count As Long
    Dim colID As Integer
    Dim rowID As Integer
    Dim Product As Double
    Dim LastRow As Long, num_year As Integer, num_dev_lag As Integer
    
    ' variables for valumn weighted LDF calculation
    Dim Weighted_Age2AgeLDF() As Variant
    Dim Weighted_Age2UltLDF() As Variant
    Dim CurrentLoss As Double
    Dim PreviousLoss As Double
    
    Dim avg_prod As Double
    Dim weighted_prod As Double
        
    LastRow = ThisWorkbook.Sheets("Data").Cells(ThisWorkbook.Sheets("Data").Rows.Count, 2).End(xlUp).Row
    num_year = Application.WorksheetFunction.Count(Application.WorksheetFunction.Unique(ThisWorkbook.Sheets("Data").Range("C1:C" & LastRow)))
    num_dev_lag = Application.WorksheetFunction.Count(Application.WorksheetFunction.Unique(ThisWorkbook.Sheets("Data").Range("E1:E" & LastRow)))
        
    ThisWorkbook.Sheets("Summary").Activate
    'rowNum = Application.WorksheetFunction.Count(Range("Accident_Year").Resize(num_year, 1))
    'colNum = Application.WorksheetFunction.Count(Range("Development_Lag").Resize(1, num_dev_lag))
    
    rowNum = num_year
    colNum = num_dev_lag
    ReDim triangle(1 To rowNum, 1 To colNum)
    triangle = Range("Triangle").Resize(rowNum, colNum).Value
    
    ReDim Age2AgeLDF(1 To colNum - 1)
    ReDim Weighted_Age2AgeLDF(1 To colNum - 1)
        
    For colID = 1 To colNum - 1
        Total = 0
        Count = 0
        
        CurrentLoss = 0
        PreviousLoss = 0

        For rowID = 1 To rowNum
            
            
            If triangle(rowID, colID + 1) <> "" Then
                
                Total = Total + triangle(rowID, colID + 1) / triangle(rowID, colID)
                Count = Count + 1
                
                CurrentLoss = CurrentLoss + triangle(rowID, colID + 1)
                PreviousLoss = PreviousLoss + triangle(rowID, colID)
            End If
            
        Next rowID
        
        Age2AgeLDF(colID) = Total / Count
        Weighted_Age2AgeLDF(colID) = CurrentLoss / PreviousLoss
    Next colID
       
    Range("Age2AgeLDF").Resize(1, colNum - 1).Value = Age2AgeLDF
    Range("WeightedLDF").Resize(1, colNum - 1).Value = Weighted_Age2AgeLDF
    
    ReDim Age2UltLDF(1 To colNum - 1)
    ReDim Weighted_Age2UltLDF(1 To colNum - 1)
    
    avg_prod = 1
    weighted_prod = 1
    For colID = colNum - 1 To 1 Step -1
        avg_prod = avg_prod * Age2AgeLDF(colID)
        Age2UltLDF(colID) = avg_prod
        
        weighted_prod = weighted_prod * Weighted_Age2AgeLDF(colID)
        Weighted_Age2UltLDF(colID) = weighted_prod
    Next colID
    
    Range("Age2UltLDF").Resize(1, colNum - 1).Value = Age2UltLDF
    Range("WeightedCDF").Resize(1, colNum - 1).Value = Weighted_Age2UltLDF
    'Stop

End Sub

Sub Summary()

    Dim Accident_Year() As Variant
    Dim Earned_Premium() As Variant
    Dim Incurred_Losses() As Variant
    Dim Cum_Paid_Loss() As Variant
    
    Dim Ult_Loss_Incur() As Variant
    Dim Reserve_Incur() As Variant
    Dim IBNR_Incur() As Variant
    Dim Ult_Loss_Paid() As Variant
    Dim Reserve_Paid() As Variant
    Dim IBNR_Paid() As Variant
    
    Dim Age() As Variant
    Dim Age_Ult_LDF() As Variant
    Dim LastRow As Long
    Dim i As Integer, rowID As Long
    
    Dim data As Worksheet
    Dim ws As Worksheet
    
    Set data = ThisWorkbook.Sheets("Data")
    Set ws = ThisWorkbook.Sheets("Summary")
    
    'before writing data, clear data contents
    ws.Range("F67:K76").ClearContents
    
    LastRow = data.Cells(data.Rows.Count, 2).End(xlUp).Row
    
    Accident_Year = Application.WorksheetFunction.Sort(Application.WorksheetFunction.Unique(data.Range("C1:C" & LastRow)))
    ws.Range("F67:F76").Value = Accident_Year
    
    ReDim Earned_Premium(1 To UBound(Accident_Year, 1))
    ReDim Incurred_Losses(1 To UBound(Accident_Year, 1))
    ReDim Cum_Paid_Loss(1 To UBound(Accident_Year, 1))
    ReDim Age(1 To UBound(Accident_Year, 1))
    ReDim Age_Ult_LDF(1 To UBound(Accident_Year, 1))
    
    ReDim Ult_Loss_Incur(1 To UBound(Accident_Year, 1))
    ReDim Reserve_Incur(1 To UBound(Accident_Year, 1))
    ReDim IBNR_Incur(1 To UBound(Accident_Year, 1))
    
    ReDim Ult_Loss_Paid(1 To UBound(Accident_Year, 1))
    ReDim Reserve_Paid(1 To UBound(Accident_Year, 1))
    ReDim IBNR_Paid(1 To UBound(Accident_Year, 1))
    
    For i = 1 To UBound(Accident_Year, 1)
        For rowID = 1 To LastRow
            If data.Cells(rowID, 3).Value = Accident_Year(i, 1) And data.Cells(rowID, 4).Value = "2007" Then
                 Earned_Premium(i) = data.Cells(rowID, 9).Value
                 Incurred_Losses(i) = data.Cells(rowID, 6).Value
                 Cum_Paid_Loss(i) = data.Cells(rowID, 7).Value
                 Age(i) = 2008 - Accident_Year(i, 1)
                 
                 If Age(i) = 10 Then
                    Age_Ult_LDF(i) = 1
                 Else
                    Age_Ult_LDF(i) = ws.Range("WeightedCDF").Offset(0, Age(i) - 1).Value
                 End If
                 
                 Ult_Loss_Incur(i) = Incurred_Losses(i) * Age_Ult_LDF(i)
                 Reserve_Incur(i) = Ult_Loss_Incur(i) - Cum_Paid_Loss(i)
                 IBNR_Incur(i) = Ult_Loss_Incur(i) - Incurred_Losses(i)
                 
                 Ult_Loss_Paid(i) = Cum_Paid_Loss(i) * Age_Ult_LDF(i)
                 Reserve_Paid(i) = Ult_Loss_Paid(i) - Cum_Paid_Loss(i)
                 IBNR_Paid(i) = Ult_Loss_Paid(i) - Cum_Paid_Loss(i)
                 
            End If
        Next rowID
    Next i
    'Stop
    ws.Range("G67:G76").Value = Application.WorksheetFunction.Transpose(Earned_Premium)
    ws.Range("H67:H76").Value = Application.WorksheetFunction.Transpose(Incurred_Losses)
    ws.Range("I67:I76").Value = Application.WorksheetFunction.Transpose(Cum_Paid_Loss)
    ws.Range("J67:J76").Value = Application.WorksheetFunction.Transpose(Age)
    ws.Range("K67:K76").Value = Application.WorksheetFunction.Transpose(Age_Ult_LDF)
    
    ws.Range("L67:L76").Value = Application.WorksheetFunction.Transpose(Ult_Loss_Incur)
    ws.Range("M67:M76").Value = Application.WorksheetFunction.Transpose(Reserve_Incur)
    ws.Range("N67:N76").Value = Application.WorksheetFunction.Transpose(IBNR_Incur)
    
    ws.Range("P67:P76").Value = Application.WorksheetFunction.Transpose(Ult_Loss_Paid)
    ws.Range("Q67:Q76").Value = Application.WorksheetFunction.Transpose(Reserve_Paid)
    ws.Range("R67:R76").Value = Application.WorksheetFunction.Transpose(IBNR_Paid)


End Sub


Sub ClearContentPPauto()

    Worksheets("ppauto_pos98-07").Cells.ClearContents
    Worksheets("Data").Cells.ClearContents
    
End Sub
