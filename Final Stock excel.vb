Sub Multiple_year_stock_data()

    Dim LastRow As Long
    Dim WorksheetName As String
    Dim Stock_Name As String
    Dim Stock_Total As Double
    Stock_Total = 0
    Dim Summary_Table_Row As Integer
    Dim First_Day As Double    
    Dim Find_Row As Double 
    
    
    For Each ws In Worksheets

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        WorksheetName = ws.Name

        First_Day = 2
        

    ws.Range("I1").EntireColumn.Insert
    ws.Cells(1,9).Value = "Ticker"

    ws.Range("J1").EntireColumn.Insert
    ws.Cells(1,10).Value = "Yearly Change"

    ws.Range("K1").EntireColumn.Insert
    ws.Cells(1,11).Value = "Percent Change"

    ws.Range("L1").EntireColumn.Insert
    ws.Cells(1,12).Value = "Total Stock Volume"

    ws.Range("O1").EntireColumn.Insert
    ws.Cells(2,15).Value = "Greatest % Increase"
    ws.Cells(3,15).Value = "Greatest % Decrease"
    ws.Cells(4,15).Value = "Greatese Total Volume"

    ws.Range("P1").EntireColumn.Insert
    ws.Cells(1,16).Value = "Ticker"

    ws.Range("Q1").EntireColumn.Insert
    ws.Cells(1,17).Value = "Value"

    Summary_Table_Row = 2

    For i = 2 To LastRow
    
        If ws.Cells(i+1,1).Value <> ws.Cells(i,1).Value Then  


            Stock_Name = ws.Cells(i,1).Value

            Stock_Total = Stock_Total + ws.Cells(i,7).Value


            ws.Range("I" & Summary_Table_Row).Value = Stock_Name
            ws.Range("L" & Summary_Table_Row).Value = Stock_Total
            
            ws.Range("J" & Summary_Table_Row).Value = ws.Cells(i,6).Value - ws.Cells(First_Day,3).Value
            ws.Range("K" & Summary_Table_Row).Value = (ws.Cells(i,6).Value - ws.Cells(First_Day,3).Value) / ws.Cells(First_Day,3).Value * 100 & "%"

            First_Day = i + 1 

            Summary_Table_Row = Summary_Table_Row +1  

            Stock_Total = 0

        Else 

            Stock_Total = Stock_Total + Cells(i,7).Value  

    End If

    Next i  

            ws.Cells(2,17).Value = WorksheetFunction.Max(ws.Range("K1:K" & LastRow))
            Find_Row = WorksheetFunction.Match(ws.Cells(2,17).Value,ws.Range("K1:K" & LastRow),0)
            ws.Cells(2,16).Value = ws.Cells(Find_Row,9).Value 
            ws.Range("Q2").NumberFormat = "0.00%"

            ws.Cells(3,17).Value = WorksheetFunction.Min(ws.Range("K1:K" & LastRow))
            Find_Row = WorksheetFunction.Match(ws.Cells(3,17).Value,ws.Range("K1:K" & LastRow),0)
            ws.Cells(3,16).Value = ws.Cells(Find_Row,9).Value 
            ws.Range("Q3").NumberFormat = "0.00%"

            ws.Cells(4,17).Value = WorksheetFunction.Max(ws.Range("L1:L" & LastRow))
            Find_Row = WorksheetFunction.Match(ws.Cells(4,17).Value,ws.Range("L1:L" & LastRow),0)
            ws.Cells(4,16).Value = ws.Cells(Find_Row,9).Value

         

            Worksheets("2018").Columns("A:R").AutoFit
            Worksheets("2019").Columns("A:R").AutoFit
            Worksheets("2020").Columns("A:R").AutoFit
            
        dataRowStart = 2
        dataRowEnd = 3001

        For j = dataRowStart To dataRowEnd
        
            If ws.Cells(j, 10) > 0 Then
            
                ws.Cells(j, 10).Interior.Color = vbGreen
                   
            Else
        
                ws.Cells(j, 10).Interior.Color = vbRed
            
        End If         
        Next j    
       
Next ws
End Sub

        