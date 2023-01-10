Attribute VB_Name = "Module1"
Sub stocks()
    
    Dim ticker() As String
    Dim percentChange, greatestIncrease, greatestDecrease, yearlyChange, openValue, closeValue As Double
    Dim volume, greatestVolume, openDate, closeDate As Long
    Dim ws As Worksheet
    Dim rowCount, tickerCount, maxRow, minRow, volumeRow As Integer
    Dim volumeRange As String
    
    For Each ws In ThisWorkbook.Worksheets
                   
        ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("L1"), Unique:=True
    
        'get unique ticker count
        tickerCount = ws.Cells(Rows.Count, "L").End(xlUp).Row - 1
        ReDim ticker(1 To tickerCount) As String
        
        'get ticker and store in an array
        For i = 1 To tickerCount
           ticker(i) = ws.Cells(i + 1, 12).Value
        Next i
        
        'set headers
        ws.Range("L1").Value = "<ticker>"
        ws.Range("M1").Value = "Yearly Change"
        ws.Range("N1").Value = "Percentage Change"
        ws.Range("O1").Value = "Total Stock Volume"
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        ws.Range("N:N").NumberFormat = "0.00%"
        ws.Range("T2:T3").NumberFormat = "0.00%"
        
        
        
        
        
        
        'loop through each value in the array
        For i = 1 To tickerCount
            volume = 0
            
            openDate = ws.Range("A:A").Find(What:=ticker(i), LookAt:=xlWhole).Row
            closeDate = ws.Range("A:A").Find(What:=ticker(i), LookAt:=xlWhole, SearchDirection:=xlPrevious).Row
            
            volumeRange = "G" & openDate & ":G" & closeDate
            volume = Application.WorksheetFunction.Sum(ws.Range(volumeRange))
            
            ws.Cells(i + 1, 15) = volume
            
            openValue = ws.Cells(openDate, 3).Value
            closeValue = ws.Cells(closeDate, 6).Value
            yearlyChange = closeValue - openValue
            ws.Cells(i + 1, 13) = yearlyChange

                If yearlyChange > 0 Then
                        ws.Cells(i + 1, 13).Interior.Color = vbGreen
                    Else
                        ws.Cells(i + 1, 13).Interior.Color = vbRed
                End If


            percentChange = (yearlyChange) / openValue
            ws.Cells(i + 1, 14) = percentChange
                               
        Next i

        'get max/min/total for increase, decrease, and volume
        greatestVolume = WorksheetFunction.Max(ws.Range("O:O"))
        volumeRow = WorksheetFunction.Match(greatestVolume, ws.Range("O:O"), 0)
        ws.Cells(4, 19).Value = ws.Cells(volumeRow, 12).Value
        ws.Cells(4, 20).Value = greatestVolume

        greatestIncrease = WorksheetFunction.Max(ws.Range("N:N"))
        maxRow = WorksheetFunction.Match(greatestIncrease, ws.Range("N:N"), 0)
        ws.Cells(2, 19).Value = ws.Cells(maxRow, 12).Value
        ws.Cells(2, 20).Value = greatestIncrease

        greatestDecrease = WorksheetFunction.Min(ws.Range("N:N"))
        minRow = WorksheetFunction.Match(greatestDecrease, ws.Range("N:N"), 0)
        ws.Cells(3, 19).Value = ws.Cells(minRow, 12).Value
        ws.Cells(3, 20).Value = greatestDecrease
        ws.UsedRange.EntireColumn.AutoFit


    Next ws
End Sub
