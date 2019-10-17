Option Explicit

Sub StockMarketData()
    On Error GoTo errHandler
    
    Dim ws As Worksheet
    
    'variables for looping through raw data
    Dim CurrentTicker, PriorTicker As String
    Dim CurrentTicker_FirstTime As Boolean
    Dim CurrentTicker_ValAtOpen, CurrentTicker_ValAtClose, _
        CurrentTicker_YearlyChange_Nominal, CurrentTicker_YearlyChange_Percent, _
        CurrentTicker_DailyVolume, CurrentTicker_YearlyVolume As Variant
    Dim CurrentRow_RawData, LastRow_RawData As Long
    Dim LastRow_Totals As Long
    Dim NumTickers, PlacementInTotalsRow As Long
    
    'variables for calculating greatest values
    Dim PercentChange_SearchRange, TotalYearlyVolume_SearchRange As Range
    Dim Max_PercentChange_Row, Min_PercentChange_Row, Max_YearlyVolume_Row As Range
    Dim Max_PercentChange, Min_PercentChange, Max_YearlyVolume As Double
    Dim Max_PercentChange_Ticker, Min_PercentChange_Ticker, Max_YearlyVolume_Ticker As String
    
    'variables for applying conditional formating:
    Dim FormatRange As Range
    Dim GreaterThanZero As FormatCondition, LessThanZero As FormatCondition
    
        
    For Each ws In Worksheets
    
        'get last row of current worksheet
        LastRow_RawData = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'init PriorTicker to nothing
        PriorTicker = ""
        
        'counter used to track individual tickers from raw data
        NumTickers = 0
        
        'setup Column names for Totals Row
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"

        
        For CurrentRow_RawData = 2 To LastRow_RawData
        
            CurrentTicker = ws.Cells(CurrentRow_RawData, 1)
            
            'test if this is the first time a new ticker has appeared
            'if yes: increment the count of individual ticker symbols and set _FirstTime variable to true
            
            If CurrentTicker <> PriorTicker Then
                CurrentTicker_FirstTime = True
                NumTickers = NumTickers + 1
            Else
                CurrentTicker_FirstTime = False
            End If
            
            'if first time ticker has appeared, store the value at open and the initial daily volume
            'otherwise, add the current daily volume to the daily volume from the prior days
            
            If CurrentTicker_FirstTime = True Then
                CurrentTicker_ValAtOpen = ws.Cells(CurrentRow_RawData, 3)
                CurrentTicker_DailyVolume = ws.Cells(CurrentRow_RawData, 7)
            Else
                CurrentTicker_DailyVolume = CurrentTicker_DailyVolume + ws.Cells(CurrentRow_RawData, 7)
            End If
                        
            'test if the next row contains a new ticker
            '   if yes:
            '           we are on the last row of the current ticker
            '           store _ValAtClose,
            '           move the _DailyVolume total into the _YearlyVolume variable
            '           calculate_YearlyChange_nominal
            '           calulate _YearlyChange_percent:
            '               return 0 if any value is 0
            '               ROUND the ouput so that that range.find() can operate properly
            '           place resulting variable into the Totals row/column areas
            '
            '   if no: skip this section and continue looping through the current ticker raw data
            
            If (CurrentTicker <> ws.Cells(CurrentRow_RawData + 1, 1)) Then
                CurrentTicker_ValAtClose = ws.Cells(CurrentRow_RawData, 6)
                CurrentTicker_YearlyVolume = CurrentTicker_DailyVolume
                CurrentTicker_YearlyChange_Nominal = CurrentTicker_ValAtClose - CurrentTicker_ValAtOpen
                
                If (CurrentTicker_ValAtClose = 0) Or (CurrentTicker_ValAtOpen = 0) Then
                    CurrentTicker_YearlyChange_Percent = 0
                Else
                    CurrentTicker_YearlyChange_Percent = (CurrentTicker_ValAtClose - CurrentTicker_ValAtOpen) / CurrentTicker_ValAtOpen
                End If

                'the totals row/column areas have to begin on the second row, but each worksheet can have different raw data counts
                'so we calculate the row needed by taking the curent value of numTickers and adding 1 to it
                'this gives us the current placementInTotalsRow value, where we store the resulting variables
                
                PlacementInTotalsRow = NumTickers + 1
                ws.Cells(PlacementInTotalsRow, 9) = CurrentTicker
                ws.Cells(PlacementInTotalsRow, 10) = CurrentTicker_YearlyChange_Nominal
                ws.Cells(PlacementInTotalsRow, 11) = FormatPercent(CurrentTicker_YearlyChange_Percent)
                ws.Cells(PlacementInTotalsRow, 12) = CurrentTicker_YearlyVolume
                
                'this is the last row, so we reset our daily volume variables to 0, so they can start fresh on the next loop
                CurrentTicker_DailyVolume = 0
                CurrentTicker_YearlyVolume = 0
            End If
            
            'set Prior Ticker to the value of CurrentTicker, so we can track where we are in the raw data
            PriorTicker = CurrentTicker
        
        Next CurrentRow_RawData
        
        'at this point we have processed all of the raw data, and stored the resulting total variables into their area
        'now, we search through those values and pull out the greatest increase, and greatest decrecease, and greatest total volume
        'and then we store them in another section
        
        'Get Last Row of Totals column
        LastRow_Totals = ws.Cells(Rows.Count, 11).End(xlUp).Row
             
            'set the Search range to find the Greatest Percentage Increase and Greatest Percent Decrease
            Set PercentChange_SearchRange = ws.Range("K2:K" & LastRow_Totals)
                
                'get the values of the greatest percent change and the greatest percent decrease
                Max_PercentChange = Application.WorksheetFunction.Max(PercentChange_SearchRange)
                Min_PercentChange = Application.WorksheetFunction.Min(PercentChange_SearchRange)
                
                'search through the range to find the row number that the Max_PercentChange value is stored at
                'if a value is not found, VBA throws an error, so test for range object being not nothing before accessing the object
                
                Set Max_PercentChange_Row = ws.Range("K:K").Find(What:=Max_PercentChange, LookIn:=xlValues)
                    
                    If Not Max_PercentChange_Row Is Nothing Then
                        Max_PercentChange_Ticker = ws.Range("I" & Max_PercentChange_Row.Row).Value
                    End If
               
               'search through the range to find the row number that the min_PercentChange value is stored at
                'if a value is not found, VBA throws an error, so test for range object being not nothing before accessing the object
                               
                Set Min_PercentChange_Row = ws.Range("K:K").Find(What:=Min_PercentChange, LookIn:=xlValues)
                    
                    If Not Min_PercentChange_Row Is Nothing Then
                        Min_PercentChange_Ticker = ws.Range("I" & Min_PercentChange_Row.Row).Value
                    End If
                
                
            'set the Search range to find the Largest Yearly Total Volume
            Set TotalYearlyVolume_SearchRange = ws.Range("L2:L" & LastRow_Totals)
                
                'get the value of the Largest Yearly Total
                Max_YearlyVolume = Application.WorksheetFunction.Max(TotalYearlyVolume_SearchRange)
       
                'search through the range to find the row number that the Max_PercentChange value is stored at
                'if a value is not found, VBA throws an error, so test for range object being not nothing before accessing the object

                Set Max_YearlyVolume_Row = ws.Range("L:L").Find(What:=Max_YearlyVolume, LookIn:=xlValues)
                
                    If Not Max_YearlyVolume_Row Is Nothing Then
                        Max_YearlyVolume_Ticker = ws.Range("I" & Max_YearlyVolume_Row.Row).Value
                    End If
                
            
            'fill headers, columns & rows with greatest values
            ws.Range("N2") = "Greatest Percent Increase"
            ws.Range("N3") = "Greatest Percent Decrease"
            ws.Range("N4") = "Greatest Total Volume"

            ws.Range("O1") = "Ticker"
            ws.Range("O2") = Max_PercentChange_Ticker
            ws.Range("O3") = Min_PercentChange_Ticker
            ws.Range("O4") = Max_YearlyVolume_Ticker
    
            ws.Range("P1") = "Value"
            ws.Range("P2") = FormatPercent(Max_PercentChange)
            ws.Range("P3") = FormatPercent(Min_PercentChange)
            ws.Range("P4") = Max_YearlyVolume
            
            ws.Columns("A:P").AutoFit
             

                
        ' set conditional formatting on Column J (Yearly Change)
        
        'Select range on which conditional formatting is to be desired
        Set FormatRange = ws.Range("J2", "J" & LastRow_Totals)

            'To delete/clear any existing conditional formatting from the range
            FormatRange.FormatConditions.Delete

            'Defining and setting the criteria for each conditional format
            Set GreaterThanZero = FormatRange.FormatConditions.Add(xlCellValue, xlGreater, "=0")
            Set LessThanZero = FormatRange.FormatConditions.Add(xlCellValue, xlLess, "=0")

            'Set GreaterThanZero to green
            With GreaterThanZero
                .Interior.ColorIndex = 4
            End With

            'Set LessThanZero to red
            With LessThanZero
                .Interior.ColorIndex = 3
            End With
        
        ' initialize LastRow variables to 0 to start over with new workbook
        LastRow_RawData = 0
        LastRow_Totals = 0
       
    Next ws

    MsgBox ("Analysis Complete")

        
exitsub:
    Exit Sub

errHandler:
    MsgBox ("An Error Has Occured: " & vbCrLf & vbCrLf & _
                "Error Number: " & Err.Number & vbCrLf & _
                "Error Description: " & Err.Description & vbCrLf & _
                "Error Source: " & Err.Source)
    
    Resume exitsub
        
End Sub
        










