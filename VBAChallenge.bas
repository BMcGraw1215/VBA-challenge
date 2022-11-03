'for use with 'alphabetical_testing.xlsx' or 'Multiple_year_stock_data.xlsx'

Sub StockSort()

For Sheet = 1 To ThisWorkbook.Worksheets.Count
Worksheets(Sheet).Activate

'Establishes most variables.

Dim TickerCount As Integer                                   ' Keeps Count of the Rows in the new table we're trying to create.
Dim YROpeningVal                                             ' Tracks Unique Ticker Opening Values
Dim YRClosingVal                                             ' Tracks Unique Ticker Closing Values
Dim EndofTable As Long                                       ' This Variable helps tell the code when to stop, given any data set.

YROpeningVal = 0
YRClosingVal = 0
TickerCount = 2
EndofTable = Range("A1").End(xlDown).Row                        ' Searches for last value located Column A.

'Adds a temporary associable end to the table.

If Cells(EndofTable, 1).Value <> 0 Then            ' Makes sure that we only add the End of Table once.
    For Col = 1 To 7
        Cells(EndofTable + 1, Col).Value = 0        'Ensures the last ticker also gets calculated.
    Next Col
End If

'Runs through the initial table, creating a new one based on parameters set.

For Row = 2 To (EndofTable + 1)                                 ' Scans whole Table up until EndofTable.

    If IsNumeric(Cells(Row, 7).Value) = True Then               ' Makes sure SumVolume cannot be a string(Header)
        SumVolume = SumVolume + Cells(Row - 1, 7).Value         ' keeps adding onto SumVolume with Every Row.
    End If
       
    If Cells(Row, 1).Value <> Cells(Row - 1, 1).Value = True Then              ' Checks to see if we're at a new ticker or not, by comparing to previous row.
           
        If Cells(Row, 1).Value <> 0 Then                        ' Keeps End of Table from populating in Column I
            Cells(TickerCount, 9).Value = Cells(Row, 1).Value   ' Fills Column I with Unique Ticker Values
        End If
           
        If IsNumeric(Cells(Row - 1, 6).Value) = True Then       ' Makes sure our YRClosingVal cannot be a string(Header)
            YRClosingVal = Cells(Row - 1, 6).Value              ' Sets our Yearly Closing Value variable equal to the previous Ticker's most recent Closing Value
            Cells(TickerCount - 1, 10).Value = YRClosingVal - YROpeningVal      ' Calculates the difference in Closing Value and Opening Value. Places Value in Column J.
            Cells(TickerCount - 1, 11).Value = YRClosingVal / YROpeningVal - 1     ' Calculates a % value between Closing Value and Opening Value. Places Value in Column K.
            Cells(TickerCount - 1, 12).Value = SumVolume        ' Adds SumVolume to our new table in Column L.
        End If
           
        SumVolume = 0                                           ' Resets our sum for Volume since we've reached a new Ticker.
        TickerCount = TickerCount + 1                           ' Increments TickerCount
        YROpeningVal = Cells(Row, 3).Value                      ' Since We've reached a new Ticker, Updates the Opening Value.
    End If

Next Row

' Cleans up the End of the original Table.
   
For Col = 1 To 7
    Cells(EndofTable + 1, Col).Clear
Next Col

'Sets the Headers for the new table.

Cells(1, 9).Value = "Tickers"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Changes Column Width to be more appropriate

Range("I:I").ColumnWidth = 9
Range("J:K").ColumnWidth = 15
Range("L:L").ColumnWidth = 19

Range("K:K").NumberFormat = "0.00%"
Range("L:L").NumberFormat = "###,###,###,###,###,###"
'Sets Conditional Formatting For Column J

Dim LTFormat As FormatCondition                                  ' For the LessThan 0 Values
Dim GTFormat As FormatCondition                                  ' For the GreaterThan 0 Values
Dim YearlyValues As Range
Set YearlyValues = Range("J2:J" & TickerCount)

YearlyValues.FormatConditions.Delete
Set LTFormat = YearlyValues.FormatConditions.Add(xlCellValue, xlLess, "0")
Set GTFormat = YearlyValues.FormatConditions.Add(xlCellValue, xlGreater, "0")

With LTFormat
    .Interior.Color = RGB(255, 180, 180)
End With
With GTFormat
    .Interior.Color = RGB(180, 255, 180)
End With
   
'Resets Variables
SumVolume = Empty
TickerCount = Empty
YROpeningVal = Empty
YRClosingVal = Empty
EndofTable = Empty
   

' BONUS


'Sets all Variables needed while reading through Table2
Dim GI As Double, GD As Double, GV As Double
Dim GIT As String, GDT As String, GVT As String
   
Dim EndofTable2 As Long
EndofTable2 = Range("I1").End(xlDown).Row

'Reads Table2 for Values needed.
For Table2Row = 2 To EndofTable2
    If Cells(Table2Row, 11).Value > GI Then
    GI = Cells(Table2Row, 11).Value
    GIT = Cells(Table2Row, 9).Value
    End If
    
    If Cells(Table2Row, 11).Value < GD Then
    GD = Cells(Table2Row, 11).Value
    GDT = Cells(Table2Row, 9).Value
    End If
    
    If Cells(Table2Row, 12).Value > GV Then
    GV = Cells(Table2Row, 12).Value
    GVT = Cells(Table2Row, 9).Value
    End If
Next Table2Row

'Places all Values
Cells(2, 16).Value = GIT
Cells(3, 16).Value = GDT
Cells(4, 16).Value = GVT

Cells(2, 17).Value = GI
Cells(3, 17).Value = GD
Cells(4, 17).Value = GV

' Set Headers for both Rows and Columns of the 3rd table.
   
' Row Headers
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
   
' Column Headers
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

' Change Column Widths for 3rd Table
Range("O:O").ColumnWidth = 20
Range("P:Q").ColumnWidth = 17

'Formats Values
Range("Q2:Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = "###,###,###,###,###,###"

'Resets Variables
GI = Empty
GD = Empty
GV = Empty
GIT = Empty
GDT = Empty
GVT = Empty

Next Sheet

MsgBox ("Sub Complete!")

End Sub
