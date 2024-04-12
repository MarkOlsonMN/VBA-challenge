Attribute VB_Name = "Module_2_Challenge_VBA_Code"
Option Explicit

Sub HW2_ModifyStockTables()

    Dim ws  As Worksheet
    
    Dim outputRowCurrent As Long
    Dim dataRowBegin As Long
    Dim dataRowEnd As Long
    Dim row As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim volumeCurrent As Long
    Dim volumeTotal As LongLong
    
    For Each ws In Worksheets
        ws.Activate

        dataRowEnd = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

'd        MsgBox ("sheetName = [" + ws.Name + "]" + " " + vbLf + _
                "dataRowEnd = [" + Str(dataRowEnd) + "]")

        Dim dataRowCurrent As Long
        Dim currentValue As String
        
        dataRowCurrent = 2
        outputRowCurrent = 2
            
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
            
        Columns("J").NumberFormat = "0.00"
        Columns("K").NumberFormat = "0.00%"
        
        Columns("I").AutoFit
        Columns("J").AutoFit
        Columns("K").AutoFit
        Columns("L").AutoFit

        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"

        Columns("O").AutoFit
        Columns("P").AutoFit
        Columns("Q").AutoFit

        'MsgBox ("Pause to verify Autofit")
        Application.Wait Now + TimeValue("00:00:03")

        ' Loop until the cell in column 1 is not a String value
        While IsEmpty(Cells(dataRowCurrent, 1)) = False And IsNumeric(Cells(dataRowCurrent, 1)) = False
            
            currentValue = Cells(dataRowCurrent, 1).Value
'd            MsgBox "While loop begin - Value in Row " & dataRowCurrent & " is: " & currentValue
            
            volumeTotal = 0
            
            ticker = Cells(dataRowCurrent, 1).Value
            openPrice = Cells(dataRowCurrent, 3).Value
            
            Dim startRow As Long
            Dim startColumn As Long
            Dim searchValue As String
            Dim foundCell As Range
            Dim lastRow As Long
            Dim foundRow As Long
            Dim foundColumn As Long

            ' Set the starting row and column position
            startRow = dataRowCurrent ' Starting row
            startColumn = 1 ' Starting column (Column A)

            ' Get the string value in the starting cell
            searchValue = ws.Cells(startRow, startColumn).Value
        
            ' Find the last row with the same string value in the same column
            Set foundCell = ws.Columns(startColumn).Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        
            If Not foundCell Is Nothing Then
                lastRow = foundCell.row
'd                MsgBox "The last row with value '" & searchValue & "' in column " & Split(Cells(1, startColumn).Address, "$")(1) & " is: " & lastRow
                
                '%%%%%%%%%%%%%%%%
                foundRow = foundCell.row
                foundColumn = foundCell.Column
        
'd                MsgBox "Found cell is at Row: " & foundRow & ", Column: " & foundColumn
        
                ' Initialize a new Range object using the found row and column indexes
                Dim newRange As Range
                Set newRange = ws.Cells(foundRow, foundColumn)
                ' Or you can use Cells method directly
                ' Set newRange = ws.Cells(foundRow, foundColumn)
        
                ' Do something with the newRange object
                ' For example, change the background color of the found cell
                newRange.Interior.Color = RGB(255, 255, 0)

                Dim startingRow As Long
                Dim endingRow As Long
                Dim totalSum As Double
                Dim sumRange As Range
                
                ' Set your known starting and ending rows
                startingRow = 2 ' Update this with your actual starting row number
                startingRow = startRow ' Update this with your actual starting row number
                
                endingRow = lastRow
            
                ' Define the range of values in column 7 from starting row to ending row
                Set sumRange = ws.Range(ws.Cells(startingRow, 7), ws.Cells(endingRow, 7))
            
                ' Calculate the total sum using the WorksheetFunction.Sum method
                totalSum = WorksheetFunction.Sum(sumRange)
            
                ' Display the total sum
'd                MsgBox "The total sum of values in column 7 from row " & startingRow & " to row " & endingRow & " is: " & totalSum
            
                volumeTotal = totalSum
            
            Else
                MsgBox "Value '" & searchValue & "' not found in column " & Split(Cells(1, startColumn).Address, "$")(1)
            End If
            
            '---------------------------------------
            
            closePrice = Cells(lastRow, 6).Value

            yearlyChange = closePrice - openPrice
            percentChange = (yearlyChange / openPrice) '* 100
            
'd            MsgBox ("ticker        = [" + ticker + "]" + " " + vbLf + _
                    "    openPrice = [" + Str(openPrice) + "]" + " " + vbLf + _
                    "   closePrice = [" + Str(closePrice) + "]" + " " + vbLf + _
                    " yearlyChange = [" + Str(yearlyChange) + "]" + " " + vbLf + _
                    "percentChange = [" + Str(percentChange) + "]" + " " + vbLf + _
                    "  volumeTotal = [" + Str(volumeTotal) + "]")
            
            Cells(outputRowCurrent, 9).Value = ticker
            Cells(outputRowCurrent, 10).Value = yearlyChange
            Cells(outputRowCurrent, 11).Value = percentChange
            Cells(outputRowCurrent, 12).Value = volumeTotal

            '----------
            Dim cell As Range
            ' Specify the row and column where you want to apply conditional formatting
            Set cell = ws.Cells(outputRowCurrent, 10) ' Update with your row and column numbers
            ' Check if the cell contains a numeric value
            If IsNumeric(cell.Value) Then
                If cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0) ' Green color
                ElseIf cell.Value < 0 Then
                    cell.Interior.Color = RGB(255, 0, 0) ' Red color
                End If
            End If
            '----------
    '        Dim cell As Range
            ' Specify the row and column where you want to apply conditional formatting
            Set cell = ws.Cells(outputRowCurrent, 11) ' Update with your row and column numbers
            ' Check if the cell contains a numeric value
            If IsNumeric(cell.Value) Then
                If cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0) ' Green color
                ElseIf cell.Value < 0 Then
                    cell.Interior.Color = RGB(255, 0, 0) ' Red color
                End If
            End If
            '----------

            ' control the While loop's next Cells(row,column) to be checked
'd            MsgBox ("While loop end - bump dataRowCurrent to next ticker ..." + vbLf + _
                    "  before = [" + Str(dataRowCurrent) + "]" + vbLf + _
                    "  after  = [" + Str((endingRow + 1)) + "]")
            
            dataRowCurrent = (endingRow + 1)
        
            outputRowCurrent = outputRowCurrent + 1

        Wend

        Columns("I").AutoFit
        Columns("J").AutoFit
        Columns("K").AutoFit
        Columns("L").AutoFit

        '--------------------------------------------------------------
        Dim summaryVolumeValues As Range
        Dim summaryTickerNames As Range
        Dim MaxVolumeValue As LongLong
        Dim MaxVolumeTicker As String
        Dim outputRowStart As Integer

        outputRowStart = 2 '4 '2

        Set summaryVolumeValues = ws.Range(Cells(outputRowStart, 12), Cells(outputRowCurrent, 12))
        Set summaryTickerNames = ws.Range(Cells(outputRowStart, 9), Cells(outputRowCurrent, 9))

        MaxVolumeValue = Application.Max(summaryVolumeValues)
        MaxVolumeTicker = Application.WorksheetFunction.Index(summaryTickerNames, Application.WorksheetFunction.Match(MaxVolumeValue, summaryVolumeValues, 0))
        
        ' Output the maximum value and the row where it was found
'd        MsgBox "MaxVolumeValue = [" & MaxVolumeValue & "]" & vbCrLf & _
            "MaxVolumeTicker = [" & MaxVolumeTicker & "]"
        Range("P4").Value = MaxVolumeTicker
        Range("Q4").Value = MaxVolumeValue
        
        '--------------------------------------------------------------
        Dim summaryPercentChangeValues As Range
        Dim MaxPercentChangeValue As Double
        Dim MaxPercentChangeTicker As String

        Set summaryPercentChangeValues = ws.Range(Cells(outputRowStart, 11), Cells(outputRowCurrent, 11))

        MaxPercentChangeValue = Application.Max(summaryPercentChangeValues)
        MaxPercentChangeTicker = Application.WorksheetFunction.Index(summaryTickerNames, Application.WorksheetFunction.Match(MaxPercentChangeValue, summaryPercentChangeValues, 0))
        
        ' Output the maximum value and the row where it was found
'd        MsgBox "MaxPercentChangeValue = [" & MaxPercentChangeValue & "]" & vbCrLf & _
               "MaxPercentChangeTicker = [" & MaxPercentChangeTicker & "]"
        Range("P2").Value = MaxPercentChangeTicker
        Range("Q2").Value = MaxPercentChangeValue
        '--------------------------------------------------------------
        Dim MinPercentChangeValue As Double
        Dim MinPercentChangeTicker As String
        
        MinPercentChangeValue = Application.Min(summaryPercentChangeValues)
        MinPercentChangeTicker = Application.WorksheetFunction.Index(summaryTickerNames, Application.WorksheetFunction.Match(MinPercentChangeValue, summaryPercentChangeValues, 0))
        
        ' Output the maximum value and the row where it was found
'd        MsgBox "MinPercentChangeValue = [" & MinPercentChangeValue & "]" & vbCrLf & _
               "MinPercentChangeTicker = [" & MinPercentChangeTicker & "]"

        Range("P3").Value = MinPercentChangeTicker
        Range("Q3").Value = MinPercentChangeValue

        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        
        Range("Q2").Cells.Interior.Color = RGB(0, 255, 0) ' Green color
        Range("Q3").Cells.Interior.Color = RGB(255, 0, 0) ' Red color

        Columns("O").AutoFit
        Columns("P").AutoFit
        Columns("Q").AutoFit

    Next ws

End Sub


