Option Explicit
' ======================================================
' Analysis of 53 years of rainfall by year and month.
' Coded by: Rick Henderson, Twitter: @rickhenderson
' ======================================================
' Given a worksheet with the following data:
' For example, containing daily rainfall readings every day for many years
'  if the date is in mm/dd/yyyy format in the Excel worksheet.
' | Date       | Rainfall |
' | 01/01/2016 | 7.21     |
'      ...           ...
' The subroutine calculate_rainfall_by_month() will calcluate the monthly totals of rainfall
' and output them to another worksheet by year and month in the following format:
'
' | Year | Jan | Feb | Mar |.........|
' | 2016 | 34.4| 2.3 | 7.3 |.........|
' | ...  | ...
'
' This module also contains a subroutine called calculate_yearly_rainfall() to
' calculate the total for each year of the data set.
'
' Both subroutines should also work when the data changes, even if some
' dates are missing.
'
' The sheet names are coded inside the sub so you will have to change those for re-use
' By using conditional formatting, this could potentially generate an Excel heatmap.
' ==============================================

Private Function getMonthFromString(stringDate As String, source_format As String) As String
    ' Silly string extraction for string dates
   
    If source_format = "dd/mm/yyyy" Then
        getMonthFromString = Mid(stringDate, 4, 2) ' Strings start at index 1, not 0
    ElseIf source_format = "mm/dd/yyyy" Then
        getMonthFromString = Left(stringDate, 2)
    End If

End Function

Private Function getYearFromString(stringDate As String, source_format As String) As String
 '   Return just the year
 ' Could add more flexibility here. Only handles strings where year is last 4 digits
     getYearFromString = Right(stringDate, 4)
End Function

Private Function getDayFromString(stringDate As String, source_format As String) As String
     ' Return just the day
     If source_format = "dd/mm/yyyy" Then
        getDayFromString = Left(stringDate, 2)
    ElseIf source_format = "mm/dd/yyyy" Then
        getDayFromString = Mid(stringDate, 4, 2) ' Strings start at index 1, not 0
    End If
End Function

Sub calculate_yearly_rainfall()

    Dim dataStartCell As Range
    Dim outputStartCell As Range
    Dim currentCell As Range
    Dim currentDate As String
    Dim currentMonth As String
    Dim currentYear As String
    Dim dataRange As Range
    Dim yearlySum As Double
    Dim previousMonth As String
    Dim previousYear As String
    Dim yearNum As Integer
    Dim monthList As String
    Dim outMonth As Byte
    Dim cell As Range
    
    Dim dateNum As Double
    Dim inputRangeSize As Integer
    
    ' Turn screen updating off so the program runs faster
    ' This stops the screen from flickering as each value is written to the screen
    Application.ScreenUpdating = False
    
    Call ClearYearlyResults
    
    ' Create a range variable to point to the start of the data
    Set dataStartCell = Worksheets("Given Data Format").Range("A1")
           
    ' Create a range variable to point to where the output should start
    Set outputStartCell = Worksheets("Yearly Rainfall").Range("A1")
    
    ' Select all the date values and name the selected range
    With dataStartCell
        Range(.Offset(1, 0), .End(xlDown)).Name = "dataRange"
    End With
           
    ' Create a variable to easily refer to the data range
    Set dataRange = Range("dataRange")
    
    ' Set the yearNum to 1 to represent the first year of data for reporting purposes
    yearNum = 1
    
    ' Initialize the variables used as counters
    dateNum = 0
    yearlySum = 0
    currentMonth = getMonthFromString(dataStartCell.Offset(1, 0).Value, "dd/mm/yyyy")
    previousMonth = getMonthFromString(dataStartCell.Offset(1, 0).Value, "dd/mm/yyyy")
    previousYear = getYearFromString(dataStartCell.Offset(1, 0).Value, "dd/mm/yyyy")
    
    ' Start the main loop
    With dataStartCell
    
        For Each cell In dataRange
            ' Go to the next row in the date set
            dateNum = dateNum + 1
            currentDate = cell.Value
            ' Get the current years and month - just string extraction
            currentYear = getYearFromString(currentDate, "dd/mm/yyyy")
            currentMonth = getMonthFromString(currentDate, "dd/mm/yyyy")
                             
                If currentYear = previousYear Then
                    ' Add the rainfall to the yearlySum
                    yearlySum = yearlySum + .Offset(dateNum, 1).Value
                Else
                    ' A new year has occurred
                    ' Output the current Year's rainfall
                    outputStartCell.Offset(yearNum, 0).Value = previousYear
                    outputStartCell.Offset(yearNum, 1).Value = yearlySum
                    yearNum = yearNum + 1
                    yearlySum = 0
                    previousYear = currentYear
                    'Debug.Print currentYear
                End If
        Next
        ' Need to output the row of actual data
        outputStartCell.Offset(yearNum, 0).Value = currentYear
        outputStartCell.Offset(yearNum, 1).Value = yearlySum
      
    End With

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub

Sub calculate_rainfall_by_month()

    Dim dataStartCell As Range
    Dim outputStartCell As Range
    Dim currentCell As Range
    Dim currentDate As String
    Dim currentDay As String
    Dim currentMonth As String
    Dim currentYear As String
    Dim currentRainfall As Double
    Dim dataRange As Range
    Dim yearlySum As Double
    Dim monthlySum As Double
    Dim previousMonth As String
    Dim previousYear As String
    Dim monthList As String
    Dim monthNum As Byte
    Dim yearNum As Byte
    
    Dim cell As Range

    ' Turn screen updating off so the program runs faster
    ' This stops the screen from flickering as each value is written to the screen
    Application.ScreenUpdating = False
    
    ' Clear the previous run of this sub by deleting all output
    ' using a user defined function (UDF)
    Call ClearMonthlyResults
    
    ' Create a range variable to point to the start of the data
    Set dataStartCell = Worksheets("Given Data Format").Range("A1")
           
    ' Create a range variable to point to where the output should start
    Set outputStartCell = Worksheets("Required Format").Range("A1")
    
    ' Select all the date values and name the selected range
    With dataStartCell
        Range(.Offset(1, 0), .End(xlDown)).Name = "dataRange"
    End With
           
    ' Create a variable to easily refer to the data range
    Set dataRange = Range("dataRange")
    
    ' Set the yearNum to 1 to represent the first year of data for reporting purposes
    yearNum = 1
    
    ' Initialize the variables used as counters
    yearlySum = 0
    monthlySum = 0
    monthNum = 1
    
    ' DatePart() works correctly for unambiguous dates, but doesn't pad leading zeroes
    previousMonth = month(dataStartCell.Offset(1, 0).Value)
    previousYear = Year(dataStartCell.Offset(1, 0).Value)
    
    ' Start the main loop
    With dataStartCell
    
        For Each cell In dataRange
            ' Get the current Date
            currentDate = cell.Value
            
            ' Get the current years and month - just string extraction
            currentDay = day(currentDate)
            currentYear = Year(currentDate)
            currentMonth = month(currentDate)
                       
            ' Current rainfall is one column over from the current cell in the loop
            currentRainfall = cell.Offset(0, 1).Value
            
                If currentMonth = previousMonth Then
                    ' Add the rainfall to the current monthly sum
                    monthlySum = monthlySum + currentRainfall
                    'monthString = monthString & "+" & currentRainfall
                    
                Else
                    ' A new month has occurred
                                       
                    ' Add the last day of the previous Month
                    ' Output the month's total rainfall
                    outputStartCell.Offset(yearNum, monthNum).Value = monthlySum
                                        
                    ' Reset the monthly sum to 0 and add the new first day of month rainfall
                    monthlySum = 0 + currentRainfall
                    
                    ' Increase the count to the next month
                    monthNum = monthNum + 1
                    
                End If
                
                If currentYear <> previousYear Then
                    ' A new year has occurred
                    
                    ' Add the December rainfall to the current monthly sum
                    monthlySum = monthlySum + currentRainfall
                    
                    ' If it was December, write the previous year in column A
                    If previousMonth = "12" Then
                        outputStartCell.Offset(yearNum, 0).Value = previousYear
                    End If
                    
                    ' Increase the count to the next month
                    'monthNum = monthNum + 1
                    
                    ' Output the current Year's December Rainfall
                    'outputStartCell.Offset(yearNum, monthNum).Value = monthlySum
                    
                    ' Set the previousMonth as the currentMonth
                    previousMonth = currentMonth
                    
                    yearlySum = 0
                    yearNum = yearNum + 1
                    
                    ' Add currentRainfall to the new January
                    monthlySum = 0 + currentRainfall
                    monthNum = 1
                    previousYear = currentYear
                    
                    'Debug.Print currentYear
                End If

                ' Move to next day on next pass through the For-Next loop
                ' Set the previousMonth as the currentMonth
                previousMonth = currentMonth
                previousYear = currentYear
                
        Next
        
        ' Output the final month total and the last year value in the first column
        outputStartCell.Offset(yearNum, monthNum) = monthlySum
        outputStartCell.Offset(yearNum, 0).Value = currentYear
        
    End With

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub

Sub ClearYearlyResults()
'
' ClearYearlyResults Macro

    Sheets("Yearly Rainfall").Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A2:B57").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

Sub ClearMonthlyResults()
'
' ClearMonthlyResults Macro

    Dim cell As Range
    
    Range("A3:C3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A3").Select

End Sub
