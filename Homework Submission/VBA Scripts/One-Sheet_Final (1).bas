Attribute VB_Name = "Module2"
Sub VBA_Homework()

' ------------PSEUDOCODE-----------

'1. Loop through one sheet to grab ticker symbol highligting by difference
    '1a. Dont hardcode total, but instead use "xldown" function to have floating end.
'2. Find difference between open price on first day of the year and close price on last day of the year
'3. Calculate % difference between the two
    '3a. Have positive changes highlight as green
    '3b. Have Negative changes highlight as red
    '3c. Create a popup for each
'4. Have running total of stock volume for each ticker symbol
    '4a. Find Ticker with the highest volume of changes
    '4b. Create Popup Message signaling which had the highest volume.
'5. After completing all of the above, make sure that it can run for all worksheets in the workbook.

'------------------------------------
    
    'Set Variable as Ticker Symbol
    Dim Ticker As String
    
    'Set Variable as Yearly Price Difference
    Dim diff_yr As Double
    
    'Set Variable as % Yearly Difference
    Dim diff_percent As Double
    
    'Set Variable Total Volume per Ticker and set to 0
    Dim Volume_Total As LongLong
    Volume_Total = 0
    
    'Set Variable Yearly close for calculations
    Dim close_yr As Double
    
    'Set Variable Yearly Open
    Dim open_yr As Double
            
    'Set Row Counter for Output Table
    Dim Output_Row As Long
    Output_Row = 3
    
    'Set Headers for Output Table
    Range("J2").Value = "Ticker"
    Range("K2").Value = "Yearly Change"
    Range("L2").Value = "Percent Change"
    Range("M2").Value = "Total Stock Volume"

    'Define Variable Yearly Open for calculations

    open_yr = Range("C2").Value
    
    'Set Variable for Last Row in each worksheet
    Last_Row1 = Range("A" & Rows.Count).End(xlUp).row
    
    'Loop Through all Ticker Symbols on first sheet
    For I = 2 To Last_Row1
    
        'Check if Ticker Symbol is the Same and set conditions if it is different
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                
            'Set Ticker Symbol and Place Ticker Symbol in Output Table
            Ticker = Cells(I, 1).Value
            Range("J" & Output_Row) = Ticker
            
            'Set Yearly Close for calculations
            close_yr = Cells(I, 6).Value
            
            'Calculate Yearly Change and place in output table
            diff_yr = close_yr - open_yr
            Range("K" & Output_Row).Value = diff_yr
            
            'Sum Final Stock Volume, Place Total Stock Volume in Output Table, Initialize Volume
            Volume_Total = Volume_Total + Cells(I, 7).Value
            Range("M" & Output_Row) = Volume_Total
            Volume_Total = 0
            
            'Calculate Yearly % Change and place in output table
            If open_yr > 0 Then
                diff_percent = diff_yr / open_yr
                Range("L" & Output_Row).Value = diff_percent
            Else
                diff_percent = 0
                Range("L" & Output_Row).Value = diff_percent
            End If
            
            'Add 1 to the output row to ensure proper placement of next ticker
            Output_Row = Output_Row + 1
            
            'Set New Yearly Open
            open_yr = Cells(I + 1, 3).Value
            
        'Set Conditions if Ticker Symbol is the same
        Else
        
            'Running Sum for Volume
            Volume_Total = Volume_Total + Cells(I, 7).Value
            
        'End If Statements
        End If
        
    'Continue for Following Row
    Next I

'------------Find Min and Max Values-------------

'Set Last Row for Summary Table
Last_Row2 = Range("K" & Rows.Count).End(xlUp).row

'Create Table
Range("O3").Value = "Greatest % Increase"
Range("O4").Value = "Greatest % Decrease"
Range("O5").Value = "Greatest Total Volume"
Range("P2").Value = "Ticker"
Range("Q2").Value = "Value"

'Find and Return Greatest % Increase

Dim Max_percent As Double
Dim Max_Ticker As Long

    'Find Max Percent with VBA Function
    Max_percent = Application.WorksheetFunction.Max(Range("L3:L" & Last_Row2))
    Range("Q3").Value = Max_percent
    
    'Return Ticker Symbol of Highest Percent Return
    Max_Ticker = 3
    For q = 3 To Last_Row2
        If Cells(q, 12).Value = Max_percent Then Exit For
        If Cells(q, 12).Value <> Max_percent Then
            Max_Ticker = Max_Ticker + 1
        End If
    Next q
    Cells(3, 16).Value = Cells(Max_Ticker, 10).Value
    
'Find and Return Greatest % Decrease

Dim Min_percent As Double
Dim Min_Ticker As Long

    'Find Min Percent with VBA Function
    Min_percent = Application.WorksheetFunction.Min(Range("L3:L" & Last_Row2))
    Range("Q4").Value = Min_percent
    
    'Return Ticker Symbol of  Minimum Return
    Min_Ticker = 3
    For r = 3 To Last_Row2
        If Cells(r, 12).Value = Min_percent Then Exit For
        If Cells(r, 12).Value <> Min_percent Then
            Min_Ticker = Min_Ticker + 1
        End If
    Next r
    Cells(4, 16).Value = Cells(Min_Ticker, 10).Value
    
Dim Max_Volume As LongLong
Dim MaxV_Ticker As Long

    'Find Max Volume Traded with VBA Function
    Max_Volume = Application.WorksheetFunction.Max(Range("M3:M" & Last_Row2))
    Range("Q5").Value = Max_Volume
    
    'Return Ticker Symbol of Highest Volume Traded
    MaxV_Ticker = 3
    For s = 3 To Last_Row2
        If Cells(s, 13).Value = Max_Volume Then Exit For
        If Cells(s, 13).Value <> Max_Volume Then
            MaxV_Ticker = MaxV_Ticker + 1
        End If
    Next s
    Cells(5, 16).Value = Cells(MaxV_Ticker, 10).Value
    
    
'----------------Cell Formatting-----------------

Dim row As Integer
Dim column As Integer

'Conditional Formatting with Positive Numbers in Green and Negative Numbers in Red and No Change remaining white

For row = 1 To Last_Row2

    For column = 11 To 12
    
        If Cells(row, column).Value > 0 Then
        
            Cells(row, column).Interior.ColorIndex = 4
            
        ElseIf Cells(row, column).Value < 0 Then
        
            Cells(row, column).Interior.ColorIndex = 3
            
        Else
        
            Cells(row, column).Interior.Color = RGB(255, 255, 255)
            
        End If
        
    Next column
    
Next row

'Format % Change column as Percentage
Columns("L").NumberFormat = "0.00%"
Range("Q3:Q4").NumberFormat = "0.00%"

'Format $ Change Amount as Currency
Columns("K").NumberFormat = "$#,##0.00"

'Bold Title of Output Chart
Range("J2:M2").Font.FontStyle = "Bold"
Range("O3:O5").Font.FontStyle = "Bold"
Range("P2:Q2").Font.FontStyle = "Bold"

'Format Volume to have commas
Columns("M").NumberFormat = "#,##0"
Range("Q5").NumberFormat = "#,##0"

'Change Width of Columns to Fit Maximum Value
Columns("K:Q").AutoFit

End Sub
