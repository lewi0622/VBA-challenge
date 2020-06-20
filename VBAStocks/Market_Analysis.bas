Attribute VB_Name = "Module1"
Option Explicit
Sub Market_Analysis()
'********************** Timer and Optimzation Setup **********************
    'Uncomment the MsgBox at the end of code to get total time to run macro
    'Run time for alphabetical_testing is approx 1.0 sec
    'Run time for multiple_year is approx 5.5 sec

    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    'Remember time when macro starts
    StartTime = Timer

    'Optimization Start
    Dim CalcState As Long
    Dim EventState As Boolean
    Dim PageBreakState As Boolean
    Application.ScreenUpdating = False
    
    EventState = Application.EnableEvents
    Application.EnableEvents = False
    
    CalcState = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    PageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
'********************** Constant Declaration **********************
    Const start_row = 2
    Const start_col = 1
    Const end_col = 7
    Const table_start_col = 9
    Const table_end_col = 12
    Const greatest_start_col = 15
    Const greatest_end_col = 17
    Const greatest_end_row = 4
    
'********************** Variable Declaration **********************
    Dim end_row As Long
    Dim sheet_vals As Variant
    Dim ws As Worksheet
    Dim i As Long
    Dim lower_bound As Long
    Dim upper_bound As Long
    Dim current_ticker As String
    Dim table_row As Integer
    Dim year_open As Double
    Dim total_volume As Variant
    Dim output_table() As Variant
    Dim greatest_table(2, 1) As Variant
    
    For Each ws In Worksheets
        With ws
'********************** Output Table **********************
            'Find end row
            end_row = .Cells(Rows.Count, 1).End(xlUp).Row
            'Read in all values
            sheet_vals = .Range(.Cells(start_row, start_col), .Cells(end_row, end_col)).Value
            lower_bound = LBound(sheet_vals, 1)
            upper_bound = UBound(sheet_vals, 1)
            
            'Initialize Sheet Variables
            current_ticker = ""
            table_row = 0
            year_open = -1
            total_volume = -1
            ReDim output_table(3, table_row)
            
            'Clear sheet for outputs
            .Range(.Columns(table_start_col), .Columns(greatest_end_col)).Clear
            
            'Create Table Header
            .Range(.Cells(1, table_start_col), .Cells(1, table_end_col)).Value = Split("Ticker,Yearly Change,Percent Change,Total Stock Volume", ",")
            
            'Create Greatest Table
            .Range(.Cells(1, 16), .Cells(1, 17)).Value = Split("Ticker,Value", ",")
            .Range(.Cells(2, 15), .Cells(4, 15)).Value = Application.Transpose(Split("Greatest % Increase,Greatest % Decrease,Greatest Total Volume", ","))
            
            For i = lower_bound To upper_bound
                'Check if next ticker is different from current_ticker
                If sheet_vals(i, start_col) <> current_ticker Or i = upper_bound Then
                    'check final case and tidy up
                    If i = upper_bound Then
                        'Sum total_vol
                        total_volume = total_volume + sheet_vals(i, end_col)
                    End If
                    
                    'Not initial case
                    If current_ticker <> "" Then
                        'Write to output
                        output_table(0, table_row) = current_ticker
                        output_table(1, table_row) = sheet_vals(i - 1, 6) - year_open
                        
                        'Handle divide by zero case
                        If year_open <> 0 Then
                            output_table(2, table_row) = output_table(1, table_row) / year_open
                        Else
                            output_table(2, table_row) = 0
                        End If
                        output_table(3, table_row) = total_volume
                        
                        'inc table row
                        table_row = table_row + 1
                        
                        'resize array to make room for next vals
                        ReDim Preserve output_table(3, table_row) As Variant
                    End If
                    'Get new ticker values
                    current_ticker = sheet_vals(i, start_col)
                    year_open = sheet_vals(i, 3)
                    total_volume = sheet_vals(i, end_col)
                Else
                    'Sum total_vol
                    total_volume = total_volume + sheet_vals(i, end_col)
                    
                End If
                
            Next i
            
            'Write output table to sheet
            .Range(.Cells(start_row, table_start_col), .Cells(UBound(output_table, 2) + 1, table_end_col)).Value = Application.Transpose(output_table)
            
'********************** Output Table Formatting **********************
            'Apply conditional formatting
            Dim greater_cond As FormatCondition
            Dim less_cond As FormatCondition
            Set greater_cond = .Range("J:J").FormatConditions.Add(xlCellValue, xlGreater, 0)
            Set less_cond = .Range("J:J").FormatConditions.Add(xlCellValue, xlLess, 0)
            'Remove formatting from first row
            .Rows(1).FormatConditions.Delete
            greater_cond.Interior.Color = vbGreen
            less_cond.Interior.Color = vbRed
            
            'Apply formatting to table
            .Columns(table_start_col).NumberFormat = "@"
            .Columns(table_start_col + 1).NumberFormat = "0.00"
            .Columns(table_start_col + 2).NumberFormat = "0.00%"
            
'********************** Greatest Table **********************
            For i = 0 To UBound(output_table, 2)
                'Initial Case
                If i = 0 Then
                    '% change
                    greatest_table(0, 1) = output_table(2, i)
                    greatest_table(1, 1) = output_table(2, i)
                    'Volume
                    greatest_table(2, 1) = output_table(3, i)
                Else
                    'Greater %
                    If output_table(2, i) > greatest_table(0, 1) Then
                        greatest_table(0, 1) = output_table(2, i)
                        greatest_table(0, 0) = output_table(0, i)
                    End If
                    'Least %
                    If output_table(2, i) < greatest_table(1, 1) Then
                        greatest_table(1, 1) = output_table(2, i)
                        greatest_table(1, 0) = output_table(0, i)
                    End If
                    'Volume
                    If output_table(3, i) > greatest_table(2, 1) Then
                        greatest_table(2, 1) = output_table(3, i)
                        greatest_table(2, 0) = output_table(0, i)
                    End If
                End If
            Next i
            
            'Write to sheet
            .Range(.Cells(start_row, greatest_start_col + 1), .Cells(greatest_end_row, greatest_end_col)).Value = greatest_table
                        
'********************** Output Table Formatting **********************
            .Range(.Cells(start_row, greatest_end_col), .Cells(start_row + 1, greatest_end_col)).NumberFormat = "0.00%"
            
'********************** General Formatting **********************
            'AutoFit all created table columns
            .Columns("I:Q").AutoFit
        End With
    Next ws
    
'********************** Timer and Optimization Cleanup **********************
    'Optimization Stop
    ActiveSheet.DisplayPageBreaks = PageBreakState
    Application.Calculation = CalcState
    Application.EnableEvents = EventState
    Application.ScreenUpdating = True
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Uncomment next line to see total macro run time
    'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
End Sub
