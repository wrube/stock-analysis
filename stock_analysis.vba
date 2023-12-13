Attribute VB_Name = "Module1"
Sub testing()
    Dim activeSheet As Worksheet
    Dim rowCounter As Long
    Dim out_row_counter As Integer
    Dim ticker_counter As Long
    Dim last_row_num As Long
    Dim ticker_ref As String
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim vol As Long
    Dim results_origin_cell As Range
    Dim ticker_data_rng As Range
    Dim results_column_num_start As Integer: results_column_num_start = 11
    Dim ticker_start_row As Long
    Dim ticker_end_row As Long
    Dim ticker_yr_start_price As Double
    Dim ticker_yr_end_price As Double
    Dim yr_change As Double
    Dim yr_perc_change As Double
    Dim total_volume As Variant
    Dim greatest_perc_inc As Double
    Dim greatest_perc_dec As Double
    Dim greatest_vol_change As LongLong
    Dim results_ticker_end_row As Long
    Dim ticker_greatest_inc As String
    Dim ticker_greatest_dec As String
    Dim ticker_greatest_vol As String
    Dim summary_table As Range
    
    
    
    'Iterate through each worksheet ref: https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

    For Each activeSheet In Worksheets
    
        Set results_origin_cell = activeSheet.Cells(1, results_column_num_start)

        ticker_ref = activeSheet.Range("A2").Value
        
        'results table row starting number
        out_row_counter = 1
        
        'set the first row index for first ticker at line 2
        ticker_start_row = 2
        
        Call set_column_titles(results_origin_cell)
        
        last_row_num = find_last_row_num(activeSheet, 1)
        
        For rowCounter = 2 To last_row_num + 1
            ticker = activeSheet.Cells(rowCounter, 1).Value


            If ticker_ref <> ticker Then
                ' store the previous row number as the end range for previous ticker
                ticker_end_row = rowCounter - 1
                
                ' start running calculations on ticker data
                ' Set the ticker table range
                Set ticker_data_rng = activeSheet.Range("A" & ticker_start_row & ":" & "G" & ticker_end_row)

                ' sort the previous ticker columns by date to ensure first ticker rows are in date order
                ticker_data_rng.Sort Key1:=activeSheet.Range("B" & ticker_start_row)
                
                ' calculate the year change in price for ticker
                ticker_yr_start_price = activeSheet.Range("C" & ticker_start_row).Value
                ticker_yr_end_price = activeSheet.Range("F" & ticker_end_row).Value
                yr_change = ticker_yr_end_price - ticker_yr_start_price
                
                'calculate the percentage change in price for ticker
                yr_perc_change = (ticker_yr_end_price - ticker_yr_start_price) / ticker_yr_start_price
                
                'calculate the total volume for ticker
                total_volume = WorksheetFunction.Sum(Range("G" & ticker_start_row & ":" & "G" & ticker_end_row))
            
                'post results in spreadsheet
                results_origin_cell.Offset(out_row_counter, 0).Value = ticker_ref
                results_origin_cell.Offset(out_row_counter, 1).Value = yr_change
                results_origin_cell.Offset(out_row_counter, 2).Value = Application.Round(yr_perc_change, 4)
                results_origin_cell.Offset(out_row_counter, 3).Value = total_volume
                
                
                out_row_counter = out_row_counter + 1
                ticker_start_row = rowCounter

            End If
            ticker_ref = ticker

        Next
        
        'Summarise the results
        greatest_perc_inc = WorksheetFunction.Max(activeSheet.Columns(results_column_num_start + 2))
        greatest_perc_dec = WorksheetFunction.Min(activeSheet.Columns(results_column_num_start + 2))
        greatest_vol_change = WorksheetFunction.Max(activeSheet.Columns(results_column_num_start + 3))
        results_ticker_end_row = find_last_row_num(activeSheet, results_column_num_start)
        Set summary_table = activeSheet.Range(results_origin_cell.Offset(1, 0), _
                results_origin_cell.Offset(results_ticker_end_row - 1, 3))
        ticker_greatest_inc = Application.XLookup(greatest_perc_inc, activeSheet.Columns(results_column_num_start + 2), _
                activeSheet.Columns(results_column_num_start), "None", 0, 1)
        ticker_greatest_dec = Application.XLookup(greatest_perc_dec, activeSheet.Columns(results_column_num_start + 2), _
                activeSheet.Columns(results_column_num_start), "None", 0, 1)
        ticker_greatest_vol = Application.XLookup(greatest_vol_change, activeSheet.Columns(results_column_num_start + 3), _
                activeSheet.Columns(results_column_num_start), "None", 0, 1)

        'post the results
        results_origin_cell.Offset(1, 7).Value = greatest_perc_inc
        results_origin_cell.Offset(1, 6).Value = ticker_greatest_inc
        results_origin_cell.Offset(2, 7).Value = greatest_perc_dec
        results_origin_cell.Offset(2, 6).Value = ticker_greatest_dec
        results_origin_cell.Offset(3, 7).Value = greatest_vol_change
        results_origin_cell.Offset(3, 6).Value = ticker_greatest_vol
        
        
        'Pretty up the results and summary
        Call fomat_yr_change_results(activeSheet, results_column_num_start + 1)
        activeSheet.Columns(results_column_num_start + 2).NumberFormat = "0.00%"
        activeSheet.Range(results_origin_cell.Offset(1, 7), results_origin_cell.Offset(2, 7)).NumberFormat = "0.00%"
        activeSheet.Range("H:Z").Columns.AutoFit
        

    Next
End Sub

Private Sub set_column_titles(origin As Range)

    
    origin.Value = "Ticker"
    origin.Offset(, 1).Value = "Yearly Change"
    origin.Offset(, 2).Value = "Percent Change"
    origin.Offset(, 3).Value = "Total Stock Volume"
    
    origin.Offset(, 6).Value = "Ticker"
    origin.Offset(, 7).Value = "Value"
    origin.Offset(1, 5).Value = "Greatest % Increase"
    origin.Offset(2, 5).Value = "Greatest % Decrease"
    origin.Offset(3, 5).Value = "Greatest Total Volume"

End Sub

Private Function find_last_row_num(sheet As Worksheet, col_num As Integer) As Long
    'Finds the last row with data in column A ref:https://www.wallstreetmojo.com/vba-last-row/
    
    find_last_row_num = sheet.Columns(col_num).CurrentRegion.Rows.Count

End Function

Private Sub fomat_yr_change_results(sheet As Worksheet, column As Integer)
    Dim rng As Range
        
    Set rng = sheet.Columns(column)
    
    ' help from https://www.automateexcel.com/vba/conditional-formatting/
    
    rng.FormatConditions.Delete
    
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    rng.FormatConditions(1).Interior.Color = vbRed
    
    rng.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    rng.FormatConditions(2).Interior.Color = vbGreen
    sheet.Range(Cells(1, column).Address).FormatConditions.Delete
    
     
End Sub

Sub clear_columns_J_to_Z()
 ' Clears all the columns from J to Z
    Dim activeSheet As Worksheet

    For Each activeSheet In Worksheets
        activeSheet.Range("I1:Z1").EntireColumn.ClearContents
        activeSheet.Range("I1:Z1").EntireColumn.ClearFormats
        activeSheet.Range("I1:Z1").EntireColumn.FormatConditions.Delete
    Next

End Sub

