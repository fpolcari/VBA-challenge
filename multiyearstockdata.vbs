Attribute VB_Name = "Module1"
Sub credit():
    
    ' define the variable to hold the open price
    Dim openPrice As Double
    
        ' use a for loop to populate the calculations into sheet B
        For Each Sheet In ThisWorkbook.Worksheets
        
        Sheet.Activate
        
        ' variable to hold the ticker symbol
        Ticker = ""
    
        ' variable to hold the total stock volume
        totalVolume = 0
    
        ' give your openPrice a value
        openPrice = Cells(2, 3).Value
    
        ' Variable to hold the close price
        closePrice = ""
    
       ' variable to hold the yearly change
        yearlyChange = ""
    
       ' variable to hold the percent change
        percentChange = ""
    
       ' variable to hold the summary table starter row
        summaryTableRow = 2
    
        ' use function to find the last row in the sheet
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
                
                ' populate the column headers in sheet A and autofit
                Range("I1").Value = "Ticker"
                Range("J1").Value = "Yearly Change"
                Range("K1").Value = "Percent Change"
                Range("L1").Value = "Total Stock Volume"
                Range("I1:L1").Columns.AutoFit

                  ' loop from row 2 in column A out to the last row
                    For Row = 2 To lastRow
            
                     ' check to see if the ticker symbol changes
                         If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                    
                        ' if the ticker symbol changes, do ....
                        ' first set the ticker symbol
                        Ticker = Cells(Row, 1).Value
                        
                        ' add the last stock volume from the row
                        totalVolume = totalVolume + Cells(Row, 7).Value
                            
                            ' add the ticker symbol to the I column in the summary table row
                            Cells(summaryTableRow, 9).Value = Ticker
                            
                            ' add the total volume to the L column in the summary table row
                            Cells(summaryTableRow, 12).Value = totalVolume
                            
                                ' define the close price of the ticker symbol
                                closePrice = Cells(Row, 6).Value
                                
                                 ' add the yearly change to column J
                                 yearlyChange = closePrice - openPrice
                                 
                                 ' add the yearly change to the J column
                                 Cells(summaryTableRow, 10).Value = yearlyChange
                         
                                     If openPrice = 0 Then
                                     percentChange = 0
                                     
                                     Else
                                     
                                     ' compute % change
                                        percentChange = (yearlyChange) / (openPrice)
                                        
                                        End If
                         
                                        ' add the percentange change to the K column
                                        Cells(summaryTableRow, 11).Value = percentChange
                                    
                                         ' add the percent format to the K column
                                           Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                                    
                                           If yearlyChange < 0 Then
                                           
                                           Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                                           ElseIf yearlyChange > 0 Then
                                           Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                                
                                            End If
                                
                                       ' set the open price for next ticker
                                        openPrice = Cells(Row + 1, 3).Value
                                        
                                       ' go to the next summary table row (add 1 on to the value of the summary table row)
                                        summaryTableRow = summaryTableRow + 1
                                   
                                       ' reset the Volume total to 0
                                       totalVolume = 0
                                   
                                       Else
                                       ' if the ticker stays the same, do....
                                       ' add on to the total volumne from the G column
                                       totalVolume = totalVolume + Cells(Row, 7).Value
                                       
                                
                                End If
                    
                        Next Row
        Next Sheet
    
End Sub


