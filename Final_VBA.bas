Attribute VB_Name = "Module1"
Sub tick()

  ' Set an initial variable for holding the brand name
Dim Ticker As String
Dim openC As Double
Dim CloseF As Double
Dim Summary_Table_Row As Long
Dim stock_Total As Double
Dim PerChange As Double
Dim Yearchange As Double
stock_Total = 0
Summary_Table_Row = 2
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row


  ' Set an initial variable for holding the total per credit card brand
  
 

  ' Keep track of the location for each credit card brand in the summary table
  


  ' Loop through all stock tickers'
 

    ' Check if we are still within the same stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the stock name
      stock_Name = Cells(i, 1).Value

      ' Add to the stock total
      stock_Total = stock_Total + Cells(i, 7).Value
      
      
      ' Print the stock name in the Summary Table
      Range("I" & Summary_Table_Row).Value = stock_Name
      
      
      
      ' Print the stock Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
     
     
     
      ' Reset the Brand Total
      stock_Total = 0

      
           
      
      
      
            
      
      
    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the stock Total
      stock_Total = stock_Total + Cells(i, 7).Value
        
        
     
      
    End If
  Next i

End Sub

Sub YearPer():
Dim Yearchange As Double
Dim Summary_Table_Row As Long
Dim openC As Double
Dim CloseF As Double
Dim PerChange As Double
Summary_Table_Row = 2
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row



    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
            openC = Cells(i, 3).Value
            'Cells(i, 14).Value = openC'
    ElseIf Cells(i, 1).Value <> Cells(i + 1, 1) Then
            
            CloseF = Cells(i, 6).Value
            'Cells(i, 15).Value = CloseF'
    'Else Cells(i, 1).Value <> Cells(i - 1, 1) Or Cells(i, 1).Value <> Cells(i + 1, 1) Then'
    
    Yearchange = CloseF - openC
    Range("J" & Summary_Table_Row).Value = Yearchange
    
      
    Summary_Table_Row = Summary_Table_Row + 1

   
    
    
   
   End If
   
      
Next i
End Sub

Sub Percent():
Dim PerChange As Double
Dim Summary_Table_Row As Long
Dim Yearchange As Double
Dim openC As Double
Dim i As LongLong
Summary_Table_Row = 2
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    Yearchange = Cells(i, 10).Value
    
    If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
        openC = Cells(i, 3).Value
    End If
    
    If openC = 0 Then
    PerChange = 0
     
     
    Else
    PerChange = Yearchange / openC

    Range("K" & Summary_Table_Row).Value = PerChange
    Summary_Table_Row = Summary_Table_Row + 1
    End If
Next i

End Sub

