Attribute VB_Name = "Module1"
Sub alpha()

  Dim Ticker As String
  Ticker = " "
  
  Dim counter As Integer
  counter = 1
  
  Dim Ticker_Total As Double
  Ticker_Total = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

' Label columns in the summary table

  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"

' Determine the Last Row
  Dim LastRow As Long
    With ActiveSheet
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
      
' Loop through rows, all ticker
  For i = 2 To LastRow

' Check if we are still within the same ticker
  If Cells(i, 1).Value <> Ticker Then

      ' Set counter for summary table row
      counter = counter + 1

      ' Set the Ticker name
      Ticker = Cells(i, 1).Value
                  
      ' Set the Open Price
      Open_Price = Cells(i, 3).Value

      ' Add first volume to the Ticker Total
      Ticker_Total = Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      Cells(counter, 9).Value = Ticker

      ' Print the Total volume in the Summary Table
      Cells(counter, 12).Value = Ticker_Total
       
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add all volume to Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 7).Value
      
      ' Print all volume in the Summary Table
      Cells(counter, 12).Value = Ticker_Total
      
    End If

 ' Check if last ticker
   If Cells(i + 1, 1).Value <> Ticker Then

    ' Set the Close Price
      Close_Price = Cells(i, 6).Value
      
    ' Calculate the Yearly Change from Open to Close
      price_change = Close_Price - Open_Price
                            
    ' Apply to "Yearly Change" Column in summary table
      Cells(counter, 10).Value = price_change
       
    ' Conditional formatting (positive change in green and negative change in red)
      If price_change < 0 Then
        Cells(counter, 10).Interior.ColorIndex = 3
            ElseIf price_change > 0 Then
            Cells(counter, 10).Interior.ColorIndex = 4
      End If
                        
    ' Calculate the Percent Change
      If Open_Price = 0 Then
        percent_change = Open_Price
            Else: percent_change = price_change / Open_Price
      End If
        
    ' Apply to "Percent Change" in summary table
      Cells(counter, 11).Value = percent_change
      Cells(counter, 11).NumberFormat = "0.00%"
        
   End If
                
  Next i
  
' Label columns in the second table
  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greatest % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Value"


  best = 0
  worst = 0
  most = 0
        
' Determine the Last Row in summary table
  Dim LastSum As Long
    With ActiveSheet
        LastSum = .Cells(.Rows.Count, "I").End(xlUp).Row
    End With

' Loop through rows, all ticker
  For i = 2 To LastSum

  ' Check for best performer
     If Cells(i, 11).Value > best Then
     best = Cells(i, 11).Value
         Range("P2") = Cells(i, 9).Value
         Range("Q2") = Cells(i, 11).Value
     End If
      
  ' Change best performer to percentage
     Range("Q2").NumberFormat = "0.00%"
     
  ' Check for worst performer
   
     If Cells(i, 11).Value < worst Then
     worst = Cells(i, 11).Value
         Range("P3") = Cells(i, 9).Value
         Range("Q3") = Cells(i, 11).Value
     End If
     
  ' Change worst performer to percentage
     Range("Q3").NumberFormat = "0.00%"
     
  ' Check for most volume
   
     If Cells(i, 12).Value > most Then
     most = Cells(i, 12).Value
         Range("P4") = Cells(i, 9).Value
         Range("Q4") = Cells(i, 12).Value
     End If
      
   Next i
   
End Sub


