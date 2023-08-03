Attribute VB_Name = "Module1"
Sub market()

For Each ws In Worksheets
   
 ' Create Columns
  
   ws.Range("K1").Value = " Ticker"
   ws.Range("K1").Font.Bold = True
   

   ws.Range("L1").Value = "Yearly Change"
   ws.Range("L1").Font.Bold = True
   ws.Range("L1").EntireColumn.AutoFit
   
    
   ws.Range("M1").Value = "Percent Change"
   ws.Range("M1").Font.Bold = True
   ws.Range("M1").EntireColumn.AutoFit
   
   
   ws.Range("N1").Value = "Total Stock Volume"
   ws.Range("N1").Font.Bold = True
   ws.Range("N1").EntireColumn.AutoFit

  
' Set an initial variables

    Dim worksheetName As String
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    Dim Volume As Variant
    Volume = 0
    
    Dim stockOpen As Double
    Dim stockClose As Double
    Dim lastrow As Double
    
    Dim Summary_table As Double
    Summary_table = 2
    
    
 ' Determine the last Row
   
   lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

' loop trough all Stock Data

   For i = 2 To lastrow
   
' Check if we are still within the same sctock , if it not ...
 
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
' Set the Ticker
      Tickers = ws.Cells(i, 1).Value

' Set the Volume
      Volume = Volume + ws.Cells(i, "g").Value
   
'Print Tickers,volume
       ws.Range("k" & Summary_table).Value = Tickers
       ws.Range("n" & Summary_table).Value = Volume
   
'Reset the Volume

       Volume = 0
   
       stockClose = ws.Cells(i, 6)
   
     If stockOpen = 0 Then
       YearlyChange = 0
       PercentChange = 0
     Else
        YearlyChange = stockClose - stockOpen
        PercentChange = (stockClose - stockOpen) / stockOpen
   End If
   
' Print YearlyChange,PercentChange
       ws.Range("L" & Summary_table).Value = YearlyChange
       ws.Range("m" & Summary_table).Value = PercentChange
       ws.Range("m" & Summary_table).Style = "Percent"
              
' Add the percentage

       ws.Range("m" & Summary_table).NumberFormat = "0.00%"
' Add one to the summary table row

       Summary_table = Summary_table + 1
    
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
      stockOpen = ws.Cells(i, 3)
      Volume = Volume + ws.Cells(i, "g").Value
      ws.Range("n" & Summary_table).Value = Volume
    
    Else: Volume = Volume + ws.Cells(i, "g").Value
      ws.Range("n" & Summary_table).Value = Volume
   
   End If
   
 Next i
   
   
   
'Conditional Formatting

   For a = 2 To lastrow

   If ws.Range("L" & a).Value > 0 Then
      ws.Range("L" & a).Interior.ColorIndex = 4
   ElseIf ws.Range("L" & a).Value < 0 Then
      ws.Range("L" & a).Interior.ColorIndex = 3

    End If
   Next a
   
   For b = 2 To lastrow
  
   If ws.Range("m" & b).Value > 0 Then
      ws.Range("M" & b).Interior.ColorIndex = 4
   ElseIf ws.Range("m" & b).Value < 0 Then
      ws.Range("M" & b).Interior.ColorIndex = 3
   
   
    End If
   Next b

        
 'BONUS!!!
 
 ' Determine the last Row
 
 Lastrow1 = ws.Cells(Rows.Count, 13).End(xlUp).Row
 
 
 'Print in the summary table
  
   ws.Range("Q2").Value = "Greatest % Increase"
   ws.Range("Q2").Font.Bold = True

   ws.Range("Q3").Value = "Greatest % Decrease"
   ws.Range("Q3").Font.Bold = True
   
   ws.Range("Q4").Value = "Greatest Total Volume"
   ws.Range("Q4").Font.Bold = True
   
   ws.Range("R1").Value = "Ticker"
   ws.Range("R1").Font.Bold = True
   
   ws.Range("S1").Value = "Value"
   ws.Range("S1").Font.Bold = True
   
' Set variables
   
  Dim GreatestIncrease As Double
  GreatestIncrease = 0
  
  Dim GreatestDecrease As Double
  GreatestDecrease = 0
  
  Dim GreatestVolume As Double
  GreatestVolume = 0
  
' loop trough all summary table

  For j = 2 To lastrow
  
  If ws.Cells(j, 13).Value > GreatestIncrease Then
     GreatestIncrease = ws.Cells(j, 13).Value
     ws.Range("S2").Value = GreatestIncrease
     ws.Range("S2").Style = "Percent"
     ws.Range("S2").NumberFormat = "0.00%"
     ws.Range("R2").Value = ws.Cells(j, 11).Value
  End If
  
  Next j
  
   
  For k = 2 To lastrow
  
  If ws.Cells(k, 13).Value < GreatestDecrease Then
     GreatestDecrease = ws.Cells(k, 13).Value
     ws.Range("S3").Value = GreatestDecrease
     ws.Range("S3").Style = "Percent"
     ws.Range("S3").NumberFormat = "0.00%"
     ws.Range("R3").Value = ws.Cells(k, 11).Value
  End If
  
  Next k
  
  For L = 2 To lastrow
  
  If ws.Cells(L, 14).Value > GreatestVolume Then
     GreatestVolume = ws.Cells(L, 14).Value
     ws.Range("S4").Value = GreatestVolume
     ws.Range("R4").Value = ws.Cells(L, 11).Value
  End If
  
  Next L
  
     
 Next ws

 
 End Sub
