
' Bootcamp Course
' Week 2 - VBA - Challenge
' Onjectives of the scripts are -
' 1. Calculate yearly change for each ticker of given year
' 2. Calculate percentage change from openingto closing price in the given year
' 3. Calculate total stock volume per year
' Bonus - Calculate Greatest % Increase, % decrease and total volume

Sub ticker()
   
'Initialise the required varaibles
   
Dim ticker As String
Dim opening As Double
Dim closing As Double
Dim sumcount As Integer
Dim tickercnt As Double
Dim stockvalue As Double

'Initialise the required varaibles for bonus

Dim greatpchange As Double
Dim greatnchange As Double
Dim greatstock As Double

Dim greatpticker As String
Dim greatnticker As String
Dim greatstockticker As String

'Looping through the worksheets to repeat the process
  
For Each ws In Worksheets
 
 'Determine the last row in the active sheet
   
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
 'Variables to keep track ticker count and stock value
 
 sumcount = 2
 tickercnt = 0
  
 'Variables for bonus variables
  
 greatpchange = 0
 greatpticker = ""
 greatnchange = 0
 greatnticker = ""
 greatstock = 0
 greatstockticker = ""
 
 
 'Write the column names for Output table and bonus columns
 
 ws.Range("L1").Value = "Ticker"
 ws.Range("M1").Value = "Yearly Change"
 ws.Range("N1").Value = "Percent Change"
 ws.Range("O1").Value = "Total Stock Volume"
 
 ws.Range("S2").Value = "Greatest % Increase"
 ws.Range("S3").Value = "Greatest % Decrease"
 ws.Range("S4").Value = "Greatest Total Volume"
 
 'Loop the rows from 3 to last row
 
 For i = 2 To lastrow
 
 'Condition to check if we are within the same ticker, if it is not..
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
       'Set Ticker Closing and Percent change values
       ticker = ws.Cells(i, 1).Value
       closing = ws.Cells(i, 6).Value
       pchange = ((closing - opening) / opening)
       
       ' Add to the stock value
       stockvalue = stockvalue + ws.Cells(i, 7).Value
       'ws.Range("J" & sumcount).Value = opening
       'ws.Range("K" & sumcount).Value = closing
       

      'Printing the values of ticker, yearlychange and percentage change in Output columns
       ws.Range("L" & sumcount).Value = ticker
       ws.Range("M" & sumcount).Value = closing - opening
       ws.Range("N" & sumcount).Value = pchange
       ws.Range("N" & sumcount).NumberFormat = "0.00%"
       
       'Checking the condition for color formating
       
       If ws.Range("M" & sumcount).Value > 0 Then
          ws.Range("M" & sumcount).Interior.ColorIndex = 4
       Else: ws.Range("M" & sumcount).Interior.ColorIndex = 3
       End If
       'Print Stock volume in output table
       ws.Range("O" & sumcount).Value = stockvalue
       
       'Bonus - Check % change to get Greater % increase and get corresponding ticker value
       If pchange > 0 And pchange > greatpchange Then
          greatpchange = pchange
          greatpticker = ticker
       End If
       'Bonus - Check % change to get Greater % decrease and get corresponding ticker value
       If pchange < 0 And pchange < greatnchange Then
          greatnchange = pchange
          greatnticker = ticker
       End If
       'Check and get Greater stock volume
       If stockvalue > greatstock Then
          greatstock = stockvalue
          greatstockticker = ticker
       End If
           
       'Reset Stock value
       stockvalue = 0
       ' Add 1 to output table row
       sumcount = sumcount + 1
       'Reset Ticket count
       tickercnt = 0
    
      'if the cell immediately following a row the sameticker
      
    Else
         ' Add 1 to the ticker count
         tickercnt = tickercnt + 1
         ' Check if the ticker is opening row, assign opening value
         If tickercnt = 1 Then
            opening = ws.Cells(i, 3).Value
         End If
         
         'add to the stock value
         
         stockvalue = stockvalue + ws.Cells(i, 7).Value
    End If
    
 Next i
 
   'Print the Greater % increase, decrease with % percent number format and stock volume to the sheet
 
     ws.Cells(2, 20).Value = greatpticker
     ws.Cells(2, 21).Value = greatpchange
     ws.Cells(2, 21).NumberFormat = "0.00%"
     ws.Cells(3, 20).Value = greatnticker
     ws.Cells(3, 21).Value = greatnchange
     ws.Cells(3, 21).NumberFormat = "0.00%"
     ws.Cells(4, 20).Value = greatstockticker
     ws.Cells(4, 21).Value = greatstock
     
     
 Next ws
 
    
End Sub

