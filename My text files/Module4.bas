Attribute VB_Name = "Module4"
Sub Finalwork()
 
 ' LOOP THROUGH ALL SHEETS


 Dim ws As Worksheet
' Loop through all sheets in the workbook
For Each ws In ThisWorkbook.Sheets
ws.Activate

'Declare the variables

' Set an initial variable for holding the Ticker name
 Dim Ticker_Name As String

' Set an initial variable for holding the Total Stock Volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0


Dim YearlyChange As Double
Dim Openprice As Double
Dim Closeprice As Double
Dim Percent_Change As Long
Dim PercentChange As Long
Dim Finalrow As Integer
Dim i As Long
Dim lastrow As Long
Dim PercentMin As Double
Dim PercentMax As Double
Dim VolumeMax As Double
Dim PercentMinTicker As String
Dim PercentMaxTicker As String
Dim VolumeMaxTicker As String

PercentMin = 0
PercentMax = 0
VolumeMax = 0




    
 
' Keep track of the location for each credit card brand in the Ticker outcome table
 
  Dim Ticker_Counter As Integer
  Ticker_Counter = 2
firstrow = 2



' Determine the Last Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Add a new column for the ticker symbols
ws.Range("I1").EntireColumn.Insert
'Add a header for Tickers
 ws.Cells(1, "I").Value = "Ticker"

'Add a new column for the Yearly change
ws.Range("J1").EntireColumn.Insert
    
    'Enter header for yearly change
    
ws.Cells(1, "J").Value = "Yearly_change"

' Add a new column for the Percent Change
   
 ws.Range("K1").EntireColumn.Insert
    'Insert header for Percent Change
    ws.Cells(1, "K").Value = "Percent_Change"
   
   'Add a new column for the Total stock Volume
   ws.Range("L1").EntireColumn.Insert
   
   'Insert header for Total stock Volume
   ws.Cells(1, "L").Value = "Total Stock Volume"
   'Loop through all the ticker
    
    

For i = 2 To lastrow

    ' Check if we are still within the same Ticker, if it is not...
    
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value
     ' Print the Ticker name in the Ticker Table
      Range("I" & Ticker_Counter).Value = Ticker_Name


      ' Add to the Stock Volume Total
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print the stock volume total to the Ticker Table
      Range("L" & Ticker_Counter).Value = Total_Stock_Volume
'take the open price
Openprice = Cells(firstrow, 3).Value
'take the close price

Closeprice = Cells(i, 6).Value
'calculate yaerlychange
YearlyChange = Closeprice - Openprice
'Insert yearly change in the column
 Range("J" & Ticker_Counter).Value = YearlyChange
 'Formating the Column Yearly Change

If YearlyChange > 0 Then
Range("J" & Ticker_Counter).Interior.ColorIndex = 4
Else
Range("J" & Ticker_Counter).Interior.ColorIndex = 3
End If


'calculate percent change
If YearlyChange = 0 Or Openprice = 0 Then
                
                Cells(Ticker_Counter, 11).Value = 0
            Else
           Cells(Ticker_Counter, 11).Value = (YearlyChange / Openprice)
            'Range("K" & Ticker_Counter).Value = PercentChange
          
           
'print percentage change in the column
'Range("K" & Ticker_Counter).Value = (YearlyChange / Openprice)
'Percent_Change = (YearlyChange / Openprice)
Cells(Ticker_Counter, 11).NumberFormat = "0.00%"

 End If
 
 Cells(2, 15).Value = "Greatest%Increase"
  Cells(3, 15).Value = "Greatest%Decrease"
   Cells(4, 15).Value = "Greatest Total Volume"
   Cells(1, 16).Value = "Ticker"
   Cells(1, 17).Value = "Value"
 

If PercentChange > PercentMax Then
PercentMax = PercentChange
PercentMaxTicker = Ticker_Name
End If


If Cells(Ticker_Counter, 11).Value < PercentMin Then
PercentMin = Cells(Ticker_Counter, 11).Value
PercentMinTicker = Ticker_Name
End If
If Cells(Ticker_Counter, 11).Value > PercentMax Then
PercentMax = Cells(Ticker_Counter, 11).Value
PercentMaxTicker = Ticker_Name
End If
If Total_Stock_Volume > VolumeMax Then
VolumeMax = Total_Stock_Volume
VolumeMaxTicker = Ticker_Name
End If

Cells(2, 16).Value = PercentMaxTicker
Cells(2, 17).Value = PercentMax
Cells(3, 16).Value = PercentMinTicker
Cells(3, 17).Value = PercentMin

Cells(4, 16).Value = VolumeMaxTicker
Cells(4, 17).Value = VolumeMax

Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"





' Add one to the Ticker Counter Table
Ticker_Counter = Ticker_Counter + 1
'Add one to Firstrow
firstrow = i + 1
 ' Reset the Stock Volume Total
      Total_Stock_Volume = 0

    ' If the cell immediately following a row is the same Ticker...
    Else

      
      ' Add to the Total stock Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
End If

Next i
' Autofit to display data
ws.Columns("A:Q").AutoFit
Next ws

End Sub



