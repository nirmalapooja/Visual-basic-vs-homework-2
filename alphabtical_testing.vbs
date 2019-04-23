Attribute VB_Name = "Module1"


' Create a sub routine with the filename
Sub Hweasy():
  ' Set an initial variable for holding the Ticker
  Dim Ticker_name As String


  ' Set an initial variable for holding the total voulme
  Dim Total_Volume As Double
  Total_Volume = 0

  ' Keep track of the location for each Ticker in the Data table
  Dim Ticker_Table_Row As Integer
  Ticker_Table_Row = 2

  ' Loop through all Ticker data
  Dim LastRow As Long
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  For i = 2 To LastRow
  
  
    ' Check if we are still within the same Ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_name = Cells(i, 1).Value

      ' Add to the Total_Volume
      Total_Volume = Total_Volume + Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Ticker_Table_Row).Value = Ticker_name

      ' Print the Brand Amount to the Summary Table
      Range("J" & Ticker_Table_Row).Value = Total_Volume

      ' Add one to the Ticker_table_row
      Ticker_Table_Row = Ticker_Table_Row + 1
      
      ' Reset the Total_volume
      Total_Volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Total_Volume
      Total_Volume = Total_Volume + Cells(i, 7).Value

    End If

  Next i
 
End Sub
