Attribute VB_Name = "Module1"
Sub ticker_symbol()

Dim Tabs As Integer
Dim Count As Integer

  ' Set an initial variable for holding the brand name
  Dim Brand_Name As String
  
    ' Set an initial variable for holding the total per credit card brand
  Dim Total_Volume As Double

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  
  Dim Open_Price_Start As Double
  
  Dim Close_Price_End As Double
  
Tabs = ActiveWorkbook.Worksheets.Count

'MsgBox (Tabs)

'For loop for worksheet tabs
For Count = 1 To Tabs

  Total_Volume = 0

  Summary_Table_Row = 2
  
  'Last Row
   lastRow = Cells(Rows.Count, 1).End(xlUp).Row
   'MsgBox (lastRow)
   
   
   Open_Price_Start = 0
   
   
   Close_Price_End = 0
   
   'Dim j as Integer
   J = 0
  

    ' Loop through all credit card purchases
    For I = 2 To lastRow
    

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
      'Set Open Price
      Open_Price_Start = Cells(I - J, 3).Value
      
       'Set Close Price
      Close_Price_End = Cells(I, 6).Value

      ' Set the Brand name
      Brand_Name = Cells(I, 1).Value

      ' Add to the Brand Total
      Total_Volume = Total_Volume + Cells(I, 7).Value
      
      'Print the Yearly Stock in the Summary Table
      If ((Open_Price_Start - Close_Price_End) / Open_Price_Start < 0) Then
        Range("J" & Summary_Table_Row).Value = (Open_Price_Start - Close_Price_End)
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
    Else
      Range("J" & Summary_Table_Row).Value = (Open_Price_Start - Close_Price_End)
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
  End If
  
  
      'Print the Percent change in the Summary Table
         Range("K" & Summary_Table_Row).Value = ((Open_Price_Start - Close_Price_End) / Open_Price_Start)
        
      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = Brand_Name

      ' Print the Brand Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Total_Volume = 0
      J = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Total_Volume = Total_Volume + Cells(I, 7).Value
      J = J + 1
      
      
    End If

  Next I
  
Next Count

End Sub


