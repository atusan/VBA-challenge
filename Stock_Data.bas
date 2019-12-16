Attribute VB_Name = "Module1"
Sub Stock_Data()

Dim ws As Worksheet
    
For Each ws In Worksheets
          
          'Last Row
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
          
            Dim Ticker_Name As String
            Dim Open_Price As Double
            Dim Close_Price As Double
            Dim Yearly_Change As Double
            Dim Percent_Change As Double
            Dim Volume As Double
            Volume = 0

            Dim i As Long
        
          
          ' Keep track of the location for item in the summary table
          Dim Summary_Table_Row As Integer
          Summary_Table_Row = 2
          
          ws.Range("I1").Value = "Ticker"
          ws.Range("J1").Value = " Yearly_Change"
          ws.Range("K1").Value = " Percent_Change"
          ws.Range("L1").Value = " Total Stock Volume"
         'Open_Price = ws.Cells(i, 3).Value
          Open_Price = Cells(2, 3).Value
           
           For i = 2 To LastRow
                
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                    
                       Ticker_Name = ws.Cells(i, 1).Value
                       ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                       
                     
                        Close_Price = Cells(i, 6).Value
                        Yearly_Change = Close_Price - Open_Price
                         ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                        
                        If (Open_Price = 0 And Close_Price = 0) Then
                           Percent_Change = 0
                            
                         ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                             Percent_Change = 1
                           
                           Else
                           
                           Percent_Change = Yearly_Change / Open_Price
                            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                       
                        End If
                          Volume = Volume + Cells(i, 7).Value
                          ws.Range("L" & Summary_Table_Row).Value = Volume
    
                        
                        
                        ''ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                        ''ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                        ''ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        ''ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                         ''ws.Range("L" & Summary_Table_Row).Value = Percent_Change
                        Summary_Table_Row = Summary_Table_Row + 1
                        Open_Price = Cells(i + 1, 3)
                        Volume = 0
                        
               Else
                      Volume = Volume + Cells(i, 7).Value
         
              End If
          
      
      Next i
      
Next ws
      
End Sub


