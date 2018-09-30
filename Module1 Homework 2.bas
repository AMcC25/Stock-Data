Attribute VB_Name = "Module1"
Sub StockEasy()
     Dim ticker As String
     Dim Volume_Total As Double
     
         Volume_Total = 0

     Dim Summary_Table_Row As Integer
     
        Summary_Table_Row = 2
     
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    

    For i = 2 To lastrow


     
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


         ticker = Cells(i, 1).Value


         Volume_Total = Volume_Total + Cells(i, 7).Value


      
         Range("I" & Summary_Table_Row).Value = ticker


         Range("J" & Summary_Table_Row).Value = Volume_Total

         Summary_Table_Row = Summary_Table_Row + 1

         Volume_Total = 0


       Else

        Volume_Total = Volume_Total + Cells(i, 7).Value
        
        End If


     Next i
     
   


   End Sub
   

Sub WorkbookLoop2()

Dim ws As Worksheet

y = 2 'starting row

For Each ws In Worksheets
  
    ws.Activate
    Debug.Print ws.Name
    
    
    
    y = y + 1
    
Next ws

End Sub

