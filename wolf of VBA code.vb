Sub wolf_of_VBA()
    
    '---Andrew Scott
    '-------------------------------------------
    
    'Dim variables
    Dim ticker As String
    Dim vol As Double
    Dim sum_table As Integer
    Dim lrow As Long
    Dim ws As Worksheet
        
    
    'Executing for all sheets
    For Each ws In Worksheets
        ws.Activate
                    
            'find total number of rows for lrow
            lrow = Cells(Rows.Count, 1).End(xlUp).Row
                
            'set initial table variable value
            sum_table = 2
                
            'not sure where to put this so it's going here--printing column headers
            Range("i1").Value = "Ticker"
            Range("J1").Value = "Volume"
            
            'Start loop
            For I = 2 To lrow
              
            'find break in ticker symbol
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            
            'sum volume for ticker
            ticker = Cells(I, 1).Value
            vol = vol + Cells(I, 7).Value
            
            'printing summary table
            Range("I" & sum_table).Value = ticker
            Range("J" & sum_table).Value = vol
            
            'move summary table down one row
            sum_table = sum_table + 1
            
            'I could use a snickers candybar right now
            
            'reset ticker volume sum
            vol = 0
            
            Else
            
            vol = vol + Cells(I, 7).Value
            
                
            End If
            
            Next I
    
    Next ws
   
End Sub