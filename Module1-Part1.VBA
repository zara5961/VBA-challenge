Attribute VB_Name = "Module1"
'Instructions:
'Create a script that loops through all the stocks for each quarter and outputs the following information:
'The ticker symbol
'The total stock volume of the stock.

Sub Stockanalysis()

'Loop for all the sheets.
'Set a variables

    
    Dim ticker As String                  'for ticker
    Dim volume As Double                'for vol
        volume = 0
    
'to identify the first cell to put result(summary)
        Dim summary As Integer
        summary = 2
      

'Title for the Summary headers as in ReadMe

        Cells(1, 9).Value = "Ticker"
        Cells(1, 12).Value = "Total Stock Volume"

'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'start the Loop on ticker (Column A)

        For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Set the ticker name
              ticker = Cells(i, 1).Value

              'Set the volume
              volume = volume + Cells(i, 7).Value

              'Print the ticker name in the summary (Column I)
              Range("I" & summary).Value = ticker

              'Print the volume for each ticker in the summary (Column L)
              Range("J" & summary).Value = volume

   
              'now we need to go to the next row in summary
              summary = summary + 1

              'Reset volume of trade to zero, to start add up for new ticker
              volume = 0

            Else
              
               'Add the volume
              volume = volume + Cells(i, 7).Value

            
            End If
        
        Next i


End Sub
