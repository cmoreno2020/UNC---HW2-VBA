Sub Button1_Click()
  Dim sheetn As String
  Dim ws As String
  Dim i As Integer
  
  'EASY CHALLENGE'
  
  'Name of the file/workbook where files are located, and then activate the file
  'ws = "Multiple_year_stock_data - CEM.xlsm"
  
  ws = ActiveWorkbook.Name
  Workbooks(ws).Activate
  
  
  'Iterate through the sheets in the workbook (-1 to account for last two sheets which include images and a pivot table for testing)
  For i = 1 To Sheets.Count - 1
       
      'Get Name of the worksheet to work on
      sheetn = Sheets(i).Name
      
      'Creates the Summary table for the "sheetn" in workbook "ws"
      Call CreateSummary(sheetn, ws)
 
  Next i
  
  'Go back to first sheet
  sheetn = Sheets(1).Name
  Workbooks(ws).Sheets(sheetn).Select
  Cells(7, 14).Select
  
End Sub

Sub CreateSummary(sheetn As String, ws As String)

  Dim i As Long
  Dim j As Integer
  Dim h As Integer
  
  Dim s As Double
  Dim n As String
  Dim e As Boolean
    
  'EASY CHALLENGE
 
  'Activate the sheetn in current workbook (ws)
  Workbooks(ws).Worksheets(sheetn).Activate
   
   'Row where consolidated data would start being placed - initialized in 7
   h = 7
   
   'Initialize the heading for the table where summary would be placed
   Workbooks(ws).Worksheets(sheetn).Cells(h - 2, 10) = "EASY CHALLENGE TABLE - YEAR " & sheetn
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, 10) = "Ticker"
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, 11) = "Total Stock Volume"
   
   Workbooks(ws).Worksheets(sheetn).Cells(h - 2, 10).Font.Bold = True
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, 10).Font.Bold = True
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, 11).Font.Bold = True
   
   'i and j are initialized to track each line in table including stock data
   i = 2
   j = 1
   
   'variable "s" is initialized with the stock volume in the first line -> "s" to store the total volume for each stock
   s = Workbooks(ws).Worksheets(sheetn).Cells(i, 7)
   
   'variable "n" is initialized with the name of stock in the first line
   n = Workbooks(ws).Worksheets(sheetn).Cells(i, j)
   
   'Initialize boolean variable used to check if the data in stock-table finished
   e = False
   
   
   'While not finished with data - start calculations for summary
   Do While Not (e)
   
     'If current line is empty, "e" is set to True, and the Do While loop will finish
     If IsEmpty(Workbooks(ws).Worksheets(sheetn).Cells(i, j)) Then
        e = True
     Else
     
        'If current line is not empty - test if there is change of stock name to start calculation for new stock
        If n <> Workbooks(ws).Worksheets(sheetn).Cells(i + 1, j) Then
           
           'If change of stock - update the summary data in summary table for current stock
           Workbooks(ws).Worksheets(sheetn).Cells(h, 10) = n
           Workbooks(ws).Worksheets(sheetn).Cells(h, 11) = s
           
           'Initializa variables for the total volume ("s") and name of current stock ("n"), to start calculating total volume for next stock
           i = i + 1
           s = Workbooks(ws).Worksheets(sheetn).Cells(i, 7)
           n = Workbooks(ws).Worksheets(sheetn).Cells(i, j)
           h = h + 1
        Else
           'if not different, add the volumen to the standing total volume fo the stock ("s"), and move to next line
           s = s + Workbooks(ws).Worksheets(sheetn).Cells(i + 1, 7).Value
           i = i + 1
        End If
     
     End If
     
   Loop

End Sub

