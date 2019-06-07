Sub Button2_Click()
  Dim sheetn As String
  Dim ws As String
  Dim i As Integer
  
  'MEDIUM AND HARD CHALLENGE
  
  'Activate the workbook/file where data is located ("ws")
'ws = "Multiple_year_stock_data - CEM.xlsm"
  
  ws = ActiveWorkbook.Name
  
  Workbooks(ws).Activate
  
  'Go through each sheet with data to calculate summary
  For i = 1 To Sheets.Count - 1
       
      'Get Name of the worksheet to work on
      sheetn = Sheets(i).Name
      
      'Generate summary for "sheetn"
      Call CreateSummary2(sheetn, ws)
 
  Next i
  
  'Go back to first sheet
  sheetn = Sheets(1).Name
  Workbooks(ws).Sheets(sheetn).Select
  Cells(7, 14).Select
  
End Sub

Sub CreateSummary2(sheetn As String, ws As String)

  Dim i As Long
  Dim j As Integer
  Dim h As Integer
  Dim c As Integer
  
  Dim s As Double
  Dim n As String
  Dim e As Boolean
  Dim a As Integer
  
  Dim v1 As Double
  Dim v2 As Double
  Dim yc As Double
  Dim ycp As Double
  
  Dim nv1 As String  'Store name of stock with greatest volume
  Dim nd1 As String  'Store name of stock with greatest decrease
  Dim ni1 As String  'Store name of stock with greatest increase
  
  Dim gv1 As String  'Store value of greatest volume
  Dim gd1 As String  'Store greatest decrease
  Dim gi1 As String  'Store greatest increase
  
 
  'Activate the sheetn in current workbook (ws)
  Workbooks(ws).Worksheets(sheetn).Activate
   
   'Row ("h") and Column ("c") where consolidated data would be placed
   h = 7
   c = 14
   
   'Initialize the heading for the table where summary would be placed - for MEDIUM CHALLENGE
   Workbooks(ws).Worksheets(sheetn).Cells(h - 2, c) = "MEDIUM CHALLENGE - YEAR " & sheetn
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c) = "Ticker"
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c + 1) = "Yearly Change"
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c + 2) = "Percent Change"
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c + 3) = "Total Stock Volume"
   
   'Make titles in Bold
   Workbooks(ws).Worksheets(sheetn).Cells(h - 2, c).Font.Bold = True
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c).Font.Bold = True
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c + 1).Font.Bold = True
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c + 2).Font.Bold = True
   Workbooks(ws).Worksheets(sheetn).Cells(h - 1, c + 3).Font.Bold = True
   
   
   'i and j are initialized to start going throug stock data table
   i = 2
   j = 1
   
   'variable "s" is initialized with the stock volume in the first line -> s will maintain the total volume for each stock
   s = Workbooks(ws).Worksheets(sheetn).Cells(i, 7).Value
   
   'variable "n" is initialized with the name of stock in the first line
   n = Workbooks(ws).Worksheets(sheetn).Cells(i, j).Value
   
   'variable "v1" is the opening value of stock in the year
   v1 = Workbooks(ws).Worksheets(sheetn).Cells(i, 3).Value
   
   'Initialize boolean variable used to check if the data in table finished
   e = False
   
   
   'Initialize variable to keep he largest volume stock, largest % decrease and largest % increase
   nv1 = n
   nd1 = n
   ni1 = n
   
   gv1 = 0
   gd1 = 0
   gi1 = 0
   
   'While not finished with data - start calculations for summary
   Do While Not (e)
   
     'Check if the current line in the table is empty - if it is empty, Do While loop finishes (e is set to True)
     If IsEmpty(Workbooks(ws).Worksheets(sheetn).Cells(i, j)) Then
        e = True
     Else
     
        'If not finished table, test if there is change in stock
        If n <> Workbooks(ws).Worksheets(sheetn).Cells(i + 1, j) Then
           
           'If the stock changes, store the closing value of stock in v2
           v2 = Workbooks(ws).Worksheets(sheetn).Cells(i, 6).Value
           
           'Calculate yearly change ("yc")
           yc = v2 - v1
            
           'Update the summary data with stock name (n), Yearly change (yc)
           Workbooks(ws).Worksheets(sheetn).Cells(h, c).Value = n  'Name of current stock
           Workbooks(ws).Worksheets(sheetn).Cells(h, c + 1).Value = yc  'Yearly change for current stock
           Workbooks(ws).Worksheets(sheetn).Cells(h, c + 1).NumberFormat = "#,##0.00000"  'Format for yearly change

           
           'if yc is negative, set cell in red, otherwise set cell in green
           If yc < 0 Then
              Workbooks(ws).Worksheets(sheetn).Cells(h, c + 1).Interior.ColorIndex = 3
           Else
              Workbooks(ws).Worksheets(sheetn).Cells(h, c + 1).Interior.ColorIndex = 4
           End If

           
           'Update yearly percentage change for current stock in summary table
           If v1 <> 0 Then 'Ensure v1 is different of zero to avoid error
              
              'If v1 is <>0, calculate year change %, store in summary table
              ycp = (v2 / v1) - 1
              Workbooks(ws).Worksheets(sheetn).Cells(h, c + 2).Value = ycp
              Workbooks(ws).Worksheets(sheetn).Cells(h, c + 2).NumberFormat = "0.00%"
              
              'Check if new year % change is smaller than current smallest (gd1)
              If ycp < gd1 Then
                 'if there is a new smallest % change -> update data for variables storing stock with smallest year change
                 gd1 = ycp
                 nd1 = n
              End If
              
              'Check if new year % change (ycp) is the largest vs. current (gi1)
              If ycp > gi1 Then
                 'If ycp is largest, then update data for variables storing stock with largest % change
                 gi1 = ycp
                 ni1 = n
              End If
              
              
           Else
              'if v1 is equal to zero, then assing a value of "NA" to year % change
              Workbooks(ws).Worksheets(sheetn).Cells(h, c + 2).Value = "NA"
           End If
           
           'Update summary table with total volume for current stock
           Workbooks(ws).Worksheets(sheetn).Cells(h, c + 3).Value = s
           Workbooks(ws).Worksheets(sheetn).Cells(h, c + 3).NumberFormat = "#,##0"

           'Check is new total volume (s) is greater than current largest one (gv1)
           If s > gv1 Then
               'If s is larger, update variable storing information for stock with largest volume
               gv1 = s
               nv1 = n
           End If
           
           
           'Initializa variables for the total volume ("s"), current name of stock ("n") and initial value ("v1"), to start for next stock
           i = i + 1
           s = Workbooks(ws).Worksheets(sheetn).Cells(i, 7)
           n = Workbooks(ws).Worksheets(sheetn).Cells(i, j)
           v1 = Workbooks(ws).Worksheets(sheetn).Cells(i, 3)
           h = h + 1
        Else
           'if not different, add the volumen to the standing total volume fo the stock ("s"), and move to next line
           s = s + Workbooks(ws).Worksheets(sheetn).Cells(i + 1, 7).Value
           i = i + 1
        End If
     
     End If
     
   Loop

    'Resize/Autofit columns for the table
    Columns("N:Q").Select
    Columns("N:Q").EntireColumn.AutoFit

    '********************************************************************************************
    'Section to populate table with greatest % increase, greatest % decrease and greatest volume
    '********************************************************************************************

    'Column (a) and row (h) where summary table with greatest decrease, increase and volume will be placed
    a = 19
    h = 7
    
    'Set titles for table of greates increase, decrease and volume
    Workbooks(ws).Worksheets(sheetn).Cells(h - 2, a).Value = "HARD CHALLENGE - YEAR " & sheetn
    Workbooks(ws).Worksheets(sheetn).Cells(h - 1, a).Value = "Dimension"
    Workbooks(ws).Worksheets(sheetn).Cells(h - 1, a + 1).Value = "Ticker"
    Workbooks(ws).Worksheets(sheetn).Cells(h - 1, a + 2).Value = "Value"

    'Set to BOLD the titles for table
    Workbooks(ws).Worksheets(sheetn).Cells(h - 2, a).Font.Bold = True
    Workbooks(ws).Worksheets(sheetn).Cells(h - 1, a).Font.Bold = True
    Workbooks(ws).Worksheets(sheetn).Cells(h - 1, a + 1).Font.Bold = True
    Workbooks(ws).Worksheets(sheetn).Cells(h - 1, a + 2).Font.Bold = True

    'Place Greatest % increase stock information
    Workbooks(ws).Worksheets(sheetn).Cells(h, a).Value = "Stock with greatest % increase"
    Workbooks(ws).Worksheets(sheetn).Cells(h, a + 1).Value = ni1
    Workbooks(ws).Worksheets(sheetn).Cells(h, a + 2).Value = gi1
    Workbooks(ws).Worksheets(sheetn).Cells(h, a + 2).NumberFormat = "0.00%"
    
    'Place Greatest % decrease stock information
    Workbooks(ws).Worksheets(sheetn).Cells(h + 1, a).Value = "Stock with greatest % decrease"
    Workbooks(ws).Worksheets(sheetn).Cells(h + 1, a + 1).Value = nd1
    Workbooks(ws).Worksheets(sheetn).Cells(h + 1, a + 2).Value = gd1
    Workbooks(ws).Worksheets(sheetn).Cells(h + 1, a + 2).NumberFormat = "0.00%"
   
    'Place Greates Total Volume stock information
    Workbooks(ws).Worksheets(sheetn).Cells(h + 2, a).Value = "Greatest Total Volume"
    Workbooks(ws).Worksheets(sheetn).Cells(h + 2, a + 1).Value = nv1
    Workbooks(ws).Worksheets(sheetn).Cells(h + 2, a + 2).Value = gv1
    Workbooks(ws).Worksheets(sheetn).Cells(h + 2, a + 2).NumberFormat = "#,##0"
    
    'Resize/Autofit columns for the table
    Columns("S:U").Select
    Columns("S:U").EntireColumn.AutoFit
    Range("N5").Select


End Sub


