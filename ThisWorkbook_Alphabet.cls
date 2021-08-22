VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Market()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Sheets(1).Select
Set starting_ws = ActiveSheet

For Each ws In Worksheets

ws.Activate

Application.ScreenUpdating = False


    'Set initial variable for holding ticker name
    Dim Ticker As String

    'set initial value for holding the total stock volume per ticker
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    'keep track of the location for each stock in the summary
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Variable for opening price
    Dim OpenPrice As Double   '*** change Variant to Double

    'variable for closing price
    Dim ClosePrice As Double  '*** change Variant to Double

    'variables for yearly change
    Dim YearlyChange As Double
    Dim PercentChange As Double
    
    Dim i As Double
    
    'Counts the number of rows

    lastrow = Cells(Rows.Count, 1).End(xlUp).Row


    ' insert headers and titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"



    'set opening price
    OpenPrice = Cells(2, 3).Value


    'Loop through all tickers

    For i = 2 To lastrow

    'Check if we are still within the same ticker, if it is not

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    'set the ticker
        Ticker = Cells(i, 1).Value

  '      MsgBox ("Values:  Ticker: " & Ticker & " Open: " & OpenPrice & " Close: " & ClosePrice)
        ClosePrice = Cells(i, 6).Value
        YearlyChange = ClosePrice - OpenPrice
        Range("J" & Summary_Table_Row).Value = YearlyChange
        If OpenPrice <> 0 Then
        PercentChange = (YearlyChange / OpenPrice)
            Range("K" & Summary_Table_Row).Value = PercentChange
        Else
            Range("K" & Summary_Table_Row).Value = 0
        End If
        OpenPrice = Cells(i + 1, 3).Value
        Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
    'add to the total stock volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    'Print the ticker in the summary table
        Range("I" & Summary_Table_Row).Value = Ticker
    
    'Print the total stock volume to the summary table
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
    
    'add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset the brand total
        Total_Stock_Volume = 0
    
    'If the cell immediately following a row is the same brand...
    
    Else
    
    'add to the total stock volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
    End If
    
    Next i
    
            ' find max increase, max decrease, and max total volume
            Range("Q2").Value = WorksheetFunction.Max(Range("K:K"))
            Range("Q2").NumberFormat = "0.00%"
            Range("Q3").Value = WorksheetFunction.Min(Range("K:K"))
                'FormatPercent (Range("Q2"))
            Range("Q3").NumberFormat = "0.00%"
            Range("Q4").Value = WorksheetFunction.Max(Range("L:L"))



    ' conditional colors
    For i = 2 To lastrow
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        'autofit column width for stock totals *in every worksheet*
            Columns("J").NumberFormat = "0.00"
            Columns("K").AutoFit
            Columns("L").AutoFit
        End If
    Next i
    
                                            'find matching ticker symbols
                                            For i = 2 To lastrow
                                            If Cells(i, 11).Value = Range("Q2").Value Then
                                            Range("P2").Value = Cells(i, 9).Value
                                            End If

                                            If Cells(i, 11).Value = Range("Q3").Value Then
                                            Range("P3").Value = Cells(i, 9).Value
                                            End If

                                            If Cells(i, 12).Value = Range("Q4").Value Then
                                            Range("P4").Value = Cells(i, 9).Value
                                            End If
                                            Next i

                                                Columns("O").AutoFit
                                                Columns("Q").AutoFit
                                                Range("Q4").NumberFormat = "#,##0"
Next ws
   
starting_ws.Activate

Application.ScreenUpdating = True

    
End Sub

