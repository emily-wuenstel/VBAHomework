Sub Stock_Analysis()

Dim Ticker As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim volume As Single


'------------------------------- Setting While Loop  -------------------
Dim WS_Count As Integer
Dim X As Integer

'get count of WS's
WS_Count = ActiveWorkbook.Worksheets.Count

'loop through worskheets
For X = 1 To WS_Count

'activate the worksheet in the loop (1,2,3)
Worksheets(X).Activate



'------------------------------- Core code ----------------------------
    'Set Column titles for Sheet
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Open Price"
    Cells(1, 12).Value = "Close Price"
    Cells(1, 13).Value = "Yearly Change"
    Cells(1, 14).Value = "Percent Change"
    Cells(1, 15).Value = "Volume"

    'Summary table counter
    Dim Summary_Row As Integer
    Summary_Row = 2

    'create a first open variable to track first open price and not overwrite it
    Dim First_Open As Integer
    First_Open = 1

    'loop through to get Tickers, Open, and Close prices for the year
    For I = 2 To 70413

        'If Cells are the same, count volume and the open price
        If Cells(I + 1, 1).Value = Cells(I, 1).Value Then
    
            volume = volume + Cells(I, 7).Value
            
        'change first open to 2 so it won't get overwritten
         If First_Open = 1 Then
                Open_Price = Cells(I, 3).Value
             First_Open = 2
         Else
        End If
              
            
        'If they are not the same, print the running results
        Else
             'set ticker name
             Ticker = Cells(I, 1).Value
        
            'add to volume
             volume = volume + Cells(I, 7).Value
        
            'Get close price
            Close_Price = Cells(I, 6).Value
        
            'Print ticker in summary table
            Range("J" & Summary_Row).Value = Ticker
        
            'Print Open price in summary table
            Range("K" & Summary_Row).Value = Open_Price
        
            'Print Close price in summary table
            Range("L" & Summary_Row).Value = Close_Price
        
            'Print Volume
            Range("O" & Summary_Row).Value = volume
        
            'Reset Volume
            volume = 0
        
            'Add Summary Row
            Summary_Row = Summary_Row + 1
        
            'Reset First open to 1 to grab next price
            First_Open = 1
           
        End If
     
    
    Next I


    'calculate yearly change and place it into column M, and Percent change into Column N

    Dim Open_Year As Double
    Dim Close_Year As Double
    Dim Year_Change As Double
    Dim Perc_Change As Double
    
    Current_Sheet = ActiveSheet.Name
    
    'find last row
    Last_Row_Total = Cells(Rows.Count, "J").End(xlUp).Row

    
    For I = 2 To Last_Row_Total

        Open_Year = Cells(I, 11).Value
        Close_Year = Cells(I, 12).Value
    
        'calculate $ change for the year
        Year_Change = (Close_Year - Open_Year)
        
        'Make sure the change is not rounded
        Year_Change = Round(Year_Change, 2)
    
        'put that into column M
        Cells(I, 13).Value = Year_Change
    
        'calculate % change (increase / original) and then format to % for excel below, but need an if statement for if the change is 0 (this happens in P)
    
        If Year_Change < 0 Then
            Perc_Change = Perc_Change * -1
            Perc_Change = (Year_Change / Open_Year)
    
            'Put that value in Column N
            Cells(I, 14).Value = Perc_Change
            
        ElseIf Year_Change > 0 Then
        
            Perc_Change = (Year_Change / Open_Year)
    
            'Put that value in Column N
            Cells(I, 14).Value = Perc_Change
            
        
        ElseIf Year_Change = 0 Then
        
            Cells(I, 14).Value = 0
            
        End If
        
    
    Next I

    'Set Column N to percentage format

    Range("N:N").NumberFormat = "00.00%"
    
    'Make columns look nice with Autofit

    Cells.Columns.AutoFit
    
    
'conditionals to turn cells green or red based on values but we can use the same last_row_total value since it's the same set of data

For I = 2 To Last_Row_Total

    If Cells(I, 14) > 0 Then
        Cells(I, 14).Interior.ColorIndex = 4
    
    ElseIf Cells(I, 14) <= 0 Then
        Cells(I, 14).Interior.ColorIndex = 3
    End If
    
Next I
    
Next X


End Sub
Sub Clear_Data()

Dim WS_Count As Integer
Dim X As Integer

WS_Count = ActiveWorkbook.Worksheets.Count

For X = 1 To WS_Count

Worksheets(X).Activate

    Range("J:O").ClearContents
    Range("J:O").ClearFormats

Next X


End Sub

