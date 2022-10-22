<<<<<<< HEAD
Attribute VB_Name = "Module1"
Sub Stocks()
    ' Declare Variables
    Dim vol As Double
    Dim LastRow As Double
    Dim TickerOpen As Double
    Dim TickerClose As Double

    ' Declare Variables and assign
    Dim RowDisplay As Integer
    For Each ws In Worksheets
        ws.Activate
        RowDisplay = 2 ' set equal to 2 because that's the 1st row displayed

        ' Get # of Rows in Sheet
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row

        ' Display Titles
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        TickerOpen = Cells(2, 3)
        TickerGreatestIncrease = "":        TickerGreatestIncreaseValue = -10000
        TickerGreatestDecrease = "":        TickerGreatestDecreaseValue = 10000
        TickerGreatestVolume = "":          TickerGreatestVolumeValue = -1

        ' Loop through each row
        For i = 2 To LastRow

           ' Compare current rows ticker w/ next rows ticker (Look ahead)
           If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                TickerClose = Cells(i, 6).Value

               ' add volume to total volume
               vol = vol + Cells(i, 7).Value

              ' Display ticker and ticker volume
              Cells(RowDisplay, 9) = Cells(i, 1)
              Cells(RowDisplay, 10) = TickerClose - TickerOpen
              Cells(RowDisplay, 11) = (TickerClose - TickerOpen) / TickerOpen
              Cells(RowDisplay, 12) = vol
              ' reset volume to
              vol = 0
             TickerOpen = Cells(i + 1, 3).Value

              ' increment display row
             RowDisplay = RowDisplay + 1
            End If
            
            ' add volume to total volume
            vol = vol + Cells(i, 7).Value
        Next i
        Next ws
    End Sub
=======
Attribute VB_Name = "Module1"
Sub Stocks()
    ' Declare Variables
    Dim vol As Double
    Dim LastRow As Double
    Dim TickerOpen As Double
    Dim TickerClose As Double

    ' Declare Variables and assign
    Dim RowDisplay As Integer
    For Each ws In Worksheets
        ws.Activate
        RowDisplay = 2 ' set equal to 2 because that's the 1st row displayed

        ' Get # of Rows in Sheet
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row

        ' Display Titles
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        TickerOpen = Cells(2, 3)
        TickerGreatestIncrease = "":        TickerGreatestIncreaseValue = -10000
        TickerGreatestDecrease = "":        TickerGreatestDecreaseValue = 10000
        TickerGreatestVolume = "":          TickerGreatestVolumeValue = -1

        ' Loop through each row
        For i = 2 To LastRow

           ' Compare current rows ticker w/ next rows ticker (Look ahead)
           If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                TickerClose = Cells(i, 6).Value

               ' add volume to total volume
               vol = vol + Cells(i, 7).Value

              ' Display ticker and ticker volume
              Cells(RowDisplay, 9) = Cells(i, 1)
              Cells(RowDisplay, 10) = TickerClose - TickerOpen
              Cells(RowDisplay, 11) = (TickerClose - TickerOpen) / TickerOpen
              Cells(RowDisplay, 12) = vol
              ' reset volume to
              vol = 0
             TickerOpen = Cells(i + 1, 3).Value

              ' increment display row
             RowDisplay = RowDisplay + 1
            End If
            
            ' add volume to total volume
            vol = vol + Cells(i, 7).Value
        Next i
        Next ws
    End Sub
>>>>>>> 3583b64feee9b01cec7b6f6768bec73e86e7b4fc
