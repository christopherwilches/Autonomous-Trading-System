Sub RunCycleprocedure()
    Dim wsControl As Worksheet
    Dim wsMACD As Worksheet
    Dim wsEMA As Worksheet
    Dim wsRSI As Worksheet
    Dim wsBreakout As Worksheet
    Dim wsADX As Worksheet
    Dim wsVolatility As Worksheet
    Dim wsSMA As Worksheet
    Dim wsBollinger As Worksheet
    Dim wsEMAMACD As Worksheet
    Dim wsRSIBollinger As Worksheet

    Dim i As Long
    Dim startRow As Long
    Dim destRowMACD As Long
    Dim destRowEMA As Long
    Dim destRowRSI As Long
    Dim destRowBreakout As Long
    Dim destRowADX As Long
    Dim destRowVolatility As Long
    Dim destRowSMA As Long
    Dim destRowBollinger As Long
    Dim destRowEMAMACD As Long
    Dim destRowRSIBollinger As Long
    Dim NumberofTickers As Long
    Dim Startrowtbl As Long

    ' Set the worksheets
    Set wsControl = ThisWorkbook.Sheets("Controlsheet")
    Set wsMACD = ThisWorkbook.Sheets("1 MACD")
    Set wsEMA = ThisWorkbook.Sheets("2 EMA")
    Set wsRSI = ThisWorkbook.Sheets("3 RSI")
    Set wsBreakout = ThisWorkbook.Sheets("4 BREAKOUT")
    Set wsADX = ThisWorkbook.Sheets("5 ADX")
    Set wsVolatility = ThisWorkbook.Sheets("6 Volatility Measure")
    Set wsSMA = ThisWorkbook.Sheets("7 SMA")
    Set wsBollinger = ThisWorkbook.Sheets("8 Bollinger Bands")
    Set wsEMAMACD = ThisWorkbook.Sheets("9 EMA & MACD")
    Set wsRSIBollinger = ThisWorkbook.Sheets("10 RSI Bollinger")

    On Error GoTo SafeExit

    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim prevStatusBar As Boolean

    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevStatusBar = Application.DisplayStatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual

    ' Initialize the starting row for the cycle table
    Startrowtbl = Worksheets("ControlSheet").Range("Q1").Value2
    destRowMACD = Startrowtbl
    destRowEMA = Startrowtbl
    destRowRSI = Startrowtbl
    destRowBreakout = Startrowtbl
    destRowADX = Startrowtbl
    destRowVolatility = Startrowtbl
    destRowSMA = Startrowtbl
    destRowBollinger = Startrowtbl
    destRowEMAMACD = Startrowtbl
    destRowRSIBollinger = Startrowtbl

    ' Store the first dataset before cycling begins
    Dim originalData As Variant
    startRow = 2 ' Ensure startRow is initialized
    originalData = wsControl.Range("A" & startRow & ":H" & (startRow + 32)).Value2

    Worksheets("ControlSheet").Range("N1").Value2 = Time
    NumberofTickers = Worksheets("ControlSheet").Range("K1").Value2

    ' Loop through the interactions
    For i = 0 To NumberofTickers - 1

        ' Calculate the starting row every time
        startRow = 55 + ((i - 1) * 53)

        ' Copy the block
        If i > 0 Then
            wsControl.Range("$A$2:$H$34").Value2 = wsControl.Range("A" & startRow).Resize(33, 8).Value2
        End If

        ' Force a single full recalc 
        Application.Calculate

        ' "1 MACD" values
        wsMACD.Range("K" & destRowMACD & ":P" & destRowMACD).Value2 = wsMACD.Range("$V$6:$AA$6").Value2
        wsMACD.Range("Q" & destRowMACD & ":V" & destRowMACD).Value2 = wsMACD.Range("$V$26:$AA$26").Value2
        wsMACD.Range("W" & destRowMACD & ":AB" & destRowMACD).Value2 = wsMACD.Range("$AM$6:$AR$6").Value2
        wsMACD.Range("AC" & destRowMACD & ":AH" & destRowMACD).Value2 = wsMACD.Range("$AM$26:$AR$26").Value2
        wsMACD.Range("AI" & destRowMACD & ":AN" & destRowMACD).Value2 = wsMACD.Range("$BD$6:$BI$6").Value2
        wsMACD.Range("AO" & destRowMACD & ":AT" & destRowMACD).Value2 = wsMACD.Range("$BD$26:$BI$26").Value2
        wsMACD.Range("AU" & destRowMACD & ":AZ" & destRowMACD).Value2 = wsMACD.Range("$BU$6:$BZ$6").Value2
        wsMACD.Range("BA" & destRowMACD & ":BF" & destRowMACD).Value2 = wsMACD.Range("$BU$26:$BZ$26").Value2
        wsMACD.Range("J" & destRowMACD).Value2 = wsMACD.Range("$A$2").Value2
        wsMACD.Range("I" & destRowMACD).Value2 = wsMACD.Range("$A$3").Value2
        wsMACD.Range("BG" & destRowMACD & ":BL" & destRowMACD).Value2 = wsMACD.Range("$B$3:$G$3").Value2

        ' "2 EMA" values
        wsEMA.Range("K" & destRowEMA & ":P" & destRowEMA).Value2 = wsEMA.Range("$V$3:$AA$3").Value2
        wsEMA.Range("Q" & destRowEMA & ":V" & destRowEMA).Value2 = wsEMA.Range("$V$23:$AA$23").Value2
        wsEMA.Range("W" & destRowEMA & ":AB" & destRowEMA).Value2 = wsEMA.Range("$AI$3:$AN$3").Value2
        wsEMA.Range("AC" & destRowEMA & ":AH" & destRowEMA).Value2 = wsEMA.Range("$AI$23:$AN$23").Value2
        wsEMA.Range("AI" & destRowEMA & ":AN" & destRowEMA).Value2 = wsEMA.Range("$AV$3:$BA$3").Value2
        wsEMA.Range("AO" & destRowEMA & ":AT" & destRowEMA).Value2 = wsEMA.Range("$AV$23:$BA$23").Value2
        wsEMA.Range("AU" & destRowEMA & ":AZ" & destRowEMA).Value2 = wsEMA.Range("$BI$3:$BN$3").Value2
        wsEMA.Range("BA" & destRowEMA & ":BF" & destRowEMA).Value2 = wsEMA.Range("$BI$23:$BN$23").Value2
        wsEMA.Range("J" & destRowEMA).Value2 = wsEMA.Range("$A$2").Value2
        wsEMA.Range("I" & destRowEMA).Value2 = wsEMA.Range("$A$3").Value2
        wsEMA.Range("BG" & destRowEMA & ":BL" & destRowEMA).Value2 = wsEMA.Range("$B$3:$G$3").Value2

        ' "3 RSI" values
        wsRSI.Range("K" & destRowRSI & ":P" & destRowRSI).Value2 = wsRSI.Range("$U$4:$Z$4").Value2
        wsRSI.Range("Q" & destRowRSI & ":V" & destRowRSI).Value2 = wsRSI.Range("$U$26:$Z$26").Value2
        wsRSI.Range("W" & destRowRSI & ":AB" & destRowRSI).Value2 = wsRSI.Range("$AK$4:$AP$4").Value2
        wsRSI.Range("AC" & destRowRSI & ":AH" & destRowRSI).Value2 = wsRSI.Range("$AK$26:$AP$26").Value2
        wsRSI.Range("AI" & destRowRSI & ":AN" & destRowRSI).Value2 = wsRSI.Range("$BA$4:$BF$4").Value2
        wsRSI.Range("AO" & destRowRSI & ":AT" & destRowRSI).Value2 = wsRSI.Range("$BA$26:$BF$26").Value2
        wsRSI.Range("AU" & destRowRSI & ":AZ" & destRowRSI).Value2 = wsRSI.Range("$BQ$4:$BV$4").Value2
        wsRSI.Range("BA" & destRowRSI & ":BF" & destRowRSI).Value2 = wsRSI.Range("$BQ$26:$BV$26").Value2
        wsRSI.Range("J" & destRowRSI).Value2 = wsRSI.Range("$A$2").Value2
        wsRSI.Range("I" & destRowRSI).Value2 = wsRSI.Range("$A$3").Value2
        wsRSI.Range("BG" & destRowRSI & ":BL" & destRowRSI).Value2 = wsRSI.Range("$B$3:$G$3").Value2

        ' "4 BREAKOUT" values
        wsBreakout.Range("K" & destRowBreakout & ":P" & destRowBreakout).Value2 = wsBreakout.Range("$O$4:$T$4").Value2
        wsBreakout.Range("Q" & destRowBreakout & ":V" & destRowBreakout).Value2 = wsBreakout.Range("$O$18:$T$18").Value2
        wsBreakout.Range("W" & destRowBreakout & ":AB" & destRowBreakout).Value2 = wsBreakout.Range("$Z$4:$AE$4").Value2
        wsBreakout.Range("AC" & destRowBreakout & ":AH" & destRowBreakout).Value2 = wsBreakout.Range("$Z$18:$AA$18").Value2
        wsBreakout.Range("AI" & destRowBreakout & ":AN" & destRowBreakout).Value2 = wsBreakout.Range("$AK$4:$AP$4").Value2
        wsBreakout.Range("AO" & destRowBreakout & ":AT" & destRowBreakout).Value2 = wsBreakout.Range("$AK$18:$AP$18").Value2
        wsBreakout.Range("AU" & destRowBreakout & ":AZ" & destRowBreakout).Value2 = wsBreakout.Range("$AV$4:$BA$4").Value2
        wsBreakout.Range("BA" & destRowBreakout & ":BF" & destRowBreakout).Value2 = wsBreakout.Range("$AV$18:$BA$18").Value2
        wsBreakout.Range("J" & destRowBreakout).Value2 = wsBreakout.Range("$A$2").Value2
        wsBreakout.Range("I" & destRowBreakout).Value2 = wsBreakout.Range("$A$3").Value2
        wsBreakout.Range("BG" & destRowBreakout & ":BL" & destRowBreakout).Value2 = wsBreakout.Range("$B$3:$G$3").Value2

        ' "5 ADX" values
        wsADX.Range("K" & destRowADX & ":P" & destRowADX).Value2 = wsADX.Range("$AB$4:$AG$4").Value2
        wsADX.Range("Q" & destRowADX & ":V" & destRowADX).Value2 = wsADX.Range("$AB$28:$AG$28").Value2
        wsADX.Range("W" & destRowADX & ":AB" & destRowADX).Value2 = wsADX.Range("$AX$4:$BC$4").Value2
        wsADX.Range("AC" & destRowADX & ":AH" & destRowADX).Value2 = wsADX.Range("$AX$28:$BC$28").Value2
        wsADX.Range("AI" & destRowADX & ":AN" & destRowADX).Value2 = wsADX.Range("$BT$4:$BY$4").Value2
        wsADX.Range("AO" & destRowADX & ":AT" & destRowADX).Value2 = wsADX.Range("$BT$28:$BY$28").Value2
        wsADX.Range("AU" & destRowADX & ":AZ" & destRowADX).Value2 = wsADX.Range("$CP$4:$CU$4").Value2
        wsADX.Range("BA" & destRowADX & ":BF" & destRowADX).Value2 = wsADX.Range("$CP$28:$CU$28").Value2
        wsADX.Range("J" & destRowADX).Value2 = wsADX.Range("$A$2").Value2
        wsADX.Range("I" & destRowADX).Value2 = wsADX.Range("$A$3").Value2
        wsADX.Range("BG" & destRowADX & ":BL" & destRowADX).Value2 = wsADX.Range("$B$3:$G$3").Value2

        ' "6 Volatility Measure" values
        wsVolatility.Range("K" & destRowVolatility & ":P" & destRowVolatility).Value2 = wsVolatility.Range("$S$4:$X$4").Value2
        wsVolatility.Range("Q" & destRowVolatility & ":V" & destRowVolatility).Value2 = wsVolatility.Range("$S$30:$X$30").Value2
        wsVolatility.Range("W" & destRowVolatility & ":AB" & destRowVolatility).Value2 = wsVolatility.Range("$AE$4:$AU$4").Value2
        wsVolatility.Range("AC" & destRowVolatility & ":AH" & destRowVolatility).Value2 = wsVolatility.Range("$AE$30:$AU$30").Value2
        wsVolatility.Range("AI" & destRowVolatility & ":AN" & destRowVolatility).Value2 = wsVolatility.Range("$AQ$4:$AU$4").Value2
        wsVolatility.Range("AO" & destRowVolatility & ":AT" & destRowVolatility).Value2 = wsVolatility.Range("$AQ$30:$AU$30").Value2
        wsVolatility.Range("AU" & destRowVolatility & ":AZ" & destRowVolatility).Value2 = wsVolatility.Range("$BC$4:$BH$4").Value2
        wsVolatility.Range("BA" & destRowVolatility & ":BF" & destRowVolatility).Value2 = wsVolatility.Range("$BC$30:$BH$30").Value2
        wsVolatility.Range("J" & destRowVolatility).Value2 = wsVolatility.Range("$A$2").Value2
        wsVolatility.Range("I" & destRowVolatility).Value2 = wsVolatility.Range("$A$3").Value2
        wsVolatility.Range("BG" & destRowVolatility & ":BL" & destRowVolatility).Value2 = wsVolatility.Range("$B$3:$G$3").Value2

        ' "7 SMA" values
        wsSMA.Range("K" & destRowSMA & ":P" & destRowSMA).Value2 = wsSMA.Range("$Q$3:$V$3").Value2
        wsSMA.Range("Q" & destRowSMA & ":V" & destRowSMA).Value2 = wsSMA.Range("$Q$24:$V$24").Value2
        wsSMA.Range("W" & destRowSMA & ":AB" & destRowSMA).Value2 = wsSMA.Range("$AC$3:$AH$3").Value2
        wsSMA.Range("AC" & destRowSMA & ":AH" & destRowSMA).Value2 = wsSMA.Range("$AC$24:$AH$24").Value2
        wsSMA.Range("AI" & destRowSMA & ":AN" & destRowSMA).Value2 = wsSMA.Range("$AO$3:$AT$3").Value2
        wsSMA.Range("AO" & destRowSMA & ":AT" & destRowSMA).Value2 = wsSMA.Range("$AO$24:$AT$24").Value2
        wsSMA.Range("AU" & destRowSMA & ":AZ" & destRowSMA).Value2 = wsSMA.Range("$BA$3:$BF$3").Value2
        wsSMA.Range("BA" & destRowSMA & ":BF" & destRowSMA).Value2 = wsSMA.Range("$BA$24:$BF$24").Value2
        wsSMA.Range("J" & destRowSMA).Value2 = wsSMA.Range("$A$2").Value2
        wsSMA.Range("I" & destRowSMA).Value2 = wsSMA.Range("$A$3").Value2
        wsSMA.Range("BG" & destRowSMA & ":BL" & destRowSMA).Value2 = wsSMA.Range("$B$3:$G$3").Value2

        ' "8 Bollinger Bands" values
        wsBollinger.Range("K" & destRowBollinger & ":P" & destRowBollinger).Value2 = wsBollinger.Range("$T$4:$Y$4").Value2
        wsBollinger.Range("Q" & destRowBollinger & ":V" & destRowBollinger).Value2 = wsBollinger.Range("$T$25:$Y$25").Value2
        wsBollinger.Range("W" & destRowBollinger & ":AB" & destRowBollinger).Value2 = wsBollinger.Range("$AJ$4:$AO$4").Value2
        wsBollinger.Range("AC" & destRowBollinger & ":AH" & destRowBollinger).Value2 = wsBollinger.Range("$AH$25:$AM$25").Value2
        wsBollinger.Range("AI" & destRowBollinger & ":AN" & destRowBollinger).Value2 = wsBollinger.Range("$AW$4:$BB$4").Value2
        wsBollinger.Range("AO" & destRowBollinger & ":AT" & destRowBollinger).Value2 = wsBollinger.Range("$AW$25:$BB$25").Value2
        wsBollinger.Range("AU" & destRowBollinger & ":AZ" & destRowBollinger).Value2 = wsBollinger.Range("$BJ$4:$BO$4").Value2
        wsBollinger.Range("BA" & destRowBollinger & ":BF" & destRowBollinger).Value2 = wsBollinger.Range("$BJ$25:$BO$25").Value2
        wsBollinger.Range("J" & destRowBollinger).Value2 = wsBollinger.Range("$A$2").Value2
        wsBollinger.Range("I" & destRowBollinger).Value2 = wsBollinger.Range("$A$3").Value2
        wsBollinger.Range("BG" & destRowBollinger & ":BL" & destRowBollinger).Value2 = wsBollinger.Range("$B$3:$G$3").Value2

        ' "9 EMA & MACD" values
        wsEMAMACD.Range("K" & destRowEMAMACD & ":P" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$T$5:$Y$5").Value2
        wsEMAMACD.Range("Q" & destRowEMAMACD & ":V" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$T$22:$Y$22").Value2
        wsEMAMACD.Range("W" & destRowEMAMACD & ":AB" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$AJ$5:$AU$5").Value2
        wsEMAMACD.Range("AC" & destRowEMAMACD & ":AH" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$AJ$22:$AU$22").Value2
        wsEMAMACD.Range("AI" & destRowEMAMACD & ":AN" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$AY$5:$BD$5").Value2
        wsEMAMACD.Range("AO" & destRowEMAMACD & ":AT" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$AY$22:$BD$22").Value2
        wsEMAMACD.Range("AU" & destRowEMAMACD & ":AZ" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$BN$5:$BS$5").Value2
        wsEMAMACD.Range("BA" & destRowEMAMACD & ":BF" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$BN$22:$BS$22").Value2
        wsEMAMACD.Range("J" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$A$2").Value2
        wsEMAMACD.Range("I" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$A$3").Value2
        wsEMAMACD.Range("BG" & destRowEMAMACD & ":BL" & destRowEMAMACD).Value2 = wsEMAMACD.Range("$B$3:$G$3").Value2

        ' "10 RSI Bollinger" values
        wsRSIBollinger.Range("K" & destRowRSIBollinger & ":P" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$AC$4:$AH$4").Value2
        wsRSIBollinger.Range("Q" & destRowRSIBollinger & ":V" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$AC$23:$AH$23").Value2
        wsRSIBollinger.Range("W" & destRowRSIBollinger & ":AB" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$AZ$4:$BE$4").Value2
        wsRSIBollinger.Range("AC" & destRowRSIBollinger & ":AH" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$AZ$23:$BE$23").Value2
        wsRSIBollinger.Range("AI" & destRowRSIBollinger & ":AN" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$BW$4:$CB$4").Value2
        wsRSIBollinger.Range("AO" & destRowRSIBollinger & ":AT" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$BW$23:$CB$23").Value2
        wsRSIBollinger.Range("AU" & destRowRSIBollinger & ":AZ" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$CT$4:$CY$4").Value2
        wsRSIBollinger.Range("BA" & destRowRSIBollinger & ":BF" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$CT$23:$CY$23").Value2
        wsRSIBollinger.Range("J" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$A$2").Value2
        wsRSIBollinger.Range("I" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$A$3").Value2
        wsRSIBollinger.Range("BG" & destRowRSIBollinger & ":BL" & destRowRSIBollinger).Value2 = wsRSIBollinger.Range("$B$3:$G$3").Value2

        ' Increment the destination row for the next batch
        destRowMACD = destRowMACD + 1
        destRowEMA = destRowEMA + 1
        destRowRSI = destRowRSI + 1
        destRowBreakout = destRowBreakout + 1
        destRowADX = destRowADX + 1
        destRowVolatility = destRowVolatility + 1
        destRowSMA = destRowSMA + 1
        destRowBollinger = destRowBollinger + 1
        destRowEMAMACD = destRowEMAMACD + 1
        destRowRSIBollinger = destRowRSIBollinger + 1

    Next i

    Worksheets("ControlSheet").Range("O1").Value2 = Time

    ' Restore the first dataset after cycling
    If Not IsEmpty(originalData) Then
        wsControl.Range("A2:H34").Value2 = originalData
    End If

SafeExit:
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayStatusBar = prevStatusBar
End Sub

