VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim I As Integer
    Dim sheetName As String
    Dim openFound As Boolean
    Dim openVal As Double
    Dim closeVal As Double
    
    Dim openCloseDiff() As Double, size As Integer, idx As Integer
    Dim ticker() As String
    Dim currTicker As String
    
    Dim printRow As Integer
    Dim percentChange As Double
    Dim currStockVolume As Double
    
    Dim maxChange As Double
    Dim minChange As Double
    Dim maxVolume As Double
    Dim maxChangeTicket As String
    Dim minChangeTicket As String
    Dim maxVolumeTicket As String
    
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    
    For I = 1 To WS_Count
    maxChange = -9999999
    minChange = 9999999
    maxVolume = 0
        
        ActiveWorkbook.Worksheets(I).Cells(1, 9) = "Ticker"
        ActiveWorkbook.Worksheets(I).Cells(1, 10).Value = "Yearly Change"
        ActiveWorkbook.Worksheets(I).Cells(1, 11).Value = "Percent Change"
        ActiveWorkbook.Worksheets(I).Cells(1, 12).Value = "Total Stock Volume"
        ActiveWorkbook.Worksheets(I).Cells(1, 16).Value = "Ticker"
        ActiveWorkbook.Worksheets(I).Cells(1, 17).Value = "Value"
                        
        openFound = False
        printRow = 2
        
        sheetName = ActiveWorkbook.Worksheets(I).Name
        
        For Row = 2 To ActiveWorkbook.Worksheets(I).UsedRange.Rows.Count
            For Column = 1 To 1
                Set curCell = ActiveWorkbook.Worksheets(I).Cells(Row, Column)
                If openFound = False Then
                    openVal = ActiveWorkbook.Worksheets(I).Cells(Row, 3)
                    openFound = True
                    currTicker = ActiveWorkbook.Worksheets(I).Cells(Row, 1).Value
                    currStockVolume = ActiveWorkbook.Worksheets(I).Cells(Row, 7).Value
                Else
                    If StrComp(ActiveWorkbook.Worksheets(I).Cells(Row, 1).Value, currTicker) <> 0 Then
                        closeVal = ActiveWorkbook.Worksheets(I).Cells(Row - 1, 6)
                        
                        Debug.Print (ActiveWorkbook.Worksheets(I).Cells(Row - 1, 1))
                        Debug.Print (closeVal - openVal)
                        ActiveWorkbook.Worksheets(I).Cells(printRow, 9) = ActiveWorkbook.Worksheets(I).Cells(Row - 1, 1)
                        ActiveWorkbook.Worksheets(I).Cells(printRow, 10).Value = closeVal - openVal
                        If (closeVal - openVal) < 0 Then
                            ActiveWorkbook.Worksheets(I).Cells(printRow, 10).Interior.Color = RGB(255, 0, 0)
                        Else
                            ActiveWorkbook.Worksheets(I).Cells(printRow, 10).Interior.Color = RGB(0, 255, 0)
                        End If
                        
                        If openVal = 0 Then
                            percentChange = 100
                        Else
                            percentChange = ((closeVal - openVal) * 100) / openVal
                        End If
                        
                        ActiveWorkbook.Worksheets(I).Cells(printRow, 11).Value = CStr(percentChange) + "%"
                        ActiveWorkbook.Worksheets(I).Cells(printRow, 12).Value = currStockVolume
                        
                        If maxChange < percentChange Then
                            maxChange = percentChange
                            maxChangeTicket = currTicker
                        End If
                        
                        If minChange > percentChange Then
                            minChange = percentChange
                            minChangeTicket = currTicker
                        End If
                        
                        If maxVolume < currStockVolume Then
                            maxVolume = currStockVolume
                            maxVolumeTicket = currTicker
                        End If
                        
                        printRow = printRow + 1
                        openVal = ActiveWorkbook.Worksheets(I).Cells(Row, 3)
                        currTicker = ActiveWorkbook.Worksheets(I).Cells(Row, 1).Value
                        openFound = True
                        currStockVolume = ActiveWorkbook.Worksheets(I).Cells(Row, 7).Value
                    Else
                        currStockVolume = currStockVolume + ActiveWorkbook.Worksheets(I).Cells(Row, 7).Value
                    End If
                End If
            Next Column
        Next Row
        
        closeVal = ActiveWorkbook.Worksheets(I).Cells(Row - 1, 6)
        Debug.Print (ActiveWorkbook.Worksheets(I).Cells(Row - 1, 1))
        Debug.Print (closeVal - openVal)
        ActiveWorkbook.Worksheets(I).Cells(printRow, 9) = ActiveWorkbook.Worksheets(I).Cells(Row - 1, 1)
        ActiveWorkbook.Worksheets(I).Cells(printRow, 10).Value = closeVal - openVal
        
        If (closeVal - openVal) < 0 Then
            ActiveWorkbook.Worksheets(I).Cells(printRow, 10).Interior.Color = RGB(255, 0, 0)
        Else
            ActiveWorkbook.Worksheets(I).Cells(printRow, 10).Interior.Color = RGB(0, 255, 0)
        End If
        
        If openVal = 0 Then
            percentChange = 100
        Else
            percentChange = ((closeVal - openVal) * 100) / openVal
        End If
        
        ActiveWorkbook.Worksheets(I).Cells(printRow, 11).Value = CStr(percentChange) + "%"
        ActiveWorkbook.Worksheets(I).Cells(printRow, 12).Value = currStockVolume
        
        If maxChange < percentChange Then
            maxChange = percentChange
            maxChangeTicket = currTicker
        End If
        
        If minChange > percentChange Then
            minChange = percentChange
            minChangeTicket = currTicker
        End If
        
        If maxVolume < currStockVolume Then
            maxVolume = currStockVolume
            maxVolumeTicket = currTicker
        End If
        
        printRow = printRow + 1
        
        ActiveWorkbook.Worksheets(I).Cells(2, 15).Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(I).Cells(2, 16).Value = maxChangeTicket
        ActiveWorkbook.Worksheets(I).Cells(2, 17).Value = CStr(maxChange) + "%"
        ActiveWorkbook.Worksheets(I).Cells(3, 15).Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(I).Cells(3, 16).Value = minChangeTicket
        ActiveWorkbook.Worksheets(I).Cells(3, 17).Value = CStr(minChange) + "%"
        ActiveWorkbook.Worksheets(I).Cells(4, 15).Value = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(I).Cells(4, 16).Value = maxVolumeTicket
        ActiveWorkbook.Worksheets(I).Cells(4, 17).Value = maxVolume

    Next I

End Sub



