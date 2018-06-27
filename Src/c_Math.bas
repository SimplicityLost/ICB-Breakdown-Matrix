Attribute VB_Name = "c_Math"
Function YtD(workingws As Worksheet, year As Long, month As Integer)
    Dim ytdval
    workingws.Range("T1:U1") = Split("Year|Month", "|")
    workingws.Range("T2").Value = year
    workingws.Range("U2").Value = "<=" & month
    workingws.Range("A:Q").AdvancedFilter _
        Action:=xlFilterInPlace, _
        criteriarange:=workingws.Range("T1:U2")
    ytdval = Application.WorksheetFunction.Sum(workingws.Range("M:M").SpecialCells(xlCellTypeVisible))
    workingws.Cells.AutoFilter
    YtD = ytdval
End Function

Function xmonthtrend(workingws As Worksheet, datein, trend As Integer)
    Dim mnthtotal()
    ReDim mnthtotal(0 To (trend - 1))
    Dim i
    workingws.Range("T1:U1") = Split("Year|Month", "|")
    
    For i = 0 To trend - 1

            workingws.Range("T2").Value = year(DateAdd("m", -i, datein)) - 2000
            workingws.Range("U2").Value = month(DateAdd("m", -i, datein))
            workingws.Range("A:Q").AdvancedFilter _
                Action:=xlFilterInPlace, _
                criteriarange:=workingws.Range("T1:U2")
            mnthtotal(i) = Application.WorksheetFunction.Sum(workingws.Range("M:M").SpecialCells(xlCellTypeVisible))
        workingws.Cells.AutoFilter
    Next i
    xmonthtrend = mnthtotal()
End Function
