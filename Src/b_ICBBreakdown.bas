Attribute VB_Name = "b_ICBBreakdown"
Option Explicit
Dim workingws As Worksheet 'worksheet where calculations happen
Dim dataws As Worksheet 'worksheet where data from data pull lives
Dim reportws As Worksheet 'worksheet where the final report will output
Dim vendorws As Worksheet 'worksheet that holds vendor info
Dim vendor 'vendor name for loop
Dim reportheaders 'array of headers for the report
Dim i 'row index for pre-report list
Dim vendrow


Function ICBBreakdown(storenum As String, venlist, datein)

    Set workingws = ThisWorkbook.Sheets("Working")
    Set dataws = ThisWorkbook.Sheets("data")
    Set reportws = ThisWorkbook.Sheets("Report")
    Set vendorws = ThisWorkbook.Sheets("Vendors")
    i = 2
    reportws.Cells.Delete
    workingws.Cells.Delete
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    reportheaders = Split("Vendor Name|Description|Contact Person|Contact Info|" _
    & MonthName(month(DateAdd("m", -1, datein))) & "|" & MonthName(month(DateAdd("m", -2, datein))) & "|" & MonthName(month(DateAdd("m", -3, datein))) & "|" _
    & year(DateAdd("m", -1, datein)) & "|" & year(DateAdd("m", -1, datein)) - 1 & "|YoY|Type", "|")
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
    reportws.Range("A1:K1") = reportheaders
    dataws.Range("AG1") = "Store"
    dataws.Range("AG2") = storenum
    dataws.Range("A:Q").AdvancedFilter _
        Action:=xlFilterCopy, _
        criteriarange:=dataws.Range("AG1:Ag2"), _
        CopyToRange:=Sheets("Working2").Range("A:Q")
    
            
    For Each vendor In venlist
        dataws.Range("AH1").Value = "Trimmed Vendor Name"
        dataws.Range("AH2").Value = Trim(vendor)
        dataws.Range("A:Q").AdvancedFilter _
        Action:=xlFilterCopy, _
        criteriarange:=dataws.Range("AG1:AH2"), _
        CopyToRange:=workingws.Range("A:Q")
    If IsEmpty(workingws.Range("a2")) Then
        GoTo gotonextvendor:
    End If
    'pull the description, vendor type, and contact info from the vendor info file
        vendrow = Application.WorksheetFunction.Match(Trim(vendor), vendorws.Range("A:A"), 0)
        If Not IsEmpty(Application.WorksheetFunction.Index(vendorws.Range("G:G"), vendrow)) Then
            reportws.Range("a" & i).Value = Trim(vendor) & "  -  " & Application.WorksheetFunction.Index(vendorws.Range("G:G"), vendrow)
        Else
            reportws.Range("a" & i).Value = Trim(vendor)
        End If
        reportws.Range("b" & i).Value = Application.WorksheetFunction.Index(vendorws.Range("C:C"), vendrow)
        reportws.Range("c" & i).Value = Application.WorksheetFunction.Index(vendorws.Range("D:D"), vendrow)
        reportws.Range("d" & i).Value = Application.WorksheetFunction.Index(vendorws.Range("E:E"), vendrow) & "   |   " & Application.WorksheetFunction.Index(vendorws.Range("F:F"), vendrow)
        reportws.Range("k" & i).Value = Application.WorksheetFunction.Index(vendorws.Range("B:B"), vendrow)
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'calculate the 3 month trend (this month, prev month, prev-1 month)
        reportws.Range("e" & i & ":g" & i) = xmonthtrend(workingws, DateAdd("m", -1, datein), 3)
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    'calculate the YTD, last year's YTD, and the Year over Year
        reportws.Range("h" & i).Value = YtD(workingws, year(DateAdd("m", -1, datein)) - 2000, month(DateAdd("m", -1, datein)))
        reportws.Range("i" & i).Value = YtD(workingws, year(DateAdd("m", -1, datein)) - 2001, month(DateAdd("m", -1, datein)))
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
        If reportws.Range("i" & i).Value <> 0 Then
            reportws.Range("j" & i).Value = (reportws.Range("h" & i).Value - reportws.Range("i" & i).Value) / Abs(reportws.Range("i" & i).Value)
        Else
            reportws.Range("j" & i).Value = 0
        End If
        
    'remove pointless vendors from ugly list
        If Application.WorksheetFunction.Sum(reportws.Range("E" & i & ":I" & i)) = 0 Then
            reportws.Rows(i).Delete
        End If
    
    i = i + 1
gotonextvendor:
    Next vendor
    'sort vendors alphabetically
        reportws.Range("A1:K200").Sort key1:=reportws.Range("A:A"), order1:=xlAscending, Header:=xlYes
        workingws.Cells.Delete
        workingws.Range("A1:Z200").Value = reportws.Range("A1:Z200").Value
        reportws.Cells.Delete
End Function
