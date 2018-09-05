Attribute VB_Name = "d_ReportMaker"
Sub ReportMaker(storenum As String, datein)
    Dim workingws As Worksheet, reportws As Worksheet, palettews As Worksheet
    Dim rprtin, rprtend
    Dim critarray, ventype
    Set workingws = ThisWorkbook.Sheets("working")
    Set reportws = ThisWorkbook.Sheets("report")
    Set palettews = ThisWorkbook.Sheets("palette")
    
    On Error Resume Next
    reportws.Range("A:AA").Ungroup
    On Error GoTo 0
    
    critarray = Split("NC/CI/MF/OTHER", "/")
    palettews.Range("Z1") = "Type"
    rprtin = 9
   
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    reportheaders = Split("Vendor Name|Description|Contact Person|Contact Info|" _
        & MonthName(month(DateAdd("m", -1, datein))) & "|" & MonthName(month(DateAdd("m", -2, datein))) & "|" & MonthName(month(DateAdd("m", -3, datein))) & "|" _
        & year(DateAdd("m", -1, datein)) & "|" & year(DateAdd("m", -1, datein)) - 1 & "|YoY", "|")
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    
    reportws.Range("B6:K6") = reportheaders
    reportws.Range("B5").Value = "Vendor Information"
    reportws.Range("G5").Value = "3 Month Trend"
    reportws.Range("J5").Value = "Annual Trend - YTD"
    
    workingws.Columns("H").Insert
    workingws.Columns("E").Insert
    reportws.Columns("I").Insert
    reportws.Columns("F").Insert
    
    With reportws
        Call formatreport(.Range("D5:E6"), 1)
        Call formatreport(.Range("B5:C6"), 1)
        Call formatreport(.Range("G5:I6"), 1)
        Call formatreport(.Range("K5:M6"), 1)
        With .Range("M6").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .color = -4210753
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    End With
    
    
    For i = 0 To UBound(critarray)
        palettews.Range("z2") = critarray(i)
        
        Select Case i
            Case 0
                typetitle = "National Contract"
            Case 1
                typetitle = "Consolidated Invoices"
            Case 2
                typetitle = "Management Fees"
            Case 3
                typetitle = "Other Fees"
        End Select
        
        If workingws.Range("M:M").Find(critarray(i)) Is Nothing Then
            workingws.Range("A500").Value = "None Found"
            workingws.Range("M500").Value = critarray(i)
        End If
        reportws.Range("B" & rprtin - 1).Value = typetitle
        workingws.Range("A1:M500").AdvancedFilter _
            Action:=xlFilterInPlace, _
            criteriarange:=palettews.Range("Z1:Z2")
        workingws.Range("A2:L500").SpecialCells(xlCellTypeVisible).Copy
        reportws.Range("B" & rprtin).PasteSpecial xlPasteAll
        workingws.Range("A:Z").AutoFilter
        Application.CutCopyMode = False
        rprtend = reportws.Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False) _
                .Row
                       
        With reportws
 
            Call formatreport(.Range("D" & rprtin - 1 & ":E" & rprtend), 2)
            Call formatreport(.Range("B" & rprtin - 1 & ":C" & rprtend), 2)
            
            Call formatreport(.Range("g" & rprtin - 1 & ":i" & rprtend), 2)
            Call formatreport(.Range("g" & rprtin & ":g" & rprtend), 4)
            Call formatreport(.Range("k" & rprtin - 1 & ":m" & rprtend), 2)
            Call formatreport(.Range("M" & rprtin & ":m" & rprtend), 3)
        End With
                       
                       
        rprtin = rprtend + 3
             
        
    Next i
    
    reportws.Columns("B:M").AutoFit
    reportws.Columns("F").ColumnWidth = 3
    reportws.Columns("J").ColumnWidth = 3
    
    If storenum = "*" Then
        reportws.Range("B2").Value = "ICB Breakdown Matrix for All Stores"
    Else
        reportws.Range("B2").Value = "ICB Breakdown Matrix for " & storenum
    End If
    reportws.Range("B2").style = "Title"
    
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    reportws.Range("C3").Value = "For Period Ending " & month(DateAdd("m", -1, datein)) & "/" & year(DateAdd("m", -1, datein))
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    reportws.Range("C3").style = "Heading 4"
    reportws.Range("G:L").style = "currency"
    reportws.Range("M:M").style = "Percent"
    
    reportws.Range("D:E").Group
    reportws.Outline.ShowLevels ColumnLevels:=1
    
    reportws.Columns("O:BDM").Delete
    
    reportws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    reportws.Range("a1").Select
End Sub
Sub FileWriter(storenum As String, datein)
Dim w As Workbook
Dim reportwb As Workbook
Set reportwb = ActiveWorkbook
Set w = Application.Workbooks.Add

    prevdate = WorksheetFunction.EDate(datein, -1)
    
    reportwb.Sheets("Working2").Copy _
        Before:=w.Sheets(1)
    w.Sheets("Working2").Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    w.Sheets("Working2").name = "Source Data"
    
    reportwb.Sheets("Report").Copy _
        Before:=w.Sheets(1)
        
    Application.DisplayAlerts = False
    
    For Each Sheet In w.Worksheets
        If InStr(Sheet.name, "Sheet") > 0 Then
            Sheet.Delete
        End If
    Next Sheet
    
'    For s = 1 To 3
'        w.Sheets("Sheet" & s).Delete
'    Next s
     Application.DisplayAlerts = True
    
    If storenum = "*" Then
        flnm = "All Stores"
    Else
        flnm = CStr(storenum)
    End If
    
    filepath = "\\Ntoscar\T-Drive\Accounts Payable\Procurement\Procurement Analyst Projects\Corey's Projects\ICB Report Project\ICB Reports " & month(prevdate) & "-" & year(prevdate)
    Call CreateDir(CStr(filepath))
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    w.SaveAs Filename:=filepath & "\" & flnm & " ICB Report (Up to " & month(prevdate) & "-" & year(prevdate) & ")"
    'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    w.Close
End Sub

Sub formatreport(rangein As Range, style As Integer)
    Select Case style
    'Header row format
        Case 1
                With rangein.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .color = 13734656
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With rangein.Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                rangein.Font.Bold = True
                With rangein
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = False
                End With
                rangein.Borders(xlDiagonalDown).LineStyle = xlNone
                rangein.Borders(xlDiagonalUp).LineStyle = xlNone
                With rangein.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                With rangein.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                With rangein.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                With rangein.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                rangein.Borders(xlInsideVertical).LineStyle = xlNone
                With rangein.Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .color = -16737793
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
    'Content boxes for most of the report
        Case 2
                'Major borders
                With rangein.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                With rangein.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                With rangein.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                With rangein.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlThick
                End With
                
                'Top row format
                With ThisWorkbook.Sheets("report").Cells(rangein.Row, rangein.Column).Resize(, rangein.Columns.Count)
                    With .Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .color = 13734656
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With .Font
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = 0
                        .Bold = True
                    End With
                End With

    'The specific formatting for YoY percentages.
        Case 3
            Dim cfIconSet
            
                With rangein.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .color = -4210753
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With rangein.FormatConditions.Add(xlCellValue, xlNotBetween, -0.25, 0.25)
                    With .Interior
                        .color = 13551615
                    End With
                    With .Font
                        .color = 393372
                    End With
                End With
                With rangein.FormatConditions.Add(xlCellValue, xlBetween, -0.25, 0.25)
                    With .Interior
                        .color = 13561798
                    End With
                    With .Font
                        .color = 24832
                    End With
                End With
 
            Case 4
            
                'With rangein.FormatConditions.Add(xlCellValue, xlEqual, 0)
                With rangein.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(RC[1]<>0,RC=0)")
                  With .Interior
                        .color = 13551615
                    End With
                    With .Font
                        .color = 393372
                    End With
                End With
                
          End Select
End Sub

Sub CreateDir(strPath As String)
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    modstrPath = Right(strPath, Len(strPath) - 10)
    For Each elm In Split(modstrPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir("\\Ntoscar\" & strCheckPath, vbDirectory)) = 0 Then MkDir "\\Ntoscar\" & strCheckPath
    Next
End Sub
