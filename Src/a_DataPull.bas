Attribute VB_Name = "a_DataPull"
Option Explicit

Dim data As Worksheet 'worksheet to hold all the data
Dim recordwb As Workbook 'workbook that holds all the compiled data
Dim vendors As Worksheet 'worksheet to hold vendor information
Dim fNameAndPath
Dim n

Function datapuller(storenum) As Boolean
    Set data = ActiveWorkbook.Sheets("Data")
    
    fNameAndPath = Application.GetOpenFilename(FileFilter:="All Files, *", Title:="Where is the consolidated record workbook?")
    If fNameAndPath = False Then
        Set recordwb = Nothing
        datapuller = True
        Exit Function
    Else
        Set recordwb = Workbooks.Open(fNameAndPath, True, True)
    End If
    
    data.Cells.Delete
    n = recordwb.Sheets(1).Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count
    data.Range("AG1").Value = "Store"
    data.Range("AG2").Value = storenum
    recordwb.Sheets("Entry List").Range("A1:Z" & n).AdvancedFilter _
        Action:=xlFilterCopy, _
        criteriarange:=data.Range("AG1:AG2"), _
        CopyToRange:=data.Range("A1")
        
    recordwb.Close False
    Set recordwb = Nothing
    data.Range("A1:Q" & n).Sort key1:=data.Range("C:C"), order1:=xlAscending, Header:=xlYes
    datapuller = False
End Function
Function vendorpuller()
    Set vendors = ActiveWorkbook.Sheets("Vendors")
    
    fNameAndPath = Application.GetOpenFilename(FileFilter:="All Files, *", Title:="Where is the vendor information workbook?")
    If fNameAndPath = False Then
        Set recordwb = Nothing
        vendorpuller = True
        Exit Function
    Else
        Set recordwb = Workbooks.Open(fNameAndPath, True, True)
    End If
    
    vendors.Cells.Clear
    recordwb.Sheets(1).Range("A:Z").Copy Destination:=vendors.Range("a1")
        
    recordwb.Close False
    Set recordwb = Nothing
    vendorpuller = False
End Function
Function UniqueVals(rangein As Range) As Variant

Dim tmp As String
Dim cell

For Each cell In rangein
      If (cell.Value <> "") And (InStr(1, tmp, cell.Value, vbTextCompare) = 0) Then
        tmp = tmp & cell.Value & "|"
      End If
   Next cell

If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - 1)

UniqueVals = Split(tmp, "|")

End Function
