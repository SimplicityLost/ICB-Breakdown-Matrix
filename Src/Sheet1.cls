VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox2_Click()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    Dim cancel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Sheets("report").Cells.Delete
    cancel = datapuller(Range("C3").Value)
    If cancel Then Exit Sub
    cancel = vendorpuller
    If cancel Then Exit Sub
    procdatein = DateAdd("m", 1, CDate(Range("C7").Value & "/1/" & Range("f7").Value))
    Call ICBBreakdown(Range("C3").Value, UniqueVals(Sheets("Data").Range("K2:K" & Sheets("Data").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count)), procdatein)
    Call ReportMaker(Range("C3").Value, procdatein)
    If newfiles.Value Then
        Call FileWriter(Range("C3").Value, procdatein)
    ElseIf inject.Value Then
        Call FileInjector(Range("C7").Value, Range("f7").Value, Range("C3").Value)
    End If
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False
    Dim missingnums
    If reportonly.Value Then
        MsgBox ("You must select either 'Create New Report Files' or 'Insert Reports into ICB Files'")
        Exit Sub
    End If
    Sheets("report").Cells.Delete
    Call datapuller("*")
    Call vendorpuller
    storelist = UniqueVals(Sheets("Data").Range("C2:C700000"))
    venlist = UniqueVals(Sheets("Data").Range("K2:K" & Sheets("Data").Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count))
    procdatein = DateAdd("m", 1, CDate(Range("C7").Value & "/1/" & Range("f7").Value))
    For Each store In storelist
        Call ICBBreakdown(CStr(store), venlist, procdatein)
        Call ReportMaker(CStr(store), procdatein)
        If newfiles.Value Then Call FileWriter(CStr(store), procdatein)
        If inject.Value Then
            If FileInjector(Range("C7").Value, Range("f7").Value, store) = 0 Then
                If injectnew.Value Then
                    Call FileWriter(CStr(store), procdatein)
                Else
                    missingnums = missingnums & store & "|"
                End If
            End If
        End If
    Next store
    
    Application.ScreenUpdating = True
    
    MsgBox "These stores are missing : " & missingnums
End Sub

Private Sub OptionButton1_Click()

End Sub
