Attribute VB_Name = "e_FileInjector"
Function FileInjector(mnthin, yrin, storenum)
    Dim path, name, missing
    Dim w As Workbook
    Dim reportwb As Workbook

    Set reportwb = ActiveWorkbook
    
    file = Filechooser(mnthin, yrin, storenum)
    
    If Len(file) = 0 Then
        FileInjector = 0
    Else
        Set w = Workbooks.Open(file)

        reportwb.Sheets("Report").Copy _
            After:=w.Sheets(w.Sheets.Count)
        w.Sheets("Report").name = "BREAKDOWN"

        w.Save
        w.Close
        FileInjector = 1

    End If
End Function
Function FilePathGet(mnth, yr)
'\\Ntoscar\Stores\L001 Motors Inter-Company Billing\2017\05May ICB\EOM
Dim thism, thisy, file

If (mnth > 9) Then
thism = UCase(CStr(mnth & " " & MonthName(mnth, True) & " 2018"))
Else
thism = UCase(CStr("0" & mnth & " " & MonthName(mnth, True) & " 2018"))
End If

If Len(yr) = 2 Then
thisy = "ICB 20" & yr
Else
thisy = "ICB " & yr
End If

file = "\\Ntoscar\Stores\L001 Motors Inter-Company Billing\" & yr & "\" & thism & "\EOM"

If Len(Dir(file, vbDirectory)) = 0 Then
    FilePathGet = "\\Ntoscar\Stores\L001 Motors Inter-Company Billing\" & yr & "\" & thism & "\EOM"
Else
    FilePathGet = file
End If

End Function

Function FileNameGet(mnthin, yrin, storenum)
Dim mnth, yr, lastday

If (mnthin > 9) Then
mnth = "" & mnthin
Else
mnth = "0" & mnthin
End If

If Len(yrin) = 2 Then
yr = "" & yrin
Else
yr = Right("" & yrin, 2)
End If

lastday = Split(DateSerial(yrin, mnthin + 1, 0), "/")(1)

FileNameGet = storenum & "-ICB" & mnth & yr & "EOM"

End Function

Function Filechooser(monthin, yearin, storenum)

filepath = FilePathGet(monthin, yearin)
filenm = FileNameGet(monthin, yearin, storenum)

current = Dir(filepath & "\" & filenm & "*")
newest = ""

Do While current <> ""
If newest = "" Then
newest = filepath & "\" & current
ElseIf FileDateTime(newest) < FileDateTime(filepath & "\" & current) Then
newest = filepath & "\" & current
End If

current = Dir
Loop

Filechooser = newest
End Function
