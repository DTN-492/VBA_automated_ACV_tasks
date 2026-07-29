# 1_PAX
```vba
Sub Pax()
Dim i As Integer
Dim j As Integer
Dim fr As Integer
Dim lr As Integer
Dim nocode As Integer
Dim code As String
Dim sale As String
Dim sys1 As String
Dim tare As Double
Dim goods As Double

'Clear old check warning
Sheets("PAX").Range("P:AC").EntireColumn.Delete

'Determine 1st row of Sale table
For i = 1 To 100
    If Sheets("PAX").Range("A" & i).Value = "P" Then
        fr = i + 1
        Exit For
    End If
Next i

'Determine last row of Sale table
For i = 1 To 1000
    If Sheets("PAX").Range("B" & i).Value = "TOTAL" Then
        lr = i - 2
        Exit For
    End If
Next i

'HK only check
If Application.WorksheetFunction.CountIfs(Sheets("Sys1").Range("R:R"), "<>", Sheets("Sys1").Range("R:R"), "<>HK") > 0 Then
        With Sheets("PAX").Range("P" & (lr + 2))
            .Value = "Sys1 contains others value than HK!"
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
Else
        With Sheets("PAX").Range("P" & (lr + 2))
            .Value = "Sys1 contains only HK"
            .Font.Color = RGB(128, 128, 128)
            .Font.Italic = True
        End With
End If

'Add title
With Sheets("PAX")
    .Range("P" & (fr - 2)).Value = "Sys1"
    .Range("Q" & (fr - 2)).Value = "Warning"
    .Range("Q" & (fr - 2) & ":R" & (fr - 2)).Merge
    .Range("P" & (fr - 1)).Value = "Nature of Goods"
    .Range("Q" & (fr - 1)).Value = "Sale"
    .Range("R" & (fr - 1)).Value = "Sys1"
    .Range("S" & (fr - 1)).Value = "ELM"
    .Range("T" & (fr - 1)).Value = "DocNo"
    .Range("U" & (fr - 1)).Value = "Weight"
    .Range("W" & (fr - 1)).Value = "ACAS"
    .Range("X" & (fr - 1)).Value = "TSA+MAWB"
    .Range("Y" & (fr - 1)).Value = "ELI"
    .Range("Z" & (fr - 1)).Value = "INTERLINE"
    .Range("AA" & (fr - 1)).Value = "FWB"
    .Range("AB" & (fr - 1)).Value = "FHL"
    .Range("AC" & (fr - 1)).Value = "RCS"
    .Range("P" & (fr - 2) & ":AC" & (fr - 1)).Font.Bold = True
    .Range("P" & (fr - 2) & ":AC" & (fr - 1)).HorizontalAlignment = xlCenter
End With
    
'Vlookup NoG from Sys1
With Sheets("PAX").Range("P" & fr & ":P" & lr)
    .Formula = "=IFERROR(VLOOKUP(B" & fr & ", Sys1!$B$2:$I$1000, 8, 0), """")"
    .Value = .Value
End With

'Warning for NoG
nocode = Application.WorksheetFunction.CountA(Sheets("Code").Range("A:A"))
For i = fr To lr
    For j = 2 To nocode
        code = Sheets("Code").Range("A" & j).Value
        sale = Sheets("PAX").Range("H" & i).Value
        sys1 = Sheets("PAX").Range("P" & i).Value
        
        If InStr(sale, code) = 0 And InStr(sys1, code) > 0 Then
            With Sheets("PAX").Range("Q" & i)
                .Value = Sheets("PAX").Range("Q" & i).Value & " " & code
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        ElseIf InStr(sale, code) > 0 And InStr(sys1, code) = 0 Then
            With Sheets("PAX").Range("R" & i)
                .Value = Sheets("PAX").Range("R" & i).Value & " " & code
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
        
    Next j
Next i

For i = fr To lr
    If Sheets("PAX").Range("Q" & i).Value = "" Then
            With Sheets("PAX").Range("Q" & i)
                .Value = ""
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
    End If
    
    If Sheets("PAX").Range("R" & i).Value = "" Then
            With Sheets("PAX").Range("R" & i)
                .Value = ""
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
    End If
Next i

'Warning if ELM Found
For i = fr To lr
    sale = Sheets("PAX").Range("H" & i).Value
    sys1 = Sheets("PAX").Range("P" & i).Value
    
    If InStr(sale, "ELM") Or InStr(sys1, "ELM") Then
        With Sheets("PAX").Range("S" & i)
            .Value = "ELM FOUND!"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
    Else
        With Sheets("PAX").Range("S" & i)
            .Value = "No ELM Found"
            .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
        End With
    End If
    

Next i

'Warning if Missing DocNo
For i = fr To lr
    j = Application.WorksheetFunction.CountIf(Sheets("Sys1").Range("B2:B1000"), Sheets("PAX").Range("B" & i).Value)
    If j = 0 Then
        With Sheets("PAX").Range("T" & i)
            .Value = "MISSING DOCNO!"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
    Else
        With Sheets("PAX").Range("T" & i)
            .Value = "No Missing DocNo"
            .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
        End With
    End If
    
Next i

'Total est weight
''Vlookup Weight
With Sheets("PAX").Range("U" & fr & ":U" & lr)
    .Formula = "=IFERROR(VLOOKUP(B" & fr & ", Sys1!$B$2:$K$1000, 10, 0), """")"
    .Value = .Value
End With

''Vlookup Pcs
With Sheets("PAX").Range("O" & fr & ":O" & lr)
    .Formula = "=IFERROR(VLOOKUP(B" & fr & ", Sys1!$B$2:$K$1000, 9, 0), """")"
    .Value = .Value
End With

'' Real weight
For i = fr To lr
    sale = Sheets("PAX").Range("H" & i).Value
    sys1 = Sheets("PAX").Range("P" & i).Value
    
    'UPLIFT
    If InStr(sale, "UPLIFT 1P ONLY") Or InStr(sys1, "UPLIFT 1P ONLY") Then
        Sheets("PAX").Range("U" & i).Value = Sheets("PAX").Range("U" & i).Value / Sheets("PAX").Range("O" & i).Value
    ElseIf InStr(sale, "UPLIFT AT LEAST 1P") Or InStr(sys1, "UPLIFT AT LEAST 1P") Then
        Sheets("PAX").Range("U" & i).Value = Sheets("PAX").Range("U" & i).Value / 2
    Else
        Sheets("PAX").Range("U" & i).Value = Sheets("PAX").Range("U" & i).Value
    End If
    
    'MD LD PLA AKE Blank
    If Sheets("PAX").Range("I" & i).MergeArea.Cells(1, 1) = "" And Sheets("PAX").Range("J" & i).MergeArea.Cells(1, 1) = "" And Sheets("PAX").Range("K" & i).MergeArea.Cells(1, 1) = "" And Sheets("PAX").Range("L" & i).MergeArea.Cells(1, 1) = "" Then
        Sheets("PAX").Range("U" & i).Value = 0
    End If
    
    'Clear redundant
    If Sheets("PAX").Range("A" & i).Value = "" Then
        Sheets("PAX").Range("Q" & i).Value = ""
        Sheets("PAX").Range("R" & i).Value = ""
        Sheets("PAX").Range("S" & i).Value = ""
        Sheets("PAX").Range("T" & i).Value = ""
        Sheets("PAX").Range("U" & i).Value = ""
    End If
    Sheets("PAX").Range("O" & i) = ""
Next i

''TARE
Dim md, ld, pla, ake As Integer
md = Sheets("PAX").Range("I" & (lr + 2)).Value
ld = Sheets("PAX").Range("J" & (lr + 2)).Value
pla = Sheets("PAX").Range("K" & (lr + 2)).Value
ake = Sheets("PAX").Range("L" & (lr + 2)).Value

tare = 120 * (md + ld) + 109 * pla + 70 * ake
Sheets("PAX").Range("U" & (lr + 2)).Value = tare

''Goods
goods = Application.WorksheetFunction.Sum(Sheets("PAX").Range("U" & fr & ":U" & lr).Value)
Sheets("PAX").Range("U" & (lr + 3)).Value = tare + goods
Sheets("PAX").Range("T" & (lr + 3)).Value = "est"

'Check table
''ACAS & TSA+MAWB
nocode = Application.WorksheetFunction.CountA(Sheets("Code").Range("B:B"))
For i = fr To lr
    code = Sheets("PAX").Range("D" & i).Value
    If code = "" Then
        Sheets("PAX").Range("W" & i & ":X" & i).Interior.Color = RGB(0, 0, 0)
    ElseIf Application.WorksheetFunction.CountIf(Sheets("Code").Range("B:B"), code) > 0 Then
        Sheets("PAX").Range("W" & i).Value = code
        Sheets("PAX").Range("X" & i).Value = code
    Else
        Sheets("PAX").Range("W" & i & ":X" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''ELI
For i = fr To lr
    If InStr(Sheets("PAX").Range("P" & i).Value, "ELI") Then
        Sheets("PAX").Range("Y" & i).Value = "NO"
    Else
        Sheets("PAX").Range("Y" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''INTERLINE
For i = fr To lr
    If Sheets("PAX").Range("B" & i).Text Like "###-########" Then
        Sheets("PAX").Range("Z" & i).Value = "NO"
    Else
        Sheets("PAX").Range("Z" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''FWB & FHL & RCS
For i = fr To lr
    If Sheets("PAX").Range("B" & i).Text Like "###-########" And Sheets("PAX").Range("B" & i).Value <> "" Then
        Sheets("PAX").Range("AA" & i & ":AC" & i).Value = "NO"
    Else
        Sheets("PAX").Range("AA" & i & ":AC" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''Conditional Formating
With Sheets("PAX").Range("W" & fr & ":AC" & lr)
    .FormatConditions.Delete
    With .FormatConditions.Add(Type:=xlExpression, Formula1:="=W" & fr & "=""OK""")
        .Interior.Color = RGB(255, 255, 0)
        .Font.Color = RGB(255, 165, 0)
        .Font.Bold = True
    End With
End With

'Column Fit
Sheets("PAX").Range("P:AD").EntireColumn.AutoFit

End Sub

```
# 1_PAXclear
```vba
Sub PAXclear()

Dim warn As Integer

warn = MsgBox("Are you sure want to CLEAR ALL", vbYesNo + vbQuestion, "Confirm")

If warn = vbYes Then
    Sheets("PAX").Cells.Clear
Else
    MsgBox "Oke iem"
End If

End Sub
```

# 1_PAXsave
```vba
Sub PAXsave()
    Dim timeStamp As String
    Dim savePath As String
    Dim newWb As Workbook

    timeStamp = UCase(Format(Now, "mmmdd_hhmmss"))
    savePath = Sheets("Code").Range("E1").Value & "\PAX " & timeStamp & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Copies PAX to a brand new, single-sheet workbook
    Sheets("PAX").Copy
    Set newWb = ActiveWorkbook

    ' Save as XLSX and close
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "File saved at " & savePath
End Sub
```

# 2_FRT
```vba
Sub Frt()
Dim i As Integer
Dim j As Integer
Dim fr As Integer
Dim lr As Integer
Dim nocode As Integer
Dim code As String
Dim sale As String
Dim sys1 As String
Dim tare As Double
Dim goods As Double

'Clear old check warning
Sheets("FRT").Range("N:AA").EntireColumn.Delete

'Insert 2 columns
Sheets("FRT").Range("K1:L1").EntireColumn.Insert

'Determine 1st row of Sale table
For i = 1 To 100
    If Sheets("FRT").Range("A" & i).Value = "P" Then
        fr = i + 1
        Exit For
    End If
Next i

'Determine last row of Sale table
For i = 1 To 1000
    If Sheets("FRT").Range("B" & i).Value = "TOTAL" Then
        lr = i - 2
        Exit For
    End If
Next i

'HK only check
If Application.WorksheetFunction.CountIfs(Sheets("Sys1").Range("R:R"), "<>", Sheets("Sys1").Range("R:R"), "<>HK") > 0 Then
        With Sheets("FRT").Range("P" & (lr + 2))
            .Value = "Sys1 contains others value than HK!"
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
Else
        With Sheets("FRT").Range("P" & (lr + 2))
            .Value = "Sys1 contains only HK"
            .Font.Color = RGB(128, 128, 128)
            .Font.Italic = True
        End With
End If

'Add title
With Sheets("FRT")
    .Range("P" & (fr - 2)).Value = "Sys1"
    .Range("Q" & (fr - 2)).Value = "Warning"
    .Range("Q" & (fr - 2) & ":R" & (fr - 2)).Merge
    .Range("P" & (fr - 1)).Value = "Nature of Goods"
    .Range("Q" & (fr - 1)).Value = "Sale"
    .Range("R" & (fr - 1)).Value = "Sys1"
    .Range("S" & (fr - 1)).Value = "ELM"
    .Range("T" & (fr - 1)).Value = "DocNo"
    .Range("U" & (fr - 1)).Value = "Weight"
    .Range("W" & (fr - 1)).Value = "ACAS"
    .Range("X" & (fr - 1)).Value = "TSA+MAWB"
    .Range("Y" & (fr - 1)).Value = "ELI"
    .Range("Z" & (fr - 1)).Value = "INTERLINE"
    .Range("AA" & (fr - 1)).Value = "FWB"
    .Range("AB" & (fr - 1)).Value = "FHL"
    .Range("AC" & (fr - 1)).Value = "RCS"
    .Range("P" & (fr - 2) & ":AC" & (fr - 1)).Font.Bold = True
    .Range("P" & (fr - 2) & ":AC" & (fr - 1)).HorizontalAlignment = xlCenter
End With
    
'Vlookup NoG from Sys1
With Sheets("FRT").Range("P" & fr & ":P" & lr)
    .Formula = "=IFERROR(VLOOKUP(B" & fr & ", Sys1!$B$2:$I$1000, 8, 0), """")"
    .Value = .Value
End With

'Warning for NoG
nocode = Application.WorksheetFunction.CountA(Sheets("Code").Range("A:A"))
For i = fr To lr
    For j = 2 To nocode
        code = Sheets("Code").Range("A" & j).Value
        sale = Sheets("FRT").Range("H" & i).Value
        sys1 = Sheets("FRT").Range("P" & i).Value
        
        If InStr(sale, code) = 0 And InStr(sys1, code) > 0 Then
            With Sheets("FRT").Range("Q" & i)
                .Value = Sheets("FRT").Range("Q" & i).Value & " " & code
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        ElseIf InStr(sale, code) > 0 And InStr(sys1, code) = 0 Then
            With Sheets("FRT").Range("R" & i)
                .Value = Sheets("FRT").Range("R" & i).Value & " " & code
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
        
    Next j
Next i

For i = fr To lr
    If Sheets("FRT").Range("Q" & i).Value = "" Then
            With Sheets("FRT").Range("Q" & i)
                .Value = ""
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
    End If
    
    If Sheets("FRT").Range("R" & i).Value = "" Then
            With Sheets("FRT").Range("R" & i)
                .Value = ""
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
    End If
Next i

'Warning if ELM Found
For i = fr To lr
    sale = Sheets("FRT").Range("H" & i).Value
    sys1 = Sheets("FRT").Range("P" & i).Value
    
    If InStr(sale, "ELM") Or InStr(sys1, "ELM") Then
        With Sheets("FRT").Range("S" & i)
            .Value = "ELM FOUND!"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
    Else
        With Sheets("FRT").Range("S" & i)
            .Value = "No ELM Found"
            .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
        End With
    End If
    

Next i

'Warning if Missing DocNo
For i = fr To lr
    j = Application.WorksheetFunction.CountIf(Sheets("Sys1").Range("B2:B1000"), Sheets("FRT").Range("B" & i).Value)
    If j = 0 Then
        With Sheets("FRT").Range("T" & i)
            .Value = "MISSING DOCNO!"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
    Else
        With Sheets("FRT").Range("T" & i)
            .Value = "No Missing DocNo"
            .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
        End With
    End If
    
Next i

'Total est weight
''Vlookup Weight
With Sheets("FRT").Range("U" & fr & ":U" & lr)
    .Formula = "=IFERROR(VLOOKUP(B" & fr & ", Sys1!$B$2:$K$1000, 10, 0), """")"
    .Value = .Value
End With

''Vlookup Pcs
With Sheets("FRT").Range("O" & fr & ":O" & lr)
    .Formula = "=IFERROR(VLOOKUP(B" & fr & ", Sys1!$B$2:$K$1000, 9, 0), """")"
    .Value = .Value
End With

'' Real weight
For i = fr To lr
    sale = Sheets("FRT").Range("H" & i).Value
    sys1 = Sheets("FRT").Range("P" & i).Value
    
    'UPLIFT
    If InStr(sale, "UPLIFT 1P ONLY") Or InStr(sys1, "UPLIFT 1P ONLY") Then
        Sheets("FRT").Range("U" & i).Value = Sheets("FRT").Range("U" & i).Value / Sheets("FRT").Range("O" & i).Value
    ElseIf InStr(sale, "UPLIFT AT LEAST 1P") Or InStr(sys1, "UPLIFT AT LEAST 1P") Then
        Sheets("FRT").Range("U" & i).Value = Sheets("FRT").Range("U" & i).Value / 2
    Else
        Sheets("FRT").Range("U" & i).Value = Sheets("FRT").Range("U" & i).Value
    End If
    
    'MD LD PLA AKE Blank
    If Sheets("FRT").Range("I" & i).MergeArea.Cells(1, 1) = "" And Sheets("FRT").Range("J" & i).MergeArea.Cells(1, 1) = "" Then
        Sheets("FRT").Range("U" & i).Value = 0
    End If
    
    'Clear redundant
    If Sheets("FRT").Range("A" & i).Value = "" Then
        Sheets("FRT").Range("Q" & i).Value = ""
        Sheets("FRT").Range("R" & i).Value = ""
        Sheets("FRT").Range("S" & i).Value = ""
        Sheets("FRT").Range("T" & i).Value = ""
        Sheets("FRT").Range("U" & i).Value = ""
    End If
    Sheets("FRT").Range("O" & i) = ""
Next i

''TARE
Dim md, ld, pla, ake As Integer
md = Sheets("FRT").Range("I" & (lr + 2)).Value
ld = Sheets("FRT").Range("J" & (lr + 2)).Value

tare = 120 * (md + ld)
Sheets("FRT").Range("U" & (lr + 2)).Value = tare

''Goods
goods = Application.WorksheetFunction.Sum(Sheets("FRT").Range("U" & fr & ":U" & lr).Value)
Sheets("FRT").Range("U" & (lr + 3)).Value = tare + goods
Sheets("FRT").Range("T" & (lr + 3)).Value = "est"

'Check table
''ACAS & TSA+MAWB
nocode = Application.WorksheetFunction.CountA(Sheets("Code").Range("B:B"))
For i = fr To lr
    code = Sheets("FRT").Range("D" & i).Value
    If code = "" Then
        Sheets("FRT").Range("W" & i & ":X" & i).Interior.Color = RGB(0, 0, 0)
    ElseIf Application.WorksheetFunction.CountIf(Sheets("Code").Range("B:B"), code) > 0 Then
        Sheets("FRT").Range("W" & i).Value = code
        Sheets("FRT").Range("X" & i).Value = code
    Else
        Sheets("FRT").Range("W" & i & ":X" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''ELI
For i = fr To lr
    If InStr(Sheets("FRT").Range("P" & i).Value, "ELI") Then
        Sheets("FRT").Range("Y" & i).Value = "NO"
    Else
        Sheets("FRT").Range("Y" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''INTERLINE
For i = fr To lr
    If Sheets("FRT").Range("B" & i).Text Like "###-########" Then
        Sheets("FRT").Range("Z" & i).Value = "NO"
    Else
        Sheets("FRT").Range("Z" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''FWB & FHL & RCS
For i = fr To lr
    If Sheets("FRT").Range("B" & i).Text Like "###-########" And Sheets("FRT").Range("B" & i).Value <> "" Then
        Sheets("FRT").Range("AA" & i & ":AC" & i).Value = "NO"
    Else
        Sheets("FRT").Range("AA" & i & ":AC" & i).Interior.Color = RGB(0, 0, 0)
    End If
Next i

''Conditional Formating
With Sheets("FRT").Range("W" & fr & ":AC" & lr)
    .FormatConditions.Delete
    With .FormatConditions.Add(Type:=xlExpression, Formula1:="=W" & fr & "=""OK""")
        .Interior.Color = RGB(255, 255, 0)
        .Font.Color = RGB(255, 165, 0)
        .Font.Bold = True
    End With
End With

'Column Fit
Sheets("FRT").Range("P:AD").EntireColumn.AutoFit

'Delete 2 column
Range("K1:L1").EntireColumn.Delete

End Sub


```
# 2_FRTclear
```vba
Sub FRTclear()

Dim warn As Integer

warn = MsgBox("Are you sure want to CLEAR ALL", vbYesNo + vbQuestion, "Confirm")

If warn = vbYes Then
    Sheets("FRT").Cells.Clear
Else
    MsgBox "Oke iem"
End If

End Sub
```

# 2_FRTsave
```vba
Sub FRTsave()
    Dim timeStamp As String
    Dim savePath As String
    Dim newWb As Workbook

    timeStamp = UCase(Format(Now, "mmmdd_hhmmss"))
    savePath = Sheets("Code").Range("E1").Value & "\FRT " & timeStamp & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Copies PAX to a brand new, single-sheet workbook
    Sheets("FRT").Copy
    Set newWb = ActiveWorkbook

    ' Save as XLSX and close
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "File saved at " & savePath
End Sub
```

# 3_AWB
```vba
Sub AWB()

Dim lr, frs, lrs As Integer
Dim i, j As Integer

'Clear old check warning
Sheets("AWB").Range("M:W").EntireColumn.Delete

'Determine last row of AWB
lr = Application.WorksheetFunction.Max(Sheets("AWB").Range("A:A")) + 2

'Determine 1st row of Sale table

''Choose PAX or FRT
Dim ch As Integer

ch = MsgBox("Yes for PAX, No for FRT", vbYesNo + vbQuestion, "Sale data today")

If ch = vbYes Then
    Sheets("AWB").Range("M1").Value = "PAX"
Else
    Sheets("AWB").Range("M1").Value = "FRT"
End If

If Sheets("AWB").Range("M1").Value = "PAX" Then
    For i = 1 To 100
        If Sheets("PAX").Range("A" & i).Value = "P" Then
            frs = i + 1
            Exit For
        End If
    Next i
    
    'Determine last row of Sale table
    For i = 1 To 1000
        If Sheets("PAX").Range("B" & i).Value = "TOTAL" Then
            lrs = i - 2
            Exit For
        End If
    Next i
ElseIf Sheets("AWB").Range("m1").Value = "FRT" Then
    For i = 1 To 100
        If Sheets("FRT").Range("A" & i).Value = "P" Then
            frs = i + 1
            Exit For
        End If
    Next i
    
    'Determine last row of Sale table
    For i = 1 To 1000
        If Sheets("FRT").Range("B" & i).Value = "TOTAL" Then
            lrs = i - 2
            Exit For
        End If
    Next i
Else
    MsgBox "Please enter in cell M1 PAX or FRT"
End If

'Add title
With Sheets("AWB")
    .Range("M2").Value = "Missing Doc"
    .Range("O2").Value = "Total Pcs"
    .Range("P2").Value = "Total Weight"
    .Range("Q2").Value = "Remain Pcs"
    .Range("R2").Value = "Remain Weight"
    .Range("U1").Value = "Sys1"
    .Range("U2").Value = "Nature of Goods"
    .Range("V2").Value = "AWB"
    .Range("W2").Value = "Sys1"
    .Range("V1").Value = "Warning"
    .Range("V1:W1").Merge
End With

'Split DocNo
For i = 3 To lr
    Sheets("AWB").Range("B" & i).Value = Format(Sheets("AWB").Range("B" & i).Value, "###-########")
Next i

'Check missing DocNo
If Sheets("AWB").Range("M1").Value = "PAX" Then
    For i = frs To lrs
        If Application.WorksheetFunction.CountIfs(Sheets("AWB").Range("B3:" & "B" & lr), Sheets("PAX").Range("B" & i)) = 0 Then
            With Sheets("AWB").Range("M" & (i - frs + 3))
                .Value = Sheets("PAX").Range("B" & i).Value
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
    Next i
ElseIf Sheets("AWB").Range("M1").Value = "FRT" Then
    For i = frs To lrs
        If Application.WorksheetFunction.CountIfs(Sheets("AWB").Range("B3:" & "B" & lr), Sheets("FRT").Range("B" & i)) = 0 Then
            With Sheets("AWB").Range("M" & (i - frs + 3))
                .Value = Sheets("FRT").Range("B" & i).Value
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
    Next i
Else
    MsgBox "Please enter in cell M1 PAX or FRT"
End If

If Application.WorksheetFunction.CountA(Sheets("AWB").Range("M3:" & "M" & (lr + 3))) = 0 Then
    With Sheets("AWB").Range("M3")
        .Value = "No missing found"
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
End If

'Fill OffBal & Partial
For i = 3 To lr
    If Application.WorksheetFunction.CountIfs(Sheets("Offbal_Partial").Range("A:A"), Sheets("AWB").Range("B" & i)) = 1 Then
        Sheets("AWB").Range("B" & i).Interior.Color = RGB(91, 155, 213)
    ElseIf Application.WorksheetFunction.CountIfs(Sheets("Offbal_Partial").Range("B:B"), Sheets("AWB").Range("B" & i)) = 1 Then
        Sheets("AWB").Range("B" & i).Interior.Color = RGB(255, 255, 0)
    End If
Next i

'Concat DocNo
i = 2
While Sheets("Sys2").Range("F" & i) <> ""
    Sheets("Sys2").Range("G" & i).Value = Sheets("Sys2").Range("E" & i).Value & "-" & Sheets("Sys2").Range("F" & i).Value
    i = i + 1
Wend

'Vlookup Total & Remain
''Total pcs
With Sheets("AWB").Range("O3" & ":O" & lr)
    .Formula = "=IFERROR(VLOOKUP(B3, Sys2!$G$2:$M$1000, 6, 0), """")"
    .Value = .Value
End With

''Total weight
With Sheets("AWB").Range("P3" & ":P" & lr)
    .Formula = "=IFERROR(VLOOKUP(B3, Sys2!$G$2:$M$1000, 7, 0), """")"
    .Value = .Value
End With

''Remian pcs
With Sheets("AWB").Range("Q3" & ":Q" & lr)
    .Formula = "=IFERROR(VLOOKUP(B3, Sys2!$G$2:$R$1000, 11, 0), """")"
    .Value = .Value
End With

''Remian weight
With Sheets("AWB").Range("R3" & ":R" & lr)
    .Formula = "=IFERROR(VLOOKUP(B3, Sys2!$G$2:$R$1000, 12, 0), """")"
    .Value = .Value
End With

'Warning
For i = 3 To lr
    'Total pcs vs SHC
    If Sheets("AWB").Range("O" & i).Value <> Sheets("AWB").Range("G" & i).Value Then
        Sheets("AWB").Range("O" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("AWB").Range("O" & i).Font.Color = RGB(0, 0, 0)
    End If
    
    'Total weight vs NATURE
    If Sheets("AWB").Range("P" & i).Value <> Sheets("AWB").Range("H" & i).Value Then
        Sheets("AWB").Range("P" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("AWB").Range("O" & i).Font.Color = RGB(0, 0, 0)
    End If
    
    'Remain pcs vs PIECES
    If Sheets("AWB").Range("Q" & i).Value <> Sheets("AWB").Range("E" & i).Value Then
        Sheets("AWB").Range("Q" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("AWB").Range("O" & i).Font.Color = RGB(0, 0, 0)
    End If
    
    'Remain weight vs NATURE
    If Sheets("AWB").Range("R" & i).Value <> Sheets("AWB").Range("H" & i).Value Then
        Sheets("AWB").Range("R" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("AWB").Range("O" & i).Font.Color = RGB(0, 0, 0)
    End If
Next i

'Warning code
Dim o, p, q, r As Long
For i = 3 To lr
    o = Sheets("AWB").Range("O" & i).Font.Color
    p = Sheets("AWB").Range("P" & i).Font.Color
    q = Sheets("AWB").Range("Q" & i).Font.Color
    r = Sheets("AWB").Range("R" & i).Font.Color

    If Sheets("AWB").Range("B" & i).Interior.Color = RGB(91, 155, 213) Then 'OffBal
        If o = RGB(0, 0, 0) And p = RGB(255, 0, 0) And q = RGB(0, 0, 0) And r = RGB(0, 0, 0) Then
            With Sheets("AWB").Range("S" & i)
                .Value = "OK"
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
        Else
            With Sheets("AWB").Range("S" & i)
                .Value = "STH WRONG!"
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
    ElseIf Sheets("AWB").Range("B" & i).Interior.Color = RGB(255, 255, 0) Then 'Partial
        If o = RGB(0, 0, 0) And p = RGB(255, 0, 0) Then
            With Sheets("AWB").Range("S" & i)
                .Value = "OK"
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
        Else
            With Sheets("AWB").Range("S" & i)
                .Value = "STH WRONG!"
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
    ElseIf o = RGB(0, 0, 0) And p = RGB(0, 0, 0) Then
        With Sheets("AWB").Range("S" & i)
            .Value = "OK"
            .Font.Italic = True
            .Font.Color = RGB(128, 128, 128)
        End With
    Else
        With Sheets("AWB").Range("S" & i)
            .Value = "STH WRONG!"
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
        End With
    End If
Next i

'Check Nature of Goods
''Vlookup NoG from Sys1
With Sheets("AWB").Range("U3" & ":U" & lr)
    .Formula = "=IFERROR(VLOOKUP(B3, Sys1!$B$2:$I$1000, 8, 0), """")"
    .Value = .Value
End With

''Warning for NoG
Dim nocode As Integer
Dim sale, sys1 As String

nocode = Application.WorksheetFunction.CountA(Sheets("Code").Range("A:A"))
For i = 3 To lr
    For j = 2 To nocode
        code = Sheets("Code").Range("A" & j).Value
        sale = Sheets("AWB").Range("I" & i).Value
        sys1 = Sheets("AWB").Range("U" & i).Value
        
        If InStr(sale, code) = 0 And InStr(sys1, code) > 0 Then
            With Sheets("AWB").Range("V" & i)
                .Value = Sheets("AWB").Range("V" & i).Value & " " & code
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        ElseIf InStr(sale, code) > 0 And InStr(sys1, code) = 0 Then
            With Sheets("AWB").Range("W" & i)
                .Value = Sheets("AWB").Range("W" & i).Value & " " & code
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
        
    Next j
Next i

For i = 3 To lr
    If Sheets("AWB").Range("V" & i).Value = "" Then
            With Sheets("AWB").Range("V" & i)
                .Value = "OK"
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
    End If
    
    If Sheets("AWB").Range("W" & i).Value = "" Then
            With Sheets("AWB").Range("W" & i)
                .Value = "OK"
                .Font.Italic = True
                .Font.Color = RGB(128, 128, 128)
            End With
    End If
Next i

'Auto fit
Sheets("AWB").Range("M:W").EntireColumn.AutoFit

End Sub

```
# 3_AWBclear
```vba
Sub AWBclear()

Dim warn As Integer

warn = MsgBox("Are you sure want to CLEAR ALL", vbYesNo + vbQuestion, "Confirm")

If warn = vbYes Then
    Sheets("AWB").Cells.Clear
Else
    MsgBox "Oke iem"
End If

End Sub
```

# 3_AWBsave
```vba
Sub AWBsave()
    Dim timeStamp As String
    Dim savePath As String
    Dim newWb As Workbook

    timeStamp = UCase(Format(Now, "mmmdd_hhmmss"))
    savePath = Sheets("Code").Range("E1").Value & "\AWB " & timeStamp & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Copies PAX to a brand new, single-sheet workbook
    Sheets("AWB").Copy
    Set newWb = ActiveWorkbook

    ' Save as XLSX and close
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "File saved at " & savePath
End Sub
```

# 4_UWS
```vba
Sub UWS()
Dim i As Integer
Dim md, ld, bk As Integer

'Determine MD, LD, BULK title row
For i = 1 To 100
    If LCase(Sheets("Final UWS").Range("A" & i).Value) = "main deck" Then
        md = i
    ElseIf LCase(Sheets("Final UWS").Range("A" & i).Value) = "low deck" Then
        ld = i
    ElseIf LCase(Sheets("Final UWS").Range("A" & i).Value) = "bulk" Then
        bk = i
    End If
Next i

'Add title
With Sheets("Final UWS")
    .Range("Q" & (md - 1)).Value = "Pallet"
    .Range("Q" & (md - 1) & ":V" & (md - 1)).Merge
    .Range("Q" & md).Value = "Unit Number"
    .Range("R" & md).Value = "Gross Wt"
    .Range("S" & md).Value = "Contour"
    .Range("T" & md).Value = "Special Code"
    .Range("U" & md).Value = "CMDT Code"
    .Range("Q" & (md - 1) & ":V" & md).HorizontalAlignment = xlCenter
    .Range("Q" & (md - 1) & ":V" & md).Font.Bold = True
End With

'Check Contour code sufficient
Dim q4, q5, q6, pwg, pld As Integer

q4 = Application.WorksheetFunction.CountIfs(Sheets("Final UWS").Range("K" & (md + 1) & ":K" & (ld - 1)), "Q4")
q5 = Application.WorksheetFunction.CountIfs(Sheets("Final UWS").Range("K" & (md + 1) & ":K" & (ld - 1)), "Q5")
q6 = Application.WorksheetFunction.CountIfs(Sheets("Final UWS").Range("K" & (md + 1) & ":K" & (ld - 1)), "Q6")

pwg = Application.WorksheetFunction.CountIfs(Sheets("Final UWS").Range("K" & (ld + 1) & ":K" & (bk - 1)), "PWG")
pld = Application.WorksheetFunction.CountIfs(Sheets("Final UWS").Range("K" & (ld + 1) & ":K" & (bk - 1)), "PLD")

Sheets("Final UWS").Range("Q3").Value = "MD"
Sheets("Final UWS").Range("Q4").Value = "LD"
Sheets("Final UWS").Range("R3:T4").ClearFormats

''Q4
If q4 < 4 Then
    With Sheets("Final UWS").Range("R3")
        .Value = (4 - q4) & " Q4 MISSING!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf q4 > 4 Then
    With Sheets("Final UWS").Range("R3")
        .Value = (q4 - 4) & " Q4 EXCESS!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf q4 = 4 Then
    With Sheets("Final UWS").Range("R3")
        .Value = "Q4 sufficient"
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
End If

''Q5
If q5 < 22 Then
    With Sheets("Final UWS").Range("S3")
        .Value = (22 - q5) & " Q5 MISSING!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf q5 > 22 Then
    With Sheets("Final UWS").Range("S3")
        .Value = (q5 - 22) & " Q5 EXCESS!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf q5 = 22 Then
    With Sheets("Final UWS").Range("S3")
        .Value = "Q5 sufficient"
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
End If

''Q6
If q6 < 1 Then
    With Sheets("Final UWS").Range("T3")
        .Value = (1 - q6) & " Q6 MISSING!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf q6 > 1 Then
    With Sheets("Final UWS").Range("T3")
        .Value = (q6 - 1) & " Q6 EXCESS!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf q6 = 1 Then
    With Sheets("Final UWS").Range("T3")
        .Value = "Q6 sufficient"
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
End If

''PWG + PLD
If pwg + pld = 10 Then
    With Sheets("Final UWS").Range("R4")
        .Value = "PWG + PLD sufficient"
        .Font.Color = RGB(128, 128, 128)
        .Font.Italic = True
    End With
ElseIf pwg + pld > 10 Then
    With Sheets("Final UWS").Range("R4")
        .Value = (pwg + pld - 10) & " PWG/PLD EXCESS!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
ElseIf pwg + pld < 10 Then
    With Sheets("Final UWS").Range("R4")
        .Value = -(pwg + pld - 10) & " PWG/PLD MISSING!"
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With
End If

'Check table
''Unit Number
For i = (md + 1) To (bk + 1)
    If i = md Or i = ld Or i = bk Then
        Sheets("Final UWS").Range("Q" & i).Value = ""
    Else
        With Sheets("Final UWS").Range("Q" & i)
            .Value = Sheets("Final UWS").Range("B" & i).Value & Sheets("Final UWS").Range("C" & i).Value & Sheets("Final UWS").Range("D" & i).Value
        End With
    End If
Next i

''Gross Weight
With Sheets("Final UWS").Range("R" & (md + 1) & ":R" & (bk - 1))
    .Formula = "=IFERROR(VLOOKUP(Q" & md + 1 & ", PALLET!$D$7:$F$45, 3, 0), """")"
    .Value = .Value
End With

Sheets("Final UWS").Range("R" & (bk + 1)).Value = Sheets("PALLET").Range("F49")

''Contour Code
With Sheets("Final UWS").Range("S" & (md + 1) & ":S" & (bk - 1))
    .Formula = "=IFERROR(VLOOKUP(Q" & md + 1 & ", PALLET!$D$7:$J$45, 7, 0), """")"
    .Value = .Value
End With

''Special Code
With Sheets("Final UWS").Range("T" & (md + 1) & ":T" & (bk - 1))
    .Formula = "=IFERROR(VLOOKUP(Q" & md + 1 & ", PALLET!$D$7:$N$45, 11, 0), """")"
    .Value = .Value
End With

''CMDT Code
With Sheets("Final UWS").Range("U" & (md + 1) & ":U" & (bk - 1))
    .Formula = "=IFERROR(VLOOKUP(Q" & md + 1 & ", PALLET!$D$7:$N$45, 10, 0), """")"
    .Value = .Value
End With

Sheets("Final UWS").Range("U" & (bk + 1)).Value = Sheets("PALLET").Range("M49")

''Remark
Sheets("Final UWS").Range("V" & (bk + 1)).Value = Sheets("PALLET").Range("P49")

'Cross Check
''Gross Weight
Sheets("Final UWS").Range("R" & (bk + 2)).Value = Sheets("PALLET").Range("F50").Value
For i = (md + 1) To (bk + 2)
    If Sheets("Final UWS").Range("R" & i).Value <> Sheets("Final UWS").Range("H" & i).Value Then
        Sheets("Final UWS").Range("R" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("Final UWS").Range("R" & i).Font.Color = RGB(0, 0, 0)
    End If
Next i

''Contour
For i = (md + 1) To (bk - 1)
    If Sheets("Final UWS").Range("S" & i).Value <> Sheets("Final UWS").Range("K" & i).Value Then
        Sheets("Final UWS").Range("S" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("Final UWS").Range("S" & i).Font.Color = RGB(0, 0, 0)
    End If
Next i

''Special Code
For i = (md + 1) To (bk - 1)
    If Sheets("Final UWS").Range("T" & i).Value <> Sheets("Final UWS").Range("N" & i).Value Then
        Sheets("Final UWS").Range("T" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("Final UWS").Range("T" & i).Font.Color = RGB(0, 0, 0)
    End If
Next i

''CMDT Code
For i = (md + 1) To (bk + 1)
    If Sheets("Final UWS").Range("U" & i).Value <> Sheets("Final UWS").Range("L" & i).Value Then
        Sheets("Final UWS").Range("U" & i).Font.Color = RGB(255, 0, 0)
    Else
        Sheets("Final UWS").Range("U" & i).Font.Color = RGB(0, 0, 0)
    End If
Next i

''Remark
If Sheets("Final UWS").Range("O" & (bk + 1)) <> Sheets("Final UWS").Range("V" & (bk + 1)) Then
    Sheets("Final UWS").Range("V" & (bk + 1)).Font.Color = RGB(255, 0, 0)
End If

'AutoFit
Sheets("Final UWS").Range("Q:U").EntireColumn.AutoFit

End Sub

```
# 4_UWSclear
```vba
```vba
Sub FRTclear()

Dim warn As Integer

warn = MsgBox("Are you sure want to CLEAR ALL", vbYesNo + vbQuestion, "Confirm")

If warn = vbYes Then
    Sheets("Final UWS").Cells.Clear
Else
    MsgBox "Oke iem"
End If

End Sub
```

# 4_UWSsave
```vba
Sub UWSsave()
    Dim timeStamp As String
    Dim savePath As String
    Dim newWb As Workbook

    timeStamp = UCase(Format(Now, "mmmdd_hhmmss"))
    savePath = Sheets("Code").Range("E1").Value & "\UWS " & timeStamp & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Copies PAX to a brand new, single-sheet workbook
    Sheets("Final UWS").Copy
    Set newWb = ActiveWorkbook

    ' Save as XLSX and close
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "File saved at " & savePath
End Sub
```

# 5_NOTOC
```vba
Sub NOTOC()

Dim i As Integer
Dim fr, lr As Integer

'Determine 1st row
i = 1
While Sheets("NOTOC").Range("A" & i).Value <> "DOH"
    i = i + 1
Wend
fr = i

'Determine last row
While Sheets("NOTOC").Range("A" & i).Value = "DOH"
    i = i + 1
Wend
lr = i - 1

'Clear format
Sheets("NOTOC").Range("T:U").ClearFormats

'Add title
With Sheets("NOTOC")
    .Range("T" & (fr - 1)).Value = "DocNo"
    .Range("U" & (fr - 1)).Value = "UnitNo"
    .Range("T" & (fr - 1) & ":U" & (fr - 1)).Font.Bold = True
    .Range("T" & (fr - 1) & ":U" & (fr - 1)).HorizontalAlignment = xlCenter
End With

'PAX or FRT
Dim ch As Integer

ch = MsgBox("Yes for PAX, No for FRT", vbYesNo + vbQuestion, "Sale data today")

If ch = vbYes Then
    Sheets("NOTOC").Range("T1").Value = "PAX"
Else
    Sheets("NOTOC").Range("T1").Value = "FRT"
End If

'Vlookup DocNo from PAX/FRT
If Sheets("NOTOC").Range("T1").Value = "PAX" Then
    With Sheets("NOTOC").Range("T" & fr & ":T" & lr)
        .Formula = "=IFERROR(VLOOKUP(B" & fr & ", PAX!$B$1:$B$1000, 1, 0), ""Mismatched"")"
        .Value = .Value
    End With
ElseIf Sheets("NOTOC").Range("T1").Value = "FRT" Then
    With Sheets("NOTOC").Range("T" & fr & ":T" & lr)
        .Formula = "=IFERROR(VLOOKUP(B" & fr & ", FRT!$B$1:$B$1000, 1, 0), ""Mismatched"")"
        .Value = .Value
    End With
End If

'Vlookup ULD from PALLET
With Sheets("NOTOC").Range("U" & fr & ":U" & lr)
    .Formula = "=IFERROR(VLOOKUP(M" & fr & ", PALLET!$D$1:$D$1000, 1, 0), ""Mismatched"")"
    .Value = .Value
End With

'Condional formating
For i = fr To lr
    'DocNo
    If Sheets("NOTOC").Range("T" & i).Value = "Mismatched" Then
        With Sheets("NOTOC").Range("T" & i)
            .Value = Sheets("NOTOC").Range("B" & i)
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
    End If
    'UnitNo
    If Sheets("NOTOC").Range("U" & i).Value = "Mismatched" Then
        With Sheets("NOTOC").Range("U" & i)
            .Value = Sheets("NOTOC").Range("M" & i)
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
        End With
    End If
Next i

'AutoFit
Sheets("NOTOC").Range("T:U").EntireColumn.AutoFit

End Sub

```

# 5_NOTOCclear
```vba
Sub NOTOCclear()

Dim warn As Integer

warn = MsgBox("Are you sure want to CLEAR ALL", vbYesNo + vbQuestion, "Confirm")

If warn = vbYes Then
    Sheets("NOTOC").Cells.Clear
Else
    MsgBox "Oke iem"
End If

End Sub
```

# 5_NOTOCsave
```vba
Sub NOTOCsave()
    Dim timeStamp As String
    Dim savePath As String
    Dim newWb As Workbook

    timeStamp = UCase(Format(Now, "mmmdd_hhmmss"))
    savePath = Sheets("Code").Range("E1").Value & "\NOTOC " & timeStamp & ".xlsx"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Copies PAX to a brand new, single-sheet workbook
    Sheets("NOTOC").Copy
    Set newWb = ActiveWorkbook

    ' Save as XLSX and close
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "File saved at " & savePath
End Sub
```
