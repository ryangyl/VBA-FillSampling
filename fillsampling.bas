Attribute VB_Name = "Module2"
Sub fillsampling()

Dim main As Worksheet
Set main = Sheets("Main")
Dim grs1 As Worksheet
Set grs1 = Sheets("GRS0")
Dim todaydate As Long
Dim todaymonth As Integer

Dim nr As Integer
nr = WorksheetFunction.CountA(grs1.Range("A:A"))

Dim i As Integer
Dim s0 As Long
Dim sc As Integer
Dim sr As Integer

For i = 3 To nr
On Error Resume Next
sc = 0
sr = 0
s0 = CLng(grs1.Range("f" & i))
sc = Application.XMatch(s0, main.Range("2:2"), 0)
sr = Application.XMatch(grs1.Range("c" & i), main.Range("b:b"), 0)
If Not IsError(sc) And Not IsError(sc) Then
        If main.Cells(sr, sc).Value = "" Then
        
            If CLng(grs1.Range("f" & i)) < CLng(Date) Then 'due
                If grs1.Range("g" & i) = "" Then 'not done
                    main.Cells(sr, sc).Interior.Color = RGB(0, 0, 255)
                    main.Cells(sr, sc).Value = "S"
                Else 'done
                    If grs1.Range("g" & i) < grs1.Range("f" & i) Then 'early
                        sc = 0
                        s0 = CLng(grs1.Range("g" & i))
                        sc = Application.XMatch(s0, main.Range("2:2"), 0)
                        main.Cells(sr, sc).Value = "ES"
                        main.Cells(sr, sc).Interior.Color = RGB(0, 255, 0)
                    ElseIf grs1.Range("g" & i) = grs1.Range("f" & i) Then 'early
                        main.Cells(sr, sc).Value = "S"
                        main.Cells(sr, sc).Interior.Color = RGB(0, 255, 0)
                    Else 'late
                        main.Cells(sr, sc).Interior.Color = RGB(225, 0, 0)
                        main.Cells(sr, sc).Value = "S"
                        sc = 0
                        s0 = CLng(grs1.Range("g" & i))
                        sc = Application.XMatch(s0, main.Range("2:2"), 0)
                        main.Cells(sr, sc).Value = "s"
                        main.Cells(sr, sc).Interior.Color = RGB(0, 255, 0)
                    End If
                End If
            Else
                main.Cells(sr, sc).Value = "S"
            End If
        ElseIf main.Cells(sr, sc).Value = "R" Or main.Cells(sr, sc).Value = "r" Or main.Cells(sr, sc).Value = "ER" Then
            If grs1.Range("g" & i) = "" Then 'not done
                    main.Cells(sr, sc).Interior.Color = RGB(0, 0, 255)
                    main.Cells(sr, sc).Value = main.Cells(sr, sc).Value & "/S"
                Else 'done
                    If grs1.Range("g" & i) < grs1.Range("f" & i) Then 'early
                        sc = 0
                        s0 = CLng(grs1.Range("g" & i))
                        sc = Application.XMatch(s0, main.Range("2:2"), 0)
                        main.Cells(sr, sc).Value = main.Cells(sr, sc).Value & "/ES"
                        main.Cells(sr, sc).Interior.Color = RGB(0, 255, 0)
                    ElseIf grs1.Range("g" & i) = grs1.Range("f" & i) Then 'early
                        main.Cells(sr, sc).Value = main.Cells(sr, sc).Value & "/S"
                        main.Cells(sr, sc).Interior.Color = RGB(0, 255, 0)
                    Else 'late
                        main.Cells(sr, sc).Interior.Color = RGB(225, 0, 0)
                        main.Cells(sr, sc).Value = main.Cells(sr, sc).Value & "/S"
                        sc = 0
                        s0 = CLng(grs1.Range("g" & i))
                        sc = Application.XMatch(s0, main.Range("2:2"), 0)
                        main.Cells(sr, sc).Value = main.Cells(sr, sc).Value & "/s"
                        main.Cells(sr, sc).Interior.Color = RGB(0, 255, 0)
                    End If
        
    End If
    End If
End If
Next i


End Sub

