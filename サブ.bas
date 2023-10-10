Attribute VB_Name = "サブ"
Option Explicit
Sub リンク作成()
    Dim 終行 As Long, 行 As Long
    Dim アドレス As String, テキスト As String
    終行 = Sheets("解析フォーム").Cells(Rows.Count, 2).End(xlUp).Row
    For 行 = 11 To 終行
        If Sheets("解析フォーム").Cells(行, 4) <> "" Then
            アドレス = Cells(行, 4)
            テキスト = Cells(行, 1)
            Sheets("解析フォーム").Hyperlinks.Add Anchor:=Cells(行, 5), Address:=アドレス, TextToDisplay:=テキスト
        End If
    Next
End Sub
Sub 解析クリア()
    With Sheets("解析フォーム")
        .Unprotect
        With Range(.Cells(11, 1), .Cells(Rows.Count, 5))
            .Value = Empty
            .Font.Color = RGB(0, 0, 0)
            .Borders.LineStyle = False
            .IndentLevel = 0
        End With
        .Protect
    End With
End Sub
Sub 保護切替()
    With Sheets("解析フォーム")
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "シート保護を解除しました"
            Case False: .Protect: MsgBox "シートを保護しました"
        End Select
    End With
End Sub
