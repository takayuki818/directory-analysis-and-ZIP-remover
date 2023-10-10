Attribute VB_Name = "ZIPリムーバ"
Option Explicit
Sub ZIP書き出し()
    Dim Z行 As Long, 行 As Long, 終行 As Long
    Application.ScreenUpdating = False
    Call ZIP書き出しクリア
    Z行 = 2
    終行 = Sheets("解析フォーム").Cells(Rows.Count, 2).End(xlUp).Row
    For 行 = 11 To 終行
        If Right(Sheets("解析フォーム").Cells(行, 1), 4) = ".zip" Then
            Sheets("ZIPリムーバ").Cells(Z行, 3) = Sheets("解析フォーム").Cells(行, 1)
            Sheets("ZIPリムーバ").Cells(Z行, 4) = Sheets("解析フォーム").Cells(行, 4)
            Z行 = Z行 + 1
        End If
    Next
    With Sheets("ZIPリムーバ")
        Range(.Cells(2, 3), .Cells(.Cells(Rows.Count, 3).End(xlUp).Row, 4)).Borders.LineStyle = True
    End With
    Application.ScreenUpdating = True
End Sub
Sub ZIP書き出しクリア()
    With Range(Sheets("ZIPリムーバ").Cells(2, 3), Sheets("ZIPリムーバ").Cells(Rows.Count, 4))
        .Value = Empty
        .IndentLevel = 0
        .Font.Color = RGB(0, 0, 0)
        .Borders.LineStyle = False
    End With
End Sub
Sub ZIPリムーブ()
    Dim FSO As New FileSystemObject
    Dim 文 As String
    Dim 行 As Long, 終行 As Long
    文 = "対象ZIP一覧にあるファイルを全てリムーブ先フォルダへ移動します" & vbCrLf & vbCrLf & "本当に実行してよろしいですか？"
    If MsgBox(文, vbYesNo) = vbYes Then
        Application.ScreenUpdating = False
        Set FSO = CreateObject("Scripting.FileSystemObject")
        With Sheets("ZIPリムーバ")
            終行 = .Cells(Rows.Count, 4).End(xlUp).Row
            For 行 = 2 To 終行
                FSO.MoveFile .Cells(行, 4), .Cells(2, 1) & "\" & .Cells(行, 3)
            Next
            Set FSO = Nothing
        End With
        Application.ScreenUpdating = True
        MsgBox "処理が完了しました"
    End If
End Sub

