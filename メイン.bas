Attribute VB_Name = "メイン"
Option Explicit
Sub D直下のみ書出開始()
    Dim FSO As New FileSystemObject
    Dim 起点フォルダ As Folder
    Dim インデント As Long
    Dim 始時 As Date, 終時 As Date
    Application.ScreenUpdating = False
    始時 = Timer
    実行中.Show vbModeless
    実行中.Repaint
    Call 解析クリア
    Set 起点フォルダ = FSO.GetFolder(Sheets("解析フォーム").Range("参照ディレクトリ"))
    インデント = 0
    Call D直下のみ再帰書出(起点フォルダ, インデント)
    With Sheets("解析フォーム")
        .Unprotect
        Range(.Cells(11, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
        .Protect
    End With
    終時 = Timer
    MsgBox "処理が完了しました。" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時
    Unload 実行中
    Application.ScreenUpdating = True
End Sub
Sub D直下のみ再帰書出(親フォルダ, インデント)
    Dim ファイル As File
    Dim フォルダ As Folder
    Dim 行 As Long
    With Sheets("解析フォーム")
        .Unprotect
        For Each フォルダ In 親フォルダ.SubFolders
            行 = .Cells(Rows.Count, 2).End(xlUp).Row + 1
            .Cells(行, 1).IndentLevel = インデント
            .Cells(行, 1).Font.Color = RGB(255, 0, 0)
            .Cells(行, 1).Value = フォルダ.Name
            .Cells(行, 2).Value = インデント
            .Cells(行, 4).Value = フォルダ.Path
        Next
        For Each ファイル In 親フォルダ.Files
            If ファイル.Name <> "Thumbs.db" Then
                行 = .Cells(Rows.Count, 2).End(xlUp).Row + 1
                .Cells(行, 1).IndentLevel = インデント + 1
                .Cells(行, 1).Value = ファイル.Name
                .Cells(行, 2).Value = インデント + 1
                .Cells(行, 3).Value = Round(ファイル.Size / 1048576, 2)
                .Cells(行, 4).Value = ファイル.Path
            End If
        Next
        .Protect
    End With
End Sub
Sub D下全書出開始()
    Dim FSO As New FileSystemObject
    Dim 起点フォルダ As Folder
    Dim インデント As Long
    Dim 始時 As Date, 終時 As Date
    Application.ScreenUpdating = False
    If MsgBox("参照ディレクトリ以下の状況により、処理が長時間に及ぶ可能性があります。" & vbCrLf & vbCrLf & "処理を開始してよろしいですか？", vbYesNo) = vbYes Then
        始時 = Timer
        実行中.Show vbModeless
        実行中.Repaint
        Call 解析クリア
        Set 起点フォルダ = FSO.GetFolder(Sheets("解析フォーム").Range("参照ディレクトリ"))
        Call D下全再帰書出(起点フォルダ, インデント)
        With Sheets("解析フォーム")
            .Unprotect
            Range(.Cells(11, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
            .Protect
        End With
        終時 = Timer
        MsgBox "処理が完了しました。" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時
        Unload 実行中
    End If
    Application.ScreenUpdating = True
End Sub
Function D下全再帰書出(親フォルダ, インデント)
    Dim ファイル As File
    Dim フォルダ As Folder
    Dim ファイルサイズ As Double, サブサイズ As Double
    Dim フォルダ行 As Long, 行 As Long
    With Sheets("解析フォーム")
        .Unprotect
        フォルダ行 = .Cells(Rows.Count, 2).End(xlUp).Row + 1
        .Cells(フォルダ行, 1).IndentLevel = インデント
        .Cells(フォルダ行, 1).Font.Color = RGB(255, 0, 0)
        .Cells(フォルダ行, 1).Value = 親フォルダ.Name
        .Cells(フォルダ行, 2).Value = インデント
        .Cells(フォルダ行, 4).Value = 親フォルダ.Path
        For Each ファイル In 親フォルダ.Files
            If ファイル.Name <> "Thumbs.db" Then
                行 = .Cells(Rows.Count, 2).End(xlUp).Row + 1
                .Cells(行, 1).IndentLevel = インデント + 1
                .Cells(行, 1).Value = ファイル.Name
                .Cells(行, 2).Value = インデント + 1
                .Cells(行, 3).Value = Round(ファイル.Size / 1048576, 2)
                .Cells(行, 4).Value = ファイル.Path
                ファイルサイズ = Round(ファイルサイズ + ファイル.Size / 1048576, 2)
            End If
        Next
        For Each フォルダ In 親フォルダ.SubFolders
            サブサイズ = サブサイズ + D下全再帰書出(フォルダ, インデント + 1)
        Next
        .Cells(フォルダ行, 3) = ファイルサイズ + サブサイズ
        D下全再帰書出 = .Cells(フォルダ行, 3)
        .Protect
    End With
End Function
Sub D下フォルダのみ書出開始()
    Dim FSO As New FileSystemObject
    Dim 文 As String
    Dim 起点フォルダ As Folder
    Dim インデント As Long
    Dim 始時 As Date, 終時 As Date
    文 = "参照ディレクトリ以下の状況により、処理が長時間に及ぶ可能性があります。" & vbCrLf & vbCrLf & "処理を開始してよろしいですか？"
    If MsgBox(文, vbYesNo) = vbYes Then
        Application.ScreenUpdating = False
        始時 = Timer
        実行中.Show vbModeless
        実行中.Repaint
        Call 解析クリア
        Set 起点フォルダ = FSO.GetFolder(Sheets("解析フォーム").Range("参照ディレクトリ"))
        インデント = 0
        Call D下フォルダのみ再帰書出(起点フォルダ, インデント)
        With Sheets("解析フォーム")
            .Unprotect
            Range(.Cells(11, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
            .Protect
        End With
        終時 = Timer
        MsgBox "処理が完了しました。" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時
        Unload 実行中
        Application.ScreenUpdating = True
    End If
End Sub
Function D下フォルダのみ再帰書出(親フォルダ, インデント)
    Dim ファイル As File
    Dim フォルダ As Folder
    Dim 行 As Long
    Dim ファイルサイズ As Double, サブサイズ As Double
    With Sheets("解析フォーム")
        .Unprotect
        行 = .Cells(Rows.Count, 2).End(xlUp).Row + 1
        .Cells(行, 1).IndentLevel = インデント
        .Cells(行, 1).Value = 親フォルダ.Name
        .Cells(行, 2).Value = インデント
        .Cells(行, 4).Value = 親フォルダ.Path
        For Each ファイル In 親フォルダ.Files
            If ファイル.Name <> "Thumbs.db" Then
                ファイルサイズ = Round(ファイルサイズ + ファイル.Size / 1048576, 2)
            End If
        Next
        For Each フォルダ In 親フォルダ.SubFolders
            サブサイズ = サブサイズ + D下フォルダのみ再帰書出(フォルダ, インデント + 1)
        Next
        .Cells(行, 3) = ファイルサイズ + サブサイズ
        D下フォルダのみ再帰書出 = .Cells(行, 3)
        .Protect
    End With
End Function

