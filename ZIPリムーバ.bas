Attribute VB_Name = "ZIP�����[�o"
Option Explicit
Sub ZIP�����o��()
    Dim Z�s As Long, �s As Long, �I�s As Long
    Application.ScreenUpdating = False
    Call ZIP�����o���N���A
    Z�s = 2
    �I�s = Sheets("��̓t�H�[��").Cells(Rows.Count, 2).End(xlUp).Row
    For �s = 11 To �I�s
        If Right(Sheets("��̓t�H�[��").Cells(�s, 1), 4) = ".zip" Then
            Sheets("ZIP�����[�o").Cells(Z�s, 3) = Sheets("��̓t�H�[��").Cells(�s, 1)
            Sheets("ZIP�����[�o").Cells(Z�s, 4) = Sheets("��̓t�H�[��").Cells(�s, 4)
            Z�s = Z�s + 1
        End If
    Next
    With Sheets("ZIP�����[�o")
        Range(.Cells(2, 3), .Cells(.Cells(Rows.Count, 3).End(xlUp).Row, 4)).Borders.LineStyle = True
    End With
    Application.ScreenUpdating = True
End Sub
Sub ZIP�����o���N���A()
    With Range(Sheets("ZIP�����[�o").Cells(2, 3), Sheets("ZIP�����[�o").Cells(Rows.Count, 4))
        .Value = Empty
        .IndentLevel = 0
        .Font.Color = RGB(0, 0, 0)
        .Borders.LineStyle = False
    End With
End Sub
Sub ZIP�����[�u()
    Dim FSO As New FileSystemObject
    Dim �� As String
    Dim �s As Long, �I�s As Long
    �� = "�Ώ�ZIP�ꗗ�ɂ���t�@�C����S�ă����[�u��t�H���_�ֈړ����܂�" & vbCrLf & vbCrLf & "�{���Ɏ��s���Ă�낵���ł����H"
    If MsgBox(��, vbYesNo) = vbYes Then
        Application.ScreenUpdating = False
        Set FSO = CreateObject("Scripting.FileSystemObject")
        With Sheets("ZIP�����[�o")
            �I�s = .Cells(Rows.Count, 4).End(xlUp).Row
            For �s = 2 To �I�s
                FSO.MoveFile .Cells(�s, 4), .Cells(2, 1) & "\" & .Cells(�s, 3)
            Next
            Set FSO = Nothing
        End With
        Application.ScreenUpdating = True
        MsgBox "�������������܂���"
    End If
End Sub

