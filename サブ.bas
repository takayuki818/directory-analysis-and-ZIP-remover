Attribute VB_Name = "�T�u"
Option Explicit
Sub �����N�쐬()
    Dim �I�s As Long, �s As Long
    Dim �A�h���X As String, �e�L�X�g As String
    �I�s = Sheets("��̓t�H�[��").Cells(Rows.Count, 2).End(xlUp).Row
    For �s = 11 To �I�s
        If Sheets("��̓t�H�[��").Cells(�s, 4) <> "" Then
            �A�h���X = Cells(�s, 4)
            �e�L�X�g = Cells(�s, 1)
            Sheets("��̓t�H�[��").Hyperlinks.Add Anchor:=Cells(�s, 5), Address:=�A�h���X, TextToDisplay:=�e�L�X�g
        End If
    Next
End Sub
Sub ��̓N���A()
    With Sheets("��̓t�H�[��")
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
Sub �ی�ؑ�()
    With Sheets("��̓t�H�[��")
        Select Case .ProtectContents
            Case True: .Unprotect: MsgBox "�V�[�g�ی���������܂���"
            Case False: .Protect: MsgBox "�V�[�g��ی삵�܂���"
        End Select
    End With
End Sub
