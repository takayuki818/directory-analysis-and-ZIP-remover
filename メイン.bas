Attribute VB_Name = "���C��"
Option Explicit
Sub D�����̂ݏ��o�J�n()
    Dim FSO As New FileSystemObject
    Dim �N�_�t�H���_ As Folder
    Dim �C���f���g As Long
    Dim �n�� As Date, �I�� As Date
    Application.ScreenUpdating = False
    �n�� = Timer
    ���s��.Show vbModeless
    ���s��.Repaint
    Call ��̓N���A
    Set �N�_�t�H���_ = FSO.GetFolder(Sheets("��̓t�H�[��").Range("�Q�ƃf�B���N�g��"))
    �C���f���g = 0
    Call D�����̂ݍċA���o(�N�_�t�H���_, �C���f���g)
    With Sheets("��̓t�H�[��")
        .Unprotect
        Range(.Cells(11, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
        .Protect
    End With
    �I�� = Timer
    MsgBox "�������������܂����B" & vbCrLf & vbCrLf & "�������ԁF" & �I�� - �n��
    Unload ���s��
    Application.ScreenUpdating = True
End Sub
Sub D�����̂ݍċA���o(�e�t�H���_, �C���f���g)
    Dim �t�@�C�� As File
    Dim �t�H���_ As Folder
    Dim �s As Long
    With Sheets("��̓t�H�[��")
        .Unprotect
        For Each �t�H���_ In �e�t�H���_.SubFolders
            �s = .Cells(Rows.Count, 2).End(xlUp).Row + 1
            .Cells(�s, 1).IndentLevel = �C���f���g
            .Cells(�s, 1).Font.Color = RGB(255, 0, 0)
            .Cells(�s, 1).Value = �t�H���_.Name
            .Cells(�s, 2).Value = �C���f���g
            .Cells(�s, 4).Value = �t�H���_.Path
        Next
        For Each �t�@�C�� In �e�t�H���_.Files
            If �t�@�C��.Name <> "Thumbs.db" Then
                �s = .Cells(Rows.Count, 2).End(xlUp).Row + 1
                .Cells(�s, 1).IndentLevel = �C���f���g + 1
                .Cells(�s, 1).Value = �t�@�C��.Name
                .Cells(�s, 2).Value = �C���f���g + 1
                .Cells(�s, 3).Value = Round(�t�@�C��.Size / 1048576, 2)
                .Cells(�s, 4).Value = �t�@�C��.Path
            End If
        Next
        .Protect
    End With
End Sub
Sub D���S���o�J�n()
    Dim FSO As New FileSystemObject
    Dim �N�_�t�H���_ As Folder
    Dim �C���f���g As Long
    Dim �n�� As Date, �I�� As Date
    Application.ScreenUpdating = False
    If MsgBox("�Q�ƃf�B���N�g���ȉ��̏󋵂ɂ��A�����������Ԃɋy�ԉ\��������܂��B" & vbCrLf & vbCrLf & "�������J�n���Ă�낵���ł����H", vbYesNo) = vbYes Then
        �n�� = Timer
        ���s��.Show vbModeless
        ���s��.Repaint
        Call ��̓N���A
        Set �N�_�t�H���_ = FSO.GetFolder(Sheets("��̓t�H�[��").Range("�Q�ƃf�B���N�g��"))
        Call D���S�ċA���o(�N�_�t�H���_, �C���f���g)
        With Sheets("��̓t�H�[��")
            .Unprotect
            Range(.Cells(11, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
            .Protect
        End With
        �I�� = Timer
        MsgBox "�������������܂����B" & vbCrLf & vbCrLf & "�������ԁF" & �I�� - �n��
        Unload ���s��
    End If
    Application.ScreenUpdating = True
End Sub
Function D���S�ċA���o(�e�t�H���_, �C���f���g)
    Dim �t�@�C�� As File
    Dim �t�H���_ As Folder
    Dim �t�@�C���T�C�Y As Double, �T�u�T�C�Y As Double
    Dim �t�H���_�s As Long, �s As Long
    With Sheets("��̓t�H�[��")
        .Unprotect
        �t�H���_�s = .Cells(Rows.Count, 2).End(xlUp).Row + 1
        .Cells(�t�H���_�s, 1).IndentLevel = �C���f���g
        .Cells(�t�H���_�s, 1).Font.Color = RGB(255, 0, 0)
        .Cells(�t�H���_�s, 1).Value = �e�t�H���_.Name
        .Cells(�t�H���_�s, 2).Value = �C���f���g
        .Cells(�t�H���_�s, 4).Value = �e�t�H���_.Path
        For Each �t�@�C�� In �e�t�H���_.Files
            If �t�@�C��.Name <> "Thumbs.db" Then
                �s = .Cells(Rows.Count, 2).End(xlUp).Row + 1
                .Cells(�s, 1).IndentLevel = �C���f���g + 1
                .Cells(�s, 1).Value = �t�@�C��.Name
                .Cells(�s, 2).Value = �C���f���g + 1
                .Cells(�s, 3).Value = Round(�t�@�C��.Size / 1048576, 2)
                .Cells(�s, 4).Value = �t�@�C��.Path
                �t�@�C���T�C�Y = Round(�t�@�C���T�C�Y + �t�@�C��.Size / 1048576, 2)
            End If
        Next
        For Each �t�H���_ In �e�t�H���_.SubFolders
            �T�u�T�C�Y = �T�u�T�C�Y + D���S�ċA���o(�t�H���_, �C���f���g + 1)
        Next
        .Cells(�t�H���_�s, 3) = �t�@�C���T�C�Y + �T�u�T�C�Y
        D���S�ċA���o = .Cells(�t�H���_�s, 3)
        .Protect
    End With
End Function
Sub D���t�H���_�̂ݏ��o�J�n()
    Dim FSO As New FileSystemObject
    Dim �� As String
    Dim �N�_�t�H���_ As Folder
    Dim �C���f���g As Long
    Dim �n�� As Date, �I�� As Date
    �� = "�Q�ƃf�B���N�g���ȉ��̏󋵂ɂ��A�����������Ԃɋy�ԉ\��������܂��B" & vbCrLf & vbCrLf & "�������J�n���Ă�낵���ł����H"
    If MsgBox(��, vbYesNo) = vbYes Then
        Application.ScreenUpdating = False
        �n�� = Timer
        ���s��.Show vbModeless
        ���s��.Repaint
        Call ��̓N���A
        Set �N�_�t�H���_ = FSO.GetFolder(Sheets("��̓t�H�[��").Range("�Q�ƃf�B���N�g��"))
        �C���f���g = 0
        Call D���t�H���_�̂ݍċA���o(�N�_�t�H���_, �C���f���g)
        With Sheets("��̓t�H�[��")
            .Unprotect
            Range(.Cells(11, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)).Borders.LineStyle = True
            .Protect
        End With
        �I�� = Timer
        MsgBox "�������������܂����B" & vbCrLf & vbCrLf & "�������ԁF" & �I�� - �n��
        Unload ���s��
        Application.ScreenUpdating = True
    End If
End Sub
Function D���t�H���_�̂ݍċA���o(�e�t�H���_, �C���f���g)
    Dim �t�@�C�� As File
    Dim �t�H���_ As Folder
    Dim �s As Long
    Dim �t�@�C���T�C�Y As Double, �T�u�T�C�Y As Double
    With Sheets("��̓t�H�[��")
        .Unprotect
        �s = .Cells(Rows.Count, 2).End(xlUp).Row + 1
        .Cells(�s, 1).IndentLevel = �C���f���g
        .Cells(�s, 1).Value = �e�t�H���_.Name
        .Cells(�s, 2).Value = �C���f���g
        .Cells(�s, 4).Value = �e�t�H���_.Path
        For Each �t�@�C�� In �e�t�H���_.Files
            If �t�@�C��.Name <> "Thumbs.db" Then
                �t�@�C���T�C�Y = Round(�t�@�C���T�C�Y + �t�@�C��.Size / 1048576, 2)
            End If
        Next
        For Each �t�H���_ In �e�t�H���_.SubFolders
            �T�u�T�C�Y = �T�u�T�C�Y + D���t�H���_�̂ݍċA���o(�t�H���_, �C���f���g + 1)
        Next
        .Cells(�s, 3) = �t�@�C���T�C�Y + �T�u�T�C�Y
        D���t�H���_�̂ݍċA���o = .Cells(�s, 3)
        .Protect
    End With
End Function

