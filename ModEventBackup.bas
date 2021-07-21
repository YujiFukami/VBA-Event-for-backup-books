Attribute VB_Name = "ModEventBackup"
Option Explicit
'�C�x���g�p�̃v���V�[�W��
'���g���R�s�[���Ďg��

Sub ���[�N�u�b�N�ۑ����Ƀt�H���_�ɏ㏑���o�b�N�A�b�v()
'���[�N�u�b�N�ۑ����Ƀt�H���_�ɏ㏑���o�b�N�A�b�v
'20210721

    Dim FilePath$, FolderName$, BookName$
    FilePath = ThisWorkbook.Path
    FolderName = "Backup" '�o�b�N�A�b�v����t�H���_�̖��O����������������������������������������������
    BookName = ThisWorkbook.Name
    
    If Dir(FilePath & "\" & FolderName, vbDirectory) = "" Then '�t�H���_���Ȃ��ꍇ�̓t�H���_���쐬
        MkDir FilePath & "\" & FolderName
    End If
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & FolderName & "\" & BookName
    
End Sub
Sub ���[�N�u�b�N�ۑ����Ƀt�H���_�ɓ��t�����ăo�b�N�A�b�v()
'���[�N�u�b�N�ۑ����Ƀt�H���_�ɓ��t�����ăo�b�N�A�b�v
'20210721

    Dim FilePath$, FolderName$, BookName$, extension$, BookName2$, StrTime$
    FilePath = ThisWorkbook.Path
    FolderName = "Backup" '�o�b�N�A�b�v����t�H���_�̖��O����������������������������������������������
    BookName = ThisWorkbook.Name
    
    If Dir(FilePath & "\" & FolderName, vbDirectory) = "" Then '�t�H���_���Ȃ��ꍇ�̓t�H���_���쐬
        MkDir FilePath & "\" & FolderName
    End If
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & extension, "")
    StrTime = Format(Now(), "YYYYMMDDhhmmss") '����������������������������������������������
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & FolderName & "\" & BookName2 & " " & StrTime & "." & extension
    
End Sub
Sub ���[�N�u�b�N�ۑ����ɓ����t�H���_��ɏ㏑���o�b�N�A�b�v()
'���[�N�u�b�N�ۑ����ɓ����t�H���_��ɏ㏑���o�b�N�A�b�v
'20210721

    Dim FilePath$, AddStr$, BookName$, extension$, BookName2$, StrTime$
    FilePath = ThisWorkbook.Path
    AddStr = "backup" '�o�b�N�A�b�v�t�@�C���̌���ɂ����O����������������������������������������������
    BookName = ThisWorkbook.Name
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & extension, "")
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & BookName2 & "_" & AddStr & "." & extension
    
End Sub

Sub ���[�N�u�b�N�ۑ����ɓ����t�H���_��ɓ��t�����ď㏑���o�b�N�A�b�v()
'���[�N�u�b�N�ۑ����ɓ����t�H���_��ɓ��t�����ď㏑���o�b�N�A�b�v
'20210721

    Dim FilePath$, AddStr$, BookName$, extension$, BookName2$, StrTime$
    FilePath = ThisWorkbook.Path
    AddStr = "backup" '�o�b�N�A�b�v�t�@�C���̌���ɂ����O����������������������������������������������
    BookName = ThisWorkbook.Name
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & extension, "")
    StrTime = Format(Now(), "YYYYMMDDhhmmss") '����������������������������������������������
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & BookName2 & "_" & AddStr & "_" & StrTime & "." & extension
    
End Sub
