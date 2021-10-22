Attribute VB_Name = "ModEventBackup"
Option Explicit

'ワークブック保存時にフォルダに上書きバックアップ                  ・・・元場所：FukamiAddins3.ModEventBackup
'ワークブック保存時にフォルダに日付をつけてバックアップ            ・・・元場所：FukamiAddins3.ModEventBackup
'ワークブック保存時に同じフォルダ上に上書きバックアップ            ・・・元場所：FukamiAddins3.ModEventBackup
'ワークブック保存時に同じフォルダ上に日付をつけて上書きバックアップ・・・元場所：FukamiAddins3.ModEventBackup



Sub ワークブック保存時にフォルダに上書きバックアップ()
'ワークブック保存時にフォルダに上書きバックアップ
'20210721

    Dim FilePath   As String
    Dim FolderName As String
    Dim BookName   As String
    FilePath = ThisWorkbook.Path
    FolderName = "Backup" 'バックアップするフォルダの名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    If Dir(FilePath & "\" & FolderName, vbDirectory) = "" Then 'フォルダがない場合はフォルダを作成
        MkDir FilePath & "\" & FolderName
    End If
    
    Dim FSO As New FileSystemObject
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & FolderName & "\" & BookName
    
End Sub

Sub ワークブック保存時にフォルダに日付をつけてバックアップ()
'ワークブック保存時にフォルダに日付をつけてバックアップ
'20210721

    Dim FilePath   As String
    Dim FolderName As String
    Dim BookName   As String
    Dim Extension  As String
    Dim BookName2  As String
    Dim StrTime    As String
    FilePath = ThisWorkbook.Path
    FolderName = "Backup" 'バックアップするフォルダの名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    If Dir(FilePath & "\" & FolderName, vbDirectory) = "" Then 'フォルダがない場合はフォルダを作成
        MkDir FilePath & "\" & FolderName
    End If
    
    Dim FSO As New FileSystemObject
    Extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & Extension, "")
    StrTime = Format(Now(), "YYYYMMDDhhmmss") '←←←←←←←←←←←←←←←←←←←←←←←
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & FolderName & "\" & BookName2 & " " & StrTime & "." & Extension
    
End Sub

Sub ワークブック保存時に同じフォルダ上に上書きバックアップ()
'ワークブック保存時に同じフォルダ上に上書きバックアップ
'20210721

    Dim FilePath  As String
    Dim AddStr    As String
    Dim BookName  As String
    Dim Extension As String
    Dim BookName2 As String
    Dim StrTime   As String
    FilePath = ThisWorkbook.Path
    AddStr = "Backup" 'バックアップファイルの語尾につく名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    Dim FSO As New FileSystemObject
    Extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & Extension, "")
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & BookName2 & "_" & AddStr & "." & Extension
    
End Sub

Sub ワークブック保存時に同じフォルダ上に日付をつけて上書きバックアップ()
'ワークブック保存時に同じフォルダ上に日付をつけて上書きバックアップ
'20210721

    Dim FilePath  As String
    Dim AddStr    As String
    Dim BookName  As String
    Dim Extension As String
    Dim BookName2 As String
    Dim StrTime   As String
    FilePath = ThisWorkbook.Path
    AddStr = "Backup" 'バックアップファイルの語尾につく名前←←←←←←←←←←←←←←←←←←←←←←←
    BookName = ThisWorkbook.Name
    
    Dim FSO As New FileSystemObject
    Extension = FSO.GetExtensionName(BookName)
    BookName2 = Replace(BookName, "." & Extension, "")
    StrTime = Format(Now(), "YYYYMMDDhhmmss") '←←←←←←←←←←←←←←←←←←←←←←←
    FSO.CopyFile FilePath & "\" & BookName, FilePath & "\" & BookName2 & "_" & AddStr & "_" & StrTime & "." & Extension
    
End Sub


